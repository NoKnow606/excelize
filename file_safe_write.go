// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// DirtySheetWriteOptions controls worksheet-level serialization during
// non-destructive writes.
type DirtySheetWriteOptions struct {
	// DirtyWorksheets contains worksheet names or worksheet XML paths that may
	// be serialized from in-memory worksheet objects.
	DirtyWorksheets map[string]bool
	// PreserveCleanWorksheetXML forces clean worksheets to be copied from their
	// original package XML or temp file instead of f.Sheet.
	PreserveCleanWorksheetXML bool
}

// WriteNonDestructive provides a function to write the file to io.Writer
// WITHOUT modifying the internal worksheet state.
//
// Unlike Write(), this function preserves the original worksheet data in memory:
// - Does NOT call trimRow() to delete empty rows
// - Does NOT modify sheet.SheetData.Row arrays
// - Does NOT delete worksheets from memory (even without KeepWorksheetInMemory)
//
// This is essential for applications that need to:
// 1. Save file to persistent storage (e.g., GridFS, S3)
// 2. Continue performing operations on the same File object
// 3. Avoid state corruption from trimRow() modifications
//
// Performance note: This function creates deep copies of worksheet data,
// so it uses more memory than Write(). Use Write() if you don't need to
// preserve internal state.
//
// Example:
//
//	f := excelize.NewFile()
//	f.SetCellValue("Sheet1", "A1", "Hello")
//	f.SetCellValue("Sheet1", "A100", "World")
//
//	// Save to GridFS without destroying internal state
//	var buf bytes.Buffer
//	err := f.WriteNonDestructive(&buf)
//	if err != nil {
//	    return err
//	}
//	gridFS.Save(buf.Bytes())
//
//	// Continue working - internal state is preserved!
//	f.SetCellValue("Sheet1", "A50", "More data")  // ✅ Works correctly
//
//	// Compare with Write():
//	var buf2 bytes.Buffer
//	f.Write(&buf2)  // ⚠️ This calls trimRow() and deletes empty rows
//	f.SetCellValue("Sheet1", "A50", "More data")  // 💥 May fail due to state corruption
func (f *File) WriteNonDestructive(w io.Writer, opts ...Options) error {
	_, err := f.WriteToNonDestructive(w, opts...)
	return err
}

// WriteNonDestructiveWithDirtySheets writes the workbook while only
// serializing worksheets explicitly marked dirty. Clean worksheets are copied
// from their original XML bytes when PreserveCleanWorksheetXML is enabled.
func (f *File) WriteNonDestructiveWithDirtySheets(w io.Writer, dirtyOpts DirtySheetWriteOptions, opts ...Options) error {
	_, err := f.WriteToNonDestructiveWithDirtySheets(w, dirtyOpts, opts...)
	return err
}

// WriteToNonDestructive implements io.WriterTo to write the file without
// modifying internal state. Returns the number of bytes written.
func (f *File) WriteToNonDestructive(w io.Writer, opts ...Options) (int64, error) {
	return f.WriteToNonDestructiveWithDirtySheets(w, DirtySheetWriteOptions{}, opts...)
}

// WriteToNonDestructiveWithDirtySheets implements io.WriterTo using dirty
// worksheet awareness.
func (f *File) WriteToNonDestructiveWithDirtySheets(w io.Writer, dirtyOpts DirtySheetWriteOptions, opts ...Options) (int64, error) {
	for i := range opts {
		f.options = &opts[i]
	}
	buf, err := f.WriteToBufferNonDestructiveWithDirtySheets(dirtyOpts)
	if err != nil {
		return 0, err
	}
	return buf.WriteTo(w)
}

// WriteToBufferNonDestructive provides a function to get bytes.Buffer from
// the file WITHOUT modifying internal worksheet state.
//
// This function creates a snapshot of all worksheets for serialization,
// leaving the original in-memory data untouched.
func (f *File) WriteToBufferNonDestructive() (*bytes.Buffer, error) {
	return f.WriteToBufferNonDestructiveWithDirtySheets(DirtySheetWriteOptions{})
}

// WriteToBufferNonDestructiveWithDirtySheets provides a function to get
// bytes.Buffer from the file while preserving clean worksheet XML.
func (f *File) WriteToBufferNonDestructiveWithDirtySheets(dirtyOpts DirtySheetWriteOptions) (*bytes.Buffer, error) {
	buf := new(bytes.Buffer)
	zw := f.ZipWriter(buf)

	if err := f.writeToZipNonDestructive(zw, &dirtyOpts); err != nil {
		_ = zw.Close()
		return buf, err
	}
	if err := zw.Close(); err != nil {
		return buf, err
	}
	f.writeZip64LFH(buf)
	if f.options != nil && f.options.Password != "" {
		b, err := Encrypt(buf.Bytes(), f.options)
		if err != nil {
			return buf, err
		}
		buf.Reset()
		buf.Write(b)
	}
	return buf, nil
}

// writeToZipNonDestructive writes file content to zip WITHOUT modifying
// internal worksheet state.
func (f *File) writeToZipNonDestructive(zw ZipWriter, dirtyOpts *DirtySheetWriteOptions) error {
	// These writers don't modify worksheet state, safe to call directly
	f.calcChainWriter()
	f.commentsWriter()
	f.contentTypesWriter()
	f.drawingsWriter()
	f.volatileDepsWriter()
	f.vmlDrawingWriter()
	f.workBookWriter()

	// ⚠️ CRITICAL: workSheetWriter() calls trimRow() which modifies state!
	// Use our non-destructive version instead
	if err := f.workSheetWriterNonDestructive(dirtyOpts); err != nil {
		return err
	}

	f.relsWriter()
	_ = f.sharedStringsLoader()
	f.sharedStringsWriter()
	f.styleSheetWriter()
	f.themeWriter()

	// Write streams (unchanged from original)
	for path, stream := range f.streams {
		fi, err := zw.Create(path)
		if err != nil {
			return err
		}
		var from io.Reader
		if from, err = stream.rawData.Reader(); err != nil {
			_ = stream.rawData.Close()
			return err
		}
		if _, err := io.Copy(fi, from); err != nil {
			return err
		}
	}

	// Write remaining files from memory.
	var (
		n                int
		err              error
		files, tempFiles []string
	)
	preserveCleanWorksheets := dirtyOpts != nil && dirtyOpts.PreserveCleanWorksheetXML
	dirtyWorksheetPaths := f.dirtyWorksheetPaths(dirtyOpts)
	cleanWorksheetPaths := map[string]bool{}
	cleanTempWorksheetPaths := map[string]bool{}
	if preserveCleanWorksheets {
		cleanWorksheetPaths = f.cleanWorksheetPaths(dirtyWorksheetPaths)
		if err := f.validateCleanWorksheetXML(cleanWorksheetPaths); err != nil {
			return err
		}
		for path := range cleanWorksheetPaths {
			if _, ok := f.tempFiles.Load(path); ok {
				cleanTempWorksheetPaths[path] = true
			}
			if _, ok := f.streams[path]; ok {
				return fmt.Errorf("clean worksheet %q has stream data but no original XML source", path)
			}
		}
	}
	f.Pkg.Range(func(path, content interface{}) bool {
		pathStr := path.(string)
		if _, ok := f.streams[pathStr]; ok {
			return true
		}
		if preserveCleanWorksheets && cleanTempWorksheetPaths[pathStr] {
			return true
		}
		files = append(files, pathStr)
		return true
	})
	sort.Sort(sort.Reverse(sort.StringSlice(files)))
	for _, path := range files {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		content, _ := f.Pkg.Load(path)
		if n, err = fi.Write(content.([]byte)); int64(n) > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
		if err != nil {
			break
		}
	}
	if err != nil {
		return err
	}

	// Preserve files that were lazily extracted to temp files and never loaded
	// into Pkg. Large worksheet XML parts commonly live here; dropping them
	// during a non-destructive save corrupts untouched sheets.
	f.tempFiles.Range(func(path, content interface{}) bool {
		pathStr := path.(string)
		_, hasPkg := f.Pkg.Load(pathStr)
		if hasPkg && !(preserveCleanWorksheets && cleanTempWorksheetPaths[pathStr]) {
			return true
		}
		if _, ok := f.streams[pathStr]; ok {
			return true
		}
		tempFiles = append(tempFiles, pathStr)
		return true
	})
	sort.Sort(sort.Reverse(sort.StringSlice(tempFiles)))
	for _, path := range tempFiles {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		var file *os.File
		if file, err = f.readTemp(path); err != nil {
			break
		}
		written, copyErr := io.Copy(fi, file)
		closeErr := file.Close()
		if written > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
		if copyErr != nil {
			err = copyErr
			break
		}
		if closeErr != nil {
			err = closeErr
			break
		}
	}
	return err
}

// workSheetWriterNonDestructive writes worksheet XML WITHOUT modifying
// the original sheet.SheetData.Row arrays.
//
// Key differences from workSheetWriter():
// 1. Creates a DEEP COPY of each worksheet before trimming
// 2. Trims the COPY, not the original
// 3. Does NOT delete worksheets from memory
// 4. Preserves original state for subsequent operations
func (f *File) workSheetWriterNonDestructive(dirtyOpts *DirtySheetWriteOptions) error {
	var (
		arr     []byte
		buffer  = bytes.NewBuffer(arr)
		encoder = xml.NewEncoder(buffer)
		err     error
	)

	preserveCleanWorksheets := dirtyOpts != nil && dirtyOpts.PreserveCleanWorksheetXML
	dirtyWorksheetPaths := f.dirtyWorksheetPaths(dirtyOpts)
	f.Sheet.Range(func(p, ws interface{}) bool {
		if ws != nil {
			path := p.(string)
			if preserveCleanWorksheets && !dirtyWorksheetPaths[path] {
				return true
			}
			originalSheet := ws.(*xlsxWorksheet)

			// 🔒 Lock the worksheet to prevent concurrent modifications during copy
			originalSheet.mu.Lock()

			// 🔥 CRITICAL: Create a DEEP COPY for serialization
			// This ensures we don't modify the original!
			sheetCopy := f.deepCopyWorksheet(originalSheet)

			// 🔓 Unlock immediately after copy
			originalSheet.mu.Unlock()

			// Process merge cells on the copy
			if sheetCopy.MergeCells != nil && len(sheetCopy.MergeCells.Cells) > 0 {
				_ = sheetCopy.mergeOverlapCells()
			}

			// Process columns on the copy
			if sheetCopy.Cols != nil && len(sheetCopy.Cols.Col) > 0 {
				f.mergeExpandedCols(sheetCopy)
			}

			// 🔥 Trim rows on the COPY, not the original!
			// This is safe because sheetCopy is independent
			sheetCopy.SheetData.Row = trimRow(&sheetCopy.SheetData)

			// Add namespaces if needed
			if sheetCopy.SheetPr != nil || sheetCopy.Drawing != nil || sheetCopy.Hyperlinks != nil || sheetCopy.Picture != nil || sheetCopy.TableParts != nil {
				f.addNameSpaces(path, SourceRelationship)
			}

			// Handle alternate content
			if sheetCopy.DecodeAlternateContent != nil {
				sheetCopy.AlternateContent = &xlsxAlternateContent{
					Content: sheetCopy.DecodeAlternateContent.Content,
					XMLNSMC: SourceRelationshipCompatibility.Value,
				}
			}
			sheetCopy.DecodeAlternateContent = nil

			// Encode the COPY
			if err = encoder.Encode(sheetCopy); err != nil {
				return false
			}
			f.saveFileList(path, replaceRelationshipsBytes(f.replaceNameSpaceBytes(path, buffer.Bytes())))

			buffer.Reset()

			// ⚠️ IMPORTANT: We do NOT delete the original worksheet!
			// It stays in memory for subsequent operations
		}
		return true
	})
	return err
}

func (f *File) dirtyWorksheetPaths(dirtyOpts *DirtySheetWriteOptions) map[string]bool {
	paths := map[string]bool{}
	if dirtyOpts == nil || !dirtyOpts.PreserveCleanWorksheetXML {
		return paths
	}
	for dirtySheet, dirty := range dirtyOpts.DirtyWorksheets {
		if !dirty {
			continue
		}
		dirtySheet = strings.TrimSpace(dirtySheet)
		if dirtySheet == "" {
			continue
		}
		if isWorksheetXMLPath(dirtySheet) {
			paths[dirtySheet] = true
			continue
		}
		for sheetName, path := range f.sheetMap {
			if strings.EqualFold(sheetName, dirtySheet) || strings.EqualFold(path, dirtySheet) {
				paths[path] = true
				break
			}
		}
	}
	return paths
}

func (f *File) cleanWorksheetPaths(dirtyWorksheetPaths map[string]bool) map[string]bool {
	paths := map[string]bool{}
	for _, path := range f.sheetMap {
		if path == "" || dirtyWorksheetPaths[path] {
			continue
		}
		paths[path] = true
	}
	return paths
}

func (f *File) validateCleanWorksheetXML(cleanWorksheetPaths map[string]bool) error {
	for path := range cleanWorksheetPaths {
		if _, ok := f.tempFiles.Load(path); ok {
			continue
		}
		if content, ok := f.Pkg.Load(path); ok {
			if b, ok := content.([]byte); ok && len(b) > 0 {
				continue
			}
		}
		return fmt.Errorf("clean worksheet %q has no original XML source", path)
	}
	return nil
}

func isWorksheetXMLPath(path string) bool {
	path = strings.ToLower(strings.ReplaceAll(path, "\\", "/"))
	return strings.HasPrefix(path, "xl/worksheets/sheet") && strings.HasSuffix(path, ".xml")
}

// deepCopyWorksheet creates a deep copy of a worksheet structure.
// This ensures that modifications to the copy don't affect the original.
func (f *File) deepCopyWorksheet(original *xlsxWorksheet) *xlsxWorksheet {
	// Create new worksheet
	copy := &xlsxWorksheet{
		// Copy basic fields (pointers are OK for these as they won't be modified)
		SheetPr:                original.SheetPr,
		Dimension:              original.Dimension,
		SheetViews:             original.SheetViews,
		SheetFormatPr:          original.SheetFormatPr,
		Cols:                   nil, // deep copied below
		SheetCalcPr:            original.SheetCalcPr,
		SheetProtection:        original.SheetProtection,
		ProtectedRanges:        original.ProtectedRanges,
		Scenarios:              original.Scenarios,
		AutoFilter:             original.AutoFilter,
		SortState:              original.SortState,
		DataConsolidate:        original.DataConsolidate,
		CustomSheetViews:       original.CustomSheetViews,
		PhoneticPr:             original.PhoneticPr,
		ConditionalFormatting:  original.ConditionalFormatting,
		DataValidations:        original.DataValidations,
		Hyperlinks:             original.Hyperlinks,
		PrintOptions:           original.PrintOptions,
		PageMargins:            original.PageMargins,
		PageSetUp:              original.PageSetUp,
		HeaderFooter:           original.HeaderFooter,
		RowBreaks:              original.RowBreaks,
		ColBreaks:              original.ColBreaks,
		CustomProperties:       original.CustomProperties,
		CellWatches:            original.CellWatches,
		IgnoredErrors:          original.IgnoredErrors,
		SmartTags:              original.SmartTags,
		Drawing:                original.Drawing,
		LegacyDrawing:          original.LegacyDrawing,
		LegacyDrawingHF:        original.LegacyDrawingHF,
		DrawingHF:              original.DrawingHF,
		Picture:                original.Picture,
		OleObjects:             original.OleObjects,
		Controls:               original.Controls,
		WebPublishItems:        original.WebPublishItems,
		TableParts:             original.TableParts,
		ExtLst:                 original.ExtLst,
		AlternateContent:       original.AlternateContent,
		DecodeAlternateContent: original.DecodeAlternateContent,
		MergeCells:             original.MergeCells,
	}

	// 🔥 CRITICAL: Deep copy SheetData.Row
	// This is where trimRow() will operate, so we must copy the entire array
	if original.SheetData.Row != nil {
		copy.SheetData.Row = make([]xlsxRow, len(original.SheetData.Row))
		for i := range original.SheetData.Row {
			copy.SheetData.Row[i] = xlsxRow{
				R:            original.SheetData.Row[i].R,
				Spans:        original.SheetData.Row[i].Spans,
				Hidden:       original.SheetData.Row[i].Hidden,
				Ht:           original.SheetData.Row[i].Ht,
				CustomHeight: original.SheetData.Row[i].CustomHeight,
				OutlineLevel: original.SheetData.Row[i].OutlineLevel,
				Collapsed:    original.SheetData.Row[i].Collapsed,
				ThickTop:     original.SheetData.Row[i].ThickTop,
				ThickBot:     original.SheetData.Row[i].ThickBot,
				Ph:           original.SheetData.Row[i].Ph,
				S:            original.SheetData.Row[i].S,
				CustomFormat: original.SheetData.Row[i].CustomFormat,
			}

			// Deep copy cells array (this is critical!)
			// Check for nil to avoid panic in concurrent scenarios
			if original.SheetData.Row[i].C != nil {
				cellCount := len(original.SheetData.Row[i].C)
				copy.SheetData.Row[i].C = make([]xlsxC, cellCount)
				// Use explicit loop with bounds check to avoid race conditions
				for j := 0; j < cellCount && j < len(original.SheetData.Row[i].C); j++ {
					copy.SheetData.Row[i].C[j] = original.SheetData.Row[i].C[j]
				}
			}
		}
	}

	// Deep copy Cols to avoid shared slice mutations (e.g. sort.Slice in mergeExpandedCols)
	if original.Cols != nil {
		copy.Cols = &xlsxCols{}
		copy.Cols.Col = make([]xlsxCol, len(original.Cols.Col))
		for i, c := range original.Cols.Col {
			copy.Cols.Col[i] = c
			if c.Width != nil {
				w := *c.Width
				copy.Cols.Col[i].Width = &w
			}
		}
	}

	return copy
}

// SaveNonDestructive provides a function to save the spreadsheet to a file
// WITHOUT modifying internal state.
//
// This is equivalent to SaveAs() but uses WriteNonDestructive() internally.
func (f *File) SaveNonDestructive(name string, opts ...Options) error {
	if len(name) > MaxFilePathLength {
		return ErrMaxFilePathLength
	}
	f.Path = name
	file, err := os.OpenFile(filepath.Clean(name), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return f.WriteNonDestructive(file, opts...)
}
