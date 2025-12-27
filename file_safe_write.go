// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
)

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

// WriteToNonDestructive implements io.WriterTo to write the file without
// modifying internal state. Returns the number of bytes written.
func (f *File) WriteToNonDestructive(w io.Writer, opts ...Options) (int64, error) {
	for i := range opts {
		f.options = &opts[i]
	}
	buf, err := f.WriteToBufferNonDestructive()
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
	buf := new(bytes.Buffer)
	zw := f.ZipWriter(buf)

	if err := f.writeToZipNonDestructive(zw); err != nil {
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
func (f *File) writeToZipNonDestructive(zw ZipWriter) error {
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
	f.workSheetWriterNonDestructive()

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

	// Write remaining files (unchanged from original)
	var (
		err   error
		files []string
	)
	f.Pkg.Range(func(path, content interface{}) bool {
		if _, ok := f.streams[path.(string)]; ok {
			return true
		}
		files = append(files, path.(string))
		return true
	})
	for _, path := range files {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		content, _ := f.Pkg.Load(path)
		if _, err = fi.Write(content.([]byte)); err != nil {
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
func (f *File) workSheetWriterNonDestructive() {
	var (
		arr     []byte
		buffer  = bytes.NewBuffer(arr)
		encoder = xml.NewEncoder(buffer)
	)

	f.Sheet.Range(func(p, ws interface{}) bool {
		if ws != nil {
			originalSheet := ws.(*xlsxWorksheet)

			// 🔥 CRITICAL: Create a DEEP COPY for serialization
			// This ensures we don't modify the original!
			sheetCopy := f.deepCopyWorksheet(originalSheet)

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
				f.addNameSpaces(p.(string), SourceRelationship)
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
			_ = encoder.Encode(sheetCopy)
			f.saveFileList(p.(string), replaceRelationshipsBytes(f.replaceNameSpaceBytes(p.(string), buffer.Bytes())))

			buffer.Reset()

			// ⚠️ IMPORTANT: We do NOT delete the original worksheet!
			// It stays in memory for subsequent operations
		}
		return true
	})
}

// deepCopyWorksheet creates a deep copy of a worksheet structure.
// This ensures that modifications to the copy don't affect the original.
func (f *File) deepCopyWorksheet(original *xlsxWorksheet) *xlsxWorksheet {
	// Create new worksheet
	copy := &xlsxWorksheet{
		// Copy basic fields (pointers are OK for these as they won't be modified)
		SheetPr:               original.SheetPr,
		Dimension:             original.Dimension,
		SheetViews:            original.SheetViews,
		SheetFormatPr:         original.SheetFormatPr,
		Cols:                  original.Cols,
		SheetCalcPr:           original.SheetCalcPr,
		SheetProtection:       original.SheetProtection,
		ProtectedRanges:       original.ProtectedRanges,
		Scenarios:             original.Scenarios,
		AutoFilter:            original.AutoFilter,
		SortState:             original.SortState,
		DataConsolidate:       original.DataConsolidate,
		CustomSheetViews:      original.CustomSheetViews,
		PhoneticPr:            original.PhoneticPr,
		ConditionalFormatting: original.ConditionalFormatting,
		DataValidations:       original.DataValidations,
		Hyperlinks:            original.Hyperlinks,
		PrintOptions:          original.PrintOptions,
		PageMargins:           original.PageMargins,
		PageSetUp:             original.PageSetUp,
		HeaderFooter:          original.HeaderFooter,
		RowBreaks:             original.RowBreaks,
		ColBreaks:             original.ColBreaks,
		CustomProperties:      original.CustomProperties,
		CellWatches:           original.CellWatches,
		IgnoredErrors:         original.IgnoredErrors,
		SmartTags:             original.SmartTags,
		Drawing:               original.Drawing,
		LegacyDrawing:         original.LegacyDrawing,
		LegacyDrawingHF:       original.LegacyDrawingHF,
		DrawingHF:             original.DrawingHF,
		Picture:               original.Picture,
		OleObjects:            original.OleObjects,
		Controls:              original.Controls,
		WebPublishItems:       original.WebPublishItems,
		TableParts:            original.TableParts,
		ExtLst:                original.ExtLst,
		AlternateContent:      original.AlternateContent,
		DecodeAlternateContent: original.DecodeAlternateContent,
		MergeCells:            original.MergeCells,
	}

	// 🔥 CRITICAL: Deep copy SheetData.Row
	// This is where trimRow() will operate, so we must copy the entire array
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
		copy.SheetData.Row[i].C = make([]xlsxC, len(original.SheetData.Row[i].C))
		for j := range original.SheetData.Row[i].C {
			copy.SheetData.Row[i].C[j] = original.SheetData.Row[i].C[j]
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
