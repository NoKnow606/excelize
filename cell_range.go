package excelize

import (
	"fmt"
	"runtime"
	"strconv"
	"sync"
)

// CellData contains value, style and formula for a cell
type CellData struct {
	Value   string
	Style   *Style
	Formula string
}

// GetRangeValuesConcurrent provides a function to get cell values from a range
// with concurrent processing. This is optimized for large ranges by splitting
// the range into chunks and processing them in parallel.
//
// The range format is "A1:Z1000". This function locks the worksheet once and
// then processes different row chunks concurrently.
//
// Example:
//
//	values, err := f.GetRangeValuesConcurrent("Sheet1", "A1:Z10000")
//	if err != nil {
//	    fmt.Println(err)
//	    return
//	}
//	for rowIdx, row := range values {
//	    for colIdx, value := range row {
//	        fmt.Printf("Cell [%d,%d]: %s\n", rowIdx, colIdx, value)
//	    }
//	}
func (f *File) GetRangeValuesConcurrent(sheet, rangeAddr string, opts ...Options) ([][]string, error) {
	// Parse range
	startCol, startRow, endCol, endRow, err := f.parseRange(rangeAddr)
	if err != nil {
		return nil, err
	}

	// Get shared strings first (sharedStringsReader has its own locking)
	sst, err := f.sharedStringsReader()
	if err != nil {
		return nil, err
	}

	// Get worksheet
	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	f.mu.Unlock()
	if err != nil {
		return nil, err
	}

	// Lock worksheet for reading
	ws.mu.Lock()
	defer ws.mu.Unlock()

	// Calculate dimensions
	numRows := endRow - startRow + 1
	numCols := endCol - startCol + 1
	results := make([][]string, numRows)
	for i := range results {
		results[i] = make([]string, numCols)
	}

	// Split into chunks by rows
	numWorkers := runtime.NumCPU()
	if numRows < numWorkers {
		numWorkers = numRows
	}
	rowsPerWorker := (numRows + numWorkers - 1) / numWorkers

	var wg sync.WaitGroup
	rawCellValue := f.getOptions(opts...).RawCellValue

	for i := 0; i < numWorkers; i++ {
		workerStartRow := startRow + i*rowsPerWorker
		workerEndRow := workerStartRow + rowsPerWorker - 1
		if workerEndRow > endRow {
			workerEndRow = endRow
		}
		if workerStartRow > endRow {
			break
		}

		wg.Add(1)
		go func(wStartRow, wEndRow int) {
			defer wg.Done()

			for row := wStartRow; row <= wEndRow; row++ {
				resultRowIdx := row - startRow
				for col := startCol; col <= endCol; col++ {
					resultColIdx := col - startCol
					cell, _ := CoordinatesToCellName(col, row)

					// Read cell value without additional locking
					value := f.getCellValueFromWorksheet(ws, sst, cell, rawCellValue)
					results[resultRowIdx][resultColIdx] = value
				}
			}
		}(workerStartRow, workerEndRow)
	}

	wg.Wait()
	return results, nil
}

// getCellValueFromWorksheet reads cell value from worksheet without locking.
// Assumes ws.mu is already locked by caller.
func (f *File) getCellValueFromWorksheet(ws *xlsxWorksheet, sst *xlsxSST, cell string, rawCellValue bool) string {
	_, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return ""
	}

	lastRowNum := 0
	if l := len(ws.SheetData.Row); l > 0 {
		lastRowNum = ws.SheetData.Row[l-1].R
	}

	if row > lastRowNum {
		return ""
	}

	// Binary search for row
	left, right := 0, len(ws.SheetData.Row)-1
	rowIdx := -1
	for left <= right {
		mid := (left + right) / 2
		if ws.SheetData.Row[mid].R == row {
			rowIdx = mid
			break
		}
		if ws.SheetData.Row[mid].R < row {
			left = mid + 1
		} else {
			right = mid - 1
		}
	}

	if rowIdx == -1 {
		return ""
	}

	// Find cell in row
	rowData := ws.SheetData.Row[rowIdx]
	for colIdx := range rowData.C {
		colData := &rowData.C[colIdx]
		if cell == colData.R {
			// Handle shared string
			if colData.T == "s" && sst != nil {
				if xlsxSI := 0; colData.V != "" {
					if v, err := strconv.Atoi(colData.V); err == nil {
						xlsxSI = v
					}
					if xlsxSI >= 0 && xlsxSI < len(sst.SI) {
						return sst.SI[xlsxSI].String()
					}
				}
				return colData.V
			}
			// Return direct value
			return colData.V
		}
	}

	return ""
}

// parseRange parses a range address like "A1:Z1000" and returns
// startCol, startRow, endCol, endRow
func (f *File) parseRange(rangeAddr string) (int, int, int, int, error) {
	// Split by ":"
	var startCell, endCell string
	for i, c := range rangeAddr {
		if c == ':' {
			startCell = rangeAddr[:i]
			endCell = rangeAddr[i+1:]
			break
		}
	}

	if startCell == "" || endCell == "" {
		return 0, 0, 0, 0, fmt.Errorf("invalid range format: %s", rangeAddr)
	}

	startCol, startRow, err := CellNameToCoordinates(startCell)
	if err != nil {
		return 0, 0, 0, 0, err
	}

	endCol, endRow, err := CellNameToCoordinates(endCell)
	if err != nil {
		return 0, 0, 0, 0, err
	}

	if startCol > endCol || startRow > endRow {
		return 0, 0, 0, 0, fmt.Errorf("invalid range: start must be before end")
	}

	return startCol, startRow, endCol, endRow, nil
}

// GetRangeDataConcurrent provides a function to get cell values, styles and formulas
// from a range with concurrent processing. This is optimized for large ranges by
// splitting the range into chunks and processing them in parallel.
//
// The range format is "A1:Z1000". This function locks the worksheet once and then
// processes different row chunks concurrently.
//
// Example:
//
//	data, err := f.GetRangeDataConcurrent("Sheet1", "A1:Z10000")
//	if err != nil {
//	    fmt.Println(err)
//	    return
//	}
//	for rowIdx, row := range data {
//	    for colIdx, cell := range row {
//	        fmt.Printf("Cell [%d,%d]: value=%s, formula=%s\n",
//	            rowIdx, colIdx, cell.Value, cell.Formula)
//	        if cell.Style != nil {
//	            fmt.Printf("  Style: fill=%v, font=%v\n", cell.Style.Fill, cell.Style.Font)
//	        }
//	    }
//	}
func (f *File) GetRangeDataConcurrent(sheet, rangeAddr string, opts ...Options) ([][]CellData, error) {
	// Parse range
	startCol, startRow, endCol, endRow, err := f.parseRange(rangeAddr)
	if err != nil {
		return nil, err
	}

	// Get shared strings first
	sst, err := f.sharedStringsReader()
	if err != nil {
		return nil, err
	}

	// Get worksheet and styles
	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		f.mu.Unlock()
		return nil, err
	}
	s, err := f.stylesReader()
	if err != nil {
		f.mu.Unlock()
		return nil, err
	}
	f.mu.Unlock()

	// Lock worksheet for reading
	ws.mu.Lock()
	defer ws.mu.Unlock()

	// Calculate dimensions
	numRows := endRow - startRow + 1
	numCols := endCol - startCol + 1
	results := make([][]CellData, numRows)
	for i := range results {
		results[i] = make([]CellData, numCols)
	}

	// Split into chunks by rows
	numWorkers := runtime.NumCPU()
	if numRows < numWorkers {
		numWorkers = numRows
	}
	rowsPerWorker := (numRows + numWorkers - 1) / numWorkers

	var wg sync.WaitGroup
	rawCellValue := f.getOptions(opts...).RawCellValue

	// Shared style cache (concurrent-safe)
	var styleCache sync.Map

	for i := 0; i < numWorkers; i++ {
		workerStartRow := startRow + i*rowsPerWorker
		workerEndRow := workerStartRow + rowsPerWorker - 1
		if workerEndRow > endRow {
			workerEndRow = endRow
		}
		if workerStartRow > endRow {
			break
		}

		wg.Add(1)
		go func(wStartRow, wEndRow int) {
			defer wg.Done()

			for row := wStartRow; row <= wEndRow; row++ {
				resultRowIdx := row - startRow
				for col := startCol; col <= endCol; col++ {
					resultColIdx := col - startCol
					cell, _ := CoordinatesToCellName(col, row)

					// Read cell data without additional locking
					cellData := f.getCellDataFromWorksheet(ws, sst, s, cell, rawCellValue, &styleCache)
					results[resultRowIdx][resultColIdx] = cellData
				}
			}
		}(workerStartRow, workerEndRow)
	}

	wg.Wait()
	return results, nil
}

// getCellDataFromWorksheet reads cell value, style and formula from worksheet without locking.
// Assumes ws.mu is already locked by caller.
func (f *File) getCellDataFromWorksheet(ws *xlsxWorksheet, sst *xlsxSST, s *xlsxStyleSheet, cell string, rawCellValue bool, styleCache *sync.Map) CellData {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return CellData{}
	}

	lastRowNum := 0
	if l := len(ws.SheetData.Row); l > 0 {
		lastRowNum = ws.SheetData.Row[l-1].R
	}

	if row > lastRowNum {
		return CellData{}
	}

	// Binary search for row
	left, right := 0, len(ws.SheetData.Row)-1
	rowIdx := -1
	for left <= right {
		mid := (left + right) / 2
		if ws.SheetData.Row[mid].R == row {
			rowIdx = mid
			break
		}
		if ws.SheetData.Row[mid].R < row {
			left = mid + 1
		} else {
			right = mid - 1
		}
	}

	if rowIdx == -1 {
		return CellData{}
	}

	// Find cell in row
	rowData := ws.SheetData.Row[rowIdx]
	for colIdx := range rowData.C {
		colData := &rowData.C[colIdx]
		if cell == colData.R {
			var value string
			// Handle shared string
			if colData.T == "s" && sst != nil {
				if xlsxSI := 0; colData.V != "" {
					if v, err := strconv.Atoi(colData.V); err == nil {
						xlsxSI = v
					}
					if xlsxSI >= 0 && xlsxSI < len(sst.SI) {
						value = sst.SI[xlsxSI].String()
					}
				}
				if value == "" {
					value = colData.V
				}
			} else {
				value = colData.V
			}

			// Get style
			styleID := ws.prepareCellStyle(col, row, colData.S)
			var style *Style

			// Check cache first (lock-free read)
			if cached, ok := styleCache.Load(styleID); ok {
				style = cached.(*Style)
			} else {
				// Parse style
				if styleID >= 0 && s.CellXfs != nil && len(s.CellXfs.Xf) > styleID {
					style = &Style{}
					xf := s.CellXfs.Xf[styleID]
					if extractStyleCondFuncs["fill"](xf, s) {
						f.extractFills(s.Fills.Fill[*xf.FillID], s, style)
					}
					if extractStyleCondFuncs["border"](xf, s) {
						f.extractBorders(s.Borders.Border[*xf.BorderID], s, style)
					}
					if extractStyleCondFuncs["font"](xf, s) {
						style.Font = extractFont(s.Fonts.Font[*xf.FontID])
					}
					if extractStyleCondFuncs["alignment"](xf, s) {
						f.extractAlignment(xf.Alignment, s, style)
					}
					if extractStyleCondFuncs["protection"](xf, s) {
						f.extractProtection(xf.Protection, s, style)
					}
					f.extractNumFmt(xf.NumFmtID, s, style)

					// Cache the parsed style (handles concurrent writes)
					styleCache.LoadOrStore(styleID, style)
				}
			}

			var formula string
			if colData.F != nil && colData.F.Content != "" {
				formula = "=" + colData.F.Content
			}

			return CellData{
				Value:   value,
				Style:   style,
				Formula: formula,
			}
		}
	}

	return CellData{}
}
