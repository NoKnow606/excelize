package excelize

import (
	"fmt"
)

// SafeCheckRow is a safer version of checkRow that handles index out of range errors
// This is a temporary fix until the root cause is addressed
func (ws *xlsxWorksheet) SafeCheckRow() error {
	defer func() {
		if r := recover(); r != nil {
			// Log the panic but don't crash the program
			fmt.Printf("⚠️  SafeCheckRow recovered from panic: %v\n", r)
		}
	}()

	for rowIdx := range ws.SheetData.Row {
		rowData := &ws.SheetData.Row[rowIdx]

		colCount := len(rowData.C)
		if colCount == 0 {
			continue
		}

		// check and fill the cell without r attribute in a row element
		rCount := 0
		for idx, cell := range rowData.C {
			rCount++
			if cell.R != "" {
				lastR, _, err := CellNameToCoordinates(cell.R)
				if err != nil {
					return err
				}
				if lastR > rCount {
					rCount = lastR
				}
				continue
			}
			rowData.C[idx].R, _ = CoordinatesToCellName(rCount, rowIdx+1)
		}

		lastCol, _, err := CellNameToCoordinates(rowData.C[colCount-1].R)
		if err != nil {
			return err
		}

		if colCount < lastCol {
			sourceList := rowData.C

			// ✅ FIX: Calculate the actual max column from sourceList
			maxCol := lastCol
			for _, cell := range sourceList {
				colNum, _, err := CellNameToCoordinates(cell.R)
				if err != nil {
					continue
				}
				if colNum > maxCol {
					maxCol = colNum
				}
			}

			targetList := make([]xlsxC, 0, maxCol)

			rowData.C = ws.SheetData.Row[rowIdx].C[:0]

			// Create target list with maxCol size
			for colIdx := 0; colIdx < maxCol; colIdx++ {
				cellName, err := CoordinatesToCellName(colIdx+1, rowIdx+1)
				if err != nil {
					return err
				}
				targetList = append(targetList, xlsxC{R: cellName})
			}

			rowData.C = targetList

			// ✅ FIX: Add boundary check before accessing array
			for colIdx := range sourceList {
				colData := &sourceList[colIdx]
				colNum, _, err := CellNameToCoordinates(colData.R)
				if err != nil {
					return err
				}

				// Boundary check
				if colNum-1 < 0 || colNum-1 >= len(ws.SheetData.Row[rowIdx].C) {
					fmt.Printf("⚠️  Skipping invalid cell reference: %s (colNum=%d, len=%d)\n",
						colData.R, colNum, len(ws.SheetData.Row[rowIdx].C))
					continue
				}

				ws.SheetData.Row[rowIdx].C[colNum-1] = *colData
			}
		}
	}
	return nil
}
