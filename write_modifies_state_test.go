package excelize

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestWriteModifiesWorksheetState verifies that f.Write() modifies worksheet internal state
func TestWriteModifiesWorksheetState(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup: Insert row at 100 and set value
	f.SetCellValue(sheetName, "A1", "Header")
	err := f.InsertRows(sheetName, 100, 1)
	assert.NoError(t, err)

	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A100", Value: "Data100"},
		{Sheet: sheetName, Cell: "B100", Value: 999},
	}
	_, err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify before Write
	fmt.Println("\n=== 写入前 ===")
	a100Before, _ := f.GetCellValue(sheetName, "A100")
	b100Before, _ := f.GetCellValue(sheetName, "B100")
	fmt.Printf("A100 = '%s', B100 = '%s'\n", a100Before, b100Before)

	// Get worksheet pointer and row count before Write
	wsBefore, _ := f.Sheet.Load(sheetXMLPath)
	ws := wsBefore.(*xlsxWorksheet)
	rowCountBefore := len(ws.SheetData.Row)
	fmt.Printf("Row count before Write: %d\n", rowCountBefore)

	// Call Write (with KeepWorksheetInMemory)
	var buf bytes.Buffer
	err = f.Write(&buf)
	assert.NoError(t, err)

	// Check if worksheet is still in memory
	wsAfter, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded, "Worksheet should still be in memory with KeepWorksheetInMemory")

	// Compare pointers
	ptrBefore := fmt.Sprintf("%p", wsBefore)
	ptrAfter := fmt.Sprintf("%p", wsAfter)
	fmt.Printf("\nPointer before: %s\n", ptrBefore)
	fmt.Printf("Pointer after:  %s\n", ptrAfter)
	assert.Equal(t, ptrBefore, ptrAfter, "Should be same worksheet instance")

	// ⚠️ But check if internal state changed
	wsModified := wsAfter.(*xlsxWorksheet)
	rowCountAfter := len(wsModified.SheetData.Row)
	fmt.Printf("\nRow count after Write: %d\n", rowCountAfter)

	if rowCountBefore != rowCountAfter {
		fmt.Printf("⚠️  WARNING: Row count changed! Before=%d, After=%d\n", rowCountBefore, rowCountAfter)
		fmt.Printf("   trimRow() modified the SheetData.Row array!\n")
	}

	// Try to get values after Write
	fmt.Println("\n=== 写入后 ===")
	a100After, _ := f.GetCellValue(sheetName, "A100")
	b100After, _ := f.GetCellValue(sheetName, "B100")
	fmt.Printf("A100 = '%s', B100 = '%s'\n", a100After, b100After)

	// Values should still be accessible
	assert.Equal(t, "Data100", a100After, "A100 should still be accessible after Write")
	assert.Equal(t, "999", b100After, "B100 should still be accessible after Write")
}

// TestWriteTrimsEmptyRows tests that trimRow() removes empty trailing rows
func TestWriteTrimsEmptyRows(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Create sparse data: row 1, row 5, row 100
	f.SetCellValue(sheetName, "A1", "Row1")
	f.SetCellValue(sheetName, "A5", "Row5")
	f.SetCellValue(sheetName, "A100", "Row100")

	// Get worksheet state before Write
	wsBefore, _ := f.Sheet.Load(sheetXMLPath)
	ws := wsBefore.(*xlsxWorksheet)

	fmt.Println("\n=== 写入前的 Row 结构 ===")
	fmt.Printf("Total rows in SheetData.Row: %d\n", len(ws.SheetData.Row))
	for i, row := range ws.SheetData.Row {
		if len(row.C) > 0 {
			fmt.Printf("Row[%d]: R=%d, cells=%d\n", i, row.R, len(row.C))
		}
	}

	// Write
	var buf bytes.Buffer
	f.Write(&buf)

	// Get worksheet state after Write
	wsAfter, _ := f.Sheet.Load(sheetXMLPath)
	wsModified := wsAfter.(*xlsxWorksheet)

	fmt.Println("\n=== 写入后的 Row 结构 ===")
	fmt.Printf("Total rows in SheetData.Row: %d\n", len(wsModified.SheetData.Row))
	for i, row := range wsModified.SheetData.Row {
		if len(row.C) > 0 {
			fmt.Printf("Row[%d]: R=%d, cells=%d\n", i, row.R, len(row.C))
		}
	}

	// Verify data is still accessible
	a1, _ := f.GetCellValue(sheetName, "A1")
	a5, _ := f.GetCellValue(sheetName, "A5")
	a100, _ := f.GetCellValue(sheetName, "A100")

	fmt.Println("\n=== 数据可访问性 ===")
	fmt.Printf("A1 = '%s'\n", a1)
	fmt.Printf("A5 = '%s'\n", a5)
	fmt.Printf("A100 = '%s'\n", a100)

	assert.Equal(t, "Row1", a1)
	assert.Equal(t, "Row5", a5)
	assert.Equal(t, "Row100", a100)
}

// TestMultipleWrites tests calling Write multiple times
func TestMultipleWritesModifyState(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Set initial data
	f.SetCellValue(sheetName, "A1", "Data")

	// Get initial worksheet
	ws1, _ := f.Sheet.Load(sheetXMLPath)
	ptr1 := fmt.Sprintf("%p", ws1)

	fmt.Println("\n=== 多次 Write 测试 ===")
	fmt.Printf("Initial pointer: %s\n", ptr1)

	// Write 3 times
	for i := 1; i <= 3; i++ {
		var buf bytes.Buffer
		err := f.Write(&buf)
		assert.NoError(t, err)

		ws, loaded := f.Sheet.Load(sheetXMLPath)
		assert.True(t, loaded)
		ptr := fmt.Sprintf("%p", ws)

		wsObj := ws.(*xlsxWorksheet)
		rowCount := len(wsObj.SheetData.Row)

		fmt.Printf("After Write #%d: ptr=%s, rows=%d, same=%v\n",
			i, ptr, rowCount, ptr == ptr1)
	}

	// Verify data still accessible
	a1, _ := f.GetCellValue(sheetName, "A1")
	assert.Equal(t, "Data", a1)
	fmt.Printf("\n✅ A1 = '%s' (still accessible)\n", a1)
}

// TestWriteWithInsertRowsAndBatchUpdate tests the full scenario
func TestWriteWithInsertRowsAndBatchUpdate(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer f.Close()

	sheetName := "Sheet1"

	fmt.Println("\n=== 完整场景测试 ===")

	// Step 1: InsertRows
	f.SetCellValue(sheetName, "A1", "Header")
	err := f.InsertRows(sheetName, 2, 1)
	assert.NoError(t, err)
	fmt.Println("✅ Step 1: InsertRows")

	// Step 2: BatchUpdate
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A2", Value: "Inserted"},
		{Sheet: sheetName, Cell: "B2", Value: 123},
	}
	_, err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)
	fmt.Println("✅ Step 2: BatchUpdate")

	// Step 3: Write
	var buf bytes.Buffer
	err = f.Write(&buf)
	assert.NoError(t, err)
	fmt.Println("✅ Step 3: Write")

	// Step 4: GetCellValue after Write
	a2, _ := f.GetCellValue(sheetName, "A2")
	b2, _ := f.GetCellValue(sheetName, "B2")
	fmt.Printf("✅ Step 4: GetCellValue - A2='%s', B2='%s'\n", a2, b2)

	assert.Equal(t, "Inserted", a2)
	assert.Equal(t, "123", b2)

	// Step 5: Another BatchUpdate after Write
	updates2 := []CellUpdate{
		{Sheet: sheetName, Cell: "C2", Value: "AfterWrite"},
	}
	_, err = f.BatchUpdateAndRecalculate(updates2)
	assert.NoError(t, err)
	fmt.Println("✅ Step 5: BatchUpdate after Write")

	// Step 6: GetCellValue again
	c2, _ := f.GetCellValue(sheetName, "C2")
	fmt.Printf("✅ Step 6: C2='%s'\n", c2)
	assert.Equal(t, "AfterWrite", c2)
}
