package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestCheckRowIndexOutOfRange tests the fix for index out of range panic
func TestCheckRowIndexOutOfRange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup: Create a scenario that might trigger the bug
	// Set cells with gaps in columns
	f.SetCellValue(sheetName, "A1", "Col1")
	f.SetCellValue(sheetName, "X1", "Col24") // Column 24

	// Get worksheet
	ws, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded)

	worksheet := ws.(*xlsxWorksheet)

	// Before fix: This might panic
	// After fix: Should handle gracefully
	fmt.Println("\n=== Testing checkRow with sparse columns ===")
	err := worksheet.checkRow()
	assert.NoError(t, err, "checkRow should not panic or return error")

	// Verify data is still accessible
	a1, _ := f.GetCellValue(sheetName, "A1")
	x1, _ := f.GetCellValue(sheetName, "X1")

	fmt.Printf("A1 = '%s'\n", a1)
	fmt.Printf("X1 = '%s'\n", x1)

	assert.Equal(t, "Col1", a1)
	assert.Equal(t, "Col24", x1)
}

// TestCheckRowWithCorruptedData tests checkRow with potentially corrupted data
func TestCheckRowWithCorruptedData(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Create sparse data that might cause issues
	tests := []struct {
		cell  string
		value interface{}
	}{
		{"A1", "Data1"},
		{"B1", "Data2"},
		{"Z1", "Data26"}, // Column 26
		{"A2", "Row2"},
		{"AA2", "Col27"}, // Column 27
	}

	for _, tt := range tests {
		f.SetCellValue(sheetName, tt.cell, tt.value)
	}

	// Get worksheet
	ws, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded)
	worksheet := ws.(*xlsxWorksheet)

	fmt.Println("\n=== Testing checkRow with wide sparse data ===")

	// This should not panic
	err := worksheet.checkRow()
	assert.NoError(t, err)

	// Verify all data is accessible
	for _, tt := range tests {
		val, err := f.GetCellValue(sheetName, tt.cell)
		assert.NoError(t, err)
		assert.Equal(t, fmt.Sprint(tt.value), val, "Cell %s should match", tt.cell)
		fmt.Printf("%s = '%s' ✅\n", tt.cell, val)
	}
}

// TestCheckRowAfterInsertRows tests checkRow after InsertRows operation
func TestCheckRowAfterInsertRows(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup initial data
	f.SetCellValue(sheetName, "A1", "Header1")
	f.SetCellValue(sheetName, "X1", "Header24")
	f.SetCellValue(sheetName, "A2", "Data1")
	f.SetCellValue(sheetName, "X2", "Data24")

	fmt.Println("\n=== Testing checkRow after InsertRows ===")

	// Insert rows
	err := f.InsertRows(sheetName, 2, 1)
	assert.NoError(t, err)

	// Get worksheet
	ws, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded)
	worksheet := ws.(*xlsxWorksheet)

	// Run checkRow
	err = worksheet.checkRow()
	assert.NoError(t, err, "checkRow should not panic after InsertRows")

	// Verify data
	a1, _ := f.GetCellValue(sheetName, "A1")
	x1, _ := f.GetCellValue(sheetName, "X1")
	a3, _ := f.GetCellValue(sheetName, "A3") // Original row 2 moved to row 3
	x3, _ := f.GetCellValue(sheetName, "X3")

	fmt.Printf("A1 = '%s'\n", a1)
	fmt.Printf("X1 = '%s'\n", x1)
	fmt.Printf("A3 = '%s' (moved from A2)\n", a3)
	fmt.Printf("X3 = '%s' (moved from X2)\n", x3)

	assert.Equal(t, "Header1", a1)
	assert.Equal(t, "Header24", x1)
	assert.Equal(t, "Data1", a3)
	assert.Equal(t, "Data24", x3)
}

// TestCheckRowWithWriteAndReload tests checkRow after Write operation
func TestCheckRowWithWriteAndReload(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup sparse data
	f.SetCellValue(sheetName, "A1", "Data")
	f.SetCellValue(sheetName, "Z1", "Col26")

	fmt.Println("\n=== Testing checkRow after Write ===")

	// Write (this calls trimRow and may modify internal state)
	tmpFile := "test_checkrow_write.xlsx"
	err := f.SaveAs(tmpFile)
	assert.NoError(t, err)

	// Worksheet should still be in memory
	ws, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded)
	worksheet := ws.(*xlsxWorksheet)

	// Run checkRow again
	err = worksheet.checkRow()
	assert.NoError(t, err, "checkRow should not panic after Write")

	// Verify data
	a1, _ := f.GetCellValue(sheetName, "A1")
	z1, _ := f.GetCellValue(sheetName, "Z1")

	fmt.Printf("A1 = '%s'\n", a1)
	fmt.Printf("Z1 = '%s'\n", z1)

	assert.Equal(t, "Data", a1)
	assert.Equal(t, "Col26", z1)

	// Cleanup
	f2, _ := OpenFile(tmpFile)
	if f2 != nil {
		f2.Close()
	}
}

// TestCheckRowMultipleTimes tests calling checkRow multiple times
func TestCheckRowMultipleTimes(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup
	f.SetCellValue(sheetName, "A1", "Data")
	f.SetCellValue(sheetName, "Z1", "Col26")

	ws, _ := f.Sheet.Load(sheetXMLPath)
	worksheet := ws.(*xlsxWorksheet)

	fmt.Println("\n=== Testing multiple checkRow calls ===")

	// Call checkRow multiple times
	for i := 1; i <= 3; i++ {
		err := worksheet.checkRow()
		assert.NoError(t, err, "checkRow call #%d should succeed", i)
		fmt.Printf("✅ checkRow call #%d succeeded\n", i)
	}

	// Verify data is still intact
	a1, _ := f.GetCellValue(sheetName, "A1")
	z1, _ := f.GetCellValue(sheetName, "Z1")

	assert.Equal(t, "Data", a1)
	assert.Equal(t, "Col26", z1)
}
