package excelize

import (
	"encoding/xml"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestRowsXMLMarshalNoPanic verifies that xml.Marshal doesn't panic after checkRow
func TestRowsXMLMarshalNoPanic(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Create sparse data that triggered the original bug
	f.SetCellValue(sheetName, "A1", "Data1")
	f.SetCellValue(sheetName, "X1", "Data24")
	f.SetCellValue(sheetName, "AA1", "Data27")

	// Get worksheet
	ws, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded)
	worksheet := ws.(*xlsxWorksheet)

	fmt.Println("\n=== 测试 xml.Marshal 不 panic ===")

	// Run checkRow
	err := worksheet.checkRow()
	assert.NoError(t, err)

	// ✅ This should not panic
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("xml.Marshal panicked: %v", r)
		}
	}()

	output, err := xml.Marshal(worksheet)
	assert.NoError(t, err)
	assert.NotEmpty(t, output)

	fmt.Printf("✅ xml.Marshal succeeded, output size: %d bytes\n", len(output))
}

// TestRowsGetRowsNoPanic verifies that GetRows doesn't panic
func TestRowsGetRowsNoPanic(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create sparse data
	f.SetCellValue(sheetName, "A1", "Data1")
	f.SetCellValue(sheetName, "Z1", "Data26")

	fmt.Println("\n=== 测试 GetRows 不 panic ===")

	// ✅ This should not panic
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("GetRows panicked: %v", r)
		}
	}()

	rows, err := f.GetRows(sheetName)
	assert.NoError(t, err)
	assert.NotEmpty(t, rows)

	fmt.Printf("✅ GetRows succeeded, got %d rows\n", len(rows))
	fmt.Printf("Row 1: %v\n", rows[0])
}

// TestRowsIteratorNoPanic verifies that Rows iterator doesn't panic
func TestRowsIteratorNoPanic(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create sparse data
	for i := 1; i <= 10; i++ {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", i), fmt.Sprintf("Row%d", i))
		f.SetCellValue(sheetName, fmt.Sprintf("Z%d", i), fmt.Sprintf("Col26-%d", i))
	}

	fmt.Println("\n=== 测试 Rows 迭代器不 panic ===")

	// ✅ This should not panic
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("Rows iterator panicked: %v", r)
		}
	}()

	rows, err := f.Rows(sheetName)
	assert.NoError(t, err)

	rowCount := 0
	for rows.Next() {
		row, err := rows.Columns()
		assert.NoError(t, err)
		rowCount++
		if rowCount <= 3 {
			fmt.Printf("Row %d: %d columns\n", rowCount, len(row))
		}
	}

	err = rows.Close()
	assert.NoError(t, err)

	fmt.Printf("✅ Rows iterator succeeded, processed %d rows\n", rowCount)
}

// TestCheckRowWithInsertRowsThenGetRows verifies the full flow
func TestCheckRowWithInsertRowsThenGetRows(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Setup
	f.SetCellValue(sheetName, "A1", "Header1")
	f.SetCellValue(sheetName, "Z1", "Header26")
	f.SetCellValue(sheetName, "A2", "Data1")
	f.SetCellValue(sheetName, "Z2", "Data26")

	fmt.Println("\n=== 完整流程：InsertRows + GetRows ===")

	// Insert rows
	err := f.InsertRows(sheetName, 2, 1)
	assert.NoError(t, err)
	fmt.Println("✅ InsertRows succeeded")

	// ✅ This should not panic
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("GetRows panicked after InsertRows: %v", r)
		}
	}()

	// Get rows
	rows, err := f.GetRows(sheetName)
	assert.NoError(t, err)
	fmt.Printf("✅ GetRows succeeded, got %d rows\n", len(rows))

	// Verify data
	assert.GreaterOrEqual(t, len(rows), 3)
	if len(rows) >= 3 {
		fmt.Printf("Row 1: %d columns\n", len(rows[0]))
		fmt.Printf("Row 2: %d columns (inserted)\n", len(rows[1]))
		fmt.Printf("Row 3: %d columns\n", len(rows[2]))

		assert.Equal(t, "Header1", rows[0][0])
		assert.Equal(t, "Data1", rows[2][0])
	}
}

// TestCheckRowWithBatchUpdateThenGetRows verifies batch update doesn't break GetRows
func TestCheckRowWithBatchUpdateThenGetRows(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Setup
	f.SetCellValue(sheetName, "A1", "Initial")

	fmt.Println("\n=== BatchUpdate → GetRows ===")

	// Batch update with wide columns
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A1", Value: "Updated"},
		{Sheet: sheetName, Cell: "Z1", Value: "Col26"},
		{Sheet: sheetName, Cell: "AA1", Value: "Col27"},
	}
	err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)
	fmt.Println("✅ BatchUpdate succeeded")

	// ✅ This should not panic
	defer func() {
		if r := recover(); r != nil {
			t.Errorf("GetRows panicked after BatchUpdate: %v", r)
		}
	}()

	// Get rows
	rows, err := f.GetRows(sheetName)
	assert.NoError(t, err)
	fmt.Printf("✅ GetRows succeeded, got %d rows\n", len(rows))

	// Verify data
	assert.GreaterOrEqual(t, len(rows), 1)
	if len(rows) >= 1 {
		fmt.Printf("Row 1: %d columns\n", len(rows[0]))
		assert.Equal(t, "Updated", rows[0][0])
	}
}
