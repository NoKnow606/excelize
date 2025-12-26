package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestInsertRowBatchUpdateThenGet reproduces the issue: insert row, batch update, then get values
func TestInsertRowBatchUpdateThenGet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Setup: Create initial data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Header1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", "Data1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", "Data2"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "Header2"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B2", 100))
	assert.NoError(t, f.SetCellValue("Sheet1", "B3", 200))

	fmt.Println("=== 初始数据 ===")
	for i := 1; i <= 3; i++ {
		a, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		b, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i))
		fmt.Printf("Row %d: A=%s, B=%s\n", i, a, b)
	}

	// Step 1: Insert a row at position 3
	fmt.Println("\n=== 插入第 3 行 ===")
	err := f.InsertRows("Sheet1", 3, 1)
	assert.NoError(t, err)

	fmt.Println("\n=== 插入后的数据 ===")
	for i := 1; i <= 4; i++ {
		a, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		b, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i))
		fmt.Printf("Row %d: A=%s, B=%s\n", i, a, b)
	}

	// Step 2: Batch update the newly inserted row (row 3)
	fmt.Println("\n=== 批量更新第 3 行 ===")
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A3", Value: "NewData"},
		{Sheet: "Sheet1", Cell: "B3", Value: 999},
		{Sheet: "Sheet1", Cell: "C3", Value: "Extra"},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)
	fmt.Printf("受影响的单元格: %d\n", len(affected))

	// Step 3: Try to get the values we just set
	fmt.Println("\n=== 获取第 3 行的值 ===")
	a3, err := f.GetCellValue("Sheet1", "A3")
	assert.NoError(t, err)
	fmt.Printf("A3 = '%s' (expected: 'NewData')\n", a3)

	b3, err := f.GetCellValue("Sheet1", "B3")
	assert.NoError(t, err)
	fmt.Printf("B3 = '%s' (expected: '999')\n", b3)

	c3, err := f.GetCellValue("Sheet1", "C3")
	assert.NoError(t, err)
	fmt.Printf("C3 = '%s' (expected: 'Extra')\n", c3)

	// Verify the values
	assert.Equal(t, "NewData", a3, "A3 should contain 'NewData'")
	assert.Equal(t, "999", b3, "B3 should contain '999'")
	assert.Equal(t, "Extra", c3, "C3 should contain 'Extra'")

	// Also check the entire row
	fmt.Println("\n=== 所有行的最终状态 ===")
	for i := 1; i <= 4; i++ {
		a, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		b, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i))
		c, _ := f.GetCellValue("Sheet1", fmt.Sprintf("C%d", i))
		fmt.Printf("Row %d: A=%s, B=%s, C=%s\n", i, a, b, c)
	}
}

// TestInsertRowBatchUpdateMultipleCells tests batch updating multiple cells in inserted row
func TestInsertRowBatchUpdateMultipleCells(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Setup with more columns
	for col := 1; col <= 5; col++ {
		colName, _ := ColumnNumberToName(col)
		f.SetCellValue("Sheet1", colName+"1", fmt.Sprintf("Col%d", col))
		f.SetCellValue("Sheet1", colName+"2", col*10)
	}

	fmt.Println("=== 初始数据 ===")
	row1, _ := f.GetRows("Sheet1")
	for i, row := range row1 {
		fmt.Printf("Row %d: %v\n", i+1, row)
	}

	// Insert row at position 2
	fmt.Println("\n=== 插入第 2 行 ===")
	err := f.InsertRows("Sheet1", 2, 1)
	assert.NoError(t, err)

	// Batch update all 5 columns in the new row
	fmt.Println("\n=== 批量更新第 2 行的 5 个单元格 ===")
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: "A"},
		{Sheet: "Sheet1", Cell: "B2", Value: "B"},
		{Sheet: "Sheet1", Cell: "C2", Value: "C"},
		{Sheet: "Sheet1", Cell: "D2", Value: "D"},
		{Sheet: "Sheet1", Cell: "E2", Value: "E"},
	}

	_, err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Get all values back
	fmt.Println("\n=== 获取第 2 行的值 ===")
	for _, update := range updates {
		val, err := f.GetCellValue(update.Sheet, update.Cell)
		assert.NoError(t, err)
		fmt.Printf("%s = '%s' (expected: '%v')\n", update.Cell, val, update.Value)
		assert.Equal(t, fmt.Sprint(update.Value), val, "Value should match")
	}

	// Check entire sheet
	fmt.Println("\n=== 所有行的最终状态 ===")
	rows, _ := f.GetRows("Sheet1")
	for i, row := range rows {
		fmt.Printf("Row %d: %v\n", i+1, row)
	}
}

// TestInsertRowBatchUpdateDifferentTypes tests different value types
func TestInsertRowBatchUpdateDifferentTypes(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Setup
	f.SetCellValue("Sheet1", "A1", "Header")

	// Insert row
	err := f.InsertRows("Sheet1", 2, 1)
	assert.NoError(t, err)

	// Batch update with different types
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: "String"},
		{Sheet: "Sheet1", Cell: "B2", Value: 123},
		{Sheet: "Sheet1", Cell: "C2", Value: 456.78},
		{Sheet: "Sheet1", Cell: "D2", Value: true},
		{Sheet: "Sheet1", Cell: "E2", Value: nil},
	}

	_, err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify each type
	tests := []struct {
		cell     string
		expected string
		desc     string
	}{
		{"A2", "String", "string value"},
		{"B2", "123", "integer value"},
		{"C2", "456.78", "float value"},
		{"D2", "TRUE", "boolean value"},
		{"E2", "", "nil value"},
	}

	fmt.Println("\n=== 验证不同类型的值 ===")
	for _, tt := range tests {
		val, err := f.GetCellValue("Sheet1", tt.cell)
		assert.NoError(t, err)
		fmt.Printf("%s = '%s' (expected: '%s') - %s\n", tt.cell, val, tt.expected, tt.desc)
		assert.Equal(t, tt.expected, val, "Value for %s should match", tt.cell)
	}
}

// TestInsertRowBatchUpdateWithSave tests if values persist after save
func TestInsertRowBatchUpdateWithSave(t *testing.T) {
	// Create temporary file
	tmpFile := "test_insert_batch.xlsx"
	defer func() {
		// Clean up
		f, _ := OpenFile(tmpFile)
		if f != nil {
			f.Close()
		}
	}()

	// Create and save
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "Header")

	// Insert and update
	f.InsertRows("Sheet1", 2, 1)
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: "TestData"},
		{Sheet: "Sheet1", Cell: "B2", Value: 999},
	}
	_, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Save
	err = f.SaveAs(tmpFile)
	assert.NoError(t, err)
	f.Close()

	// Reopen and verify
	fmt.Println("\n=== 重新打开文件验证 ===")
	f2, err := OpenFile(tmpFile)
	assert.NoError(t, err)
	defer f2.Close()

	a2, _ := f2.GetCellValue("Sheet1", "A2")
	b2, _ := f2.GetCellValue("Sheet1", "B2")

	fmt.Printf("重新打开后: A2='%s', B2='%s'\n", a2, b2)
	assert.Equal(t, "TestData", a2, "A2 should persist after save")
	assert.Equal(t, "999", b2, "B2 should persist after save")
}
