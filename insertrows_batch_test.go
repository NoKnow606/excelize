package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestInsertRowsThenBatchUpdate tests the scenario of inserting rows then batch updating
func TestInsertRowsThenBatchUpdate(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Setup: Create initial data with formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", 30))
	assert.NoError(t, f.SetCellValue("Sheet1", "A4", 40))

	// Setup formulas in column B
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
		{Sheet: "Sheet1", Cell: "B4", Formula: "=A4*2"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify initial state
	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "40", b2, "B2 should be 20*2=40 initially")

	fmt.Println("=== 初始状态 ===")
	for i := 1; i <= 4; i++ {
		a, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		b, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i))
		fmt.Printf("Row %d: A=%s, B=%s\n", i, a, b)
	}

	// Step 1: Insert a row at position 3 (between row 2 and 3)
	fmt.Println("\n=== 在第 3 行插入新行 ===")
	err = f.InsertRows("Sheet1", 3, 1)
	assert.NoError(t, err)

	// Check state after insert
	fmt.Println("\n=== 插入后的状态 ===")
	for i := 1; i <= 5; i++ {
		a, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		b, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i))
		bFormula, _ := f.GetCellFormula("Sheet1", fmt.Sprintf("B%d", i))
		fmt.Printf("Row %d: A=%s, B=%s, Formula=%s\n", i, a, b, bFormula)
	}

	// Step 2: Batch update the newly inserted row (row 3)
	fmt.Println("\n=== 批量更新第 3 行 ===")
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A3", Value: 100},
		{Sheet: "Sheet1", Cell: "B3", Value: 200},
		{Sheet: "Sheet1", Cell: "C3", Value: 300},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("受影响的单元格数: %d\n", len(affected))
	for _, cell := range affected {
		val, _ := f.GetCellValue(cell.Sheet, cell.Cell)
		fmt.Printf("  - %s!%s = %s\n", cell.Sheet, cell.Cell, val)
	}

	// Verify the updated values
	a3, _ := f.GetCellValue("Sheet1", "A3")
	b3, _ := f.GetCellValue("Sheet1", "B3")
	c3, _ := f.GetCellValue("Sheet1", "C3")

	assert.Equal(t, "100", a3, "A3 should be 100")
	assert.Equal(t, "200", b3, "B3 should be 200")
	assert.Equal(t, "300", c3, "C3 should be 300")

	// Verify that formulas were adjusted correctly
	fmt.Println("\n=== 最终状态 ===")
	for i := 1; i <= 5; i++ {
		a, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		b, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", i))
		bFormula, _ := f.GetCellFormula("Sheet1", fmt.Sprintf("B%d", i))
		fmt.Printf("Row %d: A=%s, B=%s, Formula=%s\n", i, a, b, bFormula)
	}

	// Original row 3 (now row 4) should still have correct formula
	b4Formula, _ := f.GetCellFormula("Sheet1", "B4")
	assert.Equal(t, "A4*2", b4Formula, "B4 formula should be adjusted to A4*2")

	b4Value, _ := f.GetCellValue("Sheet1", "B4")
	assert.Equal(t, "60", b4Value, "B4 should be 30*2=60")
}

// TestInsertRowsWithFormulaDependencies tests formula dependencies after row insertion
func TestInsertRowsWithFormulaDependencies(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create a sum formula that references a range
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", 30))

	_, err := f.BatchSetFormulasAndRecalculate([]FormulaUpdate{
		{Sheet: "Sheet1", Cell: "A4", Formula: "=SUM(A1:A3)"},
	})
	assert.NoError(t, err)

	// Verify initial sum
	a4, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "60", a4, "Initial sum should be 60")

	// Insert a row in the middle of the range
	err = f.InsertRows("Sheet1", 2, 1)
	assert.NoError(t, err)

	// Update the newly inserted row
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: 15},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\n受影响的单元格: %d\n", len(affected))
	for _, cell := range affected {
		val, _ := f.GetCellValue(cell.Sheet, cell.Cell)
		formula, _ := f.GetCellFormula(cell.Sheet, cell.Cell)
		fmt.Printf("  %s = %s (formula: %s)\n", cell.Cell, val, formula)
	}

	// Check if formula was adjusted
	a5Formula, _ := f.GetCellFormula("Sheet1", "A5")
	fmt.Printf("\nA5 formula after insert: %s\n", a5Formula)

	// The sum should now include the new row
	a5, _ := f.GetCellValue("Sheet1", "A5")
	fmt.Printf("A5 value: %s\n", a5)
}

// TestInsertRowsEmptyRowBatchUpdate tests inserting empty row then batch update
func TestInsertRowsEmptyRowBatchUpdate(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Setup simple data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Header1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", "Data1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", "Data2"))

	// Insert empty row at position 2
	err := f.InsertRows("Sheet1", 2, 1)
	assert.NoError(t, err)

	// Batch update the empty row
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: "NewData"},
		{Sheet: "Sheet1", Cell: "B2", Value: 123},
		{Sheet: "Sheet1", Cell: "C2", Value: 456.78},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Since there are no formulas, affected should be empty
	assert.Empty(t, affected, "No formulas, so no affected cells")

	// Verify the values
	a2, _ := f.GetCellValue("Sheet1", "A2")
	b2, _ := f.GetCellValue("Sheet1", "B2")
	c2, _ := f.GetCellValue("Sheet1", "C2")

	assert.Equal(t, "NewData", a2)
	assert.Equal(t, "123", b2)
	assert.Equal(t, "456.78", c2)

	// Verify original data shifted down
	a3, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Data1", a3, "Original row 2 should be at row 3 now")
}
