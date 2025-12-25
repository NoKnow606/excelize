package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestUpdateCellAndRecalculate(t *testing.T) {
	// Test 1: Basic recalculation
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a simple formula chain: A1 -> B1 -> C1
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=B1+5"))

	// Manually create a calcChain to simulate Excel's calculation order
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
			{R: "C1", I: 0}, // Same sheet as previous
		},
	}

	// Update A1 to 20 and recalculate dependent cells
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 20))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// Verify B1 was recalculated correctly (20*2 = 40)
	b1Value, err := f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "40", b1Value, "B1 should be recalculated to 40")

	// Verify C1 was recalculated correctly (40+5 = 45)
	c1Value, err := f.GetCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "45", c1Value, "C1 should be recalculated to 45")
}

func TestUpdateCellAndRecalculateWithoutCalcChain(t *testing.T) {
	// Test that it works even without calcChain
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 100))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))

	// Update A1
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 200))

	// Should work even without calcChain (but only recalculates if cell itself has formula)
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))
}

func TestUpdateCellAndRecalculateComplexChain(t *testing.T) {
	// Test with a more complex calculation chain
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up: A1=10, A2=20, B1=A1*2, B2=A2*2, C1=B1+B2
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "=A2*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=B1+B2"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
			{R: "B2", I: 0},
			{R: "C1", I: 0},
		},
	}

	// Update A1 to 30
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 30))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// B1 should be 30*2 = 60
	b1Value, err := f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "60", b1Value, "B1 should be recalculated to 60")

	// B2 should still be 40 (A2 didn't change)
	b2Value, err := f.GetCellValue("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "40", b2Value, "B2 should still be 40")

	// C1 should be 60+40 = 100
	c1Value, err := f.GetCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "100", c1Value, "C1 should be recalculated to 100")
}

func TestUpdateCellAndRecalculateNonExistentCell(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Should not error when updating non-existent cell
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "Z99"))
}

func TestUpdateCellAndRecalculateInvalidSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Should error for non-existent sheet
	err := f.UpdateCellAndRecalculate("NonExistentSheet", "A1")
	assert.Error(t, err)
}

func TestUpdateCellAndRecalculateInvalidCell(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Should error for invalid cell reference
	err := f.UpdateCellAndRecalculate("Sheet1", "INVALID")
	assert.Error(t, err)
}

func TestUpdateCellAndRecalculateMultiSheet(t *testing.T) {
	// Test that only cells in the same sheet are recalculated
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create second sheet
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set up formulas in both sheets
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	assert.NoError(t, f.SetCellValue("Sheet2", "A1", 100))
	assert.NoError(t, f.SetCellFormula("Sheet2", "B1", "=A1*2"))

	// Create calcChain with both sheets
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1}, // Sheet1
			{R: "B1", I: 2}, // Sheet2
		},
	}

	// Pre-calculate formulas using UpdateCellAndRecalculate for both sheets
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet2", "A1"))

	// Update A1 in Sheet1
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 20))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// Sheet1 B1 should be updated to 40
	b1Sheet1, err := f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "40", b1Sheet1, "Sheet1 B1 should be recalculated")

	// Sheet2 B1 should still be 200 (not affected)
	b1Sheet2, err := f.GetCellValue("Sheet2", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "200", b1Sheet2, "Sheet2 B1 should not change")
}

func TestUpdateCellAndRecalculateWithStrings(t *testing.T) {
	// Test that string formulas work correctly
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Hello"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1&\" World\""))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
		},
	}

	// Update A1
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Goodbye"))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// B1 should be "Goodbye World"
	b1Value, err := f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "Goodbye World", b1Value, "B1 should be recalculated with new string")
}

func TestUpdateCellAndRecalculateWithBooleans(t *testing.T) {
	// Test that boolean formulas work correctly
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1>5"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
		},
	}

	// Update A1 to 3
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 3))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// B1 should be FALSE
	b1Value, err := f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "FALSE", b1Value, "B1 should be FALSE")
}

func TestUpdateCellAndRecalculatePerformance(t *testing.T) {
	// Test with a chain of 100 formulas
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a chain: A1=1, A2=A1+1, A3=A2+1, ..., A100=A99+1
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 1))
	var calcChainCells []xlsxCalcChainC
	for i := 2; i <= 100; i++ {
		cell, _ := CoordinatesToCellName(1, i)
		prevCell, _ := CoordinatesToCellName(1, i-1)
		assert.NoError(t, f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1"))
		if i == 2 {
			calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 1})
		} else {
			calcChainCells = append(calcChainCells, xlsxCalcChainC{R: cell, I: 0})
		}
	}

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{C: calcChainCells}

	// Update A1 to 10
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// A100 should be 10+99 = 109
	a100Value, err := f.GetCellValue("Sheet1", "A100")
	assert.NoError(t, err)
	assert.Equal(t, "109", a100Value, "A100 should be recalculated to 109")
}
