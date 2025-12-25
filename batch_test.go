package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestBatchSetCellValue(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Test 1: Basic batch update
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
		{Sheet: "Sheet1", Cell: "A2", Value: 200},
		{Sheet: "Sheet1", Cell: "A3", Value: 300},
		{Sheet: "Sheet1", Cell: "B1", Value: "text"},
		{Sheet: "Sheet1", Cell: "B2", Value: true},
	}

	err := f.BatchSetCellValue(updates)
	assert.NoError(t, err)

	// Verify values
	a1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "100", a1)

	a2, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "200", a2)

	a3, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "300", a3)

	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "text", b1)

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "TRUE", b2)
}

func TestBatchSetCellValueMultiSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create second sheet
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Update multiple sheets
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
		{Sheet: "Sheet1", Cell: "A2", Value: 200},
		{Sheet: "Sheet2", Cell: "A1", Value: 1000},
		{Sheet: "Sheet2", Cell: "A2", Value: 2000},
	}

	err = f.BatchSetCellValue(updates)
	assert.NoError(t, err)

	// Verify Sheet1
	a1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "100", a1)

	// Verify Sheet2
	a1Sheet2, _ := f.GetCellValue("Sheet2", "A1")
	assert.Equal(t, "1000", a1Sheet2)
}

func TestBatchSetCellValueInvalidSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	updates := []CellUpdate{
		{Sheet: "NonExistent", Cell: "A1", Value: 100},
	}

	err := f.BatchSetCellValue(updates)
	assert.Error(t, err)
}

func TestBatchSetCellValueInvalidCell(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "INVALID", Value: 100},
	}

	err := f.BatchSetCellValue(updates)
	assert.Error(t, err)
}

func TestRecalculateSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up formulas
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

	// Initial calculation
	assert.NoError(t, f.RecalculateSheet("Sheet1"))

	// Verify initial values
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1)

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "40", b2)

	c1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "60", c1)

	// Update values
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 30))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 40))

	// Recalculate
	assert.NoError(t, f.RecalculateSheet("Sheet1"))

	// Verify updated values
	b1, _ = f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "60", b1)

	b2, _ = f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "80", b2)

	c1, _ = f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "140", c1)
}

func TestRecalculateSheetInvalidSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	err := f.RecalculateSheet("NonExistent")
	assert.Error(t, err)
}

func TestRecalculateSheetNoCalcChain(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Should not error when there's no calcChain
	err := f.RecalculateSheet("Sheet1")
	assert.NoError(t, err)
}

func TestBatchUpdateAndRecalculate(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up formulas
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A10)"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "=AVERAGE(A1:A10)"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "=MAX(A1:A10)"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
			{R: "B2", I: 0},
			{R: "B3", I: 0},
		},
	}

	// Batch update 10 cells
	updates := make([]CellUpdate, 10)
	for i := 0; i < 10; i++ {
		cell, _ := CoordinatesToCellName(1, i+1)
		updates[i] = CellUpdate{
			Sheet: "Sheet1",
			Cell:  cell,
			Value: (i + 1) * 10,
		}
	}

	// Batch update and recalculate
	err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify formulas were recalculated
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "550", b1, "SUM(10,20,30,...,100) = 550")

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "55", b2, "AVERAGE(10,20,30,...,100) = 55")

	b3, _ := f.GetCellValue("Sheet1", "B3")
	assert.Equal(t, "100", b3, "MAX(10,20,30,...,100) = 100")
}

func TestBatchUpdateAndRecalculateMultiSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create second sheet
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set up formulas in both sheets
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A5)"))
	assert.NoError(t, f.SetCellFormula("Sheet2", "B1", "=SUM(A1:A5)"))

	// Create calcChain for both sheets
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1}, // Sheet1
			{R: "B1", I: 2}, // Sheet2
		},
	}

	// Batch update both sheets
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 10},
		{Sheet: "Sheet1", Cell: "A2", Value: 20},
		{Sheet: "Sheet1", Cell: "A3", Value: 30},
		{Sheet: "Sheet2", Cell: "A1", Value: 100},
		{Sheet: "Sheet2", Cell: "A2", Value: 200},
	}

	err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify Sheet1
	b1Sheet1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "60", b1Sheet1, "Sheet1: SUM(10,20,30,0,0) = 60")

	// Verify Sheet2
	b1Sheet2, _ := f.GetCellValue("Sheet2", "B1")
	assert.Equal(t, "300", b1Sheet2, "Sheet2: SUM(100,200,0,0,0) = 300")
}

func TestBatchUpdateAndRecalculateNoFormulas(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Batch update without any formulas (should not error)
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
		{Sheet: "Sheet1", Cell: "A2", Value: 200},
	}

	err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify values
	a1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "100", a1)
}

func TestBatchUpdateAndRecalculateComplexFormulas(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a complex formula dependency
	// A1, A2, A3 -> B1, B2 -> C1
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1+A2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "=A2+A3"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=B1+B2"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
			{R: "B2", I: 0},
			{R: "C1", I: 0},
		},
	}

	// Batch update
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 10},
		{Sheet: "Sheet1", Cell: "A2", Value: 20},
		{Sheet: "Sheet1", Cell: "A3", Value: 30},
	}

	err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify cascading calculation
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "30", b1, "B1 = A1+A2 = 10+20 = 30")

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "50", b2, "B2 = A2+A3 = 20+30 = 50")

	c1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "80", c1, "C1 = B1+B2 = 30+50 = 80")
}

func TestBatchUpdateAndRecalculateLargeDataset(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a formula that depends on many cells
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A100)"))

	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B1", I: 1}},
	}

	// Batch update 100 cells
	updates := make([]CellUpdate, 100)
	expectedSum := 0
	for i := 0; i < 100; i++ {
		cell, _ := CoordinatesToCellName(1, i+1)
		value := i + 1
		updates[i] = CellUpdate{
			Sheet: "Sheet1",
			Cell:  cell,
			Value: value,
		}
		expectedSum += value
	}

	err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify SUM
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, fmt.Sprintf("%d", expectedSum), b1, "SUM(1..100) = 5050")
}
