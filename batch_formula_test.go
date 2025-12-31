package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchSetFormulas tests BatchSetFormulas function
func TestBatchSetFormulas(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up data
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellValue("Sheet1", "A3", 30)

	// Batch set formulas
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
	}

	err := f.BatchSetFormulas(formulas)
	assert.NoError(t, err)

	// Verify formulas are set
	formula1, _ := f.GetCellFormula("Sheet1", "B1")
	assert.Equal(t, "=A1*2", formula1)

	formula2, _ := f.GetCellFormula("Sheet1", "B2")
	assert.Equal(t, "=A2*2", formula2)

	formula3, _ := f.GetCellFormula("Sheet1", "B3")
	assert.Equal(t, "=A3*2", formula3)
}

// TestBatchSetFormulasAndRecalculate tests BatchSetFormulasAndRecalculate function
func TestBatchSetFormulasAndRecalculate(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up data
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellValue("Sheet1", "A3", 30)

	// Batch set formulas and recalculate
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
		{Sheet: "Sheet1", Cell: "C1", Formula: "=SUM(B1:B3)"},
	}

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify formulas are set
	formula1, _ := f.GetCellFormula("Sheet1", "B1")
	assert.Equal(t, "=A1*2", formula1)

	formula4, _ := f.GetCellFormula("Sheet1", "C1")
	assert.Equal(t, "=SUM(B1:B3)", formula4)

	// Verify values are calculated
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1, "B1 should be 10*2=20")

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "40", b2, "B2 should be 20*2=40")

	b3, _ := f.GetCellValue("Sheet1", "B3")
	assert.Equal(t, "60", b3, "B3 should be 30*2=60")

	c1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "120", c1, "C1 should be SUM(20,40,60)=120")

	// Verify calcChain is updated
	assert.NotNil(t, f.CalcChain, "CalcChain should be created")
	assert.True(t, len(f.CalcChain.C) >= 4, "CalcChain should contain at least 4 entries")
}

// TestBatchSetFormulasAndRecalculate_WithFormulaPrefix tests formulas with '=' prefix
func TestBatchSetFormulasAndRecalculate_WithFormulaPrefix(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	f.SetCellValue("Sheet1", "A1", 100)

	// Test with '=' prefix
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"}, // With '='
		{Sheet: "Sheet1", Cell: "B2", Formula: "A1*3"},  // Without '='
	}

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Both should work
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "200", b1)

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "300", b2)
}

// TestBatchSetFormulasAndRecalculate_MultiSheet tests with multiple sheets
func TestBatchSetFormulasAndRecalculate_MultiSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create Sheet2
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set up data in both sheets
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet2", "A1", 100)

	// Batch set formulas in both sheets
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=B1+10"},
		{Sheet: "Sheet2", Cell: "B1", Formula: "=A1*3"},
		{Sheet: "Sheet2", Cell: "B2", Formula: "=B1+20"},
	}

	err = f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify Sheet1
	b1Sheet1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1Sheet1, "Sheet1 B1 should be 10*2=20")

	b2Sheet1, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "30", b2Sheet1, "Sheet1 B2 should be 20+10=30")

	// Verify Sheet2
	b1Sheet2, _ := f.GetCellValue("Sheet2", "B1")
	assert.Equal(t, "300", b1Sheet2, "Sheet2 B1 should be 100*3=300")

	b2Sheet2, _ := f.GetCellValue("Sheet2", "B2")
	assert.Equal(t, "320", b2Sheet2, "Sheet2 B2 should be 300+20=320")
}

// TestBatchSetFormulasAndRecalculate_ComplexDependencies tests complex formula dependencies
func TestBatchSetFormulasAndRecalculate_ComplexDependencies(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up base data
	for i := 1; i <= 10; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i*10)
	}

	// Create complex dependencies: A->B->C->D
	formulas := []FormulaUpdate{
		// B column: double A values
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},

		// C column: sum of B values
		{Sheet: "Sheet1", Cell: "C1", Formula: "=SUM(B1:B3)"},

		// D column: average of A values
		{Sheet: "Sheet1", Cell: "D1", Formula: "=AVERAGE(A1:A10)"},

		// E column: depends on both C and D
		{Sheet: "Sheet1", Cell: "E1", Formula: "=C1+D1"},
	}

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify calculated values
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1, "B1 should be 10*2=20")

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "40", b2, "B2 should be 20*2=40")

	b3, _ := f.GetCellValue("Sheet1", "B3")
	assert.Equal(t, "60", b3, "B3 should be 30*2=60")

	c1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "120", c1, "C1 should be SUM(20,40,60)=120")

	d1, _ := f.GetCellValue("Sheet1", "D1")
	assert.Equal(t, "55", d1, "D1 should be AVERAGE(10,20,...,100)=55")

	e1, _ := f.GetCellValue("Sheet1", "E1")
	assert.Equal(t, "175", e1, "E1 should be 120+55=175")
}

// TestBatchSetFormulasAndRecalculate_EmptyList tests with empty formula list
func TestBatchSetFormulasAndRecalculate_EmptyList(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Empty list should not error
	err := f.BatchSetFormulasAndRecalculate([]FormulaUpdate{})
	assert.NoError(t, err)
}

// TestBatchSetFormulasAndRecalculate_InvalidSheet tests error handling
func TestBatchSetFormulasAndRecalculate_InvalidSheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	formulas := []FormulaUpdate{
		{Sheet: "NonExistent", Cell: "B1", Formula: "=A1*2"},
	}

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.Error(t, err, "Should error for non-existent sheet")
}

// TestBatchSetFormulasAndRecalculate_LargeDataset tests with large dataset
func TestBatchSetFormulasAndRecalculate_LargeDataset(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up 100 data cells
	for i := 1; i <= 100; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
	}

	// Create 100 formula cells
	formulas := make([]FormulaUpdate, 100)
	for i := 1; i <= 100; i++ {
		formulas[i-1] = FormulaUpdate{
			Sheet:   "Sheet1",
			Cell:    fmt.Sprintf("B%d", i),
			Formula: fmt.Sprintf("=A%d*2", i),
		}
	}

	// Add a SUM formula
	formulas = append(formulas, FormulaUpdate{
		Sheet:   "Sheet1",
		Cell:    "C1",
		Formula: "=SUM(B1:B100)",
	})

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify some values
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "2", b1)

	b50, _ := f.GetCellValue("Sheet1", "B50")
	assert.Equal(t, "100", b50, "B50 should be 50*2=100")

	b100, _ := f.GetCellValue("Sheet1", "B100")
	assert.Equal(t, "200", b100, "B100 should be 100*2=200")

	// Verify SUM
	c1, _ := f.GetCellValue("Sheet1", "C1")
	expectedSum := (1 + 100) * 100 // Sum of 2,4,6,...,200 = 2*(1+2+...+100) = 2*5050 = 10100
	assert.Equal(t, fmt.Sprintf("%d", expectedSum), c1, "C1 should be SUM of all B values")
}

// TestBatchSetFormulasAndRecalculate_CalcChainUpdate tests calcChain is properly updated
func TestBatchSetFormulasAndRecalculate_CalcChainUpdate(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	f.SetCellValue("Sheet1", "A1", 10)

	// Set formulas
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=B1+10"},
	}

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// CalcChain should now exist and not be empty
	assert.NotNil(t, f.CalcChain, "CalcChain should be created")
	assert.True(t, len(f.CalcChain.C) >= 2, "CalcChain should contain at least 2 entries")

	// Verify calcChain contains our cells
	hasB1 := false
	hasB2 := false
	for _, entry := range f.CalcChain.C {
		if entry.R == "B1" {
			hasB1 = true
		}
		if entry.R == "B2" {
			hasB2 = true
		}
	}
	assert.True(t, hasB1, "CalcChain should contain B1")
	assert.True(t, hasB2, "CalcChain should contain B2")
}

// TestBatchSetFormulasAndRecalculate_UpdateExistingFormulas tests updating existing formulas
func TestBatchSetFormulasAndRecalculate_UpdateExistingFormulas(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	f.SetCellValue("Sheet1", "A1", 10)

	// Set initial formula
	f.SetCellFormula("Sheet1", "B1", "=A1*2")
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}
	f.UpdateCellAndRecalculate("Sheet1", "A1")

	b1Before, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1Before)

	// Update formula using batch API
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*3"}, // Changed from *2 to *3
	}

	err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify formula is updated
	formulaAfter, _ := f.GetCellFormula("Sheet1", "B1")
	assert.Equal(t, "=A1*3", formulaAfter)

	// Verify value is recalculated
	b1After, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "30", b1After, "B1 should be updated to 10*3=30")
}
