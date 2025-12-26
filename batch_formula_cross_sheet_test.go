package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchSetFormulasAndRecalculate_CrossSheetReference tests cross-sheet formula references
func TestBatchSetFormulasAndRecalculate_CrossSheetReference(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create Sheet2
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set up initial data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 100))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 200))

	// Set formulas with cross-sheet references
	formulas := []FormulaUpdate{
		// Sheet1 formulas
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},

		// Sheet2 formulas referencing Sheet1
		{Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+10"},
		{Sheet: "Sheet2", Cell: "C2", Formula: "=Sheet1!B2+10"},
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Verify initial calculations
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "200", b1, "Sheet1.B1 should be 100*2=200")

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "400", b2, "Sheet1.B2 should be 200*2=400")

	// ✅ Verify Sheet2 formulas are calculated
	c1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "210", c1, "Sheet2.C1 should be 200+10=210")

	c2, _ := f.GetCellValue("Sheet2", "C2")
	assert.Equal(t, "410", c2, "Sheet2.C2 should be 400+10=410")
}

// TestBatchSetFormulasAndRecalculate_CrossSheetChain tests cross-sheet dependency chain
func TestBatchSetFormulasAndRecalculate_CrossSheetChain(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create multiple sheets
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))

	// Create dependency chain: Sheet1 -> Sheet2 -> Sheet3
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},          // 20
		{Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+5"},   // 25
		{Sheet: "Sheet3", Cell: "D1", Formula: "=Sheet2!C1*3"},   // 75
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Verify chain: 10 -> 20 -> 25 -> 75
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1, "Sheet1.B1 should be 10*2=20")

	c1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "25", c1, "Sheet2.C1 should be 20+5=25")

	d1, _ := f.GetCellValue("Sheet3", "D1")
	assert.Equal(t, "75", d1, "Sheet3.D1 should be 25*3=75")
}

// TestBatchSetFormulasAndRecalculate_MixedSheetFormulas tests mixed sheet formulas
func TestBatchSetFormulasAndRecalculate_MixedSheetFormulas(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set initial data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 50))
	assert.NoError(t, f.SetCellValue("Sheet2", "A1", 100))

	// Set formulas that reference each other
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},          // 100
		{Sheet: "Sheet2", Cell: "B1", Formula: "=A1*3"},          // 300
		{Sheet: "Sheet1", Cell: "C1", Formula: "=Sheet2!B1+B1"},  // 400 (300+100)
		{Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+B1"},  // 400 (100+300)
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Verify all calculations
	sheet1B1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "100", sheet1B1)

	sheet2B1, _ := f.GetCellValue("Sheet2", "B1")
	assert.Equal(t, "300", sheet2B1)

	sheet1C1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "400", sheet1C1, "Sheet1.C1 should be 300+100=400")

	sheet2C1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "400", sheet2C1, "Sheet2.C1 should be 100+300=400")
}
