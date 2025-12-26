package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateAndRecalculate_CrossSheet tests cross-sheet formula recalculation
func TestBatchUpdateAndRecalculate_CrossSheet(t *testing.T) {
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

	// Create cross-sheet formulas
	formulas := []FormulaUpdate{
		// Sheet1 formulas
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},

		// Sheet2 formulas referencing Sheet1
		{Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+10"},
		{Sheet: "Sheet2", Cell: "C2", Formula: "=Sheet1!B2+10"},
		{Sheet: "Sheet2", Cell: "C3", Formula: "=Sheet1!A1+Sheet1!A2"},
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Verify initial calculations
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "200", b1, "Sheet1.B1 should be 100*2=200")

	c1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "210", c1, "Sheet2.C1 should be 200+10=210")

	c3, _ := f.GetCellValue("Sheet2", "C3")
	assert.Equal(t, "300", c3, "Sheet2.C3 should be 100+200=300")

	// ✅ Now update Sheet1 values and verify cross-sheet recalculation
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 500},
		{Sheet: "Sheet1", Cell: "A2", Value: 600},
	}

	assert.NoError(t, f.BatchUpdateAndRecalculate(updates))

	// Verify Sheet1 formulas are recalculated
	b1After, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "1000", b1After, "Sheet1.B1 should be 500*2=1000")

	b2After, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "1200", b2After, "Sheet1.B2 should be 600*2=1200")

	// ✅ Verify Sheet2 formulas are also recalculated (cross-sheet dependency)
	c1After, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "1010", c1After, "Sheet2.C1 should be 1000+10=1010")

	c2After, _ := f.GetCellValue("Sheet2", "C2")
	assert.Equal(t, "1210", c2After, "Sheet2.C2 should be 1200+10=1210")

	c3After, _ := f.GetCellValue("Sheet2", "C3")
	assert.Equal(t, "1100", c3After, "Sheet2.C3 should be 500+600=1100")
}

// TestBatchUpdateAndRecalculate_CrossSheetComplex tests complex cross-sheet dependencies
func TestBatchUpdateAndRecalculate_CrossSheetComplex(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create multiple sheets
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)

	// Set up data: Sheet1 -> Sheet2 -> Sheet3 (chain)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))

	formulas := []FormulaUpdate{
		// Sheet1: B1 = A1 * 2
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},

		// Sheet2: C1 = Sheet1!B1 + 5
		{Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+5"},

		// Sheet3: D1 = Sheet2!C1 * 3
		{Sheet: "Sheet3", Cell: "D1", Formula: "=Sheet2!C1*3"},
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Verify initial chain: 10 -> 20 -> 25 -> 75
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1)

	c1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "25", c1)

	d1, _ := f.GetCellValue("Sheet3", "D1")
	assert.Equal(t, "75", d1)

	// Update Sheet1 and verify entire chain recalculates
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 50},
	}

	assert.NoError(t, f.BatchUpdateAndRecalculate(updates))

	// Verify chain: 50 -> 100 -> 105 -> 315
	b1After, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "100", b1After, "Sheet1.B1 should be 50*2=100")

	c1After, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "105", c1After, "Sheet2.C1 should be 100+5=105")

	d1After, _ := f.GetCellValue("Sheet3", "D1")
	assert.Equal(t, "315", d1After, "Sheet3.D1 should be 105*3=315")
}

// TestBatchUpdateAndRecalculate_CrossSheetMultipleUpdates tests multiple updates affecting cross-sheet formulas
func TestBatchUpdateAndRecalculate_CrossSheetMultipleUpdates(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set initial data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", 30))

	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
		{Sheet: "Sheet2", Cell: "C1", Formula: "=SUM(Sheet1!B1:B3)"},
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Initial: SUM(20,40,60) = 120
	c1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "120", c1)

	// Update multiple cells in Sheet1
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
		{Sheet: "Sheet1", Cell: "A2", Value: 200},
		{Sheet: "Sheet1", Cell: "A3", Value: 300},
	}

	assert.NoError(t, f.BatchUpdateAndRecalculate(updates))

	// Verify Sheet2 formula recalculates: SUM(200,400,600) = 1200
	c1After, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "1200", c1After, "Sheet2.C1 should be SUM(200,400,600)=1200")
}

// TestBatchUpdateAndRecalculate_SingleSheetStillWorks tests that single-sheet updates still work
func TestBatchUpdateAndRecalculate_SingleSheetStillWorks(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Single sheet scenario
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))

	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "C1", Formula: "=B1+5"},
	}

	assert.NoError(t, f.BatchSetFormulasAndRecalculate(formulas))

	// Verify initial
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1)

	c1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "25", c1)

	// Update and recalculate
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
	}

	assert.NoError(t, f.BatchUpdateAndRecalculate(updates))

	// Verify recalculation
	b1After, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "200", b1After)

	c1After, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "205", c1After)
}
