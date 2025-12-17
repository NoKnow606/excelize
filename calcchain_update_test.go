package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestUpdateCellCache(t *testing.T) {
	// Test 1: Basic single cell cache update
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a simple formula chain: A1 -> B1 -> C1
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=B1+5"))

	// Calculate to populate cache
	_, err := f.CalcCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	_, err = f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)

	// Update A1 and clear its cache
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 20))
	assert.NoError(t, f.UpdateCellCache("Sheet1", "A1"))

	// Test 2: Update cell cache without calcChain
	f2 := NewFile()
	defer func() {
		assert.NoError(t, f2.Close())
	}()

	assert.NoError(t, f2.SetCellValue("Sheet1", "A1", 100))
	assert.NoError(t, f2.SetCellFormula("Sheet1", "B1", "=A1*2"))

	// Should work even without calcChain
	assert.NoError(t, f2.UpdateCellCache("Sheet1", "A1"))

	// Test 3: Update non-existent cell
	assert.NoError(t, f2.UpdateCellCache("Sheet1", "Z99"))

	// Test 4: Update cell in non-existent sheet
	err = f2.UpdateCellCache("NonExistentSheet", "A1")
	assert.Error(t, err)

	// Test 5: Invalid cell reference
	err = f2.UpdateCellCache("Sheet1", "INVALID")
	assert.Error(t, err)
}

func TestUpdateCellCacheWithCalcChain(t *testing.T) {
	// Create a file with calcChain
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

	// Manually create a calcChain for testing
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
			{R: "B2", I: 0}, // I=0 means same sheet as previous
			{R: "C1", I: 0},
		},
	}

	// Calculate to populate cache
	_, err := f.CalcCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	_, err = f.CalcCellValue("Sheet1", "B2")
	assert.NoError(t, err)
	_, err = f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)

	// Update A1 - should clear B1 and C1 caches
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 30))
	assert.NoError(t, f.UpdateCellCache("Sheet1", "A1"))

	// Verify that B1's cache was cleared
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)

	// Find B1 and check if cache is cleared
	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			if cell.R == "B1" && cell.F != nil {
				// Cache should be cleared
				assert.Equal(t, "", cell.V, "B1 cache should be cleared")
			}
		}
	}
}

func TestFindCellInCalcChain(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create a test calcChain
	calcChain := &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "A1", I: 1},
			{R: "B1", I: 1},
			{R: "C1", I: 0}, // Same sheet as previous
			{R: "A1", I: 2}, // Different sheet
		},
	}

	// Test finding cells
	index := f.findCellInCalcChain(calcChain, 1, "A1")
	assert.Equal(t, 0, index, "Should find A1 in sheet 1 at index 0")

	index = f.findCellInCalcChain(calcChain, 1, "B1")
	assert.Equal(t, 1, index, "Should find B1 in sheet 1 at index 1")

	index = f.findCellInCalcChain(calcChain, 1, "C1")
	assert.Equal(t, 2, index, "Should find C1 in sheet 1 at index 2")

	index = f.findCellInCalcChain(calcChain, 2, "A1")
	assert.Equal(t, 3, index, "Should find A1 in sheet 2 at index 3")

	index = f.findCellInCalcChain(calcChain, 1, "Z99")
	assert.Equal(t, -1, index, "Should not find non-existent cell")

	index = f.findCellInCalcChain(calcChain, 3, "A1")
	assert.Equal(t, -1, index, "Should not find cell in non-existent sheet")
}

func TestClearCellCache(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a formula with cache (simulating an Excel file with cached values)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))

	// Manually set cache value to simulate Excel's cached result
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	for i := range ws.SheetData.Row {
		for j := range ws.SheetData.Row[i].C {
			if ws.SheetData.Row[i].C[j].R == "B1" {
				ws.SheetData.Row[i].C[j].V = "20"
				ws.SheetData.Row[i].C[j].T = "n"
			}
		}
	}

	// Verify cache exists
	ws, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)

	var cacheExists bool
	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			if cell.R == "B1" && cell.F != nil && cell.V != "" {
				cacheExists = true
				break
			}
		}
	}
	assert.True(t, cacheExists, "Cache should exist before clearing")

	// Clear the cache
	assert.NoError(t, f.clearCellFormulaCache("Sheet1", "B1"))

	// Verify cache is cleared
	ws, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)

	cacheExists = false
	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			if cell.R == "B1" && cell.F != nil && cell.V != "" {
				cacheExists = true
				break
			}
		}
	}
	assert.False(t, cacheExists, "Cache should be cleared")
}

func TestUpdateCellCachePerformance(t *testing.T) {
	// Create a file with many formulas
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up a chain of 100 formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 1))
	for i := 2; i <= 100; i++ {
		cell, _ := CoordinatesToCellName(1, i)
		prevCell, _ := CoordinatesToCellName(1, i-1)
		assert.NoError(t, f.SetCellFormula("Sheet1", cell, "="+prevCell+"+1"))
	}

	// Update the first cell
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))

	// This should be fast even without calcChain
	err := f.UpdateCellCache("Sheet1", "A1")
	assert.NoError(t, err)
}
