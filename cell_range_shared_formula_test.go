package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestGetRangeDataConcurrentSharedFormula tests that GetRangeDataConcurrent
// correctly returns formulas for cells with shared formulas
func TestGetRangeDataConcurrentSharedFormula(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create data: B1=10, C1=20
	f.SetCellValue(sheetName, "B1", 10)
	f.SetCellValue(sheetName, "C1", 20)

	// Set formula in D1: =C1-1
	f.SetCellFormula(sheetName, "D1", "C1-1")

	// Copy D1 to D2 (this creates a shared formula)
	f.SetCellFormula(sheetName, "D2", "C2-1")

	fmt.Println("\n=== Test GetCellFormula vs GetRangeDataConcurrent ===")

	// Test 1: GetCellFormula should work
	d1Formula, err := f.GetCellFormula(sheetName, "D1")
	assert.NoError(t, err)
	fmt.Printf("GetCellFormula D1: '%s'\n", d1Formula)
	assert.Equal(t, "C1-1", d1Formula, "D1 should have formula C1-1")

	d2Formula, err := f.GetCellFormula(sheetName, "D2")
	assert.NoError(t, err)
	fmt.Printf("GetCellFormula D2: '%s'\n", d2Formula)
	assert.Equal(t, "C2-1", d2Formula, "D2 should have formula C2-1")

	// Test 2: GetRangeDataConcurrent should also return formulas
	data, err := f.GetRangeDataConcurrent(sheetName, "D1:D2")
	assert.NoError(t, err)

	fmt.Printf("\nGetRangeDataConcurrent results:\n")
	fmt.Printf("D1: Formula='%s', Value='%s'\n", data[0][0].Formula, data[0][0].Value)
	fmt.Printf("D2: Formula='%s', Value='%s'\n", data[1][0].Formula, data[1][0].Value)

	// Both should have formulas now
	assert.NotEmpty(t, data[0][0].Formula, "D1 should have formula in GetRangeDataConcurrent")
	assert.NotEmpty(t, data[1][0].Formula, "D2 should have formula in GetRangeDataConcurrent")

	// Verify formulas match GetCellFormula
	assert.Equal(t, "="+d1Formula, data[0][0].Formula, "D1 formula should match")
	assert.Equal(t, "="+d2Formula, data[1][0].Formula, "D2 formula should match")
}

// TestGetRangeDataConcurrentSharedFormulaDetailed tests shared formula handling
func TestGetRangeDataConcurrentSharedFormulaDetailed(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create a range with shared formulas
	f.SetCellValue(sheetName, "A1", 10)
	f.SetCellValue(sheetName, "A2", 20)
	f.SetCellValue(sheetName, "A3", 30)

	// Set formula in B1: =A1*2
	f.SetCellFormula(sheetName, "B1", "A1*2")
	// Copy to B2, B3 (creates shared formula)
	f.SetCellFormula(sheetName, "B2", "A2*2")
	f.SetCellFormula(sheetName, "B3", "A3*2")

	fmt.Println("\n=== Detailed Shared Formula Test ===")

	// Check each cell with GetCellFormula
	for row := 1; row <= 3; row++ {
		cell := fmt.Sprintf("B%d", row)
		formula, _ := f.GetCellFormula(sheetName, cell)
		fmt.Printf("GetCellFormula %s: '%s'\n", cell, formula)
	}

	// Check with GetRangeDataConcurrent
	data, err := f.GetRangeDataConcurrent(sheetName, "B1:B3")
	assert.NoError(t, err)

	fmt.Println("\nGetRangeDataConcurrent results:")
	for i, cellData := range data {
		cell := fmt.Sprintf("B%d", i+1)
		fmt.Printf("%s: Formula='%s', Value='%s'\n", cell, cellData[0].Formula, cellData[0].Value)

		// All cells should have formulas
		assert.NotEmpty(t, cellData[0].Formula, "%s should have formula", cell)
	}
}
