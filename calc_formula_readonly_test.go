package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestCalcFormulaValueReadOnly tests that CalcFormulaValue doesn't create rows
func TestCalcFormulaValueReadOnly(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set some base data
	f.SetCellValue(sheet, "B1", 10)
	f.SetCellValue(sheet, "B2", 20)
	f.SetCellValue(sheet, "B3", 30)

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	t.Logf("Initial rows: %d", initialRows)

	// Test 1: Calculate formula on non-existent far cell
	result, err := f.CalcFormulaValue(sheet, "Z9999", "SUM(B1:B3)")
	if err != nil {
		t.Errorf("CalcFormulaValue failed: %v", err)
	}
	if result != "60" {
		t.Errorf("Expected 60, got %s", result)
	}

	// Check row count - should have created row 9999 temporarily, then cleaned up
	ws, _ = f.workSheetReader(sheet)
	afterRows := len(ws.SheetData.Row)

	t.Logf("After CalcFormulaValue on Z9999:")
	t.Logf("  Rows: %d (initial: %d)", afterRows, initialRows)
	t.Logf("  Result: %s", result)

	// The row might have been created temporarily but should be cleaned up
	if afterRows >= 9999 {
		t.Errorf("CalcFormulaValue created too many rows: %d (should not create up to 9999)", afterRows)
	}

	// Test 2: Verify cell Z9999 doesn't have a persisted formula
	formula, _ := f.GetCellFormula(sheet, "Z9999")
	if formula != "" {
		t.Errorf("Formula should not be persisted, but got: %s", formula)
	}

	t.Logf("✓ CalcFormulaValue is read-only (no persistent changes)")
}

// TestCalcFormulaValueMinimalRowCreation tests minimal row creation
func TestCalcFormulaValueMinimalRowCreation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Only set A1
	f.SetCellValue(sheet, "A1", 100)

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	// Calculate on row 100 (far away)
	result, err := f.CalcFormulaValue(sheet, "A100", "A1*2")
	if err != nil {
		t.Errorf("CalcFormulaValue failed: %v", err)
	}
	if result != "200" {
		t.Errorf("Expected 200, got %s", result)
	}

	ws, _ = f.workSheetReader(sheet)
	afterRows := len(ws.SheetData.Row)

	t.Logf("Minimal row creation test:")
	t.Logf("  Initial rows: %d", initialRows)
	t.Logf("  After calculating A100: %d", afterRows)
	t.Logf("  Rows created: %d", afterRows-initialRows)

	// Should create row 100, but after cleanup should reset or minimal
	if afterRows > 100 {
		t.Logf("  Note: %d rows exist (temporary row 100 was created)", afterRows)
	}
}

// TestCalcFormulaValueVsGetCellStyle compares memory footprint
func TestCalcFormulaValueVsGetCellStyle(t *testing.T) {
	const farRow = 5000

	// Scenario 1: Using old approach with prepareCell
	f1 := NewFile()
	defer f1.Close()

	f1.SetCellValue("Sheet1", "B1", 10)
	f1.SetCellValue("Sheet1", "B2", 20)

	// Old way: SetCellFormula would call prepareCell
	f1.SetCellFormula("Sheet1", fmt.Sprintf("A%d", farRow), "SUM(B1:B2)")
	f1.CalcCellValue("Sheet1", fmt.Sprintf("A%d", farRow))

	ws1, _ := f1.workSheetReader("Sheet1")
	rows1 := len(ws1.SheetData.Row)

	// Scenario 2: Using new CalcFormulaValue
	f2 := NewFile()
	defer f2.Close()

	f2.SetCellValue("Sheet1", "B1", 10)
	f2.SetCellValue("Sheet1", "B2", 20)

	// New way: CalcFormulaValue (optimized)
	f2.CalcFormulaValue("Sheet1", fmt.Sprintf("A%d", farRow), "SUM(B1:B2)")

	ws2, _ := f2.workSheetReader("Sheet1")
	rows2 := len(ws2.SheetData.Row)

	t.Logf("\n=== Memory Footprint Comparison ===")
	t.Logf("Old approach (SetCellFormula + CalcCellValue):")
	t.Logf("  Rows created: %d", rows1)
	t.Logf("\nNew approach (CalcFormulaValue):")
	t.Logf("  Rows created: %d", rows2)
	t.Logf("\nMemory saved: %d rows", rows1-rows2)

	if rows2 < rows1 {
		t.Logf("✓ CalcFormulaValue uses less memory")
	}
}

// TestCalcFormulaValueBatchReadOnly tests batch calculations don't bloat memory
func TestCalcFormulaValueBatchReadOnly(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set base data
	for i := 1; i <= 10; i++ {
		f.SetCellValue(sheet, fmt.Sprintf("B%d", i), i*10)
	}

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	// Calculate formulas on far-away cells (row 1000-1100)
	formulas := make(map[string]string)
	for i := 1000; i < 1100; i++ {
		formulas[fmt.Sprintf("A%d", i)] = "SUM(B1:B10)"
	}

	start := time.Now()
	results, err := f.CalcFormulasValues(sheet, formulas)
	duration := time.Since(start)

	if err != nil {
		t.Errorf("CalcFormulasValues failed: %v", err)
	}

	if len(results) != 100 {
		t.Errorf("Expected 100 results, got %d", len(results))
	}

	ws, _ = f.workSheetReader(sheet)
	afterRows := len(ws.SheetData.Row)

	t.Logf("\n=== Batch Calculation Test ===")
	t.Logf("Calculated 100 formulas in %v", duration)
	t.Logf("Initial rows: %d", initialRows)
	t.Logf("After batch: %d", afterRows)
	t.Logf("Rows created: %d", afterRows-initialRows)

	// Check one result
	if results["A1000"] != "550" {
		t.Errorf("Expected 550, got %s", results["A1000"])
	}

	t.Logf("✓ Batch calculation completed")
}

// TestCalcFormulaValueReadOnlyPerformance benchmarks the improvement
func TestCalcFormulaValueReadOnlyPerformance(t *testing.T) {
	const iterations = 1000
	sheet := "Sheet1"

	// Setup data
	f := NewFile()
	defer f.Close()

	for i := 1; i <= 100; i++ {
		f.SetCellValue(sheet, fmt.Sprintf("B%d", i), i)
	}

	// Warm up
	for i := 0; i < 10; i++ {
		f.CalcFormulaValue(sheet, "A1", "SUM(B1:B100)")
	}

	// Test: Calculate formulas on different cells
	start := time.Now()
	for i := 0; i < iterations; i++ {
		cell := fmt.Sprintf("Z%d", (i%1000)+1)
		f.CalcFormulaValue(sheet, cell, "SUM(B1:B100)")
	}
	duration := time.Since(start)

	ws, _ := f.workSheetReader(sheet)
	finalRows := len(ws.SheetData.Row)

	t.Logf("\n=== Performance Test ===")
	t.Logf("Iterations: %d", iterations)
	t.Logf("Duration: %v", duration)
	t.Logf("Avg per call: %v", duration/iterations)
	t.Logf("Final row count: %d", finalRows)
	t.Logf("Throughput: %.0f calculations/sec", float64(iterations)/duration.Seconds())

	if finalRows > 1000 {
		t.Logf("Warning: Created %d rows (expected ≤1000)", finalRows)
	}
}
