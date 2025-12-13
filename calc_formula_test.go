package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestCalcFormulaValue tests the basic functionality of CalcFormulaValue
func TestCalcFormulaValue(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set some base values
	_ = f.SetCellValue(sheet, "B1", 10)
	_ = f.SetCellValue(sheet, "B2", 20)
	_ = f.SetCellValue(sheet, "B3", 30)

	// Test 1: Calculate formula without modifying cell
	t.Run("Calculate without modifying cell", func(t *testing.T) {
		result, err := f.CalcFormulaValue(sheet, "A1", "SUM(B1:B3)")
		if err != nil {
			t.Errorf("CalcFormulaValue failed: %v", err)
		}

		if result != "60" {
			t.Errorf("Expected 60, got %s", result)
		}

		// Verify A1 still has no formula
		formula, err := f.GetCellFormula(sheet, "A1")
		if err != nil {
			t.Errorf("GetCellFormula failed: %v", err)
		}
		if formula != "" {
			t.Errorf("Expected no formula in A1, got %s", formula)
		}
	})

	// Test 2: Calculate formula on cell with existing formula
	t.Run("Calculate on cell with existing formula", func(t *testing.T) {
		// Set an existing formula
		_ = f.SetCellFormula(sheet, "A2", "B1+B2")

		// Calculate a different formula temporarily
		result, err := f.CalcFormulaValue(sheet, "A2", "B1*B2")
		if err != nil {
			t.Errorf("CalcFormulaValue failed: %v", err)
		}

		if result != "200" {
			t.Errorf("Expected 200, got %s", result)
		}

		// Verify original formula is preserved
		formula, err := f.GetCellFormula(sheet, "A2")
		if err != nil {
			t.Errorf("GetCellFormula failed: %v", err)
		}
		if formula != "B1+B2" {
			t.Errorf("Expected B1+B2, got %s", formula)
		}
	})

	// Test 3: Multiple calculations on same cell
	t.Run("Multiple calculations on same cell", func(t *testing.T) {
		formulas := []string{"SUM(B1:B3)", "AVERAGE(B1:B3)", "MAX(B1:B3)", "MIN(B1:B3)"}
		expected := []string{"60", "20", "30", "10"}

		for i, formula := range formulas {
			result, err := f.CalcFormulaValue(sheet, "A3", formula)
			if err != nil {
				t.Errorf("CalcFormulaValue(%s) failed: %v", formula, err)
			}
			if result != expected[i] {
				t.Errorf("Formula %s: expected %s, got %s", formula, expected[i], result)
			}
		}

		// Verify A3 still has no formula
		formula, err := f.GetCellFormula(sheet, "A3")
		if err != nil {
			t.Errorf("GetCellFormula failed: %v", err)
		}
		if formula != "" {
			t.Errorf("Expected no formula in A3, got %s", formula)
		}
	})
}

// TestCalcFormulasValues tests batch formula calculation
func TestCalcFormulasValues(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set base values
	for i := 1; i <= 10; i++ {
		cell := fmt.Sprintf("B%d", i)
		_ = f.SetCellValue(sheet, cell, i*10)
	}

	// Test batch calculation
	formulas := map[string]string{
		"A1": "SUM(B1:B10)",
		"A2": "AVERAGE(B1:B10)",
		"A3": "MAX(B1:B10)",
		"A4": "MIN(B1:B10)",
		"A5": "COUNT(B1:B10)",
	}

	results, err := f.CalcFormulasValues(sheet, formulas)
	if err != nil {
		t.Errorf("CalcFormulasValues failed: %v", err)
	}

	expected := map[string]string{
		"A1": "550",
		"A2": "55",
		"A3": "100",
		"A4": "10",
		"A5": "10",
	}

	for cell, expectedValue := range expected {
		if results[cell] != expectedValue {
			t.Errorf("Cell %s: expected %s, got %s", cell, expectedValue, results[cell])
		}
	}

	// Verify no formulas were persisted
	for cell := range formulas {
		formula, err := f.GetCellFormula(sheet, cell)
		if err != nil {
			t.Errorf("GetCellFormula(%s) failed: %v", cell, err)
		}
		if formula != "" {
			t.Errorf("Cell %s should have no formula, got %s", cell, formula)
		}
	}
}

// TestCalcFormulaValuePerformance compares performance of different approaches
func TestCalcFormulaValuePerformance(t *testing.T) {
	const iterations = 1000

	// Setup
	f1 := NewFile()
	defer f1.Close()
	f2 := NewFile()
	defer f2.Close()

	sheet := "Sheet1"

	// Set base values
	for i := 1; i <= 100; i++ {
		cell := fmt.Sprintf("B%d", i)
		_ = f1.SetCellValue(sheet, cell, i)
		_ = f2.SetCellValue(sheet, cell, i)
	}

	// Test 1: SetCellFormula + CalcCellValue (traditional approach)
	t.Logf("\n=== Test 1: SetCellFormula + CalcCellValue ===")
	start := time.Now()
	for i := 0; i < iterations; i++ {
		_ = f1.SetCellFormula(sheet, "A1", "SUM(B1:B100)")
		_, _ = f1.CalcCellValue(sheet, "A1")
	}
	duration1 := time.Since(start)

	t.Logf("Iterations: %d", iterations)
	t.Logf("Duration: %v", duration1)
	t.Logf("Avg per iteration: %v", duration1/time.Duration(iterations))

	// Test 2: CalcFormulaValue (new approach)
	t.Logf("\n=== Test 2: CalcFormulaValue ===")
	start = time.Now()
	for i := 0; i < iterations; i++ {
		_, _ = f2.CalcFormulaValue(sheet, "A1", "SUM(B1:B100)")
	}
	duration2 := time.Since(start)

	t.Logf("Iterations: %d", iterations)
	t.Logf("Duration: %v", duration2)
	t.Logf("Avg per iteration: %v", duration2/time.Duration(iterations))

	// Performance comparison
	t.Logf("\n=== Performance Summary ===")
	t.Logf("Traditional approach: %v", duration1)
	t.Logf("CalcFormulaValue:     %v", duration2)
	speedup := float64(duration1) / float64(duration2)
	t.Logf("Speedup: %.2fx faster", speedup)

	if speedup < 1.5 {
		t.Logf("WARNING: Expected at least 1.5x speedup, got %.2fx", speedup)
	}
}

// TestCalcFormulaValueCacheNotCleared verifies cache is not cleared
func TestCalcFormulaValueCacheNotCleared(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set base values
	_ = f.SetCellValue(sheet, "B1", 10)
	_ = f.SetCellValue(sheet, "B2", 20)

	// Set a formula in A1
	_ = f.SetCellFormula(sheet, "A1", "B1+B2")

	// First calculation (cold cache)
	start := time.Now()
	result1, _ := f.CalcCellValue(sheet, "A1")
	duration1 := time.Since(start)

	// Second calculation (hot cache)
	start = time.Now()
	result2, _ := f.CalcCellValue(sheet, "A1")
	duration2 := time.Since(start)

	t.Logf("First calc (cold): %v, result: %s", duration1, result1)
	t.Logf("Second calc (hot): %v, result: %s", duration2, result2)

	// Use CalcFormulaValue on different cell
	_, _ = f.CalcFormulaValue(sheet, "A2", "B1*B2")

	// Calculate A1 again - cache should still be hot
	start = time.Now()
	result3, _ := f.CalcCellValue(sheet, "A1")
	duration3 := time.Since(start)

	t.Logf("Third calc (after CalcFormulaValue): %v, result: %s", duration3, result3)

	// Verify cache was not cleared (duration3 should be similar to duration2)
	if duration3 > duration2*10 {
		t.Errorf("Cache appears to have been cleared: duration2=%v, duration3=%v", duration2, duration3)
	} else {
		t.Logf("✓ Cache preserved after CalcFormulaValue")
	}
}

// TestCalcFormulaValueErrorHandling tests error scenarios
func TestCalcFormulaValueErrorHandling(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Test 1: Invalid formula
	t.Run("Invalid formula", func(t *testing.T) {
		result, err := f.CalcFormulaValue(sheet, "A1", "INVALID()")
		if err == nil {
			t.Error("Expected error for invalid formula")
		}
		t.Logf("Result: %s, Error: %v", result, err)
	})

	// Test 2: Invalid sheet
	t.Run("Invalid sheet", func(t *testing.T) {
		_, err := f.CalcFormulaValue("NonExistentSheet", "A1", "1+1")
		if err == nil {
			t.Error("Expected error for invalid sheet")
		}
	})

	// Test 3: Invalid cell reference
	t.Run("Invalid cell reference", func(t *testing.T) {
		_, err := f.CalcFormulaValue(sheet, "INVALID", "1+1")
		if err == nil {
			t.Error("Expected error for invalid cell")
		}
	})
}

// TestCalcFormulasValuesWithErrors tests partial success scenarios
func TestCalcFormulasValuesWithErrors(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	_ = f.SetCellValue(sheet, "B1", 10)

	formulas := map[string]string{
		"A1": "B1*2",           // Valid
		"A2": "INVALID()",      // Invalid formula
		"A3": "SUM(B1:B10)",    // Valid
		"A4": "NONEXIST()",     // Invalid function
	}

	results, err := f.CalcFormulasValues(sheet, formulas)

	// Should return partial results
	if len(results) == 0 {
		t.Error("Expected some successful results")
	}

	// Should return error
	if err == nil {
		t.Error("Expected error due to invalid formulas")
	}

	t.Logf("Successful calculations: %d", len(results))
	t.Logf("Error: %v", err)

	// Verify successful calculations
	if results["A1"] != "20" {
		t.Errorf("A1: expected 20, got %s", results["A1"])
	}
}
