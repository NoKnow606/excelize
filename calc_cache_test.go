package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestCalcCellValuesCachePerformance tests the performance impact of calculation cache
func TestCalcCellValuesCachePerformance(t *testing.T) {
	const rows = 10000
	const cols = 100

	f := NewFile()
	sheet := "Sheet1"

	t.Logf("Preparing test data with complex formulas: %d rows x %d cols...", rows, cols)

	// Create data with interdependent formulas to benefit from caching
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f.SetCellValue(sheet, cell, c*100)
	}

	// Create formulas that reference multiple cells
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			if r%5 == 0 {
				// Complex formula referencing multiple cells
				prevCell1, _ := CoordinatesToCellName(c, r-1)
				prevCell2, _ := CoordinatesToCellName(c, r-2)
				formula := fmt.Sprintf("=%s+%s+%d", prevCell1, prevCell2, c)
				_ = f.SetCellFormula(sheet, cell, formula)
			} else if r%3 == 0 {
				// Reference formula
				prevCell, _ := CoordinatesToCellName(c, r-1)
				formula := fmt.Sprintf("=%s*2", prevCell)
				_ = f.SetCellFormula(sheet, cell, formula)
			} else {
				_ = f.SetCellValue(sheet, cell, r*c)
			}
		}
	}

	t.Logf("Data preparation completed")

	// Build list of cells to calculate
	cells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			cells = append(cells, cell)
		}
	}

	totalCells := len(cells)

	// Test 1: First calculation (cold cache)
	t.Logf("\n=== Test 1: First Calculation (Cold Cache) ===")
	start := time.Now()
	results1, err1 := f.CalcCellValues(sheet, cells)
	duration1 := time.Since(start)

	if err1 != nil {
		t.Logf("Warning: %v", err1)
	}

	t.Logf("Duration: %v", duration1)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration1.Seconds())
	t.Logf("Avg per cell: %v", duration1/time.Duration(totalCells))
	t.Logf("Successful: %d/%d", len(results1), totalCells)

	// Test 2: Second calculation (hot cache)
	t.Logf("\n=== Test 2: Second Calculation (Hot Cache) ===")
	start = time.Now()
	results2, err2 := f.CalcCellValues(sheet, cells)
	duration2 := time.Since(start)

	if err2 != nil {
		t.Logf("Warning: %v", err2)
	}

	t.Logf("Duration: %v", duration2)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration2.Seconds())
	t.Logf("Avg per cell: %v", duration2/time.Duration(totalCells))
	t.Logf("Successful: %d/%d", len(results2), totalCells)

	// Test 3: Third calculation after cache clear
	t.Logf("\n=== Test 3: After Cache Clear (Cold Cache Again) ===")
	f.calcCache.Clear()
	start = time.Now()
	results3, err3 := f.CalcCellValues(sheet, cells)
	duration3 := time.Since(start)

	if err3 != nil {
		t.Logf("Warning: %v", err3)
	}

	t.Logf("Duration: %v", duration3)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration3.Seconds())
	t.Logf("Avg per cell: %v", duration3/time.Duration(totalCells))
	t.Logf("Successful: %d/%d", len(results3), totalCells)

	// Performance comparison
	t.Logf("\n=== Performance Summary ===")
	t.Logf("Total cells: %d", totalCells)
	t.Logf("")
	t.Logf("Cold Cache (1st): %v", duration1)
	t.Logf("Hot Cache (2nd):  %v (%.2fx faster)", duration2, float64(duration1)/float64(duration2))
	t.Logf("Cold Cache (3rd): %v (%.2fx slower than hot)", duration3, float64(duration3)/float64(duration2))
	t.Logf("")
	t.Logf("Cache Speedup: %.2fx", float64(duration1)/float64(duration2))

	// Verify results are consistent
	if len(results1) != len(results2) || len(results1) != len(results3) {
		t.Errorf("Result counts differ: %d vs %d vs %d", len(results1), len(results2), len(results3))
	}
}

// TestCalcCellValuesPartialRecalculation tests performance when recalculating a subset of cells
func TestCalcCellValuesPartialRecalculation(t *testing.T) {
	const rows = 5000
	const cols = 100

	f := NewFile()
	sheet := "Sheet1"

	t.Logf("Preparing test data: %d rows x %d cols...", rows, cols)

	// Set base values
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f.SetCellValue(sheet, cell, c*100)
	}

	// Create formulas
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			if r%10 == 0 {
				prevCell, _ := CoordinatesToCellName(c, r-1)
				formula := fmt.Sprintf("=%s+%d", prevCell, c)
				_ = f.SetCellFormula(sheet, cell, formula)
			} else {
				_ = f.SetCellValue(sheet, cell, r*c)
			}
		}
	}

	// Scenario 1: Calculate all cells first
	t.Logf("\n=== Scenario 1: Calculate All Cells ===")
	allCells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			allCells = append(allCells, cell)
		}
	}

	start := time.Now()
	results1, _ := f.CalcCellValues(sheet, allCells)
	duration1 := time.Since(start)

	t.Logf("All cells (%d): %v, %.0f cells/sec", len(allCells), duration1, float64(len(allCells))/duration1.Seconds())
	t.Logf("Calculated: %d/%d", len(results1), len(allCells))

	// Scenario 2: Recalculate only 10% of cells (should be much faster due to cache)
	t.Logf("\n=== Scenario 2: Recalculate 10%% of Cells (Hot Cache) ===")
	partialCells := make([]string, 0, len(allCells)/10)
	for i := 0; i < len(allCells); i += 10 {
		partialCells = append(partialCells, allCells[i])
	}

	start = time.Now()
	results2, _ := f.CalcCellValues(sheet, partialCells)
	duration2 := time.Since(start)

	t.Logf("Partial cells (%d): %v, %.0f cells/sec", len(partialCells), duration2, float64(len(partialCells))/duration2.Seconds())
	t.Logf("Calculated: %d/%d", len(results2), len(partialCells))

	// Scenario 3: Recalculate after modifying one cell
	t.Logf("\n=== Scenario 3: After Modifying One Cell ===")
	_ = f.SetCellValue(sheet, "A1", 99999) // This clears the cache

	start = time.Now()
	results3, _ := f.CalcCellValues(sheet, partialCells)
	duration3 := time.Since(start)

	t.Logf("Partial cells after change (%d): %v, %.0f cells/sec", len(partialCells), duration3, float64(len(partialCells))/duration3.Seconds())
	t.Logf("Calculated: %d/%d", len(results3), len(partialCells))

	t.Logf("\n=== Summary ===")
	t.Logf("Full calculation: %v", duration1)
	t.Logf("Partial (cached):  %v (%.0fx faster per cell)", duration2, (float64(len(allCells))/duration1.Seconds())/(float64(len(partialCells))/duration2.Seconds()))
	t.Logf("Partial (no cache): %v", duration3)
}
