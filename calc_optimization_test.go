package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestCalcCellValuesOptimizationComparison compares different optimization strategies
func TestCalcCellValuesOptimizationComparison(t *testing.T) {
	scenarios := []struct {
		name string
		rows int
		cols int
	}{
		{"Medium (5k x 50)", 5000, 50},
		{"Large (10k x 100)", 10000, 100},
		{"Extra Large (20k x 100)", 20000, 100},
	}

	for _, scenario := range scenarios {
		t.Run(scenario.name, func(t *testing.T) {
			compareOptimizations(t, scenario.rows, scenario.cols)
		})
	}
}

func compareOptimizations(t *testing.T, rows, cols int) {
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

	// Build list of cells
	cells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			cells = append(cells, cell)
		}
	}

	totalCells := len(cells)

	// Test 1: Original implementation
	t.Logf("\n=== Test 1: Original CalcCellValues ===")
	start := time.Now()
	results1, err1 := f.CalcCellValues(sheet, cells)
	duration1 := time.Since(start)

	if err1 != nil {
		t.Logf("Warning: %v", err1)
	}
	t.Logf("Duration: %v", duration1)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration1.Seconds())
	t.Logf("Successful: %d/%d", len(results1), totalCells)

	// Test 2: Optimized implementation (with strings.Builder)
	t.Logf("\n=== Test 2: Optimized CalcCellValues (strings.Builder) ===")
	// Clear cache to have fair comparison
	f.calcCache.Clear()

	start = time.Now()
	results2, err2 := f.CalcCellValuesOptimized(sheet, cells)
	duration2 := time.Since(start)

	if err2 != nil {
		t.Logf("Warning: %v", err2)
	}
	t.Logf("Duration: %v", duration2)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration2.Seconds())
	t.Logf("Successful: %d/%d", len(results2), totalCells)
	t.Logf("Speedup vs original: %.2fx", float64(duration1)/float64(duration2))

	// Test 3: Concurrent implementation
	t.Logf("\n=== Test 3: Concurrent CalcCellValues ===")
	// Clear cache to have fair comparison
	f.calcCache.Clear()

	start = time.Now()
	results3, err3 := f.CalcCellValuesConcurrent(sheet, cells)
	duration3 := time.Since(start)

	if err3 != nil {
		t.Logf("Warning: %v", err3)
	}
	t.Logf("Duration: %v", duration3)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration3.Seconds())
	t.Logf("Successful: %d/%d", len(results3), totalCells)
	t.Logf("Speedup vs original: %.2fx", float64(duration1)/float64(duration3))

	// Test 4: Concurrent with hot cache (best case)
	t.Logf("\n=== Test 4: Concurrent CalcCellValues (Hot Cache) ===")

	start = time.Now()
	results4, err4 := f.CalcCellValuesConcurrent(sheet, cells)
	duration4 := time.Since(start)

	if err4 != nil {
		t.Logf("Warning: %v", err4)
	}
	t.Logf("Duration: %v", duration4)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration4.Seconds())
	t.Logf("Successful: %d/%d", len(results4), totalCells)
	t.Logf("Speedup vs original: %.2fx", float64(duration1)/float64(duration4))

	// Summary
	t.Logf("\n=== Performance Summary ===")
	t.Logf("Total cells: %d", totalCells)
	t.Logf("")
	t.Logf("Original (cold cache):     %v (baseline)", duration1)
	t.Logf("Optimized (cold cache):    %v (%.2fx)", duration2, float64(duration1)/float64(duration2))
	t.Logf("Concurrent (cold cache):   %v (%.2fx)", duration3, float64(duration1)/float64(duration3))
	t.Logf("Concurrent (hot cache):    %v (%.2fx)", duration4, float64(duration1)/float64(duration4))
	t.Logf("")
	t.Logf("Best improvement: Concurrent with hot cache is %.2fx faster", float64(duration1)/float64(duration4))

	// Verify results consistency
	if len(results1) != len(results2) || len(results1) != len(results3) || len(results1) != len(results4) {
		t.Errorf("Result counts differ: %d vs %d vs %d vs %d", len(results1), len(results2), len(results3), len(results4))
	}
}

// BenchmarkCalcStrategies benchmarks different calculation strategies
func BenchmarkCalcStrategies(b *testing.B) {
	const rows = 1000
	const cols = 100

	f := NewFile()
	sheet := "Sheet1"

	// Setup data
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f.SetCellValue(sheet, cell, c*100)
	}

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

	cells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			cells = append(cells, cell)
		}
	}

	b.Run("Original", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f.calcCache.Clear()
			_, _ = f.CalcCellValues(sheet, cells)
		}
	})

	b.Run("Optimized", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f.calcCache.Clear()
			_, _ = f.CalcCellValuesOptimized(sheet, cells)
		}
	})

	b.Run("Concurrent", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f.calcCache.Clear()
			_, _ = f.CalcCellValuesConcurrent(sheet, cells)
		}
	})
}
