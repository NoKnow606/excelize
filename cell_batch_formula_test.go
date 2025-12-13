package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestSetCellValuesBatchWithFormulasPerformance tests the real benefit with formulas
func TestSetCellValuesBatchWithFormulasPerformance(t *testing.T) {
	scenarios := []struct {
		name string
		rows int
		cols int
	}{
		{"Medium (5k x 50)", 5000, 50},
		{"Large (10k x 100)", 10000, 100},
	}

	for _, scenario := range scenarios {
		t.Run(scenario.name, func(t *testing.T) {
			testBatchWithFormulas(t, scenario.rows, scenario.cols)
		})
	}
}

func testBatchWithFormulas(t *testing.T, rows, cols int) {
	// Test scenario: First half columns are formulas, second half are values
	// We'll update the values multiple times and see the performance difference

	t.Logf("Setup: %d rows x %d cols (50%% formulas, 50%% values)", rows, cols)

	// Test 1: Original approach (SetCellValue in loop)
	t.Logf("\n=== Test 1: SetCellValue in Loop (Original) ===")
	f1 := NewFile()
	sheet := "Sheet1"

	// Set formulas in first half
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols/2; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			refCell, _ := CoordinatesToCellName(c+cols/2, r)
			formula := fmt.Sprintf("=%s*2", refCell)
			_ = f1.SetCellFormula(sheet, cell, formula)
		}
	}

	// Update values in second half multiple times
	iterations := 3
	start := time.Now()
	for iter := 0; iter < iterations; iter++ {
		for r := 1; r <= rows; r++ {
			for c := cols/2 + 1; c <= cols; c++ {
				cell, _ := CoordinatesToCellName(c, r)
				_ = f1.SetCellValue(sheet, cell, r*c*iter)
			}
		}
	}
	duration1 := time.Since(start)

	totalOps := rows * cols / 2 * iterations
	t.Logf("Updated %d values %d times", rows*cols/2, iterations)
	t.Logf("Total operations: %d", totalOps)
	t.Logf("Duration: %v", duration1)
	t.Logf("Throughput: %.0f ops/sec", float64(totalOps)/duration1.Seconds())

	// Test 2: Batch approach (SetCellValues)
	t.Logf("\n=== Test 2: SetCellValues (Batch) ===")
	f2 := NewFile()

	// Set formulas in first half
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols/2; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			refCell, _ := CoordinatesToCellName(c+cols/2, r)
			formula := fmt.Sprintf("=%s*2", refCell)
			_ = f2.SetCellFormula(sheet, cell, formula)
		}
	}

	// Update values in second half multiple times using batch
	start = time.Now()
	for iter := 0; iter < iterations; iter++ {
		values := make(map[string]interface{}, rows*cols/2)
		for r := 1; r <= rows; r++ {
			for c := cols/2 + 1; c <= cols; c++ {
				cell, _ := CoordinatesToCellName(c, r)
				values[cell] = r * c * iter
			}
		}
		_ = f2.SetCellValues(sheet, values)
	}
	duration2 := time.Since(start)

	t.Logf("Updated %d values %d times", rows*cols/2, iterations)
	t.Logf("Total operations: %d", totalOps)
	t.Logf("Duration: %v", duration2)
	t.Logf("Throughput: %.0f ops/sec", float64(totalOps)/duration2.Seconds())

	// Performance comparison
	t.Logf("\n=== Performance Summary ===")
	t.Logf("SetCellValue (loop): %v", duration1)
	t.Logf("SetCellValues (batch): %v", duration2)
	t.Logf("Speedup: %.2fx faster", float64(duration1)/float64(duration2))

	if duration2 > duration1 {
		t.Logf("WARNING: Batch is slower. This might be expected for small datasets without formulas.")
	}
}
