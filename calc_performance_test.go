package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestCalcCellValuesPerformanceScaling tests performance at different scales
func TestCalcCellValuesPerformanceScaling(t *testing.T) {
	scenarios := []struct {
		name string
		rows int
		cols int
	}{
		{"Small (1k x 10)", 1000, 10},
		{"Medium (5k x 50)", 5000, 50},
		{"Large (10k x 100)", 10000, 100},
		{"Extra Large (40k x 100)", 40000, 100},
	}

	for _, scenario := range scenarios {
		t.Run(scenario.name, func(t *testing.T) {
			runPerformanceTest(t, scenario.rows, scenario.cols)
		})
	}
}

func runPerformanceTest(t *testing.T, rows, cols int) {
	f := NewFile()
	sheet := "Sheet1"

	t.Logf("Preparing test data: %d rows x %d cols...", rows, cols)
	startPrep := time.Now()

	// Set base values in first row
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f.SetCellValue(sheet, cell, c*100)
	}

	// Set formulas and values for remaining rows
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

	prepDuration := time.Since(startPrep)
	t.Logf("Data preparation: %v", prepDuration)

	// Build list of cells to calculate
	cells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			cells = append(cells, cell)
		}
	}

	totalCells := len(cells)

	// Test CalcCellValues (batch)
	start := time.Now()
	results, err := f.CalcCellValues(sheet, cells)
	batchDuration := time.Since(start)

	if err != nil {
		t.Logf("Warning: %v", err)
	}

	// Test individual CalcCellValue calls for comparison (sample 1000 cells)
	sampleSize := 1000
	if totalCells < sampleSize {
		sampleSize = totalCells
	}

	start = time.Now()
	for i := 0; i < sampleSize; i++ {
		_, _ = f.CalcCellValue(sheet, cells[i])
	}
	singleDuration := time.Since(start)
	projectedSingleDuration := time.Duration(float64(singleDuration) * float64(totalCells) / float64(sampleSize))

	t.Logf("\n=== Performance Results ===")
	t.Logf("Total cells: %d", totalCells)
	t.Logf("Successful: %d, Failed: %d", len(results), totalCells-len(results))
	t.Logf("")
	t.Logf("Batch CalcCellValues:")
	t.Logf("  Duration: %v", batchDuration)
	t.Logf("  Throughput: %.0f cells/sec", float64(totalCells)/batchDuration.Seconds())
	t.Logf("  Avg per cell: %v", batchDuration/time.Duration(totalCells))
	t.Logf("")
	t.Logf("Individual CalcCellValue (projected from %d samples):", sampleSize)
	t.Logf("  Projected duration: %v", projectedSingleDuration)
	t.Logf("  Projected throughput: %.0f cells/sec", float64(totalCells)/projectedSingleDuration.Seconds())
	t.Logf("")
	t.Logf("Speedup: %.2fx faster with batch", float64(projectedSingleDuration)/float64(batchDuration))
}
