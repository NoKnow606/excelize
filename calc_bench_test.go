package excelize

import (
	"fmt"
	"testing"
	"time"
)

// BenchmarkCalcCellValues40kx100 tests CalcCellValues performance with 40k rows * 100 columns
func BenchmarkCalcCellValues40kx100(b *testing.B) {
	const rows = 40000
	const cols = 100

	f := NewFile()
	sheet := "Sheet1"

	// Prepare test data: populate cells with values and formulas
	fmt.Printf("Preparing test data: %d rows x %d cols...\n", rows, cols)
	startPrep := time.Now()

	// Set base values in first row
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f.SetCellValue(sheet, cell, c*100)
	}

	// Set formulas and values for remaining rows
	// Every 10th row uses a formula, others use static values for realistic mix
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			if r%10 == 0 {
				// Formula: reference cell above + column number
				prevCell, _ := CoordinatesToCellName(c, r-1)
				formula := fmt.Sprintf("=%s+%d", prevCell, c)
				_ = f.SetCellFormula(sheet, cell, formula)
			} else {
				// Static value
				_ = f.SetCellValue(sheet, cell, r*c)
			}
		}

		if r%5000 == 0 {
			fmt.Printf("  Prepared %d/%d rows (%.1f%%)...\n", r, rows, float64(r)/float64(rows)*100)
		}
	}

	prepDuration := time.Since(startPrep)
	fmt.Printf("Data preparation completed in %v\n", prepDuration)

	// Build list of cells to calculate
	cells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			cells = append(cells, cell)
		}
	}

	totalCells := len(cells)
	fmt.Printf("\nStarting benchmark with %d total cells...\n", totalCells)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		start := time.Now()
		results, err := f.CalcCellValues(sheet, cells)
		duration := time.Since(start)

		if err != nil {
			b.Logf("CalcCellValues returned error: %v", err)
		}

		fmt.Printf("\nBenchmark iteration %d:\n", i+1)
		fmt.Printf("  Total cells: %d\n", totalCells)
		fmt.Printf("  Successful calculations: %d\n", len(results))
		fmt.Printf("  Failed calculations: %d\n", totalCells-len(results))
		fmt.Printf("  Duration: %v\n", duration)
		fmt.Printf("  Cells/sec: %.0f\n", float64(totalCells)/duration.Seconds())
		fmt.Printf("  Avg per cell: %v\n", duration/time.Duration(totalCells))
	}
}

// TestCalcCellValuesPerformance is a regular test that can be run with go test
func TestCalcCellValuesPerformance(t *testing.T) {
	const rows = 40000
	const cols = 100

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

		if r%5000 == 0 {
			t.Logf("  Prepared %d/%d rows (%.1f%%)...", r, rows, float64(r)/float64(rows)*100)
		}
	}

	prepDuration := time.Since(startPrep)
	t.Logf("Data preparation completed in %v", prepDuration)

	// Build list of cells to calculate
	cells := make([]string, 0, rows*cols)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			cells = append(cells, cell)
		}
	}

	totalCells := len(cells)
	t.Logf("Starting calculation of %d cells...", totalCells)

	start := time.Now()
	results, err := f.CalcCellValues(sheet, cells)
	duration := time.Since(start)

	if err != nil {
		t.Logf("CalcCellValues returned error: %v", err)
	}

	t.Logf("\n=== Performance Results ===")
	t.Logf("Total cells: %d", totalCells)
	t.Logf("Successful calculations: %d", len(results))
	t.Logf("Failed calculations: %d", totalCells-len(results))
	t.Logf("Total duration: %v", duration)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration.Seconds())
	t.Logf("Average per cell: %v", duration/time.Duration(totalCells))
}
