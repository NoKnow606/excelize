package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestSetCellValuesRealWorldScenario tests a realistic scenario where we replace formulas with values
func TestSetCellValuesRealWorldScenario(t *testing.T) {
	// Scenario: User has a spreadsheet with formulas, then wants to "paste values" to replace formulas with their calculated values
	// This is a common operation that triggers cache clearing multiple times

	const rows = 5000
	const cols = 100

	t.Logf("Real-world scenario: Replace %d formulas with values", rows*cols)

	// Setup: Create a file with formulas everywhere
	f1 := NewFile()
	sheet := "Sheet1"

	// Set base values in first row
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f1.SetCellValue(sheet, cell, c*100)
	}

	// Set formulas in all other rows
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			prevCell, _ := CoordinatesToCellName(c, r-1)
			formula := fmt.Sprintf("=%s+%d", prevCell, r)
			_ = f1.SetCellFormula(sheet, cell, formula)
		}
	}

	t.Logf("Setup complete: %d cells with formulas", (rows-1)*cols)

	// Test 1: Replace formulas with values using SetCellValue in loop
	t.Logf("\n=== Test 1: Replace Formulas with SetCellValue Loop ===")
	f2 := NewFile()
	// Copy same setup
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f2.SetCellValue(sheet, cell, c*100)
	}
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			prevCell, _ := CoordinatesToCellName(c, r-1)
			formula := fmt.Sprintf("=%s+%d", prevCell, r)
			_ = f2.SetCellFormula(sheet, cell, formula)
		}
	}

	// Now replace formulas with values
	start := time.Now()
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			// Replace formula with value
			_ = f2.SetCellValue(sheet, cell, r*c)
		}
	}
	duration1 := time.Since(start)

	totalOps := (rows - 1) * cols
	t.Logf("Replaced %d formulas", totalOps)
	t.Logf("Duration: %v", duration1)
	t.Logf("Throughput: %.0f ops/sec", float64(totalOps)/duration1.Seconds())

	// Test 2: Replace formulas with values using SetCellValues batch
	t.Logf("\n=== Test 2: Replace Formulas with SetCellValues Batch ===")
	f3 := NewFile()
	// Copy same setup
	for c := 1; c <= cols; c++ {
		cell, _ := CoordinatesToCellName(c, 1)
		_ = f3.SetCellValue(sheet, cell, c*100)
	}
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			prevCell, _ := CoordinatesToCellName(c, r-1)
			formula := fmt.Sprintf("=%s+%d", prevCell, r)
			_ = f3.SetCellFormula(sheet, cell, formula)
		}
	}

	// Now replace formulas with values using batch
	start = time.Now()
	values := make(map[string]interface{}, totalOps)
	for r := 2; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			values[cell] = r * c
		}
	}
	err := f3.SetCellValues(sheet, values)
	duration2 := time.Since(start)

	if err != nil {
		t.Errorf("SetCellValues failed: %v", err)
	}

	t.Logf("Replaced %d formulas", totalOps)
	t.Logf("Duration: %v", duration2)
	t.Logf("Throughput: %.0f ops/sec", float64(totalOps)/duration2.Seconds())

	// Performance comparison
	t.Logf("\n=== Performance Summary ===")
	t.Logf("Scenario: Replacing %d formulas with values", totalOps)
	t.Logf("SetCellValue (loop):   %v (%.0f ops/sec)", duration1, float64(totalOps)/duration1.Seconds())
	t.Logf("SetCellValues (batch): %v (%.0f ops/sec)", duration2, float64(totalOps)/duration2.Seconds())
	if duration2 < duration1 {
		t.Logf("Speedup: %.2fx faster ✓", float64(duration1)/float64(duration2))
	} else {
		t.Logf("Speedup: %.2fx (batch is slower)", float64(duration1)/float64(duration2))
	}
}
