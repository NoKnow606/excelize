package excelize

import (
	"fmt"
	"strconv"
	"testing"
	"time"
)

// TestSetCellValuesBatchPerformance tests the performance improvement of SetCellValues
func TestSetCellValuesBatchPerformance(t *testing.T) {
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
			compareBatchPerformance(t, scenario.rows, scenario.cols)
		})
	}
}

func compareBatchPerformance(t *testing.T, rows, cols int) {
	totalCells := rows * cols

	// Test 1: Original approach (SetCellValue in loop)
	t.Logf("\n=== Test 1: SetCellValue in Loop (Original) ===")
	f1 := NewFile()
	sheet := "Sheet1"

	start := time.Now()
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			_ = f1.SetCellValue(sheet, cell, r*c)
		}
	}
	duration1 := time.Since(start)

	t.Logf("Total cells: %d", totalCells)
	t.Logf("Duration: %v", duration1)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration1.Seconds())

	// Test 2: Batch approach (SetCellValues)
	t.Logf("\n=== Test 2: SetCellValues (Batch) ===")
	f2 := NewFile()

	// Prepare batch data
	values := make(map[string]interface{}, totalCells)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			values[cell] = r * c
		}
	}

	start = time.Now()
	err := f2.SetCellValues(sheet, values)
	duration2 := time.Since(start)

	if err != nil {
		t.Errorf("SetCellValues failed: %v", err)
	}

	t.Logf("Total cells: %d", totalCells)
	t.Logf("Duration: %v", duration2)
	t.Logf("Throughput: %.0f cells/sec", float64(totalCells)/duration2.Seconds())

	// Performance comparison
	t.Logf("\n=== Performance Summary ===")
	t.Logf("Total cells: %d", totalCells)
	t.Logf("SetCellValue (loop): %v", duration1)
	t.Logf("SetCellValues (batch): %v", duration2)
	t.Logf("Speedup: %.2fx faster", float64(duration1)/float64(duration2))

	// Verify correctness - sample check
	for i := 0; i < 10; i++ {
		r := i*rows/10 + 1
		c := i*cols/10 + 1
		cell, _ := CoordinatesToCellName(c, r)
		v1, _ := f1.GetCellValue(sheet, cell)
		v2, _ := f2.GetCellValue(sheet, cell)
		if v1 != v2 {
			t.Errorf("Mismatch at %s: %s vs %s", cell, v1, v2)
		}
	}
}

// TestSetCellValuesWithFormulas tests batch setting with formulas
func TestSetCellValuesWithFormulas(t *testing.T) {
	const rows = 1000
	const cols = 10

	f := NewFile()
	sheet := "Sheet1"

	t.Logf("Testing batch mode with formulas...")

	// First, set some base values
	values1 := make(map[string]interface{}, rows*cols/2)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols/2; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			values1[cell] = r * c * 100
		}
	}

	start := time.Now()
	err := f.SetCellValues(sheet, values1)
	duration1 := time.Since(start)

	if err != nil {
		t.Errorf("First batch failed: %v", err)
	}

	t.Logf("Set %d base values in %v", len(values1), duration1)

	// Then set formulas that reference the base values
	for r := 1; r <= rows; r++ {
		for c := cols/2 + 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			refCell, _ := CoordinatesToCellName(c-cols/2, r)
			formula := fmt.Sprintf("=%s*2", refCell)
			_ = f.SetCellFormula(sheet, cell, formula)
		}
	}

	t.Logf("Set %d formulas", rows*cols/2)

	// Now update base values in batch
	values2 := make(map[string]interface{}, rows*cols/2)
	for r := 1; r <= rows; r++ {
		for c := 1; c <= cols/2; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			values2[cell] = r * c * 200 // Different values
		}
	}

	start = time.Now()
	err = f.SetCellValues(sheet, values2)
	duration2 := time.Since(start)

	if err != nil {
		t.Errorf("Second batch failed: %v", err)
	}

	t.Logf("Updated %d base values in %v", len(values2), duration2)

	// Verify cache was cleared and formulas can be recalculated
	cell1, _ := CoordinatesToCellName(1, 1)
	cell2, _ := CoordinatesToCellName(cols/2+1, 1)

	val1, _ := f.GetCellValue(sheet, cell1)
	result, err := f.CalcCellValue(sheet, cell2)

	if err != nil {
		t.Errorf("Formula calculation failed: %v", err)
	}

	t.Logf("Base value: %s, Formula result: %s", val1, result)

	expected := strconv.Itoa(200 * 2) // 1*1*200*2
	if result != expected {
		t.Errorf("Formula result incorrect: got %s, want %s", result, expected)
	}
}

// TestSetCellValuesMixedTypes tests batch setting with mixed data types
func TestSetCellValuesMixedTypes(t *testing.T) {
	f := NewFile()
	sheet := "Sheet1"

	values := map[string]interface{}{
		"A1": 100,
		"A2": 3.14,
		"A3": "Hello",
		"A4": true,
		"A5": time.Date(2025, 1, 1, 0, 0, 0, 0, time.UTC),
		"B1": int64(9999999999),
		"B2": float32(2.718),
		"B3": []byte("World"),
		"B4": nil,
	}

	err := f.SetCellValues(sheet, values)
	if err != nil {
		t.Errorf("SetCellValues with mixed types failed: %v", err)
	}

	// Verify all values
	tests := []struct {
		cell     string
		expected string
	}{
		{"A1", "100"},
		{"A2", "3.14"},
		{"A3", "Hello"},
		{"A4", "TRUE"},
		{"A5", "44927"},
		{"B1", "9999999999"},
		{"B3", "World"},
		{"B4", ""},
	}

	for _, tt := range tests {
		val, err := f.GetCellValue(sheet, tt.cell)
		if err != nil {
			t.Errorf("GetCellValue(%s) failed: %v", tt.cell, err)
		}
		if val != tt.expected {
			t.Errorf("Cell %s: got %s, want %s", tt.cell, val, tt.expected)
		}
	}

	t.Logf("All mixed type values set correctly")
}
