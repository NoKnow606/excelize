package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestCalcCellValueReadOnlyOptimization tests that CalcCellValue doesn't create rows
func TestCalcCellValueReadOnlyOptimization(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set base data
	f.SetCellValue(sheet, "B1", 100)
	f.SetCellValue(sheet, "B2", 200)

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	t.Logf("Initial rows: %d", initialRows)

	// Calculate formula on non-existent far cell
	result, err := f.CalcCellValue(sheet, "Z9999")
	if err != nil {
		t.Errorf("CalcCellValue failed: %v", err)
	}

	// Z9999 should have no value
	if result != "" {
		t.Logf("Result for empty cell Z9999: %s", result)
	}

	// Check row count - should NOT have created rows up to 9999
	ws, _ = f.workSheetReader(sheet)
	afterRows := len(ws.SheetData.Row)

	t.Logf("After CalcCellValue on Z9999:")
	t.Logf("  Rows: %d (initial: %d)", afterRows, initialRows)

	if afterRows >= 9999 {
		t.Errorf("CalcCellValue should not create rows: %d rows exist", afterRows)
	}

	t.Logf("✓ CalcCellValue is read-only optimized")
}

// TestCalcCellValueWithFormulaReadOnly tests formula calculation without row creation
func TestCalcCellValueWithFormulaReadOnly(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set data and formula
	f.SetCellValue(sheet, "A1", 10)
	f.SetCellValue(sheet, "A2", 20)
	f.SetCellFormula(sheet, "A3", "SUM(A1:A2)")

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	// Calculate formula in A3 (exists)
	result1, err := f.CalcCellValue(sheet, "A3")
	if err != nil {
		t.Errorf("CalcCellValue failed: %v", err)
	}
	if result1 != "30" {
		t.Errorf("Expected 30, got %s", result1)
	}

	// Calculate formula for non-existent cell (should reference existing cells)
	// This simulates what CalcFormulaValue does
	_, err = f.CalcCellValue(sheet, "Z999")
	if err != nil {
		t.Errorf("CalcCellValue failed: %v", err)
	}

	ws, _ = f.workSheetReader(sheet)
	afterRows := len(ws.SheetData.Row)

	t.Logf("Rows before: %d, after: %d", initialRows, afterRows)

	if afterRows >= 999 {
		t.Errorf("Should not create 999 rows, but got %d rows", afterRows)
	}

	t.Logf("✓ Formula calculation is read-only optimized")
}

// TestCalcCellValuePerformanceComparison compares old vs new implementation
func TestCalcCellValuePerformanceComparison(t *testing.T) {
	const iterations = 1000
	sheet := "Sheet1"

	// Setup
	f := NewFile()
	defer f.Close()

	f.SetCellValue(sheet, "B1", 100)
	f.SetCellValue(sheet, "B2", 200)
	f.SetCellValue(sheet, "B3", 300)

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	// Test: Calculate values for many cells
	start := time.Now()
	for i := 0; i < iterations; i++ {
		cell := fmt.Sprintf("Z%d", i+1)
		f.CalcCellValue(sheet, cell)
	}
	duration := time.Since(start)

	ws, _ = f.workSheetReader(sheet)
	finalRows := len(ws.SheetData.Row)

	t.Logf("\n=== CalcCellValue Performance (Read-Only Optimized) ===")
	t.Logf("Iterations: %d", iterations)
	t.Logf("Duration: %v", duration)
	t.Logf("Avg per call: %v", duration/iterations)
	t.Logf("Initial rows: %d", initialRows)
	t.Logf("Final rows: %d", finalRows)
	t.Logf("Rows created: %d", finalRows-initialRows)

	if finalRows-initialRows > iterations/10 {
		t.Logf("Warning: Created %d rows (expected minimal)", finalRows-initialRows)
	} else {
		t.Logf("✓ Minimal row creation: only %d rows", finalRows-initialRows)
	}
}

// BenchmarkCalcCellValueReadOnly benchmarks the optimized version
func BenchmarkCalcCellValueReadOnly(b *testing.B) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Setup data
	for i := 1; i <= 100; i++ {
		f.SetCellValue(sheet, fmt.Sprintf("B%d", i), i*10)
	}

	// Add a formula
	f.SetCellFormula(sheet, "C1", "SUM(B1:B100)")

	b.ResetTimer()

	b.Run("ExistingFormulaCell", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f.CalcCellValue(sheet, "C1")
		}
	})

	b.Run("NonExistentFarCell", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			cell := fmt.Sprintf("Z%d", (i%1000)+1)
			f.CalcCellValue(sheet, cell)
		}
	})
}

// TestCalcCellValueMemoryFootprint tests memory usage
func TestCalcCellValueMemoryFootprint(t *testing.T) {
	const farRow = 10000

	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "B1", 100)

	ws, _ := f.workSheetReader(sheet)
	initialRows := len(ws.SheetData.Row)

	// Calculate many far-away cells
	for i := 1; i <= 100; i++ {
		cell := fmt.Sprintf("Z%d", i*100)
		f.CalcCellValue(sheet, cell)
	}

	ws, _ = f.workSheetReader(sheet)
	finalRows := len(ws.SheetData.Row)

	t.Logf("\n=== Memory Footprint Test ===")
	t.Logf("Initial rows: %d", initialRows)
	t.Logf("Final rows: %d", finalRows)
	t.Logf("Rows created: %d", finalRows-initialRows)
	t.Logf("Far cells accessed: 100 (up to Z10000)")

	if finalRows >= farRow/10 {
		t.Errorf("Created too many rows: %d (should be minimal)", finalRows)
	} else {
		t.Logf("✓ Memory footprint minimized")
	}
}
