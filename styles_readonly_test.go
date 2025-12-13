package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestGetCellStyleReadOnly tests the basic functionality
func TestGetCellStyleReadOnly(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Test 1: Cell with no style (doesn't exist)
	style, err := f.GetCellStyleReadOnly(sheet, "A1")
	if err != nil {
		t.Errorf("GetCellStyleReadOnly failed: %v", err)
	}
	if style != 0 {
		t.Errorf("Expected style 0 for non-existent cell, got %d", style)
	}

	// Test 2: Set a cell value with style
	styleIdx, err := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	if err != nil {
		t.Fatalf("NewStyle failed: %v", err)
	}

	f.SetCellValue(sheet, "B1", "Test")
	f.SetCellStyle(sheet, "B1", "B1", styleIdx)

	// Test 3: Read existing cell style
	readStyle, err := f.GetCellStyleReadOnly(sheet, "B1")
	if err != nil {
		t.Errorf("GetCellStyleReadOnly failed: %v", err)
	}
	if readStyle != styleIdx {
		t.Errorf("Expected style %d, got %d", styleIdx, readStyle)
	}

	// Test 4: Read cell in far-away location (shouldn't create rows/cols)
	farStyle, err := f.GetCellStyleReadOnly(sheet, "ZZ9999")
	if err != nil {
		t.Errorf("GetCellStyleReadOnly failed for far cell: %v", err)
	}
	if farStyle != 0 {
		t.Errorf("Expected style 0 for far cell, got %d", farStyle)
	}

	// Verify that ZZ9999 row was NOT created
	ws, _ := f.workSheetReader(sheet)
	if len(ws.SheetData.Row) >= 9999 {
		t.Errorf("GetCellStyleReadOnly should not create rows, but row count is %d", len(ws.SheetData.Row))
	}
}

// TestGetCellStyleReadOnlyInheritance tests style inheritance from row/column
func TestGetCellStyleReadOnlyInheritance(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create a row style
	rowStyle, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#00FF00"}, Pattern: 1},
	})

	// Set row style manually (simulate row default style)
	ws, _ := f.workSheetReader(sheet)
	ws.prepareSheetXML(1, 1)
	ws.SheetData.Row[0].S = rowStyle

	// Test: Cell should inherit row style
	style, err := f.GetCellStyleReadOnly(sheet, "A1")
	if err != nil {
		t.Errorf("GetCellStyleReadOnly failed: %v", err)
	}
	if style != rowStyle {
		t.Errorf("Expected inherited row style %d, got %d", rowStyle, style)
	}

	t.Logf("Row style inheritance works: cell A1 inherited style %d from row", style)
}

// TestGetCellStyleReadOnlyVsNormal compares behavior
func TestGetCellStyleReadOnlyVsNormal(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set style on B2
	styleIdx, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#0000FF"}, Pattern: 1},
	})
	f.SetCellValue(sheet, "B2", 100)
	f.SetCellStyle(sheet, "B2", "B2", styleIdx)

	// Test existing cell: both methods should return same result
	normalStyle, _ := f.GetCellStyle(sheet, "B2")
	readOnlyStyle, _ := f.GetCellStyleReadOnly(sheet, "B2")

	if normalStyle != readOnlyStyle {
		t.Errorf("Style mismatch: GetCellStyle=%d, GetCellStyleReadOnly=%d", normalStyle, readOnlyStyle)
	}

	// Count rows before reading far cell
	ws, _ := f.workSheetReader(sheet)
	rowCountBefore := len(ws.SheetData.Row)

	// Read far cell with ReadOnly (should NOT create rows)
	f.GetCellStyleReadOnly(sheet, "Z999")

	rowCountAfterReadOnly := len(ws.SheetData.Row)

	if rowCountAfterReadOnly != rowCountBefore {
		t.Errorf("GetCellStyleReadOnly created rows: before=%d, after=%d", rowCountBefore, rowCountAfterReadOnly)
	}

	// Read far cell with normal GetCellStyle (WILL create rows)
	f.GetCellStyle(sheet, "Z999")

	rowCountAfterNormal := len(ws.SheetData.Row)

	if rowCountAfterNormal <= rowCountBefore {
		t.Errorf("GetCellStyle should create rows: before=%d, after=%d", rowCountBefore, rowCountAfterNormal)
	}

	t.Logf("Behavior comparison:")
	t.Logf("  Initial rows: %d", rowCountBefore)
	t.Logf("  After GetCellStyleReadOnly: %d (no change ✓)", rowCountAfterReadOnly)
	t.Logf("  After GetCellStyle: %d (created rows)", rowCountAfterNormal)
}

// TestGetCellStyleReadOnlyErrors tests error handling
func TestGetCellStyleReadOnlyErrors(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Test 1: Invalid sheet name
	_, err := f.GetCellStyleReadOnly("NonExistentSheet", "A1")
	if err == nil {
		t.Error("Expected error for non-existent sheet")
	}

	// Test 2: Invalid cell reference
	_, err = f.GetCellStyleReadOnly("Sheet1", "INVALID")
	if err == nil {
		t.Error("Expected error for invalid cell reference")
	}

	t.Logf("Error handling works correctly")
}

// BenchmarkGetCellStyleReadOnly benchmarks the performance difference
func BenchmarkGetCellStyleReadOnly(b *testing.B) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Setup: Create some data
	for i := 1; i <= 100; i++ {
		cell := fmt.Sprintf("A%d", i)
		f.SetCellValue(sheet, cell, i)
	}

	b.Run("GetCellStyle", func(b *testing.B) {
		f1 := NewFile()
		defer f1.Close()

		for i := 1; i <= 100; i++ {
			cell := fmt.Sprintf("A%d", i)
			f1.SetCellValue(sheet, cell, i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			cell := fmt.Sprintf("Z%d", (i%1000)+1)
			f1.GetCellStyle(sheet, cell)
		}
	})

	b.Run("GetCellStyleReadOnly", func(b *testing.B) {
		f2 := NewFile()
		defer f2.Close()

		for i := 1; i <= 100; i++ {
			cell := fmt.Sprintf("A%d", i)
			f2.SetCellValue(sheet, cell, i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			cell := fmt.Sprintf("Z%d", (i%1000)+1)
			f2.GetCellStyleReadOnly(sheet, cell)
		}
	})
}

// TestGetCellStyleReadOnlyPerformance tests performance with timing
func TestGetCellStyleReadOnlyPerformance(t *testing.T) {
	const iterations = 10000
	sheet := "Sheet1"

	// Test 1: GetCellStyle (creates rows/cols)
	f1 := NewFile()
	defer f1.Close()

	start := time.Now()
	for i := 0; i < iterations; i++ {
		cell := fmt.Sprintf("Z%d", (i%1000)+1)
		f1.GetCellStyle(sheet, cell)
	}
	duration1 := time.Since(start)

	// Test 2: GetCellStyleReadOnly (read-only)
	f2 := NewFile()
	defer f2.Close()

	start = time.Now()
	for i := 0; i < iterations; i++ {
		cell := fmt.Sprintf("Z%d", (i%1000)+1)
		f2.GetCellStyleReadOnly(sheet, cell)
	}
	duration2 := time.Since(start)

	// Report results
	t.Logf("\n=== Performance Comparison (%d iterations) ===", iterations)
	t.Logf("GetCellStyle:         %v (%.2f μs/op)", duration1, float64(duration1.Microseconds())/float64(iterations))
	t.Logf("GetCellStyleReadOnly: %v (%.2f μs/op)", duration2, float64(duration2.Microseconds())/float64(iterations))
	t.Logf("Speedup:              %.2fx faster", float64(duration1)/float64(duration2))

	// Check memory usage
	ws1, _ := f1.workSheetReader(sheet)
	ws2, _ := f2.workSheetReader(sheet)

	t.Logf("\n=== Memory Usage ===")
	t.Logf("GetCellStyle rows created:         %d", len(ws1.SheetData.Row))
	t.Logf("GetCellStyleReadOnly rows created: %d", len(ws2.SheetData.Row))

	if len(ws2.SheetData.Row) > 0 {
		t.Error("GetCellStyleReadOnly should not create any rows")
	}
}
