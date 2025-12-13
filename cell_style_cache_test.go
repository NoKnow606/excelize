package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestPrepareCellStyleCache tests column style caching
func TestPrepareCellStyleCache(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create column styles
	style1, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	style2, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#00FF00"}, Pattern: 1},
	})

	// Set column styles
	f.SetColStyle(sheet, "A:C", style1)
	f.SetColStyle(sheet, "D:F", style2)

	ws, _ := f.workSheetReader(sheet)

	// Test 1: Get style for column in range (should cache)
	s1 := ws.prepareCellStyle(1, 1, 0) // Column A
	if s1 != style1 {
		t.Errorf("Expected style %d for column A, got %d", style1, s1)
	}

	// Test 2: Get same column again (should use cache)
	s2 := ws.prepareCellStyle(1, 1, 0) // Column A again
	if s2 != style1 {
		t.Errorf("Expected cached style %d, got %d", style1, s2)
	}

	// Test 3: Get style for different column
	s3 := ws.prepareCellStyle(4, 1, 0) // Column D
	if s3 != style2 {
		t.Errorf("Expected style %d for column D, got %d", style2, s3)
	}

	// Test 4: Verify cache is populated
	if cachedStyle, ok := ws.colStyleCache.Load(1); !ok {
		t.Error("Column 1 should be cached")
	} else if cachedStyle.(int) != style1 {
		t.Errorf("Cached style mismatch: expected %d, got %d", style1, cachedStyle.(int))
	}

	t.Logf("✓ Column style caching works correctly")
}

// TestPrepareCellStyleCacheInvalidation tests cache clearing
func TestPrepareCellStyleCacheInvalidation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	style1, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	style2, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#00FF00"}, Pattern: 1},
	})

	// Set initial style
	f.SetColStyle(sheet, "A:A", style1)

	ws, _ := f.workSheetReader(sheet)

	// Get style (should cache)
	s1 := ws.prepareCellStyle(1, 1, 0)
	if s1 != style1 {
		t.Errorf("Expected style %d, got %d", style1, s1)
	}

	// Verify cached
	if _, ok := ws.colStyleCache.Load(1); !ok {
		t.Error("Column 1 should be cached")
	}

	// Change column style
	f.SetColStyle(sheet, "A:A", style2)

	// Verify cache was cleared
	if _, ok := ws.colStyleCache.Load(1); ok {
		t.Error("Cache should be cleared after SetColStyle")
	}

	// Get style again (should use new style)
	s2 := ws.prepareCellStyle(1, 1, 0)
	if s2 != style2 {
		t.Errorf("Expected new style %d, got %d", style2, s2)
	}

	t.Logf("✓ Cache invalidation works correctly")
}

// TestPrepareCellStylePerformance benchmarks the caching improvement
func TestPrepareCellStylePerformance(t *testing.T) {
	const iterations = 100000
	const numCols = 100

	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create many column styles
	for i := 1; i <= numCols; i++ {
		style, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i*1000)}, Pattern: 1},
		})
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, style)
	}

	ws, _ := f.workSheetReader(sheet)

	// Test: Repeatedly access column styles (simulates GetCellStyleReadOnly usage)
	start := time.Now()
	for i := 0; i < iterations; i++ {
		col := (i % numCols) + 1
		ws.prepareCellStyle(col, 1, 0)
	}
	duration := time.Since(start)

	t.Logf("\n=== Column Style Cache Performance ===")
	t.Logf("Columns: %d", numCols)
	t.Logf("Iterations: %d", iterations)
	t.Logf("Duration: %v", duration)
	t.Logf("Avg per call: %v", duration/iterations)
	t.Logf("Throughput: %.0f calls/sec", float64(iterations)/duration.Seconds())

	// Verify cache hit rate
	cacheHits := 0
	ws.colStyleCache.Range(func(key, value interface{}) bool {
		cacheHits++
		return true
	})

	t.Logf("Cache entries: %d (expected %d)", cacheHits, numCols)

	if cacheHits != numCols {
		t.Errorf("Expected %d cache entries, got %d", numCols, cacheHits)
	}
}

// BenchmarkPrepareCellStyle benchmarks with and without caching
func BenchmarkPrepareCellStyle(b *testing.B) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create 50 column styles
	for i := 1; i <= 50; i++ {
		style, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i*1000)}, Pattern: 1},
		})
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, style)
	}

	ws, _ := f.workSheetReader(sheet)

	b.ResetTimer()

	b.Run("WithCache", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			col := (i % 50) + 1
			ws.prepareCellStyle(col, 1, 0)
		}
	})

	b.Run("ClearCacheBetweenCalls", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			col := (i % 50) + 1
			ws.colStyleCache.Delete(col) // Simulate cache miss
			ws.prepareCellStyle(col, 1, 0)
		}
	})
}

// TestPrepareCellStylePriorityOrder tests style priority
func TestPrepareCellStylePriorityOrder(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	cellStyle, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	rowStyle, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#00FF00"}, Pattern: 1},
	})
	colStyle, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#0000FF"}, Pattern: 1},
	})

	// Set column style
	f.SetColStyle(sheet, "A:A", colStyle)

	ws, _ := f.workSheetReader(sheet)

	// Test 1: Cell style has highest priority
	style := ws.prepareCellStyle(1, 1, cellStyle)
	if style != cellStyle {
		t.Errorf("Cell style should have priority: expected %d, got %d", cellStyle, style)
	}

	// Test 2: Row style has second priority (simulate by setting row style)
	ws.prepareSheetXML(1, 1)
	ws.SheetData.Row[0].S = rowStyle
	style = ws.prepareCellStyle(1, 1, 0)
	if style != rowStyle {
		t.Errorf("Row style should have priority: expected %d, got %d", rowStyle, style)
	}

	// Test 3: Column style has third priority
	ws.SheetData.Row[0].S = 0 // Clear row style
	style = ws.prepareCellStyle(1, 1, 0)
	if style != colStyle {
		t.Errorf("Column style should be used: expected %d, got %d", colStyle, style)
	}

	t.Logf("✓ Style priority order correct: cell > row > column")
}

// TestPrepareCellStyleEdgeCases tests edge cases
func TestPrepareCellStyleEdgeCases(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	ws, _ := f.workSheetReader(sheet)

	// Test 1: No columns defined
	style := ws.prepareCellStyle(1, 1, 0)
	if style != 0 {
		t.Errorf("Expected 0 when no columns defined, got %d", style)
	}

	// Test 2: Empty Cols
	ws.Cols = &xlsxCols{}
	style = ws.prepareCellStyle(1, 1, 0)
	if style != 0 {
		t.Errorf("Expected 0 with empty Cols, got %d", style)
	}

	// Test 3: Column not in any range
	colStyle, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	f.SetColStyle(sheet, "A:C", colStyle)

	ws, _ = f.workSheetReader(sheet)
	style = ws.prepareCellStyle(10, 1, 0) // Column J (not styled)
	if style != 0 {
		t.Errorf("Expected 0 for unstyled column, got %d", style)
	}

	// Verify cache stores 0 for unstyled column
	if cachedStyle, ok := ws.colStyleCache.Load(10); !ok {
		t.Error("Unstyled column should be cached as 0")
	} else if cachedStyle.(int) != 0 {
		t.Errorf("Expected cached 0, got %d", cachedStyle.(int))
	}

	t.Logf("✓ Edge cases handled correctly")
}
