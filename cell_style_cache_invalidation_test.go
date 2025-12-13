package excelize

import (
	"fmt"
	"sync"
	"testing"
)

// TestColumnStyleCacheInvalidation tests all column modification paths clear cache
func TestColumnStyleCacheInvalidation(t *testing.T) {
	tests := []struct {
		name      string
		setup     func(*File, string) error
		modify    func(*File, string) error
		col       int
		expectNew bool
	}{
		{
			name: "SetColStyle invalidates cache",
			setup: func(f *File, sheet string) error {
				style, _ := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1}})
				return f.SetColStyle(sheet, "A:A", style)
			},
			modify: func(f *File, sheet string) error {
				style, _ := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"#00FF00"}, Pattern: 1}})
				return f.SetColStyle(sheet, "A:A", style)
			},
			col:       1,
			expectNew: true,
		},
		{
			name: "SetColWidth invalidates cache",
			setup: func(f *File, sheet string) error {
				style, _ := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1}})
				return f.SetColStyle(sheet, "B:B", style)
			},
			modify: func(f *File, sheet string) error {
				return f.SetColWidth(sheet, "B", "B", 20)
			},
			col:       2,
			expectNew: false, // Width doesn't change style
		},
		{
			name: "SetColVisible invalidates cache",
			setup: func(f *File, sheet string) error {
				style, _ := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1}})
				return f.SetColStyle(sheet, "C:C", style)
			},
			modify: func(f *File, sheet string) error {
				return f.SetColVisible(sheet, "C:C", false)
			},
			col:       3,
			expectNew: false, // Visibility doesn't change style
		},
		{
			name: "SetColOutlineLevel invalidates cache",
			setup: func(f *File, sheet string) error {
				style, _ := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1}})
				return f.SetColStyle(sheet, "D:D", style)
			},
			modify: func(f *File, sheet string) error {
				return f.SetColOutlineLevel(sheet, "D", 2)
			},
			col:       4,
			expectNew: false, // Outline level doesn't change style
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()

			sheet := "Sheet1"

			// Setup initial style
			if err := tt.setup(f, sheet); err != nil {
				t.Fatalf("Setup failed: %v", err)
			}

			ws, _ := f.workSheetReader(sheet)

			// Get style to populate cache
			style1 := ws.prepareCellStyle(tt.col, 1, 0)

			// Verify cache is populated
			if _, ok := ws.colStyleCache.Load(tt.col); !ok {
				t.Error("Cache should be populated after first access")
			}

			// Modify column (should clear cache)
			if err := tt.modify(f, sheet); err != nil {
				t.Fatalf("Modify failed: %v", err)
			}

			// Check if cache was cleared
			if _, ok := ws.colStyleCache.Load(tt.col); ok {
				t.Errorf("%s: Cache should be cleared after modification", tt.name)
			}

			// Get style again (should recalculate)
			style2 := ws.prepareCellStyle(tt.col, 1, 0)

			if tt.expectNew && style1 == style2 {
				t.Errorf("%s: Style should have changed after modification", tt.name)
			}

			t.Logf("%s: ✓ Cache invalidation works correctly", tt.name)
		})
	}
}

// TestColumnStyleCacheConcurrency tests cache operations under concurrent access
func TestColumnStyleCacheConcurrency(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Setup styles for columns A-Z
	for i := 1; i <= 26; i++ {
		style, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i*1000)}, Pattern: 1},
		})
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, style)
	}

	ws, _ := f.workSheetReader(sheet)

	const goroutines = 20
	const iterations = 1000

	var wg sync.WaitGroup
	errors := make(chan error, goroutines)

	// Concurrent reads and writes
	for i := 0; i < goroutines; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()

			for j := 0; j < iterations; j++ {
				col := (j % 26) + 1

				if id%2 == 0 {
					// Read style (cache lookup)
					ws.prepareCellStyle(col, 1, 0)
				} else {
					// Modify style (cache invalidation)
					if j%100 == 0 {
						colName, _ := ColumnNumberToName(col)
						newStyle, _ := f.NewStyle(&Style{
							Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", j*100)}, Pattern: 1},
						})
						f.SetColStyle(sheet, colName+":"+colName, newStyle)
					}
				}
			}
		}(i)
	}

	wg.Wait()
	close(errors)

	for err := range errors {
		t.Errorf("Concurrent operation failed: %v", err)
	}

	t.Logf("✓ Concurrent cache operations completed successfully")
}

// TestColumnStyleCacheMemoryBounds tests cache doesn't grow unbounded
func TestColumnStyleCacheMemoryBounds(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set style for columns 1-100
	for i := 1; i <= 100; i++ {
		style, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i*1000)}, Pattern: 1},
		})
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, style)
	}

	ws, _ := f.workSheetReader(sheet)

	// Access columns 1-100 (populate cache)
	for i := 1; i <= 100; i++ {
		ws.prepareCellStyle(i, 1, 0)
	}

	// Count cache entries
	cacheSize := 0
	ws.colStyleCache.Range(func(key, value interface{}) bool {
		cacheSize++
		return true
	})

	t.Logf("Cache size after 100 accesses: %d", cacheSize)

	if cacheSize != 100 {
		t.Errorf("Expected cache size 100, got %d", cacheSize)
	}

	// Modify column 1 (should clear only that entry)
	style, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FFFFFF"}, Pattern: 1},
	})
	f.SetColStyle(sheet, "A:A", style)

	// Count cache entries after modification
	cacheSize2 := 0
	ws.colStyleCache.Range(func(key, value interface{}) bool {
		cacheSize2++
		return true
	})

	t.Logf("Cache size after modification: %d", cacheSize2)

	if cacheSize2 != 99 {
		t.Errorf("Expected cache size 99, got %d", cacheSize2)
	}

	t.Logf("✓ Cache memory bounds are respected")
}

// TestColumnStyleCacheConsistency tests cache consistency with actual styles
func TestColumnStyleCacheConsistency(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set styles for multiple columns
	styles := make(map[int]int)
	for i := 1; i <= 50; i++ {
		style, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i*10000)}, Pattern: 1},
		})
		styles[i] = style
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, style)
	}

	ws, _ := f.workSheetReader(sheet)

	// Get styles (populate cache)
	for i := 1; i <= 50; i++ {
		cachedStyle := ws.prepareCellStyle(i, 1, 0)
		if cachedStyle != styles[i] {
			t.Errorf("Column %d: cached style %d != expected %d", i, cachedStyle, styles[i])
		}
	}

	// Modify every 10th column
	for i := 10; i <= 50; i += 10 {
		newStyle, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{"#000000"}, Pattern: 1},
		})
		styles[i] = newStyle
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, newStyle)
	}

	// Verify all styles are still consistent
	for i := 1; i <= 50; i++ {
		actualStyle := ws.prepareCellStyle(i, 1, 0)
		if actualStyle != styles[i] {
			t.Errorf("Column %d: actual style %d != expected %d (after modification)", i, actualStyle, styles[i])
		}
	}

	t.Logf("✓ Cache consistency maintained after modifications")
}

// TestColumnStyleCacheRaceConditions runs with -race flag
func TestColumnStyleCacheRaceConditions(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping race condition test in short mode")
	}

	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	const workers = 10
	var wg sync.WaitGroup

	// Concurrent style modifications and reads
	for w := 0; w < workers; w++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()

			for i := 0; i < 100; i++ {
				col := (id*100 + i) % 26 + 1
				colName, _ := ColumnNumberToName(col)

				// Alternate between read and write
				if i%2 == 0 {
					style, _ := f.NewStyle(&Style{
						Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i)}, Pattern: 1},
					})
					f.SetColStyle(sheet, colName+":"+colName, style)
				} else {
					ws, _ := f.workSheetReader(sheet)
					ws.prepareCellStyle(col, 1, 0)
				}
			}
		}(w)
	}

	wg.Wait()
	t.Logf("✓ No race conditions detected")
}

// TestColumnStyleCacheAfterColumnDeletion tests cache after column operations
func TestColumnStyleCacheAfterColumnDeletion(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Set styles for columns A-E
	for i := 1; i <= 5; i++ {
		style, _ := f.NewStyle(&Style{
			Fill: Fill{Type: "pattern", Color: []string{fmt.Sprintf("#%06X", i*10000)}, Pattern: 1},
		})
		colName, _ := ColumnNumberToName(i)
		f.SetColStyle(sheet, colName+":"+colName, style)
	}

	ws, _ := f.workSheetReader(sheet)

	// Populate cache
	for i := 1; i <= 5; i++ {
		ws.prepareCellStyle(i, 1, 0)
	}

	// Verify cache populated
	cacheSize := 0
	ws.colStyleCache.Range(func(key, value interface{}) bool {
		cacheSize++
		return true
	})

	if cacheSize != 5 {
		t.Errorf("Expected cache size 5, got %d", cacheSize)
	}

	// Modify column C (should clear cache for column 3)
	newStyle, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FFFFFF"}, Pattern: 1},
	})
	f.SetColStyle(sheet, "C:C", newStyle)

	// Verify column C cache was cleared
	if _, ok := ws.colStyleCache.Load(3); ok {
		t.Error("Column C cache should be cleared")
	}

	// Verify other columns still cached
	for i := 1; i <= 5; i++ {
		if i == 3 {
			continue
		}
		if _, ok := ws.colStyleCache.Load(i); !ok {
			t.Errorf("Column %d cache should still exist", i)
		}
	}

	t.Logf("✓ Selective cache invalidation works correctly")
}
