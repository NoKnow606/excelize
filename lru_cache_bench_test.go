package excelize

import (
	"fmt"
	"runtime"
	"testing"
)

func BenchmarkLRUCacheMemory(b *testing.B) {
	// Test memory usage with LRU cache limiting to 50 entries
	cache := newLRUCache(50)

	// Force GC and get baseline
	runtime.GC()
	var m1 runtime.MemStats
	runtime.ReadMemStats(&m1)

	// Store 200 large matrices (4x the capacity)
	for i := 0; i < 200; i++ {
		key := fmt.Sprintf("Sheet%d!A1:Z10000", i)
		matrix := make([][]formulaArg, 10000)
		for j := range matrix {
			matrix[j] = make([]formulaArg, 26)
			for k := range matrix[j] {
				matrix[j][k] = formulaArg{
					Type:   ArgNumber,
					Number: float64(i*10000 + j*26 + k),
				}
			}
		}
		cache.Store(key, matrix)
	}

	// Force GC and measure final memory
	runtime.GC()
	var m2 runtime.MemStats
	runtime.ReadMemStats(&m2)

	cacheLen := cache.Len()
	allocMB := float64(m2.Alloc-m1.Alloc) / 1024 / 1024

	b.Logf("LRU Cache Results:")
	b.Logf("  Stored: 200 matrices")
	b.Logf("  Capacity: 50 matrices")
	b.Logf("  Actual entries: %d", cacheLen)
	b.Logf("  Memory allocated: %.2f MB", allocMB)
	b.Logf("  Memory per entry: %.2f MB", allocMB/float64(cacheLen))

	if cacheLen != 50 {
		b.Errorf("Expected cache to hold 50 entries, got %d", cacheLen)
	}

	// Verify only the most recent 50 are kept
	oldestKept := 200 - 50 // Should be 150
	for i := 0; i < oldestKept; i++ {
		key := fmt.Sprintf("Sheet%d!A1:Z10000", i)
		if _, ok := cache.Load(key); ok {
			b.Errorf("Expected Sheet%d to be evicted", i)
		}
	}

	for i := oldestKept; i < 200; i++ {
		key := fmt.Sprintf("Sheet%d!A1:Z10000", i)
		if _, ok := cache.Load(key); !ok {
			b.Errorf("Expected Sheet%d to be in cache", i)
		}
	}
}

func TestLRUCacheMemoryLimit(t *testing.T) {
	// Compare memory usage with and without LRU
	t.Run("WithLRU_50", func(t *testing.T) {
		cache := newLRUCache(50)
		runtime.GC()
		var m1 runtime.MemStats
		runtime.ReadMemStats(&m1)

		// Store 100 matrices
		for i := 0; i < 100; i++ {
			key := fmt.Sprintf("Sheet%d!A1:J1000", i)
			matrix := make([][]formulaArg, 1000)
			for j := range matrix {
				matrix[j] = make([]formulaArg, 10)
			}
			cache.Store(key, matrix)
		}

		runtime.GC()
		var m2 runtime.MemStats
		runtime.ReadMemStats(&m2)

		allocMB := float64(m2.Alloc-m1.Alloc) / 1024 / 1024
		t.Logf("LRU(50) - Stored 100, Kept %d, Memory: %.2f MB", cache.Len(), allocMB)

		if cache.Len() != 50 {
			t.Errorf("Expected 50 entries, got %d", cache.Len())
		}
	})

	t.Run("WithLRU_10", func(t *testing.T) {
		cache := newLRUCache(10)
		runtime.GC()
		var m1 runtime.MemStats
		runtime.ReadMemStats(&m1)

		// Store 100 matrices
		for i := 0; i < 100; i++ {
			key := fmt.Sprintf("Sheet%d!A1:J1000", i)
			matrix := make([][]formulaArg, 1000)
			for j := range matrix {
				matrix[j] = make([]formulaArg, 10)
			}
			cache.Store(key, matrix)
		}

		runtime.GC()
		var m2 runtime.MemStats
		runtime.ReadMemStats(&m2)

		allocMB := float64(m2.Alloc-m1.Alloc) / 1024 / 1024
		t.Logf("LRU(10) - Stored 100, Kept %d, Memory: %.2f MB", cache.Len(), allocMB)

		if cache.Len() != 10 {
			t.Errorf("Expected 10 entries, got %d", cache.Len())
		}
	})
}
