package excelize

import (
	"fmt"
	"testing"
)

func TestLRUCache(t *testing.T) {
	// Test basic LRU cache functionality
	cache := newLRUCache(3)

	// Test Store and Load
	cache.Store("key1", "value1")
	cache.Store("key2", "value2")
	cache.Store("key3", "value3")

	if cache.Len() != 3 {
		t.Errorf("Expected cache length 3, got %d", cache.Len())
	}

	// Test Load
	if val, ok := cache.Load("key1"); !ok || val.(string) != "value1" {
		t.Errorf("Expected to load key1=value1, got %v, %v", val, ok)
	}

	// Test eviction when capacity exceeded
	evicted := cache.Store("key4", "value4")
	if !evicted {
		t.Error("Expected eviction when adding 4th item to cache with capacity 3")
	}

	if cache.Len() != 3 {
		t.Errorf("Expected cache length 3 after eviction, got %d", cache.Len())
	}

	// key2 should be evicted (least recently used)
	if _, ok := cache.Load("key2"); ok {
		t.Error("Expected key2 to be evicted")
	}

	// key1, key3, key4 should still be present
	if _, ok := cache.Load("key1"); !ok {
		t.Error("Expected key1 to still be in cache")
	}
	if _, ok := cache.Load("key3"); !ok {
		t.Error("Expected key3 to still be in cache")
	}
	if _, ok := cache.Load("key4"); !ok {
		t.Error("Expected key4 to be in cache")
	}

	// Test Delete
	if !cache.Delete("key1") {
		t.Error("Expected Delete to return true for existing key")
	}
	if cache.Len() != 2 {
		t.Errorf("Expected cache length 2 after delete, got %d", cache.Len())
	}

	// Test Clear
	cache.Clear()
	if cache.Len() != 0 {
		t.Errorf("Expected cache length 0 after clear, got %d", cache.Len())
	}
}

func TestLRUCacheWithRanges(t *testing.T) {
	// Test LRU cache with realistic range data
	cache := newLRUCache(2)

	// Store two large ranges
	matrix1 := make([][]formulaArg, 100)
	for i := range matrix1 {
		matrix1[i] = make([]formulaArg, 10)
	}
	matrix2 := make([][]formulaArg, 100)
	for i := range matrix2 {
		matrix2[i] = make([]formulaArg, 10)
	}

	cache.Store("Sheet1!A1:J100", matrix1)
	cache.Store("Sheet2!A1:J100", matrix2)

	// Both should be in cache
	if _, ok := cache.Load("Sheet1!A1:J100"); !ok {
		t.Error("Expected Sheet1!A1:J100 to be in cache")
	}
	if _, ok := cache.Load("Sheet2!A1:J100"); !ok {
		t.Error("Expected Sheet2!A1:J100 to be in cache")
	}

	// Add third range - should evict Sheet1 (least recently used)
	matrix3 := make([][]formulaArg, 100)
	for i := range matrix3 {
		matrix3[i] = make([]formulaArg, 10)
	}
	evicted := cache.Store("Sheet3!A1:J100", matrix3)

	if !evicted {
		t.Error("Expected eviction when adding 3rd range to cache with capacity 2")
	}

	// Sheet1 should be evicted
	if _, ok := cache.Load("Sheet1!A1:J100"); ok {
		t.Error("Expected Sheet1!A1:J100 to be evicted")
	}

	// Sheet2 and Sheet3 should still be present
	if _, ok := cache.Load("Sheet2!A1:J100"); !ok {
		t.Error("Expected Sheet2!A1:J100 to still be in cache")
	}
	if _, ok := cache.Load("Sheet3!A1:J100"); !ok {
		t.Error("Expected Sheet3!A1:J100 to be in cache")
	}

	fmt.Printf("LRU cache test passed: capacity=%d, entries=%d\n", 2, cache.Len())
}

func TestLRUCacheRange(t *testing.T) {
	cache := newLRUCache(5)

	cache.Store("key1", "value1")
	cache.Store("key2", "value2")
	cache.Store("key3", "value3")

	count := 0
	cache.Range(func(key string, value interface{}) bool {
		count++
		return true
	})

	if count != 3 {
		t.Errorf("Expected to iterate 3 items, got %d", count)
	}

	// Test early termination
	count = 0
	cache.Range(func(key string, value interface{}) bool {
		count++
		return count < 2 // Stop after 2 items
	})

	if count != 2 {
		t.Errorf("Expected to iterate 2 items before stopping, got %d", count)
	}
}
