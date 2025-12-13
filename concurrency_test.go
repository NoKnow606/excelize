package excelize

import (
	"fmt"
	"sync"
	"testing"
	"time"
)

// TestSetCellValuesConcurrency tests concurrent batch operations
// NOTE: This tests our batch operations, not the underlying File operations
// Excelize File object is NOT designed for concurrent access without external locking
func TestSetCellValuesConcurrency(t *testing.T) {
	// Important: Each goroutine should use its own File instance for thread safety
	const goroutines = 10
	const cellsPerGoroutine = 100

	var wg sync.WaitGroup
	errors := make(chan error, goroutines)

	// Each goroutine uses its own File instance (thread-safe pattern)
	for i := 0; i < goroutines; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()

			// Each goroutine creates its own File
			f := NewFile()
			defer f.Close()

			values := make(map[string]interface{}, cellsPerGoroutine)
			for j := 0; j < cellsPerGoroutine; j++ {
				cell := fmt.Sprintf("A%d", j+1)
				values[cell] = j * id
			}

			if err := f.SetCellValues("Sheet1", values); err != nil {
				errors <- err
			}
		}(i)
	}

	wg.Wait()
	close(errors)

	// Check for errors
	for err := range errors {
		t.Errorf("Concurrent operation failed: %v", err)
	}

	t.Logf("Concurrent batch operations (separate Files) completed successfully")
}

// TestSetCellValuesPanicRecovery tests panic recovery
func TestSetCellValuesPanicRecovery(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// This should NOT panic even with invalid data
	values := map[string]interface{}{
		"A1":      100,
		"INVALID": "bad cell",  // Invalid cell reference
		"A2":      200,
	}

	err := f.SetCellValues("Sheet1", values)
	if err == nil {
		t.Error("Expected error for invalid cell reference")
	}

	t.Logf("Panic recovery test passed, error: %v", err)

	// Verify batch mode was reset
	f.mu.Lock()
	inBatch := f.inBatchMode
	f.mu.Unlock()

	if inBatch {
		t.Error("Batch mode should be reset after error")
	}
}

// TestSetCellValuesBatchModeIsolation tests that batch mode doesn't leak
func TestSetCellValuesBatchModeIsolation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Start batch operation
	values := map[string]interface{}{
		"A1": 100,
		"A2": 200,
	}

	// Batch operation
	err := f.SetCellValues("Sheet1", values)
	if err != nil {
		t.Errorf("SetCellValues failed: %v", err)
	}

	// Verify batch mode is off
	f.mu.Lock()
	inBatch := f.inBatchMode
	f.mu.Unlock()

	if inBatch {
		t.Error("Batch mode should be reset after SetCellValues")
	}

	// Verify normal operations work correctly
	err = f.SetCellValue("Sheet1", "A3", 300)
	if err != nil {
		t.Errorf("SetCellValue after batch failed: %v", err)
	}
}

// TestSetCellValuesNestedBatch tests nested batch operations
func TestSetCellValuesNestedBatch(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Outer batch
	values1 := map[string]interface{}{
		"A1": 100,
	}

	err := f.SetCellValues("Sheet1", values1)
	if err != nil {
		t.Errorf("First batch failed: %v", err)
	}

	// Inner batch (should work independently)
	values2 := map[string]interface{}{
		"B1": 200,
	}

	err = f.SetCellValues("Sheet1", values2)
	if err != nil {
		t.Errorf("Second batch failed: %v", err)
	}

	// Verify both values are set
	val1, _ := f.GetCellValue("Sheet1", "A1")
	val2, _ := f.GetCellValue("Sheet1", "B1")

	if val1 != "100" || val2 != "200" {
		t.Errorf("Values not set correctly: A1=%s, B1=%s", val1, val2)
	}
}

// TestCalcCellValuesConcurrency tests concurrent calculations
// Each goroutine uses its own File instance for thread safety
func TestCalcCellValuesConcurrency(t *testing.T) {
	const goroutines = 10
	var wg sync.WaitGroup
	errors := make(chan error, goroutines)

	// Each goroutine uses its own File instance
	for i := 0; i < goroutines; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()

			f := NewFile()
			defer f.Close()

			// Setup data
			for j := 1; j <= 10; j++ {
				cell := fmt.Sprintf("B%d", j)
				f.SetCellValue("Sheet1", cell, j*10)
			}

			// Set formulas
			for j := 1; j <= 10; j++ {
				cell := fmt.Sprintf("A%d", j)
				refCell := fmt.Sprintf("B%d", j)
				f.SetCellFormula("Sheet1", cell, fmt.Sprintf("=%s*2", refCell))
			}

			// Calculate
			cells := make([]string, 10)
			for j := 0; j < 10; j++ {
				cells[j] = fmt.Sprintf("A%d", j+1)
			}

			results, err := f.CalcCellValues("Sheet1", cells)
			if err != nil {
				errors <- fmt.Errorf("goroutine %d: %w", id, err)
				return
			}

			if len(results) != len(cells) {
				errors <- fmt.Errorf("goroutine %d: expected %d results, got %d", id, len(cells), len(results))
			}
		}(i)
	}

	wg.Wait()
	close(errors)

	for err := range errors {
		t.Errorf("Concurrent calculation failed: %v", err)
	}
}

// TestCalcFormulaValueConcurrency tests concurrent temporary formula calculations
func TestCalcFormulaValueConcurrency(t *testing.T) {
	const goroutines = 10
	var wg sync.WaitGroup
	errors := make(chan error, goroutines)

	// Each goroutine uses its own File instance
	for i := 0; i < goroutines; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()

			f := NewFile()
			defer f.Close()

			// Setup data
			for j := 1; j <= 10; j++ {
				cell := fmt.Sprintf("B%d", j)
				f.SetCellValue("Sheet1", cell, j)
			}

			// Temporary calculations
			for j := 0; j < 10; j++ {
				cell := fmt.Sprintf("A%d", j+1)
				refCell := fmt.Sprintf("B%d", j+1)
				formula := fmt.Sprintf("%s*3", refCell)

				result, err := f.CalcFormulaValue("Sheet1", cell, formula)
				if err != nil {
					errors <- fmt.Errorf("goroutine %d, cell %s: %w", id, cell, err)
					return
				}

				// Verify no formula was persisted
				persistedFormula, _ := f.GetCellFormula("Sheet1", cell)
				if persistedFormula != "" {
					errors <- fmt.Errorf("goroutine %d: formula leaked to %s", id, cell)
					return
				}

				_ = result // Use result to avoid unused variable
			}
		}(i)
	}

	wg.Wait()
	close(errors)

	for err := range errors {
		t.Errorf("Concurrent CalcFormulaValue failed: %v", err)
	}
}

// TestRaceConditions tests for race conditions using stress testing
func TestRaceConditions(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping race condition test in short mode")
	}

	f := NewFile()
	defer f.Close()

	const iterations = 100
	const goroutines = 10

	var wg sync.WaitGroup

	// Mixed operations
	for i := 0; i < goroutines; i++ {
		wg.Add(1)
		go func(id int) {
			defer wg.Done()

			for j := 0; j < iterations; j++ {
				switch j % 4 {
				case 0:
					// Batch set
					values := map[string]interface{}{
						fmt.Sprintf("A%d", id*iterations+j): j,
					}
					f.SetCellValues("Sheet1", values)

				case 1:
					// Single set
					f.SetCellValue("Sheet1", fmt.Sprintf("B%d", id*iterations+j), j)

				case 2:
					// Calculate formula
					f.CalcFormulaValue("Sheet1", fmt.Sprintf("C%d", id*iterations+j), "1+1")

				case 3:
					// Get value
					f.GetCellValue("Sheet1", fmt.Sprintf("A%d", id*iterations+j))
				}
			}
		}(i)
	}

	wg.Wait()
	t.Logf("Stress test completed: %d goroutines × %d iterations", goroutines, iterations)
}

// TestBatchOperationTimeout tests that batch operations don't hang
func TestBatchOperationTimeout(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Large batch operation with timeout
	const size = 10000
	values := make(map[string]interface{}, size)
	for i := 1; i <= size; i++ {
		values[fmt.Sprintf("A%d", i)] = i
	}

	done := make(chan bool)
	go func() {
		err := f.SetCellValues("Sheet1", values)
		if err != nil {
			t.Errorf("Large batch failed: %v", err)
		}
		done <- true
	}()

	select {
	case <-done:
		t.Log("Batch operation completed successfully")
	case <-time.After(10 * time.Second):
		t.Error("Batch operation timed out")
	}
}
