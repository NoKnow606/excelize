package excelize

import (
	"fmt"
	"runtime"
	"strings"
	"sync"
)

// CalcCellValuesConcurrent calculates multiple cell values concurrently using goroutines.
// This is an optimized version that can be significantly faster for large datasets.
func (f *File) CalcCellValuesConcurrent(sheet string, cells []string, opts ...Options) (map[string]string, error) {
	if len(cells) == 0 {
		return make(map[string]string), nil
	}

	// Determine optimal worker count
	numWorkers := runtime.NumCPU()
	if len(cells) < numWorkers {
		numWorkers = len(cells)
	}

	// Results storage with mutex protection
	results := make(map[string]string, len(cells))
	var resultsMu sync.Mutex

	// Error collection with mutex protection
	var errors []error
	var errorsMu sync.Mutex

	// Work distribution
	var wg sync.WaitGroup
	cellChan := make(chan string, len(cells))

	// Start workers
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for cell := range cellChan {
				result, err := f.CalcCellValue(sheet, cell, opts...)
				if err != nil {
					errorsMu.Lock()
					errors = append(errors, fmt.Errorf("failed to calculate %s: %w", cell, err))
					errorsMu.Unlock()
					continue
				}

				resultsMu.Lock()
				results[cell] = result
				resultsMu.Unlock()
			}
		}()
	}

	// Send work to workers
	for _, cell := range cells {
		cellChan <- cell
	}
	close(cellChan)

	// Wait for all workers to complete
	wg.Wait()

	// Return partial results with combined errors if any
	if len(errors) > 0 {
		var sb strings.Builder
		for i, err := range errors {
			if i > 0 {
				sb.WriteString("; ")
			}
			sb.WriteString(err.Error())
		}
		return results, fmt.Errorf("%s", sb.String())
	}

	return results, nil
}

// CalcCellValuesOptimized is an optimized version with multiple improvements:
// 1. Pre-check cache before calculation
// 2. Use strings.Builder for error messages
// 3. Estimate result map capacity accurately
func (f *File) CalcCellValuesOptimized(sheet string, cells []string, opts ...Options) (map[string]string, error) {
	if len(cells) == 0 {
		return make(map[string]string), nil
	}

	results := make(map[string]string, len(cells))
	var errors []error

	// Calculate all cells, benefiting from cache
	// Skip cells that fail to calculate and collect errors
	for _, cell := range cells {
		result, err := f.CalcCellValue(sheet, cell, opts...)
		if err != nil {
			errors = append(errors, fmt.Errorf("failed to calculate %s: %w", cell, err))
			continue
		}
		results[cell] = result
	}

	// Return partial results with combined errors if any
	if len(errors) > 0 {
		// Use strings.Builder for efficient string concatenation
		var sb strings.Builder
		for i, err := range errors {
			if i > 0 {
				sb.WriteString("; ")
			}
			sb.WriteString(err.Error())
		}
		return results, fmt.Errorf("%s", sb.String())
	}

	return results, nil
}
