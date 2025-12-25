package excelize

import (
	"fmt"
	"runtime"
	"sort"
	"strings"
	"sync"
	"time"
)

// CalcCellValuesDependencyAwareOptions provides options for dependency-aware calculation
type CalcCellValuesDependencyAwareOptions struct {
	// EnableDebug enables debug output for performance monitoring
	EnableDebug bool
	// Options for cell value calculation
	CalcOptions Options
}

// CalcCellValuesDependencyAware calculates multiple cell values with dependency-aware optimization.
// It uses a two-phase strategy:
// 1. Phase 1 (Warmup): Sequentially calculates a small batch of formulas to warm up caches
//    (especially useful for VLOOKUP/XLOOKUP that share lookup tables)
// 2. Phase 2 (Concurrent): Calculates remaining formulas using multiple goroutines
//
// This is most effective when:
//   - You have many formulas (> 1000 cells)
//   - Formulas share dependencies (e.g., multiple VLOOKUP referencing same table)
//   - You need maximum performance
//
// For simpler cases or when formulas are independent, consider using CalcCellValuesConcurrent instead.
func (f *File) CalcCellValuesDependencyAware(sheet string, cells []string, options CalcCellValuesDependencyAwareOptions) (map[string]string, error) {
	if len(cells) == 0 {
		return make(map[string]string), nil
	}

	f.mu.Lock()
	if !f.formulaChecked {
		if err := f.setArrayFormulaCells(); err != nil {
			f.mu.Unlock()
			return nil, err
		}
		f.formulaChecked = true
	}
	f.mu.Unlock()

	results := make(map[string]string, len(cells))
	var allErrors []error

	// Phase 1: Warmup - calculate first cell of each column to warm up caches
	// This is especially helpful for lookup formulas (VLOOKUP, XLOOKUP, etc.)
	warmupCells := make(map[string]string) // col -> first cell
	for _, cell := range cells {
		col, _, err := CellNameToCoordinates(cell)
		if err != nil {
			continue
		}
		colName, _ := ColumnNumberToName(col)
		if _, exists := warmupCells[colName]; !exists {
			warmupCells[colName] = cell
		}
	}

	// Sort warmup cells for deterministic order
	warmupList := make([]string, 0, len(warmupCells))
	for _, cell := range warmupCells {
		warmupList = append(warmupList, cell)
	}
	sort.Strings(warmupList)

	var warmupStart time.Time
	if options.EnableDebug {
		warmupStart = time.Now()
		fmt.Printf("Phase 1: Warming up cache with %d cells (one per column)...\n", len(warmupList))
	}

	// Pre-warm worksheet cache to avoid repeated namespace conversion
	// This significantly reduces memory allocation during formula calculation
	_, _ = f.workSheetReader(sheet)

	// Pre-warm commonly referenced sheets (e.g., lookup tables)
	// This is a heuristic: warm up first 3 sheets which often contain lookup data
	sheetList := f.GetSheetList()
	for i := 0; i < len(sheetList) && i < 3; i++ {
		if sheetList[i] != sheet {
			_, _ = f.workSheetReader(sheetList[i])
		}
	}

	// Calculate warmup cells sequentially
	for _, cell := range warmupList {
		result, err := f.CalcCellValue(sheet, cell, options.CalcOptions)
		if err != nil {
			// Collect error but continue (consistent with phase 2)
			allErrors = append(allErrors, fmt.Errorf("failed to calculate %s: %w", cell, err))
			if options.EnableDebug {
				fmt.Printf("  Warning: %s failed: %v\n", cell, err)
			}
			continue
		}
		results[cell] = result
	}

	if options.EnableDebug {
		warmupElapsed := time.Since(warmupStart)
		cacheCount := f.rangeCache.Len()
		fmt.Printf("Phase 1 completed: %d cells calculated, rangeCache entries: %d, elapsed: %v\n",
			len(warmupList), cacheCount, warmupElapsed)
	}

	// Phase 2: Concurrent calculation of remaining formulas
	remaining := make([]string, 0, len(cells)-len(warmupList))
	warmupSet := make(map[string]bool, len(warmupList))
	for _, cell := range warmupList {
		warmupSet[cell] = true
	}
	for _, cell := range cells {
		if !warmupSet[cell] {
			remaining = append(remaining, cell)
		}
	}

	if len(remaining) == 0 {
		// All cells were in warmup phase
		if len(allErrors) > 0 {
			return results, combineErrors(allErrors)
		}
		return results, nil
	}

	var concurrentStart time.Time
	if options.EnableDebug {
		concurrentStart = time.Now()
		fmt.Printf("Phase 2: Concurrent calculation of %d cells...\n", len(remaining))
	}

	// Calculate optimal number of workers and chunk size
	numWorkers := runtime.NumCPU()
	if len(remaining) < numWorkers {
		numWorkers = len(remaining)
	}
	chunkSize := (len(remaining) + numWorkers - 1) / numWorkers

	type workerResult struct {
		results map[string]string
		errors  []error
	}

	resultChan := make(chan workerResult, numWorkers)
	var wg sync.WaitGroup

	// Launch workers
	for i := 0; i < numWorkers; i++ {
		chunkStart := i * chunkSize
		chunkEnd := chunkStart + chunkSize
		if chunkEnd > len(remaining) {
			chunkEnd = len(remaining)
		}
		if chunkStart >= len(remaining) {
			break
		}

		wg.Add(1)
		go func(cellChunk []string, workerID int) {
			defer wg.Done()
			localResults := make(map[string]string, len(cellChunk))
			var localErrors []error

			for _, cell := range cellChunk {
				result, err := f.CalcCellValue(sheet, cell, options.CalcOptions)
				if err != nil {
					localErrors = append(localErrors, fmt.Errorf("failed to calculate %s: %w", cell, err))
					continue
				}
				localResults[cell] = result
			}

			resultChan <- workerResult{
				results: localResults,
				errors:  localErrors,
			}
		}(remaining[chunkStart:chunkEnd], i)
	}

	// Wait for all workers and close channel
	go func() {
		wg.Wait()
		close(resultChan)
	}()

	// Collect results from all workers
	for wr := range resultChan {
		for k, v := range wr.results {
			results[k] = v
		}
		allErrors = append(allErrors, wr.errors...)
	}

	if options.EnableDebug {
		concurrentElapsed := time.Since(concurrentStart)
		fmt.Printf("Phase 2 completed: %d cells calculated, elapsed: %v\n", len(remaining), concurrentElapsed)
		fmt.Printf("Total: %d cells, %d successful, %d failed\n",
			len(cells), len(results), len(allErrors))
	}

	// Clear range cache after batch calculation to free memory
	// rangeCache can hold large matrices that are no longer needed after batch completion
	// This prevents memory exhaustion in long-running processes or multiple batch calculations
	f.rangeCache.Clear()

	// Return partial results with combined errors if any
	if len(allErrors) > 0 {
		return results, combineErrors(allErrors)
	}

	return results, nil
}

// combineErrors combines multiple errors into a single error message
func combineErrors(errors []error) error {
	if len(errors) == 0 {
		return nil
	}
	var sb strings.Builder
	for i, err := range errors {
		if i > 0 {
			sb.WriteString("; ")
		}
		sb.WriteString(err.Error())
	}
	return fmt.Errorf("%s", sb.String())
}
