package excelize

import (
	"fmt"
	"runtime"
	"strings"
	"sync"
	"time"
)

// CalcCellValuesDependencyAware 依赖感知的并发计算
// 策略: 先串行计算一小批公式来预热缓存,然后并发计算剩余公式
func (f *File) CalcCellValuesDependencyAware(sheet string, cells []string, opts ...Options) (map[string]string, error) {
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

	start := time.Now()
	results := make(map[string]string, len(cells))

	// 阶段1: 预热缓存 - 只计算每列的第一个单元格
	warmupCells := make(map[string]string) // col -> first cell
	for _, cell := range cells {
		col, _, _ := CellNameToCoordinates(cell)
		colName, _ := ColumnNumberToName(col)
		if _, exists := warmupCells[colName]; !exists {
			warmupCells[colName] = cell
		}
	}

	warmupList := make([]string, 0, len(warmupCells))
	for _, cell := range warmupCells {
		warmupList = append(warmupList, cell)
	}

	fmt.Printf("Phase 1: Warming up cache with %d cells (one per column)...\n", len(warmupList))
	warmupStart := time.Now()
	for _, cell := range warmupList {
		fmt.Printf("  Warming up: %s\n", cell)
		t1 := time.Now()
		result, err := f.CalcCellValue(sheet, cell, opts...)
		fmt.Printf("  %s took %v, result: %s\n", cell, time.Since(t1), result)
		if err != nil {
			return results, fmt.Errorf("warmup failed at %s: %w", cell, err)
		}
		results[cell] = result
	}
	warmupElapsed := time.Since(warmupStart)
	fmt.Printf("Phase 1 completed: %d cells in %v (%.1f cells/sec)\n",
		len(warmupList), warmupElapsed, float64(len(warmupList))/warmupElapsed.Seconds())

	// 检查 rangeCache
	cacheCount := 0
	f.rangeCache.Range(func(key, value interface{}) bool {
		cacheCount++
		return true
	})
	fmt.Printf("  rangeCache entries: %d\n", cacheCount)

	// 阶段2: 并发计算剩余公式
	remaining := make([]string, 0, len(cells)-len(warmupList))
	warmupSet := make(map[string]bool)
	for _, cell := range warmupList {
		warmupSet[cell] = true
	}
	for _, cell := range cells {
		if !warmupSet[cell] {
			remaining = append(remaining, cell)
		}
	}

	if len(remaining) == 0 {
		return results, nil
	}

	fmt.Printf("Phase 2: Concurrent calculation of %d cells...\n", len(remaining))
	concurrentStart := time.Now()

	numWorkers := runtime.NumCPU()
	chunkSize := (len(remaining) + numWorkers - 1) / numWorkers

	type workerResult struct {
		results map[string]string
		errors  []error
		elapsed time.Duration
	}

	resultChan := make(chan workerResult, numWorkers)
	var wg sync.WaitGroup

	for i := 0; i < numWorkers; i++ {
		start := i * chunkSize
		end := start + chunkSize
		if end > len(remaining) {
			end = len(remaining)
		}
		if start >= len(remaining) {
			break
		}

		wg.Add(1)
		go func(cellChunk []string, workerID int) {
			defer wg.Done()
			localResults := make(map[string]string)
			var localErrors []error

			workerStart := time.Now()
			for idx, cell := range cellChunk {
				t1 := time.Now()
				result, err := f.CalcCellValue(sheet, cell, opts...)
				elapsed := time.Since(t1)

				// 打印前10个和慢的公式
				if idx < 10 || elapsed > 10*time.Millisecond {
					fmt.Printf("Worker %d: cell %s took %v\n", workerID, cell, elapsed)
				}

				if err != nil {
					localErrors = append(localErrors, fmt.Errorf("failed to calculate %s: %w", cell, err))
					continue
				}
				localResults[cell] = result
			}
			workerElapsed := time.Since(workerStart)

			resultChan <- workerResult{
				results: localResults,
				errors:  localErrors,
				elapsed: workerElapsed,
			}
		}(remaining[start:end], i)
	}

	go func() {
		wg.Wait()
		close(resultChan)
	}()

	var errors []error
	var maxTime, minTime time.Duration
	minTime = time.Hour

	for wr := range resultChan {
		for k, v := range wr.results {
			results[k] = v
		}
		errors = append(errors, wr.errors...)
		if wr.elapsed > maxTime {
			maxTime = wr.elapsed
		}
		if wr.elapsed < minTime {
			minTime = wr.elapsed
		}
	}

	concurrentElapsed := time.Since(concurrentStart)
	fmt.Printf("Phase 2 completed: %d cells in %v (%.1f cells/sec)\n",
		len(remaining), concurrentElapsed, float64(len(remaining))/concurrentElapsed.Seconds())
	fmt.Printf("  Worker times: min=%v, max=%v, diff=%v (%.1f%% efficiency)\n",
		minTime, maxTime, maxTime-minTime, float64(minTime)/float64(maxTime)*100)

	totalElapsed := time.Since(start)
	fmt.Printf("\nTotal: %d cells in %v (%.1f cells/sec)\n",
		len(cells), totalElapsed, float64(len(cells))/totalElapsed.Seconds())

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
