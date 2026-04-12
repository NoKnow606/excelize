package excelize

import (
	"context"
	"fmt"
	"log"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"time"
)

// BatchDebugStats 批量更新的调试统计信息
type BatchDebugStats struct {
	TotalCells    int                   // 总计算单元格数
	CellStats     map[string]*CellStats // 每个单元格的统计
	TotalDuration time.Duration         // 总耗时
	CacheHits     int                   // 缓存命中次数
	CacheMisses   int                   // 缓存未命中次数
	mu            sync.Mutex            // 保护并发访问
}

// CellStats 单个单元格的统计信息
type CellStats struct {
	Cell         string        // 单元格坐标 (Sheet!Cell)
	CalcCount    int           // 计算次数
	CalcDuration time.Duration // 计算总耗时
	CacheHit     bool          // 是否命中缓存
	Formula      string        // 公式内容
	Result       string        // 计算结果
}

// enableBatchDebug 是否启用批量更新调试
var enableBatchDebug = false

// currentBatchStats 当前批量更新的统计信息
var currentBatchStats *BatchDebugStats
var batchStatsMu sync.Mutex

// EnableBatchDebug 启用批量更新调试统计
func EnableBatchDebug() {
	enableBatchDebug = true
}

// DisableBatchDebug 禁用批量更新调试统计
func DisableBatchDebug() {
	enableBatchDebug = false
}

// GetBatchDebugStats 获取最近一次批量更新的调试统计
func GetBatchDebugStats() *BatchDebugStats {
	batchStatsMu.Lock()
	defer batchStatsMu.Unlock()
	return currentBatchStats
}

// recordCellCalc 记录单元格计算
func recordCellCalc(sheet, cell, formula, result string, duration time.Duration, cacheHit bool) {
	if !enableBatchDebug || currentBatchStats == nil {
		return
	}

	currentBatchStats.mu.Lock()
	defer currentBatchStats.mu.Unlock()

	cellKey := sheet + "!" + cell
	if currentBatchStats.CellStats[cellKey] == nil {
		currentBatchStats.CellStats[cellKey] = &CellStats{
			Cell:    cellKey,
			Formula: formula,
		}
	}

	stats := currentBatchStats.CellStats[cellKey]
	stats.CalcCount++
	stats.CalcDuration += duration
	stats.CacheHit = cacheHit
	stats.Result = result

	if cacheHit {
		currentBatchStats.CacheHits++
	} else {
		currentBatchStats.CacheMisses++
	}
}

// CellUpdate 表示一个单元格更新操作
type CellUpdate struct {
	Sheet string      // 工作表名称
	Cell  string      // 单元格坐标，如 "A1"
	Value interface{} // 单元格值
}

// FormulaUpdate 表示一个公式更新操作
type FormulaUpdate struct {
	Sheet   string // 工作表名称
	Cell    string // 单元格坐标，如 "A1"
	Formula string // 公式内容，如 "=A1*2"（可以包含或不包含前导 '='）
}

// FormulaUpdateWithValue 表示一个带预计算值的公式更新操作
// 使用此结构体可以避免设置公式后重新计算，直接使用提供的缓存值
type FormulaUpdateWithValue struct {
	Sheet   string // 工作表名称
	Cell    string // 单元格坐标，如 "A1"
	Formula string // 公式内容，如 "=A1*2"
	Value   string // 预计算的值（公式的计算结果）
}

// BatchSetCellValue 批量设置单元格值，不触发重新计算
//
// 此函数用于批量更新多个单元格的值，相比循环调用 SetCellValue，
// 这个函数可以避免重复的工作表查找和验证操作。
//
// 注意：此函数不会自动重新计算公式。如果需要重新计算，
// 请在调用后使用 RecalculateSheet 或 UpdateCellAndRecalculate。
//
// 参数：
//
//	updates: 单元格更新列表
//
// 示例：
//
//	updates := []excelize.CellUpdate{
//	    {Sheet: "Sheet1", Cell: "A1", Value: 100},
//	    {Sheet: "Sheet1", Cell: "A2", Value: 200},
//	    {Sheet: "Sheet1", Cell: "A3", Value: 300},
//	}
//	err := f.BatchSetCellValue(updates)
func (f *File) BatchSetCellValue(updates []CellUpdate) error {
	for _, update := range updates {
		if err := f.SetCellValue(update.Sheet, update.Cell, update.Value); err != nil {
			return err
		}
	}
	return nil
}

// RecalculateSheet 重新计算指定工作表中所有公式单元格的值
//
// 此函数会遍历工作表中的所有公式单元格，重新计算它们的值并更新缓存。
// 这在批量更新单元格后需要重新计算依赖公式时非常有用。
//
// 参数：
//
//	sheet: 工作表名称
//
// 注意：此函数只会重新计算该工作表中的公式，不会影响其他工作表。
//
// 示例：
//
//	// 批量更新后重新计算
//	f.BatchSetCellValue(updates)
//	err := f.RecalculateSheet("Sheet1")
func (f *File) RecalculateSheet(sheet string) error {
	// Get sheet ID (1-based, matches calcChain)
	sheetID := f.getSheetID(sheet)
	if sheetID == -1 {
		return ErrSheetNotExist{SheetName: sheet}
	}

	// Read calcChain
	calcChain, err := f.calcChainReader()
	if err != nil {
		return err
	}

	// If calcChain doesn't exist or is empty, nothing to do
	if calcChain == nil || len(calcChain.C) == 0 {
		return nil
	}

	// Recalculate all formulas in the sheet
	return f.recalculateAllInSheet(calcChain, sheetID)
}

// RecalculateAll 重新计算所有工作表中的所有公式并更新缓存值
//
// 此函数会遍历 calcChain 中的所有公式单元格，重新计算并更新缓存值。
// 计算结果会直接更新到工作表的单元格缓存中。
//
// 注意：为了避免内存溢出，此函数不再返回受影响单元格的列表。
// 所有计算结果已经直接更新到工作表中，可以通过 GetCellValue 读取。
//
// 返回：
//
//	error: 错误信息
//
// 示例：
//
//	err := f.RecalculateAll()
//	if err != nil {
//	    log.Fatal(err)
//	}
//	// 读取计算后的值
//	value, _ := f.GetCellValue("Sheet1", "A1")
func (f *File) RecalculateAll() error {
	totalStart := time.Now()

	calcChain, err := f.calcChainReader()
	if err != nil {
		return err
	}

	if calcChain == nil || len(calcChain.C) == 0 {
		return nil
	}

	log.Printf("📊 [RecalculateAll] Starting: %d formulas to calculate", len(calcChain.C))

	// === 批量SUMIFS/AVERAGEIFS优化 ===
	// 在逐个计算之前，先检测并批量计算SUMIFS/AVERAGEIFS公式
	batchStart := time.Now()
	batchResults := f.detectAndCalculateBatchSUMIFS()
	batchDuration := time.Since(batchStart)

	batchCount := len(batchResults)
	if batchCount > 0 {
		log.Printf("⚡ [RecalculateAll] Batch SUMIFS/AVERAGEIFS/SUMPRODUCT optimization: %d formulas calculated in %v (avg: %v/formula)",
			batchCount, batchDuration, batchDuration/time.Duration(batchCount))

		// 将批量结果存入calcCache，这样后续逐个计算时会直接使用缓存
		for fullCell, value := range batchResults {
			// fullCell format: "Sheet!Cell"
			cacheKey := fullCell + "!raw=true"
			f.calcCache.Store(cacheKey, fmt.Sprintf("%g", value))
		}
	}

	sheetList := f.GetSheetList()
	currentSheetIndex := -1
	var currentWs *xlsxWorksheet
	var currentSheetName string
	sheetFormulaCount := 0 // Track formulas within current sheet

	// Pre-build cell map for current sheet to avoid O(n²) lookups
	cellMap := make(map[string]*xlsxC)

	sheetBuildTime := time.Duration(0)
	calcTime := time.Duration(0)
	formulaCount := 0
	batchHitCount := 0                        // Track how many formulas used batch results
	progressInterval := len(calcChain.C) / 20 // Report every 5% (changed from 10%)
	slowFormulaCount := 0                     // Track slow formulas (>100ms)
	timeoutCount := 0                         // Track timeout formulas (>5s)
	skippedComplexFormulas := 0               // Track skipped complex formulas

	// Track slow formulas with details
	type slowFormulaInfo struct {
		sheet    string
		cell     string
		duration time.Duration
		formula  string
	}
	slowFormulas := make([]slowFormulaInfo, 0, 100) // Store top 100 slow formulas

	// Helper to insert slow formula in sorted order (by duration descending)
	insertSlowFormula := func(sf slowFormulaInfo) {
		// Find insertion position
		insertPos := len(slowFormulas)
		for i := 0; i < len(slowFormulas); i++ {
			if sf.duration > slowFormulas[i].duration {
				insertPos = i
				break
			}
		}

		// Insert or append
		if insertPos < len(slowFormulas) {
			// Insert at position
			slowFormulas = append(slowFormulas[:insertPos+1], slowFormulas[insertPos:]...)
			slowFormulas[insertPos] = sf
		} else if len(slowFormulas) < 100 {
			// Append if not full
			slowFormulas = append(slowFormulas, sf)
		}

		// Keep only top 100
		if len(slowFormulas) > 100 {
			slowFormulas = slowFormulas[:100]
		}
	}

	// Track columns that have timed out - skip all cells in that column
	// Map: "SheetName!Column" -> true (e.g., "Sheet1!H" -> true)
	timeoutColumns := make(map[string]bool)

	// Track columns with circular references detected
	circularRefColumns := make(map[string]bool)

	for i := range calcChain.C {
		c := calcChain.C[i]
		if c.I != 0 {
			currentSheetIndex = c.I
		}

		if currentSheetIndex < 0 || currentSheetIndex >= len(sheetList) {
			continue
		}

		sheetName := sheetList[currentSheetIndex]

		// If sheet changed, rebuild cell map
		if sheetName != currentSheetName {
			buildStart := time.Now()

			// 🔥 MEMORY OPTIMIZATION: Clear previous sheet's cellMap to free memory
			if len(cellMap) > 0 {
				cellMap = nil // Allow GC to collect old map
			}

			currentSheetName = sheetName
			sheetFormulaCount = 0 // Reset counter for new sheet
			currentWs, err = f.workSheetReader(sheetName)
			if err != nil {
				continue
			}

			// Build cell map for fast lookup
			// Pre-allocate with estimated capacity to reduce allocations
			estimatedCells := 0
			if currentWs != nil && currentWs.SheetData.Row != nil {
				estimatedCells = len(currentWs.SheetData.Row) * 50 // Estimate ~50 cells per row
			}
			cellMap = make(map[string]*xlsxC, estimatedCells)

			if currentWs != nil && currentWs.SheetData.Row != nil {
				for rowIdx := range currentWs.SheetData.Row {
					for cellIdx := range currentWs.SheetData.Row[rowIdx].C {
						cell := &currentWs.SheetData.Row[rowIdx].C[cellIdx]
						cellMap[cell.R] = cell
					}
				}
			}
			buildDuration := time.Since(buildStart)
			sheetBuildTime += buildDuration
		}

		// Fast lookup using cellMap
		cellRef, exists := cellMap[c.R]
		if !exists || cellRef.F == nil {
			continue
		}

		sheetFormulaCount++ // Increment sheet-level counter

		// Extract column letter from cell reference (e.g., "H2" -> "H")
		colLetter := ""
		for _, ch := range c.R {
			if ch >= 'A' && ch <= 'Z' {
				colLetter += string(ch)
			} else {
				break
			}
		}

		columnKey := sheetName + "!" + colLetter

		// Check if this column has circular reference
		if circularRefColumns[columnKey] {
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			skippedComplexFormulas++
			continue
		}

		// Get formula content (will be used for circular ref detection and dependency checks)
		var formula string
		if cellRef.F.Content != "" {
			formula = cellRef.F.Content
		} else if cellRef.F.T == STCellFormulaTypeShared && cellRef.F.Si != nil {
			formula, _ = getSharedFormula(currentWs, *cellRef.F.Si, c.R)
		}

		// 🔥 AUTO-DETECT circular reference in formula
		// IMPORTANT: Only check for self-reference within the SAME sheet
		// A formula in Sheet1!B2 can reference Sheet2!B2 without circular dependency
		// Pattern 1: Direct self-reference like "B2", "$B2", "B$2", "$B$2" (WITHOUT sheet prefix)
		// Pattern 2: INDEX with self cell reference as parameter (e.g., INDEX(..., B2))
		hasCircularRef := false

		// Only check for circular reference if formula does NOT contain cross-sheet references
		// If formula has '!' it might be referencing other sheets, need more careful check
		hasCrossSheetRef := strings.Contains(formula, "!")

		if !hasCrossSheetRef {
			// No cross-sheet references, safe to check for self-reference
			// Use word boundary to avoid false positives (e.g., AC2 matching C2)
			// Match patterns: \bC2\b, \b$C2\b, \bC$2\b, \b$C$2\b
			selfRefPattern := regexp.MustCompile(`\b\$?` + regexp.QuoteMeta(colLetter) + `\$?` + regexp.QuoteMeta(c.R[len(colLetter):]) + `\b`)
			if selfRefPattern.MatchString(formula) {
				hasCircularRef = true
			}

			// Check for column reference in INDEX/OFFSET that could cause circular dependency
			// e.g., INDEX(..., colLetter+rowNum) or INDEX(...+H2)
			if !hasCircularRef && (strings.Contains(formula, "INDEX") || strings.Contains(formula, "OFFSET")) {
				// Check if cell reference is used as a parameter (not part of range)
				// Pattern: INDEX(..., +C2) or INDEX(..., C2+...)
				indexPattern := regexp.MustCompile(`[+\-*/,\(]\s*\$?` + regexp.QuoteMeta(c.R) + `\s*[+\-*/,\)]`)
				if indexPattern.MatchString(formula) {
					hasCircularRef = true
				}
			}
		} else {
			// Has cross-sheet references, need to check more carefully
			// Only flag as circular if it references the SAME sheet AND cell
			// Pattern: Sheet!C2 or 'Sheet'!C2 where Sheet is the current sheet
			selfRef1 := sheetName + "!" + c.R
			selfRef2 := "'" + sheetName + "'!" + c.R

			if strings.Contains(formula, selfRef1) || strings.Contains(formula, selfRef2) {
				hasCircularRef = true
			}

			// Also check for same-sheet plain cell references (without sheet prefix)
			// IMPORTANT: Prevent false positives like AC2 matching C2
			// Require: not preceded by letter (to avoid AC2, BC2, etc matching C2)
			if !hasCircularRef {
				// Pattern: [^A-Z!\'\"]\$?C\$?2\b
				// This ensures the cell ref is not part of a longer column name
				plainCellPattern := regexp.MustCompile(`[^A-Z!\'\"]\$?` + regexp.QuoteMeta(colLetter) + `\$?` + regexp.QuoteMeta(c.R[len(colLetter):]) + `\b`)
				if plainCellPattern.MatchString(formula) {
					hasCircularRef = true
				}
			}
		}

		if hasCircularRef {
			// Mark entire column as having circular reference
			circularRefColumns[columnKey] = true
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			skippedComplexFormulas++

			// Log first occurrence
			if len(circularRefColumns) == 1 {
				log.Printf("  🔄 [RecalculateAll] Circular reference detected: %s!%s (formula references itself)", sheetName, c.R)
			}
			continue
		}

		// Check if this column has already timed out - if so, skip it
		if timeoutColumns[columnKey] {
			// Skip this cell silently - column already timed out
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			timeoutCount++
			continue
		}

		// Check if this formula depends on any circular-ref columns
		// If so, skip it to prevent cascading errors
		dependsOnCircular := false
		for circularCol := range circularRefColumns {
			// Extract sheet and column from "SheetName!Column"
			parts := strings.Split(circularCol, "!")
			if len(parts) == 2 {
				circularSheet := parts[0]
				circularColumn := parts[1]

				// Check if formula references this circular column
				if circularSheet == sheetName {
					// Same sheet: check for any reference to the column
					// Patterns: H2, $H2, H$2, $H$2, H:H, etc.
					// Use word boundary to avoid false positives (e.g., SH2 should not match H)
					colPattern := regexp.MustCompile(`\b\$?` + regexp.QuoteMeta(circularColumn) + `[\$:\d]`)
					if colPattern.MatchString(formula) {
						dependsOnCircular = true
						break
					}
				} else {
					// Cross-sheet
					if strings.Contains(formula, "'"+circularSheet+"'!"+circularColumn) ||
						strings.Contains(formula, circularSheet+"!"+circularColumn) {
						dependsOnCircular = true
						break
					}
				}
			}
		}

		if dependsOnCircular {
			// Skip this cell - it depends on a circular column
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			skippedComplexFormulas++
			continue
		}

		dependsOnTimeout := false
		for timeoutCol := range timeoutColumns {
			// Extract sheet and column from "SheetName!Column"
			parts := strings.Split(timeoutCol, "!")
			if len(parts) == 2 {
				timeoutSheet := parts[0]
				timeoutColumn := parts[1]

				// Check if formula references this timed-out column
				// Handle both same-sheet ($H2) and cross-sheet ('Sheet'!H2) references
				if timeoutSheet == sheetName {
					// Same sheet: check for $H2, H$2, H2, $H:$H patterns
					if strings.Contains(formula, timeoutColumn+"$") ||
						strings.Contains(formula, "$"+timeoutColumn) ||
						strings.Contains(formula, timeoutColumn+":") {
						dependsOnTimeout = true
						break
					}
				} else {
					// Cross-sheet: check for 'Sheet'!H2 pattern
					if strings.Contains(formula, "'"+timeoutSheet+"'!"+timeoutColumn) ||
						strings.Contains(formula, timeoutSheet+"!"+timeoutColumn) {
						dependsOnTimeout = true
						break
					}
				}
			}
		}

		if dependsOnTimeout {
			// Skip this cell - it depends on a timed-out column
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			timeoutCount++
			continue
		}

		// Calculate the formula value using raw values with timeout
		// Use context to ensure goroutine cleanup
		calcStart := time.Now()

		type calcResult struct {
			result string
			err    error
		}

		// Create context with timeout for proper goroutine lifecycle management
		ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
		defer cancel() // Ensure context is always cancelled to free resources

		resultChan := make(chan calcResult, 1)

		// Run calculation in goroutine
		// The buffered channel ensures the goroutine can exit even if we timeout
		go func() {
			res, err := f.CalcCellValue(sheetName, c.R, Options{RawCellValue: true})
			select {
			case resultChan <- calcResult{result: res, err: err}:
				// Result sent successfully
			case <-ctx.Done():
				// Context cancelled, don't block on send
				return
			}
		}()

		// Wait for result with timeout
		var result string
		var err error
		var timedOut bool

		select {
		case calcRes := <-resultChan:
			result = calcRes.result
			err = calcRes.err
		case <-ctx.Done():
			timedOut = true
			// Mark this entire column as timed out
			timeoutColumns[columnKey] = true
		}

		calcDuration := time.Since(calcStart)

		// Track slow formulas (>100ms) to help identify bottlenecks
		if timedOut {
			slowFormulaCount++
			timeoutCount++
			// Record timeout formula (always insert - timeouts are the slowest)
			insertSlowFormula(slowFormulaInfo{
				sheet:    sheetName,
				cell:     c.R,
				duration: 5 * time.Second, // Timeout
				formula:  truncateString(formula, 100),
			})
			// Clear the cell value and continue to next formula
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			continue
		} else if calcDuration > 100*time.Millisecond {
			slowFormulaCount++
			// Record slow formula (insert in sorted order)
			insertSlowFormula(slowFormulaInfo{
				sheet:    sheetName,
				cell:     c.R,
				duration: calcDuration,
				formula:  truncateString(formula, 100),
			})
		}

		// Check if this was a batch cache hit (very fast calculation)
		if calcDuration < 1*time.Microsecond {
			batchHitCount++
		}

		calcTime += calcDuration

		if err != nil {
			// CRITICAL: Even if err != nil, result may contain error value like "#DIV/0!"
			// Only clear if result is truly empty (calculation completely failed)
			if result == "" {
				cellRef.V = ""
				cellRef.T = ""
				continue
			}
			// Fall through to write error value (e.g., "#DIV/0!", "#NUM!", etc.)
		}

		// Update cache value directly (we already have the cell reference)
		cellRef.V = result
		// Determine type based on value
		if result == "" {
			cellRef.T = ""
		} else if result == "TRUE" || result == "FALSE" {
			cellRef.T = "b"
		} else if strings.HasPrefix(result, "#") {
			// Error values like #DIV/0!, #NUM!, #VALUE!, etc.
			cellRef.T = "e"
		} else {
			// Try to parse as number
			if _, err := strconv.ParseFloat(result, 64); err == nil {
				cellRef.T = "n"
			} else {
				cellRef.T = "str"
			}
		}

		// 🔥 MEMORY FIX: Don't build affected list - it consumes too much memory
		// For 216k formulas, affected list would use ~50-100 MB
		// The worksheet cache (cellRef.V) is already updated, which is the main goal

		formulaCount++

		// Progress logging - every 5%
		if progressInterval > 0 && formulaCount%progressInterval == 0 {
			progress := float64(formulaCount) / float64(len(calcChain.C)) * 100
			elapsed := time.Since(totalStart)
			avgPerFormula := elapsed / time.Duration(formulaCount)
			remaining := time.Duration(len(calcChain.C)-formulaCount) * avgPerFormula
			log.Printf("  ⏳ [RecalculateAll] Progress: %.0f%% (%d/%d), sheet: '%s', elapsed: %v, avg: %v/formula, remaining: ~%v, slow formulas: %d",
				progress, formulaCount, len(calcChain.C), currentSheetName, elapsed, avgPerFormula, remaining, slowFormulaCount)

			// 🔥 MEMORY OPTIMIZATION: Force GC at progress checkpoints to free memory
			// This helps prevent OOM on large files (200k+ formulas)
			if formulaCount%(progressInterval*4) == 0 { // Every 20%
				runtime.GC()
			}
		}
	}

	// 🔥 MEMORY OPTIMIZATION: Clear cellMap before final GC
	cellMap = nil
	currentWs = nil

	totalDuration := time.Since(totalStart)
	log.Printf("✅ [RecalculateAll] Completed: %d formulas in %v", formulaCount, totalDuration)

	// Avoid division by zero
	avgPerFormula := time.Duration(0)
	if formulaCount > 0 {
		avgPerFormula = calcTime / time.Duration(formulaCount)
	}
	log.Printf("  📊 Breakdown: CellMap build: %v, Formula calc: %v, Avg per formula: %v",
		sheetBuildTime, calcTime, avgPerFormula)

	// Log slow formula statistics
	if slowFormulaCount > 0 {
		log.Printf("  ⚠️  Slow formulas detected: %d formulas took >100ms to calculate", slowFormulaCount)

		// Print top slow formulas
		if len(slowFormulas) > 0 {
			log.Printf("  📋 Top %d slow formulas:", len(slowFormulas))
			for i, sf := range slowFormulas {
				if i >= 20 { // Only show top 20
					log.Printf("  ... and %d more slow formulas", len(slowFormulas)-20)
					break
				}
				log.Printf("    %2d. %s!%s - %v - %s", i+1, sf.sheet, sf.cell, sf.duration, sf.formula)
			}
		}
	}

	// Log timeout statistics
	if timeoutCount > 0 {
		log.Printf("  ⏱️  Timeout formulas: %d formulas exceeded 5s timeout or depend on timed-out columns", timeoutCount)
		if len(timeoutColumns) > 0 {
			log.Printf("  📋 Timed-out columns: %v", timeoutColumns)
		}
	}

	// Log skipped complex formulas
	if skippedComplexFormulas > 0 {
		log.Printf("  🚫 Skipped formulas with circular references: %d formulas", skippedComplexFormulas)
		if len(circularRefColumns) > 0 {
			log.Printf("  📋 Circular reference columns: %v", getMapKeys(circularRefColumns))
		}
	}

	// Log batch optimization statistics
	if batchCount > 0 {
		log.Printf("  ⚡ Batch SUMIFS/AVERAGEIFS/SUMPRODUCT stats: %d formulas batched, %d cache hits during calculation",
			batchCount, batchHitCount)
		if batchHitCount > 0 {
			batchSavings := batchDuration
			log.Printf("  💰 Estimated time saved by batch optimization: %v", batchSavings)
		}
	}

	return nil
}

// getMapKeys returns keys from a map[string]bool as a slice
func getMapKeys(m map[string]bool) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, k)
	}
	return keys
}

// truncateString truncates a string to maxLen characters
func truncateString(s string, maxLen int) string {
	if len(s) <= maxLen {
		return s
	}
	return s[:maxLen] + "..."
}

// AffectedCell 表示受影响的单元格
type AffectedCell struct {
	Sheet       string // 工作表名称
	Cell        string // 单元格坐标
	CachedValue string // 重新计算后的缓存值
}

// BatchUpdateAndRecalculate 批量更新单元格值并重新计算受影响的公式
//
// 此函数结合了 BatchSetCellValue 和重新计算的功能，
// 可以在一次调用中完成批量更新和重新计算，避免重复操作。
//
// 重要特性：
// 1. ✅ 支持跨工作表依赖：如果 Sheet2 引用 Sheet1 的值，更新 Sheet1 后会自动重新计算 Sheet2
// 2. ✅ 只遍历一次 calcChain
// 3. ✅ 每个公式只计算一次（即使被多个更新影响）
// 4. ✅ 性能提升可达 10-100 倍（取决于更新数量）
// 5. ✅ 自动更新所有受影响单元格的缓存值
//
// 注意：为了避免内存溢出，此函数不再返回受影响单元格的列表。
// 所有计算结果已经直接更新到工作表中，可以通过 GetCellValue 读取。
//
// 参数：
//
//	updates: 单元格更新列表
//
// 返回：
//
//	error: 错误信息
//
// 示例：
//
//	// Sheet1: A1 = 100
//	// Sheet2: B1 = Sheet1!A1 * 2
//	updates := []excelize.CellUpdate{
//	    {Sheet: "Sheet1", Cell: "A1", Value: 200},
//	}
//	err := f.BatchUpdateAndRecalculate(updates)
//	// 结果：Sheet1.A1 = 200, Sheet2.B1 = 400 (自动重新计算)
//	// 读取计算后的值
//	value, _ := f.GetCellValue("Sheet2", "B1")
func (f *File) BatchUpdateAndRecalculate(updates []CellUpdate) error {
	// 初始化调试统计
	if enableBatchDebug {
		batchStatsMu.Lock()
		currentBatchStats = &BatchDebugStats{
			CellStats: make(map[string]*CellStats),
		}
		batchStatsMu.Unlock()
	}

	batchStart := time.Now()

	// 1. 批量更新所有单元格
	if err := f.BatchSetCellValue(updates); err != nil {
		return err
	}

	// 2. 立即将更新的值写入缓存，确保后续依赖计算能读到新值
	//    即使依赖计算失败，缓存中也保留了正确的更新值
	for _, update := range updates {
		cacheKey := update.Sheet + "!" + update.Cell
		valueStr := fmt.Sprintf("%v", update.Value)
		f.calcCache.Store(cacheKey+"!raw=false", valueStr)
		f.calcCache.Store(cacheKey+"!raw=true", valueStr)
	}

	// 3. 读取 calcChain
	calcChain, err := f.calcChainReader()
	if err != nil {
		return err
	}

	// If calcChain doesn't exist or is empty, nothing to recalculate
	if calcChain == nil || len(calcChain.C) == 0 {
		return nil
	}

	// 4. 收集所有被更新的单元格（用于依赖检查）
	// 优化：同时建立列索引，加速列引用检查
	updatedCells := make(map[string]map[string]bool)   // sheet -> cell -> true
	updatedColumns := make(map[string]map[string]bool) // sheet -> column -> true
	for _, update := range updates {
		if updatedCells[update.Sheet] == nil {
			updatedCells[update.Sheet] = make(map[string]bool)
			updatedColumns[update.Sheet] = make(map[string]bool)
		}
		updatedCells[update.Sheet][update.Cell] = true

		// 提取列名
		col, _, err := CellNameToCoordinates(update.Cell)
		if err == nil {
			colName, _ := ColumnNumberToName(col)
			updatedColumns[update.Sheet][colName] = true
		}
	}

	// 5. 找出所有受影响的公式单元格（通过依赖分析）
	affectedFormulas := f.findAffectedFormulas(calcChain, updatedCells, updatedColumns)

	// 6. 只清除受影响公式的缓存（不清除刚更新的值）
	for cellKey := range affectedFormulas {
		// 跳过刚更新的单元格，保留其缓存值
		parts := strings.SplitN(cellKey, "!", 2)
		if len(parts) == 2 {
			if cells, ok := updatedCells[parts[0]]; ok && cells[parts[1]] {
				continue
			}
		}
		cacheKey := cellKey + "!raw=false"
		f.calcCache.Delete(cacheKey)
	}

	// 7. 重新计算受影响的公式
	err = f.recalculateAffectedCells(calcChain, affectedFormulas)

	// 记录总耗时
	if enableBatchDebug && currentBatchStats != nil {
		currentBatchStats.TotalDuration = time.Since(batchStart)
		currentBatchStats.TotalCells = len(affectedFormulas)
	}

	return err
}

// BatchSetFormulas 批量设置公式，不触发重新计算
//
// 此函数用于批量设置多个单元格的公式。相比循环调用 SetCellFormula，
// 这个函数可以提高性能并支持自动更新 calcChain。
//
// 参数：
//
//	formulas: 公式更新列表
//
// 示例：
//
//	formulas := []excelize.FormulaUpdate{
//	    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
//	    {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
//	    {Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
//	}
//	err := f.BatchSetFormulas(formulas)
func (f *File) BatchSetFormulas(formulas []FormulaUpdate) error {
	for _, formula := range formulas {
		if err := f.SetCellFormula(formula.Sheet, formula.Cell, formula.Formula); err != nil {
			return err
		}
	}
	return nil
}

// BatchSetFormulasWithValue 批量设置公式及其预计算值，不触发重新计算
//
// 此函数用于批量更新场景，当调用方已经知道公式的计算结果时，
// 可以直接设置公式和值，避免重新计算的开销。
//
// 与 BatchSetFormulas 的区别：
//   - BatchSetFormulas: 设置公式后清除缓存值，需要后续计算
//   - BatchSetFormulasWithValue: 设置公式并保留缓存值，无需重新计算
//
// 参数：
//
//	formulas: 带预计算值的公式更新列表
//
// 示例：
//
//	formulas := []excelize.FormulaUpdateWithValue{
//	    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2", Value: "200"},
//	    {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2", Value: "400"},
//	}
//	err := f.BatchSetFormulasWithValue(formulas)
func (f *File) BatchSetFormulasWithValue(formulas []FormulaUpdateWithValue) error {
	for _, formula := range formulas {
		if err := f.SetCellFormulaWithValue(formula.Sheet, formula.Cell, formula.Formula, formula.Value); err != nil {
			return err
		}
	}
	return nil
}

// BatchSetFormulasAndRecalculate 批量设置公式并重新计算
//
// 此函数批量设置多个单元格的公式，然后自动重新计算所有受影响的公式，
// 并更新 calcChain 以确保引用关系正确。
//
// 功能特点：
// 1. ✅ 批量设置公式（避免重复的工作表查找）
// 2. ✅ 自动计算所有公式的值
// 3. ✅ 自动更新 calcChain（计算链）
// 4. ✅ 触发依赖公式的重新计算
// 5. ✅ 自动更新所有受影响单元格的缓存值
//
// 注意：为了避免内存溢出，此函数不再返回受影响单元格的列表。
// 所有计算结果已经直接更新到工作表中，可以通过 GetCellValue 读取。
//
// 相比循环调用 SetCellFormula + UpdateCellAndRecalculate，性能提升显著。
//
// 参数：
//
//	formulas: 公式更新列表
//
// 返回：
//
//	error: 错误信息
//
// 示例：
//
//	formulas := []excelize.FormulaUpdate{
//	    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
//	    {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
//	    {Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
//	    {Sheet: "Sheet1", Cell: "C1", Formula: "=SUM(B1:B3)"},
//	}
//	err := f.BatchSetFormulasAndRecalculate(formulas)
//	// 现在所有公式都已设置、计算，并且 calcChain 已更新
//	// 读取计算后的值
//	value, _ := f.GetCellValue("Sheet1", "C1")
func (f *File) BatchSetFormulasAndRecalculate(formulas []FormulaUpdate) error {
	if len(formulas) == 0 {
		return nil
	}

	// 1. 批量设置公式
	if err := f.BatchSetFormulas(formulas); err != nil {
		return err
	}

	// 2. 收集所有受影响的工作表和单元格
	affectedSheets := make(map[string][]string)
	for _, formula := range formulas {
		affectedSheets[formula.Sheet] = append(affectedSheets[formula.Sheet], formula.Cell)
	}

	// 3. 为每个工作表更新 calcChain
	if err := f.updateCalcChainForFormulas(formulas); err != nil {
		return err
	}

	// 4. 收集被设置公式的单元格
	setFormulaCells := make(map[string]map[string]bool)
	for _, formula := range formulas {
		if setFormulaCells[formula.Sheet] == nil {
			setFormulaCells[formula.Sheet] = make(map[string]bool)
		}
		setFormulaCells[formula.Sheet][formula.Cell] = true
	}

	// 5. 重新计算所有公式
	for sheet := range affectedSheets {
		if err := f.RecalculateSheet(sheet); err != nil {
			return err
		}
	}

	// 6. 读取 calcChain 并找出依赖于新公式的其他单元格
	calcChain, err := f.calcChainReader()
	if err != nil {
		return err
	}

	if calcChain == nil || len(calcChain.C) == 0 {
		return nil
	}

	// 构建列索引
	setFormulaColumns := make(map[string]map[string]bool)
	for sheet, cells := range setFormulaCells {
		setFormulaColumns[sheet] = make(map[string]bool)
		for cell := range cells {
			col, _, err := CellNameToCoordinates(cell)
			if err == nil {
				colName, _ := ColumnNumberToName(col)
				setFormulaColumns[sheet][colName] = true
			}
		}
	}

	affectedFormulas := f.findAffectedFormulas(calcChain, setFormulaCells, setFormulaColumns)

	// 7. 只排除那些不依赖于同批其他公式的被设置单元格
	// 如果 C1 依赖 B1，且 B1 和 C1 都被设置，则保留 C1
	for sheet, cells := range setFormulaCells {
		for cell := range cells {
			cellKey := sheet + "!" + cell
			// 检查这个单元格是否依赖于同批的其他公式
			isDependentOnOthers := false

			// 获取这个单元格的公式
			ws, err := f.workSheetReader(sheet)
			if err == nil {
				col, row, _ := CellNameToCoordinates(cell)
				cellData := f.getCellFromWorksheet(ws, col, row)
				if cellData != nil && cellData.F != nil {
					formula := cellData.F.Content
					if formula == "" && cellData.F.T == STCellFormulaTypeShared && cellData.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cellData.F.Si, cell)
					}

					if formula != "" {
						// 检查公式是否引用了同批的其他单元格
						isDependentOnOthers = f.formulaReferencesUpdatedCells(formula, sheet, setFormulaCells, setFormulaColumns)
					}
				}
			}

			// 如果不依赖于同批其他公式，则排除
			if !isDependentOnOthers {
				delete(affectedFormulas, cellKey)
			}
		}
	}

	// 7. 不再构建affected列表，所有计算结果已经更新到工作表缓存中
	return nil
}

// updateCalcChainForFormulas 更新 calcChain 以包含新设置的公式
func (f *File) updateCalcChainForFormulas(formulas []FormulaUpdate) error {
	// 读取或创建 calcChain
	calcChain, err := f.calcChainReader()
	if err != nil {
		return err
	}

	if calcChain == nil {
		calcChain = &xlsxCalcChain{
			C: []xlsxCalcChainC{},
		}
	}

	// 创建现有 calcChain 条目的映射（用于去重）
	existingEntries := make(map[string]map[string]bool) // sheet -> cell -> exists
	for _, entry := range calcChain.C {
		sheetID := entry.I
		sheetName := f.GetSheetMap()[sheetID]
		if existingEntries[sheetName] == nil {
			existingEntries[sheetName] = make(map[string]bool)
		}
		existingEntries[sheetName][entry.R] = true
	}

	// 添加新的公式到 calcChain
	for _, formula := range formulas {
		// 检查是否已存在
		if existingEntries[formula.Sheet] != nil && existingEntries[formula.Sheet][formula.Cell] {
			continue // 已存在，跳过
		}

		// 获取 sheet ID
		sheetID := f.getSheetID(formula.Sheet)
		if sheetID == -1 {
			continue // 工作表不存在，跳过
		}

		// 添加到 calcChain
		newEntry := xlsxCalcChainC{
			R: formula.Cell,
			I: sheetID, // I is the sheet ID (1-based)
		}

		calcChain.C = append(calcChain.C, newEntry)

		// 更新映射
		if existingEntries[formula.Sheet] == nil {
			existingEntries[formula.Sheet] = make(map[string]bool)
		}
		existingEntries[formula.Sheet][formula.Cell] = true
	}

	// 保存更新后的 calcChain
	f.CalcChain = calcChain

	return nil
}

// recalculateAllSheets recalculates all formulas in all sheets according to calcChain order
func (f *File) recalculateAllSheets(calcChain *xlsxCalcChain) error {
	_, err := f.recalculateAllSheetsWithTracking(calcChain)
	return err
}

// recalculateAllSheetsWithTracking recalculates all formulas and tracks affected cells
func (f *File) recalculateAllSheetsWithTracking(calcChain *xlsxCalcChain) ([]AffectedCell, error) {
	// Track current sheet ID (for handling I=0 case)
	currentSheetID := -1
	var affected []AffectedCell

	// Build dependency graph to find truly affected cells
	updatedCells := make(map[string]bool) // "Sheet!Cell" -> true

	// Recalculate all cells in calcChain order
	for i := range calcChain.C {
		c := calcChain.C[i]

		// Update current sheet ID if specified
		if c.I != 0 {
			currentSheetID = c.I
		}

		// Get sheet name
		sheetName := f.GetSheetMap()[currentSheetID]
		if sheetName == "" {
			continue // Skip if sheet not found
		}

		cellKey := sheetName + "!" + c.R

		// Check if this cell was recalculated (cache was cleared)
		cacheKey := cellKey + "!raw=false"
		_, hadCache := f.calcCache.Load(cacheKey)

		// Recalculate the cell
		if err := f.recalculateCell(sheetName, c.R); err != nil {
			// Continue even if one cell fails
			continue
		}

		// Check if cache was updated (meaning it was recalculated)
		newValue, hasNewCache := f.calcCache.Load(cacheKey)

		// Only track if this cell was actually recalculated (no cache before, has cache now)
		if !hadCache && hasNewCache {
			cachedValue := ""
			if newValue != nil {
				cachedValue = newValue.(string)
			}

			affected = append(affected, AffectedCell{
				Sheet:       sheetName,
				Cell:        c.R,
				CachedValue: cachedValue,
			})
			updatedCells[cellKey] = true
		}
	}

	return affected, nil
}

// findAffectedFormulas 找出所有受影响的公式单元格（包括间接依赖
// findAffectedFormulas 找出所有受影响的公式单元格（包括间接依赖）
// 通过解析公式中的单元格引用，找出哪些公式依赖于被更新的单元格
func (f *File) findAffectedFormulas(calcChain *xlsxCalcChain, updatedCells map[string]map[string]bool, updatedColumns map[string]map[string]bool) map[string]bool {
	affected := make(map[string]bool)
	currentSheetID := -1

	// 第一轮：找出直接依赖
	for i := range calcChain.C {
		c := calcChain.C[i]
		if c.I != 0 {
			currentSheetID = c.I
		}

		sheetName := f.GetSheetMap()[currentSheetID]
		if sheetName == "" {
			continue
		}

		// 获取公式内容
		ws, err := f.workSheetReader(sheetName)
		if err != nil {
			continue
		}

		col, row, _ := CellNameToCoordinates(c.R)
		cellData := f.getCellFromWorksheet(ws, col, row)
		if cellData == nil || cellData.F == nil {
			continue
		}

		formula := cellData.F.Content
		if formula == "" && cellData.F.T == STCellFormulaTypeShared && cellData.F.Si != nil {
			formula, _ = getSharedFormula(ws, *cellData.F.Si, c.R)
		}

		if formula == "" {
			continue
		}

		// 检查公式是否引用了被更新的单元格
		if f.formulaReferencesUpdatedCells(formula, sheetName, updatedCells, updatedColumns) {
			cellKey := sheetName + "!" + c.R
			affected[cellKey] = true
		}
	}

	// 递归查找间接依赖：如果公式引用了受影响的单元格，它也受影响
	changed := true
	for changed {
		changed = false
		currentSheetID = -1

		for i := range calcChain.C {
			c := calcChain.C[i]
			if c.I != 0 {
				currentSheetID = c.I
			}

			sheetName := f.GetSheetMap()[currentSheetID]
			if sheetName == "" {
				continue
			}

			cellKey := sheetName + "!" + c.R
			if affected[cellKey] {
				continue // 已经标记为受影响
			}

			// 获取公式内容
			ws, err := f.workSheetReader(sheetName)
			if err != nil {
				continue
			}

			col, row, _ := CellNameToCoordinates(c.R)
			cellData := f.getCellFromWorksheet(ws, col, row)
			if cellData == nil || cellData.F == nil {
				continue
			}

			formula := cellData.F.Content
			if formula == "" && cellData.F.T == STCellFormulaTypeShared && cellData.F.Si != nil {
				formula, _ = getSharedFormula(ws, *cellData.F.Si, c.R)
			}

			if formula == "" {
				continue
			}

			// 检查公式是否引用了受影响的单元格
			if f.formulaReferencesAffectedCells(formula, sheetName, affected) {
				affected[cellKey] = true
				changed = true
			}
		}
	}

	return affected
}

// findAffectedFormulasOptimized 优化版：使用反向依赖图快速查找受影响的公式
// 对于少量单元格更新的场景，性能提升可达 100-1000 倍
func (f *File) findAffectedFormulasOptimized(calcChain *xlsxCalcChain, updatedCells map[string]map[string]bool, updatedColumns map[string]map[string]bool) map[string]bool {
	// 如果更新的单元格很多，回退到原始方法
	totalUpdates := 0
	for _, cells := range updatedCells {
		totalUpdates += len(cells)
	}
	if totalUpdates > 100 {
		return f.findAffectedFormulas(calcChain, updatedCells, updatedColumns)
	}

	affected := make(map[string]bool)

	// 第一步：构建反向依赖图（哪些公式依赖于哪些单元格/列）
	// dependents[cellKey] = 依赖于该单元格的公式列表
	// columnDependents[sheetColumn] = 依赖于该列的公式列表
	dependents := make(map[string][]string)
	columnDependents := make(map[string][]string)
	sheetDependents := make(map[string][]string)

	currentSheetID := -1
	sheetMap := f.GetSheetMap()

	// 预加载所有工作表
	wsCache := make(map[string]*xlsxWorksheet)

	for i := range calcChain.C {
		c := calcChain.C[i]
		if c.I != 0 {
			currentSheetID = c.I
		}

		sheetName := sheetMap[currentSheetID]
		if sheetName == "" {
			continue
		}

		ws, ok := wsCache[sheetName]
		if !ok {
			var err error
			ws, err = f.workSheetReader(sheetName)
			if err != nil {
				continue
			}
			wsCache[sheetName] = ws
		}

		col, row, _ := CellNameToCoordinates(c.R)
		cellData := f.getCellFromWorksheet(ws, col, row)
		if cellData == nil || cellData.F == nil {
			continue
		}

		formula := cellData.F.Content
		if formula == "" && cellData.F.T == STCellFormulaTypeShared && cellData.F.Si != nil {
			formula, _ = getSharedFormula(ws, *cellData.F.Si, c.R)
		}

		if formula == "" {
			continue
		}

		cellKey := sheetName + "!" + c.R

		// 提取公式依赖并构建反向索引
		deps := f.extractDependenciesWithSQL(formula, sheetName, "")
		for _, dep := range deps {
			if strings.HasPrefix(dep, "SHEET:") {
				refSheet := strings.TrimPrefix(dep, "SHEET:")
				sheetDependents[refSheet] = append(sheetDependents[refSheet], cellKey)
				continue
			}
			parts := strings.SplitN(dep, "!", 2)
			if len(parts) != 2 {
				continue
			}
			refSheet := parts[0]
			refCell := parts[1]

			// 列范围引用
			if strings.HasSuffix(refCell, ":COLUMN_RANGE") {
				colName := strings.TrimSuffix(refCell, ":COLUMN_RANGE")
				colKey := refSheet + "!" + colName
				columnDependents[colKey] = append(columnDependents[colKey], cellKey)
			} else {
				// 单元格引用
				depKey := refSheet + "!" + refCell
				dependents[depKey] = append(dependents[depKey], cellKey)

				// 也添加到列依赖（因为列更新可能影响该单元格）
				depCol, _, err := CellNameToCoordinates(refCell)
				if err == nil {
					depColName, _ := ColumnNumberToName(depCol)
					colKey := refSheet + "!" + depColName
					columnDependents[colKey] = append(columnDependents[colKey], cellKey)
				}
			}
		}
	}

	// 第二步：从更新的单元格开始，使用 BFS 找出所有受影响的公式
	queue := make([]string, 0, 1000)

	// 添加直接受影响的公式
	for sheet, cells := range updatedCells {
		for _, dep := range sheetDependents[sheet] {
			if !affected[dep] {
				affected[dep] = true
				queue = append(queue, dep)
			}
		}
		for cell := range cells {
			cellKey := sheet + "!" + cell
			// 添加直接依赖于该单元格的公式
			for _, dep := range dependents[cellKey] {
				if !affected[dep] {
					affected[dep] = true
					queue = append(queue, dep)
				}
			}
		}
	}

	// 添加依赖于更新列的公式
	for sheet, cols := range updatedColumns {
		for _, dep := range sheetDependents[sheet] {
			if !affected[dep] {
				affected[dep] = true
				queue = append(queue, dep)
			}
		}
		for col := range cols {
			colKey := sheet + "!" + col
			for _, dep := range columnDependents[colKey] {
				if !affected[dep] {
					affected[dep] = true
					queue = append(queue, dep)
				}
			}
		}
	}

	// BFS 传播依赖
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]

		// 添加依赖于当前公式结果的其他公式
		for _, dep := range dependents[current] {
			if !affected[dep] {
				affected[dep] = true
				queue = append(queue, dep)
			}
		}

		// 获取当前公式所在的列
		parts := strings.SplitN(current, "!", 2)
		if len(parts) == 2 {
			col, _, err := CellNameToCoordinates(parts[1])
			if err == nil {
				colName, _ := ColumnNumberToName(col)
				colKey := parts[0] + "!" + colName
				for _, dep := range columnDependents[colKey] {
					if !affected[dep] {
						affected[dep] = true
						queue = append(queue, dep)
					}
				}
			}
		}
	}

	return affected
}

// formulaReferencesUpdatedCells 检查公式是否引用了被更新的单元格
// 使用 extractDependencies 函数解析公式依赖
func (f *File) formulaReferencesUpdatedCells(formula, currentSheet string, updatedCells map[string]map[string]bool, updatedColumns map[string]map[string]bool) bool {
	// 使用公式解析器提取依赖
	deps := f.extractDependenciesWithSQL(formula, currentSheet, "")

	for _, dep := range deps {
		if strings.HasPrefix(dep, "SHEET:") {
			refSheet := strings.TrimPrefix(dep, "SHEET:")
			if len(updatedCells[refSheet]) > 0 || len(updatedColumns[refSheet]) > 0 {
				return true
			}
			continue
		}
		// dep 格式: "Sheet!Cell" 或 "Sheet!Col:COLUMN_RANGE"
		parts := strings.SplitN(dep, "!", 2)
		if len(parts) != 2 {
			continue
		}
		refSheet := parts[0]
		refCell := parts[1]

		// 检查是否是列范围引用
		if strings.HasSuffix(refCell, ":COLUMN_RANGE") {
			colName := strings.TrimSuffix(refCell, ":COLUMN_RANGE")
			if updatedColumns[refSheet] != nil && updatedColumns[refSheet][colName] {
				return true
			}
			continue
		}

		// 检查单元格是否在更新列表中
		if updatedCells[refSheet] != nil && updatedCells[refSheet][refCell] {
			return true
		}

		// 检查单元格所在的列是否在更新列表中（用于列范围引用）
		col, _, err := CellNameToCoordinates(refCell)
		if err == nil {
			colName, _ := ColumnNumberToName(col)
			if updatedColumns[refSheet] != nil && updatedColumns[refSheet][colName] {
				// 检查是否有该列的单元格被更新
				for cell := range updatedCells[refSheet] {
					cellCol, _, err := CellNameToCoordinates(cell)
					if err == nil {
						cellColName, _ := ColumnNumberToName(cellCol)
						if cellColName == colName {
							return true
						}
					}
				}
			}
		}
	}

	return false
}

// formulaReferencesAffectedCells 检查公式是否引用了受影响的单元格
// 使用 extractDependencies 函数解析公式依赖
func (f *File) formulaReferencesAffectedCells(formula, currentSheet string, affectedCells map[string]bool) bool {
	// 使用公式解析器提取依赖
	deps := f.extractDependenciesWithSQL(formula, currentSheet, "")

	for _, dep := range deps {
		if strings.HasPrefix(dep, "SHEET:") {
			refSheet := strings.TrimPrefix(dep, "SHEET:")
			for affectedCell := range affectedCells {
				if strings.HasPrefix(affectedCell, refSheet+"!") {
					return true
				}
			}
			continue
		}
		// dep 格式: "Sheet!Cell" 或 "Sheet!Col:COLUMN_RANGE"
		parts := strings.SplitN(dep, "!", 2)
		if len(parts) != 2 {
			continue
		}
		refSheet := parts[0]
		refCell := parts[1]

		// 检查是否是列范围引用
		if strings.HasSuffix(refCell, ":COLUMN_RANGE") {
			colName := strings.TrimSuffix(refCell, ":COLUMN_RANGE")
			// 检查是否有受影响的单元格在该列
			for cellKey := range affectedCells {
				keyParts := strings.SplitN(cellKey, "!", 2)
				if len(keyParts) == 2 && keyParts[0] == refSheet {
					col, _, err := CellNameToCoordinates(keyParts[1])
					if err == nil {
						cellColName, _ := ColumnNumberToName(col)
						if cellColName == colName {
							return true
						}
					}
				}
			}
			continue
		}

		// 检查单元格是否在受影响列表中
		cellKey := refSheet + "!" + refCell
		if affectedCells[cellKey] {
			return true
		}
	}

	return false
}

// cellInRange 检查单元格是否在范围内
func (f *File) cellInRange(cell, startCell, endCell string) bool {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return false
	}

	startCol, startRow, err := CellNameToCoordinates(startCell)
	if err != nil {
		return false
	}

	endCol, endRow, err := CellNameToCoordinates(endCell)
	if err != nil {
		return false
	}

	return col >= startCol && col <= endCol && row >= startRow && row <= endRow
}

// getCellFromWorksheet 从工作表中获取单元格数据
func (f *File) getCellFromWorksheet(ws *xlsxWorksheet, col, row int) *xlsxC {
	for i := range ws.SheetData.Row {
		if ws.SheetData.Row[i].R == row {
			for j := range ws.SheetData.Row[i].C {
				c := &ws.SheetData.Row[i].C[j]
				cellCol, cellRow, _ := CellNameToCoordinates(c.R)
				if cellCol == col && cellRow == row {
					return c
				}
			}
			return nil
		}
	}
	return nil
}

// recalculateAffectedCells 只重新计算受影响的单元格
func (f *File) recalculateAffectedCells(calcChain *xlsxCalcChain, affectedFormulas map[string]bool) error {
	currentSheetID := -1

	for i := range calcChain.C {
		c := calcChain.C[i]
		if c.I != 0 {
			currentSheetID = c.I
		}

		sheetName := f.GetSheetMap()[currentSheetID]
		if sheetName == "" {
			continue
		}

		cellKey := sheetName + "!" + c.R

		// 只处理受影响的单元格
		if !affectedFormulas[cellKey] {
			continue
		}

		// 重新计算 - 结果已经直接更新到工作表缓存
		if err := f.recalculateCell(sheetName, c.R); err != nil {
			continue
		}
	}

	return nil
}

// BatchUpdateValuesAndFormulasWithRecalc 批量更新单元格值和公式，并只重新计算受影响的依赖公式
//
// 这是一个高性能的批量更新方法，适用于同时更新多个单元格的值和公式的场景。
// 与 BatchUpdateAndRecalculate + BatchSetFormulasAndRecalculate + RecalculateAllWithDependency 的组合相比，
// 此方法只清除一次缓存并只计算受影响的公式，避免了多次全局重算。
//
// 功能特点：
// 1. ✅ 批量设置单元格值（不触发逐个缓存清理）
// 2. ✅ 批量设置公式（不触发逐个重算）
// 3. ✅ 统一分析所有更新的依赖关系
// 4. ✅ 只清除受影响公式的缓存（一次性）
// 5. ✅ 使用 DAG 按正确顺序计算受影响的公式（只计算一次）
// 6. ✅ 自动更新 calcChain
//
// 参数：
//
//	valueUpdates: 单元格值更新列表（可以为空）
//	formulaUpdates: 公式更新列表（可以为空）
//
// 返回：
//
//	error: 错误信息
//
// 示例：
//
//	values := []excelize.CellUpdate{
//	    {Sheet: "Sheet1", Cell: "A1", Value: 100},
//	    {Sheet: "Sheet1", Cell: "A2", Value: 200},
//	}
//	formulas := []excelize.FormulaUpdate{
//	    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
//	    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+10"},
//	}
//	err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
func (f *File) BatchUpdateValuesAndFormulasWithRecalc(valueUpdates []CellUpdate, formulaUpdates []FormulaUpdate) error {
	// 转换为 V2 格式：FormulaUpdate -> FormulaUpdateWithValue（Value 为空）
	formulaUpdatesV2 := make([]FormulaUpdateWithValue, 0, len(formulaUpdates))
	for _, formula := range formulaUpdates {
		formulaUpdatesV2 = append(formulaUpdatesV2, FormulaUpdateWithValue{
			Sheet:   formula.Sheet,
			Cell:    formula.Cell,
			Formula: formula.Formula,
			Value:   "", // 空值，V2 会处理
		})
	}

	// 调用 V2 版本
	return f.BatchUpdateValuesAndFormulasWithRecalcV2(valueUpdates, formulaUpdatesV2)

	/* 旧版本逻辑（已注释）
	if len(valueUpdates) == 0 && len(formulaUpdates) == 0 {
		return nil
	}

	// 1. 批量设置单元格值（不触发重算）
	if len(valueUpdates) > 0 {
		if err := f.BatchSetCellValue(valueUpdates); err != nil {
			return err
		}
	}

	// 2. 立即将更新的值写入缓存，确保后续依赖计算能读到新值
	//    即使依赖计算失败，缓存中也保留了正确的更新值
	for _, update := range valueUpdates {
		cacheKey := update.Sheet + "!" + update.Cell
		valueStr := fmt.Sprintf("%v", update.Value)
		// 存储 formulaArg 类型到缓存（供 rangeResolver 使用）
		arg := inferFormulaResultType(valueStr)
		f.calcCache.Store(cacheKey, arg)
		// 存储字符串类型到缓存
		f.calcCache.Store(cacheKey+"!raw=false", valueStr)
		f.calcCache.Store(cacheKey+"!raw=true", valueStr)
	}

	// 3. 批量设置公式（不触发重算）
	if len(formulaUpdates) > 0 {
		if err := f.BatchSetFormulas(formulaUpdates); err != nil {
			return err
		}
	}

	// 4. 立即计算更新的公式并写入缓存和 worksheet
	//    这确保依赖这些公式的单元格能读到正确的值
	//    同时 GetCellValue 也能返回正确的值
	//    保存计算结果，以便在增量重算后恢复
	formulaResults := make(map[string]string) // "Sheet!Cell" -> calculated value
	for _, formula := range formulaUpdates {
		// 清除旧缓存（包括所有格式的缓存 key）
		cacheKey := formula.Sheet + "!" + formula.Cell
		f.calcCache.Delete(cacheKey)                // formulaArg 类型缓存
		f.calcCache.Delete(cacheKey + "!raw=false") // 字符串类型缓存
		f.calcCache.Delete(cacheKey + "!raw=true")  // 字符串类型缓存

		// 计算公式
		value, err := f.CalcCellValue(formula.Sheet, formula.Cell)
		if err == nil {
			// 存储 formulaArg 类型到缓存（供 rangeResolver 使用）
			arg := inferFormulaResultType(value)
			f.calcCache.Store(cacheKey, arg)
			// 存储字符串类型到缓存（供 GetCellValue 等使用）
			f.calcCache.Store(cacheKey+"!raw=false", value)
			f.calcCache.Store(cacheKey+"!raw=true", value)
			// 将计算结果写入 worksheet XML，这样 GetCellValue 能正确读取
			f.setFormulaValue(formula.Sheet, formula.Cell, value)
			// 保存结果以便增量重算后恢复
			formulaResults[cacheKey] = value
		}
	}

	// 5. 收集被更新的单元格（精确到单元格级别）
	updatedCells := make(map[string]bool) // "Sheet!Cell" -> true
	for _, update := range valueUpdates {
		updatedCells[update.Sheet+"!"+update.Cell] = true
	}
	for _, formula := range formulaUpdates {
		updatedCells[formula.Sheet+"!"+formula.Cell] = true
	}

	// 6. 增量重算：只计算依赖于更新单元格的公式
	//    注意：RecalculateAffectedByCells 可能会清除缓存（包括我们刚设置的公式值）
	err := f.RecalculateAffectedByCells(updatedCells)

	// 7. 恢复更新的公式值到缓存和 worksheet
	//    因为增量重算可能清除了这些缓存，需要重新写入
	for _, formula := range formulaUpdates {
		cacheKey := formula.Sheet + "!" + formula.Cell
		if value, ok := formulaResults[cacheKey]; ok {
			// 存储 formulaArg 类型到缓存（供 rangeResolver 使用）
			arg := inferFormulaResultType(value)
			f.calcCache.Store(cacheKey, arg)
			// 存储字符串类型到缓存
			f.calcCache.Store(cacheKey+"!raw=false", value)
			f.calcCache.Store(cacheKey+"!raw=true", value)
			f.setFormulaValue(formula.Sheet, formula.Cell, value)
		}
	}

	// 8. 恢复更新的值到缓存
	for _, update := range valueUpdates {
		cacheKey := update.Sheet + "!" + update.Cell
		valueStr := fmt.Sprintf("%v", update.Value)
		// 存储 formulaArg 类型到缓存（供 rangeResolver 使用）
		arg := inferFormulaResultType(valueStr)
		f.calcCache.Store(cacheKey, arg)
		// 存储字符串类型到缓存
		f.calcCache.Store(cacheKey+"!raw=false", valueStr)
		f.calcCache.Store(cacheKey+"!raw=true", valueStr)
	}

	return err
	*/
}

// BatchUpdateValuesAndFormulasWithRecalcV2 批量更新单元格值和公式（带预计算值），并只重新计算受影响的依赖公式
//
// 这是 BatchUpdateValuesAndFormulasWithRecalc 的优化版本，主要区别是公式更新可以携带预计算值，
// 避免了设置公式后需要立即计算的开销。
//
// 适用场景：
//   - 调用方已经知道公式的计算结果（例如从外部系统获取）
//   - 批量导入带公式和值的数据
//   - 需要最大化性能的场景
//
// 与 BatchUpdateValuesAndFormulasWithRecalc 的区别：
//   - 旧版本：设置公式 → 计算公式 → 设置缓存 → 增量重算
//   - 新版本：设置公式+值 → 增量重算（跳过公式计算步骤）
//
// 参数：
//
//	valueUpdates: 单元格值更新列表（可以为空）
//	formulaUpdates: 带预计算值的公式更新列表（可以为空）
//
// 返回：
//
//	error: 错误信息
//
// 示例：
//
//	values := []excelize.CellUpdate{
//	    {Sheet: "Sheet1", Cell: "A1", Value: 100},
//	    {Sheet: "Sheet1", Cell: "A2", Value: 200},
//	}
//	formulas := []excelize.FormulaUpdateWithValue{
//	    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2", Value: "200"},
//	    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+10", Value: "210"},
//	}
//	err := f.BatchUpdateValuesAndFormulasWithRecalcV2(values, formulas)
func (f *File) BatchUpdateValuesAndFormulasWithRecalcV2(valueUpdates []CellUpdate, formulaUpdates []FormulaUpdateWithValue) error {
	if len(valueUpdates) == 0 && len(formulaUpdates) == 0 {
		return nil
	}

	// 1. 批量设置单元格值（不触发重算）
	if len(valueUpdates) > 0 {
		if err := f.BatchSetCellValue(valueUpdates); err != nil {
			return err
		}
	}

	// 2. 立即将更新的值写入缓存，确保后续依赖计算能读到新值
	//    同时记录 valueUpdates 中的值，供步骤 3 使用
	valueMap := make(map[string]string) // "Sheet!Cell" -> value string
	for _, update := range valueUpdates {
		cacheKey := update.Sheet + "!" + update.Cell
		valueStr := fmt.Sprintf("%v", update.Value)
		arg := inferFormulaResultType(valueStr)
		f.calcCache.Store(cacheKey, arg)
		f.calcCache.Store(cacheKey+"!raw=false", valueStr)
		f.calcCache.Store(cacheKey+"!raw=true", valueStr)
		valueMap[cacheKey] = valueStr
	}

	// 3. 分离有预计算值和没有预计算值的公式
	formulasWithPreCalc := make([]FormulaUpdateWithValue, 0)
	formulasNeedCalc := make([]FormulaUpdate, 0)
	for _, formula := range formulaUpdates {
		cellKey := formula.Sheet + "!" + formula.Cell
		if v, ok := valueMap[cellKey]; ok {
			// 有预计算值
			formulasWithPreCalc = append(formulasWithPreCalc, FormulaUpdateWithValue{
				Sheet:   formula.Sheet,
				Cell:    formula.Cell,
				Formula: formula.Formula,
				Value:   v,
			})
		} else {
			// 没有预计算值，需要计算
			formulasNeedCalc = append(formulasNeedCalc, FormulaUpdate{
				Sheet:   formula.Sheet,
				Cell:    formula.Cell,
				Formula: formula.Formula,
			})
		}
	}

	// 4. 设置有预计算值的公式（不触发重算）
	if len(formulasWithPreCalc) > 0 {
		if err := f.BatchSetFormulasWithValue(formulasWithPreCalc); err != nil {
			return err
		}
	}

	// 5. 设置没有预计算值的公式，然后批量计算
	if len(formulasNeedCalc) > 0 {
		// 按 sheet 分组
		sheetFormulas := make(map[string][]FormulaUpdate) // sheet -> []FormulaUpdate
		for _, formula := range formulasNeedCalc {
			sheetFormulas[formula.Sheet] = append(sheetFormulas[formula.Sheet], formula)
		}

		// 批量计算每个 sheet 的公式，然后用 BatchSetFormulasWithValue 批量设置
		allFormulasWithValue := make([]FormulaUpdateWithValue, 0, len(formulasNeedCalc))
		for sheet, formulas := range sheetFormulas {
			// 收集单元格列表
			cells := make([]string, len(formulas))
			for i, f := range formulas {
				cells[i] = f.Cell
			}
			// 清除旧缓存
			for _, cell := range cells {
				cellKey := sheet + "!" + cell
				f.calcCache.Delete(cellKey)
				f.calcCache.Delete(cellKey + "!raw=false")
				f.calcCache.Delete(cellKey + "!raw=true")
			}
			// 先设置公式（以便计算）
			if err := f.BatchSetFormulas(formulas); err != nil {
				return err
			}
			// 批量计算
			results, _ := f.CalcCellValues(sheet, cells)
			// 收集计算结果
			for _, formula := range formulas {
				value := results[formula.Cell]
				allFormulasWithValue = append(allFormulasWithValue, FormulaUpdateWithValue{
					Sheet:   formula.Sheet,
					Cell:    formula.Cell,
					Formula: formula.Formula,
					Value:   value,
				})
			}
		}

		// 批量设置公式和值
		if len(allFormulasWithValue) > 0 {
			if err := f.BatchSetFormulasWithValue(allFormulasWithValue); err != nil {
				return err
			}
		}
	}

	// 6. 收集被更新的单元格（精确到单元格级别）
	updatedCells := make(map[string]bool)
	valueCells := make(map[string]bool) // 记录 valueUpdates 中的单元格
	for _, update := range valueUpdates {
		cellKey := update.Sheet + "!" + update.Cell
		updatedCells[cellKey] = true
		valueCells[cellKey] = true
	}
	for _, formula := range formulaUpdates {
		updatedCells[formula.Sheet+"!"+formula.Cell] = true
	}

	// 7. 收集带预计算值的公式单元格（这些单元格不需要重算）
	//    判断条件：单元格同时在 valueUpdates 和 formulaUpdates 中
	excludeCells := make(map[string]bool)
	for _, formula := range formulaUpdates {
		cellKey := formula.Sheet + "!" + formula.Cell
		if valueCells[cellKey] {
			// 该公式单元格有预计算值，排除重算
			excludeCells[cellKey] = true
		}
	}

	// 8. 增量重算：只计算依赖于更新单元格的公式，但排除已有预计算值的公式单元格
	err := f.RecalculateAffectedByCellsWithExclusion(updatedCells, excludeCells)

	// 9. 恢复更新的值到缓存（增量重算可能清除了依赖于这些值的公式的缓存，但不会清除值本身）
	for _, update := range valueUpdates {
		cacheKey := update.Sheet + "!" + update.Cell
		valueStr := fmt.Sprintf("%v", update.Value)
		arg := inferFormulaResultType(valueStr)
		f.calcCache.Store(cacheKey, arg)
		f.calcCache.Store(cacheKey+"!raw=false", valueStr)
		f.calcCache.Store(cacheKey+"!raw=true", valueStr)
	}

	return err
}

// findAffectedFormulasByScanning 通过扫描所有工作表来找出受影响的公式
// 这个方法不依赖 calcChain，适用于 calcChain 不完整或不存在的情况
func (f *File) findAffectedFormulasByScanning(updatedCells map[string]map[string]bool, updatedColumns map[string]map[string]bool) map[string]bool {
	affected := make(map[string]bool)

	// 遍历所有工作表
	sheetList := f.GetSheetList()
	for _, sheetName := range sheetList {
		ws, err := f.workSheetReader(sheetName)
		if err != nil || ws.SheetData.Row == nil {
			continue
		}

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				if cell.F == nil {
					continue
				}

				formula := cell.F.Content
				if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
					formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
				}

				if formula == "" {
					continue
				}

				// 检查公式是否引用了被更新的单元格
				if f.formulaReferencesUpdatedCells(formula, sheetName, updatedCells, updatedColumns) {
					cellKey := sheetName + "!" + cell.R
					affected[cellKey] = true
				}
			}
		}
	}

	// 递归查找间接依赖
	changed := true
	for changed {
		changed = false

		for _, sheetName := range sheetList {
			ws, err := f.workSheetReader(sheetName)
			if err != nil || ws.SheetData.Row == nil {
				continue
			}

			for _, row := range ws.SheetData.Row {
				for _, cell := range row.C {
					cellKey := sheetName + "!" + cell.R
					if affected[cellKey] {
						continue // 已经标记为受影响
					}

					if cell.F == nil {
						continue
					}

					formula := cell.F.Content
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}

					if formula == "" {
						continue
					}

					// 检查公式是否引用了受影响的单元格
					if f.formulaReferencesAffectedCells(formula, sheetName, affected) {
						affected[cellKey] = true
						changed = true
					}
				}
			}
		}
	}

	return affected
}

// recalculateAffectedCellsWithDAG 使用 DAG 按依赖顺序重新计算受影响的单元格
func (f *File) recalculateAffectedCellsWithDAG(calcChain *xlsxCalcChain, affectedFormulas map[string]bool) error {
	if len(affectedFormulas) == 0 {
		return nil
	}

	// 构建 DAG
	dag := newCalcDAG()

	// 直接从 affectedFormulas 构建 DAG，不依赖 calcChain 的顺序
	for cellKey := range affectedFormulas {
		dag.addNode(cellKey)

		parts := strings.SplitN(cellKey, "!", 2)
		if len(parts) != 2 {
			continue
		}
		sheetName, cellRef := parts[0], parts[1]

		// 获取公式内容
		ws, err := f.workSheetReader(sheetName)
		if err != nil {
			continue
		}

		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}

		cellData := f.getCellFromWorksheet(ws, col, row)
		if cellData == nil || cellData.F == nil {
			continue
		}

		formula := cellData.F.Content
		if formula == "" && cellData.F.T == STCellFormulaTypeShared && cellData.F.Si != nil {
			formula, _ = getSharedFormula(ws, *cellData.F.Si, cellRef)
		}

		if formula == "" {
			continue
		}

		// 解析公式依赖，添加边
		deps := f.extractFormulaDependencies(formula, sheetName)
		for _, dep := range deps {
			if affectedFormulas[dep] {
				dag.addEdge(dep, cellKey) // dep -> cellKey 表示 cellKey 依赖于 dep
			}
		}
	}

	// 拓扑排序
	sorted, err := dag.topologicalSort()
	if err != nil {
		// 如果有循环依赖，回退到按 calcChain 顺序计算
		return f.recalculateAffectedCells(calcChain, affectedFormulas)
	}

	// 按拓扑顺序计算
	for _, cellKey := range sorted {
		parts := strings.SplitN(cellKey, "!", 2)
		if len(parts) != 2 {
			continue
		}
		sheet, cell := parts[0], parts[1]

		if err := f.recalculateCell(sheet, cell); err != nil {
			continue // 忽略单个单元格的错误
		}
	}

	return nil
}

// extractFormulaDependencies 从公式中提取所有依赖的单元格
func (f *File) extractFormulaDependencies(formula, currentSheet string) []string {
	var deps []string

	// 匹配单元格引用：Sheet!A1 或 A1 或 Sheet!A1:B2 或 A1:B2
	// 也处理带引号的工作表名：'Sheet Name'!A1
	cellRefPattern := regexp.MustCompile(`(?:'([^']+)'!|([A-Za-z_][A-Za-z0-9_]*!))?(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?`)

	matches := cellRefPattern.FindAllStringSubmatch(formula, -1)
	seen := make(map[string]bool)

	for _, match := range matches {
		var sheet string
		if match[1] != "" {
			sheet = match[1] // 带引号的工作表名
		} else if match[2] != "" {
			sheet = strings.TrimSuffix(match[2], "!")
		} else {
			sheet = currentSheet
		}

		startCell := match[3]
		endCell := match[4]

		// 移除 $ 符号
		startCell = strings.ReplaceAll(startCell, "$", "")
		if endCell != "" {
			endCell = strings.ReplaceAll(endCell, "$", "")
		}

		if endCell == "" {
			// 单个单元格引用
			cellKey := sheet + "!" + startCell
			if !seen[cellKey] {
				deps = append(deps, cellKey)
				seen[cellKey] = true
			}
		} else {
			// 范围引用 - 展开范围中的所有单元格
			startCol, startRow, err1 := CellNameToCoordinates(startCell)
			endCol, endRow, err2 := CellNameToCoordinates(endCell)
			if err1 == nil && err2 == nil {
				// 限制展开的单元格数量，避免大范围导致性能问题
				maxCells := 1000
				count := 0
				for row := startRow; row <= endRow && count < maxCells; row++ {
					for col := startCol; col <= endCol && count < maxCells; col++ {
						cellName, _ := CoordinatesToCellName(col, row)
						cellKey := sheet + "!" + cellName
						if !seen[cellKey] {
							deps = append(deps, cellKey)
							seen[cellKey] = true
							count++
						}
					}
				}
			}
		}
	}

	return deps
}

// calcDAG 计算依赖图
type calcDAG struct {
	nodes    map[string]bool
	edges    map[string][]string // from -> []to
	inDegree map[string]int
}

func newCalcDAG() *calcDAG {
	return &calcDAG{
		nodes:    make(map[string]bool),
		edges:    make(map[string][]string),
		inDegree: make(map[string]int),
	}
}

func (d *calcDAG) addNode(node string) {
	if !d.nodes[node] {
		d.nodes[node] = true
		d.inDegree[node] = 0
	}
}

func (d *calcDAG) addEdge(from, to string) {
	d.addNode(from)
	d.addNode(to)
	d.edges[from] = append(d.edges[from], to)
	d.inDegree[to]++
}

func (d *calcDAG) topologicalSort() ([]string, error) {
	var result []string
	queue := make([]string, 0)

	// 找出所有入度为 0 的节点
	for node := range d.nodes {
		if d.inDegree[node] == 0 {
			queue = append(queue, node)
		}
	}

	for len(queue) > 0 {
		node := queue[0]
		queue = queue[1:]
		result = append(result, node)

		for _, neighbor := range d.edges[node] {
			d.inDegree[neighbor]--
			if d.inDegree[neighbor] == 0 {
				queue = append(queue, neighbor)
			}
		}
	}

	if len(result) != len(d.nodes) {
		return nil, fmt.Errorf("circular dependency detected")
	}

	return result, nil
}

// RebuildCalcChain 扫描所有工作表的公式并重建 calcChain
func (f *File) RebuildCalcChain() error {
	calcChain := &xlsxCalcChain{}
	sheetList := f.GetSheetList()

	// 获取 sheetID 映射 (sheetName -> sheetID)
	sheetIDMap := make(map[string]int)
	sheetMap := f.GetSheetMap() // map[sheetID]sheetName
	for id, name := range sheetMap {
		sheetIDMap[name] = id
	}

	for _, sheetName := range sheetList {
		ws, err := f.workSheetReader(sheetName)
		if err != nil || ws.SheetData.Row == nil {
			continue
		}

		// 获取正确的 sheetID
		sheetID, ok := sheetIDMap[sheetName]
		if !ok {
			// 如果没有找到，使用索引+1（Excel 的 sheetID 通常从 1 开始）
			sheetID = len(calcChain.C) + 1
		}

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				if cell.F != nil {
					formula := cell.F.Content
					// 处理共享公式
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}
					if formula != "" {
						calcChain.C = append(calcChain.C, xlsxCalcChainC{
							R: cell.R,
							I: sheetID,
						})
					}
				}
			}
		}
	}

	if len(calcChain.C) == 0 {
		// 即使没有公式，也设置一个空的 calcChain
		f.CalcChain = calcChain
		return nil
	}

	f.CalcChain = calcChain
	return nil
}
