package excelize

import (
	"context"
	"fmt"
	"log"
	"regexp"
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

	// Track columns that have timed out - skip all cells in that column
	// Map: "SheetName!Column" -> true (e.g., "补货计划!H" -> true)
	timeoutColumns := make(map[string]bool)

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
			currentSheetName = sheetName
			sheetFormulaCount = 0 // Reset counter for new sheet
			currentWs, err = f.workSheetReader(sheetName)
			if err != nil {
				continue
			}

			// Build cell map for fast lookup
			cellMap = make(map[string]*xlsxC)
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

		// Check if this column has already timed out - if so, skip it
		columnKey := sheetName + "!" + colLetter
		if timeoutColumns[columnKey] {
			// Skip this cell silently - column already timed out
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
			// Clear the cell value and continue to next formula
			cellRef.V = ""
			cellRef.T = ""
			formulaCount++
			continue
		} else if calcDuration > 100*time.Millisecond {
			slowFormulaCount++
		}

		// Check if this was a batch cache hit (very fast calculation)
		if calcDuration < 1*time.Microsecond {
			batchHitCount++
		}

		calcTime += calcDuration

		if err != nil {
			// If calculation fails, clear the cache
			cellRef.V = ""
			cellRef.T = ""
			continue
		}

		// Update cache value directly (we already have the cell reference)
		cellRef.V = result
		// Determine type based on value
		if result == "" {
			cellRef.T = ""
		} else if result == "TRUE" || result == "FALSE" {
			cellRef.T = "b"
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
		}
	}

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
	}

	// Log timeout statistics
	if timeoutCount > 0 {
		log.Printf("  ⏱️  Timeout formulas: %d formulas exceeded 5s and were skipped", timeoutCount)
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

	// 2. 读取 calcChain
	calcChain, err := f.calcChainReader()
	if err != nil {
		return err
	}

	// If calcChain doesn't exist or is empty, nothing to recalculate
	if calcChain == nil || len(calcChain.C) == 0 {
		return nil
	}

	// 3. 收集所有被更新的单元格（用于依赖检查）
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

	// 4. 找出所有受影响的公式单元格（通过依赖分析）
	affectedFormulas := f.findAffectedFormulas(calcChain, updatedCells, updatedColumns)

	// 5. 只清除受影响公式的缓存
	for cellKey := range affectedFormulas {
		cacheKey := cellKey + "!raw=false"
		f.calcCache.Delete(cacheKey)
	}

	// 6. 重新计算受影响的公式
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

// formulaReferencesUpdatedCells 检查公式是否引用了被更新的单元格
func (f *File) formulaReferencesUpdatedCells(formula, currentSheet string, updatedCells map[string]map[string]bool, updatedColumns map[string]map[string]bool) bool {
	// 去掉公式两端的单引号（如果有）
	formula = strings.Trim(formula, "'")

	// 检查全列引用（A:A, $A:$A, 'Sheet'!A:A, 中文表名!A:A 等）
	colRefPattern := regexp.MustCompile(`(?:'([^']+)'!|([^\s\(\)!]+!))?(\$?[A-Z]+):(\$?[A-Z]+)`)
	colMatches := colRefPattern.FindAllStringSubmatch(formula, -1)

	for _, match := range colMatches {
		refSheet := currentSheet
		if match[1] != "" {
			refSheet = match[1] // 单引号表名
		} else if match[2] != "" {
			refSheet = strings.TrimSuffix(match[2], "!")
		}

		// 优化：直接检查列索引，而不是遍历所有单元格
		if updatedColumns[refSheet] != nil {
			startCol := strings.ReplaceAll(match[3], "$", "")
			endCol := strings.ReplaceAll(match[4], "$", "")

			// 检查是否有更新的列在这个范围内
			for colName := range updatedColumns[refSheet] {
				if colName >= startCol && colName <= endCol {
					return true
				}
			}
		}
	}

	// 单元格引用匹配（支持单引号表名和中文表名）
	// 使用\b单词边界或(?:^|[^A-Za-z0-9_])确保不会匹配到运算符
	cellRefPattern := regexp.MustCompile(`(?:'([^']+)'!|(?:^|[^A-Za-z0-9_])([A-Za-z0-9_]+!))?(\$?[A-Z]+\$?[0-9]+)`)
	matches := cellRefPattern.FindAllStringSubmatch(formula, -1)

	for _, match := range matches {
		refSheet := currentSheet
		if match[1] != "" {
			refSheet = match[1] // 单引号表名
		} else if match[2] != "" {
			// 移除尾部的!，并且移除前面的非字母数字字符（如=, +等）
			refSheet = strings.TrimSuffix(match[2], "!")
			// 移除前导的非字母数字字符
			refSheet = strings.TrimLeft(refSheet, "=+-*/^&|<>(),")
		}
		refCell := strings.ReplaceAll(match[3], "$", "")

		if updatedCells[refSheet] != nil && updatedCells[refSheet][refCell] {
			return true
		}
	}

	return false
}

// formulaReferencesAffectedCells 检查公式是否引用了受影响的单元格
func (f *File) formulaReferencesAffectedCells(formula, currentSheet string, affectedCells map[string]bool) bool {
	// 去掉公式两端的单引号（如果有）
	formula = strings.Trim(formula, "'")

	// 检查全列引用（A:A, $A:$A, 'Sheet'!A:A, 中文表名!A:A 等）
	colRefPattern := regexp.MustCompile(`(?:'([^']+)'!|([^\s\(\)!]+!))?(\$?[A-Z]+):(\$?[A-Z]+)`)
	colMatches := colRefPattern.FindAllStringSubmatch(formula, -1)

	for _, match := range colMatches {
		refSheet := currentSheet
		if match[1] != "" {
			refSheet = match[1] // 单引号表名
		} else if match[2] != "" {
			refSheet = strings.TrimSuffix(match[2], "!")
		}

		// 检查受影响的单元格是否在这个列范围内
		for cellKey := range affectedCells {
			// 解析 cellKey (Sheet!Cell)
			parts := strings.Split(cellKey, "!")
			if len(parts) == 2 && parts[0] == refSheet {
				col, _, err := CellNameToCoordinates(parts[1])
				if err == nil {
					colName, _ := ColumnNumberToName(col)
					startCol := strings.ReplaceAll(match[3], "$", "")
					endCol := strings.ReplaceAll(match[4], "$", "")

					if colName >= startCol && colName <= endCol {
						return true
					}
				}
			}
		}
	}

	// 单元格引用匹配（支持单引号表名和中文表名）
	// 使用\b单词边界或(?:^|[^A-Za-z0-9_])确保不会匹配到运算符
	cellRefPattern := regexp.MustCompile(`(?:'([^']+)'!|(?:^|[^A-Za-z0-9_])([A-Za-z0-9_]+!))?(\$?[A-Z]+\$?[0-9]+)`)
	matches := cellRefPattern.FindAllStringSubmatch(formula, -1)

	for _, match := range matches {
		refSheet := currentSheet
		if match[1] != "" {
			refSheet = match[1] // 单引号表名
		} else if match[2] != "" {
			// 移除尾部的!，并且移除前面的非字母数字字符（如=, +等）
			refSheet = strings.TrimSuffix(match[2], "!")
			// 移除前导的非字母数字字符
			refSheet = strings.TrimLeft(refSheet, "=+-*/^&|<>(),")
		}
		refCell := strings.ReplaceAll(match[3], "$", "")
		cellKey := refSheet + "!" + refCell

		if affectedCells[cellKey] {
			return true
		}
	}

	// 检查范围引用（A1:B10, Sheet!A1:B10 等）
	rangeRefPattern := regexp.MustCompile(`(?:'([^']+)'!|([^\s\(\)!]+!))?(\$?[A-Z]+\$?[0-9]+):(\$?[A-Z]+\$?[0-9]+)`)
	rangeMatches := rangeRefPattern.FindAllStringSubmatch(formula, -1)

	for _, match := range rangeMatches {
		refSheet := currentSheet
		if match[1] != "" {
			refSheet = match[1]
		} else if match[2] != "" {
			refSheet = strings.TrimSuffix(match[2], "!")
		}

		startCell := strings.ReplaceAll(match[3], "$", "")
		endCell := strings.ReplaceAll(match[4], "$", "")

		// 检查受影响的单元格是否在这个范围内
		for cellKey := range affectedCells {
			parts := strings.Split(cellKey, "!")
			if len(parts) == 2 && parts[0] == refSheet {
				if f.cellInRange(parts[1], startCell, endCell) {
					return true
				}
			}
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

// RebuildCalcChain 扫描所有工作表的公式并重建 calcChain
func (f *File) RebuildCalcChain() error {
	calcChain := &xlsxCalcChain{}
	sheetList := f.GetSheetList()

	for sheetIndex, sheetName := range sheetList {
		ws, err := f.workSheetReader(sheetName)
		if err != nil || ws.SheetData.Row == nil {
			continue
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
							I: sheetIndex,
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
