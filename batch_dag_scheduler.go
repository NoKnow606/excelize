package excelize

import (
	"log"
	"sort"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"
	"time"
)

// DAGScheduler implements a dynamic dependency-aware scheduler
// that executes formulas as soon as their dependencies are satisfied
type DAGScheduler struct {
	f               *File
	graph           *dependencyGraph
	readyQueue      chan string         // 准备好执行的公式队列
	completedCount  atomic.Int64        // 已完成的公式数量
	inFlightCount   atomic.Int64        // 正在执行的公式数量
	results         sync.Map            // 结果缓存 map[string]string
	dependencyCount map[string]int      // 每个公式还有多少依赖未完成
	dependents      map[string][]string // 反向依赖：哪些公式依赖这个公式
	mu              sync.Mutex          // 保护 dependencyCount 的锁
	totalFormulas   int
	numWorkers      int
	queueClosed     atomic.Bool         // 标记队列是否已关闭
	subExprCache    *SubExpressionCache // 子表达式缓存（用于复合公式）
	worksheetCache  *WorksheetCache     // 统一的worksheet缓存（用于存储所有计算结果）
}

// NewDAGScheduler creates a new DAG scheduler
func (f *File) NewDAGScheduler(graph *dependencyGraph, numWorkers int, subExprCache *SubExpressionCache) *DAGScheduler {
	// 统计总公式数和 Level 0 公式数
	totalFormulas := 0
	level0Count := 0
	for _, cells := range graph.levels {
		totalFormulas += len(cells)
	}
	if len(graph.levels) > 0 {
		level0Count = len(graph.levels[0])
	}

	// readyQueue 的缓冲区要足够大，至少能容纳所有 Level 0 的公式
	// 加上一些余量以应对后续的依赖完成通知
	queueSize := level0Count + 1000000
	if queueSize < 10000 {
		queueSize = 10000
	}

	scheduler := &DAGScheduler{
		f:               f,
		graph:           graph,
		readyQueue:      make(chan string, queueSize),
		dependencyCount: make(map[string]int),
		dependents:      make(map[string][]string),
		numWorkers:      numWorkers,
		totalFormulas:   totalFormulas,
		subExprCache:    subExprCache,
	}

	// 构建依赖计数和反向依赖关系
	for cell, node := range graph.nodes {
		// 统计有多少formula依赖（不计算data cell）
		formulaDeps := 0
		for _, dep := range node.dependencies {
			if _, isFormula := graph.nodes[dep]; isFormula {
				formulaDeps++
				// 构建反向依赖：dep -> cell
				scheduler.dependents[dep] = append(scheduler.dependents[dep], cell)
			}
		}
		scheduler.dependencyCount[cell] = formulaDeps

		// 如果没有依赖，直接加入ready queue
		if formulaDeps == 0 {
			scheduler.readyQueue <- cell
		}
	}

	return scheduler
}

// NewDAGSchedulerForLevel creates a DAG scheduler for a specific level
// Only formulas within the level are scheduled (dependencies from previous levels are already completed)
// Returns nil,false if level contains circular dependencies (no ready nodes)
func (f *File) NewDAGSchedulerForLevel(graph *dependencyGraph, levelIdx int, levelCells []string, numWorkers int, subExprCache *SubExpressionCache, worksheetCache *WorksheetCache) (*DAGScheduler, bool) {
	// 创建当前层的公式集合
	levelCellsMap := make(map[string]bool)
	for _, cell := range levelCells {
		levelCellsMap[cell] = true
	}

	// readyQueue 缓冲区要足够大，至少能容纳当前层所有可能同时准备好的公式
	queueSize := len(levelCells) + 10000
	if queueSize < 10000 {
		queueSize = 10000
	}

	scheduler := &DAGScheduler{
		f:               f,
		graph:           graph,
		readyQueue:      make(chan string, queueSize),
		dependencyCount: make(map[string]int),
		dependents:      make(map[string][]string),
		numWorkers:      numWorkers,
		totalFormulas:   len(levelCells),
		subExprCache:    subExprCache,
		worksheetCache:  worksheetCache,
	}

	readyCount := 0

	// 构建当前层内部的依赖关系
	// 只考虑当前层内部的依赖（层与层之间的依赖已经满足）
	for _, cell := range levelCells {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}

		// 统计当前层内部的依赖数量
		levelInternalDeps := 0
		for _, dep := range node.dependencies {
			// 只统计同层内部的依赖
			if levelCellsMap[dep] {
				levelInternalDeps++
				// 构建反向依赖：dep -> cell（只在当前层内部）
				scheduler.dependents[dep] = append(scheduler.dependents[dep], cell)
			}
		}
		scheduler.dependencyCount[cell] = levelInternalDeps

		// 如果没有层内依赖，直接加入ready queue
		if levelInternalDeps == 0 {
			scheduler.readyQueue <- cell
			readyCount++
		}
	}

	if len(levelCells) > 0 && readyCount == 0 {
		return nil, false
	}

	return scheduler, true
}

// Run executes the DAG scheduler
func (scheduler *DAGScheduler) Run() {
	startTime := time.Now()
	log.Printf("🚀 [DAG Scheduler] Starting: %d formulas with %d workers", scheduler.totalFormulas, scheduler.numWorkers)

	// 统计初始 ready queue 中有多少公式
	initialReady := len(scheduler.readyQueue)
	log.Printf("  📊 [DAG Scheduler] Initial ready queue size: %d formulas (no dependencies)", initialReady)

	// 边界情况：空图直接返回
	if scheduler.totalFormulas == 0 {
		log.Printf("✅ [DAG Scheduler] No formulas to calculate, exiting immediately")
		return
	}

	// 检查是否有依赖问题（如果没有任何公式准备好）
	if initialReady == 0 && scheduler.totalFormulas > 0 {
		log.Printf("⚠️ [DAG Scheduler] WARNING: No formulas ready! Possible circular dependency or dependency issue")
		// 打印一些有依赖的公式示例
		count := 0
		for cell, depCount := range scheduler.dependencyCount {
			if depCount > 0 && count < 5 {
				if node, exists := scheduler.graph.nodes[cell]; exists {
					log.Printf("    Example blocked formula: %s (waiting for %d deps) = %s", cell, depCount, node.formula[:min(100, len(node.formula))])
					log.Printf("      Dependencies: %v", node.dependencies[:min(5, len(node.dependencies))])
				}
				count++
			}
		}
		return
	}

	// 确保在函数退出时关闭队列，防止 goroutine 泄漏
	defer scheduler.closeReadyQueue()

	var wg sync.WaitGroup

	// 启动worker pool
	for i := 0; i < scheduler.numWorkers; i++ {
		wg.Add(1)
		go scheduler.worker(&wg, i)
	}

	// 启动进度报告和死锁检测 goroutine
	done := make(chan struct{})
	go func() {
		ticker := time.NewTicker(5 * time.Second) // 每5秒报告一次进度
		defer ticker.Stop()
		lastCompleted := int64(0)
		stallCount := 0
		for {
			select {
			case <-done:
				return
			case <-ticker.C:
				currentCompleted := scheduler.completedCount.Load()
				inFlight := scheduler.inFlightCount.Load()
				queueLen := len(scheduler.readyQueue)
				elapsed := time.Since(startTime)
				rate := float64(currentCompleted) / elapsed.Seconds()

				log.Printf("  📊 [Progress] %d/%d (%.1f%%) completed, %d in-flight, %d queued, %.1f/sec",
					currentCompleted, scheduler.totalFormulas,
					float64(currentCompleted)*100/float64(scheduler.totalFormulas),
					inFlight, queueLen, rate)

				// 检查是否停滞
				if currentCompleted == lastCompleted && inFlight == 0 && currentCompleted < int64(scheduler.totalFormulas) {
					stallCount++
					log.Printf("  ⚠️ [Progress] Stall detected: no progress for %d checks", stallCount)

					if stallCount >= 6 { // 30秒后强制关闭
						log.Printf("⚠️ [DAG Scheduler] Forcing close after stall")
						scheduler.closeReadyQueue()
						return
					}
				} else {
					stallCount = 0
				}
				lastCompleted = currentCompleted
			}
		}
	}()

	// 等待所有worker完成
	wg.Wait()
	close(done)

	duration := time.Since(startTime)
	if scheduler.totalFormulas > 0 {
		log.Printf("✅ [DAG Scheduler] Completed %d formulas in %v (avg: %v/formula)",
			scheduler.totalFormulas, duration, duration/time.Duration(scheduler.totalFormulas))
	} else {
		log.Printf("✅ [DAG Scheduler] Completed in %v", duration)
	}
}

// worker processes formulas from the ready queue
func (scheduler *DAGScheduler) worker(wg *sync.WaitGroup, workerID int) {
	defer wg.Done()

	for cell := range scheduler.readyQueue {
		scheduler.executeFormula(cell)
	}
}

// executeFormula calculates a single formula and notifies dependents
func (scheduler *DAGScheduler) executeFormula(cell string) {
	scheduler.inFlightCount.Add(1)
	defer scheduler.inFlightCount.Add(-1)

	// Parse cell reference
	parts := strings.Split(cell, "!")
	if len(parts) != 2 {
		scheduler.notifyDependents(cell)
		scheduler.markFormulaDone()
		return
	}

	sheet := parts[0]
	cellName := parts[1]

	// 优化：先检查 worksheetCache 是否已有批量预计算的结果
	if scheduler.worksheetCache != nil {
		if cachedArg, found := scheduler.worksheetCache.Get(sheet, cellName); found {
			value := cachedArg.Value()
			scheduler.results.Store(cell, value)
			scheduler.f.setFormulaValue(sheet, cellName, value)
			scheduler.notifyDependents(cell)
			scheduler.markFormulaDone()
			return
		}
	}

	// 也检查 calcCache（兼容旧的缓存路径）
	cacheKey := cell + "!raw=true"
	if cached, ok := scheduler.f.calcCache.Load(cacheKey); ok {
		if value, isStr := cached.(string); isStr {
			scheduler.results.Store(cell, value)
			scheduler.f.setFormulaValue(sheet, cellName, value)
			scheduler.notifyDependents(cell)
			scheduler.markFormulaDone()
			return
		}
	}

	// 获取公式（从 graph 中，避免重复读取）
	formula := ""
	if node, exists := scheduler.graph.nodes[cell]; exists {
		formula = node.formula
	}

	// 使用带子表达式缓存的计算
	opts := Options{RawCellValue: true, MaxCalcIterations: 100}

	value, err := scheduler.f.CalcCellValueWithSubExprCache(sheet, cellName, formula, scheduler.subExprCache, scheduler.worksheetCache, opts)

	// CRITICAL: Even if err != nil, value may contain error string like "#DIV/0!"
	// We should still store and write back error values so they display in Excel
	if err != nil && value == "" {
		// True error case: calculation failed without producing error value
		scheduler.notifyDependents(cell)
		scheduler.markFormulaDone()
		return
	}

	// 保存结果 (包括错误值)
	scheduler.results.Store(cell, value)

	// 写回缓存和 worksheet (包括错误值)
	scheduler.f.storeCalculatedValue(sheet, cellName, value, scheduler.worksheetCache)

	// 通知依赖此公式的其他公式
	scheduler.notifyDependents(cell)

	// 标记完成
	scheduler.markFormulaDone()
}

// notifyDependents decrements dependency count for dependents and enqueues ready formulas
func (scheduler *DAGScheduler) notifyDependents(completedCell string) {
	dependents, exists := scheduler.dependents[completedCell]
	if !exists || len(dependents) == 0 {
		return
	}

	scheduler.mu.Lock()
	defer scheduler.mu.Unlock()

	for _, dependent := range dependents {
		scheduler.dependencyCount[dependent]--
		if scheduler.dependencyCount[dependent] == 0 {
			// 所有依赖都完成了，可以执行
			select {
			case scheduler.readyQueue <- dependent:
			default:
				// Queue full, this shouldn't happen with large buffer
				log.Printf("⚠️ [DAG Scheduler] Ready queue full, dropping %s", dependent)
			}
		}
	}
}

// writeBackToWorksheet writes calculated value back to worksheet
// GetResults returns all calculated results
func (scheduler *DAGScheduler) GetResults() map[string]string {
	results := make(map[string]string)
	scheduler.results.Range(func(key, value interface{}) bool {
		results[key.(string)] = value.(string)
		return true
	})
	return results
}

func (scheduler *DAGScheduler) markFormulaDone() {
	newCount := scheduler.completedCount.Add(1)
	if newCount == int64(scheduler.totalFormulas) {
		scheduler.closeReadyQueue()
	}
}

func (scheduler *DAGScheduler) closeReadyQueue() {
	if scheduler.queueClosed.CompareAndSwap(false, true) {
		close(scheduler.readyQueue)
	}
}

// storeCalculatedValue persists the computed formula result to caches and worksheet
// Phase 1: 改为接收 formulaArg 并存储类型信息
func (f *File) storeCalculatedValue(sheet, cellName, value string, worksheetCache *WorksheetCache) {
	f.persistFormulaResult(sheet, cellName, value, worksheetCache, true, true)
}

// inferFormulaResultType 根据公式返回值推断类型
// 与 inferCellValueType 不同，这个函数专门用于处理公式计算结果
func inferFormulaResultType(value string) formulaArg {
	// 空字符串：返回字符串类型（很多公式用 "" 表示空值）
	if value == "" {
		return newStringFormulaArg("")
	}

	// 只将真实的 Excel 错误字面量识别为错误值。
	// 普通文本也可能以 # 开头，例如 Markdown 标题。
	if isExcelErrorLiteral(value) {
		return newErrorFormulaArg(value, value)
	}

	// 尝试解析为数字
	if num, err := strconv.ParseFloat(value, 64); err == nil {
		return newNumberFormulaArg(num)
	}

	// 检查布尔值
	upper := strings.ToUpper(value)
	if upper == "TRUE" || upper == "FALSE" {
		return newBoolFormulaArg(upper == "TRUE")
	}

	// 其他情况：字符串
	return newStringFormulaArg(value)
}

func isExcelErrorLiteral(value string) bool {
	switch value {
	case formulaErrorDIV,
		formulaErrorNAME,
		formulaErrorNA,
		formulaErrorNUM,
		formulaErrorVALUE,
		formulaErrorREF,
		formulaErrorNULL,
		formulaErrorSPILL,
		formulaErrorCALC,
		formulaErrorGETTINGDATA:
		return true
	default:
		return false
	}
}

func (f *File) setFormulaValue(sheet, cellName, value string) {
	f.persistFormulaResult(sheet, cellName, value, nil, false, true)
}

func (f *File) persistFormulaResult(sheet, cellName, value string, worksheetCache *WorksheetCache, updateCaches, notify bool) {
	if notify {
		formula, err := f.GetCellFormula(sheet, cellName)
		if err == nil && isExternalCachedFormula(formula) {
			notify = false
		}
	}
	if handled := f.persistSQLFormulaResult(sheet, cellName, value, worksheetCache, updateCaches, notify); handled {
		return
	}
	f.persistScalarFormulaResult(sheet, cellName, value, worksheetCache, updateCaches, notify)
}

func (f *File) persistScalarFormulaResult(sheet, cellName, value string, worksheetCache *WorksheetCache, updateCaches, notify bool) {
	arg := inferFormulaResultType(value)
	if updateCaches {
		f.storeFormulaResultCache(sheet, cellName, value, arg, worksheetCache)
	}

	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	f.mu.Unlock()
	if err != nil {
		log.Printf("  ⚠️  [persistScalarFormulaResult] workSheetReader failed for %s!%s: %v", sheet, cellName, err)
		return
	}

	ws.mu.Lock()
	c, err := prepareWorksheetCell(ws, cellName)
	if err != nil {
		ws.mu.Unlock()
		log.Printf("  ⚠️  [persistScalarFormulaResult] prepareWorksheetCell failed for %s!%s: %v", sheet, cellName, err)
		return
	}

	oldValue := c.V
	c.V = value
	c.T = inferXMLCellType(value)
	ws.mu.Unlock()

	if notify && f.OnCellCalculated != nil && oldValue != value {
		f.OnCellCalculated(sheet, cellName, oldValue, value)
	}
}

func (f *File) persistSQLFormulaResult(sheet, cellName, fallbackValue string, worksheetCache *WorksheetCache, updateCaches, notify bool) bool {
	formula, err := f.GetCellFormula(sheet, cellName)
	if err != nil || !IsSQLFormula(formula) {
		return false
	}

	result, err := f.CalcCellValueWithMatrix(sheet, cellName, Options{RawCellValue: true})
	if err != nil {
		// Keep the original behavior: still attempt to write the spill range below
		// using whatever Matrix the calc engine returned (may be empty); the caller
		// has already received the error via the matrix result. Errors from
		// CalcCellValueWithMatrix are intentionally swallowed here so persistence
		// can fall back to the scalar fallbackValue when the matrix is empty.
		_ = err
	}
	return f.applySQLFormulaMatrix(sheet, cellName, fallbackValue, result, worksheetCache, updateCaches, notify)
}

// persistSQLFormulaResultWithMatrix writes a precomputed SQL spill matrix into
// the workbook without re-executing the SQL query. This is the recommended hot
// path for callers that already produced a matrix via CalcCellValueWithMatrix
// (e.g. CalcCellValues), because re-running the same SQL multiple times
// allocates a new in-memory SQLite database and re-materializes every source
// worksheet on each invocation.
func (f *File) persistSQLFormulaResultWithMatrix(sheet, cellName, fallbackValue string, result CalcCellValueWithMatrixResult, worksheetCache *WorksheetCache, updateCaches, notify bool) bool {
	formula, err := f.GetCellFormula(sheet, cellName)
	if err != nil || !isExternalSpillFormula(formula) {
		return false
	}
	return f.applySQLFormulaMatrix(sheet, cellName, fallbackValue, result, worksheetCache, updateCaches, notify)
}

func isExternalSpillFormula(formula string) bool {
	if IsSQLFormula(formula) {
		return true
	}
	trimmed := strings.TrimSpace(formula)
	trimmed = strings.TrimPrefix(trimmed, "=")
	return strings.HasPrefix(strings.ToUpper(trimmed), "MAYBE_PIVOT(")
}

// PersistSQLFormulaResultWithMatrix writes a precomputed SQL spill matrix into
// the workbook for the given anchor cell, mirroring what UpdateSheetFormulaCache
// does for SQL formulas but without re-executing the SQL query.
//
// Callers that have already executed the query (e.g. via ExecuteSQL or
// CalcCellValueWithMatrix) can pass the resulting matrix directly to this
// method to persist the spill range and the cached top-left value. This avoids
// a second materialization of every source worksheet into a fresh in-memory
// SQLite database, which is otherwise the dominant memory cost of saving a
// workbook that contains SQL formulas.
//
// The cell at cellName must already contain the SQL formula (typically set
// via SetCellFormula) before calling this method. fallbackValue is written
// when the matrix is empty so that the cell still receives a sensible value.
//
// Returns true if the cell holds a SQL formula and the spill range was
// applied (or the fallback was written); returns false if the cell does not
// contain a SQL formula, in which case the caller should fall back to the
// regular non-SQL persistence path.
func (f *File) PersistSQLFormulaResultWithMatrix(sheet, cellName, fallbackValue string, result CalcCellValueWithMatrixResult) bool {
	return f.persistSQLFormulaResultWithMatrix(sheet, cellName, fallbackValue, result, nil, false, false)
}

// applySQLFormulaMatrix writes the spill range produced by a SQL formula into
// the worksheet. It does not execute SQL; the caller must supply the matrix.
func (f *File) applySQLFormulaMatrix(sheet, cellName, fallbackValue string, result CalcCellValueWithMatrixResult, worksheetCache *WorksheetCache, updateCaches, notify bool) bool {
	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	f.mu.Unlock()
	if err != nil {
		log.Printf("  ⚠️  [persistSQLFormulaResult] workSheetReader failed for %s!%s: %v", sheet, cellName, err)
		return true
	}

	ws.mu.Lock()
	c, err := prepareWorksheetCell(ws, cellName)
	if err != nil {
		ws.mu.Unlock()
		log.Printf("  ⚠️  [persistSQLFormulaResult] prepareWorksheetCell failed for %s!%s: %v", sheet, cellName, err)
		return true
	}

	oldValue := c.V
	oldRef := ""
	if c.F != nil {
		oldRef = c.F.Ref
	}

	if len(result.Matrix) == 0 || len(result.Matrix[0]) == 0 {
		if oldRef != "" {
			clearWorksheetRangeValues(ws, oldRef, cellName)
		}
		if c.F != nil {
			c.F.Ref = ""
		}
		c.V = fallbackValue
		c.T = inferXMLCellType(fallbackValue)
		ws.mu.Unlock()

		if oldRef != "" {
			clearSpillRangeCache(f, worksheetCache, sheet, oldRef, cellName)
		}
		if updateCaches {
			f.storeFormulaResultCache(sheet, cellName, fallbackValue, inferFormulaResultType(fallbackValue), worksheetCache)
		}
		if notify && f.OnCellCalculated != nil && oldValue != fallbackValue {
			f.OnCellCalculated(sheet, cellName, oldValue, fallbackValue)
		}
		return true
	}

	startCol, startRow, err := CellNameToCoordinates(cellName)
	if err != nil {
		ws.mu.Unlock()
		return false
	}

	endCol := startCol + len(result.Matrix[0]) - 1
	endRow := startRow + len(result.Matrix) - 1
	ref, err := CoordinatesToCellName(endCol, endRow)
	if err != nil {
		ws.mu.Unlock()
		return false
	}
	newRef := cellName + ":" + ref
	clearRef := buildSQLResultClearRef(ws, cellName, oldRef, newRef)

	type spillValue struct {
		cell  string
		value string
		arg   formulaArg
	}
	values := make([]spillValue, 0, len(result.Matrix)*len(result.Matrix[0]))
	topLeftValue := fallbackValue
	for r := range result.Matrix {
		for c := range result.Matrix[r] {
			cellRef, err := CoordinatesToCellName(startCol+c, startRow+r)
			if err != nil {
				ws.mu.Unlock()
				return false
			}
			arg := interfaceToFormulaArg(result.Matrix[r][c])
			value := arg.Value()
			if r == 0 && c == 0 {
				topLeftValue = value
			}
			values = append(values, spillValue{
				cell:  cellRef,
				value: value,
				arg:   arg,
			})
		}
	}

	if clearRef != "" {
		clearWorksheetRangeValues(ws, clearRef, cellName)
	}
	if c.F != nil {
		c.F.Ref = newRef
	}
	for _, item := range values {
		target, err := prepareWorksheetCell(ws, item.cell)
		if err != nil {
			continue
		}
		if item.cell != cellName {
			target.F = nil
		}
		target.V = item.value
		target.T = inferXMLCellType(item.value)
	}
	refreshWorksheetDimension(ws)
	ws.mu.Unlock()

	if clearRef != "" {
		clearSpillRangeCache(f, worksheetCache, sheet, clearRef, cellName)
	}
	if updateCaches {
		for _, item := range values {
			f.storeFormulaResultCache(sheet, item.cell, item.value, item.arg, worksheetCache)
		}
	}

	if notify && f.OnCellCalculated != nil && oldValue != topLeftValue {
		f.OnCellCalculated(sheet, cellName, oldValue, topLeftValue)
	}
	return true
}

func clearWorksheetRangeValues(ws *xlsxWorksheet, ref, anchor string) {
	coordinates, err := rangeRefToCoordinates(ref)
	if err != nil {
		return
	}
	_ = sortCoordinates(coordinates)

	for rowIdx := range ws.SheetData.Row {
		rowData := &ws.SheetData.Row[rowIdx]
		if rowData.R < coordinates[1] || rowData.R > coordinates[3] {
			continue
		}
		for cellIdx := range rowData.C {
			c := &rowData.C[cellIdx]
			col, row, err := CellNameToCoordinates(c.R)
			if err != nil {
				continue
			}
			if row < coordinates[1] || row > coordinates[3] || col < coordinates[0] || col > coordinates[2] {
				continue
			}
			if c.R == anchor {
				continue
			}
			c.F = nil
			c.V = ""
			c.T = ""
			c.IS = nil
		}
	}
	trimWorksheetContiguousEmptyTail(ws)
	refreshWorksheetDimension(ws)
}

func trimWorksheetContiguousEmptyTail(ws *xlsxWorksheet) {
	if ws == nil {
		return
	}

	for rowIdx := range ws.SheetData.Row {
		rowData := &ws.SheetData.Row[rowIdx]
		lastUsed := -1
		for cellIdx := range rowData.C {
			if cellContributesToUsedRange(rowData.C[cellIdx]) {
				lastUsed = cellIdx
			}
		}
		if lastUsed == -1 {
			rowData.C = rowData.C[:0]
			continue
		}
		rowData.C = rowData.C[:lastUsed+1]
	}

	lastUsedRow := len(ws.SheetData.Row) - 1
	for lastUsedRow >= 0 {
		row := &ws.SheetData.Row[lastUsedRow]
		if len(row.C) != 0 || rowHasMeaningfulAttrs(ws, row) {
			break
		}
		lastUsedRow--
	}
	ws.SheetData.Row = ws.SheetData.Row[:lastUsedRow+1]
}

func rowHasMeaningfulAttrs(ws *xlsxWorksheet, row *xlsxRow) bool {
	if row == nil {
		return false
	}
	if row.Spans != "" || row.S != 0 || row.CustomFormat || row.Hidden ||
		row.OutlineLevel != 0 || row.Collapsed || row.ThickTop || row.ThickBot || row.Ph {
		return true
	}
	if row.Ht != nil || row.CustomHeight {
		return !rowUsesDefaultSheetHeight(ws, row)
	}
	return false
}

func rowUsesDefaultSheetHeight(ws *xlsxWorksheet, row *xlsxRow) bool {
	if ws == nil || row == nil || ws.SheetFormatPr == nil || !ws.SheetFormatPr.CustomHeight {
		return false
	}
	if !row.CustomHeight || row.Ht == nil {
		return false
	}
	return *row.Ht == ws.SheetFormatPr.DefaultRowHeight
}

func refreshWorksheetDimension(ws *xlsxWorksheet) {
	maxRow, maxCol := 0, 0
	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			if !cellContributesToUsedRange(cell) {
				continue
			}
			colNum, rowNum, err := CellNameToCoordinates(cell.R)
			if err != nil {
				continue
			}
			if rowNum > maxRow {
				maxRow = rowNum
			}
			if colNum > maxCol {
				maxCol = colNum
			}
		}
	}

	if maxRow == 0 || maxCol == 0 {
		ws.Dimension = &xlsxDimension{Ref: "A1"}
		return
	}

	endCell, err := CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return
	}
	if maxRow == 1 && maxCol == 1 {
		ws.Dimension = &xlsxDimension{Ref: "A1"}
		return
	}
	ws.Dimension = &xlsxDimension{Ref: "A1:" + endCell}
}

func cellContributesToUsedRange(cell xlsxC) bool {
	return cell.F != nil || cell.V != "" || cell.T != "" || cell.IS != nil
}

func clearSpillRangeCache(f *File, worksheetCache *WorksheetCache, sheet, ref, anchor string) {
	coordinates, err := rangeRefToCoordinates(ref)
	if err != nil {
		return
	}
	_ = sortCoordinates(coordinates)
	for col := coordinates[0]; col <= coordinates[2]; col++ {
		for row := coordinates[1]; row <= coordinates[3]; row++ {
			cellRef, err := CoordinatesToCellName(col, row)
			if err != nil || cellRef == anchor {
				continue
			}
			cacheKey := sheet + "!" + cellRef
			f.calcCache.Delete(cacheKey)
			f.calcCache.Delete(cacheKey + "!raw=false")
			f.calcCache.Delete(cacheKey + "!raw=true")
			if worksheetCache != nil {
				worksheetCache.Delete(sheet, cellRef)
			}
		}
	}
}

func buildSQLResultClearRef(ws *xlsxWorksheet, anchor, oldRef, newRef string) string {
	startCol, startRow, err := CellNameToCoordinates(anchor)
	if err != nil {
		return oldRef
	}

	maxCol, maxRow := startCol, startRow
	extendMax := func(ref string) {
		if ref == "" {
			return
		}
		coordinates, err := rangeRefToCoordinates(ref)
		if err != nil {
			return
		}
		_ = sortCoordinates(coordinates)
		if coordinates[2] > maxCol {
			maxCol = coordinates[2]
		}
		if coordinates[3] > maxRow {
			maxRow = coordinates[3]
		}
	}

	extendMax(oldRef)
	extendMax(newRef)
	if ws != nil && ws.Dimension != nil {
		extendMax(ws.Dimension.Ref)
	}

	endCell, err := CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return oldRef
	}
	if anchor == endCell {
		return anchor
	}
	return anchor + ":" + endCell
}

func prepareWorksheetCell(ws *xlsxWorksheet, cell string) (*xlsxC, error) {
	var err error
	cell, err = ws.mergeCellsParser(cell)
	if err != nil {
		return nil, err
	}

	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return nil, err
	}
	if col < 1 || row < 1 {
		return nil, newCellNameToCoordinatesError(cell, newInvalidCellNameError(cell))
	}

	sizeHint := 0
	if rowCount := len(ws.SheetData.Row); rowCount > 0 {
		sizeHint = len(ws.SheetData.Row[rowCount-1].C)
	}

	var ht *float64
	var customHeight bool
	if ws.SheetFormatPr != nil && ws.SheetFormatPr.CustomHeight {
		ht = float64Ptr(ws.SheetFormatPr.DefaultRowHeight)
		customHeight = true
	}

	rowIdx := sort.Search(len(ws.SheetData.Row), func(i int) bool {
		return ws.SheetData.Row[i].R >= row
	})
	if rowIdx == len(ws.SheetData.Row) || ws.SheetData.Row[rowIdx].R != row {
		ws.SheetData.Row = append(ws.SheetData.Row, xlsxRow{})
		copy(ws.SheetData.Row[rowIdx+1:], ws.SheetData.Row[rowIdx:])
		ws.SheetData.Row[rowIdx] = xlsxRow{
			R:            row,
			CustomHeight: customHeight,
			Ht:           ht,
			C:            make([]xlsxC, 0, sizeHint),
		}
	}

	rowData := &ws.SheetData.Row[rowIdx]
	if col > len(rowData.C) {
		fillColumns(rowData, col, row)
	}
	return &rowData.C[col-1], nil
}

func (f *File) storeFormulaResultCache(sheet, cellName, value string, arg formulaArg, worksheetCache *WorksheetCache) {
	if worksheetCache != nil {
		worksheetCache.Set(sheet, cellName, arg)
	}
	cacheKey := sheet + "!" + cellName
	f.calcCache.Store(cacheKey, arg)
	f.calcCache.Store(cacheKey+"!raw=false", value)
	f.calcCache.Store(cacheKey+"!raw=true", value)
}

// inferXMLCellType 推断 XML 单元格类型（不是 formulaArg）
func inferXMLCellType(value string) string {
	if value == "" {
		return ""
	}
	if _, err := strconv.ParseFloat(value, 64); err == nil {
		return ""
	}
	upper := strings.ToUpper(value)
	if upper == "TRUE" || upper == "FALSE" {
		return "b"
	}
	if strings.HasPrefix(value, "#") {
		return "e"
	}
	return "str"
}
