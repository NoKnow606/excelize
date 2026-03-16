package excelize

import (
	"log"
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
	// Phase 1: 对于公式计算结果，应该根据返回值本身推断类型，而不是根据单元格类型
	// 因为公式单元格的 cellType 始终是 CellTypeUnset
	arg := inferFormulaResultType(value)

	// Phase 1: 存储 formulaArg 而不是字符串
	if worksheetCache != nil {
		worksheetCache.Set(sheet, cellName, arg)
	}

	// 保持与旧缓存的兼容性（暂时）
	cacheKey := sheet + "!" + cellName
	f.calcCache.Store(cacheKey, arg)
	f.calcCache.Store(cacheKey+"!raw=true", value)

	f.setFormulaValue(sheet, cellName, value)
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
	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	f.mu.Unlock()
	if err != nil {
		log.Printf("  ⚠️  [setFormulaValue] workSheetReader failed for %s!%s: %v", sheet, cellName, err)
		return
	}

	ws.mu.Lock()
	c, _, _, err := ws.prepareCell(cellName)
	if err != nil {
		ws.mu.Unlock()
		log.Printf("  ⚠️  [setFormulaValue] prepareCell failed for %s!%s: %v", sheet, cellName, err)
		return
	}

	oldValue := c.V
	c.V = value
	c.T = inferXMLCellType(value)
	ws.mu.Unlock()

	if f.OnCellCalculated != nil && oldValue != value {
		f.OnCellCalculated(sheet, cellName, oldValue, value)
	}
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
