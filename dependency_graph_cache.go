package excelize

import (
	"log"
	"strings"
)

func (f *File) markDependencyGraphDirty() {
	f.mu.Lock()
	f.dependencyGraphCache = nil
	f.dependencyGraphDirty = true
	f.mu.Unlock()
	f.removeEmbeddedCalculationSnapshot()
	f.clearPGSnapshotCacheEntries()
}

func (f *File) getCachedDependencyGraph() *dependencyGraph {
	f.mu.Lock()
	defer f.mu.Unlock()
	if f.dependencyGraphDirty || f.dependencyGraphCache == nil {
		return nil
	}
	return f.dependencyGraphCache
}

func (f *File) storeDependencyGraph(graph *dependencyGraph) *dependencyGraph {
	f.mu.Lock()
	f.dependencyGraphCache = graph
	f.dependencyGraphDirty = false
	f.mu.Unlock()
	return graph
}

func (f *File) loadOrBuildDependencyGraph(builder func() *dependencyGraph) *dependencyGraph {
	if graph := f.getCachedDependencyGraph(); graph != nil {
		log.Printf("  📦 [Dependency Analysis] Using cached dependency graph")
		return graph
	}
	return f.storeDependencyGraph(builder())
}

func (f *File) markCellValueDirty(sheet, cell string) {
	f.mu.Lock()
	f.dirtyValueCells[sheet+"!"+cell] = struct{}{}
	f.mu.Unlock()
}

func (f *File) snapshotDirtyValueCells() []string {
	f.mu.Lock()
	defer f.mu.Unlock()
	if len(f.dirtyValueCells) == 0 {
		return nil
	}
	cells := make([]string, 0, len(f.dirtyValueCells))
	for cell := range f.dirtyValueCells {
		cells = append(cells, cell)
	}
	return cells
}

func (f *File) clearDirtyValueCells() {
	f.mu.Lock()
	if len(f.dirtyValueCells) > 0 {
		f.dirtyValueCells = make(map[string]struct{})
	}
	f.mu.Unlock()
}

func (f *File) hasVolatileDependencies() bool {
	if f.VolatileDeps != nil {
		for _, volType := range f.VolatileDeps.VolType {
			for _, main := range volType.Main {
				for _, topic := range main.Tp {
					if len(topic.Tr) > 0 {
						return true
					}
				}
			}
		}
	}
	_, ok := f.Pkg.Load(defaultXMLPathVolatileDeps)
	return ok
}

func (f *File) clearWorksheetCacheAll() {
	f.mu.Lock()
	cache := f.worksheetCache
	f.mu.Unlock()
	if cache != nil {
		cache.Clear()
	}
}

func (f *File) clearPGLookupResultCache() {
	var keysToDelete []interface{}
	f.pgLookupResultCache.Range(func(key, value interface{}) bool {
		keysToDelete = append(keysToDelete, key)
		return true
	})
	for _, key := range keysToDelete {
		f.pgLookupResultCache.Delete(key)
	}
}

func (f *File) bumpSheetVersion(sheet string) {
	if sheet == "" {
		return
	}
	f.mu.Lock()
	if f.sheetVersion == nil {
		f.sheetVersion = make(map[string]uint64)
	}
	f.sheetVersion[sheet]++
	f.mu.Unlock()
}

func (f *File) bumpAllSheetVersions() {
	sheets := f.GetSheetList()
	f.mu.Lock()
	if f.sheetVersion == nil {
		f.sheetVersion = make(map[string]uint64)
	}
	for _, sheet := range sheets {
		f.sheetVersion[sheet]++
	}
	f.mu.Unlock()
}

func (f *File) renameSheetVersion(source, target string) {
	if source == "" || target == "" {
		return
	}
	f.mu.Lock()
	if f.sheetVersion == nil {
		f.sheetVersion = make(map[string]uint64)
	}
	version := f.sheetVersion[source]
	delete(f.sheetVersion, source)
	if existing := f.sheetVersion[target]; existing > version {
		version = existing
	}
	f.sheetVersion[target] = version + 1
	f.mu.Unlock()
}

func (f *File) deleteSheetVersion(sheet string) {
	if sheet == "" {
		return
	}
	f.mu.Lock()
	delete(f.sheetVersion, sheet)
	f.mu.Unlock()
}

func (f *File) getSheetVersion(sheet string) uint64 {
	f.mu.Lock()
	defer f.mu.Unlock()
	return f.sheetVersion[sheet]
}

func (f *File) clearWorksheetCacheSheet(sheet string) {
	f.mu.Lock()
	cache := f.worksheetCache
	f.mu.Unlock()
	if cache != nil {
		cache.ClearSheet(sheet)
	}
}

func (f *File) clearWorksheetCacheCell(sheet, cell string) {
	f.mu.Lock()
	cache := f.worksheetCache
	f.mu.Unlock()
	if cache != nil {
		cache.Delete(sheet, cell)
	}
}

func (f *File) setWorksheetCacheCell(sheet, cell string, value formulaArg) {
	f.mu.Lock()
	cache := f.worksheetCache
	f.mu.Unlock()
	if cache != nil {
		cache.Set(sheet, cell, value)
	}
}

func (f *File) prepareWorksheetCacheForGraph(graph *dependencyGraph) *WorksheetCache {
	f.mu.Lock()
	cache := f.worksheetCache
	if cache == nil {
		cache = NewWorksheetCache()
		f.worksheetCache = cache
	}
	f.mu.Unlock()

	for cell := range graph.nodes {
		parts := strings.SplitN(cell, "!", 2)
		if len(parts) != 2 {
			continue
		}
		cache.Delete(parts[0], parts[1])
	}

	return cache
}

func (g *dependencyGraph) buildAffectedSubgraph(dirtyCells []string) *dependencyGraph {
	if g == nil || len(g.nodes) == 0 || len(dirtyCells) == 0 {
		return nil
	}

	dependents := make(map[string][]string, len(g.nodes))
	for cell, node := range g.nodes {
		for _, dep := range node.dependencies {
			dependents[dep] = append(dependents[dep], cell)
		}
	}

	affected := make(map[string]struct{})
	visited := make(map[string]struct{}, len(dirtyCells))
	queue := make([]string, 0, len(dirtyCells))
	queue = append(queue, dirtyCells...)

	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]
		if _, ok := visited[current]; ok {
			continue
		}
		visited[current] = struct{}{}

		for _, dependent := range dependents[current] {
			if _, ok := affected[dependent]; ok {
				continue
			}
			affected[dependent] = struct{}{}
			queue = append(queue, dependent)
		}

		if columnDep, ok := dependencyColumnKeyForCell(current); ok {
			for _, dependent := range dependents[columnDep] {
				if _, ok := affected[dependent]; ok {
					continue
				}
				affected[dependent] = struct{}{}
				queue = append(queue, dependent)
			}
		}
	}

	if len(affected) == 0 {
		return &dependencyGraph{
			nodes:          make(map[string]*formulaNode),
			columnMetadata: g.columnMetadata,
		}
	}

	subgraph := &dependencyGraph{
		nodes:          make(map[string]*formulaNode, len(affected)),
		levels:         make([][]string, 0, len(g.levels)),
		columnMetadata: g.columnMetadata,
	}

	for cell, node := range g.nodes {
		if _, ok := affected[cell]; ok {
			subgraph.nodes[cell] = node
		}
	}

	for _, level := range g.levels {
		filtered := make([]string, 0, len(level))
		for _, cell := range level {
			if _, ok := affected[cell]; ok {
				filtered = append(filtered, cell)
			}
		}
		if len(filtered) > 0 {
			subgraph.levels = append(subgraph.levels, filtered)
		}
	}

	return subgraph
}

func dependencyColumnKeyForCell(fullCell string) (string, bool) {
	if strings.HasPrefix(fullCell, "COLUMN:") {
		return fullCell, true
	}
	parts := strings.SplitN(fullCell, "!", 2)
	if len(parts) != 2 {
		return "", false
	}
	colNum, _, err := CellNameToCoordinates(parts[1])
	if err != nil {
		return "", false
	}
	colName, err := ColumnNumberToName(colNum)
	if err != nil {
		return "", false
	}
	return "COLUMN:" + parts[0] + "!" + colName, true
}

func formatDirtyCellsForLog(cells []string) string {
	if len(cells) == 0 {
		return ""
	}
	const limit = 5
	if len(cells) <= limit {
		return strings.Join(cells, ", ")
	}
	return strings.Join(cells[:limit], ", ") + ", ..."
}

func (f *File) recalculateDirtyValueSubgraph(dirtyCells []string, logContext string) bool {
	if len(dirtyCells) == 0 || f.hasVolatileDependencies() {
		return false
	}

	graph := f.buildDependencyGraph()
	if graph == nil {
		return false
	}

	subgraph := graph.buildAffectedSubgraph(dirtyCells)
	if subgraph == nil {
		return false
	}
	if len(subgraph.nodes) == 0 {
		log.Printf("  ⏭️ [%s] No affected formulas for dirty cells: %s", logContext, formatDirtyCellsForLog(dirtyCells))
		f.clearDirtyValueCells()
		log.Printf("✅ [%s] Completed", logContext)
		return true
	}

	affectedCells := make([]string, 0, len(subgraph.nodes))
	for cell := range subgraph.nodes {
		affectedCells = append(affectedCells, cell)
	}
	f.clearFormulaCaches(affectedCells)

	log.Printf("  ⚡ [%s] Recalculating affected subgraph: %d formulas for dirty cells %s",
		logContext, len(subgraph.nodes), formatDirtyCellsForLog(dirtyCells))
	f.calculateByDAG(subgraph)
	f.clearDirtyValueCells()
	log.Printf("✅ [%s] Completed", logContext)
	return true
}

func (f *File) clearFormulaResultCaches(cells []string) {
	if len(cells) == 0 {
		return
	}

	subExprPrefixes := make([]string, 0, len(cells))
	for _, cell := range cells {
		f.calcCache.Delete(cell)
		f.calcCache.Delete(cell + "!raw=true")
		f.calcCache.Delete(cell + "!raw=false")
		f.pgWholeCellCache.Delete(cell)
		subExprPrefixes = append(subExprPrefixes, cell+"!subexpr:")
	}

	var keysToDelete []string
	f.calcCache.Range(func(key, value interface{}) bool {
		keyStr, ok := key.(string)
		if !ok {
			return true
		}
		for _, prefix := range subExprPrefixes {
			if strings.HasPrefix(keyStr, prefix) {
				keysToDelete = append(keysToDelete, keyStr)
				break
			}
		}
		return true
	})
	for _, key := range keysToDelete {
		f.calcCache.Delete(key)
	}

	f.mu.Lock()
	cache := f.worksheetCache
	f.mu.Unlock()
	if cache != nil {
		for _, cell := range cells {
			parts := strings.SplitN(cell, "!", 2)
			if len(parts) != 2 {
				continue
			}
			cache.Delete(parts[0], parts[1])
		}
	}
}

func (f *File) invalidateValueChangeCaches(sheet, cell string) {
	if f.hasVolatileDependencies() {
		f.calcCache.Clear()
		f.rangeCache.Clear()
		return
	}

	graph := f.getCachedDependencyGraph()
	if graph == nil {
		f.calcCache.Clear()
		f.rangeCache.Clear()
		return
	}

	f.clearCellCache(sheet, cell)

	subgraph := graph.buildAffectedSubgraph([]string{sheet + "!" + cell})
	if subgraph == nil || len(subgraph.nodes) == 0 {
		return
	}

	affectedCells := make([]string, 0, len(subgraph.nodes))
	for formulaCell := range subgraph.nodes {
		affectedCells = append(affectedCells, formulaCell)
	}
	f.clearFormulaResultCaches(affectedCells)
}

func (f *File) clearFormulaCaches(cells []string) {
	f.clearFormulaResultCaches(cells)
	if f.rangeCache.Len() > 0 {
		f.rangeCache.Clear()
	}
}
