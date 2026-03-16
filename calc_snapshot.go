package excelize

import (
	"encoding/json"
	"errors"
	"sort"
	"strings"
)

const (
	calcSnapshotWorkbookPath        = "xl/excelize/calc-snapshot.json"
	calcSnapshotWorkbookPartName    = "/" + calcSnapshotWorkbookPath
	calcSnapshotWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.calc-snapshot+json"
	calcSnapshotVersion             = 1
)

var calcSnapshotBuildTestHook func()

type calculationSnapshot struct {
	Version               int                                `json:"version"`
	DependencyGraph       calculationSnapshotDependencyGraph `json:"dependency_graph"`
	FormulaOptimizations  []calculationSnapshotFormulaClass  `json:"formula_optimizations,omitempty"`
	PGLookupOptimizations []calculationSnapshotPGLookupClass `json:"pg_lookup_optimizations,omitempty"`
	PGCaches              calculationSnapshotPGCaches        `json:"pg_caches,omitempty"`
}

type calculationSnapshotDependencyGraph struct {
	Nodes          []calculationSnapshotNode       `json:"nodes"`
	Levels         [][]string                      `json:"levels,omitempty"`
	ColumnMetadata []calculationSnapshotColumnMeta `json:"column_metadata,omitempty"`
}

type calculationSnapshotNode struct {
	Cell         string   `json:"cell"`
	Formula      string   `json:"formula"`
	Dependencies []string `json:"dependencies,omitempty"`
	Level        int      `json:"level"`
}

type calculationSnapshotColumnMeta struct {
	Key         string `json:"key"`
	HasFormulas bool   `json:"has_formulas"`
	FormulaRows []int  `json:"formula_rows,omitempty"`
	MaxRow      int    `json:"max_row"`
}

type calculationSnapshotFormulaClass struct {
	Formula          string `json:"formula"`
	HasAverageOffset bool   `json:"has_average_offset"`
	SumifsExpr       string `json:"sumifs_expr,omitempty"`
	PureSumifs       bool   `json:"pure_sumifs"`
	SumifsSource     string `json:"sumifs_source,omitempty"`
	IndexMatchExpr   string `json:"index_match_expr,omitempty"`
	IndexSource      string `json:"index_source,omitempty"`
	IsBatchType      bool   `json:"is_batch_type"`
}

type calculationSnapshotPGLookupClass struct {
	Sheet   string   `json:"sheet"`
	Formula string   `json:"formula"`
	Whole   bool     `json:"whole"`
	Exprs   []string `json:"exprs,omitempty"`
}

type calculationSnapshotPGCaches struct {
	WholeFormulaResults []calculationSnapshotPGWholeFormulaResult `json:"whole_formula_results,omitempty"`
	WholeCellResults    []calculationSnapshotPGWholeFormulaResult `json:"whole_cell_results,omitempty"`
	CalcEntries         []calculationSnapshotPGCalcEntry          `json:"calc_entries,omitempty"`
}

type calculationSnapshotPGWholeFormulaResult struct {
	Key   string                              `json:"key"`
	Value calculationSnapshotScalarFormulaArg `json:"value"`
}

type calculationSnapshotPGCalcEntry struct {
	Key       string                                 `json:"key"`
	Match     *calculationSnapshotPGExactMatchResult `json:"match,omitempty"`
	CellValue *calculationSnapshotPGMirrorCellValue  `json:"cell_value,omitempty"`
}

type calculationSnapshotPGExactMatchResult struct {
	Position int  `json:"position"`
	RowNum   int  `json:"row_num"`
	ColNum   int  `json:"col_num"`
	Found    bool `json:"found"`
}

type calculationSnapshotPGMirrorCellValue struct {
	Value     string `json:"value"`
	ValueType string `json:"value_type"`
}

type calculationSnapshotScalarFormulaArg struct {
	Type    string  `json:"type"`
	Number  float64 `json:"number,omitempty"`
	String  string  `json:"string,omitempty"`
	Boolean bool    `json:"boolean,omitempty"`
	Error   string  `json:"error,omitempty"`
}

// BuildCalculationSnapshot serializes the current dependency graph and formula classification caches.
func (f *File) BuildCalculationSnapshot() ([]byte, error) {
	if calcSnapshotBuildTestHook != nil {
		calcSnapshotBuildTestHook()
	}
	snapshot := f.buildCalculationSnapshot()
	return json.Marshal(snapshot)
}

// LoadCalculationSnapshot restores a previously serialized dependency graph and optimization caches.
//
// The snapshot must match the current workbook structure and formulas.
func (f *File) LoadCalculationSnapshot(data []byte) error {
	if len(data) == 0 {
		return errors.New("calculation snapshot is empty")
	}
	var snapshot calculationSnapshot
	if err := json.Unmarshal(data, &snapshot); err != nil {
		return err
	}
	if err := f.applyCalculationSnapshot(&snapshot); err != nil {
		return err
	}
	f.cacheCalculationSnapshot(data)
	return nil
}

// EmbedCalculationSnapshot stores a serialized calculation snapshot inside the workbook package.
func (f *File) EmbedCalculationSnapshot() error {
	data, err := f.BuildCalculationSnapshot()
	if err != nil {
		return err
	}
	if err := f.replaceWorkbookPartContentType(calcSnapshotWorkbookPartName, calcSnapshotWorkbookContentType); err != nil {
		return err
	}
	f.Pkg.Store(calcSnapshotWorkbookPath, data)
	f.mu.Lock()
	f.calcSnapshotAuto = true
	f.calcSnapshotDirty = false
	f.calcSnapshotData = append(f.calcSnapshotData[:0], data...)
	f.mu.Unlock()
	return nil
}

// SetCalculationSnapshotAutoEmbed enables or disables automatic snapshot refresh during save/write.
func (f *File) SetCalculationSnapshotAutoEmbed(enabled bool) {
	f.mu.Lock()
	f.calcSnapshotAuto = enabled
	f.mu.Unlock()
	if !enabled {
		f.removeEmbeddedCalculationSnapshot()
	}
}

// CalculationSnapshotAutoEmbedEnabled reports whether automatic snapshot refresh is enabled.
func (f *File) CalculationSnapshotAutoEmbedEnabled() bool {
	f.mu.Lock()
	defer f.mu.Unlock()
	return f.calcSnapshotAuto
}

func (f *File) buildCalculationSnapshot() calculationSnapshot {
	graph := f.buildDependencyGraph()
	snapshot := calculationSnapshot{
		Version: calcSnapshotVersion,
		DependencyGraph: calculationSnapshotDependencyGraph{
			Nodes:          make([]calculationSnapshotNode, 0, len(graph.nodes)),
			Levels:         cloneDependencyLevels(graph.levels),
			ColumnMetadata: make([]calculationSnapshotColumnMeta, 0, len(graph.columnMetadata)),
		},
	}

	nodeKeys := make([]string, 0, len(graph.nodes))
	for cell := range graph.nodes {
		nodeKeys = append(nodeKeys, cell)
	}
	sort.Strings(nodeKeys)

	seenFormulas := make(map[string]struct{}, len(nodeKeys))
	seenPGLookups := make(map[string]struct{}, len(nodeKeys))
	for _, cell := range nodeKeys {
		node := graph.nodes[cell]
		if node == nil {
			continue
		}
		deps := append([]string(nil), node.dependencies...)
		snapshot.DependencyGraph.Nodes = append(snapshot.DependencyGraph.Nodes, calculationSnapshotNode{
			Cell:         node.cell,
			Formula:      node.formula,
			Dependencies: deps,
			Level:        node.level,
		})

		if node.formula != "" {
			if _, ok := seenFormulas[node.formula]; !ok {
				seenFormulas[node.formula] = struct{}{}
				class := getFormulaOptimizationClass(node.formula)
				snapshot.FormulaOptimizations = append(snapshot.FormulaOptimizations, calculationSnapshotFormulaClass{
					Formula:          node.formula,
					HasAverageOffset: class.hasAverageOffset,
					SumifsExpr:       class.sumifsExpr,
					PureSumifs:       class.pureSumifs,
					SumifsSource:     class.sumifsSource,
					IndexMatchExpr:   class.indexMatchExpr,
					IndexSource:      class.indexSource,
					IsBatchType:      class.isBatchType,
				})
			}

			if parts := stringsSplitCellRef(node.cell); len(parts) == 2 {
				key := parts[0] + "\x00" + node.formula
				if _, ok := seenPGLookups[key]; !ok {
					seenPGLookups[key] = struct{}{}
					class := getPGLookupOptimizationClass(parts[0], node.formula)
					snapshot.PGLookupOptimizations = append(snapshot.PGLookupOptimizations, calculationSnapshotPGLookupClass{
						Sheet:   parts[0],
						Formula: node.formula,
						Whole:   class.whole,
						Exprs:   append([]string(nil), class.exprs...),
					})
				}
			}
		}
	}

	columnKeys := make([]string, 0, len(graph.columnMetadata))
	for key := range graph.columnMetadata {
		columnKeys = append(columnKeys, key)
	}
	sort.Strings(columnKeys)
	for _, key := range columnKeys {
		meta := graph.columnMetadata[key]
		if meta == nil {
			continue
		}
		rows := make([]int, 0, len(meta.formulaRows))
		for row := range meta.formulaRows {
			rows = append(rows, row)
		}
		sort.Ints(rows)
		snapshot.DependencyGraph.ColumnMetadata = append(snapshot.DependencyGraph.ColumnMetadata, calculationSnapshotColumnMeta{
			Key:         key,
			HasFormulas: meta.hasFormulas,
			FormulaRows: rows,
			MaxRow:      meta.maxRow,
		})
	}
	snapshot.PGCaches = f.buildCalculationSnapshotPGCaches()

	return snapshot
}

func (f *File) applyCalculationSnapshot(snapshot *calculationSnapshot) error {
	if snapshot == nil {
		return errors.New("calculation snapshot is nil")
	}
	if snapshot.Version != calcSnapshotVersion {
		return errors.New("unsupported calculation snapshot version")
	}

	graph := &dependencyGraph{
		nodes:          make(map[string]*formulaNode, len(snapshot.DependencyGraph.Nodes)),
		levels:         cloneDependencyLevels(snapshot.DependencyGraph.Levels),
		columnMetadata: make(map[string]*columnMeta, len(snapshot.DependencyGraph.ColumnMetadata)),
	}
	for _, node := range snapshot.DependencyGraph.Nodes {
		graph.nodes[node.Cell] = &formulaNode{
			cell:         node.Cell,
			formula:      node.Formula,
			dependencies: append([]string(nil), node.Dependencies...),
			level:        node.Level,
		}
	}
	if len(graph.levels) == 0 {
		graph.levels = buildDependencyLevelsFromNodes(graph.nodes)
	}
	for _, meta := range snapshot.DependencyGraph.ColumnMetadata {
		rowSet := make(map[int]bool, len(meta.FormulaRows))
		for _, row := range meta.FormulaRows {
			rowSet[row] = true
		}
		graph.columnMetadata[meta.Key] = &columnMeta{
			hasFormulas: meta.HasFormulas,
			formulaRows: rowSet,
			maxRow:      meta.MaxRow,
		}
	}
	for _, class := range snapshot.FormulaOptimizations {
		globalFormulaOptimizationCache.Store(class.Formula, formulaOptimizationClass{
			hasAverageOffset: class.HasAverageOffset,
			sumifsExpr:       class.SumifsExpr,
			pureSumifs:       class.PureSumifs,
			sumifsSource:     class.SumifsSource,
			indexMatchExpr:   class.IndexMatchExpr,
			indexSource:      class.IndexSource,
			isBatchType:      class.IsBatchType,
		})
	}
	for _, class := range snapshot.PGLookupOptimizations {
		globalPGLookupFormulaCache.Store(class.Sheet+"\x00"+class.Formula, pgLookupOptimizationClass{
			whole: class.Whole,
			exprs: append([]string(nil), class.Exprs...),
		})
	}
	f.clearPGSnapshotCacheEntries()
	for _, entry := range snapshot.PGCaches.WholeFormulaResults {
		arg, ok := deserializeSnapshotScalarFormulaArg(entry.Value)
		if !ok {
			continue
		}
		f.pgLookupResultCache.Store(entry.Key, arg)
	}
	for _, entry := range snapshot.PGCaches.WholeCellResults {
		arg, ok := deserializeSnapshotScalarFormulaArg(entry.Value)
		if !ok {
			continue
		}
		f.pgWholeCellCache.Store(entry.Key, arg)
	}
	for _, entry := range snapshot.PGCaches.CalcEntries {
		switch {
		case entry.Match != nil:
			f.pgCalcCache.Store(entry.Key, pgExactMatchResult{
				Position: entry.Match.Position,
				RowNum:   entry.Match.RowNum,
				ColNum:   entry.Match.ColNum,
				Found:    entry.Match.Found,
			})
		case entry.CellValue != nil:
			f.pgCalcCache.Store(entry.Key, pgMirrorCellValue{
				Value:     entry.CellValue.Value,
				ValueType: entry.CellValue.ValueType,
			})
		}
	}
	f.storeDependencyGraph(graph)
	return nil
}

func (f *File) tryLoadEmbeddedCalculationSnapshot() {
	content, ok := f.Pkg.Load(calcSnapshotWorkbookPath)
	if !ok {
		return
	}
	data, ok := content.([]byte)
	if !ok || len(data) == 0 {
		return
	}
	if err := f.LoadCalculationSnapshot(data); err == nil {
		f.mu.Lock()
		f.calcSnapshotAuto = true
		f.mu.Unlock()
	}
}

func (f *File) removeEmbeddedCalculationSnapshot() {
	f.mu.Lock()
	f.calcSnapshotData = nil
	f.calcSnapshotDirty = true
	f.mu.Unlock()
	f.Pkg.Delete(calcSnapshotWorkbookPath)
	_ = f.removeContentTypesPart(calcSnapshotWorkbookContentType, calcSnapshotWorkbookPartName)
}

func (f *File) prepareCalculationSnapshotForWrite() error {
	f.mu.Lock()
	auto := f.calcSnapshotAuto
	dirty := f.calcSnapshotDirty
	data := append([]byte(nil), f.calcSnapshotData...)
	f.mu.Unlock()
	if !auto {
		if _, ok := f.Pkg.Load(calcSnapshotWorkbookPath); !ok {
			return nil
		}
	}
	if !dirty && len(data) > 0 {
		if err := f.replaceWorkbookPartContentType(calcSnapshotWorkbookPartName, calcSnapshotWorkbookContentType); err != nil {
			return err
		}
		f.Pkg.Store(calcSnapshotWorkbookPath, data)
		return nil
	}
	return f.EmbedCalculationSnapshot()
}

func (f *File) cacheCalculationSnapshot(data []byte) {
	f.mu.Lock()
	defer f.mu.Unlock()
	f.calcSnapshotData = append(f.calcSnapshotData[:0], data...)
	f.calcSnapshotDirty = false
}

func (f *File) markCalculationSnapshotDirty() {
	f.mu.Lock()
	f.calcSnapshotDirty = true
	f.mu.Unlock()
}

func (f *File) storePGCalcCache(key string, value interface{}) {
	f.pgCalcCache.Store(key, value)
	f.markCalculationSnapshotDirty()
}

func (f *File) loadPGWholeCellCache(key string) (formulaArg, bool) {
	if key == "" {
		return formulaArg{}, false
	}
	cached, ok := f.pgWholeCellCache.Load(key)
	if !ok {
		return formulaArg{}, false
	}
	result, valid := cached.(formulaArg)
	return result, valid
}

func (f *File) storePGWholeCellCache(key string, value formulaArg) {
	if key == "" {
		return
	}
	f.pgWholeCellCache.Store(key, value)
	f.markCalculationSnapshotDirty()
}

func (f *File) buildCalculationSnapshotPGCaches() calculationSnapshotPGCaches {
	var caches calculationSnapshotPGCaches

	f.pgLookupResultCache.Range(func(key, value interface{}) bool {
		keyStr, ok := key.(string)
		if !ok {
			return true
		}
		arg, ok := value.(formulaArg)
		if !ok {
			return true
		}
		serialized, ok := serializeSnapshotScalarFormulaArg(arg)
		if !ok {
			return true
		}
		caches.WholeFormulaResults = append(caches.WholeFormulaResults, calculationSnapshotPGWholeFormulaResult{
			Key:   keyStr,
			Value: serialized,
		})
		return true
	})
	sort.Slice(caches.WholeFormulaResults, func(i, j int) bool {
		return caches.WholeFormulaResults[i].Key < caches.WholeFormulaResults[j].Key
	})

	f.pgWholeCellCache.Range(func(key, value interface{}) bool {
		keyStr, ok := key.(string)
		if !ok {
			return true
		}
		arg, ok := value.(formulaArg)
		if !ok {
			return true
		}
		serialized, ok := serializeSnapshotScalarFormulaArg(arg)
		if !ok {
			return true
		}
		caches.WholeCellResults = append(caches.WholeCellResults, calculationSnapshotPGWholeFormulaResult{
			Key:   keyStr,
			Value: serialized,
		})
		return true
	})
	sort.Slice(caches.WholeCellResults, func(i, j int) bool {
		return caches.WholeCellResults[i].Key < caches.WholeCellResults[j].Key
	})

	f.pgCalcCache.Range(func(key, value interface{}) bool {
		keyStr, ok := key.(string)
		if !ok {
			return true
		}
		entry := calculationSnapshotPGCalcEntry{Key: keyStr}
		switch v := value.(type) {
		case pgExactMatchResult:
			entry.Match = &calculationSnapshotPGExactMatchResult{
				Position: v.Position,
				RowNum:   v.RowNum,
				ColNum:   v.ColNum,
				Found:    v.Found,
			}
		case pgMirrorCellValue:
			entry.CellValue = &calculationSnapshotPGMirrorCellValue{
				Value:     v.Value,
				ValueType: v.ValueType,
			}
		default:
			return true
		}
		caches.CalcEntries = append(caches.CalcEntries, entry)
		return true
	})
	sort.Slice(caches.CalcEntries, func(i, j int) bool {
		return caches.CalcEntries[i].Key < caches.CalcEntries[j].Key
	})

	return caches
}

func serializeSnapshotScalarFormulaArg(arg formulaArg) (calculationSnapshotScalarFormulaArg, bool) {
	switch arg.Type {
	case ArgEmpty:
		return calculationSnapshotScalarFormulaArg{Type: "empty"}, true
	case ArgString:
		return calculationSnapshotScalarFormulaArg{Type: "string", String: arg.String}, true
	case ArgNumber:
		return calculationSnapshotScalarFormulaArg{Type: "number", Number: arg.Number, Boolean: arg.Boolean}, true
	case ArgError:
		return calculationSnapshotScalarFormulaArg{Type: "error", String: arg.String, Error: arg.Error}, true
	default:
		return calculationSnapshotScalarFormulaArg{}, false
	}
}

func deserializeSnapshotScalarFormulaArg(arg calculationSnapshotScalarFormulaArg) (formulaArg, bool) {
	switch arg.Type {
	case "empty":
		return newEmptyFormulaArg(), true
	case "string":
		return newStringFormulaArg(arg.String), true
	case "number":
		if arg.Boolean {
			return newBoolFormulaArg(arg.Number != 0), true
		}
		return newNumberFormulaArg(arg.Number), true
	case "error":
		msg := arg.Error
		if msg == "" {
			msg = arg.String
		}
		return newErrorFormulaArg(arg.String, msg), true
	default:
		return formulaArg{}, false
	}
}

func (f *File) clearPGSnapshotCacheEntries() {
	f.clearPGLookupResultCache()
	var deleteWholeCellKeys []interface{}
	f.pgWholeCellCache.Range(func(key, _ interface{}) bool {
		deleteWholeCellKeys = append(deleteWholeCellKeys, key)
		return true
	})
	for _, key := range deleteWholeCellKeys {
		f.pgWholeCellCache.Delete(key)
	}
	var deleteKeys []interface{}
	f.pgCalcCache.Range(func(key, _ interface{}) bool {
		deleteKeys = append(deleteKeys, key)
		return true
	})
	for _, key := range deleteKeys {
		f.pgCalcCache.Delete(key)
	}
}

func (f *File) replaceWorkbookPartContentType(partName, contentType string) error {
	if err := f.removeContentTypesPart(contentType, partName); err != nil {
		return err
	}
	return f.setContentTypes(partName, contentType)
}

func buildDependencyLevelsFromNodes(nodes map[string]*formulaNode) [][]string {
	maxLevel := -1
	for _, node := range nodes {
		if node != nil && node.level > maxLevel {
			maxLevel = node.level
		}
	}
	if maxLevel < 0 {
		return nil
	}
	levels := make([][]string, maxLevel+1)
	for cell, node := range nodes {
		if node == nil || node.level < 0 {
			continue
		}
		levels[node.level] = append(levels[node.level], cell)
	}
	for i := range levels {
		sort.Strings(levels[i])
	}
	return levels
}

func cloneDependencyLevels(levels [][]string) [][]string {
	if len(levels) == 0 {
		return nil
	}
	cloned := make([][]string, len(levels))
	for i := range levels {
		cloned[i] = append([]string(nil), levels[i]...)
	}
	return cloned
}

func stringsSplitCellRef(ref string) []string {
	return strings.SplitN(ref, "!", 2)
}
