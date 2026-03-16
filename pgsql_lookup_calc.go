package excelize

import (
	"context"
	"database/sql"
	"errors"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

type pgLookupRange struct {
	Sheet    string
	StartRow int
	EndRow   int
	StartCol int
	EndCol   int
}

func (r pgLookupRange) isVertical() bool {
	return r.StartCol == r.EndCol
}

func (r pgLookupRange) isHorizontal() bool {
	return r.StartRow == r.EndRow
}

type pgLookupEnvelope struct {
	body        string
	fallback    string
	hasFallback bool
}

type pgConditionalLookupFormula struct {
	condition pgLookupCondition
	whenTrue  string
	whenFalse string
}

type pgLookupCondition struct {
	operator      string
	lookupFormula string
	lookupOnLeft  bool
	otherExpr     string
}

type pgMatchLookupFormula struct {
	lookupExpr  string
	lookupRange pgLookupRange
}

type pgVLookupFormula struct {
	lookupExpr    string
	tableRange    pgLookupRange
	returnColIdx1 int
}

type pgIndexMatchFormula struct {
	lookupExpr  string
	matchRange  pgLookupRange
	returnRange pgLookupRange
}

type pgExactMatchResult struct {
	Position int
	RowNum   int
	ColNum   int
	Found    bool
}

type pgBatchMatchItem struct {
	fullCell string
	sheet    string
	env      pgLookupEnvelope
	formula  *pgMatchLookupFormula
	lookup   formulaArg
}

type pgBatchVLookupItem struct {
	fullCell string
	sheet    string
	env      pgLookupEnvelope
	formula  *pgVLookupFormula
	lookup   formulaArg
}

type pgBatchIndexMatchItem struct {
	fullCell string
	sheet    string
	env      pgLookupEnvelope
	formula  *pgIndexMatchFormula
	lookup   formulaArg
}

type pgLookupBatchTarget struct {
	id      string
	sheet   string
	formula string
}

type pgBatchConditionalItem struct {
	target  pgLookupBatchTarget
	formula *pgConditionalLookupFormula
}

var (
	pgExactMatchCacheTestHook func(key string, hit bool, loaded bool, valueType string)
	pgCellValueCacheTestHook  func(key string, hit bool, loaded bool, valueType string)
)

// PreloadPostgresLookupCache scans formulas and warms PostgreSQL lookup caches.
//
// It pre-populates both whole-formula result cache and the underlying PG query
// caches used by supported lookup sub-expressions. sheets is optional; when
// omitted, all worksheets are scanned.
func (f *File) PreloadPostgresLookupCache(ctx context.Context, sheets ...string) error {
	db, _, ok := f.getPostgresLookupRuntime()
	if !ok || db == nil {
		return nil
	}
	if ctx == nil {
		ctx = context.Background()
	}

	targetSheets := sheets
	if len(targetSheets) == 0 {
		targetSheets = f.GetSheetList()
	}

	graph := f.buildDependencyGraph()
	worksheetCache := f.prepareWorksheetCacheForGraph(graph)
	wholeTargets := make(map[string]pgLookupBatchTarget)
	subExprTargets := make(map[string]pgLookupBatchTarget)
	scalarRefs := make(map[string]map[string]struct{})

	for _, sheet := range targetSheets {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}
		for rowIdx := range ws.SheetData.Row {
			for colIdx := range ws.SheetData.Row[rowIdx].C {
				cell := &ws.SheetData.Row[rowIdx].C[colIdx]
				if cell.F == nil {
					continue
				}
				formula := cell.F.Content
				if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
					formula, err = getSharedFormula(ws, *cell.F.Si, cell.R)
					if err != nil {
						continue
					}
				}
				if formula == "" {
					continue
				}

				class := getPGLookupOptimizationClass(sheet, formula)
				f.collectPGFormulaScalarRefs(sheet, formula, scalarRefs)
				if class.whole {
					fullCell := sheet + "!" + cell.R
					wholeTargets[fullCell] = pgLookupBatchTarget{
						id:      fullCell,
						sheet:   sheet,
						formula: formula,
					}
					continue
				}
				for idx, expr := range class.exprs {
					targetID := sheet + "!" + cell.R + "!pgwarm:" + strconv.Itoa(idx)
					subExprTargets[targetID] = pgLookupBatchTarget{
						id:      targetID,
						sheet:   sheet,
						formula: expr,
					}
				}
			}
		}
	}

	if len(wholeTargets) == 0 && len(subExprTargets) == 0 {
		return nil
	}
	if err := f.preloadWorksheetCacheRefs(scalarRefs, worksheetCache); err != nil {
		return err
	}

	if len(subExprTargets) > 0 {
		_ = f.batchCalculatePostgresLookupTargetsWithCache(subExprTargets, worksheetCache)
	}
	if len(wholeTargets) == 0 {
		return nil
	}

	results := f.batchCalculatePostgresLookupTargetsWithCache(wholeTargets, worksheetCache)
	for id, result := range results {
		target, ok := wholeTargets[id]
		if !ok {
			continue
		}
		f.storePGWholeCellCache(id, result)
		cacheKey, cacheable, err := f.buildPGWholeFormulaResultCacheKey(target.sheet, target.formula, worksheetCache)
		if err != nil || !cacheable {
			continue
		}
		f.storePGLookupResultCache(cacheKey, result)
	}
	return nil
}

func (f *File) collectPGFormulaScalarRefs(defaultSheet, formula string, refs map[string]map[string]struct{}) {
	if conditional, ok := parsePGConditionalLookupFormula(defaultSheet, formula); ok {
		f.collectPGFormulaScalarRefs(defaultSheet, conditional.condition.lookupFormula, refs)
		f.collectPGScalarExprRefs(defaultSheet, conditional.condition.otherExpr, refs)
		f.collectPGScalarExprRefs(defaultSheet, conditional.whenTrue, refs)
		f.collectPGScalarExprRefs(defaultSheet, conditional.whenFalse, refs)
		return
	}

	env, ok := parsePGLookupEnvelope(formula)
	if !ok {
		f.collectPGScalarExprRefs(defaultSheet, formula, refs)
		return
	}

	switch parsed, ok := parsePGMatchLookupFormula(defaultSheet, env.body); {
	case ok:
		f.collectPGScalarExprRefs(defaultSheet, parsed.lookupExpr, refs)
	case false:
		if parsed, ok := parsePGVLookupFormula(defaultSheet, env.body); ok {
			f.collectPGScalarExprRefs(defaultSheet, parsed.lookupExpr, refs)
		} else if parsed, ok := parsePGIndexMatchFormula(defaultSheet, env.body); ok {
			f.collectPGScalarExprRefs(defaultSheet, parsed.lookupExpr, refs)
		}
	}
	if env.hasFallback {
		f.collectPGScalarExprRefs(defaultSheet, env.fallback, refs)
	}
}

func (f *File) collectPGScalarExprRefs(defaultSheet, expr string, refs map[string]map[string]struct{}) {
	expr = strings.TrimSpace(expr)
	if expr == "" || isPGLiteralScalarExpr(expr) {
		return
	}
	if _, ok := parsePGConditionalLookupFormula(defaultSheet, expr); ok || isSupportedPGBaseLookupFormula(defaultSheet, expr) {
		f.collectPGFormulaScalarRefs(defaultSheet, expr, refs)
		return
	}

	ref, _, _, err := parseRef(strings.ReplaceAll(expr, "$", ""))
	if err != nil {
		return
	}
	ref.Sheet = strings.Trim(ref.Sheet, "'")
	if ref.Sheet == "" {
		ref.Sheet = defaultSheet
	}
	cellName, err := CoordinatesToCellName(ref.Col, ref.Row)
	if err != nil {
		return
	}
	if refs[ref.Sheet] == nil {
		refs[ref.Sheet] = make(map[string]struct{})
	}
	refs[ref.Sheet][cellName] = struct{}{}
}

func (f *File) preloadWorksheetCacheRefs(refs map[string]map[string]struct{}, worksheetCache *WorksheetCache) error {
	if worksheetCache == nil || len(refs) == 0 {
		return nil
	}
	for sheet, cells := range refs {
		for cell := range cells {
			value, err := f.GetCellValue(sheet, cell, Options{RawCellValue: true})
			if err != nil {
				return err
			}
			cellType, err := f.GetCellType(sheet, cell)
			if err != nil {
				return err
			}
			worksheetCache.Set(sheet, cell, inferCellValueType(value, cellType))
		}
	}
	return nil
}

func (f *File) getPostgresLookupRuntime() (*sql.DB, *PGSyncOptions, bool) {
	f.mu.Lock()
	defer f.mu.Unlock()

	if !f.pgMirrorEnabled || f.pgMirrorDB == nil || f.pgMirrorOpts == nil || f.pgLookupBypass > 0 {
		return nil, nil, false
	}
	return f.pgMirrorDB, clonePGSyncOptions(f.pgMirrorOpts), true
}

func (f *File) tryCalcPostgresLookupFormula(sheet, cell, formula string, worksheetCache *WorksheetCache) (formulaArg, bool, error) {
	db, cfg, ok := f.getPostgresLookupRuntime()
	if !ok {
		return formulaArg{}, false, nil
	}

	if cell != "" {
		fullCell := sheet + "!" + cell
		if class := getPGLookupOptimizationClass(sheet, formula); class.whole {
			if cached, ok := f.loadPGWholeCellCache(fullCell); ok {
				return cached, true, nil
			}
		}
	}

	cacheKey, cacheable, cacheErr := f.buildPGWholeFormulaResultCacheKey(sheet, formula, worksheetCache)
	if cacheErr == nil && cacheable {
		if cached, ok := f.loadPGLookupResultCache(cacheKey); ok {
			if cell != "" {
				f.storePGWholeCellCache(sheet+"!"+cell, cached)
			}
			return cached, true, nil
		}
	}

	var (
		result  formulaArg
		matched bool
		evalErr error
	)
	if conditional, ok := parsePGConditionalLookupFormula(sheet, formula); ok {
		result, evalErr = f.evalPGConditionalLookupFormula(context.Background(), db, cfg, sheet, conditional, worksheetCache)
		matched = true
	} else {
		env, ok := parsePGLookupEnvelope(formula)
		if !ok {
			return formulaArg{}, false, nil
		}
		result, matched, evalErr = f.evalPGSupportedLookupEnvelope(context.Background(), db, cfg, sheet, env, worksheetCache)
	}
	if evalErr == nil && matched && cacheErr == nil && cacheable {
		f.storePGLookupResultCache(cacheKey, result)
	}
	if evalErr == nil && matched && cell != "" {
		f.storePGWholeCellCache(sheet+"!"+cell, result)
	}
	return result, matched, evalErr
}

func isSupportedPGWholeLookupFormula(sheet, formula string) bool {
	if _, ok := parsePGConditionalLookupFormula(sheet, formula); ok {
		return true
	}
	return isSupportedPGBaseLookupFormula(sheet, formula)
}

func isSupportedPGBaseLookupFormula(sheet, formula string) bool {
	env, ok := parsePGLookupEnvelope(formula)
	if !ok {
		return false
	}
	if _, ok := parsePGMatchLookupFormula(sheet, env.body); ok {
		return true
	}
	if _, ok := parsePGVLookupFormula(sheet, env.body); ok {
		return true
	}
	if _, ok := parsePGIndexMatchFormula(sheet, env.body); ok {
		return true
	}
	return false
}

func (f *File) evalPGSupportedLookupEnvelope(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, env pgLookupEnvelope, worksheetCache *WorksheetCache) (formulaArg, bool, error) {
	if matchFormula, ok := parsePGMatchLookupFormula(sheet, env.body); ok {
		result, evalErr := f.evalPGMatchFormula(ctx, db, cfg, sheet, matchFormula, env, worksheetCache)
		return result, true, evalErr
	}
	if vlookupFormula, ok := parsePGVLookupFormula(sheet, env.body); ok {
		result, evalErr := f.evalPGVLookupFormula(ctx, db, cfg, sheet, vlookupFormula, env, worksheetCache)
		return result, true, evalErr
	}
	if indexMatchFormula, ok := parsePGIndexMatchFormula(sheet, env.body); ok {
		result, evalErr := f.evalPGIndexMatchFormula(ctx, db, cfg, sheet, indexMatchFormula, env, worksheetCache)
		return result, true, evalErr
	}
	return formulaArg{}, false, nil
}

func extractAllPGSupportedLookupExprs(sheet, formula string) []string {
	var exprs []string
	remaining := formula
	for {
		expr := extractFirstPGSupportedLookupExpr(sheet, remaining)
		if expr == "" {
			return exprs
		}
		exprs = append(exprs, expr)
		idx := strings.Index(remaining, expr)
		if idx < 0 {
			return exprs
		}
		remaining = remaining[idx+len(expr):]
	}
}

func (f *File) batchCalculatePostgresLookupsWithCache(formulas map[string]string, worksheetCache *WorksheetCache) map[string]string {
	targets := make(map[string]pgLookupBatchTarget, len(formulas))
	for fullCell, formula := range formulas {
		parts := strings.SplitN(fullCell, "!", 2)
		if len(parts) != 2 {
			continue
		}
		targets[fullCell] = pgLookupBatchTarget{
			id:      fullCell,
			sheet:   parts[0],
			formula: formula,
		}
	}
	typedResults := f.batchCalculatePostgresLookupTargetsWithCache(targets, worksheetCache)
	results := make(map[string]string, len(typedResults))
	for cell, value := range typedResults {
		results[cell] = value.Value()
	}
	return results
}

func (f *File) batchCalculatePostgresLookupTargetsWithCache(targets map[string]pgLookupBatchTarget, worksheetCache *WorksheetCache) map[string]formulaArg {
	if len(targets) == 0 {
		return nil
	}

	results := make(map[string]formulaArg, len(targets))
	baseTargets := make(map[string]pgLookupBatchTarget, len(targets))
	conditionalItems := make([]pgBatchConditionalItem, 0)
	for id, target := range targets {
		if target.id != "" && strings.Count(target.id, "!") == 1 {
			if class := getPGLookupOptimizationClass(target.sheet, target.formula); class.whole {
				if cached, ok := f.loadPGWholeCellCache(target.id); ok {
					results[target.id] = cached
					continue
				}
			}
		}
		if conditional, ok := parsePGConditionalLookupFormula(target.sheet, target.formula); ok {
			conditionalItems = append(conditionalItems, pgBatchConditionalItem{
				target:  target,
				formula: conditional,
			})
			continue
		}
		baseTargets[id] = target
	}

	baseResults := f.batchCalculatePostgresBaseLookupArgsWithCache(baseTargets, worksheetCache)
	for cell, value := range baseResults {
		results[cell] = value
		if strings.Count(cell, "!") == 1 {
			f.storePGWholeCellCache(cell, value)
		}
	}
	if len(conditionalItems) == 0 {
		return results
	}

	lookupTargets := make(map[string]pgLookupBatchTarget, len(conditionalItems))
	for _, item := range conditionalItems {
		lookupTargets[item.target.id] = pgLookupBatchTarget{
			id:      item.target.id,
			sheet:   item.target.sheet,
			formula: item.formula.condition.lookupFormula,
		}
	}
	lookupResults := f.batchCalculatePostgresBaseLookupArgsWithCache(lookupTargets, worksheetCache)
	for _, item := range conditionalItems {
		lookupValue, ok := lookupResults[item.target.id]
		if !ok {
			continue
		}
		otherValue, err := f.resolvePGScalarExpr(item.target.sheet, item.formula.condition.otherExpr, worksheetCache)
		if err != nil {
			continue
		}

		lhs, rhs := lookupValue, otherValue
		if !item.formula.condition.lookupOnLeft {
			lhs, rhs = otherValue, lookupValue
		}
		if isPGQuotedStringLiteral(item.formula.condition.otherExpr) {
			if item.formula.condition.lookupOnLeft {
				lhs = newStringFormulaArg(lookupValue.Value())
			} else {
				rhs = newStringFormulaArg(lookupValue.Value())
			}
		}

		matched, err := evalPGCondition(lhs, rhs, item.formula.condition.operator)
		if err != nil {
			continue
		}
		branchExpr := item.formula.whenFalse
		if matched {
			branchExpr = item.formula.whenTrue
		}
		branchValue, err := f.resolvePGScalarExpr(item.target.sheet, branchExpr, worksheetCache)
		if err != nil {
			continue
		}
		results[item.target.id] = branchValue
		if strings.Count(item.target.id, "!") == 1 {
			f.storePGWholeCellCache(item.target.id, branchValue)
		}
	}
	return results
}

func (f *File) batchCalculatePostgresBaseLookupArgsWithCache(targets map[string]pgLookupBatchTarget, worksheetCache *WorksheetCache) map[string]formulaArg {
	db, cfg, ok := f.getPostgresLookupRuntime()
	if !ok || len(targets) == 0 {
		return nil
	}

	ctx := context.Background()
	matchGroups := make(map[string][]pgBatchMatchItem)
	vlookupGroups := make(map[string][]pgBatchVLookupItem)
	indexMatchGroups := make(map[string][]pgBatchIndexMatchItem)

	for _, target := range targets {
		env, ok := parsePGLookupEnvelope(target.formula)
		if !ok {
			continue
		}
		if parsed, ok := parsePGMatchLookupFormula(target.sheet, env.body); ok {
			lookup, err := f.resolvePGLookupExpr(target.sheet, parsed.lookupExpr, worksheetCache)
			if err != nil {
				continue
			}
			key := fmt.Sprintf("match:%s:%d:%d:%d:%d",
				parsed.lookupRange.Sheet,
				parsed.lookupRange.StartRow, parsed.lookupRange.EndRow,
				parsed.lookupRange.StartCol, parsed.lookupRange.EndCol,
			)
			matchGroups[key] = append(matchGroups[key], pgBatchMatchItem{
				fullCell: target.id,
				sheet:    target.sheet,
				env:      env,
				formula:  parsed,
				lookup:   lookup,
			})
			continue
		}
		if parsed, ok := parsePGVLookupFormula(target.sheet, env.body); ok {
			lookup, err := f.resolvePGLookupExpr(target.sheet, parsed.lookupExpr, worksheetCache)
			if err != nil {
				continue
			}
			key := fmt.Sprintf("vlookup:%s:%d:%d:%d:%d:%d",
				parsed.tableRange.Sheet,
				parsed.tableRange.StartRow, parsed.tableRange.EndRow,
				parsed.tableRange.StartCol, parsed.tableRange.EndCol,
				parsed.returnColIdx1,
			)
			vlookupGroups[key] = append(vlookupGroups[key], pgBatchVLookupItem{
				fullCell: target.id,
				sheet:    target.sheet,
				env:      env,
				formula:  parsed,
				lookup:   lookup,
			})
			continue
		}
		if parsed, ok := parsePGIndexMatchFormula(target.sheet, env.body); ok {
			lookup, err := f.resolvePGLookupExpr(target.sheet, parsed.lookupExpr, worksheetCache)
			if err != nil {
				continue
			}
			key := fmt.Sprintf("indexmatch:%s:%d:%d:%d:%d:%s:%d:%d:%d:%d",
				parsed.matchRange.Sheet,
				parsed.matchRange.StartRow, parsed.matchRange.EndRow,
				parsed.matchRange.StartCol, parsed.matchRange.EndCol,
				parsed.returnRange.Sheet,
				parsed.returnRange.StartRow, parsed.returnRange.EndRow,
				parsed.returnRange.StartCol, parsed.returnRange.EndCol,
			)
			indexMatchGroups[key] = append(indexMatchGroups[key], pgBatchIndexMatchItem{
				fullCell: target.id,
				sheet:    target.sheet,
				env:      env,
				formula:  parsed,
				lookup:   lookup,
			})
		}
	}

	results := make(map[string]formulaArg)
	for _, items := range matchGroups {
		groupResults, err := f.evalPGBatchMatchGroup(ctx, db, cfg, items, worksheetCache)
		if err != nil {
			continue
		}
		for cell, value := range groupResults {
			results[cell] = value
		}
	}
	for _, items := range vlookupGroups {
		groupResults, err := f.evalPGBatchVLookupGroup(ctx, db, cfg, items, worksheetCache)
		if err != nil {
			continue
		}
		for cell, value := range groupResults {
			results[cell] = value
		}
	}
	for _, items := range indexMatchGroups {
		groupResults, err := f.evalPGBatchIndexMatchGroup(ctx, db, cfg, items, worksheetCache)
		if err != nil {
			continue
		}
		for cell, value := range groupResults {
			results[cell] = value
		}
	}
	return results
}

func extractFirstPGSupportedLookupExpr(sheet, formula string) string {
	type candidate struct {
		name string
		idx  int
	}

	candidates := make([]candidate, 0, 5)
	for _, name := range []string{"IFERROR", "IFNA", "VLOOKUP", "INDEX", "MATCH"} {
		if idx := findPGFunctionToken(formula, name); idx >= 0 {
			candidates = append(candidates, candidate{name: name, idx: idx})
		}
	}
	if len(candidates) == 0 {
		return ""
	}

	best := candidates[0]
	for _, candidate := range candidates[1:] {
		if candidate.idx < best.idx {
			best = candidate
		}
	}

	expr := extractPGFunctionExprAt(formula, best.idx, best.name)
	if expr == "" {
		return ""
	}
	switch best.name {
	case "IFERROR", "IFNA":
		env, ok := parsePGLookupEnvelope(expr)
		if !ok {
			return ""
		}
		if _, ok := parsePGMatchLookupFormula(sheet, env.body); ok {
			return expr
		}
		if _, ok := parsePGVLookupFormula(sheet, env.body); ok {
			return expr
		}
		if _, ok := parsePGIndexMatchFormula(sheet, env.body); ok {
			return expr
		}
	case "VLOOKUP":
		if _, ok := parsePGVLookupFormula(sheet, expr); ok {
			return expr
		}
	case "INDEX":
		if _, ok := parsePGIndexMatchFormula(sheet, expr); ok {
			return expr
		}
	case "MATCH":
		if _, ok := parsePGMatchLookupFormula(sheet, expr); ok {
			return expr
		}
	}
	return ""
}

func parsePGLookupEnvelope(formula string) (pgLookupEnvelope, bool) {
	body := strings.TrimSpace(strings.TrimPrefix(formula, "="))
	if body == "" {
		return pgLookupEnvelope{}, false
	}

	for _, wrapper := range []string{"IFERROR", "IFNA"} {
		if !hasPGFunctionPrefix(body, wrapper) {
			continue
		}
		args := extractFunctionArgs(body)
		if len(args) != 2 {
			return pgLookupEnvelope{}, false
		}
		return pgLookupEnvelope{
			body:        strings.TrimSpace(args[0]),
			fallback:    strings.TrimSpace(args[1]),
			hasFallback: true,
		}, true
	}

	return pgLookupEnvelope{body: body}, true
}

func parsePGConditionalLookupFormula(defaultSheet, formula string) (*pgConditionalLookupFormula, bool) {
	body := strings.TrimSpace(strings.TrimPrefix(formula, "="))
	if !hasPGFunctionPrefix(body, "IF") {
		return nil, false
	}
	args := extractFunctionArgs(body)
	if len(args) != 3 {
		return nil, false
	}
	condition, ok := parsePGLookupCondition(defaultSheet, args[0])
	if !ok {
		return nil, false
	}
	return &pgConditionalLookupFormula{
		condition: condition,
		whenTrue:  strings.TrimSpace(args[1]),
		whenFalse: strings.TrimSpace(args[2]),
	}, true
}

func (f *File) buildPGWholeFormulaResultCacheKey(defaultSheet, formula string, worksheetCache *WorksheetCache) (string, bool, error) {
	formula = strings.TrimSpace(formula)
	if formula == "" {
		return "", false, nil
	}
	normalizedFormula := strings.TrimSpace(strings.TrimPrefix(formula, "="))
	if conditional, ok := parsePGConditionalLookupFormula(defaultSheet, formula); ok {
		return f.buildPGConditionalLookupResultCacheKey(defaultSheet, normalizedFormula, conditional, worksheetCache)
	}
	env, ok := parsePGLookupEnvelope(formula)
	if !ok {
		return "", false, nil
	}
	return f.buildPGBaseLookupResultCacheKey(defaultSheet, normalizedFormula, env, worksheetCache)
}

func (f *File) buildPGConditionalLookupResultCacheKey(defaultSheet, formula string, conditional *pgConditionalLookupFormula, worksheetCache *WorksheetCache) (string, bool, error) {
	if conditional == nil {
		return "", false, nil
	}
	if !isPGLiteralScalarExpr(conditional.condition.otherExpr) ||
		!isPGLiteralScalarExpr(conditional.whenTrue) ||
		!isPGLiteralScalarExpr(conditional.whenFalse) {
		return "", false, nil
	}
	return f.buildPGBaseLookupResultCacheKey(defaultSheet, formula, pgLookupEnvelope{body: conditional.condition.lookupFormula}, worksheetCache)
}

func (f *File) buildPGBaseLookupResultCacheKey(defaultSheet, formula string, env pgLookupEnvelope, worksheetCache *WorksheetCache) (string, bool, error) {
	if env.hasFallback && !isPGLiteralScalarExpr(env.fallback) {
		return "", false, nil
	}

	var (
		lookupExpr  string
		versionKey  string
		lookupValue formulaArg
		err         error
	)

	switch parsed, ok := parsePGMatchLookupFormula(defaultSheet, env.body); {
	case ok:
		if !isPGSimpleLookupExpr(parsed.lookupExpr) {
			return "", false, nil
		}
		lookupExpr = parsed.lookupExpr
		versionKey = f.pgSourceSheetVersionKey(parsed.lookupRange.Sheet)
	case false:
		if parsed, ok := parsePGVLookupFormula(defaultSheet, env.body); ok {
			if !isPGSimpleLookupExpr(parsed.lookupExpr) {
				return "", false, nil
			}
			lookupExpr = parsed.lookupExpr
			versionKey = f.pgSourceSheetVersionKey(parsed.tableRange.Sheet)
		} else if parsed, ok := parsePGIndexMatchFormula(defaultSheet, env.body); ok {
			if !isPGSimpleLookupExpr(parsed.lookupExpr) {
				return "", false, nil
			}
			lookupExpr = parsed.lookupExpr
			versionKey = f.pgSourceSheetVersionKey(parsed.matchRange.Sheet, parsed.returnRange.Sheet)
		} else {
			return "", false, nil
		}
	}

	lookupValue, err = f.resolvePGScalarExpr(defaultSheet, lookupExpr, worksheetCache)
	if err != nil {
		return "", false, err
	}
	return fmt.Sprintf("pg:formula:%s:%s:%s", formula, pgFormulaArgTypedKey(lookupValue), versionKey), true, nil
}

func (f *File) pgSourceSheetVersionKey(sheets ...string) string {
	seen := make(map[string]struct{}, len(sheets))
	uniq := make([]string, 0, len(sheets))
	for _, sheet := range sheets {
		if sheet == "" {
			continue
		}
		if _, ok := seen[sheet]; ok {
			continue
		}
		seen[sheet] = struct{}{}
		uniq = append(uniq, sheet)
	}
	sort.Strings(uniq)

	parts := make([]string, 0, len(uniq))
	for _, sheet := range uniq {
		parts = append(parts, sheet+"="+strconv.FormatUint(f.getSheetVersion(sheet), 10))
	}
	return strings.Join(parts, "|")
}

func (f *File) loadPGLookupResultCache(key string) (formulaArg, bool) {
	if key == "" {
		return formulaArg{}, false
	}
	cached, ok := f.pgLookupResultCache.Load(key)
	if !ok {
		return formulaArg{}, false
	}
	result, valid := cached.(formulaArg)
	return result, valid
}

func (f *File) storePGLookupResultCache(key string, value formulaArg) {
	if key == "" {
		return
	}
	f.pgLookupResultCache.Store(key, value)
	f.markCalculationSnapshotDirty()
}

func parsePGLookupCondition(defaultSheet, expr string) (pgLookupCondition, bool) {
	lhs, op, rhs, ok := splitPGConditionComparison(expr)
	if !ok {
		return pgLookupCondition{}, false
	}
	if isSupportedPGBaseLookupFormula(defaultSheet, lhs) {
		return pgLookupCondition{
			operator:      op,
			lookupFormula: lhs,
			lookupOnLeft:  true,
			otherExpr:     rhs,
		}, true
	}
	if isSupportedPGBaseLookupFormula(defaultSheet, rhs) {
		return pgLookupCondition{
			operator:      op,
			lookupFormula: rhs,
			lookupOnLeft:  false,
			otherExpr:     lhs,
		}, true
	}
	return pgLookupCondition{}, false
}

func parsePGMatchLookupFormula(defaultSheet, body string) (*pgMatchLookupFormula, bool) {
	if !hasPGFunctionPrefix(body, "MATCH") {
		return nil, false
	}
	args := extractFunctionArgs(body)
	if len(args) != 3 {
		return nil, false
	}
	if !isZeroFormulaArg(args[2]) {
		return nil, false
	}
	lookupRange, ok := parsePGLookupRange(defaultSheet, args[1])
	if !ok || (!lookupRange.isVertical() && !lookupRange.isHorizontal()) {
		return nil, false
	}
	return &pgMatchLookupFormula{
		lookupExpr:  strings.TrimSpace(args[0]),
		lookupRange: lookupRange,
	}, true
}

func parsePGVLookupFormula(defaultSheet, body string) (*pgVLookupFormula, bool) {
	if !hasPGFunctionPrefix(body, "VLOOKUP") {
		return nil, false
	}
	args := extractFunctionArgs(body)
	if len(args) != 4 {
		return nil, false
	}
	colIdx, err := strconv.Atoi(strings.TrimSpace(args[2]))
	if err != nil || colIdx <= 0 {
		return nil, false
	}
	if !isFalseFormulaArg(args[3]) {
		return nil, false
	}
	tableRange, ok := parsePGLookupRange(defaultSheet, args[1])
	if !ok {
		return nil, false
	}
	return &pgVLookupFormula{
		lookupExpr:    strings.TrimSpace(args[0]),
		tableRange:    tableRange,
		returnColIdx1: colIdx,
	}, true
}

func parsePGIndexMatchFormula(defaultSheet, body string) (*pgIndexMatchFormula, bool) {
	if !hasPGFunctionPrefix(body, "INDEX") || !strings.Contains(body, "MATCH(") {
		return nil, false
	}
	args := extractFunctionArgs(body)
	if len(args) < 2 || len(args) > 3 {
		return nil, false
	}
	if len(args) == 3 && !isOneFormulaArg(args[2]) {
		return nil, false
	}
	matchExpr := strings.TrimSpace(args[1])
	if !hasPGFunctionPrefix(matchExpr, "MATCH") {
		return nil, false
	}
	matchArgs := extractFunctionArgs(matchExpr)
	if len(matchArgs) != 3 || !isZeroFormulaArg(matchArgs[2]) {
		return nil, false
	}
	returnRange, ok := parsePGLookupRange(defaultSheet, args[0])
	if !ok || (!returnRange.isVertical() && !returnRange.isHorizontal()) {
		return nil, false
	}
	matchRange, ok := parsePGLookupRange(defaultSheet, matchArgs[1])
	if !ok || (!matchRange.isVertical() && !matchRange.isHorizontal()) {
		return nil, false
	}
	if returnRange.isVertical() != matchRange.isVertical() || returnRange.isHorizontal() != matchRange.isHorizontal() {
		return nil, false
	}
	return &pgIndexMatchFormula{
		lookupExpr:  strings.TrimSpace(matchArgs[0]),
		matchRange:  matchRange,
		returnRange: returnRange,
	}, true
}

func splitPGConditionComparison(expr string) (lhs, op, rhs string, ok bool) {
	expr = strings.TrimSpace(expr)
	depth := 0
	inString := false
	for i := 0; i < len(expr); i++ {
		switch expr[i] {
		case '"':
			if i+1 < len(expr) && expr[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '<', '>', '=':
			if inString || depth != 0 {
				continue
			}
			if i+1 < len(expr) {
				if expr[i] == '<' && expr[i+1] == '>' {
					return strings.TrimSpace(expr[:i]), "<>", strings.TrimSpace(expr[i+2:]), true
				}
				if expr[i+1] == '=' {
					return strings.TrimSpace(expr[:i]), expr[i : i+2], strings.TrimSpace(expr[i+2:]), true
				}
			}
			return strings.TrimSpace(expr[:i]), expr[i : i+1], strings.TrimSpace(expr[i+1:]), true
		}
	}
	return "", "", "", false
}

func parsePGLookupRange(defaultSheet, expr string) (pgLookupRange, bool) {
	expr = strings.TrimSpace(strings.ReplaceAll(expr, "$", ""))
	parts := strings.Split(expr, ":")
	if len(parts) == 0 || len(parts) > 2 {
		return pgLookupRange{}, false
	}

	var cr cellRange
	for i, part := range parts {
		ref, col, row, err := parseRef(strings.TrimSpace(part))
		if err != nil {
			return pgLookupRange{}, false
		}
		ref.Sheet = strings.Trim(ref.Sheet, "'")
		if i == 0 && ref.Sheet == "" {
			ref.Sheet = defaultSheet
		}
		if i == 0 {
			if col {
				ref.Row = 1
			}
			if row {
				ref.Col = 1
			}
			cr.From, cr.To = ref, ref
			if len(parts) == 1 {
				if col {
					cr.To.Row = TotalRows
				}
				if row {
					cr.To.Col = MaxColumns
				}
			}
			continue
		}
		if err := cr.prepareCellRange(col, row, ref); err != nil {
			return pgLookupRange{}, false
		}
	}

	if cr.From.Col > cr.To.Col {
		cr.From.Col, cr.To.Col = cr.To.Col, cr.From.Col
	}
	if cr.From.Row > cr.To.Row {
		cr.From.Row, cr.To.Row = cr.To.Row, cr.From.Row
	}

	return pgLookupRange{
		Sheet:    cr.From.Sheet,
		StartRow: cr.From.Row,
		EndRow:   cr.To.Row,
		StartCol: cr.From.Col,
		EndCol:   cr.To.Col,
	}, true
}

func (f *File) evalPGMatchFormula(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, formula *pgMatchLookupFormula, env pgLookupEnvelope, worksheetCache *WorksheetCache) (formulaArg, error) {
	lookupValue, err := f.resolvePGLookupExpr(sheet, formula.lookupExpr, worksheetCache)
	if err != nil {
		return formulaArg{}, err
	}
	position, _, _, found, err := f.queryPGExactMatch(ctx, db, cfg, formula.lookupRange, lookupValue)
	if err != nil {
		return formulaArg{}, err
	}
	if !found {
		return f.pgLookupFallbackOrNA(sheet, env, worksheetCache)
	}
	return newNumberFormulaArg(float64(position)), nil
}

func (f *File) evalPGVLookupFormula(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, formula *pgVLookupFormula, env pgLookupEnvelope, worksheetCache *WorksheetCache) (formulaArg, error) {
	lookupValue, err := f.resolvePGLookupExpr(sheet, formula.lookupExpr, worksheetCache)
	if err != nil {
		return formulaArg{}, err
	}
	if lookupValue.Type == ArgEmpty {
		lookupValue = newNumberFormulaArg(0)
	}

	position, rowNum, _, found, err := f.queryPGExactMatch(ctx, db, cfg, pgLookupRange{
		Sheet:    formula.tableRange.Sheet,
		StartRow: formula.tableRange.StartRow,
		EndRow:   formula.tableRange.EndRow,
		StartCol: formula.tableRange.StartCol,
		EndCol:   formula.tableRange.StartCol,
	}, lookupValue)
	_ = position
	if err != nil {
		return formulaArg{}, err
	}
	if !found {
		return f.pgLookupFallbackOrNA(sheet, env, worksheetCache)
	}

	targetCol := formula.tableRange.StartCol + formula.returnColIdx1 - 1
	if targetCol < formula.tableRange.StartCol || targetCol > formula.tableRange.EndCol {
		return newErrorFormulaArg(formulaErrorNA, "VLOOKUP has invalid column index"), fmt.Errorf("VLOOKUP has invalid column index")
	}
	value, err := f.queryPGCellValue(ctx, db, cfg, formula.tableRange.Sheet, rowNum, targetCol)
	if err != nil {
		return formulaArg{}, err
	}
	return value.formulaArg(), nil
}

func (f *File) evalPGIndexMatchFormula(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, formula *pgIndexMatchFormula, env pgLookupEnvelope, worksheetCache *WorksheetCache) (formulaArg, error) {
	lookupValue, err := f.resolvePGLookupExpr(sheet, formula.lookupExpr, worksheetCache)
	if err != nil {
		return formulaArg{}, err
	}
	position, _, _, found, err := f.queryPGExactMatch(ctx, db, cfg, formula.matchRange, lookupValue)
	if err != nil {
		return formulaArg{}, err
	}
	if !found {
		return f.pgLookupFallbackOrNA(sheet, env, worksheetCache)
	}

	targetRow := formula.returnRange.StartRow
	targetCol := formula.returnRange.StartCol
	if formula.returnRange.isVertical() {
		targetRow = formula.returnRange.StartRow + position - 1
		if targetRow > formula.returnRange.EndRow {
			return newErrorFormulaArg(formulaErrorREF, "INDEX row_num out of range"), fmt.Errorf("INDEX row_num out of range")
		}
	} else {
		targetCol = formula.returnRange.StartCol + position - 1
		if targetCol > formula.returnRange.EndCol {
			return newErrorFormulaArg(formulaErrorREF, "INDEX col_num out of range"), fmt.Errorf("INDEX col_num out of range")
		}
	}
	value, err := f.queryPGCellValue(ctx, db, cfg, formula.returnRange.Sheet, targetRow, targetCol)
	if err != nil {
		return formulaArg{}, err
	}
	return value.formulaArg(), nil
}

func (f *File) evalPGConditionalLookupFormula(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, formula *pgConditionalLookupFormula, worksheetCache *WorksheetCache) (formulaArg, error) {
	lookupValue, matched, err := f.tryCalcPostgresLookupFormula(sheet, "", formula.condition.lookupFormula, worksheetCache)
	if err != nil {
		return formulaArg{}, err
	}
	if !matched {
		return formulaArg{}, fmt.Errorf("unsupported lookup condition %q", formula.condition.lookupFormula)
	}

	otherValue, err := f.resolvePGScalarExpr(sheet, formula.condition.otherExpr, worksheetCache)
	if err != nil {
		return formulaArg{}, err
	}

	lhs, rhs := lookupValue, otherValue
	if !formula.condition.lookupOnLeft {
		lhs, rhs = otherValue, lookupValue
	}
	if isPGQuotedStringLiteral(formula.condition.otherExpr) {
		if formula.condition.lookupOnLeft {
			lhs = newStringFormulaArg(lookupValue.Value())
		} else {
			rhs = newStringFormulaArg(lookupValue.Value())
		}
	}

	conditionMet, err := evalPGCondition(lhs, rhs, formula.condition.operator)
	if err != nil {
		return formulaArg{}, err
	}
	if conditionMet {
		return f.resolvePGScalarExpr(sheet, formula.whenTrue, worksheetCache)
	}
	return f.resolvePGScalarExpr(sheet, formula.whenFalse, worksheetCache)
}

func (f *File) resolvePGLookupExpr(defaultSheet, expr string, worksheetCache *WorksheetCache) (formulaArg, error) {
	return f.resolvePGScalarExpr(defaultSheet, expr, worksheetCache)
}

func (f *File) resolvePGScalarExpr(defaultSheet, expr string, worksheetCache *WorksheetCache) (formulaArg, error) {
	expr = strings.TrimSpace(expr)
	if len(expr) >= 2 && strings.HasPrefix(expr, `"`) && strings.HasSuffix(expr, `"`) {
		return newStringFormulaArg(strings.ReplaceAll(strings.Trim(expr, `"`), `""`, `"`)), nil
	}
	upper := strings.ToUpper(expr)
	if upper == "TRUE" {
		return newBoolFormulaArg(true), nil
	}
	if upper == "FALSE" {
		return newBoolFormulaArg(false), nil
	}
	if num, err := strconv.ParseFloat(expr, 64); err == nil {
		return newNumberFormulaArg(num), nil
	}
	if env, ok := parsePGLookupEnvelope(expr); ok {
		if db, cfg, enabled := f.getPostgresLookupRuntime(); enabled {
			if result, matched, err := f.evalPGSupportedLookupEnvelope(context.Background(), db, cfg, defaultSheet, env, worksheetCache); matched || err != nil {
				return result, err
			}
		}
	}

	ref, _, _, err := parseRef(strings.ReplaceAll(expr, "$", ""))
	if err != nil {
		return formulaArg{}, fmt.Errorf("unsupported lookup expression %q", expr)
	}
	ref.Sheet = strings.Trim(ref.Sheet, "'")
	if ref.Sheet == "" {
		ref.Sheet = defaultSheet
	}
	cellName, err := CoordinatesToCellName(ref.Col, ref.Row)
	if err != nil {
		return formulaArg{}, err
	}
	if worksheetCache != nil {
		if cachedArg, ok := worksheetCache.Get(ref.Sheet, cellName); ok {
			return cachedArg, nil
		}
	}
	cacheKey := ref.Sheet + "!" + cellName
	if cached, ok := f.calcCache.Load(cacheKey); ok {
		if cachedArg, valid := cached.(formulaArg); valid {
			return cachedArg, nil
		}
	}
	value, err := f.GetCellValue(ref.Sheet, cellName, Options{RawCellValue: true})
	if err != nil {
		return formulaArg{}, err
	}
	cellType, err := f.GetCellType(ref.Sheet, cellName)
	if err != nil {
		return formulaArg{}, err
	}
	return inferCellValueType(value, cellType), nil
}

func (f *File) pgLookupFallbackOrNA(defaultSheet string, env pgLookupEnvelope, worksheetCache *WorksheetCache) (formulaArg, error) {
	if !env.hasFallback {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA), fmt.Errorf(formulaErrorNA)
	}
	fallback, err := f.resolvePGFallbackExpr(defaultSheet, env.fallback, worksheetCache)
	if err != nil {
		return formulaArg{}, err
	}
	return fallback, nil
}

func (f *File) resolvePGFallbackExpr(defaultSheet, expr string, worksheetCache *WorksheetCache) (formulaArg, error) {
	return f.resolvePGScalarExpr(defaultSheet, expr, worksheetCache)
}

func (f *File) evalPGBatchMatchGroup(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, items []pgBatchMatchItem, worksheetCache *WorksheetCache) (map[string]formulaArg, error) {
	results := make(map[string]formulaArg, len(items))
	if len(items) == 0 {
		return results, nil
	}

	batchMatches, err := f.queryPGExactMatchesBatch(ctx, db, cfg, items[0].formula.lookupRange, collectPGLookupKeysFromMatch(items))
	if err != nil {
		return nil, err
	}
	for _, item := range items {
		match := batchMatches[pgFormulaArgTypedKey(item.lookup)]
		if !match.Found {
			fallback, fallbackErr := f.pgLookupFallbackOrNA(item.sheet, item.env, worksheetCache)
			if fallbackErr != nil {
				return nil, fallbackErr
			}
			results[item.fullCell] = fallback
			continue
		}
		results[item.fullCell] = newNumberFormulaArg(float64(match.Position))
	}
	return results, nil
}

func (f *File) evalPGBatchVLookupGroup(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, items []pgBatchVLookupItem, worksheetCache *WorksheetCache) (map[string]formulaArg, error) {
	results := make(map[string]formulaArg, len(items))
	if len(items) == 0 {
		return results, nil
	}

	first := items[0].formula
	targetCol := first.tableRange.StartCol + first.returnColIdx1 - 1
	if targetCol < first.tableRange.StartCol || targetCol > first.tableRange.EndCol {
		return nil, fmt.Errorf("VLOOKUP has invalid column index")
	}

	matches, err := f.queryPGExactMatchesBatch(ctx, db, cfg, pgLookupRange{
		Sheet:    first.tableRange.Sheet,
		StartRow: first.tableRange.StartRow,
		EndRow:   first.tableRange.EndRow,
		StartCol: first.tableRange.StartCol,
		EndCol:   first.tableRange.StartCol,
	}, collectPGLookupKeysFromVLookup(items))
	if err != nil {
		return nil, err
	}

	rowsNeeded := make([]int, 0, len(items))
	seenRows := make(map[int]struct{})
	for _, item := range items {
		match := matches[pgFormulaArgTypedKey(item.lookup)]
		if match.Found {
			if _, ok := seenRows[match.RowNum]; !ok {
				seenRows[match.RowNum] = struct{}{}
				rowsNeeded = append(rowsNeeded, match.RowNum)
			}
		}
	}
	valueMap, err := f.queryPGColumnCellsBatch(ctx, db, cfg, first.tableRange.Sheet, targetCol, rowsNeeded)
	if err != nil {
		return nil, err
	}

	for _, item := range items {
		match := matches[pgFormulaArgTypedKey(item.lookup)]
		if !match.Found {
			fallback, fallbackErr := f.pgLookupFallbackOrNA(item.sheet, item.env, worksheetCache)
			if fallbackErr != nil {
				return nil, fallbackErr
			}
			results[item.fullCell] = fallback
			continue
		}
		results[item.fullCell] = valueMap[match.RowNum].formulaArg()
	}
	return results, nil
}

func (f *File) evalPGBatchIndexMatchGroup(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, items []pgBatchIndexMatchItem, worksheetCache *WorksheetCache) (map[string]formulaArg, error) {
	results := make(map[string]formulaArg, len(items))
	if len(items) == 0 {
		return results, nil
	}

	first := items[0].formula
	matches, err := f.queryPGExactMatchesBatch(ctx, db, cfg, first.matchRange, collectPGLookupKeysFromIndexMatch(items))
	if err != nil {
		return nil, err
	}

	if first.returnRange.isVertical() {
		rowsNeeded := make([]int, 0, len(items))
		seenRows := make(map[int]struct{})
		for _, item := range items {
			match := matches[pgFormulaArgTypedKey(item.lookup)]
			if !match.Found {
				continue
			}
			targetRow := item.formula.returnRange.StartRow + match.Position - 1
			if targetRow > item.formula.returnRange.EndRow {
				return nil, fmt.Errorf("INDEX row_num out of range")
			}
			if _, ok := seenRows[targetRow]; !ok {
				seenRows[targetRow] = struct{}{}
				rowsNeeded = append(rowsNeeded, targetRow)
			}
		}
		valueMap, err := f.queryPGColumnCellsBatch(ctx, db, cfg, first.returnRange.Sheet, first.returnRange.StartCol, rowsNeeded)
		if err != nil {
			return nil, err
		}
		for _, item := range items {
			match := matches[pgFormulaArgTypedKey(item.lookup)]
			if !match.Found {
				fallback, fallbackErr := f.pgLookupFallbackOrNA(item.sheet, item.env, worksheetCache)
				if fallbackErr != nil {
					return nil, fallbackErr
				}
				results[item.fullCell] = fallback
				continue
			}
			targetRow := item.formula.returnRange.StartRow + match.Position - 1
			results[item.fullCell] = valueMap[targetRow].formulaArg()
		}
		return results, nil
	}

	colsNeeded := make([]int, 0, len(items))
	seenCols := make(map[int]struct{})
	for _, item := range items {
		match := matches[pgFormulaArgTypedKey(item.lookup)]
		if !match.Found {
			continue
		}
		targetCol := item.formula.returnRange.StartCol + match.Position - 1
		if targetCol > item.formula.returnRange.EndCol {
			return nil, fmt.Errorf("INDEX col_num out of range")
		}
		if _, ok := seenCols[targetCol]; !ok {
			seenCols[targetCol] = struct{}{}
			colsNeeded = append(colsNeeded, targetCol)
		}
	}
	valueMap, err := f.queryPGRowCellsBatch(ctx, db, cfg, first.returnRange.Sheet, first.returnRange.StartRow, colsNeeded)
	if err != nil {
		return nil, err
	}
	for _, item := range items {
		match := matches[pgFormulaArgTypedKey(item.lookup)]
		if !match.Found {
			fallback, fallbackErr := f.pgLookupFallbackOrNA(item.sheet, item.env, worksheetCache)
			if fallbackErr != nil {
				return nil, fallbackErr
			}
			results[item.fullCell] = fallback
			continue
		}
		targetCol := item.formula.returnRange.StartCol + match.Position - 1
		results[item.fullCell] = valueMap[targetCol].formulaArg()
	}
	return results, nil
}

func (f *File) queryPGExactMatch(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, rng pgLookupRange, lookupValue formulaArg) (position int, rowNum int, colNum int, found bool, err error) {
	tables := getPGMirrorTables(cfg.TablePrefix)
	lookup := pgFormulaArgValue(lookupValue)
	lookupType := pgMirrorValueTypeFromFormulaArg(lookupValue)
	version := f.getSheetVersion(rng.Sheet)
	cacheKey := pgExactMatchCacheKey(cfg, rng, version, lookupType, lookup)
	if cached, ok := f.pgCalcCache.Load(cacheKey); ok {
		if result, valid := cached.(pgExactMatchResult); valid {
			if pgExactMatchCacheTestHook != nil {
				pgExactMatchCacheTestHook(cacheKey, true, true, fmt.Sprintf("%T", cached))
			}
			return result.Position, result.RowNum, result.ColNum, result.Found, nil
		}
		if pgExactMatchCacheTestHook != nil {
			pgExactMatchCacheTestHook(cacheKey, false, true, fmt.Sprintf("%T", cached))
		}
	} else if pgExactMatchCacheTestHook != nil {
		pgExactMatchCacheTestHook(cacheKey, false, false, "")
	}

	if rng.isVertical() {
		sqlStr := fmt.Sprintf(`
SELECT row_num
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND col_num = $3
  AND row_num BETWEEN $4 AND $5 AND cell_value_type = $6 AND cell_value = $7
ORDER BY row_num
LIMIT 1`, qualifyPGTable(cfg.Schema, tables.lookup))
		if err = db.QueryRowContext(ctx, sqlStr, cfg.WorkbookID, rng.Sheet, rng.StartCol, rng.StartRow, rng.EndRow, lookupType, lookup).Scan(&rowNum); err != nil {
			if err == sql.ErrNoRows {
				f.storePGCalcCache(cacheKey, pgExactMatchResult{})
				return 0, 0, 0, false, nil
			}
			return 0, 0, 0, false, err
		}
		result := pgExactMatchResult{
			Position: rowNum - rng.StartRow + 1,
			RowNum:   rowNum,
			ColNum:   rng.StartCol,
			Found:    true,
		}
		f.storePGCalcCache(cacheKey, result)
		return result.Position, result.RowNum, result.ColNum, result.Found, nil
	}

	sqlStr := fmt.Sprintf(`
SELECT col_num
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND row_num = $3
  AND col_num BETWEEN $4 AND $5 AND cell_value_type = $6 AND cell_value = $7
ORDER BY col_num
LIMIT 1`, qualifyPGTable(cfg.Schema, tables.lookup))
	if err = db.QueryRowContext(ctx, sqlStr, cfg.WorkbookID, rng.Sheet, rng.StartRow, rng.StartCol, rng.EndCol, lookupType, lookup).Scan(&colNum); err != nil {
		if err == sql.ErrNoRows {
			f.storePGCalcCache(cacheKey, pgExactMatchResult{})
			return 0, 0, 0, false, nil
		}
		return 0, 0, 0, false, err
	}
	result := pgExactMatchResult{
		Position: colNum - rng.StartCol + 1,
		RowNum:   rng.StartRow,
		ColNum:   colNum,
		Found:    true,
	}
	f.storePGCalcCache(cacheKey, result)
	return result.Position, result.RowNum, result.ColNum, result.Found, nil
}

func (f *File) queryPGExactMatchesBatch(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, rng pgLookupRange, lookupValues []formulaArg) (map[string]pgExactMatchResult, error) {
	results := make(map[string]pgExactMatchResult, len(lookupValues))
	version := f.getSheetVersion(rng.Sheet)
	type typedLookup struct {
		key       string
		value     string
		valueType string
	}
	missesByType := make(map[string][]typedLookup)
	seen := make(map[string]struct{}, len(lookupValues))
	for _, lookupArg := range lookupValues {
		lookup := pgFormulaArgValue(lookupArg)
		lookupType := pgMirrorValueTypeFromFormulaArg(lookupArg)
		key := pgFormulaArgTypedKey(lookupArg)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}
		cacheKey := pgExactMatchCacheKey(cfg, rng, version, lookupType, lookup)
		if cached, ok := f.pgCalcCache.Load(cacheKey); ok {
			if result, valid := cached.(pgExactMatchResult); valid {
				results[key] = result
				continue
			}
		}
		missesByType[lookupType] = append(missesByType[lookupType], typedLookup{
			key:       key,
			value:     lookup,
			valueType: lookupType,
		})
	}
	if len(missesByType) == 0 {
		return results, nil
	}

	tables := getPGMirrorTables(cfg.TablePrefix)
	for lookupType, misses := range missesByType {
		args := []interface{}{cfg.WorkbookID, rng.Sheet}
		placeholders := make([]string, 0, len(misses))
		if rng.isVertical() {
			args = append(args, rng.StartCol, rng.StartRow, rng.EndRow, lookupType)
			for i, lookup := range misses {
				args = append(args, lookup.value)
				placeholders = append(placeholders, fmt.Sprintf("$%d", i+7))
			}
			sqlStr := fmt.Sprintf(`
SELECT cell_value, MIN(row_num) AS row_num
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND col_num = $3
  AND row_num BETWEEN $4 AND $5 AND cell_value_type = $6 AND cell_value IN (%s)
GROUP BY cell_value`, qualifyPGTable(cfg.Schema, tables.lookup), strings.Join(placeholders, ","))
			rows, err := db.QueryContext(ctx, sqlStr, args...)
			if err != nil {
				return nil, err
			}
			defer func() { _ = rows.Close() }()

			for rows.Next() {
				var (
					lookup string
					rowNum int
				)
				if err := rows.Scan(&lookup, &rowNum); err != nil {
					return nil, err
				}
				result := pgExactMatchResult{
					Position: rowNum - rng.StartRow + 1,
					RowNum:   rowNum,
					ColNum:   rng.StartCol,
					Found:    true,
				}
				typedKey := lookupType + "\x00" + lookup
				results[typedKey] = result
				f.storePGCalcCache(pgExactMatchCacheKey(cfg, rng, version, lookupType, lookup), result)
			}
			if err := rows.Err(); err != nil {
				return nil, err
			}
		} else {
			args = append(args, rng.StartRow, rng.StartCol, rng.EndCol, lookupType)
			for i, lookup := range misses {
				args = append(args, lookup.value)
				placeholders = append(placeholders, fmt.Sprintf("$%d", i+7))
			}
			sqlStr := fmt.Sprintf(`
SELECT cell_value, MIN(col_num) AS col_num
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND row_num = $3
  AND col_num BETWEEN $4 AND $5 AND cell_value_type = $6 AND cell_value IN (%s)
GROUP BY cell_value`, qualifyPGTable(cfg.Schema, tables.lookup), strings.Join(placeholders, ","))
			rows, err := db.QueryContext(ctx, sqlStr, args...)
			if err != nil {
				return nil, err
			}
			defer func() { _ = rows.Close() }()

			for rows.Next() {
				var (
					lookup string
					colNum int
				)
				if err := rows.Scan(&lookup, &colNum); err != nil {
					return nil, err
				}
				result := pgExactMatchResult{
					Position: colNum - rng.StartCol + 1,
					RowNum:   rng.StartRow,
					ColNum:   colNum,
					Found:    true,
				}
				typedKey := lookupType + "\x00" + lookup
				results[typedKey] = result
				f.storePGCalcCache(pgExactMatchCacheKey(cfg, rng, version, lookupType, lookup), result)
			}
			if err := rows.Err(); err != nil {
				return nil, err
			}
		}
		for _, lookup := range misses {
			if _, ok := results[lookup.key]; ok {
				continue
			}
			f.storePGCalcCache(pgExactMatchCacheKey(cfg, rng, version, lookup.valueType, lookup.value), pgExactMatchResult{})
			results[lookup.key] = pgExactMatchResult{}
		}
	}
	return results, nil
}

func (f *File) queryPGCellValue(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, rowNum, colNum int) (pgMirrorCellValue, error) {
	tables := getPGMirrorTables(cfg.TablePrefix)
	version := f.getSheetVersion(sheet)
	cacheKey := pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum)
	if cached, ok := f.pgCalcCache.Load(cacheKey); ok {
		if value, valid := cached.(pgMirrorCellValue); valid {
			if pgCellValueCacheTestHook != nil {
				pgCellValueCacheTestHook(cacheKey, true, true, fmt.Sprintf("%T", cached))
			}
			return value, nil
		}
		if pgCellValueCacheTestHook != nil {
			pgCellValueCacheTestHook(cacheKey, false, true, fmt.Sprintf("%T", cached))
		}
	} else if pgCellValueCacheTestHook != nil {
		pgCellValueCacheTestHook(cacheKey, false, false, "")
	}
	sqlStr := fmt.Sprintf(`
SELECT value, value_type
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND row_num = $3 AND col_num = $4
LIMIT 1`, qualifyPGTable(cfg.Schema, tables.cell))
	var value pgMirrorCellValue
	err := db.QueryRowContext(ctx, sqlStr, cfg.WorkbookID, sheet, rowNum, colNum).Scan(&value.Value, &value.ValueType)
	if err == sql.ErrNoRows {
		value = pgMirrorCellValue{ValueType: pgMirrorValueTypeBlank}
		f.storePGCalcCache(cacheKey, value)
		return value, nil
	}
	if err == nil {
		f.storePGCalcCache(cacheKey, value)
	}
	return value, err
}

func (f *File) queryPGColumnCellsBatch(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, colNum int, rowNums []int) (map[int]pgMirrorCellValue, error) {
	results := make(map[int]pgMirrorCellValue, len(rowNums))
	version := f.getSheetVersion(sheet)
	misses := make([]int, 0, len(rowNums))
	seen := make(map[int]struct{}, len(rowNums))
	for _, rowNum := range rowNums {
		if _, ok := seen[rowNum]; ok {
			continue
		}
		seen[rowNum] = struct{}{}
		cacheKey := pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum)
		if cached, ok := f.pgCalcCache.Load(cacheKey); ok {
			if value, valid := cached.(pgMirrorCellValue); valid {
				results[rowNum] = value
				continue
			}
		}
		misses = append(misses, rowNum)
	}
	if len(misses) == 0 {
		return results, nil
	}

	tables := getPGMirrorTables(cfg.TablePrefix)
	args := []interface{}{cfg.WorkbookID, sheet, colNum}
	placeholders := make([]string, 0, len(misses))
	for i, rowNum := range misses {
		args = append(args, rowNum)
		placeholders = append(placeholders, fmt.Sprintf("$%d", i+4))
	}
	sqlStr := fmt.Sprintf(`
SELECT row_num, value, value_type
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND col_num = $3
  AND row_num IN (%s)`, qualifyPGTable(cfg.Schema, tables.cell), strings.Join(placeholders, ","))
	rows, err := db.QueryContext(ctx, sqlStr, args...)
	if err != nil {
		return nil, err
	}
	defer func() { _ = rows.Close() }()

	foundRows := make(map[int]struct{}, len(misses))
	for rows.Next() {
		var (
			rowNum    int
			value     string
			valueType string
		)
		if err := rows.Scan(&rowNum, &value, &valueType); err != nil {
			return nil, err
		}
		results[rowNum] = pgMirrorCellValue{Value: value, ValueType: valueType}
		foundRows[rowNum] = struct{}{}
		f.storePGCalcCache(pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum), results[rowNum])
	}
	if err := rows.Err(); err != nil {
		return nil, err
	}
	for _, rowNum := range misses {
		if _, ok := foundRows[rowNum]; ok {
			continue
		}
		results[rowNum] = pgMirrorCellValue{ValueType: pgMirrorValueTypeBlank}
		f.storePGCalcCache(pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum), results[rowNum])
	}
	return results, nil
}

func (f *File) queryPGRowCellsBatch(ctx context.Context, db *sql.DB, cfg *PGSyncOptions, sheet string, rowNum int, colNums []int) (map[int]pgMirrorCellValue, error) {
	results := make(map[int]pgMirrorCellValue, len(colNums))
	version := f.getSheetVersion(sheet)
	misses := make([]int, 0, len(colNums))
	seen := make(map[int]struct{}, len(colNums))
	for _, colNum := range colNums {
		if _, ok := seen[colNum]; ok {
			continue
		}
		seen[colNum] = struct{}{}
		cacheKey := pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum)
		if cached, ok := f.pgCalcCache.Load(cacheKey); ok {
			if value, valid := cached.(pgMirrorCellValue); valid {
				results[colNum] = value
				continue
			}
		}
		misses = append(misses, colNum)
	}
	if len(misses) == 0 {
		return results, nil
	}

	tables := getPGMirrorTables(cfg.TablePrefix)
	args := []interface{}{cfg.WorkbookID, sheet, rowNum}
	placeholders := make([]string, 0, len(misses))
	for i, colNum := range misses {
		args = append(args, colNum)
		placeholders = append(placeholders, fmt.Sprintf("$%d", i+4))
	}
	sqlStr := fmt.Sprintf(`
SELECT col_num, value, value_type
FROM %s
WHERE workbook_id = $1 AND sheet_name = $2 AND row_num = $3
  AND col_num IN (%s)`, qualifyPGTable(cfg.Schema, tables.cell), strings.Join(placeholders, ","))
	rows, err := db.QueryContext(ctx, sqlStr, args...)
	if err != nil {
		return nil, err
	}
	defer func() { _ = rows.Close() }()

	foundCols := make(map[int]struct{}, len(misses))
	for rows.Next() {
		var (
			colNum    int
			value     string
			valueType string
		)
		if err := rows.Scan(&colNum, &value, &valueType); err != nil {
			return nil, err
		}
		results[colNum] = pgMirrorCellValue{Value: value, ValueType: valueType}
		foundCols[colNum] = struct{}{}
		f.storePGCalcCache(pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum), results[colNum])
	}
	if err := rows.Err(); err != nil {
		return nil, err
	}
	for _, colNum := range misses {
		if _, ok := foundCols[colNum]; ok {
			continue
		}
		results[colNum] = pgMirrorCellValue{ValueType: pgMirrorValueTypeBlank}
		f.storePGCalcCache(pgCellValueCacheKey(cfg, sheet, version, rowNum, colNum), results[colNum])
	}
	return results, nil
}

func pgFormulaArgValue(arg formulaArg) string {
	value := arg.Value()
	if arg.Type == ArgNumber && !arg.Boolean && strings.ContainsAny(value, "eE") {
		return strconv.FormatFloat(arg.Number, 'f', -1, 64)
	}
	return value
}

func isZeroFormulaArg(expr string) bool {
	expr = strings.TrimSpace(expr)
	return expr == "0" || expr == "0.0"
}

func isOneFormulaArg(expr string) bool {
	expr = strings.TrimSpace(expr)
	return expr == "1" || expr == "1.0"
}

func isFalseFormulaArg(expr string) bool {
	expr = strings.ToUpper(strings.TrimSpace(expr))
	return expr == "FALSE" || expr == "0"
}

func hasPGFunctionPrefix(expr, funcName string) bool {
	expr = strings.TrimSpace(expr)
	return strings.HasPrefix(expr, funcName+"(")
}

func findPGFunctionToken(expr, funcName string) int {
	token := funcName + "("
	searchFrom := 0
	for {
		idx := strings.Index(expr[searchFrom:], token)
		if idx < 0 {
			return -1
		}
		idx += searchFrom
		if idx == 0 {
			return idx
		}
		prev := expr[idx-1]
		if !((prev >= 'A' && prev <= 'Z') || (prev >= 'a' && prev <= 'z') || (prev >= '0' && prev <= '9') || prev == '_') {
			return idx
		}
		searchFrom = idx + len(funcName)
	}
}

func extractPGFunctionExprAt(expr string, idx int, funcName string) string {
	token := funcName + "("
	if idx < 0 || idx+len(token) > len(expr) || expr[idx:idx+len(token)] != token {
		return ""
	}
	start := idx + len(funcName)
	depth := 0
	inQuote := false
	for i := start; i < len(expr); i++ {
		switch expr[i] {
		case '"', '\'':
			inQuote = !inQuote
		case '(':
			if !inQuote {
				depth++
			}
		case ')':
			if !inQuote {
				depth--
				if depth == 0 {
					return expr[idx : i+1]
				}
			}
		}
	}
	return ""
}

func pgExactMatchCacheKey(cfg *PGSyncOptions, rng pgLookupRange, version uint64, lookupType, lookup string) string {
	return fmt.Sprintf("pg:match:%s:%s:%s:%d:%d:%d:%d:%s:%d:%s",
		cfg.Schema, cfg.TablePrefix, cfg.WorkbookID,
		rng.StartRow, rng.EndRow, rng.StartCol, rng.EndCol, rng.Sheet, version, lookupType+"="+lookup)
}

func pgCellValueCacheKey(cfg *PGSyncOptions, sheet string, version uint64, rowNum, colNum int) string {
	return fmt.Sprintf("pg:cell:%s:%s:%s:%s:%d:%d:%d",
		cfg.Schema, cfg.TablePrefix, cfg.WorkbookID, sheet, version, rowNum, colNum)
}

func isPGQuotedStringLiteral(expr string) bool {
	expr = strings.TrimSpace(expr)
	return len(expr) >= 2 && strings.HasPrefix(expr, `"`) && strings.HasSuffix(expr, `"`)
}

func isPGLiteralScalarExpr(expr string) bool {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return false
	}
	if isPGQuotedStringLiteral(expr) {
		return true
	}
	upper := strings.ToUpper(expr)
	if upper == "TRUE" || upper == "FALSE" {
		return true
	}
	_, err := strconv.ParseFloat(expr, 64)
	return err == nil
}

func isPGSimpleLookupExpr(expr string) bool {
	expr = strings.TrimSpace(expr)
	if isPGLiteralScalarExpr(expr) {
		return true
	}
	_, _, _, err := parseRef(strings.ReplaceAll(expr, "$", ""))
	return err == nil
}

func evalPGCondition(lhs, rhs formulaArg, operator string) (bool, error) {
	if lhs.Type == ArgError {
		return false, errors.New(lhs.String)
	}
	if rhs.Type == ArgError {
		return false, errors.New(rhs.String)
	}
	cmp := compareFormulaArg(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false)
	switch operator {
	case "=":
		return cmp == criteriaEq, nil
	case "<>":
		return cmp != criteriaEq, nil
	case "<":
		return cmp == criteriaL, nil
	case "<=":
		return cmp == criteriaL || cmp == criteriaEq, nil
	case ">":
		return cmp == criteriaG, nil
	case ">=":
		return cmp == criteriaG || cmp == criteriaEq, nil
	default:
		return false, fmt.Errorf("unsupported operator %q", operator)
	}
}

func collectPGLookupKeysFromMatch(items []pgBatchMatchItem) []formulaArg {
	keys := make([]formulaArg, 0, len(items))
	for _, item := range items {
		keys = append(keys, item.lookup)
	}
	return keys
}

func collectPGLookupKeysFromVLookup(items []pgBatchVLookupItem) []formulaArg {
	keys := make([]formulaArg, 0, len(items))
	for _, item := range items {
		keys = append(keys, item.lookup)
	}
	return keys
}

func collectPGLookupKeysFromIndexMatch(items []pgBatchIndexMatchItem) []formulaArg {
	keys := make([]formulaArg, 0, len(items))
	for _, item := range items {
		keys = append(keys, item.lookup)
	}
	return keys
}
