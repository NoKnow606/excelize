package excelize

import (
	"fmt"
	"strings"
	"sync"
)

// SubExpressionCache stores pre-calculated sub-expression results
// Key format: "SUMIFS_expression" -> value
type SubExpressionCache struct {
	mu    sync.RWMutex
	cache map[string]formulaArg
}

// NewSubExpressionCache creates a new sub-expression cache
func NewSubExpressionCache() *SubExpressionCache {
	return &SubExpressionCache{
		cache: make(map[string]formulaArg),
	}
}

// Store saves a sub-expression result
func (c *SubExpressionCache) Store(expr string, value formulaArg) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache[expr] = value
}

// Load retrieves a sub-expression result
func (c *SubExpressionCache) Load(expr string) (formulaArg, bool) {
	if c == nil {
		return formulaArg{}, false
	}
	c.mu.RLock()
	defer c.mu.RUnlock()
	value, ok := c.cache[expr]
	return value, ok
}

// Clear clears the cache
func (c *SubExpressionCache) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache = make(map[string]formulaArg)
}

// Len returns the number of cached expressions
func (c *SubExpressionCache) Len() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.cache)
}

func subExprCacheScopedKey(sheet, expr string) string {
	return sheet + "\x00" + expr
}

func loadSubExprCacheValue(cache *SubExpressionCache, sheet, expr string) (formulaArg, bool) {
	if cache == nil {
		return formulaArg{}, false
	}
	if value, ok := cache.Load(subExprCacheScopedKey(sheet, expr)); ok {
		return value, true
	}
	return cache.Load(expr)
}

func formulaArgToFormulaLiteral(arg formulaArg) string {
	switch arg.Type {
	case ArgNumber:
		if arg.Boolean {
			return arg.Value()
		}
		return pgFormulaArgValue(arg)
	case ArgError:
		return arg.Value()
	case ArgEmpty:
		return `""`
	default:
		return `"` + strings.ReplaceAll(arg.Value(), `"`, `""`) + `"`
	}
}

// CalcCellValueWithSubExprCache calculates a cell value with sub-expression cache support
// This is optimized for dependency-based calculation where SUMIFS/AVERAGEIFS/INDEX-MATCH are pre-calculated
// formula parameter is provided to avoid re-reading from worksheet (lock-free)
// worksheetCache provides access to recently calculated values during batch calculation
func (f *File) CalcCellValueWithSubExprCache(sheet, cell, formula string, subExprCache *SubExpressionCache, worksheetCache *WorksheetCache, opts Options) (string, error) {
	if formula == "" {
		// Not a formula, return the cell value directly
		return f.GetCellValue(sheet, cell, opts)
	}

	// 首先检查 calcCache，看是否已经有完整的计算结果（比如批量计算的结果）
	cacheKey := fmt.Sprintf("%s!%s!raw=%t", sheet, cell, opts.RawCellValue)
	if cachedResult, found := f.calcCache.Load(cacheKey); found {
		return cachedResult.(string), nil
	}

	// Try to replace ALL SUMIFS/AVERAGEIFS/INDEX-MATCH in the formula with cached values
	modifiedFormula := formula
	replacements := 0
	missedCount := 0

	remainingFormula := modifiedFormula
	for {
		pgLookupExpr := extractFirstPGSupportedLookupExpr(sheet, remainingFormula)
		if pgLookupExpr == "" {
			break
		}

		if cachedValue, ok := loadSubExprCacheValue(subExprCache, sheet, pgLookupExpr); ok {
			modifiedFormula = strings.Replace(modifiedFormula, pgLookupExpr, formulaArgToFormulaLiteral(cachedValue), 1)
			replacements++
		} else {
			missedCount++
		}

		idx := strings.Index(remainingFormula, pgLookupExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(pgLookupExpr):]
		} else {
			break
		}
	}

	// Extract and replace ALL INDEX-MATCH expressions (do this first as they may be nested in IFERROR)
	remainingFormula = modifiedFormula
	for {
		indexMatchExpr := extractINDEXMATCHFromFormula(remainingFormula)
		if indexMatchExpr == "" {
			break
		}

		if cachedValue, ok := loadSubExprCacheValue(subExprCache, sheet, indexMatchExpr); ok {
			modifiedFormula = strings.Replace(modifiedFormula, indexMatchExpr, formulaArgToFormulaLiteral(cachedValue), 1)
			replacements++
		} else {
			missedCount++
		}

		// Remove the processed INDEX-MATCH to find the next one
		idx := strings.Index(remainingFormula, indexMatchExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(indexMatchExpr):]
		} else {
			break
		}
	}

	// Extract and replace ALL SUMIFS expressions (not just the first one)
	remainingFormula = modifiedFormula
	for {
		sumifsExpr := extractSUMIFSFromFormula(remainingFormula)
		if sumifsExpr == "" {
			break
		}

		if cachedValue, ok := loadSubExprCacheValue(subExprCache, sheet, sumifsExpr); ok {
			modifiedFormula = strings.Replace(modifiedFormula, sumifsExpr, formulaArgToFormulaLiteral(cachedValue), 1)
			replacements++
		} else {
			missedCount++
		}

		// Remove the processed SUMIFS to find the next one
		idx := strings.Index(remainingFormula, sumifsExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(sumifsExpr):]
		} else {
			break
		}
	}

	// Extract and replace ALL AVERAGEIFS expressions
	remainingFormula = modifiedFormula
	for {
		averageifsExpr := extractAVERAGEIFSFromFormula(remainingFormula)
		if averageifsExpr == "" {
			break
		}

		if cachedValue, ok := loadSubExprCacheValue(subExprCache, sheet, averageifsExpr); ok {
			modifiedFormula = strings.Replace(modifiedFormula, averageifsExpr, formulaArgToFormulaLiteral(cachedValue), 1)
			replacements++
		} else {
			missedCount++
		}

		// Remove the processed AVERAGEIFS to find the next one
		idx := strings.Index(remainingFormula, averageifsExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(averageifsExpr):]
		} else {
			break
		}
	}

	// Evaluate the formula
	var result string
	var err error

	// If we replaced sub-expressions, evaluate the simplified formula
	if replacements > 0 {
		result, err = f.evalFormulaString(sheet, cell, modifiedFormula, worksheetCache, opts)
	} else if missedCount > 0 {
		// Cache miss - use evalFormulaString to keep worksheetCache
		result, err = f.evalFormulaString(sheet, cell, formula, worksheetCache, opts)
	} else {
		// No SUMIFS/AVERAGEIFS/INDEX-MATCH in this formula, use evalFormulaString with worksheetCache
		result, err = f.evalFormulaString(sheet, cell, formula, worksheetCache, opts)
	}

	return result, err
}

// evalFormulaString evaluates a formula string directly (without reading from cell)
// This is used when the formula has been modified (e.g., SUMIFS replaced with value)
func (f *File) evalFormulaString(sheet, cell, formula string, worksheetCache *WorksheetCache, opts Options) (string, error) {
	if token, ok, err := f.tryCalcPostgresLookupFormula(sheet, cell, "="+formula, worksheetCache); ok {
		return token.Value(), err
	}

	// Remove leading =
	formula = strings.TrimPrefix(formula, "=")

	// Check cache first
	cacheKey := fmt.Sprintf("%s!%s!subexpr:%s", sheet, cell, formula)
	if opts.RawCellValue {
		cacheKey += "!raw=true"
	}

	if value, ok := f.calcCache.Load(cacheKey); ok {
		return value.(string), nil
	}

	tokens := f.parseFormulaTokensCached(formula)
	if tokens == nil {
		return "", fmt.Errorf("failed to parse formula: %s", formula)
	}

	// Create a calc context - matching CalcCellValue's context creation
	ctx := &calcContext{
		entry:             fmt.Sprintf("%s!%s", sheet, cell),
		maxCalcIterations: opts.MaxCalcIterations,
		iterations:        make(map[string]uint),
		iterationsCache:   make(map[string]formulaArg),
		// rangeCache is sync.Map, no initialization needed
		worksheetCache: worksheetCache, // Pass worksheetCache to formula engine
	}

	result, err := f.evalInfixExp(ctx, sheet, cell, tokens)

	// Convert result to string - result is formulaArg with String/Number/etc fields
	// CRITICAL: Even if err != nil, result may contain error value like "#DIV/0!"
	// We should return the error value string so it can be displayed in Excel
	resultStr := result.Value()

	if err != nil {
		// Return error value string (e.g., "#DIV/0!") along with error
		// This allows calling code to write the error value back to the cell
		return resultStr, err
	}

	// Cache the result
	f.calcCache.Store(cacheKey, resultStr)

	return resultStr, nil
}
