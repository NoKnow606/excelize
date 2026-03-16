package excelize

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/efp"
)

// CalcCellValueLockFree provides a lock-free version of CalcCellValue specifically
// designed for use within the dependency analysis system (DAG).
//
// This function is safe to use when:
// 1. All dependencies of the cell have already been calculated
// 2. The cell is being calculated within a dependency level where no circular dependencies exist
// 3. Results will be cached after calculation
//
// Benefits:
// - No context locks (ctx.mu) needed since dependencies are already resolved
// - Uses ReadOnly functions to minimize lock contention
// - Suitable for high-concurrency scenarios within dependency levels
func (f *File) CalcCellValueLockFree(sheet, cell string, opts ...Options) (result string, err error) {
	options := f.getOptions(opts...)
	var (
		rawCellValue = options.RawCellValue
		styleIdx     int
		token        formulaArg
	)

	// Check cache first
	cacheKey := fmt.Sprintf("%s!%s!raw=%t", sheet, cell, rawCellValue)
	if cachedResult, found := f.calcCache.Load(cacheKey); found {
		return cachedResult.(string), nil
	}

	// Get formula and parse it
	formula, err := f.getCellFormulaReadOnly(sheet, cell, true)
	if err != nil {
		return "", err
	}

	// If no formula, get the cell value directly
	if formula == "" {
		return f.getCellValueLockFree(sheet, cell, rawCellValue)
	}

	// Parse and evaluate formula without context locks
	tokens := f.parseFormulaTokensCached(formula)
	if tokens == nil {
		return f.getCellValueLockFree(sheet, cell, rawCellValue)
	}

	// Evaluate formula with lock-free context
	// 关键：不使用 context 锁，因为在 DAG 中已经保证了依赖顺序
	token, err = f.evalInfixExpLockFree(sheet, cell, tokens, options.MaxCalcIterations)
	if err != nil {
		result = token.String
		return
	}

	// Format result
	if !rawCellValue {
		styleIdx, _ = f.GetCellStyleReadOnly(sheet, cell)
	}

	if token.Type == ArgNumber && !token.Boolean {
		_, precision, decimal := isNumeric(token.Value())
		if precision > 15 {
			result, err = f.formattedValue(&xlsxC{S: styleIdx, V: strings.ToUpper(strconv.FormatFloat(decimal, 'G', 15, 64))}, rawCellValue, CellTypeNumber)
			if err == nil {
				f.calcCache.Store(cacheKey, result)
			}
			return
		}
		if !strings.HasPrefix(result, "0") {
			result, err = f.formattedValue(&xlsxC{S: styleIdx, V: strings.ToUpper(strconv.FormatFloat(decimal, 'f', -1, 64))}, rawCellValue, CellTypeNumber)
		}
		if err == nil {
			f.calcCache.Store(cacheKey, result)
		}
		return
	}

	result, err = f.formattedValue(&xlsxC{S: styleIdx, V: token.Value()}, rawCellValue, CellTypeInlineString)
	if err == nil {
		f.calcCache.Store(cacheKey, result)
	}
	return
}

// getCellValueLockFree gets cell value without using locks
func (f *File) getCellValueLockFree(sheet, cell string, rawCellValue bool) (string, error) {
	// Try cache first
	ref := fmt.Sprintf("%s!%s", sheet, cell)
	if cached, ok := f.calcCache.Load(ref); ok {
		switch v := cached.(type) {
		case string:
			return v, nil
		case formulaArg:
			return v.Value(), nil
		}
	}

	// Use regular GetCellValue - in DAG context this is safe because dependencies are resolved
	value, err := f.GetCellValue(sheet, cell, Options{RawCellValue: rawCellValue})
	if err != nil {
		return "", err
	}

	return value, nil
}

// evalInfixExpLockFree evaluates formula without context locks
// This is safe in DAG system because dependencies are already calculated
func (f *File) evalInfixExpLockFree(sheet, cell string, tokens []efp.Token, maxIter uint) (formulaArg, error) {
	// 创建一个简化的context，不需要锁保护
	// 因为在DAG系统中，同一level内的公式不会相互依赖
	ctx := &calcContext{
		entry:             fmt.Sprintf("%s!%s", sheet, cell),
		maxCalcIterations: maxIter,
		iterations:        make(map[string]uint),
		iterationsCache:   make(map[string]formulaArg),
	}

	// 使用原有的 evalInfixExp，但context的锁不会被竞争
	// 因为每个CalcCellValueLockFree都有自己独立的context
	return f.evalInfixExp(ctx, sheet, cell, tokens)
}
