package excelize

import (
	"sync"

	"github.com/xuri/efp"
)

var globalFormulaParseCache sync.Map

// parseFormulaTokensCached reuses parsed token streams for read-only evaluation paths.
func parseFormulaTokensCached(formula string) []efp.Token {
	if formula == "" {
		return nil
	}
	if cached, ok := globalFormulaParseCache.Load(formula); ok {
		return cached.([]efp.Token)
	}
	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)
	if tokens == nil {
		return nil
	}
	actual, _ := globalFormulaParseCache.LoadOrStore(formula, tokens)
	return actual.([]efp.Token)
}

func (f *File) parseFormulaTokensCached(formula string) []efp.Token {
	if formula == "" {
		return nil
	}
	if cached, ok := f.formulaParseCache.Load(formula); ok {
		return cached.([]efp.Token)
	}
	tokens := parseFormulaTokensCached(formula)
	if tokens == nil {
		return nil
	}
	actual, _ := f.formulaParseCache.LoadOrStore(formula, tokens)
	return actual.([]efp.Token)
}
