package excelize

import (
	"strings"
	"sync"
)

var (
	globalFormulaOptimizationCache sync.Map
	globalPGLookupFormulaCache     sync.Map
)

type formulaOptimizationClass struct {
	hasAverageOffset bool
	sumifsExpr       string
	pureSumifs       bool
	sumifsSource     string
	indexMatchExpr   string
	indexSource      string
	isBatchType      bool
}

type pgLookupOptimizationClass struct {
	whole bool
	exprs []string
}

func getFormulaOptimizationClass(formula string) formulaOptimizationClass {
	if formula == "" {
		return formulaOptimizationClass{}
	}
	if cached, ok := globalFormulaOptimizationCache.Load(formula); ok {
		return cached.(formulaOptimizationClass)
	}
	class := buildFormulaOptimizationClass(formula)
	actual, _ := globalFormulaOptimizationCache.LoadOrStore(formula, class)
	return actual.(formulaOptimizationClass)
}

func buildFormulaOptimizationClass(formula string) formulaOptimizationClass {
	class := formulaOptimizationClass{
		hasAverageOffset: isAverageOffsetFormula(formula),
	}
	cleanFormula := strings.TrimSpace(strings.TrimPrefix(formula, "="))

	if expr := extractSUMIFSFromFormula(formula); expr != "" {
		class.sumifsExpr = expr
		class.pureSumifs = cleanFormula == strings.TrimSpace(expr)
		class.sumifsSource = extractFormulaSourceSheet(expr)
	} else if expr := extractAVERAGEIFSFromFormula(formula); expr != "" {
		class.sumifsExpr = expr
		class.pureSumifs = cleanFormula == strings.TrimSpace(expr)
		class.sumifsSource = extractFormulaSourceSheet(expr)
	}

	if strings.Contains(formula, "INDEX(") && strings.Contains(formula, "MATCH(") {
		class.indexMatchExpr = extractINDEXMATCHFromFormula(formula)
		if class.indexMatchExpr != "" {
			class.indexSource = extractIndexMatchSourceSheet(class.indexMatchExpr)
		}
	}

	class.isBatchType = class.hasAverageOffset || class.sumifsExpr != "" || class.indexMatchExpr != ""
	return class
}

func getPGLookupOptimizationClass(sheet, formula string) pgLookupOptimizationClass {
	if sheet == "" || formula == "" {
		return pgLookupOptimizationClass{}
	}
	key := sheet + "\x00" + formula
	if cached, ok := globalPGLookupFormulaCache.Load(key); ok {
		return cached.(pgLookupOptimizationClass)
	}
	class := pgLookupOptimizationClass{
		whole: isSupportedPGWholeLookupFormula(sheet, formula),
	}
	if !class.whole {
		class.exprs = extractAllPGSupportedLookupExprs(sheet, formula)
	}
	actual, _ := globalPGLookupFormulaCache.LoadOrStore(key, class)
	return actual.(pgLookupOptimizationClass)
}

func extractFormulaSourceSheet(expr string) string {
	if idx := strings.Index(expr, "!"); idx >= 0 {
		sheetName := strings.Trim(expr[:idx], "'")
		sheetName = strings.TrimPrefix(sheetName, "SUMIFS(")
		sheetName = strings.TrimPrefix(sheetName, "AVERAGEIFS(")
		sheetName = strings.TrimSpace(strings.Trim(sheetName, "'"))
		return sheetName
	}
	return ""
}

func extractIndexMatchSourceSheet(expr string) string {
	idx := strings.Index(expr, "INDEX(")
	if idx == -1 {
		return ""
	}
	remaining := expr[idx+len("INDEX("):]
	commaIdx := strings.Index(remaining, ",")
	if commaIdx == -1 {
		return ""
	}
	rangeRef := strings.TrimSpace(remaining[:commaIdx])
	if !strings.Contains(rangeRef, "!") {
		return ""
	}
	parts := strings.SplitN(rangeRef, "!", 2)
	return strings.TrimSpace(strings.Trim(parts[0], "'"))
}
