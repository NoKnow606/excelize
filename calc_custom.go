package excelize

import (
	"container/list"
)

// Custom formula stubs for functions handled externally (MCP tools, frontend rendering, etc.).
// These prevent #VALUE! errors during recalculation.
//
// Note: excelize strips underscores from function names before lookup, so:
//   AI_IMAGE -> AIIMAGE, FORMULA_BUTTON -> FORMULABUTTON, etc.

func cachedCellValue(fn *formulaFuncs) formulaArg {
	val, err := fn.f.GetCellValue(fn.sheet, fn.cell)
	if err != nil {
		return newStringFormulaArg("")
	}
	return newStringFormulaArg(val)
}

// IMAGE returns the first argument as-is.
// Syntax: =IMAGE(source)
func (fn *formulaFuncs) IMAGE(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMAGE requires 1 argument")
	}
	arg := argsList.Front().Value.(formulaArg)
	return newStringFormulaArg(arg.Value())
}

// VIDEO returns the first argument (url) as-is, same as IMAGE.
// Syntax: =VIDEO(url, ...)
func (fn *formulaFuncs) VIDEO(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "VIDEO requires at least 1 argument")
	}
	arg := argsList.Front().Value.(formulaArg)
	return newStringFormulaArg(arg.Value())
}

// AIIMAGE returns the cell's current cached value.
// Syntax: =AI_IMAGE(...) — excelize strips underscore → AIIMAGE
func (fn *formulaFuncs) AIIMAGE(argsList *list.List) formulaArg {
	return cachedCellValue(fn)
}

// HTML returns the cell's current cached value.
// Syntax: =HTML(...)
func (fn *formulaFuncs) HTML(argsList *list.List) formulaArg {
	return cachedCellValue(fn)
}

// FORMULABUTTON returns the cell's current cached value.
// Syntax: =FORMULA_BUTTON(...) — excelize strips underscore → FORMULABUTTON
func (fn *formulaFuncs) FORMULABUTTON(argsList *list.List) formulaArg {
	return cachedCellValue(fn)
}

// REFRESHBUTTON returns the cell's current cached value.
// Syntax: =REFRESH_BUTTON(...) — excelize strips underscore → REFRESHBUTTON
func (fn *formulaFuncs) REFRESHBUTTON(argsList *list.List) formulaArg {
	return cachedCellValue(fn)
}

// SKILLGENERATE returns the cell's current cached value.
// Syntax: =SKILL_GENERATE(...) — excelize strips underscore → SKILLGENERATE
func (fn *formulaFuncs) SKILLGENERATE(argsList *list.List) formulaArg {
	return cachedCellValue(fn)
}
