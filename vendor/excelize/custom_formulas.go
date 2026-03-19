// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"container/list"
)

// Custom formula stubs for functions that are handled externally (by MCP tools,
// frontend rendering, etc.). These functions are not natively supported by
// excelize, so without these stubs, recalculation would produce #VALUE! errors.
//
// Each stub simply returns the cell's current cached value, preserving whatever
// value was previously stored in the cell.

// cachedValueFormula is a helper that returns the current cell's cached value.
// All custom formula stubs delegate to this function.
func cachedValueFormula(fn *formulaFuncs) formulaArg {
	cachedValue, err := fn.f.GetCellValue(fn.sheet, fn.cell)
	if err != nil {
		return newStringFormulaArg("")
	}
	return newStringFormulaArg(cachedValue)
}

// VIDEO returns the first argument (URL) as-is, same as IMAGE.
// Syntax: =VIDEO(url, ...)
func (fn *formulaFuncs) VIDEO(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newStringFormulaArg("")
	}
	arg := argsList.Front().Value.(formulaArg)
	return newStringFormulaArg(arg.Value())
}

// AIIMAGE returns the current cell's cached value.
// Syntax: =AI_IMAGE(...)
// Note: excelize strips underscores from function names, so AI_IMAGE becomes AIIMAGE.
func (fn *formulaFuncs) AIIMAGE(argsList *list.List) formulaArg {
	return cachedValueFormula(fn)
}

// HTML returns the current cell's cached value.
// Syntax: =HTML(...)
func (fn *formulaFuncs) HTML(argsList *list.List) formulaArg {
	return cachedValueFormula(fn)
}

// FORMULABUTTON returns the current cell's cached value.
// Syntax: =FORMULA_BUTTON(...)
// Note: excelize strips underscores, so FORMULA_BUTTON becomes FORMULABUTTON.
func (fn *formulaFuncs) FORMULABUTTON(argsList *list.List) formulaArg {
	return cachedValueFormula(fn)
}

// REFRESHBUTTON returns the current cell's cached value.
// Syntax: =REFRESH_BUTTON(...)
// Note: excelize strips underscores, so REFRESH_BUTTON becomes REFRESHBUTTON.
func (fn *formulaFuncs) REFRESHBUTTON(argsList *list.List) formulaArg {
	return cachedValueFormula(fn)
}

// SKILLGENERATE returns the current cell's cached value.
// Syntax: =SKILL_GENERATE(...)
// Note: excelize strips underscores, so SKILL_GENERATE becomes SKILLGENERATE.
func (fn *formulaFuncs) SKILLGENERATE(argsList *list.List) formulaArg {
	return cachedValueFormula(fn)
}
