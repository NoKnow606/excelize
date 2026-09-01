package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcAI(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 测试 AI 函数 - 单元格没有缓存值时返回空字符串
	f.SetCellValue("Sheet1", "A1", "workflow_123")
	f.SetCellValue("Sheet1", "B1", "param_value")
	f.SetCellFormula("Sheet1", "C1", "AI(A1,B1)")
	result, err := f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "", result) // 没有缓存值，返回空字符串

	// 测试 AI 函数 - 单元格有缓存值时返回缓存值
	// 先设置公式，然后直接修改单元格的 V 属性来模拟已有缓存值
	f.SetCellFormula("Sheet1", "C2", `AI("wf_id","input_data")`)
	ws, _ := f.workSheetReader("Sheet1")
	// 直接设置单元格的缓存值 (V 属性)
	ws.SheetData.Row[1].C[2].V = "cached_result"
	result, err = f.CalcCellValue("Sheet1", "C2")
	assert.NoError(t, err)
	assert.Equal(t, "cached_result", result) // 返回缓存值

	// 测试 AI 函数参数错误 - 无参数
	f.SetCellFormula("Sheet1", "C5", "AI()")
	result, err = f.CalcCellValue("Sheet1", "C5")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "AI requires 2 arguments")

	// 测试 AI 函数参数错误 - 只有一个参数
	f.SetCellFormula("Sheet1", "C6", `AI("only_one")`)
	result, err = f.CalcCellValue("Sheet1", "C6")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "AI requires 2 arguments")

	// 测试 AI 函数参数错误 - 过多参数
	f.SetCellFormula("Sheet1", "C7", `AI("a","b","c")`)
	result, err = f.CalcCellValue("Sheet1", "C7")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "AI requires 2 arguments")
}

func TestExternalCachedFormulaIgnoredByDependencyGraph(t *testing.T) {
	f := NewFile()
	defer f.Close()

	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "prompt"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "context"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "AI(A1,B1)"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=C1&\"!\""))

	graph := f.buildDependencyGraph()
	if _, ok := graph.nodes["Sheet1!C1"]; ok {
		t.Fatalf("external cached formula C1 should not participate in dependency graph")
	}

	node, ok := graph.nodes["Sheet1!D1"]
	if !ok {
		t.Fatalf("expected dependent formula D1 in dependency graph")
	}
	assert.Contains(t, node.dependencies, "Sheet1!C1")

	affected := f.findAffectedCellsByCells(graph, map[string]bool{"Sheet1!A1": true})
	assert.False(t, affected["Sheet1!C1"], "AI formula should not be marked affected by A1 updates")
	assert.False(t, affected["Sheet1!D1"], "downstream formulas should not be marked affected until AI cell itself changes")

	affected = f.findAffectedCellsByCells(graph, map[string]bool{"Sheet1!C1": true})
	assert.True(t, affected["Sheet1!D1"], "downstream formulas should still react to explicit AI cell value changes")
}

func TestExternalCachedFormulaCacheIsPreserved(t *testing.T) {
	f := NewFile()
	defer f.Close()

	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "AI(\"wf\",\"input\")"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=1+1"))

	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	for i := range ws.SheetData.Row {
		for j := range ws.SheetData.Row[i].C {
			switch ws.SheetData.Row[i].C[j].R {
			case "C1":
				ws.SheetData.Row[i].C[j].V = "cached-ai"
				ws.SheetData.Row[i].C[j].T = "str"
			case "D1":
				ws.SheetData.Row[i].C[j].V = "2"
				ws.SheetData.Row[i].C[j].T = "n"
			}
		}
	}

	calcChain := &xlsxCalcChain{C: []xlsxCalcChainC{{R: "C1", I: 1}, {R: "D1", I: 0}}}
	assert.NoError(t, f.clearCachesFromIndex(calcChain, 1, 0))

	ws, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			switch cell.R {
			case "C1":
				assert.Equal(t, "cached-ai", cell.V, "AI cache should be preserved")
			case "D1":
				assert.Equal(t, "", cell.V, "regular formula cache should still be cleared")
			}
		}
	}
}

func TestExternalCachedFormulaDoesNotTriggerCalculatedCallback(t *testing.T) {
	f := NewFile()
	defer f.Close()

	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "AI(\"wf\",\"input\")"))

	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	for i := range ws.SheetData.Row {
		for j := range ws.SheetData.Row[i].C {
			if ws.SheetData.Row[i].C[j].R == "C1" {
				ws.SheetData.Row[i].C[j].V = "old"
				ws.SheetData.Row[i].C[j].T = "str"
			}
		}
	}

	callCount := 0
	f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
		callCount++
	}

	f.setFormulaValue("Sheet1", "C1", "new")
	assert.Equal(t, 0, callCount, "external cached formulas should not trigger OnCellCalculated")

	ws, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	for _, row := range ws.SheetData.Row {
		for _, cell := range row.C {
			if cell.R == "C1" {
				assert.Equal(t, "new", cell.V, "cached value should still be updated")
			}
		}
	}

	assert.NoError(t, f.recalculateCell("Sheet1", "C1"))
	assert.Equal(t, 0, callCount, "recalculateCell should skip external cached formulas")
}

func TestExternalCachedFormulaExcludedFromCalcChainMaintenance(t *testing.T) {
	f := NewFile()
	defer f.Close()

	aiFormula := `AI("wf","input")`
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", aiFormula))
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=1+1"))

	assert.NoError(t, f.updateCalcChainForFormulas([]FormulaUpdate{
		{Sheet: "Sheet1", Cell: "C1", Formula: aiFormula},
		{Sheet: "Sheet1", Cell: "D1", Formula: "=1+1"},
	}))
	if assert.NotNil(t, f.CalcChain) {
		if assert.Len(t, f.CalcChain.C, 1) {
			assert.Equal(t, "D1", f.CalcChain.C[0].R)
		}
	}

	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "C1", I: 1}, {R: "D1", I: 0}}}
	f.calcChainWriter()
	if assert.NotNil(t, f.CalcChain) {
		if assert.Len(t, f.CalcChain.C, 1) {
			assert.Equal(t, "D1", f.CalcChain.C[0].R)
			assert.Equal(t, 1, f.CalcChain.C[0].I)
		}
	}

	assert.NoError(t, f.RebuildCalcChain())
	if assert.NotNil(t, f.CalcChain) {
		if assert.Len(t, f.CalcChain.C, 1) {
			assert.Equal(t, "D1", f.CalcChain.C[0].R)
		}
	}

	f.Pkg.Store(defaultXMLPathCalcChain, []byte("<calcChain/>"))
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "C1", I: 1}}}
	f.calcChainWriter()
	assert.Nil(t, f.CalcChain)
	_, ok := f.Pkg.Load(defaultXMLPathCalcChain)
	assert.False(t, ok, "stale calcChain xml should be removed when only external cached formulas remain")
}

func TestExternalCachedFormulaSettersRemoveCalcChainEntries(t *testing.T) {
	f := NewFile()
	defer f.Close()

	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "C1", I: 1}}}
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", `AI("wf","input")`))
	assert.Nil(t, f.CalcChain)

	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "C1", I: 1}}}
	assert.NoError(t, f.SetCellFormulaWithValue("Sheet1", "C1", `AI("wf","input")`, "cached"))
	assert.Nil(t, f.CalcChain)
}
