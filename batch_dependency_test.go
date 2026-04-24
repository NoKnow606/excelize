package excelize

import (
	"sort"
	"strings"
	"testing"
)

func TestDependencyGraphAssignLevelsAndMerge(t *testing.T) {
	graph := &dependencyGraph{
		nodes: map[string]*formulaNode{
			"S1!A1": {dependencies: []string{}, level: -1},
			"S1!A2": {dependencies: []string{"S1!A1"}, level: -1},
			"S1!B1": {dependencies: []string{}, level: -1},
			"S1!B2": {dependencies: []string{"S1!B1"}, level: -1},
			"S1!C1": {dependencies: []string{"S1!A2", "S1!B2"}, level: -1},
		},
	}

	graph.assignLevels()

	if len(graph.levels) == 0 {
		t.Fatalf("expected levels assigned")
	}

	levelMap := make(map[string]int)
	for idx, cells := range graph.levels {
		for _, cell := range cells {
			levelMap[cell] = idx
		}
	}

	if levelMap["S1!A2"] <= levelMap["S1!A1"] {
		t.Fatalf("expected S1!A2 level higher than S1!A1")
	}
	if levelMap["S1!C1"] <= levelMap["S1!A2"] || levelMap["S1!C1"] <= levelMap["S1!B2"] {
		t.Fatalf("S1!C1 level incorrect: %+v", levelMap)
	}
}

func TestExtractDependencies(t *testing.T) {
	formula := "=SUM(Sheet2!$A$1:Sheet2!$A$3)+SUM($B$1:$B$2)+Sheet3!C5+Sheet4!$D:$D"
	deps := extractDependencies(formula, "Sheet1", "A1")

	expected := []string{
		"Sheet1!B1", "Sheet1!B2",
		"Sheet2!A1", "Sheet2!A3",
		"Sheet3!C5",
		"Sheet4!D:COLUMN_RANGE",
	}

	for _, want := range expected {
		if !containsDep(deps, want) {
			t.Fatalf("missing dependency %s in %+v", want, deps)
		}
	}
}

func TestExtractDependenciesWithColumnIndex(t *testing.T) {
	columnIndex := map[string][]string{
		"Sheet2!A": {"Sheet2!A1", "Sheet2!A2"},
		"Sheet1!B": {"Sheet1!B1"},
		"Sheet1!C": {"Sheet1!C1"},
	}

	formula := "=SUM(Sheet2!$A:$A)+SUM($B:$C)"
	deps := extractDependenciesWithColumnIndex(formula, "Sheet1", "D1", columnIndex)
	sort.Strings(deps)

	expected := []string{"Sheet1!B1", "Sheet1!C1", "Sheet2!A1", "Sheet2!A2"}
	if len(deps) != len(expected) {
		t.Fatalf("unexpected deps length: %d vs %d (%+v)", len(deps), len(expected), deps)
	}

	for i, want := range expected {
		if deps[i] != want {
			t.Fatalf("expected %s at pos %d, got %s (all=%+v)", want, i, deps[i], deps)
		}
	}
}

func TestExpandCellRef(t *testing.T) {
	deps := make(map[string]bool)
	expandCellRef("Sheet1", "$A$1:$A$3", deps)
	expandCellRef("Sheet1", "$B$2", deps)

	if !deps["Sheet1!A1"] || !deps["Sheet1!A3"] || !deps["Sheet1!B2"] {
		t.Fatalf("unexpected deps map: %+v", deps)
	}
}

func TestFormatFloat(t *testing.T) {
	if formatFloat(0) != "0" {
		t.Fatalf("formatFloat zero failed")
	}
	if formatFloat(-12.34)[:3] != "-12" {
		t.Fatalf("formatFloat negative integer part incorrect: %s", formatFloat(-12.34))
	}
	if got := formatFloat(123.4567); got[:3] != "123" {
		t.Fatalf("formatFloat positive incorrect: %s", got)
	}
	// Ensure fractional digits trimmed
	if strings.HasSuffix(formatFloat(1.2000001), "0000") {
		t.Fatalf("fractional trimming failed: %s", formatFloat(1.2000001))
	}
}

func containsDep(deps []string, want string) bool {
	for _, dep := range deps {
		if dep == want {
			return true
		}
		// Column range expansion may duplicate sheet prefix, permit contains substring
		if strings.Contains(dep, want) {
			return true
		}
	}
	return false
}

// TestRecalculateAffectedByCells tests the incremental recalculation API
func TestRecalculateAffectedByCells(t *testing.T) {
	t.Run("DirectDependency", func(t *testing.T) {
		f := NewFile()

		// Set initial values and formula
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "20" {
			t.Fatalf("expected B1=20 before update, got %s", b1Before)
		}

		// Update A1 and use incremental recalculation
		f.SetCellValue("Sheet1", "A1", 50)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify B1 was recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "100" {
			t.Errorf("expected B1=100 after update, got %s", b1After)
		}
	})

	t.Run("CrossSheetDependency", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has source values
		f.SetCellValue("Data", "A1", 100)

		// Sheet1 references Data sheet
		f.SetCellFormula("Sheet1", "B1", "=Data!A1*3")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "300" {
			t.Fatalf("expected Sheet1!B1=300 before update, got %s", b1Before)
		}

		// Update Data!A1 and use incremental recalculation
		f.SetCellValue("Data", "A1", 200)

		updatedCells := map[string]bool{
			"Data!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify Sheet1!B1 was recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "600" {
			t.Errorf("expected Sheet1!B1=600 after update, got %s", b1After)
		}
	})

	t.Run("ChainDependency", func(t *testing.T) {
		f := NewFile()

		// Create chain dependency: A1 -> B1 -> C1 -> D1
		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "B1", "=A1+1")
		f.SetCellFormula("Sheet1", "C1", "=B1+1")
		f.SetCellFormula("Sheet1", "D1", "=C1+1")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		d1Before, _ := f.GetCellValue("Sheet1", "D1")
		if d1Before != "4" {
			t.Fatalf("expected D1=4 before update, got %s", d1Before)
		}

		// Update A1 and use incremental recalculation
		f.SetCellValue("Sheet1", "A1", 10)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify entire chain was recalculated correctly
		b1, _ := f.GetCellValue("Sheet1", "B1")
		c1, _ := f.GetCellValue("Sheet1", "C1")
		d1, _ := f.GetCellValue("Sheet1", "D1")

		if b1 != "11" {
			t.Errorf("expected B1=11, got %s", b1)
		}
		if c1 != "12" {
			t.Errorf("expected C1=12, got %s", c1)
		}
		if d1 != "13" {
			t.Errorf("expected D1=13, got %s", d1)
		}
	})

	t.Run("MultipleUpdates", func(t *testing.T) {
		f := NewFile()

		// Set initial values and formulas
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellFormula("Sheet1", "B1", "=A1+A2")
		f.SetCellFormula("Sheet1", "C1", "=A1*A2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if b1Before != "30" || c1Before != "200" {
			t.Fatalf("expected B1=30, C1=200 before update, got B1=%s, C1=%s", b1Before, c1Before)
		}

		// Update both A1 and A2
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellValue("Sheet1", "A2", 200)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
			"Sheet1!A2": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify both formulas were recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		c1After, _ := f.GetCellValue("Sheet1", "C1")

		if b1After != "300" {
			t.Errorf("expected B1=300 after update, got %s", b1After)
		}
		if c1After != "20000" {
			t.Errorf("expected C1=20000 after update, got %s", c1After)
		}
	})

	t.Run("UnaffectedCellsUnchanged", func(t *testing.T) {
		f := NewFile()

		// Set up two independent formula chains
		// Chain 1: A1 -> B1
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")

		// Chain 2: C1 -> D1 (independent)
		f.SetCellValue("Sheet1", "C1", 100)
		f.SetCellFormula("Sheet1", "D1", "=C1*2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		d1Before, _ := f.GetCellValue("Sheet1", "D1")
		if b1Before != "20" || d1Before != "200" {
			t.Fatalf("expected B1=20, D1=200 before update, got B1=%s, D1=%s", b1Before, d1Before)
		}

		// Update only A1 (should not affect D1)
		f.SetCellValue("Sheet1", "A1", 50)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify B1 was recalculated
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "100" {
			t.Errorf("expected B1=100 after update, got %s", b1After)
		}

		// Verify D1 is unchanged (still 200)
		d1After, _ := f.GetCellValue("Sheet1", "D1")
		if d1After != "200" {
			t.Errorf("expected D1=200 (unchanged), got %s", d1After)
		}
	})

	t.Run("EmptyUpdates", func(t *testing.T) {
		f := NewFile()

		// Empty updates should return nil without error
		err := f.RecalculateAffectedByCells(nil)
		if err != nil {
			t.Fatalf("empty updates should not fail: %v", err)
		}

		err = f.RecalculateAffectedByCells(map[string]bool{})
		if err != nil {
			t.Fatalf("empty map should not fail: %v", err)
		}
	})

	t.Run("SUMFormula", func(t *testing.T) {
		f := NewFile()

		// Set up SUM formula
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellValue("Sheet1", "A3", 30)
		f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A3)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "60" {
			t.Fatalf("expected B1=60 before update, got %s", b1Before)
		}

		// Update A2
		f.SetCellValue("Sheet1", "A2", 100)

		updatedCells := map[string]bool{
			"Sheet1!A2": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify SUM was recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "140" {
			t.Errorf("expected B1=140 after update, got %s", b1After)
		}
	})
}

// TestRecalculateAffectedByCellsWithExclusion tests the incremental recalculation with exclusion API
func TestRecalculateAffectedByCellsWithExclusion(t *testing.T) {
	t.Run("ExcludeSingleCell", func(t *testing.T) {
		// 场景：A1 -> B1 -> C1
		// 更新 A1，排除 B1，只有 C1 被重算
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")
		f.SetCellFormula("Sheet1", "C1", "=B1+5")

		f.RecalculateAllWithDependency()

		// 验证初始值
		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if b1Before != "20" {
			t.Fatalf("expected B1=20 before, got %s", b1Before)
		}
		if c1Before != "25" {
			t.Fatalf("expected C1=25 before, got %s", c1Before)
		}

		// 更新 A1，同时手动设置 B1 的预计算值（不是正确的计算结果）
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellValue("Sheet1", "B1", 999) // 预计算值，正确值应该是 200

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
			"Sheet1!B1": true,
		}
		excludeCells := map[string]bool{
			"Sheet1!B1": true, // 排除 B1，不要重算
		}

		// 写入 B1 缓存
		f.SetCellValue("Sheet1", "B1", 999)

		err := f.RecalculateAffectedByCellsWithExclusion(updatedCells, excludeCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCellsWithExclusion failed: %v", err)
		}

		// B1 应该保持预计算值 999，不被重算
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "999" {
			t.Errorf("expected B1=999 (excluded from recalc), got %s", b1After)
		}

		// C1 应该使用 B1 的值重算: 999 + 5 = 1004
		c1After, _ := f.GetCellValue("Sheet1", "C1")
		if c1After != "1004" {
			t.Errorf("expected C1=1004, got %s", c1After)
		}
	})

	t.Run("ExcludeMultipleCells", func(t *testing.T) {
		// 场景：A1 -> B1 -> C1 -> D1
		// 更新 A1，排除 B1 和 C1，只有 D1 被重算
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")
		f.SetCellFormula("Sheet1", "C1", "=B1+10")
		f.SetCellFormula("Sheet1", "D1", "=C1*3")

		f.RecalculateAllWithDependency()

		// 更新 A1，手动设置 B1 和 C1 的预计算值
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellValue("Sheet1", "B1", 500)  // 预计算值
		f.SetCellValue("Sheet1", "C1", 1000) // 预计算值

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
			"Sheet1!B1": true,
			"Sheet1!C1": true,
		}
		excludeCells := map[string]bool{
			"Sheet1!B1": true,
			"Sheet1!C1": true,
		}

		err := f.RecalculateAffectedByCellsWithExclusion(updatedCells, excludeCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCellsWithExclusion failed: %v", err)
		}

		// B1 和 C1 保持预计算值
		b1, _ := f.GetCellValue("Sheet1", "B1")
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if b1 != "500" {
			t.Errorf("expected B1=500 (excluded), got %s", b1)
		}
		if c1 != "1000" {
			t.Errorf("expected C1=1000 (excluded), got %s", c1)
		}

		// D1 使用 C1 的预计算值重算: 1000 * 3 = 3000
		d1, _ := f.GetCellValue("Sheet1", "D1")
		if d1 != "3000" {
			t.Errorf("expected D1=3000, got %s", d1)
		}
	})

	t.Run("ExcludeCrossSheet", func(t *testing.T) {
		// 场景：Data!A1 -> Data!B1 -> Sheet1!C1
		// 更新 Data!A1，排除 Data!B1
		f := NewFile()
		f.NewSheet("Data")

		f.SetCellValue("Data", "A1", 10)
		f.SetCellFormula("Data", "B1", "=A1*5")
		f.SetCellFormula("Sheet1", "C1", "=Data!B1+100")

		f.RecalculateAllWithDependency()

		// 更新 Data!A1，手动设置 Data!B1 的预计算值
		f.SetCellValue("Data", "A1", 100)
		f.SetCellValue("Data", "B1", 888) // 预计算值

		updatedCells := map[string]bool{
			"Data!A1": true,
			"Data!B1": true,
		}
		excludeCells := map[string]bool{
			"Data!B1": true,
		}

		err := f.RecalculateAffectedByCellsWithExclusion(updatedCells, excludeCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCellsWithExclusion failed: %v", err)
		}

		// Data!B1 保持预计算值
		b1, _ := f.GetCellValue("Data", "B1")
		if b1 != "888" {
			t.Errorf("expected Data!B1=888 (excluded), got %s", b1)
		}

		// Sheet1!C1 使用 Data!B1 的预计算值重算: 888 + 100 = 988
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "988" {
			t.Errorf("expected Sheet1!C1=988, got %s", c1)
		}
	})

	t.Run("EmptyExcludeList", func(t *testing.T) {
		// 空排除列表，行为应该和 RecalculateAffectedByCells 一样
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")

		f.RecalculateAllWithDependency()

		f.SetCellValue("Sheet1", "A1", 50)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}

		err := f.RecalculateAffectedByCellsWithExclusion(updatedCells, nil)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCellsWithExclusion failed: %v", err)
		}

		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "100" {
			t.Errorf("expected B1=100, got %s", b1)
		}

		// 空 map 也应该正常工作
		f.SetCellValue("Sheet1", "A1", 25)
		updatedCells["Sheet1!A1"] = true

		err = f.RecalculateAffectedByCellsWithExclusion(updatedCells, map[string]bool{})
		if err != nil {
			t.Fatalf("RecalculateAffectedByCellsWithExclusion with empty map failed: %v", err)
		}

		b1, _ = f.GetCellValue("Sheet1", "B1")
		if b1 != "50" {
			t.Errorf("expected B1=50, got %s", b1)
		}
	})

	t.Run("ExcludeNonAffectedCell", func(t *testing.T) {
		// 排除列表中的单元格不在 affected 中，不应该影响结果
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")
		f.SetCellValue("Sheet1", "D1", 999) // 不相关的单元格

		f.RecalculateAllWithDependency()

		f.SetCellValue("Sheet1", "A1", 50)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		excludeCells := map[string]bool{
			"Sheet1!D1": true, // D1 不在 affected 中
		}

		err := f.RecalculateAffectedByCellsWithExclusion(updatedCells, excludeCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCellsWithExclusion failed: %v", err)
		}

		// B1 应该正常重算
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "100" {
			t.Errorf("expected B1=100, got %s", b1)
		}
	})
}

// TestBatchUpdateValuesAndFormulasWithRecalc tests the new incremental update API
func TestBatchUpdateValuesAndFormulasWithRecalc(t *testing.T) {
	t.Run("BasicFunctionality", func(t *testing.T) {
		f := NewFile()

		// Set initial values and formula
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellFormula("Sheet1", "C1", "=A1+A2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if c1Before != "30" {
			t.Fatalf("expected C1=30 before update, got %s", c1Before)
		}

		// Use batch update with new values and formulas
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
			{Sheet: "Sheet1", Cell: "A2", Value: 200},
		}
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
			{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify results
		b1, _ := f.GetCellValue("Sheet1", "B1")
		b2, _ := f.GetCellValue("Sheet1", "B2")
		c1, _ := f.GetCellValue("Sheet1", "C1")

		if b1 != "200" {
			t.Errorf("expected B1=200, got %s", b1)
		}
		if b2 != "400" {
			t.Errorf("expected B2=400, got %s", b2)
		}
		if c1 != "300" {
			t.Errorf("expected C1=300, got %s", c1)
		}
	})

	t.Run("CrossSheetDependency", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has source values
		f.SetCellValue("Data", "A1", 100)
		f.SetCellValue("Data", "A2", 200)

		// Sheet1 references Data sheet
		f.SetCellFormula("Sheet1", "B1", "=Data!A1*2")
		f.SetCellFormula("Sheet1", "B2", "=Data!A2*2")
		f.SetCellFormula("Sheet1", "C1", "=B1+B2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Update Data sheet values
		values := []CellUpdate{
			{Sheet: "Data", Cell: "A1", Value: 500},
			{Sheet: "Data", Cell: "A2", Value: 600},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, nil)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify Sheet1 formulas are recalculated
		b1, _ := f.GetCellValue("Sheet1", "B1")
		b2, _ := f.GetCellValue("Sheet1", "B2")
		c1, _ := f.GetCellValue("Sheet1", "C1")

		if b1 != "1000" {
			t.Errorf("expected Sheet1!B1=1000, got %s", b1)
		}
		if b2 != "1200" {
			t.Errorf("expected Sheet1!B2=1200, got %s", b2)
		}
		if c1 != "2200" {
			t.Errorf("expected Sheet1!C1=2200, got %s", c1)
		}
	})

	t.Run("ChainDependency", func(t *testing.T) {
		f := NewFile()

		// Create chain dependency: A1 -> B1 -> C1 -> D1 -> E1
		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "B1", "=A1+1")
		f.SetCellFormula("Sheet1", "C1", "=B1+1")
		f.SetCellFormula("Sheet1", "D1", "=C1+1")
		f.SetCellFormula("Sheet1", "E1", "=D1+1")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Update A1
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, nil)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify entire chain
		b1, _ := f.GetCellValue("Sheet1", "B1")
		c1, _ := f.GetCellValue("Sheet1", "C1")
		d1, _ := f.GetCellValue("Sheet1", "D1")
		e1, _ := f.GetCellValue("Sheet1", "E1")

		if b1 != "101" {
			t.Errorf("expected B1=101, got %s", b1)
		}
		if c1 != "102" {
			t.Errorf("expected C1=102, got %s", c1)
		}
		if d1 != "103" {
			t.Errorf("expected D1=103, got %s", d1)
		}
		if e1 != "104" {
			t.Errorf("expected E1=104, got %s", e1)
		}
	})

	t.Run("EmptyUpdates", func(t *testing.T) {
		f := NewFile()

		// Empty updates should return nil without error
		err := f.BatchUpdateValuesAndFormulasWithRecalc(nil, nil)
		if err != nil {
			t.Fatalf("empty updates should not fail: %v", err)
		}

		err = f.BatchUpdateValuesAndFormulasWithRecalc([]CellUpdate{}, []FormulaUpdate{})
		if err != nil {
			t.Fatalf("empty slices should not fail: %v", err)
		}
	})

	t.Run("FormulaOnlyUpdate", func(t *testing.T) {
		// 测试只传公式更新的场景
		// 这个场景之前有 bug：公式设置后值为空，导致依赖它的单元格计算错误
		f := NewFile()

		// 设置初始数据
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2") // 初始公式
		f.SetCellFormula("Sheet1", "C1", "=B1+5") // 依赖 B1

		// 初始计算
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if b1Before != "20" {
			t.Fatalf("expected B1=20 before update, got %s", b1Before)
		}
		if c1Before != "25" {
			t.Fatalf("expected C1=25 before update, got %s", c1Before)
		}

		// 只传公式，不传值：将 B1 的公式改为固定值 100
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=100"},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(nil, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// 验证 B1 的值不为空
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "100" {
			t.Errorf("expected B1=100 after formula update, got '%s'", b1After)
		}

		// 验证依赖 B1 的 C1 也被正确重算
		c1After, _ := f.GetCellValue("Sheet1", "C1")
		if c1After != "105" {
			t.Errorf("expected C1=105 after B1 formula update, got '%s'", c1After)
		}
	})

	t.Run("FormulaOnlyUpdateWithDependentIF", func(t *testing.T) {
		// 测试带 IF 条件的依赖场景
		// 模拟实际场景：IF(IFERROR(INDEX(...),0)=0,"断货",...)
		f := NewFile()

		// 设置初始数据
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellFormula("Sheet1", "B1", "=A1") // 初始公式
		// C1 依赖 B1，如果 B1=0 则显示"断货"
		f.SetCellFormula("Sheet1", "C1", `=IF(B1=0,"断货",B1*2)`)

		// 初始计算
		f.RecalculateAllWithDependency()

		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if c1Before != "200" {
			t.Fatalf("expected C1=200 before update, got %s", c1Before)
		}

		// 只传公式更新 B1
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=50"},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(nil, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// B1 应该有值 50，不是空
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "50" {
			t.Errorf("expected B1=50, got '%s'", b1After)
		}

		// C1 应该是 100，不是"断货"
		c1After, _ := f.GetCellValue("Sheet1", "C1")
		if c1After != "100" {
			t.Errorf("expected C1=100 (50*2), got '%s' - formula may have seen empty B1", c1After)
		}
	})

	t.Run("FormulaOnlyUpdateCrossSheet", func(t *testing.T) {
		// 测试跨 sheet 的公式更新场景
		f := NewFile()
		f.NewSheet("Data")

		// Data!A1 有公式
		f.SetCellFormula("Data", "A1", "=10+5")
		// Sheet1!B1 依赖 Data!A1
		f.SetCellFormula("Sheet1", "B1", "=Data!A1*2")

		// 初始计算
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "30" {
			t.Fatalf("expected Sheet1!B1=30 before update, got %s", b1Before)
		}

		// 只更新 Data!A1 的公式
		formulas := []FormulaUpdate{
			{Sheet: "Data", Cell: "A1", Formula: "=100"},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(nil, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Data!A1 应该有值 100
		dataA1, _ := f.GetCellValue("Data", "A1")
		if dataA1 != "100" {
			t.Errorf("expected Data!A1=100, got '%s'", dataA1)
		}

		// Sheet1!B1 应该重算为 200
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "200" {
			t.Errorf("expected Sheet1!B1=200 after Data!A1 formula update, got '%s'", b1After)
		}
	})

	t.Run("MixedValueAndFormulaUpdate", func(t *testing.T) {
		// 测试同时传值和公式的场景
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")
		f.SetCellFormula("Sheet1", "C1", "=B1+A1")

		f.RecalculateAllWithDependency()

		// 同时更新值和公式
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 50},
		}
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*3"}, // 改为 *3
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// A1 = 50 (值更新)
		a1, _ := f.GetCellValue("Sheet1", "A1")
		if a1 != "50" {
			t.Errorf("expected A1=50, got '%s'", a1)
		}

		// B1 = 150 (公式 =A1*3 = 50*3)
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "150" {
			t.Errorf("expected B1=150, got '%s'", b1)
		}

		// C1 = 200 (公式 =B1+A1 = 150+50)
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "200" {
			t.Errorf("expected C1=200, got '%s'", c1)
		}
	})

	t.Run("PreCalculatedValueNotOverwritten", func(t *testing.T) {
		// 测试预计算值不被增量重算覆盖的场景
		// 场景：同时在 valueUpdates 和 formulaUpdates 中传入同一个单元格
		// valueUpdates 中的值作为预计算值，不应该被重算覆盖
		f := NewFile()

		// 初始设置
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "C1", "=B1+5") // C1 依赖 B1

		// 同时传入 B1 的值（预计算值）和公式
		// 预计算值故意设为 9999（不是正确的 200），用于验证不被重算覆盖
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},  // A1 改为 100
			{Sheet: "Sheet1", Cell: "B1", Value: 9999}, // B1 的预计算值
		}
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"}, // B1 公式，正确值应该是 200
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// B1 应该保持预计算值 9999，不被重算为 200
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "9999" {
			t.Errorf("expected B1=9999 (pre-calculated value), got '%s' - value was overwritten by recalculation", b1)
		}

		// B1 的公式应该正确设置
		b1Formula, _ := f.GetCellFormula("Sheet1", "B1")
		if b1Formula != "=A1*2" {
			t.Errorf("expected B1 formula '=A1*2', got '%s'", b1Formula)
		}

		// C1 应该使用 B1 的预计算值进行重算: 9999 + 5 = 10004
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "10004" {
			t.Errorf("expected C1=10004 (using B1's pre-calculated value), got '%s'", c1)
		}
	})

	t.Run("PreCalculatedValueChainDependency", func(t *testing.T) {
		// 测试链式依赖中多个预计算值的场景
		// A1 -> B1 -> C1 -> D1
		// B1 和 C1 都有预计算值
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "D1", "=C1+1") // D1 依赖 C1

		// B1 和 C1 都有预计算值
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
			{Sheet: "Sheet1", Cell: "B1", Value: 500},  // 预计算值（正确值应该是 200）
			{Sheet: "Sheet1", Cell: "C1", Value: 1000}, // 预计算值（正确值应该是 505）
		}
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"}, // 正确值 200
			{Sheet: "Sheet1", Cell: "C1", Formula: "=B1+5"}, // 正确值 505
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// B1 保持预计算值 500
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "500" {
			t.Errorf("expected B1=500 (pre-calculated), got '%s'", b1)
		}

		// C1 保持预计算值 1000
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "1000" {
			t.Errorf("expected C1=1000 (pre-calculated), got '%s'", c1)
		}

		// D1 使用 C1 的预计算值重算: 1000 + 1 = 1001
		d1, _ := f.GetCellValue("Sheet1", "D1")
		if d1 != "1001" {
			t.Errorf("expected D1=1001 (using C1's pre-calculated value), got '%s'", d1)
		}
	})

	t.Run("PreCalculatedValueCrossSheet", func(t *testing.T) {
		// 测试跨 sheet 的预计算值场景
		f := NewFile()
		f.NewSheet("Data")

		f.SetCellValue("Data", "A1", 10)
		f.SetCellFormula("Sheet1", "C1", "=Data!B1*10") // Sheet1!C1 依赖 Data!B1

		// Data!B1 有预计算值
		values := []CellUpdate{
			{Sheet: "Data", Cell: "A1", Value: 100},
			{Sheet: "Data", Cell: "B1", Value: 888}, // 预计算值
		}
		formulas := []FormulaUpdate{
			{Sheet: "Data", Cell: "B1", Formula: "=A1*2"}, // 正确值应该是 200
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Data!B1 保持预计算值 888
		b1, _ := f.GetCellValue("Data", "B1")
		if b1 != "888" {
			t.Errorf("expected Data!B1=888 (pre-calculated), got '%s'", b1)
		}

		// Sheet1!C1 使用 Data!B1 的预计算值重算: 888 * 10 = 8880
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "8880" {
			t.Errorf("expected Sheet1!C1=8880 (using Data!B1's pre-calculated value), got '%s'", c1)
		}
	})

	t.Run("MixedPreCalculatedAndNormalFormulas", func(t *testing.T) {
		// 测试混合场景：部分公式有预计算值，部分没有
		f := NewFile()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "D1", "=B1+C1") // D1 依赖 B1 和 C1

		// B1 有预计算值，C1 没有（需要计算）
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
			{Sheet: "Sheet1", Cell: "B1", Value: 777}, // B1 预计算值
		}
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"}, // B1 有预计算值，不重算
			{Sheet: "Sheet1", Cell: "C1", Formula: "=A1*3"}, // C1 没有预计算值，需要计算
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// B1 保持预计算值 777
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "777" {
			t.Errorf("expected B1=777 (pre-calculated), got '%s'", b1)
		}

		// C1 没有预计算值，需要计算: 100 * 3 = 300
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "300" {
			t.Errorf("expected C1=300 (calculated), got '%s'", c1)
		}

		// D1 使用 B1 的预计算值和 C1 的计算值: 777 + 300 = 1077
		d1, _ := f.GetCellValue("Sheet1", "D1")
		if d1 != "1077" {
			t.Errorf("expected D1=1077, got '%s'", d1)
		}
	})
}

// TestColumnRangeDependencyInIncrementalRecalc tests that formulas referencing column ranges
// (like $B:$B) are correctly recalculated when data in that column is updated.
// This is a critical test for the fix that ensures column range dependencies are always tracked,
// even for pure data columns (columns without formulas).
func TestColumnRangeDependencyInIncrementalRecalc(t *testing.T) {
	t.Run("PureDataColumnWithINDEX_MATCH", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has pure data (no formulas) in column A (SKU) and column B (values)
		skus := []string{"SKU001", "SKU002", "SKU003", "SKU004", "SKU005"}
		values := []int{100, 200, 300, 400, 500}
		for i, sku := range skus {
			f.SetCellValue("Data", "A"+string(rune('1'+i)), sku)
			f.SetCellValue("Data", "B"+string(rune('1'+i)), values[i])
		}

		// Sheet1 has formulas that reference Data sheet using column ranges
		// INDEX($B:$B, MATCH(..., $A:$A, 0))
		f.SetCellValue("Sheet1", "A1", "SKU003")
		f.SetCellFormula("Sheet1", "B1", "=INDEX(Data!$B:$B,MATCH(A1,Data!$A:$A,0))")

		f.SetCellValue("Sheet1", "A2", "SKU001")
		f.SetCellFormula("Sheet1", "B2", "=INDEX(Data!$B:$B,MATCH(A2,Data!$A:$A,0))")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Verify initial values
		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		b2Before, _ := f.GetCellValue("Sheet1", "B2")
		if b1Before != "300" {
			t.Fatalf("expected Sheet1!B1=300 before update, got %s", b1Before)
		}
		if b2Before != "100" {
			t.Fatalf("expected Sheet1!B2=100 before update, got %s", b2Before)
		}

		// Update Data!B3 (SKU003's value) from 300 to 999
		f.SetCellValue("Data", "B3", 999)

		// Use incremental recalculation
		updatedCells := map[string]bool{
			"Data!B3": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify Sheet1!B1 was recalculated (it references SKU003 which was updated)
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "999" {
			t.Errorf("expected Sheet1!B1=999 after update, got %s (formula depends on Data!$B:$B column range)", b1After)
		}

		// Verify Sheet1!B2 is still correct (it references SKU001 which was NOT updated)
		b2After, _ := f.GetCellValue("Sheet1", "B2")
		if b2After != "100" {
			t.Errorf("expected Sheet1!B2=100 (unchanged), got %s", b2After)
		}
	})

	t.Run("PureDataColumnWithSUMIFS", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet: Category (A), Date (B), Value (C)
		data := [][]interface{}{
			{"Cat1", "2024-01-01", 100},
			{"Cat1", "2024-01-02", 200},
			{"Cat2", "2024-01-01", 300},
			{"Cat2", "2024-01-02", 400},
		}
		for i, row := range data {
			f.SetCellValue("Data", "A"+string(rune('1'+i)), row[0])
			f.SetCellValue("Data", "B"+string(rune('1'+i)), row[1])
			f.SetCellValue("Data", "C"+string(rune('1'+i)), row[2])
		}

		// Sheet1: SUMIFS referencing column ranges
		f.SetCellValue("Sheet1", "A1", "Cat1")
		f.SetCellFormula("Sheet1", "B1", "=SUMIFS(Data!$C:$C,Data!$A:$A,A1)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Verify initial value: Cat1 has 100+200=300
		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "300" {
			t.Fatalf("expected Sheet1!B1=300 before update, got %s", b1Before)
		}

		// Update Data!C1 (Cat1's first value) from 100 to 500
		f.SetCellValue("Data", "C1", 500)

		// Use incremental recalculation
		updatedCells := map[string]bool{
			"Data!C1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify SUMIFS was recalculated: Cat1 now has 500+200=700
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "700" {
			t.Errorf("expected Sheet1!B1=700 after update, got %s (SUMIFS depends on Data!$C:$C column range)", b1After)
		}
	})

	t.Run("MultiColumnRangeReference", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Source")

		// Source sheet has data in columns A, B, C (all pure data, no formulas)
		f.SetCellValue("Source", "A1", 10)
		f.SetCellValue("Source", "B1", 20)
		f.SetCellValue("Source", "C1", 30)
		f.SetCellValue("Source", "A2", 40)
		f.SetCellValue("Source", "B2", 50)
		f.SetCellValue("Source", "C2", 60)

		// Sheet1 uses SUMPRODUCT with column ranges
		f.SetCellFormula("Sheet1", "D1", "=SUM(Source!$A:$A)")
		f.SetCellFormula("Sheet1", "E1", "=SUM(Source!$B:$B)")
		f.SetCellFormula("Sheet1", "F1", "=SUM(Source!$C:$C)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Verify initial values
		d1, _ := f.GetCellValue("Sheet1", "D1")
		e1, _ := f.GetCellValue("Sheet1", "E1")
		f1Val, _ := f.GetCellValue("Sheet1", "F1")
		if d1 != "50" || e1 != "70" || f1Val != "90" {
			t.Fatalf("expected D1=50, E1=70, F1=90 before update, got D1=%s, E1=%s, F1=%s", d1, e1, f1Val)
		}

		// Update only Source!A1
		f.SetCellValue("Source", "A1", 100)

		updatedCells := map[string]bool{
			"Source!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify D1 was recalculated (depends on column A)
		d1After, _ := f.GetCellValue("Sheet1", "D1")
		if d1After != "140" { // 100 + 40 = 140
			t.Errorf("expected D1=140 after update, got %s", d1After)
		}

		// Verify E1 and F1 are still correct (they don't depend on column A)
		e1After, _ := f.GetCellValue("Sheet1", "E1")
		f1After, _ := f.GetCellValue("Sheet1", "F1")
		if e1After != "70" {
			t.Errorf("expected E1=70 (unchanged), got %s", e1After)
		}
		if f1After != "90" {
			t.Errorf("expected F1=90 (unchanged), got %s", f1After)
		}
	})

	t.Run("ColumnRangeWithMixedContent", func(t *testing.T) {
		f := NewFile()

		// Sheet1 has mixed content: some data cells, some formula cells
		f.SetCellValue("Sheet1", "A1", 10)         // data
		f.SetCellValue("Sheet1", "A2", 20)         // data
		f.SetCellFormula("Sheet1", "A3", "=A1+A2") // formula
		f.SetCellValue("Sheet1", "A4", 40)         // data

		// B1 uses column range reference to A
		f.SetCellFormula("Sheet1", "B1", "=SUM($A:$A)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// A3 should be 30, B1 should be 10+20+30+40=100
		a3, _ := f.GetCellValue("Sheet1", "A3")
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if a3 != "30" {
			t.Fatalf("expected A3=30, got %s", a3)
		}
		if b1 != "100" {
			t.Fatalf("expected B1=100 before update, got %s", b1)
		}

		// Update A1 (should affect both A3 and B1)
		f.SetCellValue("Sheet1", "A1", 100)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// A3 should be 100+20=120, B1 should be 100+20+120+40=280
		a3After, _ := f.GetCellValue("Sheet1", "A3")
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if a3After != "120" {
			t.Errorf("expected A3=120 after update, got %s", a3After)
		}
		if b1After != "280" {
			t.Errorf("expected B1=280 after update (depends on $A:$A column range), got %s", b1After)
		}
	})
}

// TestExtractDependenciesOptimizedColumnRange tests that extractDependenciesOptimized
// correctly adds column dependencies for column range references.
func TestExtractDependenciesOptimizedColumnRange(t *testing.T) {
	t.Run("ColumnRangeAlwaysAddsDependency", func(t *testing.T) {
		// Formula: =INDEX(Data!$B:$B,MATCH(A1,Data!$A:$A,0))
		// Should add COLUMN:Data!B and COLUMN:Data!A as dependencies
		formula := "INDEX(Data!$B:$B,MATCH(A1,Data!$A:$A,0))"
		deps := extractDependenciesOptimized(formula, "Sheet1", "B1", nil, nil)

		hasColB := false
		hasColA := false
		for _, dep := range deps {
			if dep == "COLUMN:Data!B" {
				hasColB = true
			}
			if dep == "COLUMN:Data!A" {
				hasColA = true
			}
		}

		if !hasColB {
			t.Errorf("expected COLUMN:Data!B in dependencies, got %v", deps)
		}
		if !hasColA {
			t.Errorf("expected COLUMN:Data!A in dependencies, got %v", deps)
		}
	})

	t.Run("ColumnRangeWithEmptyMetadata", func(t *testing.T) {
		// Even with empty columnMetadata, column range should add dependency
		formula := "SUM(Source!$C:$C)"
		columnMetadata := make(map[string]*columnMeta)
		// Don't add any metadata for Source!C - simulating a pure data column

		deps := extractDependenciesOptimized(formula, "Sheet1", "A1", nil, columnMetadata)

		hasColC := false
		for _, dep := range deps {
			if dep == "COLUMN:Source!C" {
				hasColC = true
			}
		}

		if !hasColC {
			t.Errorf("expected COLUMN:Source!C in dependencies even for pure data column, got %v", deps)
		}
	})

	t.Run("MultiColumnRangeReference", func(t *testing.T) {
		// Formula with multiple column range references: =SUMIFS($H:$H,$A:$A,B1,$C:$C,D1)
		formula := "SUMIFS($H:$H,$A:$A,B1,$C:$C,D1)"
		deps := extractDependenciesOptimized(formula, "Sheet1", "E1", nil, nil)

		expectedCols := []string{"COLUMN:Sheet1!H", "COLUMN:Sheet1!A", "COLUMN:Sheet1!C"}
		for _, expected := range expectedCols {
			found := false
			for _, dep := range deps {
				if dep == expected {
					found = true
					break
				}
			}
			if !found {
				t.Errorf("expected %s in dependencies, got %v", expected, deps)
			}
		}
	})
}

// TestConcurrentFormulaCalculationRace tests that parallel formula calculation
// does not cause data races when multiple goroutines access worksheet data.
// This test should be run with -race flag: go test -race -run TestConcurrentFormulaCalculationRace
func TestConcurrentFormulaCalculationRace(t *testing.T) {
	t.Run("RecalculateAllWithDependency_NoRace", func(t *testing.T) {
		// Create a file with multiple formulas that depend on each other
		f := NewFile()

		// Set up data
		for i := 1; i <= 10; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i*10)
		}

		// Set up formulas with dependencies
		// B column depends on A column
		for i := 1; i <= 10; i++ {
			f.SetCellFormula("Sheet1", "B"+itoa(i), "=A"+itoa(i)+"*2")
		}

		// C column depends on B column
		for i := 1; i <= 10; i++ {
			f.SetCellFormula("Sheet1", "C"+itoa(i), "=B"+itoa(i)+"+10")
		}

		// D column uses SUM across multiple cells
		f.SetCellFormula("Sheet1", "D1", "=SUM(A1:A10)")
		f.SetCellFormula("Sheet1", "D2", "=SUM(B1:B10)")
		f.SetCellFormula("Sheet1", "D3", "=SUM(C1:C10)")

		// Run recalculation - this should not cause race conditions
		err := f.RecalculateAllWithDependency()
		if err != nil {
			t.Fatalf("RecalculateAllWithDependency failed: %v", err)
		}

		// Verify results
		d1, _ := f.GetCellValue("Sheet1", "D1")
		if d1 != "550" { // 10+20+30+...+100 = 550
			t.Errorf("expected D1=550, got %s", d1)
		}
	})

	t.Run("RecalculateAllWithDependency_CrossSheet_NoRace", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")
		f.NewSheet("Calc")

		// Set up data in Data sheet
		for i := 1; i <= 20; i++ {
			f.SetCellValue("Data", "A"+itoa(i), i)
			f.SetCellValue("Data", "B"+itoa(i), i*2)
		}

		// Set up cross-sheet formulas in Calc sheet
		for i := 1; i <= 10; i++ {
			f.SetCellFormula("Calc", "A"+itoa(i), "=Data!A"+itoa(i)+"*3")
			f.SetCellFormula("Calc", "B"+itoa(i), "=Data!B"+itoa(i)+"+Calc!A"+itoa(i))
		}

		// Run recalculation - cross-sheet access should not cause race
		err := f.RecalculateAllWithDependency()
		if err != nil {
			t.Fatalf("RecalculateAllWithDependency failed: %v", err)
		}
	})

	t.Run("BatchUpdateValuesAndFormulasWithRecalc_Concurrent_NoRace", func(t *testing.T) {
		f := NewFile()

		// Set up initial data
		for i := 1; i <= 10; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i)
			f.SetCellFormula("Sheet1", "B"+itoa(i), "=A"+itoa(i)+"*2")
			f.SetCellFormula("Sheet1", "C"+itoa(i), "=B"+itoa(i)+"+A"+itoa(i))
		}

		// Initial calculation
		f.RecalculateAllWithDependency()

		// Batch update with value and formula changes
		valueUpdates := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
			{Sheet: "Sheet1", Cell: "A2", Value: 200},
		}
		formulaUpdates := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*10"},
			{Sheet: "Sheet1", Cell: "B4", Formula: "=A4*10"},
		}

		// This should not cause race conditions
		err := f.BatchUpdateValuesAndFormulasWithRecalc(valueUpdates, formulaUpdates)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify A1 update propagated
		c1, _ := f.GetCellValue("Sheet1", "C1")
		// C1 = B1 + A1 = (A1*2) + A1 = 100*2 + 100 = 300
		if c1 != "300" {
			t.Errorf("expected C1=300 after update, got %s", c1)
		}
	})

	t.Run("PreCalculateSimpleFormulas_Parallel_NoRace", func(t *testing.T) {
		f := NewFile()

		// Create many simple formulas to trigger parallel calculation
		for i := 1; i <= 50; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i)
		}

		// Simple formulas that will be pre-calculated in parallel
		for i := 1; i <= 50; i++ {
			// Simple arithmetic formulas (not SUMIFS/INDEX-MATCH)
			f.SetCellFormula("Sheet1", "B"+itoa(i), "=A"+itoa(i)+"*2+1")
		}

		// Dependent formulas
		for i := 1; i <= 50; i++ {
			f.SetCellFormula("Sheet1", "C"+itoa(i), "=B"+itoa(i)+"+A"+itoa(i))
		}

		// Run recalculation - parallel pre-calculation should not race
		err := f.RecalculateAllWithDependency()
		if err != nil {
			t.Fatalf("RecalculateAllWithDependency failed: %v", err)
		}

		// Verify a few results
		// B1 = A1*2+1 = 1*2+1 = 3
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "3" {
			t.Errorf("expected B1=3, got %s", b1)
		}

		// C1 = B1 + A1 = 3 + 1 = 4
		c1, _ := f.GetCellValue("Sheet1", "C1")
		if c1 != "4" {
			t.Errorf("expected C1=4, got %s", c1)
		}
	})

	t.Run("ArrayFormulaCells_PreCheck_NoRace", func(t *testing.T) {
		// Test that setArrayFormulaCells is called before parallel calculation
		// to avoid race conditions when getCellFormulaReadOnly triggers it
		f := NewFile()

		// Set up data and formulas
		for i := 1; i <= 20; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i)
			f.SetCellFormula("Sheet1", "B"+itoa(i), "=A"+itoa(i)+"*2")
		}

		// First recalculation should set formulaChecked = true
		err := f.RecalculateAllWithDependency()
		if err != nil {
			t.Fatalf("First RecalculateAllWithDependency failed: %v", err)
		}

		// Modify some values
		for i := 1; i <= 20; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i*10)
		}

		// Second recalculation should not trigger setArrayFormulaCells again
		// because formulaChecked is already true
		err = f.RecalculateAllWithDependency()
		if err != nil {
			t.Fatalf("Second RecalculateAllWithDependency failed: %v", err)
		}
	})

	t.Run("MultipleRecalculations_NoRace", func(t *testing.T) {
		// Test multiple sequential recalculations don't cause race
		f := NewFile()

		for i := 1; i <= 30; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i)
			f.SetCellFormula("Sheet1", "B"+itoa(i), "=A"+itoa(i)+"*2")
			f.SetCellFormula("Sheet1", "C"+itoa(i), "=B"+itoa(i)+"+A"+itoa(i))
		}

		// Multiple recalculations
		for round := 0; round < 3; round++ {
			err := f.RecalculateAllWithDependency()
			if err != nil {
				t.Fatalf("RecalculateAllWithDependency round %d failed: %v", round, err)
			}

			// Update some values between rounds
			for i := 1; i <= 10; i++ {
				f.SetCellValue("Sheet1", "A"+itoa(i), (round+1)*i*10)
			}
		}
	})
}

func TestRecalculateSheetWithDependency(t *testing.T) {
	t.Run("BasicFormulas", func(t *testing.T) {
		f := NewFile()
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellFormula("Sheet1", "B1", "A1*2")
		f.SetCellFormula("Sheet1", "B2", "A2+B1")

		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("RecalculateSheetWithDependency failed: %v", err)
		}

		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "20" {
			t.Errorf("expected B1=20, got %s", b1)
		}
		b2, _ := f.GetCellValue("Sheet1", "B2")
		if b2 != "40" {
			t.Errorf("expected B2=40, got %s", b2)
		}
	})

	t.Run("CrossSheetDataRead", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has source values (no formulas)
		f.SetCellValue("Data", "A1", 100)
		f.SetCellValue("Data", "A2", 200)

		// Sheet1 references Data sheet
		f.SetCellFormula("Sheet1", "B1", "Data!A1+Data!A2")
		f.SetCellFormula("Sheet1", "B2", "B1*2")

		// Only recalculate Sheet1 (Data sheet values are read as external data)
		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("RecalculateSheetWithDependency failed: %v", err)
		}

		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "300" {
			t.Errorf("expected B1=300, got %s", b1)
		}
		b2, _ := f.GetCellValue("Sheet1", "B2")
		if b2 != "600" {
			t.Errorf("expected B2=600, got %s", b2)
		}
	})

	t.Run("OnlyTargetSheetRecalculated", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Sheet2")

		// Both sheets have formulas
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "A1*3")
		f.SetCellValue("Sheet2", "A1", 5)
		f.SetCellFormula("Sheet2", "B1", "A1*7")

		// Only recalculate Sheet1
		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("RecalculateSheetWithDependency failed: %v", err)
		}

		// Sheet1 formula should be recalculated
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "30" {
			t.Errorf("expected Sheet1!B1=30, got %s", b1)
		}

		// Sheet2 formula should NOT be recalculated (still empty/original)
		b1s2, _ := f.GetCellValue("Sheet2", "B1")
		if b1s2 == "35" {
			t.Errorf("Sheet2!B1 should not have been recalculated by RecalculateSheetWithDependency(Sheet1)")
		}
	})

	t.Run("ChainDependency", func(t *testing.T) {
		f := NewFile()
		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "B1", "A1+1")
		f.SetCellFormula("Sheet1", "C1", "B1+1")
		f.SetCellFormula("Sheet1", "D1", "C1+1")

		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("RecalculateSheetWithDependency failed: %v", err)
		}

		d1, _ := f.GetCellValue("Sheet1", "D1")
		if d1 != "4" {
			t.Errorf("expected D1=4, got %s", d1)
		}
	})

	t.Run("UpdateAndRecalculate", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		f.SetCellValue("Data", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "Data!A1*2")

		// First recalculate
		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("first RecalculateSheetWithDependency failed: %v", err)
		}

		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "20" {
			t.Errorf("expected B1=20, got %s", b1)
		}

		// Update source value and recalculate again
		f.SetCellValue("Data", "A1", 50)
		err = f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("second RecalculateSheetWithDependency failed: %v", err)
		}

		b1, _ = f.GetCellValue("Sheet1", "B1")
		if b1 != "100" {
			t.Errorf("expected B1=100 after update, got %s", b1)
		}
	})

	t.Run("NonExistentSheet", func(t *testing.T) {
		f := NewFile()
		err := f.RecalculateSheetWithDependency("NonExistent")
		if err == nil {
			t.Errorf("expected error for non-existent sheet")
		}
	})

	t.Run("EmptySheet", func(t *testing.T) {
		f := NewFile()
		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Errorf("expected no error for empty sheet, got %v", err)
		}
	})

	t.Run("SUMFormula", func(t *testing.T) {
		f := NewFile()
		for i := 1; i <= 10; i++ {
			f.SetCellValue("Sheet1", "A"+itoa(i), i)
		}
		f.SetCellFormula("Sheet1", "B1", "SUM(A1:A10)")

		err := f.RecalculateSheetWithDependency("Sheet1")
		if err != nil {
			t.Fatalf("RecalculateSheetWithDependency failed: %v", err)
		}

		b1, _ := f.GetCellValue("Sheet1", "B1")
		if b1 != "55" {
			t.Errorf("expected B1=55, got %s", b1)
		}
	})
}

// itoa converts int to string for cell references
func itoa(i int) string {
	if i == 0 {
		return "0"
	}
	result := ""
	for i > 0 {
		result = string(rune('0'+i%10)) + result
		i /= 10
	}
	return result
}

func TestBatchUpdateAndRecalculateNilValueDoesNotBecomeNilString(t *testing.T) {
	f := NewFile()
	defer f.Close()

	if err := f.SetCellValue("Sheet1", "A1", "seed"); err != nil {
		t.Fatalf("SetCellValue: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `=A1&""`); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B1", I: 1}},
	}

	if err := f.BatchUpdateAndRecalculate([]CellUpdate{{Sheet: "Sheet1", Cell: "A1", Value: nil}}); err != nil {
		t.Fatalf("BatchUpdateAndRecalculate: %v", err)
	}

	got, err := f.GetCellValue("Sheet1", "B1", Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if got != "" {
		t.Fatalf("expected empty string, got %q", got)
	}
}

func TestStringifyCellUpdateValueNilSafe(t *testing.T) {
	if got := stringifyCellUpdateValue(nil); got != "" {
		t.Fatalf("expected empty string for nil, got %q", got)
	}

	var ptr *string
	if got := stringifyCellUpdateValue(ptr); got != "" {
		t.Fatalf("expected empty string for typed nil, got %q", got)
	}

	if got := stringifyCellUpdateValue(123); got != "123" {
		t.Fatalf("expected numeric string, got %q", got)
	}
}

func TestSetCellValueTypedNilDoesNotBecomeNilString(t *testing.T) {
	f := NewFile()
	defer f.Close()

	var ptr *string
	if err := f.SetCellValue("Sheet1", "A1", ptr); err != nil {
		t.Fatalf("SetCellValue: %v", err)
	}

	got, err := f.GetCellValue("Sheet1", "A1", Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("GetCellValue: %v", err)
	}
	if got != "" {
		t.Fatalf("expected empty string, got %q", got)
	}
}
