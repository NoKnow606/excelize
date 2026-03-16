package excelize

import "testing"

func TestDependencyGraphCacheReuseAndInvalidation(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("SetCellValue(A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+1"); err != nil {
		t.Fatalf("SetCellFormula(B1): %v", err)
	}

	graph1 := f.buildDependencyGraph()
	graph2 := f.buildDependencyGraph()
	if graph1 != graph2 {
		t.Fatalf("expected dependency graph cache reuse")
	}

	if err := f.SetCellValue("Sheet1", "A1", 20); err != nil {
		t.Fatalf("SetCellValue(A1) second time: %v", err)
	}
	graph3 := f.buildDependencyGraph()
	if graph3 != graph1 {
		t.Fatalf("value-only update should not invalidate dependency graph")
	}

	if err := f.SetCellFormula("Sheet1", "B1", "=A1+2"); err != nil {
		t.Fatalf("SetCellFormula(B1) update: %v", err)
	}
	graph4 := f.buildDependencyGraph()
	if graph4 == graph3 {
		t.Fatalf("formula update should invalidate dependency graph")
	}
	if node := graph4.nodes["Sheet1!B1"]; node == nil || node.formula != "=A1+2" {
		t.Fatalf("expected updated formula in dependency graph, got %#v", node)
	}

	if err := f.SetCellValue("Sheet1", "B1", 7); err != nil {
		t.Fatalf("SetCellValue(B1): %v", err)
	}
	graph5 := f.buildDependencyGraph()
	if graph5 == graph4 {
		t.Fatalf("removing formula should invalidate dependency graph")
	}
	if _, ok := graph5.nodes["Sheet1!B1"]; ok {
		t.Fatalf("expected formula node to be removed after SetCellValue")
	}
}

func TestDependencyGraphCacheInvalidatesOnSheetRename(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("SetCellValue(A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+1"); err != nil {
		t.Fatalf("SetCellFormula(B1): %v", err)
	}

	graph1 := f.buildDependencyGraph()
	if err := f.SetSheetName("Sheet1", "Data"); err != nil {
		t.Fatalf("SetSheetName(): %v", err)
	}

	graph2 := f.buildDependencyGraph()
	if graph2 == graph1 {
		t.Fatalf("sheet rename should invalidate dependency graph")
	}
	if _, ok := graph2.nodes["Data!B1"]; !ok {
		t.Fatalf("expected renamed sheet node in dependency graph")
	}
}

func TestRecalculateAllWithDependencySkipsWhenUnchanged(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("SetCellValue(A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+1"); err != nil {
		t.Fatalf("SetCellFormula(B1): %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency(): %v", err)
	}

	f.calcCache.Store("custom-skip-sentinel", "keep")
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency() second time: %v", err)
	}

	if got, ok := f.calcCache.Load("custom-skip-sentinel"); !ok || got != "keep" {
		t.Fatalf("expected unchanged recalculation to skip cache clearing, got ok=%v value=%v", ok, got)
	}
}

func TestRecalculateAllWithDependencyUsesAffectedSubgraph(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	must := func(err error) {
		t.Helper()
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
	}

	must(f.SetCellValue("Sheet1", "A1", 10))
	must(f.SetCellFormula("Sheet1", "B1", "=A1+1"))
	must(f.SetCellFormula("Sheet1", "C1", "=B1*2"))
	must(f.SetCellValue("Sheet1", "D1", 3))
	must(f.SetCellFormula("Sheet1", "E1", "=D1+5"))

	must(f.RecalculateAllWithDependency())

	must(f.SetCellValue("Sheet1", "A1", 20))
	f.calcCache.Store("Sheet1!E1!raw=false", "sentinel")
	must(f.RecalculateAllWithDependency())

	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "21" {
		t.Fatalf("GetCellValue(B1): got=%q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "C1"); err != nil || got != "42" {
		t.Fatalf("GetCellValue(C1): got=%q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "E1"); err != nil || got != "8" {
		t.Fatalf("GetCellValue(E1): got=%q err=%v", got, err)
	}

	if got, ok := f.calcCache.Load("Sheet1!E1!raw=false"); !ok || got != "sentinel" {
		t.Fatalf("expected unaffected formula cache to remain during subgraph recalc, got ok=%v value=%v", ok, got)
	}
}

func TestRecalculateAllWithDependencyUsesAffectedSubgraphForColumnRanges(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	must := func(err error) {
		t.Helper()
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
	}

	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	must(f.SetCellValue("Data", "A2", "SKU1"))
	must(f.SetCellValue("Data", "B2", 10))
	must(f.SetCellValue("Sheet1", "A1", "SKU1"))
	must(f.SetCellFormula("Sheet1", "B1", `=IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`))
	must(f.RecalculateAllWithDependency())

	must(f.SetCellValue("Data", "B2", 20))
	must(f.RecalculateAllWithDependency())

	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "20" {
		t.Fatalf("GetCellValue(B1): got=%q err=%v", got, err)
	}
}

func TestSetCellValueInvalidatesOnlyAffectedFormulaCaches(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	must := func(err error) {
		t.Helper()
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
	}

	must(f.SetCellValue("Sheet1", "A1", 10))
	must(f.SetCellFormula("Sheet1", "B1", "=A1+1"))
	must(f.SetCellFormula("Sheet1", "C1", "=B1*2"))
	must(f.SetCellValue("Sheet1", "D1", 3))
	must(f.SetCellFormula("Sheet1", "E1", "=D1+5"))
	must(f.RecalculateAllWithDependency())

	f.calcCache.Store("Sheet1!B1!raw=false", "affected")
	f.calcCache.Store("Sheet1!E1!raw=false", "unaffected")

	must(f.SetCellValue("Sheet1", "A1", 20))

	if _, ok := f.calcCache.Load("Sheet1!B1!raw=false"); ok {
		t.Fatalf("expected affected formula cache to be invalidated on input change")
	}
	if got, ok := f.calcCache.Load("Sheet1!E1!raw=false"); !ok || got != "unaffected" {
		t.Fatalf("expected unaffected formula cache to remain, got ok=%v value=%v", ok, got)
	}
}
