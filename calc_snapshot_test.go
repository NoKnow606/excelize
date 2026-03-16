package excelize

import (
	"bytes"
	"context"
	"strconv"
	"sync/atomic"
	"testing"
)

func TestBuildLoadCalculationSnapshot(t *testing.T) {
	f1 := NewFile()
	t.Cleanup(func() { _ = f1.Close() })
	if _, err := f1.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	if err := f1.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Data!A2): %v", err)
	}
	if err := f1.SetCellValue("Data", "B2", 100); err != nil {
		t.Fatalf("SetCellValue(Data!B2): %v", err)
	}
	if err := f1.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A1): %v", err)
	}
	if err := f1.SetCellFormula("Sheet1", "B1", `=IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B1): %v", err)
	}
	if err := f1.SetCellFormula("Sheet1", "C1", `=B1&"-done"`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!C1): %v", err)
	}

	data, err := f1.BuildCalculationSnapshot()
	if err != nil {
		t.Fatalf("BuildCalculationSnapshot(): %v", err)
	}
	if len(data) == 0 {
		t.Fatal("expected non-empty calculation snapshot")
	}

	f2 := NewFile()
	t.Cleanup(func() { _ = f2.Close() })
	if _, err := f2.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	if err := f2.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Data!A2): %v", err)
	}
	if err := f2.SetCellValue("Data", "B2", 100); err != nil {
		t.Fatalf("SetCellValue(Data!B2): %v", err)
	}
	if err := f2.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A1): %v", err)
	}
	if err := f2.SetCellFormula("Sheet1", "B1", `=IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B1): %v", err)
	}
	if err := f2.SetCellFormula("Sheet1", "C1", `=B1&"-done"`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!C1): %v", err)
	}

	if err := f2.LoadCalculationSnapshot(data); err != nil {
		t.Fatalf("LoadCalculationSnapshot(): %v", err)
	}
	graph := f2.getCachedDependencyGraph()
	if graph == nil {
		t.Fatal("expected dependency graph cache from snapshot")
	}
	if node := graph.nodes["Sheet1!B1"]; node == nil || node.formula == "" {
		t.Fatalf("expected Sheet1!B1 node in snapshot graph, got %#v", node)
	}
	if got := f2.buildDependencyGraph(); got != graph {
		t.Fatal("expected buildDependencyGraph to reuse imported snapshot graph")
	}
}

func TestEmbedCalculationSnapshotAutoLoadOnOpenReader(t *testing.T) {
	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Data!A2): %v", err)
	}
	if err := f.SetCellValue("Data", "B2", 100); err != nil {
		t.Fatalf("SetCellValue(Data!B2): %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `=IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B1): %v", err)
	}
	if err := f.EmbedCalculationSnapshot(); err != nil {
		t.Fatalf("EmbedCalculationSnapshot(): %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer(): %v", err)
	}
	reopened, err := OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("OpenReader(): %v", err)
	}
	t.Cleanup(func() { _ = reopened.Close() })

	graph := reopened.getCachedDependencyGraph()
	if graph == nil {
		t.Fatal("expected embedded snapshot to auto-load on open")
	}
	if got := reopened.buildDependencyGraph(); got != graph {
		t.Fatal("expected buildDependencyGraph to reuse auto-loaded snapshot graph")
	}
}

func TestCalculationSnapshotInvalidatedOnFormulaChange(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })
	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("SetCellValue(A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+1"); err != nil {
		t.Fatalf("SetCellFormula(B1): %v", err)
	}
	if err := f.EmbedCalculationSnapshot(); err != nil {
		t.Fatalf("EmbedCalculationSnapshot(): %v", err)
	}
	if _, ok := f.Pkg.Load(calcSnapshotWorkbookPath); !ok {
		t.Fatal("expected embedded calculation snapshot to exist")
	}

	if err := f.SetCellFormula("Sheet1", "B1", "=A1+2"); err != nil {
		t.Fatalf("SetCellFormula(B1) update: %v", err)
	}
	if _, ok := f.Pkg.Load(calcSnapshotWorkbookPath); ok {
		t.Fatal("expected embedded calculation snapshot to be removed after formula change")
	}
}

func TestCalculationSnapshotAutoEmbedOnWrite(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Data!A2): %v", err)
	}
	if err := f.SetCellValue("Data", "B2", 100); err != nil {
		t.Fatalf("SetCellValue(Data!B2): %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `=IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B1): %v", err)
	}

	f.SetCalculationSnapshotAutoEmbed(true)
	if !f.CalculationSnapshotAutoEmbedEnabled() {
		t.Fatal("expected calculation snapshot auto embed to be enabled")
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer(): %v", err)
	}
	reopened, err := OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("OpenReader(): %v", err)
	}
	t.Cleanup(func() { _ = reopened.Close() })

	if !reopened.CalculationSnapshotAutoEmbedEnabled() {
		t.Fatal("expected reopened workbook to keep snapshot auto embed enabled")
	}
	if graph := reopened.getCachedDependencyGraph(); graph == nil {
		t.Fatal("expected auto-embedded snapshot to auto-load on open")
	}
}

func TestCalculationSnapshotAutoEmbedRefreshesAfterFormulaChangeOnWrite(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })
	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("SetCellValue(A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+1"); err != nil {
		t.Fatalf("SetCellFormula(B1): %v", err)
	}
	f.SetCalculationSnapshotAutoEmbed(true)
	if _, err := f.WriteToBuffer(); err != nil {
		t.Fatalf("WriteToBuffer() first: %v", err)
	}
	if _, ok := f.Pkg.Load(calcSnapshotWorkbookPath); !ok {
		t.Fatal("expected snapshot to exist after auto-embed write")
	}

	if err := f.SetCellFormula("Sheet1", "B1", "=A1+2"); err != nil {
		t.Fatalf("SetCellFormula(B1) update: %v", err)
	}
	if _, ok := f.Pkg.Load(calcSnapshotWorkbookPath); ok {
		t.Fatal("expected dirty formula change to clear embedded snapshot before next write")
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer() second: %v", err)
	}
	reopened, err := OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("OpenReader(): %v", err)
	}
	t.Cleanup(func() { _ = reopened.Close() })

	graph := reopened.getCachedDependencyGraph()
	if graph == nil {
		t.Fatal("expected refreshed snapshot to auto-load on reopen")
	}
	node := graph.nodes["Sheet1!B1"]
	if node == nil || node.formula != "=A1+2" {
		t.Fatalf("expected refreshed snapshot formula, got %#v", node)
	}
}

func TestCalculationSnapshotAutoEmbedReusesCachedSnapshotAcrossWrites(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })
	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("SetCellValue(A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+1"); err != nil {
		t.Fatalf("SetCellFormula(B1): %v", err)
	}
	f.SetCalculationSnapshotAutoEmbed(true)

	var builds atomic.Int32
	calcSnapshotBuildTestHook = func() {
		builds.Add(1)
	}
	t.Cleanup(func() {
		calcSnapshotBuildTestHook = nil
	})

	if _, err := f.WriteToBuffer(); err != nil {
		t.Fatalf("WriteToBuffer() first: %v", err)
	}
	if got := builds.Load(); got != 1 {
		t.Fatalf("unexpected snapshot build count after first write: got %d want 1", got)
	}

	if _, err := f.WriteToBuffer(); err != nil {
		t.Fatalf("WriteToBuffer() second: %v", err)
	}
	if got := builds.Load(); got != 1 {
		t.Fatalf("unexpected snapshot build count after second write: got %d want 1", got)
	}

	if err := f.SetCellValue("Sheet1", "A1", 20); err != nil {
		t.Fatalf("SetCellValue(A1) update: %v", err)
	}
	if _, err := f.WriteToBuffer(); err != nil {
		t.Fatalf("WriteToBuffer() after value update: %v", err)
	}
	if got := builds.Load(); got != 1 {
		t.Fatalf("unexpected snapshot build count after value-only write: got %d want 1", got)
	}

	if err := f.SetCellFormula("Sheet1", "B1", "=A1+2"); err != nil {
		t.Fatalf("SetCellFormula(B1) formula update: %v", err)
	}
	if _, err := f.WriteToBuffer(); err != nil {
		t.Fatalf("WriteToBuffer() after formula update: %v", err)
	}
	if got := builds.Load(); got != 2 {
		t.Fatalf("unexpected snapshot build count after formula write: got %d want 2", got)
	}
}

func TestCalculationSnapshotRestoresPostgresWarmupCachesAcrossReopen(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 3, 1, "SKU2")
	state.seedCellRow("book1", "Data", 3, 2, "200")

	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU2"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `=IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B1): %v", err)
	}

	opts := PGSyncOptions{WorkbookID: "book1"}
	cfg, err := f.getPGSyncOptions(opts)
	if err != nil {
		t.Fatalf("getPGSyncOptions(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}
	if err := f.PreloadPostgresLookupCache(context.Background()); err != nil {
		t.Fatalf("PreloadPostgresLookupCache(): %v", err)
	}

	matchKey := pgExactMatchCacheKey(cfg, pgLookupRange{
		Sheet:    "Data",
		StartRow: 1,
		EndRow:   TotalRows,
		StartCol: 1,
		EndCol:   1,
	}, f.getSheetVersion("Data"), pgMirrorValueTypeString, "SKU2")
	cellKey := pgCellValueCacheKey(cfg, "Data", f.getSheetVersion("Data"), 3, 2)
	if cached, ok := f.pgCalcCache.Load(matchKey); !ok {
		t.Fatalf("expected warmed exact-match cache key %q", matchKey)
	} else if result, ok := cached.(pgExactMatchResult); !ok || !result.Found || result.RowNum != 3 || result.ColNum != 1 {
		t.Fatalf("unexpected warmed exact-match cache value: %#v", cached)
	}
	if cached, ok := f.pgCalcCache.Load(cellKey); !ok {
		t.Fatalf("expected warmed cell-value cache key %q", cellKey)
	} else if value, ok := cached.(pgMirrorCellValue); !ok || value.Value != "200" || value.ValueType != pgMirrorValueTypeNumber {
		t.Fatalf("unexpected warmed cell-value cache value: %#v", cached)
	}
	if got := countPGSnapshotCalcCacheEntries(f); got == 0 {
		t.Fatal("expected PG calc cache entries to be warmed")
	}

	if err := f.EmbedCalculationSnapshot(); err != nil {
		t.Fatalf("EmbedCalculationSnapshot(): %v", err)
	}
	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer(): %v", err)
	}

	reopened, err := OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("OpenReader(): %v", err)
	}
	t.Cleanup(func() { _ = reopened.Close() })
	if err := reopened.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("reopened EnablePostgresMirror(): %v", err)
	}
	if cached, ok := reopened.pgCalcCache.Load(matchKey); !ok {
		t.Fatalf("expected embedded snapshot to restore exact-match cache key %q", matchKey)
	} else if result, ok := cached.(pgExactMatchResult); !ok || !result.Found || result.RowNum != 3 || result.ColNum != 1 {
		t.Fatalf("unexpected restored exact-match cache value: %#v", cached)
	}
	if cached, ok := reopened.pgCalcCache.Load(cellKey); !ok {
		t.Fatalf("expected embedded snapshot to restore cell-value cache key %q", cellKey)
	} else if value, ok := cached.(pgMirrorCellValue); !ok || value.Value != "200" || value.ValueType != pgMirrorValueTypeNumber {
		t.Fatalf("unexpected restored cell-value cache value: %#v", cached)
	}
	if got := countPGSnapshotCalcCacheEntries(reopened); got == 0 {
		t.Fatal("expected embedded snapshot to restore PG calc cache entries")
	}
	if cached, ok := reopened.loadPGWholeCellCache("Sheet1!B1"); !ok {
		t.Fatal("expected embedded snapshot to restore whole-cell PG cache")
	} else if cached.Value() != "200" {
		t.Fatalf("unexpected restored whole-cell PG cache value: %#v", cached)
	}

	state.reset()
	var (
		runtimeExactKey    string
		runtimeExactHit    bool
		runtimeExactLoaded bool
		runtimeExactType   string
		runtimeCellKey     string
		runtimeCellHit     bool
		runtimeCellLoaded  bool
		runtimeCellType    string
	)
	pgExactMatchCacheTestHook = func(key string, hit bool, loaded bool, valueType string) {
		runtimeExactKey = key
		runtimeExactHit = hit
		runtimeExactLoaded = loaded
		runtimeExactType = valueType
	}
	pgCellValueCacheTestHook = func(key string, hit bool, loaded bool, valueType string) {
		runtimeCellKey = key
		runtimeCellHit = hit
		runtimeCellLoaded = loaded
		runtimeCellType = valueType
	}
	t.Cleanup(func() {
		pgExactMatchCacheTestHook = nil
		pgCellValueCacheTestHook = nil
	})
	got, err := reopened.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue(): %v", err)
	}
	if got != "200" {
		t.Fatalf("unexpected CalcCellValue result: got %q want %q", got, "200")
	}
	if got := state.countQueries("SELECT row_num"); got != 0 {
		currentExact, _ := reopened.pgCalcCache.Load(matchKey)
		wholeCell, _ := reopened.loadPGWholeCellCache("Sheet1!B1")
		class := getPGLookupOptimizationClass("Sheet1", `IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`)
		t.Fatalf("unexpected exact-match query count after reopen with snapshot: got %d want 0, whole=%t wholeCell=%#v exactKey=%q exactHit=%t exactLoaded=%t exactType=%q currentExact=%#v",
			got, class.whole, wholeCell, runtimeExactKey, runtimeExactHit, runtimeExactLoaded, runtimeExactType, currentExact)
	}
	if got := state.countQueries(`FROM "public"."excelize_cell_mirror"`); got != 0 {
		currentCell, _ := reopened.pgCalcCache.Load(cellKey)
		t.Fatalf("unexpected cell query count after reopen with snapshot: got %d want 0, cellKey=%q cellHit=%t cellLoaded=%t cellType=%q currentCell=%#v",
			got, runtimeCellKey, runtimeCellHit, runtimeCellLoaded, runtimeCellType, currentCell)
	}
}

func TestCalculationSnapshotRestoresPostgresWarmupCachesAcrossBatchLookup(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedLookupRow("book1", "Data", 3, 1, "SKU2")
	state.seedCellRow("book1", "Data", 2, 2, "100")
	state.seedCellRow("book1", "Data", 3, 2, "200")

	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data): %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A1): %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A2", "SKU2"); err != nil {
		t.Fatalf("SetCellValue(Sheet1!A2): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B1): %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B2", `IFERROR(VLOOKUP(A2,Data!A:B,2,FALSE),"")`); err != nil {
		t.Fatalf("SetCellFormula(Sheet1!B2): %v", err)
	}

	opts := PGSyncOptions{WorkbookID: "book1"}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}
	if err := f.PreloadPostgresLookupCache(context.Background()); err != nil {
		t.Fatalf("PreloadPostgresLookupCache(): %v", err)
	}
	if err := f.EmbedCalculationSnapshot(); err != nil {
		t.Fatalf("EmbedCalculationSnapshot(): %v", err)
	}

	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer(): %v", err)
	}
	reopened, err := OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatalf("OpenReader(): %v", err)
	}
	t.Cleanup(func() { _ = reopened.Close() })
	if err := reopened.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("reopened EnablePostgresMirror(): %v", err)
	}
	if _, ok := reopened.loadPGWholeCellCache("Sheet1!B1"); !ok {
		t.Fatal("expected restored whole-cell cache for Sheet1!B1")
	}
	if _, ok := reopened.loadPGWholeCellCache("Sheet1!B2"); !ok {
		t.Fatal("expected restored whole-cell cache for Sheet1!B2")
	}

	state.reset()
	results := reopened.batchCalculatePostgresLookupsWithCache(map[string]string{
		"Sheet1!B1": `IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`,
		"Sheet1!B2": `IFERROR(VLOOKUP(A2,Data!A:B,2,FALSE),"")`,
	}, nil)
	if got := results["Sheet1!B1"]; got != "100" {
		t.Fatalf("unexpected batch result for Sheet1!B1: got %q want %q", got, "100")
	}
	if got := results["Sheet1!B2"]; got != "200" {
		t.Fatalf("unexpected batch result for Sheet1!B2: got %q want %q", got, "200")
	}
	if got := state.countQueries("SELECT row_num"); got != 0 {
		t.Fatalf("unexpected exact-match query count after batch restore: got %d want 0", got)
	}
	if got := state.countQueries(`FROM "public"."excelize_cell_mirror"`); got != 0 {
		t.Fatalf("unexpected cell query count after batch restore: got %d want 0", got)
	}
}

func countPGSnapshotCalcCacheEntries(f *File) int {
	count := 0
	f.pgCalcCache.Range(func(key, value interface{}) bool {
		count++
		return true
	})
	return count
}

func BenchmarkCalculationSnapshotOpenAndBuildGraph(b *testing.B) {
	silenceBenchmarkLogs(b)

	newWorkbookBytes := func(embed bool) []byte {
		b.Helper()
		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(Data): %v", err)
		}
		for i := 1; i <= 1000; i++ {
			row := i + 1
			if err := f.SetCellValue("Data", "A"+calcSnapshotItoa(row), "SKU"+calcSnapshotItoa(i)); err != nil {
				b.Fatalf("SetCellValue(Data!A%d): %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+calcSnapshotItoa(row), i); err != nil {
				b.Fatalf("SetCellValue(Data!B%d): %v", row, err)
			}
			if err := f.SetCellValue("Sheet1", "A"+calcSnapshotItoa(i), "SKU"+calcSnapshotItoa(i)); err != nil {
				b.Fatalf("SetCellValue(Sheet1!A%d): %v", i, err)
			}
			if err := f.SetCellFormula("Sheet1", "B"+calcSnapshotItoa(i), `=IFERROR(INDEX(Data!B:B,MATCH(A`+calcSnapshotItoa(i)+`,Data!A:A,0)),"")`); err != nil {
				b.Fatalf("SetCellFormula(Sheet1!B%d): %v", i, err)
			}
		}
		if embed {
			if err := f.EmbedCalculationSnapshot(); err != nil {
				b.Fatalf("EmbedCalculationSnapshot(): %v", err)
			}
		}
		buf, err := f.WriteToBuffer()
		if err != nil {
			b.Fatalf("WriteToBuffer(): %v", err)
		}
		return buf.Bytes()
	}

	noSnapshot := newWorkbookBytes(false)
	withSnapshot := newWorkbookBytes(true)

	b.Run("OpenAndBuildGraph1000NoSnapshot", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f, err := OpenReader(bytes.NewReader(noSnapshot))
			if err != nil {
				b.Fatalf("OpenReader(): %v", err)
			}
			graph := f.buildDependencyGraph()
			if graph == nil || len(graph.nodes) == 0 {
				b.Fatalf("expected dependency graph")
			}
			_ = f.Close()
		}
	})

	b.Run("OpenAndBuildGraph1000EmbeddedSnapshot", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f, err := OpenReader(bytes.NewReader(withSnapshot))
			if err != nil {
				b.Fatalf("OpenReader(): %v", err)
			}
			graph := f.buildDependencyGraph()
			if graph == nil || len(graph.nodes) == 0 {
				b.Fatalf("expected dependency graph")
			}
			_ = f.Close()
		}
	})
}

func BenchmarkCalculationSnapshotWriteToBuffer(b *testing.B) {
	silenceBenchmarkLogs(b)

	newWorkbook := func(auto bool) *File {
		b.Helper()
		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(Data): %v", err)
		}
		for i := 1; i <= 1000; i++ {
			row := i + 1
			if err := f.SetCellValue("Data", "A"+calcSnapshotItoa(row), "SKU"+calcSnapshotItoa(i)); err != nil {
				b.Fatalf("SetCellValue(Data!A%d): %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+calcSnapshotItoa(row), i); err != nil {
				b.Fatalf("SetCellValue(Data!B%d): %v", row, err)
			}
			if err := f.SetCellValue("Sheet1", "A"+calcSnapshotItoa(i), "SKU"+calcSnapshotItoa(i)); err != nil {
				b.Fatalf("SetCellValue(Sheet1!A%d): %v", i, err)
			}
			if err := f.SetCellFormula("Sheet1", "B"+calcSnapshotItoa(i), `=IFERROR(INDEX(Data!B:B,MATCH(A`+calcSnapshotItoa(i)+`,Data!A:A,0)),"")`); err != nil {
				b.Fatalf("SetCellFormula(Sheet1!B%d): %v", i, err)
			}
		}
		f.SetCalculationSnapshotAutoEmbed(auto)
		return f
	}

	b.Run("WriteToBuffer1000NoAutoEmbed", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f := newWorkbook(false)
			if _, err := f.WriteToBuffer(); err != nil {
				b.Fatalf("WriteToBuffer(): %v", err)
			}
			_ = f.Close()
		}
	})

	b.Run("WriteToBuffer1000AutoEmbed", func(b *testing.B) {
		for i := 0; i < b.N; i++ {
			f := newWorkbook(true)
			if _, err := f.WriteToBuffer(); err != nil {
				b.Fatalf("WriteToBuffer(): %v", err)
			}
			_ = f.Close()
		}
	})

	b.Run("WriteToBufferNonDestructive1000AutoEmbedRepeated", func(b *testing.B) {
		f := newWorkbook(true)
		defer func() { _ = f.Close() }()

		if _, err := f.WriteToBufferNonDestructive(); err != nil {
			b.Fatalf("WriteToBufferNonDestructive() warmup: %v", err)
		}
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if _, err := f.WriteToBufferNonDestructive(); err != nil {
				b.Fatalf("WriteToBufferNonDestructive(): %v", err)
			}
		}
	})
}

func calcSnapshotItoa(v int) string {
	return strconv.Itoa(v)
}
