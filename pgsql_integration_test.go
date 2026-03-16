package excelize

import (
	"bytes"
	"context"
	"database/sql"
	"io"
	"log"
	"net/url"
	"os"
	"strconv"
	"strings"
	"testing"
	"time"

	_ "github.com/jackc/pgx/v5/stdlib"
)

func openPostgresIntegrationDB(tb testing.TB) (*sql.DB, PGSyncOptions) {
	tb.Helper()

	dsn := strings.TrimSpace(os.Getenv("DATABASE_URL"))
	if dsn == "" {
		tb.Skip("DATABASE_URL is not set")
	}

	parsed, err := url.Parse(dsn)
	if err != nil {
		tb.Fatalf("parse DATABASE_URL: %v", err)
	}

	schema := parsed.Query().Get("schema")
	if schema == "" {
		schema = "public"
	}
	query := parsed.Query()
	query.Del("schema")
	if query.Get("sslmode") == "" {
		query.Set("sslmode", "disable")
	}
	if query.Get("connect_timeout") == "" {
		query.Set("connect_timeout", "5")
	}
	parsed.RawQuery = query.Encode()

	db, err := sql.Open("pgx", parsed.String())
	if err != nil {
		tb.Fatalf("sql.Open(): %v", err)
	}
	tb.Cleanup(func() {
		_ = db.Close()
	})

	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	if err := db.PingContext(ctx); err != nil {
		tb.Fatalf("PingContext(): %v", err)
	}

	prefix := "excelize_it_" + strconv.FormatInt(time.Now().UnixNano(), 10)
	opts := PGSyncOptions{
		WorkbookID:  prefix + "_book",
		Schema:      schema,
		TablePrefix: prefix,
		BatchSize:   128,
	}

	tables := getPGMirrorTables(opts.TablePrefix)
	tb.Cleanup(func() {
		cleanupCtx, cleanupCancel := context.WithTimeout(context.Background(), 5*time.Second)
		defer cleanupCancel()
		for _, table := range []string{tables.lookup, tables.cell, tables.sheet, tables.workbook} {
			_, _ = db.ExecContext(cleanupCtx, `DROP TABLE IF EXISTS `+qualifyPGTable(opts.Schema, table))
		}
	})

	return db, opts
}

func silenceBenchmarkLogs(tb testing.TB) {
	tb.Helper()

	originalWriter := log.Writer()
	originalFlags := log.Flags()
	originalPrefix := log.Prefix()
	log.SetOutput(io.Discard)
	log.SetFlags(0)
	log.SetPrefix("")
	tb.Cleanup(func() {
		log.SetOutput(originalWriter)
		log.SetFlags(originalFlags)
		log.SetPrefix(originalPrefix)
	})
}

func buildPostgresBatchLookupBenchmarkFixture(tb testing.TB, db *sql.DB, opts PGSyncOptions, rows int) ([]byte, []byte, map[string]string) {
	tb.Helper()

	f := NewFile()
	tb.Cleanup(func() { _ = f.Close() })
	if _, err := f.NewSheet("Data"); err != nil {
		tb.Fatalf("NewSheet(): %v", err)
	}

	formulas := make(map[string]string, rows)
	for i := 1; i <= rows; i++ {
		row := i + 1
		sku := "SKU" + strconv.Itoa(i)
		value := strconv.Itoa(i * 100)
		if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
			tb.Fatalf("SetCellValue Data!A%d: %v", row, err)
		}
		if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
			tb.Fatalf("SetCellValue Data!B%d: %v", row, err)
		}
		if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
			tb.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
		}
		formula := `IFERROR(INDEX(Data!B:B,MATCH(A` + strconv.Itoa(i) + `,Data!A:A,0)),"")`
		if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), formula); err != nil {
			tb.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
		}
		formulas["Sheet1!B"+strconv.Itoa(i)] = formula
	}

	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		tb.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}

	noSnapshotBuf, err := f.WriteToBuffer()
	if err != nil {
		tb.Fatalf("WriteToBuffer() no snapshot: %v", err)
	}

	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		tb.Fatalf("EnablePostgresMirror(): %v", err)
	}
	if err := f.PreloadPostgresLookupCache(ctx); err != nil {
		tb.Fatalf("PreloadPostgresLookupCache(): %v", err)
	}
	if err := f.EmbedCalculationSnapshot(); err != nil {
		tb.Fatalf("EmbedCalculationSnapshot(): %v", err)
	}
	withSnapshotBuf, err := f.WriteToBuffer()
	if err != nil {
		tb.Fatalf("WriteToBuffer() with snapshot: %v", err)
	}

	return noSnapshotBuf.Bytes(), withSnapshotBuf.Bytes(), formulas
}

func TestPostgresMirrorIntegrationSyncAndLookupFastPath(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", "100"); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
	}
	if err := f.SetCellValue("Data", "A3", "SKU2"); err != nil {
		t.Fatalf("SetCellValue Data!A3: %v", err)
	}
	if err := f.SetCellValue("Data", "B3", "200"); err != nil {
		t.Fatalf("SetCellValue Data!B3: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU2"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "C1", `MATCH(A1,Data!A2:A3,0)`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!C1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "D1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!D1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "E1", `IF(IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")="200","hit","miss")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!E1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "F1", `IF("200"=IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),""),"left","right")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!F1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	if got, err := f.CalcCellValue("Sheet1", "B1"); err != nil || got != "200" {
		t.Fatalf("CalcCellValue B1: got %q err=%v", got, err)
	}
	if got, err := f.CalcCellValue("Sheet1", "C1"); err != nil || got != "2" {
		t.Fatalf("CalcCellValue C1: got %q err=%v", got, err)
	}
	if got, err := f.CalcCellValue("Sheet1", "D1"); err != nil || got != "200" {
		t.Fatalf("CalcCellValue D1: got %q err=%v", got, err)
	}
	if got, err := f.CalcCellValue("Sheet1", "E1"); err != nil || got != "hit" {
		t.Fatalf("CalcCellValue E1: got %q err=%v", got, err)
	}
	if got, err := f.CalcCellValue("Sheet1", "F1"); err != nil || got != "left" {
		t.Fatalf("CalcCellValue F1: got %q err=%v", got, err)
	}

	snapshot, err := f.LoadCellSnapshotFromPostgres(ctx, db, "Data", "B3", opts)
	if err != nil {
		t.Fatalf("LoadCellSnapshotFromPostgres(): %v", err)
	}
	if snapshot.Value != "200" || snapshot.Row != 3 || snapshot.Col != 2 {
		t.Fatalf("unexpected snapshot: %#v", snapshot)
	}
}

func TestPostgresMirrorIntegrationTypedExactLookup(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellInt("Data", "A2", 200); err != nil {
		t.Fatalf("SetCellInt Data!A2: %v", err)
	}
	if err := f.SetCellStr("Data", "A3", "200"); err != nil {
		t.Fatalf("SetCellStr Data!A3: %v", err)
	}
	if err := f.SetCellStr("Data", "B2", "num"); err != nil {
		t.Fatalf("SetCellStr Data!B2: %v", err)
	}
	if err := f.SetCellStr("Data", "B3", "str"); err != nil {
		t.Fatalf("SetCellStr Data!B3: %v", err)
	}
	if err := f.SetCellInt("Sheet1", "A1", 200); err != nil {
		t.Fatalf("SetCellInt Sheet1!A1: %v", err)
	}
	if err := f.SetCellStr("Sheet1", "A2", "200"); err != nil {
		t.Fatalf("SetCellStr Sheet1!A2: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `VLOOKUP(A1,Data!A2:B3,2,FALSE)`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B2", `VLOOKUP(A2,Data!A2:B3,2,FALSE)`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B2: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	if got, err := f.CalcCellValue("Sheet1", "B1"); err != nil || got != "num" {
		t.Fatalf("CalcCellValue B1: got %q err=%v", got, err)
	}
	if got, err := f.CalcCellValue("Sheet1", "B2"); err != nil || got != "str" {
		t.Fatalf("CalcCellValue B2: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationSyncWorkbookUsesPGXCopyPath(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	for i := 1; i <= 8; i++ {
		row := i + 1
		if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), "SKU"+strconv.Itoa(i)); err != nil {
			t.Fatalf("SetCellValue Data!A%d: %v", row, err)
		}
		if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), strconv.Itoa(i*10)); err != nil {
			t.Fatalf("SetCellValue Data!B%d: %v", row, err)
		}
	}

	copySheets := make([]string, 0, 2)
	copyRows := make([]int, 0, 2)
	pgMirrorCopyFromTestHook = func(sheet string, rowCount int) {
		copySheets = append(copySheets, sheet)
		copyRows = append(copyRows, rowCount)
	}
	t.Cleanup(func() {
		pgMirrorCopyFromTestHook = nil
	})

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}

	if len(copySheets) == 0 {
		t.Fatalf("expected pgx copy path hook to run")
	}
	foundData := false
	for i, sheet := range copySheets {
		if sheet == "Data" && copyRows[i] == 16 {
			foundData = true
			break
		}
	}
	if !foundData {
		t.Fatalf("expected Data sheet to use copy path, got sheets=%v rows=%v", copySheets, copyRows)
	}
}

func TestPostgresMirrorIntegrationRefreshAndRecalculate(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", "10"); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "C1", `B1&"-done"`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!C1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}

	tables := getPGMirrorTables(opts.TablePrefix)
	if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.cell)+`
SET value = $1, updated_at = NOW()
WHERE workbook_id = $2 AND sheet_name = $3 AND cell_ref = $4
`, "999", opts.WorkbookID, "Data", "B2"); err != nil {
		t.Fatalf("update cell mirror: %v", err)
	}
	if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.lookup)+`
SET cell_value = $1, updated_at = NOW()
WHERE workbook_id = $2 AND sheet_name = $3 AND cell_ref = $4
`, "999", opts.WorkbookID, "Data", "B2"); err != nil {
		t.Fatalf("update lookup mirror: %v", err)
	}

	refreshed, err := f.RefreshCellFromPostgresAndRecalculate(ctx, db, "Data", "B2", opts)
	if err != nil {
		t.Fatalf("RefreshCellFromPostgresAndRecalculate(): %v", err)
	}
	if refreshed != "999" {
		t.Fatalf("unexpected refreshed value: got %q, want %q", refreshed, "999")
	}

	if got, err := f.GetCellValue("Data", "B2"); err != nil || got != "999" {
		t.Fatalf("GetCellValue Data!B2: got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "999" {
		t.Fatalf("GetCellValue Sheet1!B1: got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "C1"); err != nil || got != "999-done" {
		t.Fatalf("GetCellValue Sheet1!C1: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationRefreshRepairsLookupMirrorFromCellSnapshot(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", "10"); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	tables := getPGMirrorTables(opts.TablePrefix)
	if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.cell)+`
SET value = $1, value_type = $2, updated_at = NOW()
WHERE workbook_id = $3 AND sheet_name = $4 AND cell_ref = $5
`, "888", pgMirrorValueTypeNumber, opts.WorkbookID, "Data", "B2"); err != nil {
		t.Fatalf("update cell mirror: %v", err)
	}

	refreshed, err := f.RefreshCellFromPostgresAndRecalculate(ctx, db, "Data", "B2", opts)
	if err != nil {
		t.Fatalf("RefreshCellFromPostgresAndRecalculate(): %v", err)
	}
	if refreshed != "888" {
		t.Fatalf("unexpected refreshed value: got %q, want %q", refreshed, "888")
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "888" {
		t.Fatalf("GetCellValue Sheet1!B1: got %q err=%v", got, err)
	}

	var mirrored string
	if err := db.QueryRowContext(ctx, `
SELECT cell_value
FROM `+qualifyPGTable(opts.Schema, tables.lookup)+`
WHERE workbook_id = $1 AND sheet_name = $2 AND cell_ref = $3
`, opts.WorkbookID, "Data", "B2").Scan(&mirrored); err != nil {
		t.Fatalf("query lookup mirror: %v", err)
	}
	if mirrored != "888" {
		t.Fatalf("lookup mirror not repaired: got %q want %q", mirrored, "888")
	}
}

func TestPostgresMirrorIntegrationEmbeddedSnapshotReopenAndRefreshFromRealPG(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellInt("Data", "B2", 100); err != nil {
		t.Fatalf("SetCellInt Data!B2: %v", err)
	}
	if err := f.SetCellValue("Data", "A3", "SKU2"); err != nil {
		t.Fatalf("SetCellValue Data!A3: %v", err)
	}
	if err := f.SetCellInt("Data", "B3", 200); err != nil {
		t.Fatalf("SetCellInt Data!B3: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU2"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}
	if err := f.PreloadPostgresLookupCache(ctx); err != nil {
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

	tables := getPGMirrorTables(opts.TablePrefix)
	if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.cell)+`
SET value = $1, value_type = $2, updated_at = NOW()
WHERE workbook_id = $3 AND sheet_name = $4 AND cell_ref = $5
`, "999", pgMirrorValueTypeNumber, opts.WorkbookID, "Data", "B3"); err != nil {
		t.Fatalf("update cell mirror: %v", err)
	}
	if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.lookup)+`
SET cell_value = $1, cell_value_type = $2, updated_at = NOW()
WHERE workbook_id = $3 AND sheet_name = $4 AND cell_ref = $5
`, "999", pgMirrorValueTypeNumber, opts.WorkbookID, "Data", "B3"); err != nil {
		t.Fatalf("update lookup mirror: %v", err)
	}

	snapshot, err := reopened.LoadCellSnapshotFromPostgres(ctx, db, "Data", "B3", opts)
	if err != nil {
		t.Fatalf("LoadCellSnapshotFromPostgres(): %v", err)
	}
	if snapshot.Value != "999" || snapshot.Row != 3 || snapshot.Col != 2 {
		t.Fatalf("unexpected PG snapshot after reopen: %#v", snapshot)
	}
	if got, err := reopened.CalcCellValue("Sheet1", "B1"); err != nil || got != "200" {
		t.Fatalf("CalcCellValue B1 with embedded warm cache: got %q err=%v", got, err)
	}

	refreshed, err := reopened.RefreshCellFromPostgresAndRecalculate(ctx, db, "Data", "B3", opts)
	if err != nil {
		t.Fatalf("RefreshCellFromPostgresAndRecalculate(): %v", err)
	}
	if refreshed != "999" {
		t.Fatalf("unexpected refreshed value: got %q want %q", refreshed, "999")
	}
	if got, err := reopened.CalcCellValue("Sheet1", "B1"); err != nil || got != "999" {
		t.Fatalf("CalcCellValue B1 after refresh: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationBatchLookupRecalculate(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}

	for i := 1; i <= 12; i++ {
		row := i + 1
		sku := "SKU" + strconv.Itoa(i)
		value := strconv.Itoa(i * 100)
		if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
			t.Fatalf("SetCellValue Data!A%d: %v", row, err)
		}
		if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
			t.Fatalf("SetCellValue Data!B%d: %v", row, err)
		}
		if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
			t.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
		}
		if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(i)+`,Data!A:A,0)),"")`); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
		}
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency(): %v", err)
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "100" {
		t.Fatalf("GetCellValue Sheet1!B1: got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B12"); err != nil || got != "1200" {
		t.Fatalf("GetCellValue Sheet1!B12: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationBatchLookupSubExprRecalculate(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}

	for i := 1; i <= 12; i++ {
		row := i + 1
		sku := "SKU" + strconv.Itoa(i)
		value := strconv.Itoa(i * 100)
		if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
			t.Fatalf("SetCellValue Data!A%d: %v", row, err)
		}
		if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
			t.Fatalf("SetCellValue Data!B%d: %v", row, err)
		}
		if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
			t.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
		}
		formula := `IF(IFERROR(INDEX(Data!B:B,MATCH(A` + strconv.Itoa(i) + `,Data!A:A,0)),"")="` + value + `","yes","no")`
		if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), formula); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
		}
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency(): %v", err)
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "yes" {
		t.Fatalf("GetCellValue Sheet1!B1: got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B12"); err != nil || got != "yes" {
		t.Fatalf("GetCellValue Sheet1!B12: got %q err=%v", got, err)
	}
	if snapshot, err := f.LoadCellSnapshotFromPostgres(ctx, db, "Sheet1", "B1", opts); err != nil || snapshot.Value != "yes" {
		t.Fatalf("LoadCellSnapshotFromPostgres Sheet1!B1: snapshot=%#v err=%v", snapshot, err)
	}
	if snapshot, err := f.LoadCellSnapshotFromPostgres(ctx, db, "Sheet1", "B12", opts); err != nil || snapshot.Value != "yes" {
		t.Fatalf("LoadCellSnapshotFromPostgres Sheet1!B12: snapshot=%#v err=%v", snapshot, err)
	}
}

func TestPostgresMirrorIntegrationAutoSyncInvalidatesWholeFormulaCache(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", "10"); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency initial: %v", err)
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "10" {
		t.Fatalf("GetCellValue Sheet1!B1 initial: got %q err=%v", got, err)
	}

	if err := f.SetCellValue("Data", "B2", "20"); err != nil {
		t.Fatalf("SetCellValue Data!B2 updated: %v", err)
	}
	snapshot, err := f.LoadCellSnapshotFromPostgres(ctx, db, "Data", "B2", opts)
	if err != nil {
		t.Fatalf("LoadCellSnapshotFromPostgres Data!B2: %v", err)
	}
	if snapshot.Value != "20" {
		t.Fatalf("unexpected mirrored value after auto sync: got %q want %q", snapshot.Value, "20")
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency updated: %v", err)
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "20" {
		t.Fatalf("GetCellValue Sheet1!B1 updated: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationEnableAutoWarmupPreloadsWholeFormulaCache(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", "200"); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	formula := `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`
	if err := f.SetCellFormula("Sheet1", "B1", formula); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}

	f.calcCache.Clear()
	opts.AutoWarmup = true
	opts.WarmupSheets = []string{"Sheet1"}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	cacheEntries := 0
	f.pgCalcCache.Range(func(key, _ interface{}) bool {
		cacheEntries++
		return true
	})
	if cacheEntries == 0 {
		t.Fatal("expected auto warmup to populate PG lookup caches")
	}

	f.calcCache.Clear()
	if got, err := f.CalcCellValue("Sheet1", "B1"); err != nil || got != "200" {
		t.Fatalf("CalcCellValue Sheet1!B1: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationEnableAsyncWarmupAndWait(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}
	if err := f.SetCellValue("Data", "A2", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Data!A2: %v", err)
	}
	if err := f.SetCellValue("Data", "B2", "200"); err != nil {
		t.Fatalf("SetCellValue Data!B2: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}

	opts.AutoWarmup = true
	opts.AsyncWarmup = true
	opts.WarmupSheets = []string{"Sheet1"}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}
	if err := f.WaitForPostgresWarmup(ctx); err != nil {
		t.Fatalf("WaitForPostgresWarmup(): %v", err)
	}

	status := f.GetPostgresWarmupStatus()
	if !status.Enabled || status.InProgress || !status.Async || status.CompletedAt.IsZero() || status.Err != nil {
		t.Fatalf("unexpected warmup status: %#v", status)
	}

	f.calcCache.Clear()
	if got, err := f.CalcCellValue("Sheet1", "B1"); err != nil || got != "200" {
		t.Fatalf("CalcCellValue Sheet1!B1: got %q err=%v", got, err)
	}
}

func TestPostgresMirrorIntegrationTypedSubExprRecalculate(t *testing.T) {
	db, opts := openPostgresIntegrationDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(): %v", err)
	}

	if err := f.SetCellStr("Data", "A2", "NUM"); err != nil {
		t.Fatalf("SetCellStr Data!A2: %v", err)
	}
	if err := f.SetCellStr("Data", "A3", "STR"); err != nil {
		t.Fatalf("SetCellStr Data!A3: %v", err)
	}
	if err := f.SetCellInt("Data", "B2", 200); err != nil {
		t.Fatalf("SetCellInt Data!B2: %v", err)
	}
	if err := f.SetCellStr("Data", "B3", "200"); err != nil {
		t.Fatalf("SetCellStr Data!B3: %v", err)
	}

	if err := f.SetCellStr("Sheet1", "A1", "NUM"); err != nil {
		t.Fatalf("SetCellStr Sheet1!A1: %v", err)
	}
	if err := f.SetCellStr("Sheet1", "A2", "STR"); err != nil {
		t.Fatalf("SetCellStr Sheet1!A2: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IF(TYPE(IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),""))=2,"string","number")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B1: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B2", `IF(TYPE(IFERROR(VLOOKUP(A2,Data!A:B,2,FALSE),""))=2,"string","number")`); err != nil {
		t.Fatalf("SetCellFormula Sheet1!B2: %v", err)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
		t.Fatalf("SyncWorkbookToPostgres(): %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, opts); err != nil {
		t.Fatalf("EnablePostgresMirror(): %v", err)
	}

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency(): %v", err)
	}
	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "number" {
		t.Fatalf("GetCellValue Sheet1!B1: got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B2"); err != nil || got != "string" {
		t.Fatalf("GetCellValue Sheet1!B2: got %q err=%v", got, err)
	}
}

func BenchmarkPostgresMirrorIntegration(b *testing.B) {
	b.Run("RealPostgresSyncWorkbook1000x2", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 1000; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), strconv.Itoa(i)); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
		}

		ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
		defer cancel()

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
		}
	})

	b.Run("RealPostgresFastPathIndexMatch", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 1000; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), strconv.Itoa(i)); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
		}
		if err := f.SetCellValue("Sheet1", "A1", "SKU900"); err != nil {
			b.Fatalf("SetCellValue Sheet1!A1: %v", err)
		}
		if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
			b.Fatalf("SetCellFormula Sheet1!B1: %v", err)
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}
		if _, err := f.CalcCellValue("Sheet1", "B1"); err != nil {
			b.Fatalf("warmup CalcCellValue(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.calcCache.Clear()
			if _, err := f.CalcCellValue("Sheet1", "B1"); err != nil {
				b.Fatalf("CalcCellValue(): %v", err)
			}
		}
	})

	b.Run("RealPostgresLoopSetCellValue100WithMirror", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		updatesA := make([]CellUpdate, 0, 100)
		updatesB := make([]CellUpdate, 0, 100)
		for i := 1; i <= 100; i++ {
			cell := "A" + strconv.Itoa(i)
			if err := f.SetCellValue("Sheet1", cell, "seed"); err != nil {
				b.Fatalf("SetCellValue Sheet1!%s: %v", cell, err)
			}
			updatesA = append(updatesA, CellUpdate{Sheet: "Sheet1", Cell: cell, Value: "v1"})
			updatesB = append(updatesB, CellUpdate{Sheet: "Sheet1", Cell: cell, Value: "v2"})
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			updates := updatesA
			if i%2 == 1 {
				updates = updatesB
			}
			for _, update := range updates {
				if err := f.SetCellValue(update.Sheet, update.Cell, update.Value); err != nil {
					b.Fatalf("SetCellValue %s!%s: %v", update.Sheet, update.Cell, err)
				}
			}
		}
	})

	b.Run("RealPostgresBatchSetCellValue100WithMirror", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		updatesA := make([]CellUpdate, 0, 100)
		updatesB := make([]CellUpdate, 0, 100)
		for i := 1; i <= 100; i++ {
			cell := "A" + strconv.Itoa(i)
			if err := f.SetCellValue("Sheet1", cell, "seed"); err != nil {
				b.Fatalf("SetCellValue Sheet1!%s: %v", cell, err)
			}
			updatesA = append(updatesA, CellUpdate{Sheet: "Sheet1", Cell: cell, Value: "v1"})
			updatesB = append(updatesB, CellUpdate{Sheet: "Sheet1", Cell: cell, Value: "v2"})
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			updates := updatesA
			if i%2 == 1 {
				updates = updatesB
			}
			if err := f.BatchSetCellValue(updates); err != nil {
				b.Fatalf("BatchSetCellValue(): %v", err)
			}
		}
	})

	b.Run("RealPostgresLoopSetCellFormula100WithMirror", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		formulasA := make([]FormulaUpdate, 0, 100)
		formulasB := make([]FormulaUpdate, 0, 100)
		for i := 1; i <= 100; i++ {
			if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), strconv.Itoa(i)); err != nil {
				b.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
			}
			cell := "B" + strconv.Itoa(i)
			formulasA = append(formulasA, FormulaUpdate{Sheet: "Sheet1", Cell: cell, Formula: "A" + strconv.Itoa(i) + "*2"})
			formulasB = append(formulasB, FormulaUpdate{Sheet: "Sheet1", Cell: cell, Formula: "A" + strconv.Itoa(i) + "*3"})
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			formulas := formulasA
			if i%2 == 1 {
				formulas = formulasB
			}
			for _, formula := range formulas {
				if err := f.SetCellFormula(formula.Sheet, formula.Cell, formula.Formula); err != nil {
					b.Fatalf("SetCellFormula %s!%s: %v", formula.Sheet, formula.Cell, err)
				}
			}
		}
	})

	b.Run("RealPostgresBatchSetFormulas100WithMirror", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		formulasA := make([]FormulaUpdate, 0, 100)
		formulasB := make([]FormulaUpdate, 0, 100)
		for i := 1; i <= 100; i++ {
			if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), strconv.Itoa(i)); err != nil {
				b.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
			}
			cell := "B" + strconv.Itoa(i)
			formulasA = append(formulasA, FormulaUpdate{Sheet: "Sheet1", Cell: cell, Formula: "A" + strconv.Itoa(i) + "*2"})
			formulasB = append(formulasB, FormulaUpdate{Sheet: "Sheet1", Cell: cell, Formula: "A" + strconv.Itoa(i) + "*3"})
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			formulas := formulasA
			if i%2 == 1 {
				formulas = formulasB
			}
			if err := f.BatchSetFormulas(formulas); err != nil {
				b.Fatalf("BatchSetFormulas(): %v", err)
			}
		}
	})

	b.Run("RealPostgresFirstBatchLookupRecalculate12Cold", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				cancel()
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresEnableMirrorAndFirstBatchLookupRecalculate12", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresEnableMirrorReturnOnly", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			b.StopTimer()
		}
	})

	b.Run("RealPostgresEnableMirrorAutoWarmupReturnOnly", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			warmOpts := opts
			warmOpts.AutoWarmup = true
			warmOpts.WarmupSheets = []string{"Sheet1"}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.EnablePostgresMirror(db, false, warmOpts); err != nil {
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			b.StopTimer()
		}
	})

	b.Run("RealPostgresEnableMirrorAsyncWarmupReturnOnly", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			warmOpts := opts
			warmOpts.AutoWarmup = true
			warmOpts.AsyncWarmup = true
			warmOpts.WarmupSheets = []string{"Sheet1"}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.EnablePostgresMirror(db, false, warmOpts); err != nil {
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			b.StopTimer()
			waitCtx, waitCancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.WaitForPostgresWarmup(waitCtx); err != nil {
				waitCancel()
				b.Fatalf("WaitForPostgresWarmup(): %v", err)
			}
			waitCancel()
		}
	})

	b.Run("RealPostgresEnableMirrorAutoWarmupAndFirstBatchLookupRecalculate12", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			warmOpts := opts
			warmOpts.AutoWarmup = true
			warmOpts.WarmupSheets = []string{"Sheet1"}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.EnablePostgresMirror(db, false, warmOpts); err != nil {
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresFirstBatchLookupRecalculate12AfterPreload", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			if _, err := f.NewSheet("Data"); err != nil {
				b.Fatalf("NewSheet(): %v", err)
			}
			for j := 1; j <= 12; j++ {
				row := j + 1
				sku := "SKU" + strconv.Itoa(j)
				value := strconv.Itoa(j * 100)
				if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
					b.Fatalf("SetCellValue Data!A%d: %v", row, err)
				}
				if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
					b.Fatalf("SetCellValue Data!B%d: %v", row, err)
				}
				if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(j), sku); err != nil {
					b.Fatalf("SetCellValue Sheet1!A%d: %v", j, err)
				}
				if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(j), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(j)+`,Data!A:A,0)),"")`); err != nil {
					b.Fatalf("SetCellFormula Sheet1!B%d: %v", j, err)
				}
			}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
				cancel()
				b.Fatalf("SyncWorkbookToPostgres(): %v", err)
			}
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				cancel()
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			if err := f.PreloadPostgresLookupCache(ctx); err != nil {
				cancel()
				b.Fatalf("PreloadPostgresLookupCache(): %v", err)
			}
			cancel()
			b.StartTimer()
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresBatchLookupRecalculate12", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 12; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			value := strconv.Itoa(i * 100)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
			if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
				b.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
			}
			if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(i)+`,Data!A:A,0)),"")`); err != nil {
				b.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
			}
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}
		if err := f.RecalculateAllWithDependency(); err != nil {
			b.Fatalf("warmup RecalculateAllWithDependency(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresBatchLookupSubExprRecalculate12", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 12; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			value := strconv.Itoa(i * 100)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
			if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
				b.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
			}
			formula := `IF(IFERROR(INDEX(Data!B:B,MATCH(A` + strconv.Itoa(i) + `,Data!A:A,0)),"")="` + value + `","yes","no")`
			if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), formula); err != nil {
				b.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
			}
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}
		if err := f.RecalculateAllWithDependency(); err != nil {
			b.Fatalf("warmup RecalculateAllWithDependency(): %v", err)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresBatchLookupRecalculate12WithInputChange", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 12; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			value := strconv.Itoa(i * 100)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
			if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
				b.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
			}
			if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), `IFERROR(INDEX(Data!B:B,MATCH(A`+strconv.Itoa(i)+`,Data!A:A,0)),"")`); err != nil {
				b.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
			}
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}
		if err := f.RecalculateAllWithDependency(); err != nil {
			b.Fatalf("warmup RecalculateAllWithDependency(): %v", err)
		}

		lookupSKUs := []string{"SKU1", "SKU2", "SKU3", "SKU4"}
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := f.SetCellValue("Sheet1", "A1", lookupSKUs[i%len(lookupSKUs)]); err != nil {
				b.Fatalf("SetCellValue Sheet1!A1: %v", err)
			}
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresBatchLookupSubExprRecalculate12WithInputChange", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 12; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			value := strconv.Itoa(i * 100)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
			if err := f.SetCellValue("Sheet1", "A"+strconv.Itoa(i), sku); err != nil {
				b.Fatalf("SetCellValue Sheet1!A%d: %v", i, err)
			}
			formula := `IF(IFERROR(INDEX(Data!B:B,MATCH(A` + strconv.Itoa(i) + `,Data!A:A,0)),"")="` + value + `","yes","no")`
			if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), formula); err != nil {
				b.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
			}
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}
		if err := f.RecalculateAllWithDependency(); err != nil {
			b.Fatalf("warmup RecalculateAllWithDependency(): %v", err)
		}

		lookupSKUs := []string{"SKU1", "SKU2", "SKU3", "SKU4"}
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := f.SetCellValue("Sheet1", "A1", lookupSKUs[i%len(lookupSKUs)]); err != nil {
				b.Fatalf("SetCellValue Sheet1!A1: %v", err)
			}
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresBatchMirrorFlushRecalculate12SharedInputChange", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)

		f := NewFile()
		if _, err := f.NewSheet("Data"); err != nil {
			b.Fatalf("NewSheet(): %v", err)
		}
		for i := 1; i <= 12; i++ {
			row := i + 1
			sku := "SKU" + strconv.Itoa(i)
			value := strconv.Itoa(i * 100)
			if err := f.SetCellValue("Data", "A"+strconv.Itoa(row), sku); err != nil {
				b.Fatalf("SetCellValue Data!A%d: %v", row, err)
			}
			if err := f.SetCellValue("Data", "B"+strconv.Itoa(row), value); err != nil {
				b.Fatalf("SetCellValue Data!B%d: %v", row, err)
			}
			if err := f.SetCellFormula("Sheet1", "B"+strconv.Itoa(i), `IFERROR(INDEX(Data!B:B,MATCH($A$1,Data!A:A,0)),"")`); err != nil {
				b.Fatalf("SetCellFormula Sheet1!B%d: %v", i, err)
			}
		}
		if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
			b.Fatalf("SetCellValue Sheet1!A1: %v", err)
		}

		ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
		defer cancel()
		if err := f.SyncWorkbookToPostgres(ctx, db, opts); err != nil {
			b.Fatalf("SyncWorkbookToPostgres(): %v", err)
		}
		if err := f.EnablePostgresMirror(db, false, opts); err != nil {
			b.Fatalf("EnablePostgresMirror(): %v", err)
		}
		if err := f.RecalculateAllWithDependency(); err != nil {
			b.Fatalf("warmup RecalculateAllWithDependency(): %v", err)
		}

		lookupSKUs := []string{"SKU1", "SKU2", "SKU3", "SKU4"}
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if err := f.SetCellValue("Sheet1", "A1", lookupSKUs[i%len(lookupSKUs)]); err != nil {
				b.Fatalf("SetCellValue Sheet1!A1: %v", err)
			}
			if err := f.RecalculateAllWithDependency(); err != nil {
				b.Fatalf("RecalculateAllWithDependency(): %v", err)
			}
		}
	})

	b.Run("RealPostgresReopenFirstBatchLookup12NoSnapshot", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)
		noSnapshotBytes, _, formulas := buildPostgresBatchLookupBenchmarkFixture(b, db, opts, 12)

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f, err := OpenReader(bytes.NewReader(noSnapshotBytes))
			if err != nil {
				b.Fatalf("OpenReader(): %v", err)
			}
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				_ = f.Close()
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			results := f.batchCalculatePostgresLookupsWithCache(formulas, nil)
			if got := results["Sheet1!B12"]; got != "1200" {
				_ = f.Close()
				b.Fatalf("unexpected batch result: got %q want %q", got, "1200")
			}
			_ = f.Close()
		}
	})

	b.Run("RealPostgresReopenFirstBatchLookup12EmbeddedSnapshot", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)
		_, withSnapshotBytes, formulas := buildPostgresBatchLookupBenchmarkFixture(b, db, opts, 12)

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f, err := OpenReader(bytes.NewReader(withSnapshotBytes))
			if err != nil {
				b.Fatalf("OpenReader(): %v", err)
			}
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				_ = f.Close()
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}
			results := f.batchCalculatePostgresLookupsWithCache(formulas, nil)
			if got := results["Sheet1!B12"]; got != "1200" {
				_ = f.Close()
				b.Fatalf("unexpected batch result: got %q want %q", got, "1200")
			}
			_ = f.Close()
		}
	})

	b.Run("RealPostgresReopenEmbeddedSnapshotRefreshThenBatchLookup12", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)
		_, withSnapshotBytes, formulas := buildPostgresBatchLookupBenchmarkFixture(b, db, opts, 12)
		tables := getPGMirrorTables(opts.TablePrefix)
		targetCellRef := "B13"

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f, err := OpenReader(bytes.NewReader(withSnapshotBytes))
			if err != nil {
				b.Fatalf("OpenReader(): %v", err)
			}
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				_ = f.Close()
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}

			value := strconv.Itoa(1200 + (i % 2))
			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			b.StopTimer()
			if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.cell)+`
SET value = $1, value_type = $2, updated_at = NOW()
WHERE workbook_id = $3 AND sheet_name = $4 AND cell_ref = $5
`, value, pgMirrorValueTypeNumber, opts.WorkbookID, "Data", targetCellRef); err != nil {
				cancel()
				_ = f.Close()
				b.Fatalf("update cell mirror: %v", err)
			}
			if _, err := db.ExecContext(ctx, `
UPDATE `+qualifyPGTable(opts.Schema, tables.lookup)+`
SET cell_value = $1, cell_value_type = $2, updated_at = NOW()
WHERE workbook_id = $3 AND sheet_name = $4 AND cell_ref = $5
`, value, pgMirrorValueTypeNumber, opts.WorkbookID, "Data", targetCellRef); err != nil {
				cancel()
				_ = f.Close()
				b.Fatalf("update lookup mirror: %v", err)
			}
			b.StartTimer()
			if _, err := f.RefreshCellFromPostgresAndRecalculate(ctx, db, "Data", targetCellRef, opts); err != nil {
				cancel()
				_ = f.Close()
				b.Fatalf("RefreshCellFromPostgresAndRecalculate(): %v", err)
			}
			cancel()

			results := f.batchCalculatePostgresLookupsWithCache(formulas, nil)
			if got := results["Sheet1!B12"]; got != value {
				_ = f.Close()
				b.Fatalf("unexpected refreshed batch result: got %q want %q", got, value)
			}
			_ = f.Close()
		}
	})

	b.Run("RealPostgresReopenEmbeddedSnapshotNoOpRefreshThenBatchLookup12", func(b *testing.B) {
		silenceBenchmarkLogs(b)
		db, opts := openPostgresIntegrationDB(b)
		_, withSnapshotBytes, formulas := buildPostgresBatchLookupBenchmarkFixture(b, db, opts, 12)
		targetCellRef := "B13"

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f, err := OpenReader(bytes.NewReader(withSnapshotBytes))
			if err != nil {
				b.Fatalf("OpenReader(): %v", err)
			}
			if err := f.EnablePostgresMirror(db, false, opts); err != nil {
				_ = f.Close()
				b.Fatalf("EnablePostgresMirror(): %v", err)
			}

			ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
			if _, err := f.RefreshCellFromPostgresAndRecalculate(ctx, db, "Data", targetCellRef, opts); err != nil {
				cancel()
				_ = f.Close()
				b.Fatalf("RefreshCellFromPostgresAndRecalculate(): %v", err)
			}
			cancel()

			results := f.batchCalculatePostgresLookupsWithCache(formulas, nil)
			if got := results["Sheet1!B12"]; got != "1200" {
				_ = f.Close()
				b.Fatalf("unexpected refreshed batch result: got %q want %q", got, "1200")
			}
			_ = f.Close()
		}
	})
}
