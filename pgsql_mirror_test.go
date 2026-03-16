package excelize

import (
	"context"
	"database/sql"
	"database/sql/driver"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
)

type pgMirrorExecCall struct {
	query string
	args  []interface{}
}

type pgMirrorQueryCall struct {
	query string
	args  []interface{}
}

type pgMirrorStubState struct {
	mu           sync.Mutex
	execs        []pgMirrorExecCall
	queries      []pgMirrorQueryCall
	failContains string
	lookupRows   []pgMirrorLookupSeed
	cellRows     []pgMirrorCellSeed
}

type pgMirrorLookupSeed struct {
	workbookID string
	sheet      string
	row        int
	col        int
	value      string
	valueType  string
}

type pgMirrorCellSeed struct {
	workbookID string
	sheet      string
	row        int
	col        int
	value      string
	valueType  string
}

func (s *pgMirrorStubState) recordExec(query string, args []driver.NamedValue) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	call := pgMirrorExecCall{
		query: query,
		args:  make([]interface{}, 0, len(args)),
	}
	for _, arg := range args {
		call.args = append(call.args, arg.Value)
	}
	s.execs = append(s.execs, call)
	if s.failContains != "" && strings.Contains(query, s.failContains) {
		return errors.New("forced exec failure")
	}
	return nil
}

func (s *pgMirrorStubState) reset() {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.execs = nil
	s.queries = nil
}

func (s *pgMirrorStubState) seedLookupRow(workbookID, sheet string, row, col int, value string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.lookupRows = append(s.lookupRows, pgMirrorLookupSeed{
		workbookID: workbookID,
		sheet:      sheet,
		row:        row,
		col:        col,
		value:      value,
		valueType:  pgMirrorValueTypeFromCell(value, CellTypeUnset),
	})
}

func (s *pgMirrorStubState) seedCellRow(workbookID, sheet string, row, col int, value string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.cellRows = append(s.cellRows, pgMirrorCellSeed{
		workbookID: workbookID,
		sheet:      sheet,
		row:        row,
		col:        col,
		value:      value,
		valueType:  pgMirrorValueTypeFromCell(value, CellTypeUnset),
	})
}

func (s *pgMirrorStubState) hasExec(querySubstr string, wantArgs ...interface{}) bool {
	s.mu.Lock()
	defer s.mu.Unlock()
	for _, exec := range s.execs {
		if !strings.Contains(exec.query, querySubstr) {
			continue
		}
		matched := true
		for _, want := range wantArgs {
			found := false
			for _, arg := range exec.args {
				if arg == want {
					found = true
					break
				}
			}
			if !found {
				matched = false
				break
			}
		}
		if matched {
			return true
		}
	}
	return false
}

func (s *pgMirrorStubState) recordQuery(query string) {
	s.recordQueryWithArgs(query, nil)
}

func (s *pgMirrorStubState) recordQueryWithArgs(query string, args []driver.NamedValue) {
	s.mu.Lock()
	defer s.mu.Unlock()
	call := pgMirrorQueryCall{
		query: query,
		args:  make([]interface{}, 0, len(args)),
	}
	for _, arg := range args {
		call.args = append(call.args, arg.Value)
	}
	s.queries = append(s.queries, call)
}

func (s *pgMirrorStubState) countQueries(querySubstr string) int {
	s.mu.Lock()
	defer s.mu.Unlock()

	count := 0
	for _, query := range s.queries {
		if strings.Contains(query.query, querySubstr) {
			count++
		}
	}
	return count
}

func (s *pgMirrorStubState) latestQuery() (pgMirrorQueryCall, bool) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if len(s.queries) == 0 {
		return pgMirrorQueryCall{}, false
	}
	return s.queries[len(s.queries)-1], true
}

func (s *pgMirrorStubState) countExecs(querySubstr string) int {
	s.mu.Lock()
	defer s.mu.Unlock()

	count := 0
	for _, exec := range s.execs {
		if strings.Contains(exec.query, querySubstr) {
			count++
		}
	}
	return count
}

type pgMirrorStubDriver struct {
	state *pgMirrorStubState
}

func (d *pgMirrorStubDriver) Open(string) (driver.Conn, error) {
	return &pgMirrorStubConn{state: d.state}, nil
}

type pgMirrorStubConn struct {
	state *pgMirrorStubState
}

func (c *pgMirrorStubConn) Prepare(string) (driver.Stmt, error) {
	return nil, errors.New("not implemented")
}
func (c *pgMirrorStubConn) Close() error              { return nil }
func (c *pgMirrorStubConn) Begin() (driver.Tx, error) { return &pgMirrorStubTx{}, nil }

func (c *pgMirrorStubConn) BeginTx(context.Context, driver.TxOptions) (driver.Tx, error) {
	return &pgMirrorStubTx{}, nil
}

func (c *pgMirrorStubConn) CheckNamedValue(*driver.NamedValue) error { return nil }

func (c *pgMirrorStubConn) ExecContext(_ context.Context, query string, args []driver.NamedValue) (driver.Result, error) {
	if err := c.state.recordExec(query, args); err != nil {
		return nil, err
	}
	return driver.RowsAffected(1), nil
}

func (c *pgMirrorStubConn) QueryContext(_ context.Context, query string, args []driver.NamedValue) (driver.Rows, error) {
	values := namedValuesToInterfaces(args)
	c.state.recordQueryWithArgs(query, args)

	c.state.mu.Lock()
	defer c.state.mu.Unlock()

	if strings.Contains(query, `"excelize_lookup_mirror"`) && strings.Contains(query, "SELECT row_num") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		col := pgMirrorStubInt(values[2])
		startRow := pgMirrorStubInt(values[3])
		endRow := pgMirrorStubInt(values[4])
		lookupType := values[5].(string)
		lookupValue := values[6].(string)
		for _, row := range c.state.lookupRows {
			if row.workbookID == workbookID && row.sheet == sheet && row.col == col &&
				row.row >= startRow && row.row <= endRow && row.valueType == lookupType && row.value == lookupValue {
				return &pgMirrorStubRows{
					columns: []string{"row_num"},
					rows:    [][]driver.Value{{int64(row.row)}},
				}, nil
			}
		}
		return &pgMirrorStubRows{columns: []string{"row_num"}}, nil
	}

	if strings.Contains(query, `"excelize_lookup_mirror"`) && strings.Contains(query, "SELECT cell_value, MIN(row_num)") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		col := pgMirrorStubInt(values[2])
		startRow := pgMirrorStubInt(values[3])
		endRow := pgMirrorStubInt(values[4])
		lookupType := values[5].(string)
		lookups := pgMirrorStubStrings(values[6:])
		lookupSet := make(map[string]struct{}, len(lookups))
		for _, lookup := range lookups {
			lookupSet[lookup] = struct{}{}
		}
		bestRows := make(map[string]int)
		for _, row := range c.state.lookupRows {
			if row.workbookID != workbookID || row.sheet != sheet || row.col != col || row.row < startRow || row.row > endRow || row.valueType != lookupType {
				continue
			}
			if _, ok := lookupSet[row.value]; !ok {
				continue
			}
			if current, exists := bestRows[row.value]; !exists || row.row < current {
				bestRows[row.value] = row.row
			}
		}
		out := make([][]driver.Value, 0, len(bestRows))
		for lookup, rowNum := range bestRows {
			out = append(out, []driver.Value{lookup, int64(rowNum)})
		}
		return &pgMirrorStubRows{
			columns: []string{"cell_value", "row_num"},
			rows:    out,
		}, nil
	}

	if strings.Contains(query, `"excelize_lookup_mirror"`) && strings.Contains(query, "SELECT col_num") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		rowNum := pgMirrorStubInt(values[2])
		startCol := pgMirrorStubInt(values[3])
		endCol := pgMirrorStubInt(values[4])
		lookupType := values[5].(string)
		lookupValue := values[6].(string)
		for _, row := range c.state.lookupRows {
			if row.workbookID == workbookID && row.sheet == sheet && row.row == rowNum &&
				row.col >= startCol && row.col <= endCol && row.valueType == lookupType && row.value == lookupValue {
				return &pgMirrorStubRows{
					columns: []string{"col_num"},
					rows:    [][]driver.Value{{int64(row.col)}},
				}, nil
			}
		}
		return &pgMirrorStubRows{columns: []string{"col_num"}}, nil
	}

	if strings.Contains(query, `"excelize_lookup_mirror"`) && strings.Contains(query, "SELECT cell_value, MIN(col_num)") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		rowNum := pgMirrorStubInt(values[2])
		startCol := pgMirrorStubInt(values[3])
		endCol := pgMirrorStubInt(values[4])
		lookupType := values[5].(string)
		lookups := pgMirrorStubStrings(values[6:])
		lookupSet := make(map[string]struct{}, len(lookups))
		for _, lookup := range lookups {
			lookupSet[lookup] = struct{}{}
		}
		bestCols := make(map[string]int)
		for _, row := range c.state.lookupRows {
			if row.workbookID != workbookID || row.sheet != sheet || row.row != rowNum || row.col < startCol || row.col > endCol || row.valueType != lookupType {
				continue
			}
			if _, ok := lookupSet[row.value]; !ok {
				continue
			}
			if current, exists := bestCols[row.value]; !exists || row.col < current {
				bestCols[row.value] = row.col
			}
		}
		out := make([][]driver.Value, 0, len(bestCols))
		for lookup, colNum := range bestCols {
			out = append(out, []driver.Value{lookup, int64(colNum)})
		}
		return &pgMirrorStubRows{
			columns: []string{"cell_value", "col_num"},
			rows:    out,
		}, nil
	}

	if strings.Contains(query, `"excelize_cell_mirror"`) && strings.Contains(query, "SELECT value") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		rowNum := pgMirrorStubInt(values[2])
		colNum := pgMirrorStubInt(values[3])
		for _, row := range c.state.cellRows {
			if row.workbookID == workbookID && row.sheet == sheet && row.row == rowNum && row.col == colNum {
				return &pgMirrorStubRows{
					columns: []string{"value", "value_type"},
					rows:    [][]driver.Value{{row.value, row.valueType}},
				}, nil
			}
		}
		return &pgMirrorStubRows{columns: []string{"value", "value_type"}}, nil
	}

	if strings.Contains(query, `"excelize_cell_mirror"`) && strings.Contains(query, "SELECT row_num, value") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		colNum := pgMirrorStubInt(values[2])
		rowNums := pgMirrorStubInts(values[3:])
		rowSet := make(map[int]struct{}, len(rowNums))
		for _, rowNum := range rowNums {
			rowSet[rowNum] = struct{}{}
		}
		out := make([][]driver.Value, 0, len(rowNums))
		for _, row := range c.state.cellRows {
			if row.workbookID == workbookID && row.sheet == sheet && row.col == colNum {
				if _, ok := rowSet[row.row]; ok {
					out = append(out, []driver.Value{int64(row.row), row.value, row.valueType})
				}
			}
		}
		return &pgMirrorStubRows{
			columns: []string{"row_num", "value", "value_type"},
			rows:    out,
		}, nil
	}

	if strings.Contains(query, `"excelize_cell_mirror"`) && strings.Contains(query, "SELECT col_num, value") {
		workbookID := values[0].(string)
		sheet := values[1].(string)
		rowNum := pgMirrorStubInt(values[2])
		colNums := pgMirrorStubInts(values[3:])
		colSet := make(map[int]struct{}, len(colNums))
		for _, colNum := range colNums {
			colSet[colNum] = struct{}{}
		}
		out := make([][]driver.Value, 0, len(colNums))
		for _, row := range c.state.cellRows {
			if row.workbookID == workbookID && row.sheet == sheet && row.row == rowNum {
				if _, ok := colSet[row.col]; ok {
					out = append(out, []driver.Value{int64(row.col), row.value, row.valueType})
				}
			}
		}
		return &pgMirrorStubRows{
			columns: []string{"col_num", "value", "value_type"},
			rows:    out,
		}, nil
	}

	return nil, fmt.Errorf("unsupported query: %s", query)
}

type pgMirrorStubTx struct{}

func (tx *pgMirrorStubTx) Commit() error   { return nil }
func (tx *pgMirrorStubTx) Rollback() error { return nil }

type pgMirrorStubRows struct {
	columns []string
	rows    [][]driver.Value
	idx     int
}

func (r *pgMirrorStubRows) Columns() []string {
	return r.columns
}

func (r *pgMirrorStubRows) Close() error {
	return nil
}

func (r *pgMirrorStubRows) Next(dest []driver.Value) error {
	if r.idx >= len(r.rows) {
		return io.EOF
	}
	copy(dest, r.rows[r.idx])
	r.idx++
	return nil
}

func namedValuesToInterfaces(args []driver.NamedValue) []interface{} {
	values := make([]interface{}, 0, len(args))
	for _, arg := range args {
		values = append(values, arg.Value)
	}
	return values
}

func pgMirrorStubInt(v interface{}) int {
	switch value := v.(type) {
	case int:
		return value
	case int8:
		return int(value)
	case int16:
		return int(value)
	case int32:
		return int(value)
	case int64:
		return int(value)
	default:
		panic(fmt.Sprintf("unexpected integer type %T", v))
	}
}

func pgMirrorStubInts(values []interface{}) []int {
	result := make([]int, 0, len(values))
	for _, value := range values {
		result = append(result, pgMirrorStubInt(value))
	}
	return result
}

func pgMirrorStubStrings(values []interface{}) []string {
	result := make([]string, 0, len(values))
	for _, value := range values {
		result = append(result, value.(string))
	}
	return result
}

var pgMirrorDriverSeq atomic.Int64

func newPGMirrorStubDB(t *testing.T) (*sql.DB, *pgMirrorStubState) {
	t.Helper()
	db, state, err := openPGMirrorStubDB()
	if err != nil {
		t.Fatalf("sql.Open() error = %v", err)
	}
	t.Cleanup(func() {
		_ = db.Close()
	})
	return db, state
}

func openPGMirrorStubDB() (*sql.DB, *pgMirrorStubState, error) {
	state := &pgMirrorStubState{}
	driverName := fmt.Sprintf("pgmirrorstub_%d", pgMirrorDriverSeq.Add(1))
	sql.Register(driverName, &pgMirrorStubDriver{state: state})
	db, err := sql.Open(driverName, "")
	return db, state, err
}

func TestPGMirrorGetSyncOptions(t *testing.T) {
	f := NewFile()
	f.Path = "/tmp/sales-report-2026.xlsx"

	cfg, err := f.getPGSyncOptions()
	if err != nil {
		t.Fatalf("getPGSyncOptions() error = %v", err)
	}
	if cfg.Schema != "public" {
		t.Fatalf("unexpected schema: got %q", cfg.Schema)
	}
	if cfg.TablePrefix != "excelize" {
		t.Fatalf("unexpected table prefix: got %q", cfg.TablePrefix)
	}
	if cfg.BatchSize != 1000 {
		t.Fatalf("unexpected batch size: got %d", cfg.BatchSize)
	}
	if cfg.WorkbookID != "sales-report-2026" {
		t.Fatalf("unexpected workbook id: got %q", cfg.WorkbookID)
	}
}

func TestPGMirrorGetSyncOptionsInvalidIdent(t *testing.T) {
	f := NewFile()

	if _, err := f.getPGSyncOptions(PGSyncOptions{Schema: "bad-name"}); err == nil {
		t.Fatal("expected schema validation error, got nil")
	}
	if _, err := f.getPGSyncOptions(PGSyncOptions{TablePrefix: "bad-prefix"}); err == nil {
		t.Fatal("expected table prefix validation error, got nil")
	}
}

func TestCollectSheetMirrorCells(t *testing.T) {
	f := NewFile()
	if err := f.SetCellValue("Sheet1", "A1", 123); err != nil {
		t.Fatalf("SetCellValue A1 error: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1*2"); err != nil {
		t.Fatalf("SetCellFormula B1 error: %v", err)
	}
	if err := f.SetCellValue("Sheet1", "C3", "tail"); err != nil {
		t.Fatalf("SetCellValue C3 error: %v", err)
	}

	cells, rows, cols, err := f.collectSheetMirrorCells("Sheet1", true)
	if err != nil {
		t.Fatalf("collectSheetMirrorCells() error = %v", err)
	}
	if rows != 3 {
		t.Fatalf("unexpected row count: got %d, want 3", rows)
	}
	if cols != 3 {
		t.Fatalf("unexpected col count: got %d, want 3", cols)
	}

	var (
		foundA1 bool
		foundB1 bool
		foundC3 bool
	)
	for _, c := range cells {
		switch c.Cell {
		case "A1":
			foundA1 = c.Value == "123"
		case "B1":
			foundB1 = c.Formula == "A1*2" || c.Formula == "=A1*2"
		case "C3":
			foundC3 = c.Value == "tail"
		}
	}
	if !foundA1 {
		t.Fatal("A1 not mirrored with expected value")
	}
	if !foundB1 {
		t.Fatal("B1 not mirrored with expected formula")
	}
	if !foundC3 {
		t.Fatal("C3 not mirrored with expected value")
	}
}

func TestBuildPGCellUpsertSQL(t *testing.T) {
	tables := getPGMirrorTables("excelize")
	sqlStr, args := buildPGCellUpsertSQL("public", tables, "book1", "Sheet1", []pgMirrorCell{
		{Cell: "A1", Row: 1, Col: 1, Value: "1", Formula: ""},
		{Cell: "B2", Row: 2, Col: 2, Value: "2", Formula: "A1+1"},
	})

	if !strings.Contains(sqlStr, `INSERT INTO "public"."excelize_cell_mirror"`) {
		t.Fatalf("unexpected insert SQL: %s", sqlStr)
	}
	if !strings.Contains(sqlStr, "FROM UNNEST(") {
		t.Fatalf("expected UNNEST bulk SQL: %s", sqlStr)
	}
	if !strings.Contains(sqlStr, "ON CONFLICT (workbook_id, sheet_name, cell_ref)") {
		t.Fatalf("missing upsert clause: %s", sqlStr)
	}
	if len(args) != 8 {
		t.Fatalf("unexpected args len: got %d, want 8", len(args))
	}
	if args[0] != "book1" || args[1] != "Sheet1" {
		t.Fatalf("unexpected first args: %#v", args[:2])
	}
	cellRefs, ok := args[2].([]string)
	if !ok || len(cellRefs) != 2 || cellRefs[0] != "A1" || cellRefs[1] != "B2" {
		t.Fatalf("unexpected cell refs arg: %#v", args[2])
	}
}

func TestBuildPGLookupUpsertSQL(t *testing.T) {
	tables := getPGMirrorTables("excelize")
	sqlStr, args := buildPGLookupUpsertSQL("public", tables, "book1", "Sheet1", []pgMirrorCell{
		{Cell: "A1", Row: 1, Col: 1, Value: "SKU1"},
		{Cell: "B2", Row: 2, Col: 2, Value: "200"},
	})

	if !strings.Contains(sqlStr, `INSERT INTO "public"."excelize_lookup_mirror"`) {
		t.Fatalf("unexpected insert SQL: %s", sqlStr)
	}
	if !strings.Contains(sqlStr, "FROM UNNEST(") {
		t.Fatalf("expected UNNEST bulk SQL: %s", sqlStr)
	}
	if !strings.Contains(sqlStr, "ON CONFLICT (workbook_id, sheet_name, cell_ref)") {
		t.Fatalf("missing upsert clause: %s", sqlStr)
	}
	if len(args) != 7 {
		t.Fatalf("unexpected args len: got %d, want 7", len(args))
	}
	if args[0] != "book1" || args[1] != "Sheet1" {
		t.Fatalf("unexpected first args: %#v", args[:2])
	}
	cellRefs, ok := args[2].([]string)
	if !ok || len(cellRefs) != 2 || cellRefs[0] != "A1" || cellRefs[1] != "B2" {
		t.Fatalf("unexpected cell refs arg: %#v", args[2])
	}
}

func TestParsePGLookupFormulas(t *testing.T) {
	if _, ok := parsePGVLookupFormula("Sheet1", `VLOOKUP(A1,Data!A:B,2,FALSE)`); !ok {
		t.Fatal("expected VLOOKUP parser to match exact formula")
	}
	if _, ok := parsePGMatchLookupFormula("Sheet1", `MATCH(A1,Data!A2:A4,0)`); !ok {
		t.Fatal("expected MATCH parser to match exact formula")
	}
	if _, ok := parsePGIndexMatchFormula("Sheet1", `INDEX(Data!B:B,MATCH(A1,Data!A:A,0))`); !ok {
		t.Fatal("expected INDEX-MATCH parser to match exact formula")
	}
}

func TestBuildPGDeleteRemovedSheetsSQL(t *testing.T) {
	tables := getPGMirrorTables("excelize")

	sqlStr, args := buildPGDeleteRemovedSheetsSQL("public", tables, "book1", []string{"Sheet1", "Sheet2"})
	if !strings.Contains(sqlStr, `DELETE FROM "public"."excelize_sheet_mirror" WHERE workbook_id = $1 AND sheet_name NOT IN ($2,$3)`) {
		t.Fatalf("unexpected delete SQL: %s", sqlStr)
	}
	if len(args) != 3 || args[0] != "book1" || args[1] != "Sheet1" || args[2] != "Sheet2" {
		t.Fatalf("unexpected args: %#v", args)
	}

	sqlStr, args = buildPGDeleteRemovedSheetsSQL("public", tables, "book1", nil)
	if !strings.Contains(sqlStr, `DELETE FROM "public"."excelize_sheet_mirror" WHERE workbook_id = $1`) {
		t.Fatalf("unexpected delete SQL without sheets: %s", sqlStr)
	}
	if len(args) != 1 || args[0] != "book1" {
		t.Fatalf("unexpected args without sheets: %#v", args)
	}
}

func TestEnableDisablePostgresMirror(t *testing.T) {
	db, _ := newPGMirrorStubDB(t)
	f := NewFile()

	if err := f.EnablePostgresMirror(db, true, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if !f.pgMirrorEnabled || !f.pgMirrorBestEffort || f.pgMirrorDB != db {
		t.Fatal("postgres mirror runtime config not enabled as expected")
	}
	if f.pgMirrorOpts == nil || f.pgMirrorOpts.WorkbookID != "book1" {
		t.Fatalf("unexpected postgres mirror options: %#v", f.pgMirrorOpts)
	}

	f.DisablePostgresMirror()
	if f.pgMirrorEnabled || f.pgMirrorBestEffort || f.pgMirrorDB != nil || f.pgMirrorOpts != nil {
		t.Fatal("postgres mirror runtime config not cleared by DisablePostgresMirror")
	}
}

func TestSetCellValueAutoSyncToPostgres(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()

	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", 42); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if !state.hasExec(`INSERT INTO "public"."excelize_cell_mirror"`, "book1", "Sheet1", "A1") {
		t.Fatal("expected cell mirror upsert for A1")
	}
	if !state.hasExec(`INSERT INTO "public"."excelize_lookup_mirror"`, "book1", "Sheet1", "A1") {
		t.Fatal("expected lookup mirror upsert for A1")
	}

	state.reset()
	f.DisablePostgresMirror()
	if err := f.SetCellValue("Sheet1", "A2", 7); err != nil {
		t.Fatalf("SetCellValue() with disabled mirror error = %v", err)
	}
	if state.hasExec(`INSERT INTO "public"."excelize_cell_mirror"`) {
		t.Fatal("did not expect cell mirror upsert after DisablePostgresMirror")
	}
}

func TestSetCellValueAutoSyncBestEffort(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.failContains = `INSERT INTO "public"."excelize_cell_mirror"`

	f := NewFile()
	if err := f.EnablePostgresMirror(db, true, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", 42); err != nil {
		t.Fatalf("SetCellValue() best-effort error = %v", err)
	}

	f = NewFile()
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", 42); err == nil {
		t.Fatal("expected strict postgres mirror error, got nil")
	}
}

func TestRecalculateSheetAutoSyncToPostgres(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if err := f.SetCellValue("Sheet1", "A1", 2); err != nil {
		t.Fatalf("SetCellValue A1 error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1*2"); err != nil {
		t.Fatalf("SetCellFormula B1 error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	if err := f.SetCellValue("Sheet1", "A1", 3); err != nil {
		t.Fatalf("SetCellValue A1 second write error = %v", err)
	}
	state.reset()

	if err := f.RecalculateSheetWithDependency("Sheet1"); err != nil {
		t.Fatalf("RecalculateSheetWithDependency() error = %v", err)
	}
	if !state.hasExec(`INSERT INTO "public"."excelize_cell_mirror"`, "book1", "Sheet1", "B1") {
		t.Fatal("expected recalculated formula cell B1 to sync to postgres")
	}
	if !state.hasExec(`INSERT INTO "public"."excelize_lookup_mirror"`, "book1", "Sheet1", "B1") {
		t.Fatal("expected recalculated formula cell B1 lookup row to sync to postgres")
	}
}

func TestSyncWorkbookToPostgresDeletesRemovedSheets(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if _, err := f.NewSheet("Sheet2"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("initial SyncWorkbookToPostgres() error = %v", err)
	}

	state.reset()
	if err := f.DeleteSheet("Sheet2"); err != nil {
		t.Fatalf("DeleteSheet() error = %v", err)
	}
	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("second SyncWorkbookToPostgres() error = %v", err)
	}
	if !state.hasExec(`DELETE FROM "public"."excelize_sheet_mirror" WHERE workbook_id = $1 AND sheet_name NOT IN ($2)`, "book1", "Sheet1") {
		t.Fatal("expected stale mirrored sheet cleanup for removed Sheet2")
	}
}

func TestCalcCellValuePostgresFastPathVLOOKUP(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedLookupRow("book1", "Data", 3, 1, "SKU2")
	state.seedCellRow("book1", "Data", 2, 2, "100")
	state.seedCellRow("book1", "Data", 3, 2, "200")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU2"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "200" {
		t.Fatalf("unexpected VLOOKUP result: got %q, want %q", got, "200")
	}
}

func TestCalcCellValuePostgresFastPathMatch(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedLookupRow("book1", "Data", 3, 1, "SKU2")
	state.seedLookupRow("book1", "Data", 4, 1, "SKU3")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU3"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `MATCH(A1,Data!A2:A4,0)`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "3" {
		t.Fatalf("unexpected MATCH result: got %q, want %q", got, "3")
	}
}

func TestCalcCellValuePostgresFastPathIndexMatch(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedLookupRow("book1", "Data", 3, 1, "SKU2")
	state.seedCellRow("book1", "Data", 2, 2, "Apple")
	state.seedCellRow("book1", "Data", 3, 2, "Banana")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "Apple" {
		t.Fatalf("unexpected INDEX-MATCH result: got %q, want %q", got, "Apple")
	}
}

func TestCalcCellValuePostgresFastPathFallbackToNative(t *testing.T) {
	db, _ := newPGMirrorStubDB(t)
	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.SetCellValue("Data", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellValue("Data", "B1", "Value1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IF(VLOOKUP(A1,Data!A:B,2,FALSE)="Value1","ok","bad")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "ok" {
		t.Fatalf("unexpected fallback result: got %q, want %q", got, "ok")
	}
}

func TestRecalculateSheetWithDependencyPostgresFastPath(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedCellRow("book1", "Data", 2, 2, "321")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	state.reset()
	if err := f.RecalculateSheetWithDependency("Sheet1"); err != nil {
		t.Fatalf("RecalculateSheetWithDependency() error = %v", err)
	}
	got, err := f.GetCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("GetCellValue() error = %v", err)
	}
	if got != "321" {
		t.Fatalf("unexpected recalculated result: got %q, want %q", got, "321")
	}
}

func TestCalcCellValuePostgresFastPathUsesQueryCache(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedCellRow("book1", "Data", 2, 2, "321")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula B1 error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "C1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula C1 error = %v", err)
	}

	state.reset()

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil || got != "321" {
		t.Fatalf("CalcCellValue B1 got %q err=%v", got, err)
	}
	got, err = f.CalcCellValue("Sheet1", "C1")
	if err != nil || got != "321" {
		t.Fatalf("CalcCellValue C1 got %q err=%v", got, err)
	}

	if got := state.countQueries("SELECT row_num"); got != 1 {
		t.Fatalf("unexpected lookup query count: got %d, want 1", got)
	}
	if got := state.countQueries("SELECT value"); got != 1 {
		t.Fatalf("unexpected cell query count: got %d, want 1", got)
	}
}

func TestRecalculateAllWithDependencyPostgresBatchLookup(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}

	for i := 1; i <= 12; i++ {
		sku := fmt.Sprintf("SKU%02d", i)
		value := fmt.Sprintf("%d", i*10)
		state.seedLookupRow("book1", "Data", i+1, 1, sku)
		state.seedCellRow("book1", "Data", i+1, 2, value)

		cellRow := strconv.Itoa(i)
		if err := f.SetCellValue("Sheet1", "A"+cellRow, sku); err != nil {
			t.Fatalf("SetCellValue Sheet1!A%s error = %v", cellRow, err)
		}
		if err := f.SetCellFormula("Sheet1", "B"+cellRow, `IFERROR(INDEX(Data!B:B,MATCH(A`+cellRow+`,Data!A:A,0)),"")`); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%s error = %v", cellRow, err)
		}
	}

	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	state.reset()
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency() error = %v", err)
	}

	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "10" {
		t.Fatalf("GetCellValue B1 got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B12"); err != nil || got != "120" {
		t.Fatalf("GetCellValue B12 got %q err=%v", got, err)
	}
	if got := state.countQueries("SELECT cell_value, MIN(row_num)"); got != 1 {
		t.Fatalf("unexpected batch lookup query count: got %d, want 1", got)
	}
	if got := state.countQueries("SELECT row_num, value"); got != 1 {
		t.Fatalf("unexpected batch cell query count: got %d, want 1", got)
	}
}

func TestRecalculateAllWithDependencyPostgresBatchLookupSubExpr(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}

	for i := 1; i <= 12; i++ {
		sku := fmt.Sprintf("SKU%02d", i)
		value := fmt.Sprintf("%d", i*10)
		state.seedLookupRow("book1", "Data", i+1, 1, sku)
		state.seedCellRow("book1", "Data", i+1, 2, value)

		cellRow := strconv.Itoa(i)
		if err := f.SetCellValue("Sheet1", "A"+cellRow, sku); err != nil {
			t.Fatalf("SetCellValue Sheet1!A%s error = %v", cellRow, err)
		}
		formula := fmt.Sprintf(`IF(IFERROR(INDEX(Data!B:B,MATCH(A%s,Data!A:A,0)),"")="%s","yes","no")`, cellRow, value)
		if err := f.SetCellFormula("Sheet1", "B"+cellRow, formula); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%s error = %v", cellRow, err)
		}
	}

	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	state.reset()
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency() error = %v", err)
	}

	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "yes" {
		t.Fatalf("GetCellValue B1 got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B12"); err != nil || got != "yes" {
		t.Fatalf("GetCellValue B12 got %q err=%v", got, err)
	}
	if got := state.countQueries("SELECT cell_value, MIN(row_num)"); got != 1 {
		t.Fatalf("unexpected batch subexpr lookup query count: got %d, want 1", got)
	}
	if got := state.countQueries("SELECT row_num, value"); got != 1 {
		t.Fatalf("unexpected batch subexpr cell query count: got %d, want 1", got)
	}
}

func TestRecalculateAllWithDependencyPostgresBatchMirrorSync(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedLookupRow("book1", "Data", 3, 1, "SKU2")
	state.seedCellRow("book1", "Data", 2, 2, "100")
	state.seedCellRow("book1", "Data", 3, 2, "200")

	for i := 1; i <= 3; i++ {
		row := strconv.Itoa(i)
		if err := f.SetCellValue("Sheet1", "A"+row, "SKU1"); err != nil {
			t.Fatalf("SetCellValue Sheet1!A%s error = %v", row, err)
		}
		if err := f.SetCellFormula("Sheet1", "B"+row, `IFERROR(INDEX(Data!B:B,MATCH($A$1,Data!A:A,0)),"")`); err != nil {
			t.Fatalf("SetCellFormula Sheet1!B%s error = %v", row, err)
		}
	}

	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("SyncWorkbookToPostgres() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("warmup RecalculateAllWithDependency() error = %v", err)
	}

	if err := f.SetCellValue("Sheet1", "A1", "SKU2"); err != nil {
		t.Fatalf("SetCellValue Sheet1!A1 second error = %v", err)
	}
	state.reset()

	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency() error = %v", err)
	}

	if got, err := f.GetCellValue("Sheet1", "B1"); err != nil || got != "200" {
		t.Fatalf("GetCellValue B1 got %q err=%v", got, err)
	}
	if got, err := f.GetCellValue("Sheet1", "B3"); err != nil || got != "200" {
		t.Fatalf("GetCellValue B3 got %q err=%v", got, err)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_cell_mirror"`); got != 1 {
		t.Fatalf("unexpected cell mirror exec count: got %d, want 1", got)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_lookup_mirror"`); got != 1 {
		t.Fatalf("unexpected lookup mirror exec count: got %d, want 1", got)
	}
}

func TestBatchSetCellValueUsesSinglePostgresMirrorFlush(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "old1"); err != nil {
		t.Fatalf("SetCellValue A1 error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A2", "old2"); err != nil {
		t.Fatalf("SetCellValue A2 error = %v", err)
	}
	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("SyncWorkbookToPostgres() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	state.reset()
	err := f.BatchSetCellValue([]CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: "new1"},
		{Sheet: "Sheet1", Cell: "A2", Value: "new2"},
	})
	if err != nil {
		t.Fatalf("BatchSetCellValue() error = %v", err)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_cell_mirror"`); got != 1 {
		t.Fatalf("unexpected cell mirror exec count: got %d, want 1", got)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_lookup_mirror"`); got != 1 {
		t.Fatalf("unexpected lookup mirror exec count: got %d, want 1", got)
	}
}

func TestSetCellValuesUsesSinglePostgresMirrorFlush(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "old1"); err != nil {
		t.Fatalf("SetCellValue A1 error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A2", "old2"); err != nil {
		t.Fatalf("SetCellValue A2 error = %v", err)
	}
	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("SyncWorkbookToPostgres() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	state.reset()
	if err := f.SetCellValues("Sheet1", map[string]interface{}{
		"A1": "new1",
		"A2": "new2",
	}); err != nil {
		t.Fatalf("SetCellValues() error = %v", err)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_cell_mirror"`); got != 1 {
		t.Fatalf("unexpected cell mirror exec count: got %d, want 1", got)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_lookup_mirror"`); got != 1 {
		t.Fatalf("unexpected lookup mirror exec count: got %d, want 1", got)
	}
}

func TestBatchSetFormulasUsesSinglePostgresMirrorFlush(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	f := NewFile()
	if err := f.SetCellValue("Sheet1", "A1", "1"); err != nil {
		t.Fatalf("SetCellValue A1 error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A2", "2"); err != nil {
		t.Fatalf("SetCellValue A2 error = %v", err)
	}
	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("SyncWorkbookToPostgres() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	state.reset()
	err := f.BatchSetFormulas([]FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "A2*2"},
	})
	if err != nil {
		t.Fatalf("BatchSetFormulas() error = %v", err)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_cell_mirror"`); got != 1 {
		t.Fatalf("unexpected cell mirror exec count: got %d, want 1", got)
	}
	if got := state.countExecs(`INSERT INTO "public"."excelize_lookup_mirror"`); got != 1 {
		t.Fatalf("unexpected lookup mirror exec count: got %d, want 1", got)
	}
}

func TestPreloadPostgresLookupCacheWarmsWholeFormulaCache(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedCellRow("book1", "Data", 2, 2, "321")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	if err := f.PreloadPostgresLookupCache(context.Background()); err != nil {
		t.Fatalf("PreloadPostgresLookupCache() error = %v", err)
	}
	state.reset()

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "321" {
		t.Fatalf("unexpected result: got %q want %q", got, "321")
	}
	if got := state.countQueries("SELECT row_num"); got != 0 {
		t.Fatalf("unexpected lookup query count after preload: got %d, want 0", got)
	}
	if got := state.countQueries("SELECT value"); got != 0 {
		t.Fatalf("unexpected cell query count after preload: got %d, want 0", got)
	}
}

func TestEnablePostgresMirrorAutoWarmupWarmsWholeFormulaCache(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedCellRow("book1", "Data", 2, 2, "321")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{
		WorkbookID:   "book1",
		AutoWarmup:   true,
		WarmupSheets: []string{"Sheet1"},
	}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}
	state.reset()

	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "321" {
		t.Fatalf("unexpected result: got %q want %q", got, "321")
	}
	if got := state.countQueries("SELECT row_num"); got != 0 {
		t.Fatalf("unexpected lookup query count after auto warmup: got %d, want 0", got)
	}
	if got := state.countQueries("SELECT value"); got != 0 {
		t.Fatalf("unexpected cell query count after auto warmup: got %d, want 0", got)
	}
}

func TestEnablePostgresMirrorAsyncWarmupStatusAndWait(t *testing.T) {
	db, state := newPGMirrorStubDB(t)
	state.seedLookupRow("book1", "Data", 2, 1, "SKU1")
	state.seedCellRow("book1", "Data", 2, 2, "321")

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "SKU1"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`); err != nil {
		t.Fatalf("SetCellFormula() error = %v", err)
	}

	started := make(chan struct{})
	release := make(chan struct{})
	pgMirrorWarmupTestHook = func() {
		close(started)
		<-release
	}
	t.Cleanup(func() {
		pgMirrorWarmupTestHook = nil
	})

	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{
		WorkbookID:   "book1",
		AutoWarmup:   true,
		AsyncWarmup:  true,
		WarmupSheets: []string{"Sheet1"},
	}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	<-started
	status := f.GetPostgresWarmupStatus()
	if !status.Enabled || !status.InProgress || !status.Async {
		t.Fatalf("unexpected warmup status while running: %#v", status)
	}
	if len(status.Sheets) != 1 || status.Sheets[0] != "Sheet1" {
		t.Fatalf("unexpected warmup sheets: %#v", status.Sheets)
	}

	waitDone := make(chan error, 1)
	go func() {
		waitDone <- f.WaitForPostgresWarmup(context.Background())
	}()
	select {
	case err := <-waitDone:
		t.Fatalf("WaitForPostgresWarmup() returned before release: %v", err)
	default:
	}

	close(release)
	if err := <-waitDone; err != nil {
		t.Fatalf("WaitForPostgresWarmup() error = %v", err)
	}

	status = f.GetPostgresWarmupStatus()
	if status.InProgress || !status.Async || status.CompletedAt.IsZero() || status.Err != nil {
		t.Fatalf("unexpected warmup status after completion: %#v", status)
	}

	state.reset()
	got, err := f.CalcCellValue("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValue() error = %v", err)
	}
	if got != "321" {
		t.Fatalf("unexpected result: got %q want %q", got, "321")
	}
	if got := state.countQueries("SELECT row_num"); got != 0 {
		t.Fatalf("unexpected lookup query count after async warmup: got %d, want 0", got)
	}
	if got := state.countQueries("SELECT value"); got != 0 {
		t.Fatalf("unexpected cell query count after async warmup: got %d, want 0", got)
	}
}

func TestEnablePostgresMirrorAutoSyncUsesFastPathAfterWorkbookSync(t *testing.T) {
	db, state := newPGMirrorStubDB(t)

	f := NewFile()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet() error = %v", err)
	}
	if err := f.SetCellValue("Sheet1", "A1", "before"); err != nil {
		t.Fatalf("SetCellValue() error = %v", err)
	}
	if err := f.SyncWorkbookToPostgres(context.Background(), db, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("SyncWorkbookToPostgres() error = %v", err)
	}
	if err := f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"}); err != nil {
		t.Fatalf("EnablePostgresMirror() error = %v", err)
	}

	state.reset()
	if err := f.SetCellValue("Sheet1", "A1", "after"); err != nil {
		t.Fatalf("SetCellValue() second error = %v", err)
	}

	if state.hasExec(`"excelize_workbook_mirror"`) {
		t.Fatalf("fast path should skip workbook mirror upsert after initial workbook sync")
	}
	if state.hasExec(`"excelize_sheet_mirror"`) {
		t.Fatalf("fast path should skip sheet mirror upsert for in-bounds cell updates")
	}
	if !state.hasExec(`"excelize_cell_mirror"`) {
		t.Fatalf("expected cell mirror upsert")
	}
	if !state.hasExec(`"excelize_lookup_mirror"`) {
		t.Fatalf("expected lookup mirror upsert")
	}
}

func BenchmarkCalcCellValuePostgresFastPath(b *testing.B) {
	b.Run("NativeIndexMatch", func(b *testing.B) {
		f := NewFile()
		_, _ = f.NewSheet("Data")
		for i := 1; i <= 1000; i++ {
			_ = f.SetCellValue("Data", "A"+strconv.Itoa(i), fmt.Sprintf("SKU%04d", i))
			_ = f.SetCellValue("Data", "B"+strconv.Itoa(i), i)
		}
		_ = f.SetCellValue("Sheet1", "A1", "SKU0900")
		_ = f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`)

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_, _ = f.CalcCellValue("Sheet1", "B1")
			f.calcCache.Clear()
		}
	})

	b.Run("PostgresFastPathIndexMatch", func(b *testing.B) {
		db, state, err := openPGMirrorStubDB()
		if err != nil {
			b.Fatalf("openPGMirrorStubDB() error = %v", err)
		}
		defer db.Close()
		for i := 1; i <= 1000; i++ {
			key := fmt.Sprintf("SKU%04d", i)
			state.seedLookupRow("book1", "Data", i, 1, key)
			state.seedCellRow("book1", "Data", i, 2, strconv.Itoa(i))
		}

		f := NewFile()
		_, _ = f.NewSheet("Data")
		_ = f.EnablePostgresMirror(db, false, PGSyncOptions{WorkbookID: "book1"})
		_ = f.SetCellValue("Sheet1", "A1", "SKU0900")
		_ = f.SetCellFormula("Sheet1", "B1", `IFERROR(INDEX(Data!B:B,MATCH(A1,Data!A:A,0)),"")`)

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_, _ = f.CalcCellValue("Sheet1", "B1")
			f.calcCache.Clear()
		}
	})
}
