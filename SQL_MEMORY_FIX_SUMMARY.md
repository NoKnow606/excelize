# SQL Formula Memory Fix Summary

**Date**: 2026-05-07  
**Branch**: `fix-memory-usage` (merged to `main`)  
**Commit**: `3ffbbad` — `perf(sql): cut redundant SQL re-execution & sheet materialization`  
**Trigger**: `excelize-mcp` was rolled back because commit `724da84` (SQL subquery support) caused a severe RSS spike — `/formula/set` on a real-world 8 MB workbook with a single complex SQL formula was materializing every source worksheet into SQLite **4 times per request**.

---

## Root Cause Analysis

### What a SQL formula costs

Every call to `ExecuteSQL` does three expensive things in sequence:

1. **`GetRows(sheet)`** — loads the entire worksheet into a `[][]string` matrix
2. **SQLite `CREATE TABLE` + `INSERT INTO`** — inserts every row into an in-memory SQLite database
3. **SQL `Prepare` + `Execute`** — runs the query

For the production workbook (`v6-daily-verifing-mem.xlsx`, ~8 MB, 5 source sheets with thousands of rows), this costs roughly **2–3 GB TotalAlloc and 8–14 seconds per execution**.

### How many times was it running?

Before this fix, a single `/formula/set` call triggered **4 SQL executions** and **3 full sheet materializations**:

| Stage | Call | SQL execs | Reason |
|---|---|---|---|
| `evaluateFormula` | `CalcCellValueWithMatrix` | 1 | Evaluate the formula, produce matrix |
| `UpdateSheetFormulaCache` phase 2 | `CalcCellValues` → `CalcCellValue` | 1 | Recalculate all formulas on the sheet |
| `UpdateSheetFormulaCache` phase 2 | `persistFormulaResult` → `persistSQLFormulaResult` → `CalcCellValueWithMatrix` | 1 | Persist spill range — but re-ran SQL to get the matrix |
| `UpdateSheetFormulaCache` phase 3 | `CalcCellValue` (the old SQL-specific branch) | 1 | One more redundant re-run |

**Total: 4 SQL executions, ~12 GB TotalAlloc, ~780 MB peak heap per `/formula/set`.**

For `/formula/compile`, `CompileSQL` was calling `prepareSQL` → full `materializeSheets` → `GetRows` + all INSERTs, costing ~2.5 GB TotalAlloc and 2.5 seconds — even though compile only needed to validate syntax and column names.

---

## Fixes

### Fix 1 — `CalcCellValues`: reuse matrix, skip double SQL (calc.go)

**Problem**: For SQL cells, `CalcCellValues` was calling `CalcCellValue` to get the scalar value, then `persistFormulaResult` → `persistSQLFormulaResult` → `CalcCellValueWithMatrix` to get the matrix for spill persistence. That's two SQL executions for one cell.

**Fix**: For SQL cells, call `CalcCellValueWithMatrix` directly (one SQL execution), then persist via `persistSQLFormulaResultWithMatrix` with the already-computed matrix.

```go
// Before
result, err := f.CalcCellValue(sheet, cell, opts...)
if isSQLFormula {
    f.persistFormulaResult(sheet, cell, result, nil, false, false)
}

// After
if isSQLFormula {
    matrixResult, mErr := f.CalcCellValueWithMatrix(sheet, cell, opts...)
    f.persistSQLFormulaResultWithMatrix(sheet, cell, matrixResult.Value, matrixResult, nil, false, false)
    ...
    continue
}
result, err := f.CalcCellValue(sheet, cell, opts...)
```

### Fix 2 — `CompileSQL`: schema-only validation (sql_formula.go)

**Problem**: `CompileSQL` was calling `prepareSQL` → `materializeSheets` → `materializeSheet` → `GetRows` + all `INSERT INTO` for every source sheet, even though it only needed to check if column names were valid SQL identifiers.

**Fix**: Introduce `prepareSQLSchemaOnly` → `materializeSheetsSchema` → `materializeSheetSchema`. The schema-only variant uses the streaming `Rows()` iterator, reads **only the first row** (headers), creates an empty `CREATE TABLE`, and calls `db.Prepare(query)` on the empty schema. SQLite validates syntax and column references without any data rows.

```go
func materializeSheetSchema(db *sql.DB, f *File, tableName, sheetName string) ([]string, error) {
    rows, _ := f.Rows(sheetName)
    defer rows.Close()
    rows.Next()                     // read header row only
    headerRow, _ := rows.Columns()
    // CREATE TABLE tableName (col1 TEXT, col2 TEXT, ...)
    db.Exec(buildCreateTableSQL(tableName, headers))
    return headers, nil
}
```

**Result**: `/formula/compile` dropped from ~2.5 s / ~2.5 GB → ~231 ms / ~242 MB (~11× faster, ~90% less memory).

### Fix 3 — `UpdateSheetFormulaCache` phase 3: skip SQL formulas (formula_cache.go)

**Problem**: After `CalcCellValues` already persisted the SQL spill range (Fix 1), `UpdateSheetFormulaCache`'s phase 3 loop was hitting every formula cell including SQL ones, calling `persistFormulaResult` which internally called `CalcCellValue` → re-entering `CalcCellValueWithMatrix` for SQL cells — another redundant full materialization.

**Fix**: Detect SQL formulas in the phase 3 loop and `continue` past them — they were already handled in phase 2.

```go
for _, fc := range formulas {
    formula, ferr := f.GetCellFormula(sheet, fc.cell)
    if ferr == nil && IsSQLFormula(formula) {
        continue  // already persisted in phase 2 via persistSQLFormulaResultWithMatrix
    }
    ...
    f.persistFormulaResult(sheet, fc.cell, value, nil, false, false)
}
```

### Fix 4 — New public API: `PersistSQLFormulaResultWithMatrix` (batch_dag_scheduler.go)

**Problem**: `excelize-mcp`'s `saveSQLFormulaResult` was calling `wb.file.UpdateSheetFormulaCache(sheet)` to persist the SQL spill range. This re-walked every formula on the sheet and ran SQL again — even though `evaluateFormula` had already computed the matrix one moment earlier.

**Fix**: Extract the spill-range writing logic into `applySQLFormulaMatrix` (no SQL execution). Wire it through an internal `persistSQLFormulaResultWithMatrix` and expose a public `PersistSQLFormulaResultWithMatrix` so `excelize-mcp` can pass the already-computed matrix directly.

```go
// New public API
func (f *File) PersistSQLFormulaResultWithMatrix(
    sheet, cellName, fallbackValue string,
    result CalcCellValueWithMatrixResult,
) bool {
    return f.persistSQLFormulaResultWithMatrix(sheet, cellName, fallbackValue, result, nil, false, false)
}
```

`excelize-mcp` then uses it in `saveSQLFormulaResult`:

```go
// Before
wb.file.UpdateSheetFormulaCache(wb.worksheetName)

// After
persistSQLFormulaSpillFromMatrix(wb.file, wb.worksheetName, result.cell, result.calculatedValue, result.rangeValues)
// which calls: f.PersistSQLFormulaResultWithMatrix(sheet, cellName, fallbackValue, matrixResult)
```

### Instrumentation — `SQLExecuteCount` / `ResetSQLExecuteCount` (sql_formula.go)

Added a process-global `uint64` atomic counter incremented in `ExecuteSQL`. Used by benchmarks to prove the fix reduces SQL executions to exactly 1 per request.

---

## Benchmarks

Real-world workbook: `v6-daily-verifing-mem.xlsx` (~8 MB, 5 source sheets, 1 SQL formula producing a 744-row × 45-col spill).

### `/formula/set` (end-to-end)

| Stage | Wall time | TotalAlloc | Peak HeapAlloc | SQL executions |
|---|---|---|---|---|
| **Before** (commit `724da84`) | ~14.0 s | ~12.0 GB | ~780 MB | 4 |
| After Fix 1+3 (lib only, `cache` mode) | ~13.9 s | ~9.5 GB | ~720 MB | 2 |
| **After Fix 1+3+4** (full, `matrix` mode) | **~8.6 s** | **~7.7 GB** | **~141 MB** | **1** |

Persistence step alone (the dominant cost after evaluation):

| | Wall time | TotalAlloc |
|---|---|---|
| `UpdateSheetFormulaCache` (old) | 8.88 s | 7.7 GB |
| `PersistSQLFormulaResultWithMatrix` (new) | 22 ms | 15 MB |

**99.7% faster, 99.8% less allocation for the persistence step.**

### `/formula/compile` (`CompileSQL`)

| | Wall time | TotalAlloc |
|---|---|---|
| Full materialization (before) | ~2.5 s | ~2.5 GB |
| Schema-only (after Fix 2) | ~231 ms | ~242 MB |

**~11× faster, ~90% less allocation.**

---

## Correctness Verification

### Round-trip survival test

`test/cmd/sql-roundtrip-verify` simulates the exact scenarios the old `UpdateSheetFormulaCache` comment warned about — "reopening, downloading, or re-uploading the workbook loses the materialized result block":

```
Round 1 (after /formula/set):    bytes=8,287,012  md5(spill)=82375b...338f  rows=744  cols=45  sqlExec=1
Round 2 (reopen + re-save):      bytes=8,287,012  md5(spill)=82375b...338f  rows=744  cols=45
Round 3 (re-upload + re-save):   bytes=8,287,012  md5(spill)=82375b...338f  rows=744  cols=45

✅ PASS — SQL spill range survives round-trip:
   • spill cells (md5)      : identical across all 3 rounds
   • spill dimensions       : 744 rows × 45 cols at anchor A1
   • SQL formula at anchor  : preserved
   • SQL re-executions      : 1 (only the 1 ExecuteSQL in Round 1)
```

### Why the concern was unfounded

The worry was: "does the new path lose the spill range on reopen/download/re-upload?"

Answer: **No.** Both the old `UpdateSheetFormulaCache` path and the new `PersistSQLFormulaResultWithMatrix` path ultimately call the same internal `applySQLFormulaMatrix` function with identical arguments `(sheet, cell, fallback, matrix, nil, false, false)`. The spill-range writing code — setting `c.F.Ref`, clearing old spill cells, writing individual `<v>` for each spill cell, refreshing sheet dimensions — is identical. The only difference is whether the matrix was computed by re-running SQL or passed in precomputed.

### Integration test (`excelize-mcp`)

`TestSQLFormulaRoundTripExportAndUpdateFilePreservesSpillRange` passes end-to-end:

```
HandleSet (new PersistSQLFormulaResultWithMatrix path)
    ↓
HandleDownload  →  assertSQLFormulaWorkbookContent ✅  (A1=Region, A2=South, B2=20, A3=North, B3=10)
    ↓
HandleUpdateFile  →  assertSQLFormulaWorkbookContent ✅  (same values after re-upload)
```

---

## File Changes

### `omnimcp-excelize` (commit `3ffbbad`, branch `fix-memory-usage`)

| File | Change |
|---|---|
| `sql_formula.go` | `SQLExecuteCount`/`ResetSQLExecuteCount` instrumentation; `CompileSQL` → schema-only path; `prepareSQLSchemaOnly` / `prepareSQLInternal(schemaOnly bool)`; `materializeSheetSchema` (header-only streaming) |
| `batch_dag_scheduler.go` | Extract `applySQLFormulaMatrix`; add internal `persistSQLFormulaResultWithMatrix`; expose public `PersistSQLFormulaResultWithMatrix` |
| `calc.go` | `CalcCellValues`: SQL cells take `CalcCellValueWithMatrix` + `persistSQLFormulaResultWithMatrix` path (1 SQL exec per cell, not 2) |
| `formula_cache.go` | `UpdateSheetFormulaCache` phase 3: `continue` past SQL formulas |
| `test/cmd/sql-mem-bench/main.go` | Benchmark tool: `--mode cache` vs `--mode matrix`, `--dump` for TSV spill export, `sqlExec` counting |
| `test/cmd/sql-compile-bench/main.go` | Benchmark tool: full vs schema-only `CompileSQL` comparison |
| `test/cmd/sql-roundtrip-verify/main.go` | 3-round reopen/download/re-upload survival test |

### `excelize-mcp` (working tree, `api/formula_api.go`)

| Change | Detail |
|---|---|
| `saveSQLFormulaResult` | Replace `UpdateSheetFormulaCache` with `persistSQLFormulaSpillFromMatrix` |
| `persistSQLFormulaSpillFromMatrix` (new helper) | Converts `result.rangeValues ([][]interface{})` → `CalcCellValueWithMatrixResult`, calls `f.PersistSQLFormulaResultWithMatrix` |
| `go.mod` | Temporary local replace: `replace github.com/xuri/excelize/v2 => ../omnimcp-excelize` |

---

## Remaining Behaviour Difference

The old `UpdateSheetFormulaCache` path had an undocumented side effect: while walking the sheet to re-persist SQL spill ranges, it also refreshed the cached `<v>` values of **all other (non-SQL) formulas** on the same sheet. The new path only persists the SQL spill range.

This is intentional and consistent with the existing API contract: `/formula/set` for SQL formulas already uses `skipRecalculation = true` and documents "saved without workbook recalculation for SQL formula". Other formulas' cached values are unchanged from their previous state and will be recalculated by Excel on open, or on the next explicit `RecalculateAllWithDependency` call.

If the non-SQL formula `<v>` refresh is needed in a future scenario, it can be added back as a targeted pass after the SQL spill is written — without re-running the SQL query.

---

## Next Steps

1. **Push `fix-memory-usage` branch** to `origin/OmniMCP-AI/excelize` and create a new tagged version (e.g. `v2.0.0-20260507-fix-memory`).
2. **Restore the remote replace** in `excelize-mcp/go.mod`:
   ```
   replace github.com/xuri/excelize/v2 => github.com/OmniMCP-AI/excelize/v2 v2.0.0-20260507-fix-memory
   ```
3. **Commit `excelize-mcp/api/formula_api.go`** on a branch and merge.
4. **Deploy and confirm** RSS on the `excelize-mcp` service is back to baseline on the same production workload.
