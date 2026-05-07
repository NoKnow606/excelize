package main

// sql-mem-bench mimics the excelize-mcp /formula/set flow for a SQL formula:
//
//	1. Open workbook from disk (simulates MongoDB load + parse).
//	2. Run f.ExecuteSQL(...) once (== api.evaluateFormula); keep the matrix.
//	3. f.SetCellFormula(...) (write the formula into target cell).
//	4. Persist the spill range, choosing one of two strategies:
//	     --mode=cache   (default) calls f.UpdateSheetFormulaCache(sheet) — this
//	                    is what current excelize-mcp does and re-runs SQL twice
//	                    inside the lib (CalcCellValues + Phase 3 persist).
//	     --mode=matrix  reuses the matrix returned by step 2 and calls
//	                    f.PersistSQLFormulaResultWithMatrix(...) — eliminates
//	                    the double execution; this is the recommended new path.
//	5. f.WriteToBuffer() (== persist updated workbook).
//	6. Reopen the buffer and dump a few cells from the spill range so the two
//	   modes can be diff-compared for correctness.
//
// Usage:
//
//	go run ./test/cmd/sql-mem-bench [--mode=cache|matrix] [--dump=N] \
//	  <xlsx_path> <sheet> <cell> <sql_file>

import (
	"bytes"
	"flag"
	"fmt"
	"os"
	"runtime"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type memSnapshot struct {
	label      string
	wallNs     int64
	heapAlloc  uint64
	heapInuse  uint64
	heapSys    uint64
	totalAlloc uint64
	sys        uint64
	numGC      uint32
	sqlCount   uint64
}

func snap(label string, started time.Time) memSnapshot {
	runtime.GC()
	var ms runtime.MemStats
	runtime.ReadMemStats(&ms)
	return memSnapshot{
		label:      label,
		wallNs:     time.Since(started).Nanoseconds(),
		heapAlloc:  ms.HeapAlloc,
		heapInuse:  ms.HeapInuse,
		heapSys:    ms.HeapSys,
		totalAlloc: ms.TotalAlloc,
		sys:        ms.Sys,
		numGC:      ms.NumGC,
		sqlCount:   excelize.SQLExecuteCount(),
	}
}

func mb(n uint64) string {
	return fmt.Sprintf("%.1f MB", float64(n)/(1024*1024))
}

func printRow(s memSnapshot, baseSQL uint64) {
	fmt.Printf(
		"  %-32s  wall=%-9s  heapAlloc=%-10s  heapInuse=%-10s  totalAlloc=%-10s  sys=%-10s  GC=%-3d  sqlExec=%d\n",
		s.label,
		time.Duration(s.wallNs).Round(time.Millisecond),
		mb(s.heapAlloc),
		mb(s.heapInuse),
		mb(s.totalAlloc),
		mb(s.sys),
		s.numGC,
		s.sqlCount-baseSQL,
	)
}

func mustReadFile(p string) string {
	b, err := os.ReadFile(p)
	if err != nil {
		fmt.Fprintf(os.Stderr, "read %s: %v\n", p, err)
		os.Exit(1)
	}
	return strings.TrimSpace(string(b))
}

func main() {
	mode := flag.String("mode", "cache", "spill persistence mode: 'cache' (UpdateSheetFormulaCache) or 'matrix' (PersistSQLFormulaResultWithMatrix)")
	dump := flag.Int("dump", 5, "number of spill rows to dump from the saved workbook for verification")
	dumpOut := flag.String("dump-out", "", "if set, write all spill values to this file (one row per line, tab-separated)")
	flag.Parse()
	if flag.NArg() < 4 {
		fmt.Fprintln(os.Stderr, "usage: sql-mem-bench [--mode=cache|matrix] [--dump=N] <xlsx_path> <sheet> <cell> <sql_file>")
		os.Exit(2)
	}
	xlsxPath := flag.Arg(0)
	sheet := flag.Arg(1)
	cell := flag.Arg(2)
	sqlPath := flag.Arg(3)

	formula := mustReadFile(sqlPath)
	formula = strings.TrimPrefix(formula, "=")

	excelize.ResetSQLExecuteCount()
	t0 := time.Now()
	baseline := snap("0_start", t0)
	printRow(baseline, 0)

	f, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	afterOpen := snap("1_after_open", t0)
	printRow(afterOpen, baseline.sqlCount)

	// 1) Mimic api.evaluateFormula: a single ExecuteSQL invocation.
	t1 := time.Now()
	queryResult, err := f.ExecuteSQL(formula)
	if err != nil {
		fmt.Fprintf(os.Stderr, "ExecuteSQL: %v\n", err)
		os.Exit(1)
	}
	executeSnap := snap("2_after_executeSQL", t1)
	printRow(executeSnap, afterOpen.sqlCount)

	// 2) Mimic api.saveSQLFormulaResult: SetCellFormula + persistence step.
	t2 := time.Now()
	if err := f.SetCellFormula(sheet, cell, formula); err != nil {
		fmt.Fprintf(os.Stderr, "SetCellFormula: %v\n", err)
		os.Exit(1)
	}
	setSnap := snap("3_after_setFormula", t2)
	printRow(setSnap, executeSnap.sqlCount)

	t3 := time.Now()
	switch *mode {
	case "cache":
		if err := f.UpdateSheetFormulaCache(sheet); err != nil {
			fmt.Fprintf(os.Stderr, "UpdateSheetFormulaCache: %v\n", err)
			os.Exit(1)
		}
	case "matrix":
		topLeft := ""
		if len(queryResult.Matrix) > 0 && len(queryResult.Matrix[0]) > 0 {
			topLeft = fmt.Sprint(queryResult.Matrix[0][0])
		}
		matrixResult := excelize.CalcCellValueWithMatrixResult{
			Value:  topLeft,
			Matrix: queryResult.Matrix,
		}
		if !f.PersistSQLFormulaResultWithMatrix(sheet, cell, topLeft, matrixResult) {
			fmt.Fprintf(os.Stderr, "PersistSQLFormulaResultWithMatrix returned false (cell not a SQL formula?)\n")
			os.Exit(1)
		}
	default:
		fmt.Fprintf(os.Stderr, "unknown mode %q\n", *mode)
		os.Exit(2)
	}
	persistSnap := snap("4_after_persist", t3)
	printRow(persistSnap, setSnap.sqlCount)

	// 3) Mimic persisting back to storage.
	t4 := time.Now()
	buf, err := f.WriteToBuffer()
	if err != nil {
		fmt.Fprintf(os.Stderr, "WriteToBuffer: %v\n", err)
		os.Exit(1)
	}
	writeSnap := snap("5_after_writeBuffer", t4)
	printRow(writeSnap, persistSnap.sqlCount)

	fmt.Printf("\nSummary: workbook=%s sheet=%s cell=%s mode=%s\n", xlsxPath, sheet, cell, *mode)
	fmt.Printf("  total wall time      = %s\n", time.Since(t0).Round(time.Millisecond))
	fmt.Printf("  total SQL executions = %d  (peak total: %d)\n",
		writeSnap.sqlCount-baseline.sqlCount, writeSnap.sqlCount)
	fmt.Printf("  peak HeapAlloc       = %s   peak HeapInuse = %s   final TotalAlloc = %s\n",
		mb(maxU64(executeSnap.heapAlloc, persistSnap.heapAlloc, writeSnap.heapAlloc)),
		mb(maxU64(executeSnap.heapInuse, persistSnap.heapInuse, writeSnap.heapInuse)),
		mb(writeSnap.totalAlloc),
	)
	fmt.Printf("  output buffer size   = %s\n", mb(uint64(buf.Len())))

	// 4) Verify by reopening the buffer and reading the spill range cells back.
	if err := verifySpill(buf.Bytes(), sheet, cell, *dump, *dumpOut); err != nil {
		fmt.Fprintf(os.Stderr, "verify spill: %v\n", err)
		os.Exit(1)
	}
}

func verifySpill(data []byte, sheet, anchor string, dumpRows int, dumpOut string) error {
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return fmt.Errorf("reopen buffer: %w", err)
	}
	defer f.Close()

	startCol, startRow, err := excelize.CellNameToCoordinates(anchor)
	if err != nil {
		return fmt.Errorf("invalid anchor %q: %w", anchor, err)
	}

	rows, err := f.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("read sheet %q: %w", sheet, err)
	}
	maxCol := 0
	totalRowsAfter := 0
	for r := startRow - 1; r < len(rows); r++ {
		if r < 0 || r >= len(rows) {
			continue
		}
		row := rows[r]
		hasContent := false
		// Determine the right-most populated cell starting at startCol.
		for c := startCol - 1; c < len(row); c++ {
			if c < 0 {
				continue
			}
			if row[c] != "" {
				hasContent = true
				if c+1 > maxCol {
					maxCol = c + 1
				}
			}
		}
		if !hasContent {
			break
		}
		totalRowsAfter++
	}

	fmt.Printf("\nSpill verification (sheet=%q, anchor=%s):\n", sheet, anchor)
	fmt.Printf("  contiguous filled rows starting at anchor = %d\n", totalRowsAfter)
	fmt.Printf("  detected right-most filled column index   = %d\n", maxCol)

	// Print the first dumpRows rows of the spill range.
	limit := dumpRows
	if limit > totalRowsAfter {
		limit = totalRowsAfter
	}
	for i := 0; i < limit; i++ {
		r := startRow - 1 + i
		row := rows[r]
		end := maxCol
		if end > len(row) {
			end = len(row)
		}
		visible := row[startCol-1 : end]
		fmt.Printf("  row[%d]: %s\n", startRow+i, abbreviate(strings.Join(visible, " | "), 240))
	}

	if dumpOut != "" {
		out, err := os.Create(dumpOut)
		if err != nil {
			return fmt.Errorf("open dump-out %q: %w", dumpOut, err)
		}
		defer out.Close()
		for i := 0; i < totalRowsAfter; i++ {
			r := startRow - 1 + i
			row := rows[r]
			end := maxCol
			if end > len(row) {
				end = len(row)
			}
			visible := row[startCol-1 : end]
			fmt.Fprintln(out, strings.Join(visible, "\t"))
		}
		fmt.Printf("  full spill written to %s\n", dumpOut)
	}
	return nil
}

func abbreviate(s string, n int) string {
	if len(s) <= n {
		return s
	}
	return s[:n] + "…"
}

func maxU64(vs ...uint64) uint64 {
	var m uint64
	for _, v := range vs {
		if v > m {
			m = v
		}
	}
	return m
}
