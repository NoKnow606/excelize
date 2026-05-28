package main

// sql-compile-bench measures the cost of CompileSQL — the path used by
// excelize-mcp's /formula/compile and /sql/compile endpoints.
//
// CompileSQL is *meant* to be a lightweight schema-level validator: it should
// only need worksheet headers to ask PostgreSQL to prepare the rewritten query.
// In practice it used to share materializeSheets with ExecuteSQL, which read
// every source row into Go memory and inserted them into temporary PostgreSQL
// tables. For workbooks with hundreds of thousands of source rows this made
// /formula/compile nearly as expensive as /formula/calculate.
//
// This benchmark exercises CompileSQL on a real workbook so we can compare
// before/after the schema-only optimization.
//
// Usage:
//
//	go run ./test/cmd/sql-compile-bench <xlsx_path> <sql_file>

import (
	"fmt"
	"os"
	"runtime"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) < 3 {
		fmt.Fprintln(os.Stderr, "usage: sql-compile-bench <xlsx_path> <sql_file>")
		os.Exit(2)
	}
	xlsxPath := os.Args[1]
	sqlPath := os.Args[2]

	body, err := os.ReadFile(sqlPath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "read %s: %v\n", sqlPath, err)
		os.Exit(1)
	}
	formula := strings.TrimSpace(string(body))

	excelize.ResetSQLExecuteCount()
	t0 := time.Now()
	printSnap("0_start", t0, 0)

	f, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()
	open := printSnap("1_after_open", t0, 0)

	t1 := time.Now()
	result, err := f.CompileSQL(formula)
	if err != nil {
		fmt.Fprintf(os.Stderr, "CompileSQL: %v\n", err)
		os.Exit(1)
	}
	cmp := printSnap("2_after_compileSQL", t1, open.totalAlloc)

	fmt.Printf("\nSummary: workbook=%s\n", xlsxPath)
	fmt.Printf("  CompileSQL wall time         = %s\n", time.Duration(cmp.wallNs).Round(time.Millisecond))
	fmt.Printf("  CompileSQL TotalAlloc delta  = %s\n", mb(cmp.totalAlloc-open.totalAlloc))
	fmt.Printf("  CompileSQL HeapAlloc delta   = %+s\n", deltaMB(cmp.heapAlloc, open.heapAlloc))
	fmt.Printf("  CompileSQL HeapInuse delta   = %+s\n", deltaMB(cmp.heapInuse, open.heapInuse))
	fmt.Printf("  CompileSQL Sys delta         = %+s\n", deltaMB(cmp.sys, open.sys))
	fmt.Printf("  ExecuteSQL invocations       = %d (expected: 0)\n", excelize.SQLExecuteCount())
	fmt.Printf("  Sources resolved             = %d  table=%q\n", len(result.Sources), result.SourceSheet)
}

type sn struct {
	wallNs                                int64
	heapAlloc, heapInuse, totalAlloc, sys uint64
}

func printSnap(label string, started time.Time, baseTotalAlloc uint64) sn {
	runtime.GC()
	var ms runtime.MemStats
	runtime.ReadMemStats(&ms)
	s := sn{
		wallNs:     time.Since(started).Nanoseconds(),
		heapAlloc:  ms.HeapAlloc,
		heapInuse:  ms.HeapInuse,
		totalAlloc: ms.TotalAlloc,
		sys:        ms.Sys,
	}
	delta := s.totalAlloc - baseTotalAlloc
	fmt.Printf(
		"  %-32s  wall=%-9s  heapAlloc=%-10s  heapInuse=%-10s  totalAlloc=%-10s  sys=%-10s  delta=%s\n",
		label,
		time.Duration(s.wallNs).Round(time.Millisecond),
		mb(s.heapAlloc),
		mb(s.heapInuse),
		mb(s.totalAlloc),
		mb(s.sys),
		mb(delta),
	)
	return s
}

func mb(n uint64) string { return fmt.Sprintf("%.1f MB", float64(n)/(1024*1024)) }

func deltaMB(after, before uint64) string {
	if after >= before {
		return fmt.Sprintf("+%.1f MB", float64(after-before)/(1024*1024))
	}
	return fmt.Sprintf("-%.1f MB", float64(before-after)/(1024*1024))
}
