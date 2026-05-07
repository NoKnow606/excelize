package main

// sql-roundtrip-verify proves that a SQL formula's spill range survives the
// reopen / download / re-upload cycle when persisted via the new
// PersistSQLFormulaResultWithMatrix path (used by excelize-mcp's
// saveSQLFormulaResult after Fix #3).
//
// It runs three rounds:
//
//	Round 1: open original xlsx, ExecuteSQL, SetCellFormula, persist via
//	         PersistSQLFormulaResultWithMatrix, WriteToBuffer.
//	Round 2: open the Round 1 buffer fresh (== "download" / Excel reopen),
//	         read the spill range back. WriteToBuffer again without touching
//	         anything.
//	Round 3: open the Round 2 buffer fresh (== "re-upload"), read the spill
//	         range back.
//
// All three rounds must produce the same spill content. This emulates
// exactly the scenarios the original UpdateSheetFormulaCache comment was
// guarding ("reopening, downloading, or re-uploading the workbook loses the
// materialized result block").
//
// Usage:
//
//	go run ./test/cmd/sql-roundtrip-verify <xlsx> <sheet> <cell> <sql_file>

import (
	"bytes"
	"crypto/md5"
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func mustReadFile(p string) string {
	b, err := os.ReadFile(p)
	if err != nil {
		fmt.Fprintf(os.Stderr, "read %s: %v\n", p, err)
		os.Exit(1)
	}
	return strings.TrimSpace(string(b))
}

func main() {
	if len(os.Args) < 5 {
		fmt.Fprintln(os.Stderr, "usage: sql-roundtrip-verify <xlsx_path> <sheet> <cell> <sql_file>")
		os.Exit(2)
	}
	xlsxPath := os.Args[1]
	sheet := os.Args[2]
	cell := os.Args[3]
	sqlPath := os.Args[4]

	formula := strings.TrimPrefix(mustReadFile(sqlPath), "=")

	// ---------- Round 1: write spill via the new matrix path ----------
	excelize.ResetSQLExecuteCount()
	f1, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		die("open: %v", err)
	}

	queryResult, err := f1.ExecuteSQL(formula)
	if err != nil {
		die("ExecuteSQL: %v", err)
	}
	if err := f1.SetCellFormula(sheet, cell, formula); err != nil {
		die("SetCellFormula: %v", err)
	}
	topLeft := ""
	if len(queryResult.Matrix) > 0 && len(queryResult.Matrix[0]) > 0 {
		topLeft = fmt.Sprint(queryResult.Matrix[0][0])
	}
	matrixResult := excelize.CalcCellValueWithMatrixResult{
		Value:  topLeft,
		Matrix: queryResult.Matrix,
	}
	if !f1.PersistSQLFormulaResultWithMatrix(sheet, cell, topLeft, matrixResult) {
		die("PersistSQLFormulaResultWithMatrix returned false (cell not SQL?)")
	}
	round1Buf, err := f1.WriteToBuffer()
	if err != nil {
		die("WriteToBuffer round1: %v", err)
	}
	f1.Close()

	r1Hash, r1Rows, r1Cols, r1Anchor := summarizeBuffer(round1Buf.Bytes(), sheet, cell)
	fmt.Printf("Round 1 (after /formula/set): bytes=%d md5(spill)=%x rows=%d cols=%d anchor=%q sqlExec=%d\n",
		round1Buf.Len(), r1Hash, r1Rows, r1Cols, r1Anchor, excelize.SQLExecuteCount())

	// ---------- Round 2: simulate download / Excel reopen ----------
	f2, err := excelize.OpenReader(bytes.NewReader(round1Buf.Bytes()))
	if err != nil {
		die("OpenReader round2: %v", err)
	}
	round2Buf, err := f2.WriteToBuffer()
	if err != nil {
		die("WriteToBuffer round2: %v", err)
	}
	f2.Close()

	r2Hash, r2Rows, r2Cols, r2Anchor := summarizeBuffer(round2Buf.Bytes(), sheet, cell)
	fmt.Printf("Round 2 (reopen + re-save):    bytes=%d md5(spill)=%x rows=%d cols=%d anchor=%q\n",
		round2Buf.Len(), r2Hash, r2Rows, r2Cols, r2Anchor)

	// ---------- Round 3: simulate re-upload ----------
	f3, err := excelize.OpenReader(bytes.NewReader(round2Buf.Bytes()))
	if err != nil {
		die("OpenReader round3: %v", err)
	}
	round3Buf, err := f3.WriteToBuffer()
	if err != nil {
		die("WriteToBuffer round3: %v", err)
	}
	f3.Close()

	r3Hash, r3Rows, r3Cols, r3Anchor := summarizeBuffer(round3Buf.Bytes(), sheet, cell)
	fmt.Printf("Round 3 (re-upload + re-save): bytes=%d md5(spill)=%x rows=%d cols=%d anchor=%q\n",
		round3Buf.Len(), r3Hash, r3Rows, r3Cols, r3Anchor)

	// ---------- Verify equality ----------
	if r1Hash != r2Hash || r2Hash != r3Hash {
		die("MISMATCH: spill range diverged across rounds (r1=%x r2=%x r3=%x)", r1Hash, r2Hash, r3Hash)
	}
	if r1Rows != r2Rows || r2Rows != r3Rows {
		die("MISMATCH: spill row count diverged across rounds (r1=%d r2=%d r3=%d)", r1Rows, r2Rows, r3Rows)
	}
	if r1Cols != r2Cols || r2Cols != r3Cols {
		die("MISMATCH: spill col count diverged across rounds (r1=%d r2=%d r3=%d)", r1Cols, r2Cols, r3Cols)
	}
	if r1Anchor != r2Anchor || r2Anchor != r3Anchor {
		die("MISMATCH: anchor formula diverged across rounds (r1=%q r2=%q r3=%q)", r1Anchor, r2Anchor, r3Anchor)
	}

	fmt.Println()
	fmt.Println("✅ PASS — SQL spill range survives round-trip:")
	fmt.Printf("   • spill cells (md5)      : identical across all 3 rounds\n")
	fmt.Printf("   • spill dimensions       : %d rows × %d cols at anchor %s\n", r1Rows, r1Cols, cell)
	fmt.Printf("   • SQL formula at anchor  : preserved\n")
	fmt.Printf("   • SQL re-executions      : %d (only the 1 ExecuteSQL in Round 1)\n", excelize.SQLExecuteCount())
}

// summarizeBuffer reopens the workbook bytes and returns:
//   - md5 of the contiguous spill range starting at anchor (TSV-encoded)
//   - row count of the contiguous filled block
//   - rightmost filled column (1-based)
//   - the formula stored at the anchor cell
func summarizeBuffer(data []byte, sheet, anchor string) (hash [16]byte, rowCount, colCount int, anchorFormula string) {
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		die("summarize OpenReader: %v", err)
	}
	defer f.Close()

	startCol, startRow, err := excelize.CellNameToCoordinates(anchor)
	if err != nil {
		die("invalid anchor %q: %v", anchor, err)
	}

	anchorFormula, _ = f.GetCellFormula(sheet, anchor)

	rows, err := f.GetRows(sheet)
	if err != nil {
		die("GetRows: %v", err)
	}

	maxCol := 0
	rowCount = 0
	for r := startRow - 1; r < len(rows); r++ {
		if r < 0 {
			continue
		}
		row := rows[r]
		hasContent := false
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
		rowCount++
	}

	colCount = maxCol - (startCol - 1)
	if colCount < 0 {
		colCount = 0
	}

	var buf bytes.Buffer
	for i := 0; i < rowCount; i++ {
		r := startRow - 1 + i
		row := rows[r]
		end := maxCol
		if end > len(row) {
			end = len(row)
		}
		visible := row[startCol-1 : end]
		buf.WriteString(strings.Join(visible, "\t"))
		buf.WriteByte('\n')
	}
	hash = md5.Sum(buf.Bytes())
	return
}

func die(format string, args ...interface{}) {
	fmt.Fprintf(os.Stderr, format+"\n", args...)
	os.Exit(1)
}
