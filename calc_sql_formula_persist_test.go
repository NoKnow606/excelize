package excelize

import (
	"os"
	"testing"
)

func TestCalcCellValuesPersistsSQLFormulaSpillRange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	results, err := f.CalcCellValues("Report", []string{"A1"}, Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("CalcCellValues: %v", err)
	}
	if got := results["A1"]; got != "Category" {
		t.Fatalf("expected CalcCellValues result Category, got %q", got)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected spill ref A1:B3 after CalcCellValues, got %q", got)
	}
	if got, err := f.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected in-memory spill value 20 at B3, got %q err=%v", got, err)
	}

	fileName := "test_sql_formula_calc_cell_values.xlsx"
	if err := f.SaveAs(fileName); err != nil {
		t.Fatalf("SaveAs: %v", err)
	}
	defer os.Remove(fileName)

	reopened, err := OpenFile(fileName)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	defer reopened.Close()

	if got, err := reopened.GetCellValue("Report", "A1", Options{RawCellValue: true}); err != nil || got != "Category" {
		t.Fatalf("expected reopened A1 header Category, got %q err=%v", got, err)
	}
	if got, err := reopened.GetCellValue("Report", "B3", Options{RawCellValue: true}); err != nil || got != "20" {
		t.Fatalf("expected reopened B3 value 20, got %q err=%v", got, err)
	}
}

func TestCalcCellValuesOptimizedPersistsSQLFormulaSpillRange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	results, err := f.CalcCellValuesOptimized("Report", []string{"A1"}, Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("CalcCellValuesOptimized: %v", err)
	}
	if got := results["A1"]; got != "Category" {
		t.Fatalf("expected CalcCellValuesOptimized result Category, got %q", got)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected spill ref A1:B3 after CalcCellValuesOptimized, got %q", got)
	}
}

func TestCalcCellValuesConcurrentPersistsSQLFormulaSpillRange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	results, err := f.CalcCellValuesConcurrent("Report", []string{"A1"}, Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("CalcCellValuesConcurrent: %v", err)
	}
	if got := results["A1"]; got != "Category" {
		t.Fatalf("expected CalcCellValuesConcurrent result Category, got %q", got)
	}

	ws, err := f.workSheetReader("Report")
	if err != nil {
		t.Fatalf("workSheetReader: %v", err)
	}
	if got := ws.SheetData.Row[0].C[0].F.Ref; got != "A1:B3" {
		t.Fatalf("expected spill ref A1:B3 after CalcCellValuesConcurrent, got %q", got)
	}
}
