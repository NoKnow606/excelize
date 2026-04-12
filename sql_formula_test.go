package excelize

import (
	"fmt"
	"testing"
)

func TestSQLFormulaWithGIDResolverSpillsMatrix(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Sales"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Report"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Sales", [][]interface{}{
		{"Region", "Revenue"},
		{"North", 10},
		{"South", 20},
		{"South", 5},
	})

	f.SetSQLSourceResolver(SQLSourceResolverFunc(func(token string, sheetList []string) (string, error) {
		switch unquoteIdentifier(token) {
		case "gid_7":
			return "Sales", nil
		case "Sales":
			return "Sales", nil
		default:
			return "", fmt.Errorf("unexpected token %s", token)
		}
	}))

	formula := `SQL("select ""Region"", sum(""Revenue"") as ""Total"" from gid_7 group by ""Region"" order by ""Total"" desc")`
	if err := f.SetCellFormula("Report", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	result, err := f.CalcCellValueWithMatrix("Report", "A1")
	if err != nil {
		t.Fatalf("CalcCellValueWithMatrix: %v", err)
	}
	if len(result.Matrix) != 3 {
		t.Fatalf("expected header plus 2 rows, got %#v", result.Matrix)
	}
	if got := result.Matrix[0][0]; got != "Region" {
		t.Fatalf("expected header Region, got %#v", got)
	}
	if got := result.Matrix[1][0]; got != "South" {
		t.Fatalf("expected first data row South, got %#v", got)
	}
	if got := result.Matrix[1][1]; got != float64(25) {
		t.Fatalf("expected South total 25, got %#v", got)
	}
	if got := result.Matrix[2][0]; got != "North" {
		t.Fatalf("expected second data row North, got %#v", got)
	}
}

func TestSQLFormulaFindAffectedCellsByCellsTracksWholeSourceSheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	defaultSheet := f.GetSheetName(0)
	if err := f.SetSheetName(defaultSheet, "Data"); err != nil {
		t.Fatalf("SetSheetName: %v", err)
	}
	if _, err := f.NewSheet("Summary"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}

	writeSQLSheetRows(t, f, "Data", [][]interface{}{
		{"Category", "Amount"},
		{"A", 10},
		{"B", 20},
	})

	formula := `SQL("select ""Category"", sum(""Amount"") as ""Total"" from ""Data"" group by ""Category"" order by ""Category""")`
	if err := f.SetCellFormula("Summary", "A1", formula); err != nil {
		t.Fatalf("SetCellFormula: %v", err)
	}

	graph := f.buildDependencyGraph()
	node := graph.nodes["Summary!A1"]
	if node == nil {
		t.Fatal("expected Summary!A1 in dependency graph")
	}
	if !containsString(node.dependencies, "SHEET:Data") {
		t.Fatalf("expected whole-sheet dependency marker, got %#v", node.dependencies)
	}

	affected := f.findAffectedCellsByCells(graph, map[string]bool{"Data!B2": true})
	if !affected["Summary!A1"] {
		t.Fatalf("expected Summary!A1 to be affected by Data!B2 update, got %#v", affected)
	}
}

func writeSQLSheetRows(t *testing.T, f *File, sheet string, rows [][]interface{}) {
	t.Helper()
	for rowIdx, row := range rows {
		cell, err := CoordinatesToCellName(1, rowIdx+1)
		if err != nil {
			t.Fatalf("CoordinatesToCellName: %v", err)
		}
		if err := f.SetSheetRow(sheet, cell, &row); err != nil {
			t.Fatalf("SetSheetRow: %v", err)
		}
	}
}

func containsString(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}
