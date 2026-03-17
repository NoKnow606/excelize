package excelize

import (
	"testing"
)

// TestSUMIFSTwoCriteriaWithDash tests SUMIFS with two criteria where one is a dash
func TestSUMIFSTwoCriteriaWithDash(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	// Create data sheet
	dataSheet := "Data"
	if err := f.SetSheetName("Sheet1", dataSheet); err != nil {
		t.Fatal(err)
	}

	// Headers
	f.SetCellValue(dataSheet, "B1", "ID")
	f.SetCellValue(dataSheet, "E1", "Status")
	f.SetCellValue(dataSheet, "J1", "Count")

	// Data rows
	f.SetCellValue(dataSheet, "B2", "A")
	f.SetCellValue(dataSheet, "E2", "-")
	f.SetCellValue(dataSheet, "J2", 100)

	f.SetCellValue(dataSheet, "B3", "A")
	f.SetCellValue(dataSheet, "E3", "active")
	f.SetCellValue(dataSheet, "J3", 200)

	f.SetCellValue(dataSheet, "B4", "B")
	f.SetCellValue(dataSheet, "E4", "-")
	f.SetCellValue(dataSheet, "J4", 300)

	// Create formula sheet
	formulaSheet := "Formulas"
	f.NewSheet(formulaSheet)

	// Test: SUMIFS with two criteria, one is string literal "-"
	// Expected: 100 (ID=A AND Status="-")
	formula := `=SUMIFS(Data!J:J,Data!B:B,"A",Data!E:E,"-")`
	f.SetCellFormula(formulaSheet, "A1", formula)

	result, err := f.CalcCellValue(formulaSheet, "A1")
	if err != nil {
		t.Fatalf("Calc failed: %v", err)
	}

	t.Logf("SUMIFS with two criteria result: '%s'", result)
	if result != "100" {
		t.Errorf("Expected '100', got '%s'", result)
	}

	// Test with cell reference for first criterion
	f.SetCellValue(formulaSheet, "B1", "A")
	formula2 := `=SUMIFS(Data!J:J,Data!B:B,B1,Data!E:E,"-")`
	f.SetCellFormula(formulaSheet, "A2", formula2)

	result2, err := f.CalcCellValue(formulaSheet, "A2")
	if err != nil {
		t.Fatalf("Calc failed: %v", err)
	}

	t.Logf("SUMIFS with cell ref + string literal result: '%s'", result2)
	if result2 != "100" {
		t.Errorf("Expected '100', got '%s'", result2)
	}
}
