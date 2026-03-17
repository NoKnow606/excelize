package excelize

import (
	"testing"
)

// TestSUMIFSCrossSheetStringLiteral tests SUMIFS with cross-sheet references and string literals
func TestSUMIFSCrossSheetStringLiteral(t *testing.T) {
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

	// Set up data
	f.SetCellValue(dataSheet, "A1", "Value")
	f.SetCellValue(dataSheet, "A2", "-")
	f.SetCellValue(dataSheet, "A3", "test")
	f.SetCellValue(dataSheet, "A4", "-")

	f.SetCellValue(dataSheet, "B1", "Count")
	f.SetCellValue(dataSheet, "B2", 10)
	f.SetCellValue(dataSheet, "B3", 20)
	f.SetCellValue(dataSheet, "B4", 30)

	// Create formula sheet
	formulaSheet := "Formulas"
	f.NewSheet(formulaSheet)

	// Test SUMIF with cross-sheet reference and string literal "-"
	formula1 := `=SUMIF(Data!A:A,"-",Data!B:B)`
	f.SetCellFormula(formulaSheet, "A1", formula1)

	result1, err := f.CalcCellValue(formulaSheet, "A1")
	if err != nil {
		t.Fatalf("CalcCellValue failed for SUMIF: %v", err)
	}

	t.Logf("SUMIF(Data!A:A,\"-\",Data!B:B) result: '%s'", result1)

	// Expected: 10 + 30 = 40
	if result1 != "40" {
		t.Errorf("SUMIF with dash literal (cross-sheet): expected '40', got '%s'", result1)
	}

	// Test SUMIFS with cross-sheet reference and string literal "-"
	formula2 := `=SUMIFS(Data!B:B,Data!A:A,"-")`
	f.SetCellFormula(formulaSheet, "A2", formula2)

	result2, err := f.CalcCellValue(formulaSheet, "A2")
	if err != nil {
		t.Fatalf("CalcCellValue failed for SUMIFS: %v", err)
	}

	t.Logf("SUMIFS(Data!B:B,Data!A:A,\"-\") result: '%s'", result2)

	// Expected: 10 + 30 = 40
	if result2 != "40" {
		t.Errorf("SUMIFS with dash literal (cross-sheet): expected '40', got '%s'", result2)
	}

	// Test with cell reference
	f.SetCellValue(formulaSheet, "C1", "-")
	formula3 := `=SUMIFS(Data!B:B,Data!A:A,C1)`
	f.SetCellFormula(formulaSheet, "A3", formula3)

	result3, err := f.CalcCellValue(formulaSheet, "A3")
	if err != nil {
		t.Fatalf("CalcCellValue failed for SUMIFS with cell ref: %v", err)
	}

	t.Logf("SUMIFS(Data!B:B,Data!A:A,C1) where C1='-' result: '%s'", result3)

	// Expected: 10 + 30 = 40
	if result3 != "40" {
		t.Errorf("SUMIFS with cell reference (cross-sheet): expected '40', got '%s'", result3)
	}
}
