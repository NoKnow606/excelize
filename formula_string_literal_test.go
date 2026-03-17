package excelize

import (
	"testing"
)

// TestFormulaStringLiteralParsing tests how string literals in formulas are parsed
func TestFormulaStringLiteralParsing(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	sheet := "Sheet1"

	// Set up simple data
	f.SetCellValue(sheet, "A1", "Value")
	f.SetCellValue(sheet, "A2", "-")
	f.SetCellValue(sheet, "A3", "test")
	f.SetCellValue(sheet, "A4", "-")

	f.SetCellValue(sheet, "B1", "Count")
	f.SetCellValue(sheet, "B2", 10)
	f.SetCellValue(sheet, "B3", 20)
	f.SetCellValue(sheet, "B4", 30)

	// Test SUMIF with string literal "-"
	formula1 := `=SUMIF(A:A,"-",B:B)`
	f.SetCellFormula(sheet, "C1", formula1)

	result1, err := f.CalcCellValue(sheet, "C1")
	if err != nil {
		t.Fatalf("CalcCellValue failed for SUMIF: %v", err)
	}

	t.Logf("SUMIF(A:A,\"-\",B:B) result: '%s'", result1)

	// Expected: 10 + 30 = 40 (rows 2 and 4)
	if result1 != "40" {
		t.Errorf("SUMIF with dash literal: expected '40', got '%s'", result1)
	}

	// Test SUMIFS with string literal "-"
	formula2 := `=SUMIFS(B:B,A:A,"-")`
	f.SetCellFormula(sheet, "C2", formula2)

	result2, err := f.CalcCellValue(sheet, "C2")
	if err != nil {
		t.Fatalf("CalcCellValue failed for SUMIFS: %v", err)
	}

	t.Logf("SUMIFS(B:B,A:A,\"-\") result: '%s'", result2)

	// Expected: 10 + 30 = 40
	if result2 != "40" {
		t.Errorf("SUMIFS with dash literal: expected '40', got '%s'", result2)
	}

	// Test with cell reference for comparison
	f.SetCellValue(sheet, "D1", "-")
	formula3 := `=SUMIFS(B:B,A:A,D1)`
	f.SetCellFormula(sheet, "C3", formula3)

	result3, err := f.CalcCellValue(sheet, "C3")
	if err != nil {
		t.Fatalf("CalcCellValue failed for SUMIFS with cell ref: %v", err)
	}

	t.Logf("SUMIFS(B:B,A:A,D1) where D1='-' result: '%s'", result3)

	// Expected: 10 + 30 = 40
	if result3 != "40" {
		t.Errorf("SUMIFS with cell reference: expected '40', got '%s'", result3)
	}

	// Test COUNTIF with string literal
	formula4 := `=COUNTIF(A:A,"-")`
	f.SetCellFormula(sheet, "C4", formula4)

	result4, err := f.CalcCellValue(sheet, "C4")
	if err != nil {
		t.Fatalf("CalcCellValue failed for COUNTIF: %v", err)
	}

	t.Logf("COUNTIF(A:A,\"-\") result: '%s'", result4)

	// Expected: 2 (rows 2 and 4)
	if result4 != "2" {
		t.Errorf("COUNTIF with dash literal: expected '2', got '%s'", result4)
	}
}
