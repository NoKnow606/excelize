package excelize

import (
	"testing"
)

// TestSUMIFSMinimal is a minimal test case for the dash criteria bug
func TestSUMIFSMinimal(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	// Create data sheet with Chinese name
	dataSheet := "大盘"
	if err := f.SetSheetName("Sheet1", dataSheet); err != nil {
		t.Fatal(err)
	}

	// Set header
	f.SetCellValue(dataSheet, "E1", "店铺")
	f.SetCellValue(dataSheet, "J1", "访客")

	// Set ONE data row with "-"
	f.SetCellValue(dataSheet, "E2", "-")
	f.SetCellValue(dataSheet, "J2", 100)

	// Create formula sheet
	formulaSheet := "Stats"
	f.NewSheet(formulaSheet)

	// Test 1: SUMIF with string literal
	formula1 := `=SUMIF(大盘!E:E,"-",大盘!J:J)`
	f.SetCellFormula(formulaSheet, "A1", formula1)

	result1, err := f.CalcCellValue(formulaSheet, "A1")
	if err != nil {
		t.Fatalf("SUMIF calc failed: %v", err)
	}

	t.Logf("SUMIF result: '%s'", result1)
	if result1 != "100" {
		t.Errorf("SUMIF: expected '100', got '%s'", result1)
	}

	// Test 2: SUMIFS with string literal
	formula2 := `=SUMIFS(大盘!J:J,大盘!E:E,"-")`
	f.SetCellFormula(formulaSheet, "A2", formula2)

	result2, err := f.CalcCellValue(formulaSheet, "A2")
	if err != nil {
		t.Fatalf("SUMIFS calc failed: %v", err)
	}

	t.Logf("SUMIFS result: '%s'", result2)
	if result2 != "100" {
		t.Errorf("SUMIFS: expected '100', got '%s'", result2)
	}
}
