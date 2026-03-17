package excelize

import (
	"testing"
)

// TestSUMIFSSameSheetVsCrossSheet compares same-sheet vs cross-sheet references
func TestSUMIFSSameSheetVsCrossSheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Set up data in Sheet1
	f.SetCellValue("Sheet1", "B2", "A")
	f.SetCellValue("Sheet1", "E2", "-")
	f.SetCellValue("Sheet1", "J2", 100)

	// Test 1: Same-sheet reference (fails?)
	f.SetCellFormula("Sheet1", "A1", `=SUMIFS(J:J,B:B,"A",E:E,"-")`)
	r1, _ := f.CalcCellValue("Sheet1", "A1")
	t.Logf("Same-sheet result: '%s'", r1)

	// Test 2: Cross-sheet reference from Sheet2
	f.NewSheet("Sheet2")
	f.SetCellFormula("Sheet2", "A1", `=SUMIFS(Sheet1!J:J,Sheet1!B:B,"A",Sheet1!E:E,"-")`)
	r2, _ := f.CalcCellValue("Sheet2", "A1")
	t.Logf("Cross-sheet result: '%s'", r2)

	if r1 != r2 {
		t.Errorf("Results differ: same-sheet='%s', cross-sheet='%s'", r1, r2)
	}

	if r1 != "100" {
		t.Errorf("Expected '100', got same-sheet='%s', cross-sheet='%s'", r1, r2)
	}
}
