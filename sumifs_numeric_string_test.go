package excelize

import (
	"testing"
)

// TestSUMIFSNumericStringID tests with numeric string IDs
func TestSUMIFSNumericStringID(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Test with short ID
	f.SetCellValue("Sheet1", "B2", "A")
	f.SetCellValue("Sheet1", "E2", "-")
	f.SetCellValue("Sheet1", "J2", 100)

	f.SetCellFormula("Sheet1", "A1", `=SUMIFS(J:J,B:B,"A",E:E,"-")`)
	r1, _ := f.CalcCellValue("Sheet1", "A1")
	t.Logf("Short ID 'A' result: '%s'", r1)

	// Test with numeric string ID (like phone number)
	f.SetCellValue("Sheet1", "B3", "12677910539")
	f.SetCellValue("Sheet1", "E3", "-")
	f.SetCellValue("Sheet1", "J3", 200)

	f.SetCellFormula("Sheet1", "A2", `=SUMIFS(J:J,B:B,"12677910539",E:E,"-")`)
	r2, _ := f.CalcCellValue("Sheet1", "A2")
	t.Logf("Numeric string '12677910539' result: '%s'", r2)

	// Test with cell reference
	f.SetCellValue("Sheet1", "C1", "12677910539")
	f.SetCellFormula("Sheet1", "A3", `=SUMIFS(J:J,B:B,C1,E:E,"-")`)
	r3, _ := f.CalcCellValue("Sheet1", "A3")
	t.Logf("Cell ref to '12677910539' result: '%s'", r3)

	if r1 != "100" {
		t.Errorf("Short ID: expected '100', got '%s'", r1)
	}
	if r2 != "200" {
		t.Errorf("Numeric string literal: expected '200', got '%s'", r2)
	}
	if r3 != "200" {
		t.Errorf("Cell ref: expected '200', got '%s'", r3)
	}
}
