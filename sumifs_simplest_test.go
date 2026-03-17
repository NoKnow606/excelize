package excelize

import (
	"testing"
)

// TestSUMIFSSimplestCase tests the absolutely simplest case
func TestSUMIFSSimplestCase(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Set data in row 2
	f.SetCellValue("Sheet1", "B2", "12677910539")
	f.SetCellValue("Sheet1", "E2", "-")
	f.SetCellValue("Sheet1", "J2", 29)

	// Test 1: Both criteria as string literals
	f.SetCellFormula("Sheet1", "A1", `=SUMIFS(J:J,B:B,"12677910539",E:E,"-")`)
	r1, err := f.CalcCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("Formula 1 failed: %v", err)
	}
	t.Logf("Formula 1 result: '%s'", r1)
	if r1 != "29" {
		t.Errorf("Formula 1: expected '29', got '%s'", r1)
	}

	// Test 2: First criterion as cell ref, second as string literal
	f.SetCellValue("Sheet1", "C1", "12677910539")
	f.SetCellFormula("Sheet1", "A2", `=SUMIFS(J:J,B:B,C1,E:E,"-")`)
	r2, err := f.CalcCellValue("Sheet1", "A2")
	if err != nil {
		t.Fatalf("Formula 2 failed: %v", err)
	}
	t.Logf("Formula 2 result: '%s'", r2)
	if r2 != "29" {
		t.Errorf("Formula 2: expected '29', got '%s'", r2)
	}

	// Test 3: Both criteria as cell refs
	f.SetCellValue("Sheet1", "D1", "-")
	f.SetCellFormula("Sheet1", "A3", `=SUMIFS(J:J,B:B,C1,E:E,D1)`)
	r3, err := f.CalcCellValue("Sheet1", "A3")
	if err != nil {
		t.Fatalf("Formula 3 failed: %v", err)
	}
	t.Logf("Formula 3 result: '%s'", r3)
	if r3 != "29" {
		t.Errorf("Formula 3: expected '29', got '%s'", r3)
	}

	// Test 4: Use J2:J2 instead of J:J to narrow the range
	f.SetCellFormula("Sheet1", "A4", `=SUMIFS(J2:J2,B2:B2,"12677910539",E2:E2,"-")`)
	r4, err := f.CalcCellValue("Sheet1", "A4")
	if err != nil {
		t.Fatalf("Formula 4 failed: %v", err)
	}
	t.Logf("Formula 4 (narrow range) result: '%s'", r4)
	if r4 != "29" {
		t.Errorf("Formula 4: expected '29', got '%s'", r4)
	}
}
