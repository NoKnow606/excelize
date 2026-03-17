package excelize

import (
	"testing"
)

// TestUserFileDetailed performs detailed analysis of the bug
func TestUserFileDetailed(t *testing.T) {
	f, err := OpenFile("/Users/dengwei/Downloads/spreadsheet_698a06cba56003ebf8667f72_1770793392652.xlsx")
	if err != nil {
		t.Fatalf("Failed to open file: %v", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	dataSheet := "大盘"
	statsSheet := "链接数据统计"

	// Get the formula
	formula, _ := f.GetCellFormula(statsSheet, "E2")
	t.Logf("Formula: %s", formula)

	// Get ID from B2
	idValue, _ := f.GetCellValue(statsSheet, "B2")
	t.Logf("Criteria 1 (ID): '%s' (len=%d)", idValue, len(idValue))

	// Check the problematic row directly
	// From previous test, we know row 479 matches
	bVal, _ := f.GetCellValue(dataSheet, "B479")
	eVal, _ := f.GetCellValue(dataSheet, "E479")
	jVal, _ := f.GetCellValue(dataSheet, "J479")

	t.Logf("Data row 479: B='%s' (len=%d), E='%s' (len=%d, bytes=%v), J='%s'",
		bVal, len(bVal), eVal, len(eVal), []byte(eVal), jVal)

	// Compare ID values
	t.Logf("ID match: '%s' == '%s' ? %v", bVal, idValue, bVal == idValue)

	// Check E column value in detail
	t.Logf("E column: is it a dash? '%s' == '-' ? %v", eVal, eVal == "-")
	t.Logf("E column byte representation: %v vs dash %v", []byte(eVal), []byte("-"))

	// Now calculate the formula
	result, err := f.CalcCellValue(statsSheet, "E2")
	if err != nil {
		t.Fatalf("Calc failed: %v", err)
	}

	t.Logf("Calculated result: '%s'", result)

	// Try to manually invoke SUMIFS to see what happens
	// We'll test both with string literal and with cell reference
	f2 := NewFile()
	defer f2.Close()

	// Copy the exact data
	f2.SetCellValue("Sheet1", "B2", bVal)
	f2.SetCellValue("Sheet1", "E2", eVal)
	f2.SetCellValue("Sheet1", "J2", jVal)

	// Test with string literal
	f2.SetCellFormula("Sheet1", "A1", `=SUMIFS(J:J,B:B,"`+idValue+`",E:E,"-")`)
	r1, _ := f2.CalcCellValue("Sheet1", "A1")
	t.Logf("Test formula 1 result: %s", r1)

	// Test with cell references
	f2.SetCellValue("Sheet1", "C1", idValue)
	f2.SetCellValue("Sheet1", "D1", "-")
	f2.SetCellFormula("Sheet1", "A2", `=SUMIFS(J:J,B:B,C1,E:E,D1)`)
	r2, _ := f2.CalcCellValue("Sheet1", "A2")
	t.Logf("Test formula 2 result: %s", r2)
}
