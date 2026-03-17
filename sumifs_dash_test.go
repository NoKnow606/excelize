package excelize

import (
	"testing"
)

// TestSUMIFSDashCriteria tests SUMIFS with "-" criteria matching
// This reproduces the bug where "-" doesn't match "-" in SUMIFS
func TestSUMIFSDashCriteria(t *testing.T) {
	// Create a test file mimicking the structure described in the issue
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	// Create sheet "大盘" with data
	dataSheet := "大盘"
	if err := f.SetSheetName("Sheet1", dataSheet); err != nil {
		t.Fatal(err)
	}

	// Add header row
	f.SetCellValue(dataSheet, "B1", "ID")
	f.SetCellValue(dataSheet, "E1", "店铺")
	f.SetCellValue(dataSheet, "J1", "访客")

	// Add data rows with various values including "-"
	testData := []struct {
		row   int
		id    string
		shop  string
		value int
	}{
		{2, "12677910539", "-", 29},
		{3, "12977912637", "店A", 1127},
		{4, "12677910539", "店B", 6},
		{5, "12977912637", "-", 156},
	}

	for _, data := range testData {
		cellB, _ := CoordinatesToCellName(2, data.row)
		cellE, _ := CoordinatesToCellName(5, data.row)
		cellJ, _ := CoordinatesToCellName(10, data.row)
		f.SetCellValue(dataSheet, cellB, data.id)
		f.SetCellValue(dataSheet, cellE, data.shop)
		f.SetCellValue(dataSheet, cellJ, data.value)
	}

	// Create sheet "链接数据统计" with SUMIFS formulas
	statsSheet := "链接数据统计"
	f.NewSheet(statsSheet)

	// Add headers
	f.SetCellValue(statsSheet, "B1", "ID")
	f.SetCellValue(statsSheet, "E1", "Total访客")

	// Add test cases
	f.SetCellValue(statsSheet, "B2", "12677910539")

	// This is the problematic formula: SUMIFS(大盘!J:J,大盘!B:B,B2,大盘!E:E,"-")
	// Expected: 29 (matching row 2 in 大盘 where ID=12677910539 AND 店铺="-")
	// Actual: 0 (bug - doesn't match "-" with "-")
	formula := `=SUMIFS(大盘!J:J,大盘!B:B,B2,大盘!E:E,"-")`
	f.SetCellFormula(statsSheet, "E2", formula)

	// Calculate the formula
	result, err := f.CalcCellValue(statsSheet, "E2")
	if err != nil {
		t.Fatalf("CalcCellValue failed: %v", err)
	}

	// Expected result should be 29
	if result != "29" {
		t.Errorf("SUMIFS with dash criteria failed: expected '29', got '%s'", result)
		t.Log("This confirms the bug: SUMIFS doesn't match '-' string criteria correctly")

		// Debug: Check what values are actually being compared
		val1, _ := f.GetCellValue(dataSheet, "E2")
		val2 := "-"
		t.Logf("Cell value E2: '%s', Criteria: '%s', Are they equal? %v", val1, val2, val1 == val2)
	}

	// Also test with different criteria to verify the function works in general
	f.SetCellValue(statsSheet, "B3", "12977912637")
	formula2 := `=SUMIFS(大盘!J:J,大盘!B:B,B3,大盘!E:E,"店A")`
	f.SetCellFormula(statsSheet, "E3", formula2)

	result2, err := f.CalcCellValue(statsSheet, "E3")
	if err != nil {
		t.Fatalf("CalcCellValue failed for formula2: %v", err)
	}

	if result2 != "1127" {
		t.Errorf("SUMIFS with '店A' criteria: expected '1127', got '%s'", result2)
	}
}

// TestSUMIFSEmptyStringCriteria tests SUMIFS with empty string criteria
func TestSUMIFSEmptyStringCriteria(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	sheet := "Sheet1"

	// Set up data
	f.SetCellValue(sheet, "A1", "ID")
	f.SetCellValue(sheet, "B1", "Status")
	f.SetCellValue(sheet, "C1", "Value")

	f.SetCellValue(sheet, "A2", "1")
	f.SetCellValue(sheet, "B2", "")
	f.SetCellValue(sheet, "C2", 100)

	f.SetCellValue(sheet, "A3", "2")
	f.SetCellValue(sheet, "B3", "active")
	f.SetCellValue(sheet, "C3", 200)

	f.SetCellValue(sheet, "A4", "1")
	f.SetCellValue(sheet, "B4", "")
	f.SetCellValue(sheet, "C4", 150)

	// Test empty string criteria
	formula := `=SUMIFS(C:C,A:A,"1",B:B,"")`
	f.SetCellFormula(sheet, "D1", formula)

	result, err := f.CalcCellValue(sheet, "D1")
	if err != nil {
		t.Fatalf("CalcCellValue failed: %v", err)
	}

	// Should match rows 2 and 4 (ID=1 and Status="")
	if result != "250" {
		t.Errorf("SUMIFS with empty string criteria: expected '250', got '%s'", result)
	}
}
