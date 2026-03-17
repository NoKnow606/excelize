package excelize

import (
	"fmt"
	"testing"
)

// TestUserFileS actual user file from the issue
func TestUserFile(t *testing.T) {
	// Open the user's file
	f, err := OpenFile("/Users/dengwei/Downloads/spreadsheet_698a06cba56003ebf8667f72_1770793392652.xlsx")
	if err != nil {
		t.Fatalf("Failed to open file: %v", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	// Check sheet names
	sheets := f.GetSheetList()
	t.Logf("Sheets in file: %v", sheets)

	// Find the sheets
	dataSheet := "大盘"
	statsSheet := "链接数据统计"

	// Check if sheets exist
	hasDataSheet := false
	hasStatsSheet := false
	for _, sheet := range sheets {
		if sheet == dataSheet {
			hasDataSheet = true
		}
		if sheet == statsSheet {
			hasStatsSheet = true
		}
	}

	if !hasDataSheet {
		t.Fatalf("Data sheet '%s' not found", dataSheet)
	}
	if !hasStatsSheet {
		t.Fatalf("Stats sheet '%s' not found", statsSheet)
	}

	// Check the formula in cell E2 of stats sheet
	formula, err := f.GetCellFormula(statsSheet, "E2")
	if err != nil {
		t.Fatalf("Failed to get formula: %v", err)
	}
	t.Logf("Formula in %s!E2: %s", statsSheet, formula)

	// Calculate the cell
	result, err := f.CalcCellValue(statsSheet, "E2")
	if err != nil {
		t.Fatalf("Failed to calculate: %v", err)
	}
	t.Logf("Calculated result: '%s'", result)

	// Get the ID from B2
	idValue, _ := f.GetCellValue(statsSheet, "B2")
	t.Logf("ID in B2: '%s'", idValue)

	// Check data sheet for matching rows
	rows, err := f.GetRows(dataSheet)
	if err != nil {
		t.Fatalf("Failed to get rows: %v", err)
	}

	t.Logf("Data sheet has %d rows", len(rows))

	// Find rows that match the criteria
	matchCount := 0
	matchSum := 0.0
	for i, row := range rows {
		if i == 0 {
			// Header row
			continue
		}
		if len(row) > 9 {
			// Check B column (index 1) and E column (index 4)
			if len(row) > 1 && len(row) > 4 {
				bVal := ""
				eVal := ""
				jVal := ""
				if len(row) > 1 {
					bVal = row[1]
				}
				if len(row) > 4 {
					eVal = row[4]
				}
				if len(row) > 9 {
					jVal = row[9]
				}

				if bVal == idValue && eVal == "-" {
					matchCount++
					t.Logf("Match at row %d: B='%s', E='%s', J='%s'", i+1, bVal, eVal, jVal)
					// Try to parse J value
					var numVal float64
					if _, err := fmt.Sscanf(jVal, "%f", &numVal); err == nil {
						matchSum += numVal
					}
				}
			}
		}
	}

	t.Logf("Found %d matching rows, sum should be: %f", matchCount, matchSum)

	// Compare with calculated result
	if result == "0" && matchSum > 0 {
		t.Errorf("Bug confirmed: Formula returned 0 but manual calculation gives %f", matchSum)
	}
}
