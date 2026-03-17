package excelize

import (
	"fmt"
	"testing"
)

// TestSUMIFSDebug helps debug the SUMIFS issue
func TestSUMIFSDebug(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	// Create sheet "大盘" with minimal data
	dataSheet := "大盘"
	if err := f.SetSheetName("Sheet1", dataSheet); err != nil {
		t.Fatal(err)
	}

	// Set header
	f.SetCellValue(dataSheet, "B1", "ID")
	f.SetCellValue(dataSheet, "E1", "店铺")
	f.SetCellValue(dataSheet, "J1", "访客")

	// Set ONE data row with "-"
	f.SetCellValue(dataSheet, "B2", "12677910539")
	f.SetCellValue(dataSheet, "E2", "-")
	f.SetCellValue(dataSheet, "J2", 29)

	// Debug: Read back the values
	idVal, _ := f.GetCellValue(dataSheet, "B2")
	shopVal, _ := f.GetCellValue(dataSheet, "E2")
	visitorVal, _ := f.GetCellValue(dataSheet, "J2")
	t.Logf("Data sheet B2='%s', E2='%s', J2='%s'", idVal, shopVal, visitorVal)

	// Debug: Use GetRows to see what the batch function sees
	rows, err := f.GetRows(dataSheet)
	if err != nil {
		t.Fatal(err)
	}
	for i, row := range rows {
		t.Logf("Row %d: %v", i, row)
		if i < len(rows) && len(row) > 9 {
			t.Logf("  Row %d Col B (idx 1): '%s'", i, row[1])
			t.Logf("  Row %d Col E (idx 4): '%s'", i, row[4])
			t.Logf("  Row %d Col J (idx 9): '%s'", i, row[9])
		}
	}

	// Create stats sheet
	statsSheet := "链接数据统计"
	f.NewSheet(statsSheet)
	f.SetCellValue(statsSheet, "B2", "12677910539")

	// Test formula
	formula := `=SUMIFS(大盘!J:J,大盘!B:B,B2,大盘!E:E,"-")`
	f.SetCellFormula(statsSheet, "E2", formula)

	// Calculate
	result, err := f.CalcCellValue(statsSheet, "E2")
	if err != nil {
		t.Fatalf("CalcCellValue failed: %v", err)
	}

	t.Logf("SUMIFS result: '%s'", result)

	if result != "29" {
		t.Errorf("Expected '29', got '%s'", result)

		// Debug the batch calculation path
		t.Log("Testing if batch optimization is being used...")

		// Manually test the pattern extraction
		pattern := f.extractSUMIFS2DPattern(statsSheet, "E2", "SUMIFS(大盘!J:J,大盘!B:B,B2,大盘!E:E,\"-\")")
		if pattern != nil {
			t.Logf("Pattern detected: sumRange=%s, criteria1=%s, criteria2=%s",
				pattern.sumRangeRef, pattern.criteriaRange1Ref, pattern.criteriaRange2Ref)

			// Extract column info
			sumCol := extractColumnFromRange(pattern.sumRangeRef)
			criteria1Col := extractColumnFromRange(pattern.criteriaRange1Ref)
			criteria2Col := extractColumnFromRange(pattern.criteriaRange2Ref)
			t.Logf("Columns: sum=%s, crit1=%s, crit2=%s", sumCol, criteria1Col, criteria2Col)

			// Check what scanRowsAndBuildResultMap would produce
			sourceSheet := extractSheetName(pattern.sumRangeRef)
			rows, _ := f.GetRows(sourceSheet)
			resultMap := f.scanRowsAndBuildResultMap(sourceSheet, rows, sumCol, criteria1Col, criteria2Col)

			t.Logf("Result map: %+v", resultMap)

			// Check if our criteria values exist in the map
			c1, _ := f.GetCellValue(statsSheet, "B2")
			c2 := "-"
			t.Logf("Looking for c1='%s', c2='%s'", c1, c2)

			if resultMap[c1] != nil {
				t.Logf("Found c1 in map, sub-map: %+v", resultMap[c1])
				if val, ok := resultMap[c1][c2]; ok {
					t.Logf("Found result: %f", val)
				} else {
					t.Logf("c2='%s' not found in sub-map keys: %v", c2, getMapKeysFromFloat(resultMap[c1]))
				}
			} else {
				t.Logf("c1='%s' not found in result map keys: %v", c1, getTopLevelKeys(resultMap))
			}
		} else {
			t.Log("Pattern not detected - will use standard calculation")
		}
	}
}

func getMapKeysFromFloat(m map[string]float64) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, fmt.Sprintf("'%s'", k))
	}
	return keys
}

func getTopLevelKeys(m map[string]map[string]float64) []string {
	keys := make([]string, 0, len(m))
	for k := range m {
		keys = append(keys, fmt.Sprintf("'%s'", k))
	}
	return keys
}
