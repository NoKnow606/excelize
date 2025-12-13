package excelize

import (
	"testing"
)

// TestGetFormulasMultipleColumns tests GetFormulas returns all formulas correctly
func TestGetFormulasMultipleColumns(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Setup: Set formulas in multiple columns
	formulas := map[string]string{
		"B1": "MAX('库存台账-all'!A:A)",
		"C1": "B1-1",
		"D1": "C1-1",
		"E1": "D1-1",
	}

	for cell, formula := range formulas {
		if err := f.SetCellFormula(sheet, cell, formula); err != nil {
			t.Fatalf("SetCellFormula failed for %s: %v", cell, err)
		}
	}

	// Test 1: GetCellFormula (逐个查询)
	t.Log("\n=== Test 1: GetCellFormula (逐个查询) ===")
	for cell, expectedFormula := range formulas {
		formula, err := f.GetCellFormula(sheet, cell)
		if err != nil {
			t.Errorf("GetCellFormula failed for %s: %v", cell, err)
		}
		if formula != expectedFormula {
			t.Errorf("Cell %s: expected formula '%s', got '%s'", cell, expectedFormula, formula)
		}
		t.Logf("%s: Formula: '%s' ✓", cell, formula)
	}

	// Test 2: GetFormulas (批量查询)
	t.Log("\n=== Test 2: GetFormulas (批量查询) ===")
	allFormulas, err := f.GetFormulas(sheet)
	if err != nil {
		t.Fatalf("GetFormulas failed: %v", err)
	}

	if len(allFormulas) == 0 {
		t.Fatal("GetFormulas returned empty result")
	}

	// 第一行应该有公式
	if len(allFormulas) < 1 {
		t.Fatal("GetFormulas should return at least 1 row")
	}

	row := allFormulas[0]
	t.Logf("GetFormulas returned row: %v", row)
	t.Logf("Row length: %d", len(row))

	// 验证每列的公式
	expectedCols := map[int]string{
		0: "",                                 // A1 无公式
		1: "MAX('库存台账-all'!A:A)",          // B1
		2: "B1-1",                             // C1
		3: "C1-1",                             // D1
		4: "D1-1",                             // E1
	}

	for col, expectedFormula := range expectedCols {
		if col >= len(row) {
			if expectedFormula != "" {
				t.Errorf("Column %d missing: expected formula '%s'", col, expectedFormula)
			}
			continue
		}

		actualFormula := row[col]
		if actualFormula != expectedFormula {
			t.Errorf("Column %d: expected '%s', got '%s'", col, expectedFormula, actualFormula)
		} else {
			t.Logf("Column %d: '%s' ✓", col, actualFormula)
		}
	}

	// 确保没有丢失公式
	nonEmptyCount := 0
	for _, formula := range row {
		if formula != "" {
			nonEmptyCount++
		}
	}

	if nonEmptyCount != len(formulas) {
		t.Errorf("Expected %d non-empty formulas, got %d", len(formulas), nonEmptyCount)
	}

	t.Log("\n✓ All formulas returned correctly")
}

// TestGetFormulasPreservesColumnPositions tests column alignment
func TestGetFormulasPreservesColumnPositions(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Setup: Set formulas with gaps
	f.SetCellFormula(sheet, "A1", "1+1")
	f.SetCellFormula(sheet, "C1", "2+2")  // Skip B1
	f.SetCellFormula(sheet, "E1", "3+3")  // Skip D1

	formulas, err := f.GetFormulas(sheet)
	if err != nil {
		t.Fatalf("GetFormulas failed: %v", err)
	}

	if len(formulas) == 0 || len(formulas[0]) < 5 {
		t.Fatalf("Expected at least 5 columns, got %d", len(formulas[0]))
	}

	row := formulas[0]

	// Verify positions
	tests := []struct {
		col      int
		expected string
		desc     string
	}{
		{0, "1+1", "A1"},
		{1, "", "B1 (empty)"},
		{2, "2+2", "C1"},
		{3, "", "D1 (empty)"},
		{4, "3+3", "E1"},
	}

	for _, tt := range tests {
		if row[tt.col] != tt.expected {
			t.Errorf("%s: expected '%s', got '%s'", tt.desc, tt.expected, row[tt.col])
		} else {
			t.Logf("%s: '%s' ✓", tt.desc, row[tt.col])
		}
	}
}

// TestGetFormulasEmptySheet tests empty sheet
func TestGetFormulasEmptySheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	formulas, err := f.GetFormulas("Sheet1")
	if err != nil {
		t.Errorf("GetFormulas failed on empty sheet: %v", err)
	}

	if len(formulas) != 0 {
		t.Errorf("Expected empty result, got %d rows", len(formulas))
	}

	t.Log("✓ Empty sheet handled correctly")
}

// TestGetFormulasMultipleRows tests multiple rows with formulas
func TestGetFormulasMultipleRows(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Setup: Set formulas in multiple rows
	testData := map[string]string{
		"A1": "SUM(B1:B10)",
		"B1": "10",
		"A2": "AVERAGE(B1:B10)",
		"B2": "20",
		"A3": "MAX(B1:B10)",
		"B3": "30",
	}

	for cell, formula := range testData {
		f.SetCellFormula(sheet, cell, formula)
	}

	formulas, err := f.GetFormulas(sheet)
	if err != nil {
		t.Fatalf("GetFormulas failed: %v", err)
	}

	if len(formulas) < 3 {
		t.Fatalf("Expected at least 3 rows, got %d", len(formulas))
	}

	// Verify each row
	expected := [][]string{
		{"SUM(B1:B10)", "10"},
		{"AVERAGE(B1:B10)", "20"},
		{"MAX(B1:B10)", "30"},
	}

	for i, expectedRow := range expected {
		if i >= len(formulas) {
			t.Errorf("Row %d missing", i+1)
			continue
		}

		actualRow := formulas[i]
		for j, expectedFormula := range expectedRow {
			if j >= len(actualRow) {
				t.Errorf("Row %d, Col %d missing", i+1, j+1)
				continue
			}

			if actualRow[j] != expectedFormula {
				t.Errorf("Row %d, Col %d: expected '%s', got '%s'", i+1, j+1, expectedFormula, actualRow[j])
			}
		}
	}

	t.Log("✓ Multiple rows handled correctly")
}
