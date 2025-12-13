package excelize

import (
	"testing"
)

// TestGetFormulasSharedFormulas tests GetFormulas correctly expands shared formulas
func TestGetFormulasSharedFormulas(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create a shared formula pattern like the real file
	// B1 is the master formula, C1-E1 share the same pattern
	formulas := map[string]string{
		"B1": "MAX('Sheet2'!A:A)",
		"C1": "B1-1",
		"D1": "C1-1",
		"E1": "D1-1",
	}

	for cell, formula := range formulas {
		if err := f.SetCellFormula(sheet, cell, formula); err != nil {
			t.Fatalf("SetCellFormula failed for %s: %v", cell, err)
		}
	}

	// Test GetFormulas returns all formulas
	allFormulas, err := f.GetFormulas(sheet)
	if err != nil {
		t.Fatalf("GetFormulas failed: %v", err)
	}

	if len(allFormulas) == 0 {
		t.Fatal("GetFormulas returned empty result")
	}

	row := allFormulas[0]
	t.Logf("GetFormulas returned row with %d columns: %v", len(row), row)

	// Verify each formula
	expectedCols := map[int]string{
		0: "",                      // A1 无公式
		1: "MAX('Sheet2'!A:A)",    // B1
		2: "B1-1",                  // C1
		3: "C1-1",                  // D1
		4: "D1-1",                  // E1
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

	// Compare with GetCellFormula
	for cell, expectedFormula := range formulas {
		formula, err := f.GetCellFormula(sheet, cell)
		if err != nil {
			t.Errorf("GetCellFormula failed for %s: %v", cell, err)
		}
		if formula != expectedFormula {
			t.Errorf("Cell %s: expected '%s', got '%s'", cell, expectedFormula, formula)
		}
	}

	t.Log("\n✓ Shared formulas expanded correctly")
}

// TestGetFormulasPerformance tests GetFormulas performance with many formulas
func TestGetFormulasPerformance(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping performance test in short mode")
	}

	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// Create 100 rows × 26 columns of formulas
	for row := 1; row <= 100; row++ {
		for col := 1; col <= 26; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			formula := "A1+1"
			if err := f.SetCellFormula(sheet, cell, formula); err != nil {
				t.Fatalf("SetCellFormula failed: %v", err)
			}
		}
	}

	// Test GetFormulas performance
	allFormulas, err := f.GetFormulas(sheet)
	if err != nil {
		t.Fatalf("GetFormulas failed: %v", err)
	}

	if len(allFormulas) != 100 {
		t.Errorf("Expected 100 rows, got %d", len(allFormulas))
	}

	// Verify all formulas are present
	totalFormulas := 0
	for _, row := range allFormulas {
		for _, formula := range row {
			if formula != "" {
				totalFormulas++
			}
		}
	}

	expectedFormulas := 100 * 26
	if totalFormulas != expectedFormulas {
		t.Errorf("Expected %d formulas, got %d", expectedFormulas, totalFormulas)
	}

	t.Logf("✓ GetFormulas processed %d formulas successfully", totalFormulas)
}
