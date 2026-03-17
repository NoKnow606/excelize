package excelize

import (
	"os"
	"testing"
)

// TestSUMIFSAfterSaveReload tests if the issue appears after saving and reloading
func TestSUMIFSAfterSaveReload(t *testing.T) {
	// Create file with data
	f := NewFile()
	f.SetCellValue("Sheet1", "B2", "12677910539")
	f.SetCellValue("Sheet1", "E2", "-")
	f.SetCellValue("Sheet1", "J2", 29)
	f.SetCellFormula("Sheet1", "A1", `=SUMIFS(J:J,B:B,"12677910539",E:E,"-")`)

	// Calculate before saving
	r1, _ := f.CalcCellValue("Sheet1", "A1")
	t.Logf("Before save: '%s'", r1)

	// Save to temp file
	tmpfile := "/tmp/test_sumifs.xlsx"
	if err := f.SaveAs(tmpfile); err != nil {
		t.Fatalf("Save failed: %v", err)
	}
	f.Close()
	defer os.Remove(tmpfile)

	// Reload file
	f2, err := OpenFile(tmpfile)
	if err != nil {
		t.Fatalf("Open failed: %v", err)
	}
	defer f2.Close()

	// Calculate after reloading
	r2, err := f2.CalcCellValue("Sheet1", "A1")
	if err != nil {
		t.Fatalf("Calc after reload failed: %v", err)
	}
	t.Logf("After reload: '%s'", r2)

	// Check data values after reload
	b2, _ := f2.GetCellValue("Sheet1", "B2")
	e2, _ := f2.GetCellValue("Sheet1", "E2")
	j2, _ := f2.GetCellValue("Sheet1", "J2")
	t.Logf("Data after reload: B2='%s', E2='%s', J2='%s'", b2, e2, j2)

	if r1 != r2 {
		t.Errorf("Results differ: before='%s', after='%s'", r1, r2)
	}

	if r2 != "29" {
		t.Errorf("Expected '29', got '%s'", r2)
	}
}
