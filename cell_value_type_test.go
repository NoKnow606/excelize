package excelize

import (
	"testing"
)

// TestCellValueType checks how numeric strings are stored
func TestCellValueType(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Set a numeric string
	f.SetCellValue("Sheet1", "A1", "12677910539")

	// Read it back
	val, _ := f.GetCellValue("Sheet1", "A1")
	t.Logf("Stored value: '%s' (len=%d)", val, len(val))

	// Try to get the raw cell value
	rawVal, _ := f.GetCellValue("Sheet1", "A1", Options{RawCellValue: true})
	t.Logf("Raw value: '%s' (len=%d)", rawVal, len(rawVal))

	// Now set it as a number (not string)
	f.SetCellValue("Sheet1", "B1", 12677910539)

	val2, _ := f.GetCellValue("Sheet1", "B1")
	t.Logf("Number stored value: '%s' (len=%d)", val2, len(val2))

	// Check if they're equal
	t.Logf("Are they equal? '%s' == '%s' => %v", val, val2, val == val2)
}
