package excelize

import (
	"testing"
)

// TestFormulaArgValueFormatting tests formulaArg.Value() formatting
func TestFormulaArgValueFormatting(t *testing.T) {
	// Create a number formula arg
	arg := newNumberFormulaArg(12677910539)

	// Get its string value
	str := arg.Value()
	t.Logf("formulaArg(12677910539).Value() = '%s'", str)

	if str != "12677910539" {
		t.Errorf("Expected '12677910539', got '%s'", str)
	}

	// Test with a decimal
	arg2 := newNumberFormulaArg(123.45)
	str2 := arg2.Value()
	t.Logf("formulaArg(123.45).Value() = '%s'", str2)

	// Test with a small integer
	arg3 := newNumberFormulaArg(5)
	str3 := arg3.Value()
	t.Logf("formulaArg(5).Value() = '%s'", str3)

	if str3 != "5" {
		t.Errorf("Expected '5', got '%s'", str3)
	}
}
