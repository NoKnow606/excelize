package excelize

import (
	"fmt"
	"testing"
)

// TestNumberFormatting tests how %g formats large numbers
func TestNumberFormatting(t *testing.T) {
	num := 12677910539.0
	formatted := fmt.Sprintf("%g", num)
	t.Logf("Number: %f, Formatted with %%g: '%s'", num, formatted)

	// Check if it matches the string
	str := "12677910539"
	if formatted != str {
		t.Logf("Mismatch! '%s' != '%s'", formatted, str)
	} else {
		t.Logf("Match! '%s' == '%s'", formatted, str)
	}
}
