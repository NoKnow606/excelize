package excelize

import (
	"fmt"
	"testing"
)

// TestLargeNumberPrecision tests precision for large numbers
func TestLargeNumberPrecision(t *testing.T) {
	num := 12677910539.0

	// Check precision
	int64Val := int64(num)
	backToFloat := float64(int64Val)

	t.Logf("Original float: %.20f", num)
	t.Logf("As int64: %d", int64Val)
	t.Logf("Back to float: %.20f", backToFloat)
	t.Logf("Are they equal? %v", num == backToFloat)

	// Format it
	formatted := fmt.Sprintf("%.0f", num)
	t.Logf("Formatted: '%s'", formatted)

	// Try with modulo
	remainder := num - float64(int64(num))
	t.Logf("Remainder: %.20f", remainder)
}
