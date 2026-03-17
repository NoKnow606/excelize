package excelize

import (
	"testing"
)

// TestIntegerCheck tests the integer detection logic
func TestIntegerCheck(t *testing.T) {
	num := 12677910539.0

	// Check if it's an integer
	isInt := num == float64(int64(num))
	t.Logf("Is %f an integer? %v", num, isInt)

	// Check the conversion
	intVal := int64(num)
	floatBack := float64(intVal)
	t.Logf("Original: %f, Int64: %d, Back to float: %f", num, intVal, floatBack)

	if num != floatBack {
		t.Errorf("Conversion mismatch!")
	}
}
