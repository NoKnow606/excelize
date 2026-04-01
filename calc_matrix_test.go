package excelize

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestCalcCellValueWithMatrix_TestWriteRange(t *testing.T) {
	f := NewFile()
	sheet := f.GetSheetName(f.GetActiveSheetIndex())
	require.NotEmpty(t, sheet)

	require.NoError(t, f.SetCellFormula(sheet, "A1", "TEST_WRITE_RANGE(2,3)"))

	result, err := f.CalcCellValueWithMatrix(sheet, "A1")
	require.NoError(t, err)
	require.Len(t, result.Matrix, 2)
	require.Len(t, result.Matrix[0], 3)
	require.Len(t, result.Matrix[1], 3)

	expected := 1.0
	for r := 0; r < 2; r++ {
		for c := 0; c < 3; c++ {
			val, ok := result.Matrix[r][c].(float64)
			require.True(t, ok, "expected numeric value at %d,%d", r, c)
			require.Equal(t, expected, val)
			expected++
		}
	}
}
