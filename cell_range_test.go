package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestGetRangeValuesConcurrent(t *testing.T) {
	f := NewFile()

	// Create test data: 100 rows × 10 columns
	for row := 1; row <= 100; row++ {
		for col := 1; col <= 10; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			value := fmt.Sprintf("R%dC%d", row, col)
			assert.NoError(t, f.SetCellValue("Sheet1", cell, value))
		}
	}

	// Test: Read range A1:J100
	values, err := f.GetRangeValuesConcurrent("Sheet1", "A1:J100")
	assert.NoError(t, err)
	assert.Equal(t, 100, len(values))
	assert.Equal(t, 10, len(values[0]))

	// Verify values
	assert.Equal(t, "R1C1", values[0][0])
	assert.Equal(t, "R1C10", values[0][9])
	assert.Equal(t, "R100C1", values[99][0])
	assert.Equal(t, "R100C10", values[99][9])

	// Test: Read partial range B2:D5
	values, err = f.GetRangeValuesConcurrent("Sheet1", "B2:D5")
	assert.NoError(t, err)
	assert.Equal(t, 4, len(values))
	assert.Equal(t, 3, len(values[0]))
	assert.Equal(t, "R2C2", values[0][0])
	assert.Equal(t, "R5C4", values[3][2])

	// Test: Invalid range
	_, err = f.GetRangeValuesConcurrent("Sheet1", "InvalidRange")
	assert.Error(t, err)

	// Test: Invalid sheet
	_, err = f.GetRangeValuesConcurrent("InvalidSheet", "A1:B2")
	assert.Error(t, err)

	// Test: Single cell
	values, err = f.GetRangeValuesConcurrent("Sheet1", "A1:A1")
	assert.NoError(t, err)
	assert.Equal(t, 1, len(values))
	assert.Equal(t, 1, len(values[0]))
	assert.Equal(t, "R1C1", values[0][0])
}

func TestGetRangeDataConcurrent(t *testing.T) {
	f := NewFile()

	// Create test data with styles and formulas
	styleID, _ := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})

	for row := 1; row <= 10; row++ {
		for col := 1; col <= 5; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			if col == 1 {
				// First column: values with style
				assert.NoError(t, f.SetCellValue("Sheet1", cell, fmt.Sprintf("R%dC%d", row, col)))
				assert.NoError(t, f.SetCellStyle("Sheet1", cell, cell, styleID))
			} else if col == 2 {
				// Second column: formulas
				assert.NoError(t, f.SetCellFormula("Sheet1", cell, fmt.Sprintf("A%d&\"_formula\"", row)))
			} else {
				// Other columns: plain values
				assert.NoError(t, f.SetCellValue("Sheet1", cell, fmt.Sprintf("R%dC%d", row, col)))
			}
		}
	}

	// Test: Read range with data
	data, err := f.GetRangeDataConcurrent("Sheet1", "A1:E10")
	assert.NoError(t, err)
	assert.Equal(t, 10, len(data))
	assert.Equal(t, 5, len(data[0]))

	// Verify first cell has value and style
	assert.Equal(t, "R1C1", data[0][0].Value)
	assert.NotNil(t, data[0][0].Style)
	assert.Equal(t, "FF0000", data[0][0].Style.Fill.Color[0])

	// Verify second column has formulas
	assert.Contains(t, data[0][1].Formula, "=A1")

	// Verify other cells have values
	assert.Equal(t, "R1C3", data[0][2].Value)
}

func TestParseRange(t *testing.T) {
	f := NewFile()

	// Test valid ranges
	startCol, startRow, endCol, endRow, err := f.parseRange("A1:Z100")
	assert.NoError(t, err)
	assert.Equal(t, 1, startCol)
	assert.Equal(t, 1, startRow)
	assert.Equal(t, 26, endCol)
	assert.Equal(t, 100, endRow)

	// Test single cell
	startCol, startRow, endCol, endRow, err = f.parseRange("B2:B2")
	assert.NoError(t, err)
	assert.Equal(t, 2, startCol)
	assert.Equal(t, 2, startRow)
	assert.Equal(t, 2, endCol)
	assert.Equal(t, 2, endRow)

	// Test invalid format
	_, _, _, _, err = f.parseRange("A1")
	assert.Error(t, err)

	_, _, _, _, err = f.parseRange("A1:InvalidCell")
	assert.Error(t, err)

	// Test reversed range
	_, _, _, _, err = f.parseRange("Z100:A1")
	assert.Error(t, err)
}

func BenchmarkGetRangeValuesConcurrent(b *testing.B) {
	f := NewFile()

	// Create test data: 1000 rows × 100 columns
	for row := 1; row <= 1000; row++ {
		for col := 1; col <= 100; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			_ = f.SetCellValue("Sheet1", cell, fmt.Sprintf("R%dC%d", row, col))
		}
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, _ = f.GetRangeValuesConcurrent("Sheet1", "A1:CV1000")
	}
}

func BenchmarkGetRows(b *testing.B) {
	f := NewFile()

	// Create test data: 1000 rows × 100 columns
	for row := 1; row <= 1000; row++ {
		for col := 1; col <= 100; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			_ = f.SetCellValue("Sheet1", cell, fmt.Sprintf("R%dC%d", row, col))
		}
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, _ = f.GetRows("Sheet1")
	}
}

func BenchmarkGetCellValueLoop(b *testing.B) {
	f := NewFile()

	// Create test data: 1000 rows × 100 columns
	for row := 1; row <= 1000; row++ {
		for col := 1; col <= 100; col++ {
			cell, _ := CoordinatesToCellName(col, row)
			_ = f.SetCellValue("Sheet1", cell, fmt.Sprintf("R%dC%d", row, col))
		}
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for row := 1; row <= 1000; row++ {
			for col := 1; col <= 100; col++ {
				cell, _ := CoordinatesToCellName(col, row)
				_, _ = f.GetCellValue("Sheet1", cell)
			}
		}
	}
}
