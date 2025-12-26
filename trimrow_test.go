package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestTrimRowWithMixedEmptyRows tests trimRow with mixed empty and non-empty rows
func TestTrimRowWithMixedEmptyRows(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create a pattern: data, empty, data, empty, empty, data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Row1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", "Row3"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A6", "Row6"))

	// Write and verify no panic
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})
}

// TestTrimRowWithAllEmptyRows tests trimRow with all empty rows
func TestTrimRowWithAllEmptyRows(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create empty rows by setting and then clearing
	for i := 1; i <= 10; i++ {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), "temp"))
	}

	// Clear all cells (creating empty rows with attributes)
	for i := 1; i <= 10; i++ {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), ""))
	}

	// Write should not panic
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})
}

// TestTrimRowWithLargeGaps tests trimRow with large gaps between data
func TestTrimRowWithLargeGaps(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create sparse data with large gaps
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "First"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A100", "Middle"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A1000", "Last"))

	// Write should not panic
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})

	// Verify data integrity
	val1, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "First", val1)

	val100, err := f.GetCellValue("Sheet1", "A100")
	assert.NoError(t, err)
	assert.Equal(t, "Middle", val100)

	val1000, err := f.GetCellValue("Sheet1", "A1000")
	assert.NoError(t, err)
	assert.Equal(t, "Last", val1000)
}

// TestTrimRowWithAlternatingPattern tests alternating empty/non-empty rows
func TestTrimRowWithAlternatingPattern(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create alternating pattern: data, empty, data, empty...
	for i := 1; i <= 100; i += 2 {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), fmt.Sprintf("Data%d", i)))
	}

	// Write should not panic
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})

	// Verify data integrity
	for i := 1; i <= 100; i += 2 {
		val, err := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		assert.NoError(t, err)
		assert.Equal(t, fmt.Sprintf("Data%d", i), val)
	}
}

// TestTrimRowMultipleWrites tests trimRow across multiple write operations
func TestTrimRowMultipleWrites(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// First write cycle
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Data1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A10", "Data10"))

	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})

	// Modify data
	assert.NoError(t, f.SetCellValue("Sheet1", "A5", "Data5"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", ""))  // Clear A1

	// Second write cycle
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})

	// Verify final state
	val5, _ := f.GetCellValue("Sheet1", "A5")
	assert.Equal(t, "Data5", val5)

	val10, _ := f.GetCellValue("Sheet1", "A10")
	assert.Equal(t, "Data10", val10)
}

// TestTrimRowWithBatchOperations tests trimRow with batch updates
func TestTrimRowWithBatchOperations(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Batch update with gaps
	updates := make([]CellUpdate, 0, 100)
	for i := 1; i <= 100; i += 3 {
		updates = append(updates, CellUpdate{
			Sheet: "Sheet1",
			Cell:  fmt.Sprintf("A%d", i),
			Value: fmt.Sprintf("Batch%d", i),
		})
	}

	assert.NoError(t, f.BatchSetCellValue(updates))

	// Write should not panic
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})

	// Verify data
	for i := 1; i <= 100; i += 3 {
		val, err := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		assert.NoError(t, err)
		assert.Equal(t, fmt.Sprintf("Batch%d", i), val)
	}
}

// TestTrimRowEdgeCases tests various edge cases
func TestTrimRowEdgeCases(t *testing.T) {
	t.Run("EmptyWorksheet", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		assert.NotPanics(t, func() {
			buf, err := f.WriteToBuffer()
			assert.NoError(t, err)
			assert.NotNil(t, buf)
		})
	})

	t.Run("SingleCell", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Single"))
		assert.NotPanics(t, func() {
			buf, err := f.WriteToBuffer()
			assert.NoError(t, err)
			assert.NotNil(t, buf)
		})
	})

	t.Run("LastRowOnly", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		assert.NoError(t, f.SetCellValue("Sheet1", "A1000", "Last"))
		assert.NotPanics(t, func() {
			buf, err := f.WriteToBuffer()
			assert.NoError(t, err)
			assert.NotNil(t, buf)
		})
	})
}
