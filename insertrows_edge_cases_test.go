package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestInsertRowBatchUpdateEdgeCases tests edge cases that might cause issues
func TestInsertRowBatchUpdateEdgeCases(t *testing.T) {
	t.Run("Insert at row 1", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		// Insert at the very beginning
		err := f.InsertRows("Sheet1", 1, 1)
		assert.NoError(t, err)

		// Batch update
		updates := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: "First"},
			{Sheet: "Sheet1", Cell: "B1", Value: 100},
		}
		_, err = f.BatchUpdateAndRecalculate(updates)
		assert.NoError(t, err)

		// Get values
		a1, _ := f.GetCellValue("Sheet1", "A1")
		b1, _ := f.GetCellValue("Sheet1", "B1")

		fmt.Printf("Row 1: A=%s, B=%s\n", a1, b1)
		assert.Equal(t, "First", a1)
		assert.Equal(t, "100", b1)
	})

	t.Run("Insert multiple rows then update", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		// Insert 3 rows at once
		err := f.InsertRows("Sheet1", 1, 3)
		assert.NoError(t, err)

		// Batch update all 3 rows
		updates := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: "R1"},
			{Sheet: "Sheet1", Cell: "A2", Value: "R2"},
			{Sheet: "Sheet1", Cell: "A3", Value: "R3"},
		}
		_, err = f.BatchUpdateAndRecalculate(updates)
		assert.NoError(t, err)

		// Verify all 3 rows
		for i := 1; i <= 3; i++ {
			val, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
			expected := fmt.Sprintf("R%d", i)
			fmt.Printf("A%d = %s (expected: %s)\n", i, val, expected)
			assert.Equal(t, expected, val)
		}
	})

	t.Run("Insert then update with GetRows", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", "Original")

		// Insert row
		err := f.InsertRows("Sheet1", 2, 1)
		assert.NoError(t, err)

		// Batch update
		updates := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A2", Value: "Inserted"},
			{Sheet: "Sheet1", Cell: "B2", Value: 123},
		}
		_, err = f.BatchUpdateAndRecalculate(updates)
		assert.NoError(t, err)

		// Get via GetRows (different code path)
		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)

		fmt.Printf("Total rows: %d\n", len(rows))
		for i, row := range rows {
			fmt.Printf("Row %d: %v\n", i+1, row)
		}

		assert.GreaterOrEqual(t, len(rows), 2)
		if len(rows) >= 2 {
			assert.Equal(t, "Inserted", rows[1][0])
			assert.Equal(t, "123", rows[1][1])
		}
	})

	t.Run("Insert row and immediately batch update without recalc", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		// Insert row
		err := f.InsertRows("Sheet1", 1, 1)
		assert.NoError(t, err)

		// Use BatchSetCellValue (no recalc) instead
		updates := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: "Test1"},
			{Sheet: "Sheet1", Cell: "B1", Value: "Test2"},
		}
		err = f.BatchSetCellValue(updates)
		assert.NoError(t, err)

		// Get values
		a1, _ := f.GetCellValue("Sheet1", "A1")
		b1, _ := f.GetCellValue("Sheet1", "B1")

		fmt.Printf("A1=%s, B1=%s\n", a1, b1)
		assert.Equal(t, "Test1", a1)
		assert.Equal(t, "Test2", b1)
	})

	t.Run("Insert row in middle of large dataset", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		// Create 100 rows of data
		for i := 1; i <= 100; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), fmt.Sprintf("Data%d", i))
		}

		// Insert in the middle
		err := f.InsertRows("Sheet1", 50, 1)
		assert.NoError(t, err)

		// Batch update the inserted row
		updates := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A50", Value: "INSERTED"},
			{Sheet: "Sheet1", Cell: "B50", Value: "NEW"},
		}
		_, err = f.BatchUpdateAndRecalculate(updates)
		assert.NoError(t, err)

		// Verify
		a50, _ := f.GetCellValue("Sheet1", "A50")
		b50, _ := f.GetCellValue("Sheet1", "B50")
		a51, _ := f.GetCellValue("Sheet1", "A51")

		fmt.Printf("A50=%s, B50=%s, A51=%s\n", a50, b50, a51)
		assert.Equal(t, "INSERTED", a50)
		assert.Equal(t, "NEW", b50)
		assert.Equal(t, "Data50", a51, "Original row 50 should move to row 51")
	})
}

// TestInsertRowWithWorksheetNotLoaded tests when worksheet is not in memory
func TestInsertRowWithWorksheetNotLoaded(t *testing.T) {
	// Create and save a file
	tmpFile := "test_unloaded.xlsx"
	defer func() {
		f, _ := OpenFile(tmpFile)
		if f != nil {
			f.Close()
		}
	}()

	// Create initial file
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", "Header")
	f.SetCellValue("Sheet1", "A2", "Data1")
	f.SaveAs(tmpFile)
	f.Close()

	// Reopen file (worksheet not loaded yet)
	f2, err := OpenFile(tmpFile)
	assert.NoError(t, err)
	defer f2.Close()

	fmt.Println("\n=== 工作表未加载状态 ===")

	// Insert row (this will load the worksheet)
	err = f2.InsertRows("Sheet1", 2, 1)
	assert.NoError(t, err)

	// Batch update
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: "InsertedInUnloaded"},
		{Sheet: "Sheet1", Cell: "B2", Value: 999},
	}
	_, err = f2.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Get values
	a2, _ := f2.GetCellValue("Sheet1", "A2")
	b2, _ := f2.GetCellValue("Sheet1", "B2")
	a3, _ := f2.GetCellValue("Sheet1", "A3")

	fmt.Printf("A2=%s, B2=%s, A3=%s\n", a2, b2, a3)
	assert.Equal(t, "InsertedInUnloaded", a2)
	assert.Equal(t, "999", b2)
	assert.Equal(t, "Data1", a3)
}

// TestInsertRowBatchUpdateCellReference tests if cell references are correct
func TestInsertRowBatchUpdateCellReference(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup
	f.SetCellValue("Sheet1", "A1", "H1")
	f.SetCellValue("Sheet1", "A2", "D1")

	// Insert row at 2
	err := f.InsertRows("Sheet1", 2, 1)
	assert.NoError(t, err)

	// Update using wrong row number (common mistake)
	wrongUpdates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: "Wrong"},  // Updating row 1 instead of 2
	}
	_, err = f.BatchUpdateAndRecalculate(wrongUpdates)
	assert.NoError(t, err)

	// Correct update
	correctUpdates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A2", Value: "Correct"},  // Row 2 is the inserted row
	}
	_, err = f.BatchUpdateAndRecalculate(correctUpdates)
	assert.NoError(t, err)

	a1, _ := f.GetCellValue("Sheet1", "A1")
	a2, _ := f.GetCellValue("Sheet1", "A2")
	a3, _ := f.GetCellValue("Sheet1", "A3")

	fmt.Printf("A1=%s, A2=%s, A3=%s\n", a1, a2, a3)
	assert.Equal(t, "Wrong", a1)   // We updated this
	assert.Equal(t, "Correct", a2) // We updated this
	assert.Equal(t, "D1", a3)      // Original data moved here
}

// TestInsertRowBatchUpdateEmptySheet tests with completely empty sheet
func TestInsertRowBatchUpdateEmptySheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Don't set any initial data
	// Insert row in empty sheet
	err := f.InsertRows("Sheet1", 1, 1)
	assert.NoError(t, err)

	// Batch update
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: "FirstEver"},
		{Sheet: "Sheet1", Cell: "B1", Value: 123},
	}
	_, err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Get values
	a1, _ := f.GetCellValue("Sheet1", "A1")
	b1, _ := f.GetCellValue("Sheet1", "B1")

	fmt.Printf("Empty sheet: A1=%s, B1=%s\n", a1, b1)
	assert.Equal(t, "FirstEver", a1)
	assert.Equal(t, "123", b1)
}
