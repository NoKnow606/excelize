package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestMoveRowSameSheet(t *testing.T) {
	// Test moving a row within the same sheet
	f := NewFile()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Row1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", "Row2-ToMove"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", "Row3"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A4", "Row4"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A5", "Row5"))

	// Set up formulas that reference row 2
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "A2*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "A2+A3"))

	// Move row 2 to row 4
	assert.NoError(t, f.MoveRow("Sheet1", 2, 4))

	// Verify data moved correctly
	val1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "Row1", val1)
	val2, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Row3", val2) // Row3 shifted up
	val3, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Row4", val3) // Row4 shifted up
	val4, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "Row2-ToMove", val4) // Original row2 is now at row4
	val5, _ := f.GetCellValue("Sheet1", "A5")
	assert.Equal(t, "Row5", val5)

	// Verify formulas updated correctly
	formula1, _ := f.GetCellFormula("Sheet1", "B1")
	assert.Equal(t, "A4", formula1, "Formula should reference moved row")
	// B2 now contains the formula that was in B3 (shifted up)
	// Original B3 had "A2+A3", after move: A2->A4, A3->A2 (shifted up with the row)
	formula2, _ := f.GetCellFormula("Sheet1", "B2")
	assert.Equal(t, "A4+A2", formula2, "Formula should shift up and update references")
	// B3 now contains the formula that was in B4 (empty), or might be empty
	formula3, _ := f.GetCellFormula("Sheet1", "B3")
	assert.Equal(t, "", formula3, "Formula should be empty (was in B4)")
	// B4 now contains the formula that was in B2, which referenced its own row A2
	formula4, _ := f.GetCellFormula("Sheet1", "B4")
	assert.Contains(t, formula4, "A4", "Formula should reference moved row")
}

func TestMoveRowCrossSheet(t *testing.T) {
	// Test that cross-sheet references update when moving rows
	f := NewFile()
	_, err := f.NewSheet("Data")
	assert.NoError(t, err)

	// Set up data in Data sheet
	assert.NoError(t, f.SetCellValue("Data", "A1", 10))
	assert.NoError(t, f.SetCellValue("Data", "A2", 20))
	assert.NoError(t, f.SetCellValue("Data", "A3", 30))
	assert.NoError(t, f.SetCellValue("Data", "A4", 40))

	// Set up formulas in Sheet1 that reference Data sheet
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "Data!A2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "Data!A2*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "Data!A2+Data!A3"))

	// Move row 2 in Data sheet to row 4
	assert.NoError(t, f.MoveRow("Data", 2, 4))

	// Verify cross-sheet formulas updated correctly
	formula1, _ := f.GetCellFormula("Sheet1", "A1")
	assert.Equal(t, "Data!A4", formula1, "Cross-sheet formula should reference moved row")
	formula2, _ := f.GetCellFormula("Sheet1", "A2")
	assert.Equal(t, "Data!A4*2", formula2, "Cross-sheet formula should reference moved row")
	// Original formula in A3 was "Data!A2+Data!A3", after move:
	// Data!A2->Data!A4, Data!A3->Data!A2 (row 3 shifted up)
	formula3, _ := f.GetCellFormula("Sheet1", "A3")
	assert.Contains(t, formula3, "Data!A4", "Cross-sheet formula should reference moved row")
	assert.Contains(t, formula3, "Data!A2", "Other references should shift accordingly")
}

func TestMoveRowUp(t *testing.T) {
	// Test moving a row upward
	f := NewFile()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Row1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", "Row2"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", "Row3"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A4", "Row4-ToMove"))

	// Set up formulas that reference row 4
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A4"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "A4*2"))

	// Move row 4 to row 2
	assert.NoError(t, f.MoveRow("Sheet1", 4, 2))

	// Verify data moved correctly
	val1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "Row1", val1)
	val2, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Row4-ToMove", val2) // Row4 is now at row2
	val3, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "Row2", val3) // Row2 shifted down
	val4, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "Row3", val4) // Row3 shifted down

	// Verify formulas updated correctly
	formula1, _ := f.GetCellFormula("Sheet1", "B1")
	assert.Equal(t, "A2", formula1, "Formula should reference moved row")
	// B2 is now empty - the row that was here moved to B3
	// The row that moved from row 4 to row 2 didn't have a formula in column B
	formula2, _ := f.GetCellFormula("Sheet1", "B2")
	assert.Equal(t, "", formula2, "Formula should be empty (moved row had no formula)")
	// B3 has the formula that was in B2, which was "A4*2", now "A2*2"
	formula3, _ := f.GetCellFormula("Sheet1", "B3")
	assert.Equal(t, "A2*2", formula3, "Formula should shift down and update references")
}

func TestMoveColSameSheet(t *testing.T) {
	// Test moving a column within the same sheet
	f := NewFile()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "ColA"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "ColB-ToMove"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", "ColC"))
	assert.NoError(t, f.SetCellValue("Sheet1", "D1", "ColD"))
	assert.NoError(t, f.SetCellValue("Sheet1", "E1", "ColE"))

	// Set up formulas that reference column B
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "B1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "B1&\"-formula\""))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C2", "B1&C1"))

	// Move column B to column D
	assert.NoError(t, f.MoveCol("Sheet1", "B", "D"))

	// Verify data moved correctly
	valA, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "ColA", valA)
	valB, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "ColC", valB) // ColC shifted left
	valC, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "ColD", valC) // ColD shifted left
	valD, _ := f.GetCellValue("Sheet1", "D1")
	assert.Equal(t, "ColB-ToMove", valD) // Original ColB is now at ColD
	valE, _ := f.GetCellValue("Sheet1", "E1")
	assert.Equal(t, "ColE", valE)

	// Verify formulas updated correctly
	formula1, _ := f.GetCellFormula("Sheet1", "A2")
	assert.Equal(t, "D1", formula1, "Formula should reference moved column")
	// B2 formula was originally in C2 (shifted left), which was "B1&C1"
	// After the move: B1->D1, C1->B1 (column C shifted left to B)
	formula2, _ := f.GetCellFormula("Sheet1", "B2")
	assert.Contains(t, formula2, "D1", "Formula should reference moved column")
	assert.Contains(t, formula2, "B1", "Formula should shift with column")
	// D2 has the formula that was in B2
	formula4, _ := f.GetCellFormula("Sheet1", "D2")
	assert.Contains(t, formula4, "D1", "Formula should reference moved column")
}

func TestMoveColCrossSheet(t *testing.T) {
	// Test that cross-sheet references update when moving columns
	f := NewFile()
	_, err := f.NewSheet("Data")
	assert.NoError(t, err)

	// Set up data in Data sheet
	assert.NoError(t, f.SetCellValue("Data", "A1", 10))
	assert.NoError(t, f.SetCellValue("Data", "B1", 20))
	assert.NoError(t, f.SetCellValue("Data", "C1", 30))
	assert.NoError(t, f.SetCellValue("Data", "D1", 40))

	// Set up formulas in Sheet1 that reference Data sheet
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "Data!B1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "Data!B1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "Data!B1+Data!C1"))

	// Move column B in Data sheet to column D
	assert.NoError(t, f.MoveCol("Data", "B", "D"))

	// Verify cross-sheet formulas updated correctly
	formula1, _ := f.GetCellFormula("Sheet1", "A1")
	assert.Equal(t, "Data!D1", formula1, "Cross-sheet formula should reference moved column")
	formula2, _ := f.GetCellFormula("Sheet1", "A2")
	assert.Equal(t, "Data!D1*2", formula2, "Cross-sheet formula should reference moved column")
	formula3, _ := f.GetCellFormula("Sheet1", "A3")
	assert.Contains(t, formula3, "Data!D1", "Cross-sheet formula should reference moved column")
}

func TestMoveColLeft(t *testing.T) {
	// Test moving a column leftward
	f := NewFile()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "ColA"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "ColB"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", "ColC"))
	assert.NoError(t, f.SetCellValue("Sheet1", "D1", "ColD-ToMove"))

	// Set up formulas that reference column D
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "D1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "D1*2"))

	// Move column D to column B
	assert.NoError(t, f.MoveCol("Sheet1", "D", "B"))

	// Verify data moved correctly
	valA, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "ColA", valA)
	valB, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "ColD-ToMove", valB) // ColD is now at ColB
	valC, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "ColB", valC) // ColB shifted right
	valD, _ := f.GetCellValue("Sheet1", "D1")
	assert.Equal(t, "ColC", valD) // ColC shifted right

	// Verify formulas updated correctly
	formula1, _ := f.GetCellFormula("Sheet1", "A2")
	assert.Equal(t, "B1", formula1, "Formula should reference moved column")
	// B2 had formula "D1*2", after D moved to B, this becomes at C2
	// And the reference D1 becomes B1
	formula2, _ := f.GetCellFormula("Sheet1", "B2")
	// The formula that was in B2 ("D1*2") is now at C2
	// B2 should have the formula from A2, which was "D1" (now "B1")
	// But wait, column A doesn't move, so A2's formula stays at A2
	// So B2 should have whatever was in A2... but A2 doesn't shift
	// Let me reconsider: Move D to B means:
	// - Column D -> B
	// - Column B -> C
	// - Column C -> D
	// So B2 formula (which was "D1*2") should now be at C2 and reference B1
	formula2, _ = f.GetCellFormula("Sheet1", "C2")
	assert.Contains(t, formula2, "B1", "Formula should shift right with its column")
}

func TestMoveRowWithFormulasInMovedRow(t *testing.T) {
	// Test that formulas inside the moved row are preserved
	f := NewFile()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "A2*2")) // Formula in row to be moved

	// Move row 2 to row 4
	assert.NoError(t, f.MoveRow("Sheet1", 2, 4))

	// Verify formula in moved row is preserved and working
	formula, _ := f.GetCellFormula("Sheet1", "B4")
	assert.NotEmpty(t, formula, "Formula should be preserved in moved row")

	// Verify the data
	val, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "20", val)
}

func TestMoveRowSamePosition(t *testing.T) {
	// Test moving a row to its own position (should be a no-op)
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", "Test"))

	assert.NoError(t, f.MoveRow("Sheet1", 2, 2))

	val, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Test", val, "Row should remain unchanged")
}

func TestMoveColSamePosition(t *testing.T) {
	// Test moving a column to its own position (should be a no-op)
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "Test"))

	assert.NoError(t, f.MoveCol("Sheet1", "B", "B"))

	val, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "Test", val, "Column should remain unchanged")
}

func TestMoveRowWithAbsoluteReferences(t *testing.T) {
	// Test that absolute references are updated correctly
	f := NewFile()

	// Set up data with absolute reference
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 100))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "$A$2"))

	// Move row 2 to row 4
	assert.NoError(t, f.MoveRow("Sheet1", 2, 4))

	// Verify absolute reference updated
	formula, _ := f.GetCellFormula("Sheet1", "B1")
	assert.Equal(t, "$A$4", formula, "Absolute reference should update to moved row")
}

func TestMoveColWithAbsoluteReferences(t *testing.T) {
	// Test that absolute references are updated correctly
	f := NewFile()

	// Set up data with absolute reference
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", 100))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "$B$1"))

	// Move column B to column D
	assert.NoError(t, f.MoveCol("Sheet1", "B", "D"))

	// Verify absolute reference updated
	formula, _ := f.GetCellFormula("Sheet1", "A2")
	assert.Equal(t, "$D$1", formula, "Absolute reference should update to moved column")
}
