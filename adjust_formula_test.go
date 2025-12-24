package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestRemoveRowWithFormulaRef(t *testing.T) {
	// Test that removing a referenced row produces #REF! error
	f := NewFile()

	// Set up data and formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", 30))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A2"))      // References A2
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "A2*2"))    // References A2
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "A3+A2"))   // References A2 and A3

	// Remove row 2 (A2 is deleted)
	assert.NoError(t, f.RemoveRow("Sheet1", 2))

	// Verify formulas now contain #REF!
	formula1, err := f.GetCellFormula("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "A#REF!", formula1, "Formula should contain #REF! when referenced row is deleted")

	formula2, err := f.GetCellFormula("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Contains(t, formula2, "#REF!", "Formula should contain #REF! when referenced row is deleted")
}

func TestRemoveRowCrossSheetRef(t *testing.T) {
	// Test that removing a referenced row in another sheet produces #REF! error
	f := NewFile()
	_, err := f.NewSheet("Data")
	assert.NoError(t, err)

	// Set up data in Data sheet
	assert.NoError(t, f.SetCellValue("Data", "A1", 10))
	assert.NoError(t, f.SetCellValue("Data", "A2", 20))
	assert.NoError(t, f.SetCellValue("Data", "A3", 30))

	// Set up formulas in Sheet1 that reference Data sheet
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "Data!A2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "Data!A2*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "Data!A3+Data!A2"))

	// Remove row 2 in Data sheet
	assert.NoError(t, f.RemoveRow("Data", 2))

	// Verify cross-sheet formulas now contain #REF!
	formula1, err := f.GetCellFormula("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Data!A#REF!", formula1)

	formula2, err := f.GetCellFormula("Sheet1", "A2")
	assert.NoError(t, err)
	assert.Equal(t, "Data!A#REF!*2", formula2)

	formula3, err := f.GetCellFormula("Sheet1", "A3")
	assert.NoError(t, err)
	assert.Contains(t, formula3, "#REF!")
}

func TestInsertRowsWithFormula(t *testing.T) {
	// Test that inserting rows correctly adjusts formulas
	f := NewFile()

	// Set up data and formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", 30))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "A2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "A3"))

	// Insert 1 row before row 2
	assert.NoError(t, f.InsertRows("Sheet1", 2, 1))

	// Verify formulas are correctly adjusted
	formula1, err := f.GetCellFormula("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "A1", formula1, "Formula before insert point should not change")

	// Row 2 is now empty (inserted row)
	formula2, err := f.GetCellFormula("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "", formula2, "Inserted row should be empty")

	// Original row 2 moved to row 3, formula should adjust from A2 to A3
	formula3, err := f.GetCellFormula("Sheet1", "B3")
	assert.NoError(t, err)
	assert.Equal(t, "A3", formula3, "Formula should adjust when row moves down")

	// Original row 3 moved to row 4, formula should adjust from A3 to A4
	formula4, err := f.GetCellFormula("Sheet1", "B4")
	assert.NoError(t, err)
	assert.Equal(t, "A4", formula4, "Formula should adjust when row moves down")
}

func TestRemoveColWithFormulaRef(t *testing.T) {
	// Test that removing a referenced column produces #REF! error
	f := NewFile()

	// Set up data and formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", 20))
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", 30))
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "B1*2"))    // References B1
	assert.NoError(t, f.SetCellFormula("Sheet1", "E1", "A1+B1"))   // References B1 and A1

	// Remove column B
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	// Verify formulas now contain #REF!
	// Column D moved to C after B was removed
	formula1, err := f.GetCellFormula("Sheet1", "C1")
	assert.NoError(t, err)
	// Note: #REF!1 is expected because the row number (1) is preserved
	assert.Contains(t, formula1, "#REF!", "Formula should contain #REF! when referenced column is deleted")

	// Column E moved to D after B was removed
	formula2, err := f.GetCellFormula("Sheet1", "D1")
	assert.NoError(t, err)
	assert.Contains(t, formula2, "#REF!", "Formula should contain #REF! when referenced column is deleted")
}
