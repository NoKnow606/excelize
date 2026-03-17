package excelize

import (
	"testing"
)

// TestCellTypeForNumericString checks cell type for numeric strings
func TestCellTypeForNumericString(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Set a numeric string
	f.SetCellValue("Sheet1", "A1", "12677910539")

	// Check its type
	cellType, _ := f.GetCellType("Sheet1", "A1")
	t.Logf("Cell type for '12677910539': %d (%s)", cellType, cellTypeToString(cellType))

	// Set explicitly as string
	f.SetCellStr("Sheet1", "A2", "12677910539")
	cellType2, _ := f.GetCellType("Sheet1", "A2")
	t.Logf("Cell type for SetCellStr: %d (%s)", cellType2, cellTypeToString(cellType2))

	// Set a non-numeric string
	f.SetCellValue("Sheet1", "A3", "-")
	cellType3, _ := f.GetCellType("Sheet1", "A3")
	t.Logf("Cell type for '-': %d (%s)", cellType3, cellTypeToString(cellType3))
}

func cellTypeToString(ct CellType) string {
	switch ct {
	case CellTypeBool:
		return "Bool"
	case CellTypeNumber:
		return "Number"
	case CellTypeInlineString:
		return "InlineString"
	case CellTypeSharedString:
		return "SharedString"
	case CellTypeFormula:
		return "Formula"
	case CellTypeDate:
		return "Date"
	case CellTypeUnset:
		return "Unset"
	default:
		return "Unknown"
	}
}
