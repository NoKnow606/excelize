package excelize

import "strconv"

const (
	pgMirrorValueTypeBlank  = "blank"
	pgMirrorValueTypeNumber = "number"
	pgMirrorValueTypeBool   = "bool"
	pgMirrorValueTypeError  = "error"
	pgMirrorValueTypeString = "string"
)

type pgMirrorCellValue struct {
	Value     string
	ValueType string
}

func pgMirrorValueTypeFromCell(value string, cellType CellType) string {
	if value == "" {
		return pgMirrorValueTypeBlank
	}
	switch cellType {
	case CellTypeBool:
		return pgMirrorValueTypeBool
	case CellTypeError:
		return pgMirrorValueTypeError
	case CellTypeInlineString, CellTypeSharedString, CellTypeFormula:
		return pgMirrorValueTypeString
	case CellTypeNumber:
		return pgMirrorValueTypeNumber
	case CellTypeUnset:
		if _, err := strconv.ParseFloat(value, 64); err == nil {
			return pgMirrorValueTypeNumber
		}
		return pgMirrorValueTypeString
	default:
		return pgMirrorValueTypeString
	}
}

func pgMirrorValueTypeFromFormulaArg(arg formulaArg) string {
	switch arg.Type {
	case ArgNumber:
		if arg.Boolean {
			return pgMirrorValueTypeBool
		}
		return pgMirrorValueTypeNumber
	case ArgError:
		return pgMirrorValueTypeError
	case ArgEmpty:
		return pgMirrorValueTypeBlank
	default:
		return pgMirrorValueTypeString
	}
}

func (v pgMirrorCellValue) formulaArg() formulaArg {
	switch v.ValueType {
	case pgMirrorValueTypeBlank:
		return newEmptyFormulaArg()
	case pgMirrorValueTypeBool:
		return newBoolFormulaArg(v.Value == "1" || v.Value == "TRUE" || v.Value == "true")
	case pgMirrorValueTypeError:
		return newErrorFormulaArg(v.Value, v.Value)
	case pgMirrorValueTypeNumber:
		if num, err := strconv.ParseFloat(v.Value, 64); err == nil {
			return newNumberFormulaArg(num)
		}
		return newStringFormulaArg(v.Value)
	default:
		return newStringFormulaArg(v.Value)
	}
}

func formulaArgTypedKey(arg formulaArg) string {
	return pgMirrorValueTypeFromFormulaArg(arg) + "\x00" + pgFormulaArgValue(arg)
}

func pgFormulaArgTypedKey(arg formulaArg) string {
	return formulaArgTypedKey(arg)
}
