package excelize

import "testing"

func TestFormulaArgToFormulaLiteral(t *testing.T) {
	tests := []struct {
		name string
		arg  formulaArg
		want string
	}{
		{name: "number", arg: newNumberFormulaArg(200), want: "200"},
		{name: "string", arg: newStringFormulaArg("200"), want: `"200"`},
		{name: "bool", arg: newBoolFormulaArg(true), want: "TRUE"},
		{name: "error", arg: newErrorFormulaArg(formulaErrorNA, formulaErrorNA), want: formulaErrorNA},
		{name: "empty", arg: newEmptyFormulaArg(), want: `""`},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			if got := formulaArgToFormulaLiteral(tc.arg); got != tc.want {
				t.Fatalf("formulaArgToFormulaLiteral() = %q, want %q", got, tc.want)
			}
		})
	}
}

func TestCalcCellValueWithSubExprCachePreservesLookupTypes(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	formula := `IF(TYPE(VLOOKUP(A1,Data!A:B,2,FALSE))=2,"string","number")`
	lookupExpr := `VLOOKUP(A1,Data!A:B,2,FALSE)`

	subExprCache := NewSubExpressionCache()
	subExprCache.Store(lookupExpr, newNumberFormulaArg(200))

	got, err := f.CalcCellValueWithSubExprCache("Sheet1", "B1", formula, subExprCache, NewWorksheetCache(), Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("CalcCellValueWithSubExprCache() number case: %v", err)
	}
	if got != "number" {
		t.Fatalf("numeric lookup should remain numeric after replacement, got %q", got)
	}

	subExprCache.Clear()
	subExprCache.Store(lookupExpr, newStringFormulaArg("200"))

	got, err = f.CalcCellValueWithSubExprCache("Sheet1", "B1", formula, subExprCache, NewWorksheetCache(), Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("CalcCellValueWithSubExprCache() string case: %v", err)
	}
	if got != "string" {
		t.Fatalf("string lookup should remain text after replacement, got %q", got)
	}
}
