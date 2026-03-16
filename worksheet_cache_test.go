package excelize

import (
	"testing"
)

func TestWorksheetCacheBasicOperations(t *testing.T) {
	wc := NewWorksheetCache()
	if wc.Len() != 0 {
		t.Fatalf("expected empty cache, got %d", wc.Len())
	}

	numArg := newNumberFormulaArg(42)
	strArg := newStringFormulaArg("hello")

	wc.Set("SheetA", "A1", numArg)
	wc.Set("SheetA", "B1", strArg)

	if got, ok := wc.Get("SheetA", "A1"); !ok || got.Number != 42 || got.Type != ArgNumber {
		t.Fatalf("Get number: ok=%v type=%v value=%v", ok, got.Type, got.Number)
	}
	if got, ok := wc.Get("SheetA", "B1"); !ok || got.String != "hello" || got.Type != ArgString {
		t.Fatalf("Get string: ok=%v type=%v value=%q", ok, got.Type, got.String)
	}

	if wc.SheetLen("SheetA") != 2 || wc.Len() != 2 {
		t.Fatalf("SheetLen/Len mismatch, got sheet=%d total=%d", wc.SheetLen("SheetA"), wc.Len())
	}

	copied := wc.GetSheet("SheetA")
	copied["A1"] = newStringFormulaArg("mutated")
	if got, _ := wc.Get("SheetA", "A1"); got.Number != 42 {
		t.Fatalf("cache mutated through copy")
	}

	wc.ClearSheet("SheetA")
	if wc.SheetLen("SheetA") != 0 {
		t.Fatalf("expected cleared sheet cache")
	}

	wc.Set("SheetB", "C3", numArg)
	wc.Clear()
	if wc.Len() != 0 {
		t.Fatalf("expected cache cleared")
	}
}

func TestWorksheetCacheLoadSheet(t *testing.T) {
	f := NewFile()
	const sheet = "Sheet1"
	must := func(err error) {
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
	}

	must(f.SetCellValue(sheet, "A1", 123.5))
	must(f.SetCellValue(sheet, "B1", "text"))
	must(f.SetCellBool(sheet, "C1", true))
	must(f.SetCellValue(sheet, "E1", ""))
	must(f.SetCellFormula(sheet, "D1", "=A1"))

	wc := NewWorksheetCache()
	if err := wc.LoadSheet(f, sheet); err != nil {
		t.Fatalf("LoadSheet failed: %v", err)
	}

	if wc.Len() != 4 {
		t.Fatalf("expected 4 cached cells, got %d", wc.Len())
	}

	if arg, ok := wc.Get(sheet, "A1"); !ok || arg.Type != ArgNumber || arg.Number != 123.5 {
		t.Fatalf("unexpected A1 cache: ok=%v type=%v num=%v", ok, arg.Type, arg.Number)
	}
	if arg, ok := wc.Get(sheet, "B1"); !ok || arg.Type != ArgString || arg.String != "text" {
		t.Fatalf("unexpected B1 cache: ok=%v type=%v str=%q", ok, arg.Type, arg.String)
	}
	if arg, ok := wc.Get(sheet, "C1"); !ok || arg.Type != ArgNumber || !arg.Boolean || arg.Number != 1 {
		t.Fatalf("unexpected C1 cache: ok=%v type=%v bool=%v num=%v", ok, arg.Type, arg.Boolean, arg.Number)
	}
	if arg, ok := wc.Get(sheet, "E1"); !ok || arg.Type != ArgString || arg.String != "" {
		t.Fatalf("unexpected E1 cache: ok=%v type=%v str=%q", ok, arg.Type, arg.String)
	}
	if _, ok := wc.Get(sheet, "D1"); ok {
		t.Fatalf("formula cell D1 should not be cached")
	}
}

func TestWorksheetCacheLoadSheetIdempotentAfterClear(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	const sheet = "Sheet1"
	if err := f.SetCellValue(sheet, "A1", 1); err != nil {
		t.Fatalf("SetCellValue A1: %v", err)
	}

	wc := NewWorksheetCache()
	if err := wc.LoadSheet(f, sheet); err != nil {
		t.Fatalf("LoadSheet first: %v", err)
	}
	if got := wc.SheetLen(sheet); got != 1 {
		t.Fatalf("expected 1 cached cell after first load, got %d", got)
	}

	if err := f.SetCellValue(sheet, "A2", 2); err != nil {
		t.Fatalf("SetCellValue A2: %v", err)
	}
	if err := wc.LoadSheet(f, sheet); err != nil {
		t.Fatalf("LoadSheet second: %v", err)
	}
	if got := wc.SheetLen(sheet); got != 1 {
		t.Fatalf("second load should be a no-op while sheet is marked loaded, got %d cells", got)
	}

	wc.ClearSheet(sheet)
	if err := wc.LoadSheet(f, sheet); err != nil {
		t.Fatalf("LoadSheet after clear: %v", err)
	}
	if got := wc.SheetLen(sheet); got != 2 {
		t.Fatalf("expected 2 cached cells after reloading cleared sheet, got %d", got)
	}
}

func TestInferCellValueType(t *testing.T) {
	tests := []struct {
		name     string
		val      string
		cellType CellType
		wantType ArgType
		wantNum  float64
		wantBool bool
		wantStr  string
	}{
		{"empty string", "", CellTypeInlineString, ArgString, 0, false, ""},
		{"boolean true", "TRUE", CellTypeBool, ArgNumber, 1, true, ""},
		{"numeric", "42.5", CellTypeNumber, ArgNumber, 42.5, false, ""},
		{"numeric parse error fallback", "abc", CellTypeNumber, ArgString, 0, false, "abc"},
		{"plain string", "SKU", CellTypeSharedString, ArgString, 0, false, "SKU"},
	}

	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			arg := inferCellValueType(tc.val, tc.cellType)
			if arg.Type != tc.wantType {
				t.Fatalf("type mismatch, want %v got %v", tc.wantType, arg.Type)
			}
			if tc.wantType == ArgNumber && arg.Number != tc.wantNum {
				t.Fatalf("number mismatch, want %v got %v", tc.wantNum, arg.Number)
			}
			if tc.wantBool && !arg.Boolean {
				t.Fatalf("expected boolean flag")
			}
			if tc.wantType == ArgString && arg.String != tc.wantStr {
				t.Fatalf("string mismatch, want %q got %q", tc.wantStr, arg.String)
			}
		})
	}
}
