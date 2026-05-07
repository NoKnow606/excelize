package excelize

import (
	"fmt"
	"strings"
	"testing"

	"github.com/stretchr/testify/require"
)

type parsedRangeCacheKey struct {
	sheet          string
	fromRow, toRow int
	fromCol, toCol int
}

func parseRangeCacheKey(t *testing.T, key string) parsedRangeCacheKey {
	t.Helper()

	parts := strings.SplitN(key, "!", 2)
	require.Len(t, parts, 2, "unexpected range cache key format %q", key)

	var parsed parsedRangeCacheKey
	parsed.sheet = parts[0]
	_, err := fmt.Sscanf(parts[1], "R%dC%d:R%dC%d", &parsed.fromRow, &parsed.fromCol, &parsed.toRow, &parsed.toCol)
	require.NoError(t, err, "unexpected range cache key format %q", key)
	return parsed
}

func requireTrimmedRangeCacheKey(t *testing.T, f *File, expectedKey string) {
	t.Helper()

	expected := parseRangeCacheKey(t, expectedKey)
	foundOptimized := false
	f.rangeCache.Range(func(key string, value interface{}) bool {
		parsed := parseRangeCacheKey(t, key)
		if key == expectedKey {
			foundOptimized = true
		}
		if parsed.sheet == expected.sheet &&
			parsed.fromRow == expected.fromRow &&
			parsed.fromCol == expected.fromCol &&
			parsed.toCol == expected.toCol &&
			parsed.toRow > expected.toRow {
			t.Fatalf("expected trimmed range cache key not to exceed %q, got %q", expectedKey, key)
		}
		return true
	})
	require.True(t, foundOptimized, "expected trimmed range cache key %q", expectedKey)
}

func TestCalcRangeOptimizationClampsBoundedSumRangeToWorksheetBounds(t *testing.T) {
	f := NewFile()
	defer f.Close()

	require.NoError(t, f.SetCellValue("Sheet1", "L3", 10))
	require.NoError(t, f.SetCellValue("Sheet1", "L10", 5))
	require.NoError(t, f.SetCellFormula("Sheet1", "M1", "SUM(L3:L1000000)"))

	got, err := f.CalcCellValue("Sheet1", "M1")
	require.NoError(t, err)
	require.Equal(t, "15", got)

	optimizedKey := "Sheet1!R3C12:R10C12"
	expected := parseRangeCacheKey(t, optimizedKey)
	foundOptimized := false
	f.rangeCache.Range(func(key string, value interface{}) bool {
		parsed := parseRangeCacheKey(t, key)
		if key == optimizedKey {
			foundOptimized = true
			matrix, ok := value.([][]formulaArg)
			require.True(t, ok, "range cache value should store the materialized matrix")
			require.Len(t, matrix, 8)
			require.Len(t, matrix[0], 1)
		}
		if parsed.sheet == expected.sheet &&
			parsed.fromRow == expected.fromRow &&
			parsed.fromCol == expected.fromCol &&
			parsed.toCol == expected.toCol &&
			parsed.toRow > expected.toRow {
			t.Fatalf("expected SUM range cache key not to exceed %q, got %q", optimizedKey, key)
		}
		return true
	})
	require.True(t, foundOptimized, "expected trimmed SUM range cache key %q", optimizedKey)
}

func TestCalcRangeOptimizationClampsMoreSafeAggregateFunctions(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		expected string
	}{
		{name: "COUNT", formula: "COUNT(L3:L1000000)", expected: "2"},
		{name: "COUNTA", formula: "COUNTA(L3:L1000000)", expected: "2"},
		{name: "MAX", formula: "MAX(L3:L1000000)", expected: "10"},
		{name: "MIN", formula: "MIN(L3:L1000000)", expected: "5"},
		{name: "PRODUCT", formula: "PRODUCT(L3:L1000000)", expected: "50"},
		{name: "MEDIAN", formula: "MEDIAN(L3:L1000000)", expected: "7.5"},
		{name: "LARGE", formula: "LARGE(L3:L1000000,1)", expected: "10"},
		{name: "SMALL", formula: "SMALL(L3:L1000000,1)", expected: "5"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()

			require.NoError(t, f.SetCellValue("Sheet1", "L3", 10))
			require.NoError(t, f.SetCellValue("Sheet1", "L10", 5))
			require.NoError(t, f.SetCellFormula("Sheet1", "M1", tt.formula))

			got, err := f.CalcCellValue("Sheet1", "M1")
			require.NoError(t, err)
			require.Equal(t, tt.expected, got)

			optimizedKey := "Sheet1!R3C12:R10C12"
			expected := parseRangeCacheKey(t, optimizedKey)
			foundOptimized := false
			f.rangeCache.Range(func(key string, value interface{}) bool {
				parsed := parseRangeCacheKey(t, key)
				if key == optimizedKey {
					foundOptimized = true
				}
				if parsed.sheet == expected.sheet &&
					parsed.fromRow == expected.fromRow &&
					parsed.fromCol == expected.fromCol &&
					parsed.toCol == expected.toCol &&
					parsed.toRow > expected.toRow {
					t.Fatalf("expected %s range cache key not to exceed %q, got %q", tt.name, optimizedKey, key)
				}
				return true
			})
			require.True(t, foundOptimized, "expected trimmed %s range cache key %q", tt.name, optimizedKey)
		})
	}
}

func TestCalcRangeOptimizationKeepsCountBlankRangeShape(t *testing.T) {
	f := NewFile()
	defer f.Close()

	require.NoError(t, f.SetCellValue("Sheet1", "L3", 1))
	require.NoError(t, f.SetCellFormula("Sheet1", "M1", "COUNTBLANK(L3:L100)"))

	got, err := f.CalcCellValue("Sheet1", "M1")
	require.NoError(t, err)
	require.Equal(t, "97", got)

	fullKey := "Sheet1!R3C12:R100C12"
	foundFull := false
	f.rangeCache.Range(func(key string, value interface{}) bool {
		if key == "Sheet1!R3C12:R3C12" {
			t.Fatalf("COUNTBLANK range should not be trimmed to worksheet bounds")
		}
		if key == fullKey {
			foundFull = true
			matrix, ok := value.([][]formulaArg)
			require.True(t, ok, "range cache value should store the materialized matrix")
			require.Len(t, matrix, 98)
			require.Len(t, matrix[0], 1)
		}
		return true
	})
	require.True(t, foundFull, "expected COUNTBLANK to keep full range cache key %q", fullKey)
}

func TestCalcRangeOptimizationWidensPastStaleSmallerSheetDimension(t *testing.T) {
	f := NewFile()
	defer f.Close()

	require.NoError(t, f.SetCellValue("Sheet1", "L3", 10))
	require.NoError(t, f.SetCellValue("Sheet1", "L10", 5))
	require.NoError(t, f.SetCellFormula("Sheet1", "M1", "SUM(L3:L1000000)"))

	ws, err := f.workSheetReader("Sheet1")
	require.NoError(t, err)
	ws.Dimension = &xlsxDimension{Ref: "A1:M3"}

	got, err := f.CalcCellValue("Sheet1", "M1")
	require.NoError(t, err)
	require.Equal(t, "15", got)

	requireTrimmedRangeCacheKey(t, f, "Sheet1!R3C12:R10C12")
}

func TestCalcRangeOptimizationFallsBackWhenSheetDimensionMissing(t *testing.T) {
	f := NewFile()
	defer f.Close()

	require.NoError(t, f.SetCellValue("Sheet1", "L3", 10))
	require.NoError(t, f.SetCellValue("Sheet1", "L10", 5))
	require.NoError(t, f.SetCellFormula("Sheet1", "M1", "SUM(L3:L1000000)"))

	ws, err := f.workSheetReader("Sheet1")
	require.NoError(t, err)
	ws.Dimension = nil

	got, err := f.CalcCellValue("Sheet1", "M1")
	require.NoError(t, err)
	require.Equal(t, "15", got)

	requireTrimmedRangeCacheKey(t, f, "Sheet1!R3C12:R10C12")
}

func TestCalcRangeOptimizationUsesCurrentInnerFunctionContext(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		expected string
	}{
		{name: "NestedIfSum", formula: `IF(SUM(L3:L1000000)>0,1,0)`, expected: "1"},
		{name: "XlfnSum", formula: `_xlfn.SUM(L3:L1000000)`, expected: "15"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			f := NewFile()
			defer f.Close()

			require.NoError(t, f.SetCellValue("Sheet1", "L3", 10))
			require.NoError(t, f.SetCellValue("Sheet1", "L10", 5))
			require.NoError(t, f.SetCellFormula("Sheet1", "M1", tt.formula))

			got, err := f.CalcCellValue("Sheet1", "M1")
			require.NoError(t, err)
			require.Equal(t, tt.expected, got)

			requireTrimmedRangeCacheKey(t, f, "Sheet1!R3C12:R10C12")
		})
	}
}
