// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"os"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestUpdateFormulaCache(t *testing.T) {
	t.Run("basic formula caching", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Set up test data
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1+A2"))

		// Update cache
		assert.NoError(t, f.UpdateFormulaCache())

		// Save and reopen
		assert.NoError(t, f.SaveAs("test_formula_cache.xlsx"))
		f2, err := OpenFile("test_formula_cache.xlsx")
		assert.NoError(t, err)
		defer func() {
			assert.NoError(t, f2.Close())
			assert.NoError(t, os.Remove("test_formula_cache.xlsx"))
		}()

		// Verify cached value is available
		value, err := f2.GetCellValue("Sheet1", "A3")
		assert.NoError(t, err)
		assert.Equal(t, "30", value)
	})

	t.Run("multiple formulas", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Set up multiple formulas
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 5))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*2"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A2+10"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A4", "A1+A2+A3"))

		// Update cache
		assert.NoError(t, f.UpdateFormulaCache())

		// Save and reopen
		assert.NoError(t, f.SaveAs("test_multiple_formulas.xlsx"))
		f2, err := OpenFile("test_multiple_formulas.xlsx")
		assert.NoError(t, err)
		defer func() {
			assert.NoError(t, f2.Close())
			assert.NoError(t, os.Remove("test_multiple_formulas.xlsx"))
		}()

		// Verify all cached values
		v2, _ := f2.GetCellValue("Sheet1", "A2")
		v3, _ := f2.GetCellValue("Sheet1", "A3")
		v4, _ := f2.GetCellValue("Sheet1", "A4")
		assert.Equal(t, "10", v2) // 5*2
		assert.Equal(t, "20", v3) // 10+10
		assert.Equal(t, "35", v4) // 5+10+20
	})

	t.Run("formatted values", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Set up formula with percentage format
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 0.5))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*2"))

		stylePercent, err := f.NewStyle(&Style{NumFmt: 9}) // 0%
		assert.NoError(t, err)
		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", stylePercent))

		// Update cache
		assert.NoError(t, f.UpdateFormulaCache())

		// Save and reopen
		assert.NoError(t, f.SaveAs("test_formatted_formula.xlsx"))
		f2, err := OpenFile("test_formatted_formula.xlsx")
		assert.NoError(t, err)
		defer func() {
			assert.NoError(t, f2.Close())
			assert.NoError(t, os.Remove("test_formatted_formula.xlsx"))
		}()

		// Verify formatted value
		formatted, _ := f2.GetCellValue("Sheet1", "A2")
		assert.Equal(t, "100%", formatted)

		// Verify raw value
		raw, _ := f2.GetCellValue("Sheet1", "A2", Options{RawCellValue: true})
		assert.Equal(t, "1", raw)
	})

	t.Run("date and time formats", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Set up formulas with date/time formats
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 44927)) // 2023-01-01
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 0.5))   // 12:00:00
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1+1"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "B1+0.25"))

		styleDate, _ := f.NewStyle(&Style{NumFmt: 14}) // m/d/yy
		styleTime, _ := f.NewStyle(&Style{NumFmt: 18}) // h:mm AM/PM
		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", styleDate))
		assert.NoError(t, f.SetCellStyle("Sheet1", "B2", "B2", styleTime))

		// Update cache
		assert.NoError(t, f.UpdateFormulaCache())

		// Save and reopen
		assert.NoError(t, f.SaveAs("test_datetime_formula.xlsx"))
		f2, err := OpenFile("test_datetime_formula.xlsx")
		assert.NoError(t, err)
		defer func() {
			assert.NoError(t, f2.Close())
			assert.NoError(t, os.Remove("test_datetime_formula.xlsx"))
		}()

		// Verify raw values are cached
		rawDate, _ := f2.GetCellValue("Sheet1", "A2", Options{RawCellValue: true})
		rawTime, _ := f2.GetCellValue("Sheet1", "B2", Options{RawCellValue: true})
		assert.Equal(t, "44928", rawDate)
		assert.Equal(t, "0.75", rawTime)

		// Verify formatted values are applied on read
		formattedDate, _ := f2.GetCellValue("Sheet1", "A2")
		formattedTime, _ := f2.GetCellValue("Sheet1", "B2")
		// Date format should be applied
		assert.NotEqual(t, "44928", formattedDate)
		assert.Contains(t, []string{"01-02-23", "1/2/23", "2023-01-02"}, formattedDate)
		// Time format should be applied
		assert.NotEqual(t, "0.75", formattedTime)
		assert.Contains(t, []string{"6:00 PM", "18:00"}, formattedTime)
	})

	t.Run("empty sheet", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// No formulas, should not error
		assert.NoError(t, f.UpdateFormulaCache())
	})

	t.Run("multiple sheets", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Create second sheet
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)

		// Add formulas to both sheets
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*2"))

		assert.NoError(t, f.SetCellValue("Sheet2", "A1", 20))
		assert.NoError(t, f.SetCellFormula("Sheet2", "A2", "A1*3"))

		// Update cache for all sheets
		assert.NoError(t, f.UpdateFormulaCache())

		// Save and reopen
		assert.NoError(t, f.SaveAs("test_multiple_sheets.xlsx"))
		f2, err := OpenFile("test_multiple_sheets.xlsx")
		assert.NoError(t, err)
		defer func() {
			assert.NoError(t, f2.Close())
			assert.NoError(t, os.Remove("test_multiple_sheets.xlsx"))
		}()

		// Verify both sheets
		v1, _ := f2.GetCellValue("Sheet1", "A2")
		v2, _ := f2.GetCellValue("Sheet2", "A2")
		assert.Equal(t, "20", v1)
		assert.Equal(t, "60", v2)
	})
}

func TestParseFormulaTokensCached(t *testing.T) {
	f := NewFile()
	defer func() { assert.NoError(t, f.Close()) }()

	tokens1 := f.parseFormulaTokensCached("SUM(A1:A3)+1")
	tokens2 := f.parseFormulaTokensCached("SUM(A1:A3)+1")
	assert.NotNil(t, tokens1)
	assert.NotNil(t, tokens2)
	assert.NotZero(t, len(tokens1))
	assert.Same(t, &tokens1[0], &tokens2[0])

	assert.Nil(t, f.parseFormulaTokensCached(""))
}

func TestUpdateSheetFormulaCache(t *testing.T) {
	t.Run("single sheet update", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Create two sheets with formulas
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*2"))

		assert.NoError(t, f.SetCellValue("Sheet2", "A1", 20))
		assert.NoError(t, f.SetCellFormula("Sheet2", "A2", "A1*3"))

		// Update only Sheet1
		assert.NoError(t, f.UpdateSheetFormulaCache("Sheet1"))

		// Save and reopen
		assert.NoError(t, f.SaveAs("test_single_sheet.xlsx"))
		f2, err := OpenFile("test_single_sheet.xlsx")
		assert.NoError(t, err)
		defer func() {
			assert.NoError(t, f2.Close())
			assert.NoError(t, os.Remove("test_single_sheet.xlsx"))
		}()

		// Sheet1 should have cached value
		v1, _ := f2.GetCellValue("Sheet1", "A2")
		assert.Equal(t, "20", v1)

		// Sheet2 should NOT have cached value (empty string)
		v2, _ := f2.GetCellValue("Sheet2", "A2")
		assert.Equal(t, "", v2)
	})

	t.Run("invalid sheet", func(t *testing.T) {
		f := NewFile()
		defer func() { assert.NoError(t, f.Close()) }()

		// Update non-existent sheet
		err := f.UpdateSheetFormulaCache("NonExistent")
		assert.Error(t, err)
	})
}

func TestCalcCellValueCacheKeyWithRawValue(t *testing.T) {
	f := NewFile()
	defer func() { assert.NoError(t, f.Close()) }()

	// Set up formula with formatting
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 0.5))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*2"))

	stylePercent, _ := f.NewStyle(&Style{NumFmt: 9})
	assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", stylePercent))

	t.Run("cache keys are different for raw vs formatted", func(t *testing.T) {
		// Clear any existing cache
		f.calcCache = sync.Map{}

		// First call with default (formatted)
		v1, err := f.CalcCellValue("Sheet1", "A2")
		assert.NoError(t, err)
		assert.Equal(t, "100%", v1)

		// Second call with RawCellValue=true should return different value
		v2, err := f.CalcCellValue("Sheet1", "A2", Options{RawCellValue: true})
		assert.NoError(t, err)
		assert.Equal(t, "1", v2)

		// Third call with default should still return formatted
		v3, err := f.CalcCellValue("Sheet1", "A2")
		assert.NoError(t, err)
		assert.Equal(t, "100%", v3)

		// Fourth call with RawCellValue=true should still return raw
		v4, err := f.CalcCellValue("Sheet1", "A2", Options{RawCellValue: true})
		assert.NoError(t, err)
		assert.Equal(t, "1", v4)
	})

	t.Run("cache works correctly with different call orders", func(t *testing.T) {
		// Clear cache and test opposite order
		f.calcCache = sync.Map{}

		// First call with RawCellValue=true
		v1, err := f.CalcCellValue("Sheet1", "A2", Options{RawCellValue: true})
		assert.NoError(t, err)
		assert.Equal(t, "1", v1)

		// Second call with default (formatted)
		v2, err := f.CalcCellValue("Sheet1", "A2")
		assert.NoError(t, err)
		assert.Equal(t, "100%", v2)

		// Verify consistency on repeated calls
		v3, err := f.CalcCellValue("Sheet1", "A2", Options{RawCellValue: true})
		assert.NoError(t, err)
		assert.Equal(t, "1", v3)

		v4, err := f.CalcCellValue("Sheet1", "A2")
		assert.NoError(t, err)
		assert.Equal(t, "100%", v4)
	})
}
