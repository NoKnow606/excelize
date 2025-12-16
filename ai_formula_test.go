//go:build ai_formula

// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"fmt"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAIFormula(t *testing.T) {
	// Test case 1: excel__read_sheet
	t.Run("excel__read_sheet", func(t *testing.T) {
		f := NewFile()
		defer func() {
			assert.NoError(t, f.Close())
		}()

		// Set the AI formula in a cell
		// Note: SetCellFormula expects the formula WITH _xlfn. prefix for custom functions
		// The JSON string needs to have quotes escaped (doubled) for Excel formula syntax
		toolName := "excel__read_sheet"
		jsonArgs := `{""uri"":""https://www.maybe.ai/docs/spreadsheets/d/6940d34c2cf1280961621a84?gid=3"", ""range_address"": ""B2:K100""}`
		formula := fmt.Sprintf(`_xlfn.AI("%s", "%s")`, toolName, jsonArgs)

		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))

		// Calculate the formula
		result, err := f.CalcCellValue("Sheet1", "A1")

		// Check that we got a result (success or error)
		assert.NoError(t, err)
		assert.NotEmpty(t, result)

		// Log the result with clear formatting
		t.Logf("\n========================================")
		t.Logf("Test 1: excel__read_sheet")
		t.Logf("========================================")
		t.Logf("Tool: %s", toolName)
		t.Logf("Args: %s", jsonArgs)
		t.Logf("\nResult:\n%s", result)
		t.Logf("========================================\n")

		// Check if it's an error or success
		if len(result) > 6 && result[:6] == "ERROR:" {
			t.Logf("Status: API returned error")
		} else {
			t.Logf("Status: API returned success")
		}
	})

	// Test case 2: ai_field_template__extract
	t.Run("ai_field_template__extract", func(t *testing.T) {
		f := NewFile()
		defer func() {
			assert.NoError(t, f.Close())
		}()

		// Set the AI formula in a cell
		// Note: SetCellFormula expects the formula WITH _xlfn. prefix for custom functions
		// The JSON string needs to have quotes escaped (doubled) for Excel formula syntax
		toolName := "ai_field_template__extract"
		jsonArgs := `{""input_raw_text"": ""盖尔·加朵是一位以色列女演员。她因在DC扩展宇宙电影中饰演神奇女侠而广为人知。 In 2018, Gadot was named one of Time's 100 most influential people and ranked by Forbes as the tenth-highest-paid actress, later rising to third in 2020"", ""item_to_extract"": ""person full name""}`
		formula := fmt.Sprintf(`_xlfn.AI("%s", "%s")`, toolName, jsonArgs)

		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))

		// Calculate the formula
		result, err := f.CalcCellValue("Sheet1", "A1")

		// Check that we got a result (success or error)
		assert.NoError(t, err)
		assert.NotEmpty(t, result)

		// Log the result with clear formatting
		t.Logf("\n========================================")
		t.Logf("Test 2: ai_field_template__extract")
		t.Logf("========================================")
		t.Logf("Tool: %s", toolName)
		t.Logf("Args (extract): person full name")
		t.Logf("Full JSON args: %s", jsonArgs)
		t.Logf("Note: Each call generates a unique UUID task_id")
		t.Logf("\nResult:\n%s", result)
		t.Logf("========================================\n")

		// Check if it's an error or success
		if len(result) > 6 && result[:6] == "ERROR:" {
			t.Logf("Status: API returned error")
			t.Logf("Error details: %s", result)

			// Check if it's the SSE URL format issue
			if strings.Contains(result, "Invalid SSE URL format") {
				t.Logf("\n⚠️  This is a backend configuration issue:")
				t.Logf("    The tool SSE URL in the database has invalid format")
				t.Logf("    Expected: {base_url}/api/v1/mcp/{sse_key}/{server_id}/sse")
				t.Logf("    This needs to be fixed in the backend database")
			}
		} else {
			t.Logf("Status: API returned success - extracted person name")
		}
	})
}

func TestAIFormulaErrors(t *testing.T) {
	// Test invalid argument count
	t.Run("invalid_argument_count", func(t *testing.T) {
		f := NewFile()
		defer func() {
			assert.NoError(t, f.Close())
		}()

		// Only one argument (should fail)
		formula := `_xlfn.AI("tool_name")`
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))

		result, err := f.CalcCellValue("Sheet1", "A1")
		// For formula errors, CalcCellValue may return both error and result
		if err != nil {
			assert.Contains(t, err.Error(), "AI requires 2 arguments")
		} else {
			assert.Contains(t, result, "#VALUE!")
		}
	})

	// Test invalid JSON
	t.Run("invalid_json", func(t *testing.T) {
		f := NewFile()
		defer func() {
			assert.NoError(t, f.Close())
		}()

		// Invalid JSON
		formula := `_xlfn.AI("tool_name", "not valid json")`
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))

		result, err := f.CalcCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Contains(t, result, "ERROR:")
		assert.Contains(t, result, "Invalid JSON")
	})
}

func TestAIFormulaWithCellReference(t *testing.T) {
	// Test with JSON in another cell (if we support it)
	t.Run("json_from_cell_reference", func(t *testing.T) {
		f := NewFile()
		defer func() {
			assert.NoError(t, f.Close())
		}()

		// Put JSON in a cell
		jsonArgs := `{"uri":"https://example.com", "range_address": "A1:B10"}`
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", jsonArgs))

		// Reference that cell in the formula
		formula := `_xlfn.AI("excel__read_sheet", B1)`
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))

		result, err := f.CalcCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.NotEmpty(t, result)

		t.Logf("\n========================================")
		t.Logf("Test 3: Cell Reference Test")
		t.Logf("========================================")
		t.Logf("JSON stored in cell B1, referenced in formula")
		t.Logf("\nResult:\n%s", result)
		t.Logf("========================================\n")
	})
}
