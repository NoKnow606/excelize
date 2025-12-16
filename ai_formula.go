//go:build !no_ai_formula

// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bytes"
	"container/list"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"strings"
	"time"

	"github.com/google/uuid"
)

// AI implements the AI custom formula for calling FastestAI API.
// Syntax: =AI(tool_name, json_args)
//
// Parameters:
//   - tool_name: The name of the MCP tool to call (e.g., "excel__read_sheet")
//   - json_args: JSON string containing the tool arguments
//
// Example: =AI("excel__read_sheet", "{\"uri\":\"https://example.com/sheet\", \"range_address\": \"A1:B10\"}")
//
// Returns the cell_value from the API response, or an error string prefixed with "ERROR:".
func (fn *formulaFuncs) AI(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "AI requires 2 arguments: tool_name and json_args")
	}

	// Get tool_name argument
	toolNameArg := argsList.Front().Value.(formulaArg)
	if toolNameArg.Type == ArgError {
		return toolNameArg
	}
	toolName := toolNameArg.String

	// Get json_args argument
	jsonArgsArg := argsList.Back().Value.(formulaArg)
	if jsonArgsArg.Type == ArgError {
		return jsonArgsArg
	}

	// Parse JSON args - try to handle both direct JSON string and cell reference
	// Excel formulas use doubled quotes ("") to escape quotes inside strings
	// We need to convert them back to single quotes for JSON parsing
	var toolArgs map[string]interface{}
	jsonStr := strings.ReplaceAll(jsonArgsArg.String, `""`, `"`)

	if err := json.Unmarshal([]byte(jsonStr), &toolArgs); err != nil {
		return newStringFormulaArg(fmt.Sprintf("ERROR: Invalid JSON arguments: %s", err.Error()))
	}

	// Call the FastestAI API
	result, err := callFastestAIAPI(fn, toolName, toolArgs)
	if err != nil {
		return newStringFormulaArg(fmt.Sprintf("ERROR: %s", err.Error()))
	}

	return newStringFormulaArg(result)
}

// apiRequest represents the request payload for FastestAI API
type apiRequest struct {
	UserID        string                 `json:"user_id"`
	SpreadsheetID string                 `json:"spreadsheet_id"`
	SheetID       string                 `json:"sheet_id"`
	Cell          apiCellLocation        `json:"cell"`
	Formula       string                 `json:"formula"`
	ToolID        string                 `json:"tool_id"`
	ToolName      string                 `json:"tool_name"`
	ToolArgs      map[string]interface{} `json:"tool_args"`
	TaskID        string                 `json:"task_id,omitempty"` // UUID for task binding
}

// apiCellLocation represents a cell location in the request
type apiCellLocation struct {
	Row    int `json:"row"`
	Column int `json:"column"`
}

// apiResponse represents the response from FastestAI API
type apiResponse struct {
	Success      bool                   `json:"success"`
	CellValue    string                 `json:"cell_value"`
	Error        *string                `json:"error,omitempty"`
	RawResponse  map[string]interface{} `json:"raw_response,omitempty"`
	Metadata     map[string]interface{} `json:"metadata,omitempty"`
	ExecutionMS  *float64               `json:"execution_time_ms,omitempty"`
	Timestamp    *string                `json:"timestamp,omitempty"`
}

// callFastestAIAPI makes the HTTP request to the FastestAI API endpoint
func callFastestAIAPI(fn *formulaFuncs, toolName string, toolArgs map[string]interface{}) (string, error) {
	// Try to derive context from the Excel file
	// For now, use defaults (Option A) - we can enhance this later for Option C
	userID := "excel-user"
	spreadsheetID := "excel-workbook"
	sheetID := fn.sheet
	if sheetID == "" {
		sheetID = "Sheet1"
	}

	// Try to parse cell location if available
	row, col := 0, 0
	if fn.cell != "" {
		// TODO: Parse cell reference to get row/column if needed
		// For now, just use defaults
	}

	// Generate a unique task_id (UUID) for this request
	// This allows the backend to call get_or_create_task_binding
	taskID := uuid.New().String()

	// Build the request payload
	// Use tool_name for both tool_id and tool_name since API requires tool_id
	reqPayload := apiRequest{
		UserID:        userID,
		SpreadsheetID: spreadsheetID,
		SheetID:       sheetID,
		Cell: apiCellLocation{
			Row:    row,
			Column: col,
		},
		Formula:  fmt.Sprintf("=AI(\"%s\", ...)", toolName),
		ToolID:   toolName, // API requires tool_id, so use tool_name for it
		ToolName: toolName,
		ToolArgs: toolArgs,
		TaskID:   taskID, // UUID for task binding to avoid SSE URL issues
	}

	// For debugging: log the task_id being sent (optional, can be removed in production)
	// fmt.Printf("DEBUG: AI formula sending task_id=%s for tool=%s\n", taskID, toolName)

	// Marshal to JSON
	jsonData, err := json.Marshal(reqPayload)
	if err != nil {
		return "", fmt.Errorf("failed to marshal request: %v", err)
	}

	// Create HTTP client with timeout
	client := &http.Client{
		Timeout: 30 * time.Second,
	}

	// Make POST request
	resp, err := client.Post(
		"https://api.fastest.ai/v1/tool/function_call",
		"application/json",
		bytes.NewBuffer(jsonData),
	)
	if err != nil {
		return "", fmt.Errorf("API request failed: %v", err)
	}
	defer resp.Body.Close()

	// Read response body
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", fmt.Errorf("failed to read response: %v", err)
	}

	// Check HTTP status
	if resp.StatusCode != http.StatusOK {
		return "", fmt.Errorf("API returned status %d: %s", resp.StatusCode, string(body))
	}

	// Parse response
	var apiResp apiResponse
	if err := json.Unmarshal(body, &apiResp); err != nil {
		return "", fmt.Errorf("failed to parse response: %v", err)
	}

	// Check if the API call was successful
	if !apiResp.Success {
		if apiResp.Error != nil {
			// Return detailed error with raw response if available
			if apiResp.RawResponse != nil {
				rawJSON, _ := json.Marshal(apiResp.RawResponse)
				return "", fmt.Errorf("%s (raw: %s)", *apiResp.Error, string(rawJSON))
			}
			return "", fmt.Errorf("%s", *apiResp.Error)
		}
		// If no error message but failed, try to show raw response
		if apiResp.RawResponse != nil {
			rawJSON, _ := json.Marshal(apiResp.RawResponse)
			return "", fmt.Errorf("API call failed (raw: %s)", string(rawJSON))
		}
		return "", fmt.Errorf("API call failed")
	}

	return apiResp.CellValue, nil
}
