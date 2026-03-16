package excelize

import (
	"fmt"
	"strconv"
	"strings"
)

// indexMatch1DPattern represents a batch INDEX-MATCH pattern with 1D lookup
// Pattern: INDEX(array, MATCH(lookup, range, 0))
type indexMatch1DPattern struct {
	arrayRange string // e.g., "日销预测!$B:$B"
	matchRange string // e.g., "日销预测!$A:$A"
	formulas   map[string]*indexMatch1DFormula
}

type indexMatch1DFormula struct {
	cell          string
	sheet         string
	lookupCell    string // e.g., "A2"
	fallbackValue string // IFERROR fallback value
}

// averageIndexMatchPattern represents AVERAGE(INDEX(range, MATCH(...), 0)) pattern
// Pattern: AVERAGE(INDEX($C:$O, MATCH(lookup, range, 0), 0))
// Returns the average of a row range (multiple columns)
type averageIndexMatchPattern struct {
	arrayRange string // e.g., "日销售!$C:$O" (multi-column range)
	matchRange string // e.g., "日销售!$A:$A"
	formulas   map[string]*averageIndexMatchFormula
}

type averageIndexMatchFormula struct {
	cell       string
	sheet      string
	lookupCell string // e.g., "$A169"
}

// indexMatch2DPattern represents a batch INDEX-MATCH pattern with 2D lookup
// Pattern: INDEX(array, MATCH(lookup1, range1, 0), MATCH(lookup2, range2, 0))
type indexMatch2DPattern struct {
	// Common ranges (same for all formulas)
	arrayRange  string // e.g., "日销预测!$G:$ZZ"
	matchRange1 string // e.g., "日销预测!$A:$A"
	matchRange2 string // e.g., "日销预测!$G$1:$ZZ$1"

	// Formula mapping: cell -> (lookup1Value, lookup2Value)
	formulas map[string]*indexMatch2DFormula
}

// indexMatch2DFormula represents a single INDEX-MATCH formula in the batch
type indexMatch2DFormula struct {
	cell        string
	sheet       string
	lookup1Cell string // e.g., "$A2"
	lookup2Cell string // e.g., "K$1"
	lookup2Expr string // e.g., "K$1-1" (expression to calculate)
}

// batchCalculateINDEXMATCH performs batch INDEX-MATCH calculation (both 1D and 2D)
func (f *File) batchCalculateINDEXMATCH(formulas map[string]string) map[string]string {
	results := make(map[string]string)

	// Group formulas by pattern
	patterns1D := make(map[string]*indexMatch1DPattern)
	patterns2D := make(map[string]*indexMatch2DPattern)
	patternsAvg := make(map[string]*averageIndexMatchPattern)
	patternsMultiCond := make(map[string]*indexMatchMultiCondPattern)

	for fullCell, formula := range formulas {
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}

		sheet := parts[0]
		cell := parts[1]

		// Try multi-condition pattern first (most complex, e.g., MATCH(1, (cond1)*(cond2), 0))
		if isMultiCondIndexMatchFormula(formula) {
			patternMulti := f.extractIndexMatchMultiCondPattern(sheet, cell, formula)
			if patternMulti != nil {
				// Create pattern key based on result range and condition structure
				key := patternMulti.resultRange
				for _, cond := range patternMulti.conditions {
					key += "|" + cond.condType + ":" + cond.sourceRange
				}
				if existing, exists := patternsMultiCond[key]; exists {
					for k, v := range patternMulti.formulas {
						existing.formulas[k] = v
					}
				} else {
					patternsMultiCond[key] = patternMulti
				}
				continue
			}
		}

		// Try AVERAGE+INDEX-MATCH pattern first (most specific)
		patternAvg := f.extractAverageIndexMatchPattern(sheet, cell, formula)
		if patternAvg != nil {
			key := patternAvg.arrayRange + "|" + patternAvg.matchRange
			if _, exists := patternsAvg[key]; !exists {
				patternsAvg[key] = patternAvg
			} else {
				for k, v := range patternAvg.formulas {
					patternsAvg[key].formulas[k] = v
				}
			}
			continue
		}

		// Try 2D pattern
		pattern2D := f.extractINDEXMATCH2DPattern(sheet, cell, formula)
		if pattern2D != nil {
			key := pattern2D.arrayRange + "|" + pattern2D.matchRange1 + "|" + pattern2D.matchRange2
			if _, exists := patterns2D[key]; !exists {
				patterns2D[key] = pattern2D
			} else {
				for k, v := range pattern2D.formulas {
					patterns2D[key].formulas[k] = v
				}
			}
			continue
		}

		// Try 1D pattern
		pattern1D := f.extractINDEXMATCH1DPattern(sheet, cell, formula)
		if pattern1D != nil {
			key := pattern1D.arrayRange + "|" + pattern1D.matchRange
			if _, exists := patterns1D[key]; !exists {
				patterns1D[key] = pattern1D
			} else {
				for k, v := range pattern1D.formulas {
					patterns1D[key].formulas[k] = v
				}
			}
		}
	}

	// Calculate multi-condition patterns first
	for _, pattern := range patternsMultiCond {
		patternResults := f.calculateIndexMatchMultiCondPatternWithCache(pattern, nil)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	// Calculate AVERAGE+INDEX-MATCH patterns
	for _, pattern := range patternsAvg {
		patternResults := f.calculateAverageIndexMatchPattern(pattern)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	// Calculate 1D patterns
	for _, pattern := range patterns1D {
		patternResults := f.calculateINDEXMATCH1DPattern(pattern)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	// Calculate 2D patterns
	for _, pattern := range patterns2D {
		patternResults := f.calculateINDEXMATCH2DPattern(pattern)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	return results
}

// extractINDEXMATCH2DPattern extracts INDEX-MATCH 2D pattern from formula
func (f *File) extractINDEXMATCH2DPattern(sheet, cell, formula string) *indexMatch2DPattern {
	// Check if formula contains INDEX and MATCH
	if !strings.Contains(formula, "INDEX(") || !strings.Contains(formula, "MATCH(") {
		return nil
	}

	// Try to parse: INDEX(array, MATCH(lookup1, range1, 0), MATCH(lookup2, range2, 0))
	// or IFERROR(INDEX(...), default)

	// Remove leading = if present
	workFormula := strings.TrimPrefix(formula, "=")

	// Remove IFERROR wrapper if present
	if strings.HasPrefix(workFormula, "IFERROR(") {
		// Extract the INDEX part
		idx := strings.Index(workFormula, "INDEX(")
		if idx > 0 {
			workFormula = workFormula[idx:]
		}
	}

	// Find INDEX(
	indexStart := strings.Index(workFormula, "INDEX(")
	if indexStart == -1 {
		return nil
	}

	// Extract the full INDEX expression
	indexExpr := extractFunctionCall(workFormula[indexStart:], "INDEX")
	if indexExpr == "" {
		return nil
	}

	// Parse INDEX arguments: array, row, col
	args := splitFunctionArgs(indexExpr)
	if len(args) < 3 {
		// Not a 2D pattern, let 1D handler try
		return nil
	}

	arrayRange := strings.TrimSpace(args[0])
	rowExpr := strings.TrimSpace(args[1])
	colExpr := strings.TrimSpace(args[2])

	// Check if rowExpr is MATCH
	if !strings.HasPrefix(rowExpr, "MATCH(") {
		return nil
	}

	// Parse first MATCH - extract the content inside MATCH()
	match1Content := extractFunctionCall(rowExpr, "MATCH")
	if match1Content == "" {
		return nil
	}
	matchArgs1 := splitFunctionArgs(match1Content)
	if len(matchArgs1) < 2 {
		return nil
	}

	lookup1Cell := strings.TrimSpace(matchArgs1[0])
	matchRange1 := strings.TrimSpace(matchArgs1[1])

	// Check if colExpr is MATCH
	var lookup2Cell, lookup2Expr, matchRange2 string
	if strings.HasPrefix(colExpr, "MATCH(") {
		// Parse second MATCH - extract the content inside MATCH()
		match2Content := extractFunctionCall(colExpr, "MATCH")
		if match2Content == "" {
			return nil
		}
		matchArgs2 := splitFunctionArgs(match2Content)
		if len(matchArgs2) < 2 {
			return nil
		}

		lookup2Expr = strings.TrimSpace(matchArgs2[0])
		lookup2Cell = lookup2Expr // May be expression like "K$1-1"
		matchRange2 = strings.TrimSpace(matchArgs2[1])
	} else {
		// Not a 2D INDEX-MATCH pattern
		return nil
	}

	// Create pattern
	pattern := &indexMatch2DPattern{
		arrayRange:  arrayRange,
		matchRange1: matchRange1,
		matchRange2: matchRange2,
		formulas:    make(map[string]*indexMatch2DFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &indexMatch2DFormula{
		cell:        cell,
		sheet:       sheet,
		lookup1Cell: lookup1Cell,
		lookup2Cell: lookup2Cell,
		lookup2Expr: lookup2Expr,
	}

	return pattern
}

// calculateINDEXMATCH2DPattern calculates a batch of INDEX-MATCH formulas
func (f *File) calculateINDEXMATCH2DPattern(pattern *indexMatch2DPattern) map[string]string {
	results := make(map[string]string)

	// Extract source sheet from array range
	sourceSheet := extractSheetName(pattern.arrayRange)
	if sourceSheet == "" {
		return results
	}

	// Parse array range to get column range
	// e.g., "日销预测!$G:$ZZ" -> columns G to ZZ
	arrayParts := strings.Split(pattern.arrayRange, "!")
	if len(arrayParts) != 2 {
		return results
	}

	colRange := arrayParts[1]
	// Extract start and end columns from $G:$ZZ
	colRange = strings.ReplaceAll(colRange, "$", "")
	colParts := strings.Split(colRange, ":")
	if len(colParts) != 2 {
		return results
	}

	startCol := colParts[0]
	endCol := colParts[1]

	// Read the array data (entire range)
	rows, err := f.GetRows(sourceSheet)
	if err != nil || len(rows) == 0 {
		return results
	}

	// Build row lookup map (first MATCH dimension)
	// Parse matchRange1: e.g., "日销预测!$A:$A"
	matchCol1 := extractColumnFromRange(pattern.matchRange1)
	matchCol1Idx, _ := ColumnNameToNumber(matchCol1)
	matchCol1Idx-- // Convert to 0-based

	rowLookupMap := make(map[string]int) // value -> row index
	if matchCol1Idx >= 0 {
		for rowIdx, row := range rows {
			if matchCol1Idx < len(row) {
				value := row[matchCol1Idx]
				if value != "" {
					rowLookupMap[value] = rowIdx
				}
			}
		}
	}

	// Build column lookup map (second MATCH dimension)
	// Parse matchRange2: e.g., "日销预测!$G$1:$ZZ$1"
	colLookupMap := make(map[string]int) // value -> column index (relative to start)

	// Get the first row for column headers
	if len(rows) > 0 {
		headerRow := rows[0]
		startColIdx, _ := ColumnNameToNumber(startCol)
		endColIdx, _ := ColumnNameToNumber(endCol)
		startColIdx-- // Convert to 0-based
		endColIdx--

		for colIdx := startColIdx; colIdx <= endColIdx && colIdx < len(headerRow); colIdx++ {
			value := headerRow[colIdx]
			if value != "" {
				colLookupMap[value] = colIdx - startColIdx
			}
		}
	}

	// Pre-calculate all lookup values in batch to avoid repeated GetCellValue calls
	// Build lookup value cache
	lookupValueCache := make(map[string]string)

	for _, info := range pattern.formulas {
		// Cache lookup1 value
		lookup1Cell := strings.ReplaceAll(info.lookup1Cell, "$", "")
		cacheKey1 := info.sheet + "!" + lookup1Cell
		if _, exists := lookupValueCache[cacheKey1]; !exists {
			// Note: This function doesn't have worksheetCache available, fallback to GetCellValue
			lookupValueCache[cacheKey1], _ = f.GetCellValue(info.sheet, lookup1Cell)
		}

		// Cache lookup2 value (handle expressions like "K$1-1")
		lookup2Cell := info.lookup2Cell
		// Extract base cell reference (before any operators)
		for _, op := range []string{"-", "+"} {
			if idx := strings.Index(lookup2Cell, op); idx > 0 {
				lookup2Cell = lookup2Cell[:idx]
				break
			}
		}
		lookup2Cell = strings.ReplaceAll(lookup2Cell, "$", "")
		cacheKey2 := info.sheet + "!" + lookup2Cell
		if _, exists := lookupValueCache[cacheKey2]; !exists {
			lookupValueCache[cacheKey2], _ = f.GetCellValue(info.sheet, lookup2Cell)
		}
	}

	// Calculate results for all formulas using cached lookup values
	startColIdx, _ := ColumnNameToNumber(startCol)
	startColIdx--

	for fullCell, info := range pattern.formulas {
		// Get lookup1 value from cache
		lookup1Cell := strings.ReplaceAll(info.lookup1Cell, "$", "")
		cacheKey1 := info.sheet + "!" + lookup1Cell
		lookup1Value := lookupValueCache[cacheKey1]

		// Get lookup2 value from cache and evaluate expression if needed
		var lookup2Value string
		if strings.Contains(info.lookup2Expr, "-") || strings.Contains(info.lookup2Expr, "+") {
			// Extract base cell reference
			lookup2Cell := info.lookup2Cell
			for _, op := range []string{"-", "+"} {
				if idx := strings.Index(lookup2Cell, op); idx > 0 {
					lookup2Cell = lookup2Cell[:idx]
					break
				}
			}
			lookup2Cell = strings.ReplaceAll(lookup2Cell, "$", "")
			cacheKey2 := info.sheet + "!" + lookup2Cell
			cellVal := lookupValueCache[cacheKey2]

			// Evaluate simple arithmetic expressions
			if strings.Contains(info.lookup2Expr, "-1") {
				// Parse as number and subtract 1
				if num, err := strconv.ParseFloat(cellVal, 64); err == nil {
					lookup2Value = strconv.FormatFloat(num-1, 'f', -1, 64)
				} else {
					lookup2Value = cellVal
				}
			} else if strings.Contains(info.lookup2Expr, "+1") {
				// Parse as number and add 1
				if num, err := strconv.ParseFloat(cellVal, 64); err == nil {
					lookup2Value = strconv.FormatFloat(num+1, 'f', -1, 64)
				} else {
					lookup2Value = cellVal
				}
			} else {
				lookup2Value = cellVal
			}
		} else {
			lookup2Cell := strings.ReplaceAll(info.lookup2Cell, "$", "")
			cacheKey2 := info.sheet + "!" + lookup2Cell
			lookup2Value = lookupValueCache[cacheKey2]
		}

		// Lookup in the 2D array
		if rowIdx, ok := rowLookupMap[lookup1Value]; ok {
			if colOffset, ok := colLookupMap[lookup2Value]; ok {
				actualColIdx := startColIdx + colOffset
				if rowIdx < len(rows) && actualColIdx < len(rows[rowIdx]) {
					results[fullCell] = rows[rowIdx][actualColIdx]
				} else {
					results[fullCell] = "0"
				}
			} else {
				results[fullCell] = "0"
			}
		} else {
			results[fullCell] = "0"
		}
	}

	return results
}

// extractFunctionCall extracts a complete function call like "FUNC(...)"
func extractFunctionCall(s string, funcName string) string {
	// Find function name
	idx := strings.Index(s, funcName+"(")
	if idx == -1 {
		return ""
	}

	start := idx + len(funcName)
	depth := 0
	inQuote := false

	for i := start; i < len(s); i++ {
		ch := s[i]

		switch ch {
		case '"', '\'':
			inQuote = !inQuote
		case '(':
			if !inQuote {
				depth++
			}
		case ')':
			if !inQuote {
				depth--
				if depth == 0 {
					// Found the closing parenthesis
					return s[start+1 : i] // Return content inside parentheses
				}
			}
		}
	}

	return ""
}

// splitFunctionArgs splits function arguments (handles nested functions and quotes)
func splitFunctionArgs(argsStr string) []string {
	var result []string
	var current strings.Builder
	depth := 0
	inQuote := false

	for i := 0; i < len(argsStr); i++ {
		ch := argsStr[i]

		switch ch {
		case '"', '\'':
			inQuote = !inQuote
			current.WriteByte(ch)
		case '(':
			if !inQuote {
				depth++
			}
			current.WriteByte(ch)
		case ')':
			if !inQuote {
				depth--
			}
			current.WriteByte(ch)
		case ',':
			if !inQuote && depth == 0 {
				// This is an argument separator
				result = append(result, current.String())
				current.Reset()
			} else {
				current.WriteByte(ch)
			}
		default:
			current.WriteByte(ch)
		}
	}

	if current.Len() > 0 {
		result = append(result, current.String())
	}

	return result
}

// extractINDEXMATCHFromFormula extracts INDEX-MATCH expression from a formula
func extractINDEXMATCHFromFormula(formula string) string {
	// Find "INDEX(" in the formula (may be nested in IFERROR)
	idx := strings.Index(formula, "INDEX(")
	if idx == -1 {
		return ""
	}

	// Extract the complete INDEX(...) expression
	start := idx
	depth := 0
	inQuote := false

	for i := start; i < len(formula); i++ {
		ch := formula[i]

		switch ch {
		case '"', '\'':
			inQuote = !inQuote
		case '(':
			if !inQuote {
				depth++
			}
		case ')':
			if !inQuote {
				depth--
				if depth == 0 {
					// Found the closing parenthesis
					return formula[start : i+1]
				}
			}
		}
	}

	return ""
}

// extractINDEXMATCH1DPattern extracts INDEX-MATCH 1D pattern from formula
func (f *File) extractINDEXMATCH1DPattern(sheet, cell, formula string) *indexMatch1DPattern {
	// Check if formula contains INDEX and MATCH
	if !strings.Contains(formula, "INDEX(") || !strings.Contains(formula, "MATCH(") {
		return nil
	}

	// Extract IFERROR fallback value if present
	var fallbackValue string
	workFormula := strings.TrimPrefix(formula, "=")
	originalFormula := workFormula

	if strings.HasPrefix(workFormula, "IFERROR(") {
		// Extract the full IFERROR expression
		iferrorContent := extractFunctionCall(workFormula, "IFERROR")
		if iferrorContent != "" {
			// Split IFERROR arguments: first is the formula, second is the fallback
			iferrorArgs := splitFunctionArgs(iferrorContent)
			if len(iferrorArgs) >= 2 {
				fallbackValue = strings.TrimSpace(iferrorArgs[1])
				// Remove quotes if it's a literal string
				fallbackValue = strings.Trim(fallbackValue, `"'`)
			}
		}
	}

	// Remove wrapper functions (IFERROR, AVERAGE, etc.) to find INDEX
	// Support patterns like:
	//   IFERROR(AVERAGE(INDEX(...)))
	//   AVERAGE(INDEX(...))
	//   IFERROR(INDEX(...))
	workFormula = originalFormula
	for {
		trimmed := false

		// Remove IFERROR wrapper
		if strings.HasPrefix(workFormula, "IFERROR(") {
			idx := strings.Index(workFormula, "INDEX(")
			if idx > 0 {
				workFormula = workFormula[idx:]
				trimmed = true
			}
		}

		// Remove AVERAGE wrapper
		if strings.HasPrefix(workFormula, "AVERAGE(") {
			idx := strings.Index(workFormula, "INDEX(")
			if idx > 0 {
				workFormula = workFormula[idx:]
				trimmed = true
			}
		}

		// If no wrapper was removed, break the loop
		if !trimmed {
			break
		}
	}

	// Find INDEX(
	indexStart := strings.Index(workFormula, "INDEX(")
	if indexStart == -1 {
		return nil
	}

	// Extract the full INDEX expression
	indexExpr := extractFunctionCall(workFormula[indexStart:], "INDEX")
	if indexExpr == "" {
		return nil
	}

	// Parse INDEX arguments
	args := splitFunctionArgs(indexExpr)
	if len(args) != 2 {
		return nil // Must have exactly 2 args for 1D
	}

	arrayRange := strings.TrimSpace(args[0])
	rowExpr := strings.TrimSpace(args[1])

	// Check if rowExpr is MATCH
	if !strings.HasPrefix(rowExpr, "MATCH(") {
		return nil
	}

	// Parse MATCH
	matchContent := extractFunctionCall(rowExpr, "MATCH")
	if matchContent == "" {
		return nil
	}
	matchArgs := splitFunctionArgs(matchContent)
	if len(matchArgs) < 2 {
		return nil
	}

	lookupCell := strings.TrimSpace(matchArgs[0])
	matchRange := strings.TrimSpace(matchArgs[1])

	// Create pattern
	pattern := &indexMatch1DPattern{
		arrayRange: arrayRange,
		matchRange: matchRange,
		formulas:   make(map[string]*indexMatch1DFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &indexMatch1DFormula{
		cell:          cell,
		sheet:         sheet,
		lookupCell:    lookupCell,
		fallbackValue: fallbackValue,
	}

	return pattern
}

// calculateINDEXMATCH1DPattern calculates a batch of 1D INDEX-MATCH formulas
func (f *File) calculateINDEXMATCH1DPattern(pattern *indexMatch1DPattern) map[string]string {
	results := make(map[string]string)

	// Extract source sheet from array range
	sourceSheet := extractSheetName(pattern.arrayRange)
	if sourceSheet == "" {
		return results
	}

	// Parse array range to get column
	arrayParts := strings.Split(pattern.arrayRange, "!")
	if len(arrayParts) != 2 {
		return results
	}

	arrayColPart := strings.ReplaceAll(arrayParts[1], "$", "")
	// Handle column range like "B:B" - extract just the first column
	if strings.Contains(arrayColPart, ":") {
		parts := strings.Split(arrayColPart, ":")
		if len(parts) > 0 {
			arrayColPart = parts[0]
		}
	}
	arrayColIdx, _ := ColumnNameToNumber(arrayColPart)
	arrayColIdx-- // Convert to 0-based

	// Build lookup map from match range
	matchCol := extractColumnFromRange(pattern.matchRange)
	matchColIdx, _ := ColumnNameToNumber(matchCol)
	matchColIdx-- // Convert to 0-based

	// Read source data
	rows, err := f.GetRows(sourceSheet)
	if err != nil || len(rows) == 0 {
		return results
	}

	// Build lookup map: value -> row index
	lookupMap := make(map[string]int)
	if matchColIdx >= 0 {
		for rowIdx, row := range rows {
			if matchColIdx < len(row) {
				value := row[matchColIdx]
				if value != "" {
					lookupMap[value] = rowIdx
				}
			}
		}
	}

	// Note: This function doesn't have worksheetCache available, so it uses the old approach
	// It's only used in non-optimized batch calculations
	// Calculate results for all formulas
	for fullCell, info := range pattern.formulas {
		// Get lookup value - using fallback without worksheetCache
		lookupCell := strings.ReplaceAll(info.lookupCell, "$", "")
		lookupValue, _ := f.GetCellValue(info.sheet, lookupCell)

		// Lookup in the array
		if rowIdx, ok := lookupMap[lookupValue]; ok {
			if rowIdx < len(rows) && arrayColIdx < len(rows[rowIdx]) {
				results[fullCell] = rows[rowIdx][arrayColIdx]
			} else {
				results[fullCell] = ""
			}
		} else {
			// No match found - use fallback value if available
			if info.fallbackValue != "" {
				// Fallback value can be a cell reference or a literal value
				if isCellReference(info.fallbackValue) {
					// It's a cell reference like "A1"
					fallbackVal, _ := f.GetCellValue(info.sheet, info.fallbackValue)
					results[fullCell] = fallbackVal
				} else {
					// It's a literal value
					results[fullCell] = info.fallbackValue
				}
			} else {
				results[fullCell] = ""
			}
		}
	}

	return results
}

// extractAverageIndexMatchPattern extracts AVERAGE(INDEX(...MATCH...)) pattern
// Pattern: AVERAGE(INDEX($C:$O, MATCH(lookup, range, 0), 0))
// or: IFERROR(AVERAGE(INDEX($C:$O, MATCH(lookup, range, 0), 0)), 0)
func (f *File) extractAverageIndexMatchPattern(sheet, cell, formula string) *averageIndexMatchPattern {
	// Check if formula contains AVERAGE, INDEX and MATCH
	if !strings.Contains(formula, "AVERAGE(") || !strings.Contains(formula, "INDEX(") || !strings.Contains(formula, "MATCH(") {
		return nil
	}

	// Remove leading = if present
	workFormula := strings.TrimPrefix(formula, "=")

	// Remove IFERROR wrapper if present
	if strings.HasPrefix(workFormula, "IFERROR(") {
		idx := strings.Index(workFormula, "AVERAGE(")
		if idx > 0 {
			workFormula = workFormula[idx:]
		}
	}

	// Check if starts with AVERAGE(
	if !strings.HasPrefix(workFormula, "AVERAGE(") {
		return nil
	}

	// Extract AVERAGE content
	averageContent := extractFunctionCall(workFormula, "AVERAGE")
	if averageContent == "" {
		return nil
	}

	// Check if AVERAGE contains INDEX
	if !strings.HasPrefix(averageContent, "INDEX(") {
		return nil
	}

	// Extract INDEX content
	indexExpr := extractFunctionCall(averageContent, "INDEX")
	if indexExpr == "" {
		return nil
	}

	// Parse INDEX arguments
	args := splitFunctionArgs(indexExpr)
	if len(args) != 3 {
		return nil // Must have 3 args: array, row, col
	}

	arrayRange := strings.TrimSpace(args[0])
	rowExpr := strings.TrimSpace(args[1])
	colExpr := strings.TrimSpace(args[2])

	// Check if colExpr is 0 (return entire row)
	if colExpr != "0" {
		return nil
	}

	// Check if rowExpr is MATCH
	if !strings.HasPrefix(rowExpr, "MATCH(") {
		return nil
	}

	// Parse MATCH
	matchContent := extractFunctionCall(rowExpr, "MATCH")
	if matchContent == "" {
		return nil
	}
	matchArgs := splitFunctionArgs(matchContent)
	if len(matchArgs) < 2 {
		return nil
	}

	lookupCell := strings.TrimSpace(matchArgs[0])
	matchRange := strings.TrimSpace(matchArgs[1])

	// Create pattern
	pattern := &averageIndexMatchPattern{
		arrayRange: arrayRange,
		matchRange: matchRange,
		formulas:   make(map[string]*averageIndexMatchFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &averageIndexMatchFormula{
		cell:       cell,
		sheet:      sheet,
		lookupCell: lookupCell,
	}

	return pattern
}

// calculateAverageIndexMatchPatternWithCache calculates AVERAGE(INDEX(...)) batch using worksheetCache
// This function reads data from worksheetCache for recalculated values, falling back to file for original data
func (f *File) calculateAverageIndexMatchPatternWithCache(pattern *averageIndexMatchPattern, worksheetCache *WorksheetCache) map[string]string {
	typedResults := f.calculateAverageIndexMatchPatternArgsWithCache(pattern, worksheetCache)
	results := make(map[string]string, len(typedResults))
	for cell, value := range typedResults {
		results[cell] = value.Value()
	}
	return results
}

func (f *File) calculateAverageIndexMatchPatternArgsWithCache(pattern *averageIndexMatchPattern, worksheetCache *WorksheetCache) map[string]formulaArg {
	results := make(map[string]formulaArg)
	if worksheetCache == nil {
		for cell, value := range f.calculateAverageIndexMatchPattern(pattern) {
			results[cell] = inferFormulaResultType(value)
		}
		return results
	}

	// Extract source sheet from array range
	sourceSheet := extractSheetName(pattern.arrayRange)
	if sourceSheet == "" {
		return results
	}

	// Parse array range to get column range (e.g., "$C:$O" -> C to O)
	arrayParts := strings.Split(pattern.arrayRange, "!")
	if len(arrayParts) != 2 {
		return results
	}

	// Parse column range like "$C:$O"
	rangePart := strings.ReplaceAll(arrayParts[1], "$", "")
	rangeParts := strings.Split(rangePart, ":")
	if len(rangeParts) != 2 {
		return results
	}

	startCol := rangeParts[0]
	endCol := rangeParts[1]
	startColIdx, _ := ColumnNameToNumber(startCol)
	endColIdx, _ := ColumnNameToNumber(endCol)

	// Build lookup map from match range
	matchCol := extractColumnFromRange(pattern.matchRange)
	matchColIdx, _ := ColumnNameToNumber(matchCol)

	// Ensure raw source cells are present; existing formula results stay in the cache.
	if err := worksheetCache.LoadSheet(f, sourceSheet); err != nil {
		return results
	}

	rows := f.convertCacheToArgRows(worksheetCache.GetSheet(sourceSheet))
	if len(rows) == 0 {
		return results
	}

	// Build lookup map: value -> row index (0-based)
	lookupMap := make(map[string]int)
	for rowIdx, row := range rows {
		if matchColIdx-1 < len(row) {
			value := row[matchColIdx-1]
			if value.Value() != "" {
				lookupMap[formulaArgTypedKey(value)] = rowIdx
			}
		}
	}

	// Calculate results for all formulas
	for fullCell, info := range pattern.formulas {
		lookupCell := strings.ReplaceAll(info.lookupCell, "$", "")
		lookupValue := f.getCellArgOrCalcCache(info.sheet, lookupCell, worksheetCache)

		// Lookup in the map
		if rowIdx, ok := lookupMap[formulaArgTypedKey(lookupValue)]; ok {
			if rowIdx >= 0 && rowIdx < len(rows) {
				// Calculate average of the row range (startColIdx to endColIdx, 1-based)
				sum := 0.0
				count := 0
				for colIdx := startColIdx - 1; colIdx <= endColIdx-1 && colIdx < len(rows[rowIdx]); colIdx++ {
					cellValue := rows[rowIdx][colIdx]
					if cellValue.Type == ArgNumber && !cellValue.Boolean {
						sum += cellValue.Number
						count++
					}
				}

				if count > 0 {
					results[fullCell] = newNumberFormulaArg(sum / float64(count))
				} else {
					results[fullCell] = newNumberFormulaArg(0)
				}
			} else {
				results[fullCell] = newNumberFormulaArg(0)
			}
		} else {
			results[fullCell] = newNumberFormulaArg(0)
		}
	}

	return results
}

// calculateAverageIndexMatchPattern calculates AVERAGE(INDEX(...)) batch (legacy version without worksheetCache)
func (f *File) calculateAverageIndexMatchPattern(pattern *averageIndexMatchPattern) map[string]string {
	results := make(map[string]string)

	// Extract source sheet from array range
	sourceSheet := extractSheetName(pattern.arrayRange)
	if sourceSheet == "" {
		return results
	}

	// Parse array range to get column range (e.g., "$C:$O" -> C to O)
	arrayParts := strings.Split(pattern.arrayRange, "!")
	if len(arrayParts) != 2 {
		return results
	}

	// Parse column range like "$C:$O"
	rangePart := strings.ReplaceAll(arrayParts[1], "$", "")
	rangeParts := strings.Split(rangePart, ":")
	if len(rangeParts) != 2 {
		return results
	}

	startCol := rangeParts[0]
	endCol := rangeParts[1]
	startColIdx, _ := ColumnNameToNumber(startCol)
	endColIdx, _ := ColumnNameToNumber(endCol)

	// Build lookup map from match range
	matchCol := extractColumnFromRange(pattern.matchRange)
	matchColIdx, _ := ColumnNameToNumber(matchCol)

	// Read data using GetRows (legacy method)
	rows, err := f.GetRows(sourceSheet, Options{RawCellValue: true})
	if err != nil || len(rows) == 0 {
		return results
	}

	// Build lookup map: value -> row index (0-based)
	lookupMap := make(map[string]int)
	for rowIdx, row := range rows {
		if matchColIdx-1 < len(row) {
			value := row[matchColIdx-1]
			if value != "" {
				lookupMap[value] = rowIdx
			}
		}
	}

	// Calculate results for all formulas
	for fullCell, info := range pattern.formulas {
		// Get lookup value
		lookupCell := strings.ReplaceAll(info.lookupCell, "$", "")
		lookupValue, _ := f.GetCellValue(info.sheet, lookupCell)

		// Lookup in the map
		if rowIdx, ok := lookupMap[lookupValue]; ok {
			if rowIdx >= 0 && rowIdx < len(rows) {
				// Calculate average of the row range (startColIdx to endColIdx, 1-based)
				sum := 0.0
				count := 0
				for colIdx := startColIdx - 1; colIdx <= endColIdx-1 && colIdx < len(rows[rowIdx]); colIdx++ {
					cellValue := rows[rowIdx][colIdx]
					if cellValue != "" {
						if val, err := strconv.ParseFloat(cellValue, 64); err == nil {
							sum += val
							count++
						}
						// Skip non-numeric values (text like "断货")
					}
				}

				if count > 0 {
					avg := sum / float64(count)
					results[fullCell] = fmt.Sprintf("%g", avg)
				} else {
					results[fullCell] = "0"
				}
			} else {
				results[fullCell] = "0"
			}
		} else {
			results[fullCell] = "0"
		}
	}

	return results
}

// convertCacheToRows converts worksheetCache map format to [][]string format
// This allows existing code to work with minimal changes
// Phase 1: 改为接收 map[string]formulaArg
func (f *File) convertCacheToRows(sheetData map[string]formulaArg) [][]string {
	if len(sheetData) == 0 {
		return [][]string{}
	}

	// Find max row and col
	maxRow, maxCol := 0, 0
	for cellRef := range sheetData {
		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}
		if row > maxRow {
			maxRow = row
		}
		if col > maxCol {
			maxCol = col
		}
	}

	// Create 2D array
	rows := make([][]string, maxRow)
	for i := range rows {
		rows[i] = make([]string, maxCol)
	}

	// Fill in values
	// Phase 1: 调用 Value() 方法将 formulaArg 转换为字符串
	for cellRef, argValue := range sheetData {
		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}
		rows[row-1][col-1] = argValue.Value()
	}

	return rows
}

func (f *File) convertCacheToArgRows(sheetData map[string]formulaArg) [][]formulaArg {
	if len(sheetData) == 0 {
		return [][]formulaArg{}
	}

	maxRow, maxCol := 0, 0
	for cellRef := range sheetData {
		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}
		if row > maxRow {
			maxRow = row
		}
		if col > maxCol {
			maxCol = col
		}
	}

	rows := make([][]formulaArg, maxRow)
	for i := range rows {
		rows[i] = make([]formulaArg, maxCol)
		for j := range rows[i] {
			rows[i][j] = newEmptyFormulaArg()
		}
	}

	for cellRef, argValue := range sheetData {
		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}
		rows[row-1][col-1] = argValue
	}

	return rows
}

func indexMatchLookupExprValue(baseArg formulaArg, expr string) formulaArg {
	if strings.Contains(expr, "-1") {
		num := baseArg.ToNumber()
		if num.Type == ArgNumber && !num.Boolean {
			return newNumberFormulaArg(num.Number - 1)
		}
		return baseArg
	}
	if strings.Contains(expr, "+1") {
		num := baseArg.ToNumber()
		if num.Type == ArgNumber && !num.Boolean {
			return newNumberFormulaArg(num.Number + 1)
		}
		return baseArg
	}
	return baseArg
}

func resolveIndexMatchCellRef(defaultSheet, ref string) (sheet, cell string, ok bool) {
	ref = strings.TrimSpace(ref)
	if strings.Contains(ref, "!") {
		parts := strings.SplitN(ref, "!", 2)
		sheet = strings.Trim(parts[0], "'")
		cell = strings.ReplaceAll(parts[1], "$", "")
		return sheet, cell, isCellReference(cell)
	}
	cell = strings.ReplaceAll(ref, "$", "")
	return defaultSheet, cell, isCellReference(cell)
}

func (f *File) resolveIndexMatchFallbackArg(defaultSheet, fallback string, worksheetCache *WorksheetCache) formulaArg {
	if sheet, cell, ok := resolveIndexMatchCellRef(defaultSheet, fallback); ok {
		return f.getCellArgOrCalcCache(sheet, cell, worksheetCache)
	}
	if literal, ok := parseFormulaLiteralArg(fallback); ok {
		return literal
	}
	return newStringFormulaArg(fallback)
}

// batchCalculateINDEXMATCHWithCache performs batch INDEX-MATCH calculation using worksheetCache
func (f *File) batchCalculateINDEXMATCHWithCache(formulas map[string]string, worksheetCache *WorksheetCache) map[string]string {
	typedResults := f.batchCalculateINDEXMATCHArgsWithCache(formulas, worksheetCache)
	results := make(map[string]string, len(typedResults))
	for cell, value := range typedResults {
		results[cell] = value.Value()
	}
	return results
}

func (f *File) batchCalculateINDEXMATCHArgsWithCache(formulas map[string]string, worksheetCache *WorksheetCache) map[string]formulaArg {
	results := make(map[string]formulaArg)
	if worksheetCache == nil {
		for cell, value := range f.batchCalculateINDEXMATCH(formulas) {
			results[cell] = inferFormulaResultType(value)
		}
		return results
	}

	// Group formulas by pattern
	patterns1D := make(map[string]*indexMatch1DPattern)
	patterns2D := make(map[string]*indexMatch2DPattern)
	patternsAvg := make(map[string]*averageIndexMatchPattern)
	patternsMultiCond := make(map[string]*indexMatchMultiCondPattern)

	for fullCell, formula := range formulas {
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}

		sheet := parts[0]
		cell := parts[1]

		// Try multi-condition pattern first (most complex, e.g., MATCH(1, (cond1)*(cond2), 0))
		if isMultiCondIndexMatchFormula(formula) {
			patternMulti := f.extractIndexMatchMultiCondPattern(sheet, cell, formula)
			if patternMulti != nil {
				// Create pattern key based on result range and condition structure
				key := patternMulti.resultRange
				for _, cond := range patternMulti.conditions {
					key += "|" + cond.condType + ":" + cond.sourceRange
				}
				if existing, exists := patternsMultiCond[key]; exists {
					for k, v := range patternMulti.formulas {
						existing.formulas[k] = v
					}
				} else {
					patternsMultiCond[key] = patternMulti
				}
				continue
			}
		}

		// Try AVERAGE+INDEX-MATCH pattern first (most specific)
		patternAvg := f.extractAverageIndexMatchPattern(sheet, cell, formula)
		if patternAvg != nil {
			key := patternAvg.arrayRange + "|" + patternAvg.matchRange
			if _, exists := patternsAvg[key]; !exists {
				patternsAvg[key] = patternAvg
			} else {
				for k, v := range patternAvg.formulas {
					patternsAvg[key].formulas[k] = v
				}
			}
			continue
		}

		// Try 2D pattern
		pattern2D := f.extractINDEXMATCH2DPattern(sheet, cell, formula)
		if pattern2D != nil {
			key := pattern2D.arrayRange + "|" + pattern2D.matchRange1 + "|" + pattern2D.matchRange2
			if _, exists := patterns2D[key]; !exists {
				patterns2D[key] = pattern2D
			} else {
				for k, v := range pattern2D.formulas {
					patterns2D[key].formulas[k] = v
				}
			}
			continue
		}

		// Try 1D pattern
		pattern1D := f.extractINDEXMATCH1DPattern(sheet, cell, formula)
		if pattern1D != nil {
			key := pattern1D.arrayRange + "|" + pattern1D.matchRange
			if _, exists := patterns1D[key]; !exists {
				patterns1D[key] = pattern1D
			} else {
				for k, v := range pattern1D.formulas {
					patterns1D[key].formulas[k] = v
				}
			}
		}
	}

	// Calculate multi-condition patterns first
	for _, pattern := range patternsMultiCond {
		patternResults := f.calculateIndexMatchMultiCondPatternWithCache(pattern, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = inferFormulaResultType(value)
		}
	}

	// Calculate AVERAGE+INDEX-MATCH patterns (use worksheetCache for recalculated values)
	for _, pattern := range patternsAvg {
		patternResults := f.calculateAverageIndexMatchPatternArgsWithCache(pattern, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	// Calculate 1D patterns (use worksheetCache)
	for _, pattern := range patterns1D {
		patternResults := f.calculateINDEXMATCH1DPatternArgsWithCache(pattern, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	// Calculate 2D patterns (use worksheetCache)
	for _, pattern := range patterns2D {
		patternResults := f.calculateINDEXMATCH2DPatternArgsWithCache(pattern, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	return results
}

// calculateINDEXMATCH2DPatternWithCache calculates a batch of INDEX-MATCH formulas using worksheetCache
func (f *File) calculateINDEXMATCH2DPatternWithCache(pattern *indexMatch2DPattern, worksheetCache *WorksheetCache) map[string]string {
	typedResults := f.calculateINDEXMATCH2DPatternArgsWithCache(pattern, worksheetCache)
	results := make(map[string]string, len(typedResults))
	for cell, value := range typedResults {
		results[cell] = value.Value()
	}
	return results
}

func (f *File) calculateINDEXMATCH2DPatternArgsWithCache(pattern *indexMatch2DPattern, worksheetCache *WorksheetCache) map[string]formulaArg {
	results := make(map[string]formulaArg)
	if worksheetCache == nil {
		for cell, value := range f.calculateINDEXMATCH2DPattern(pattern) {
			results[cell] = inferFormulaResultType(value)
		}
		return results
	}

	// Extract source sheet from array range
	sourceSheet := extractSheetName(pattern.arrayRange)
	if sourceSheet == "" {
		return results
	}

	// Parse array range to get column range
	arrayParts := strings.Split(pattern.arrayRange, "!")
	if len(arrayParts) != 2 {
		return results
	}

	colRange := arrayParts[1]
	colRange = strings.ReplaceAll(colRange, "$", "")
	colParts := strings.Split(colRange, ":")
	if len(colParts) != 2 {
		return results
	}

	startCol := colParts[0]
	endCol := colParts[1]

	// CRITICAL FIX: Always read from file first, then merge cached formula results
	// The worksheetCache only contains formula calculation results, NOT original data.
	// For INDEX-MATCH, we need the original data (e.g., A column for MATCH lookup)
	sheetData := worksheetCache.GetSheet(sourceSheet)

	// Build lookup maps
	matchCol1 := extractColumnFromRange(pattern.matchRange1)
	matchCol1Idx, _ := ColumnNameToNumber(matchCol1)
	matchCol1Idx--

	if err := worksheetCache.LoadSheet(f, sourceSheet); err != nil {
		return results
	}
	sheetData = worksheetCache.GetSheet(sourceSheet)
	rows := f.convertCacheToArgRows(sheetData)
	if len(rows) == 0 {
		return results
	}

	rowLookupMap := make(map[string]int)
	if matchCol1Idx >= 0 {
		for rowIdx, row := range rows {
			if matchCol1Idx < len(row) {
				value := row[matchCol1Idx]
				if value.Value() != "" {
					rowLookupMap[formulaArgTypedKey(value)] = rowIdx
				}
			}
		}
	}

	colLookupMap := make(map[string]int)
	if len(rows) > 0 {
		headerRow := rows[0]
		startColIdx, _ := ColumnNameToNumber(startCol)
		endColIdx, _ := ColumnNameToNumber(endCol)
		startColIdx--
		endColIdx--

		for colIdx := startColIdx; colIdx <= endColIdx && colIdx < len(headerRow); colIdx++ {
			value := headerRow[colIdx]
			if value.Value() != "" {
				colLookupMap[formulaArgTypedKey(value)] = colIdx - startColIdx
			}
		}
	}

	startColIdx, _ := ColumnNameToNumber(startCol)
	startColIdx--

	// Pre-calculate all lookup values
	lookupValueCache := make(map[string]formulaArg)

	for _, info := range pattern.formulas {
		lookup1Cell := strings.ReplaceAll(info.lookup1Cell, "$", "")
		cacheKey1 := info.sheet + "!" + lookup1Cell
		if _, exists := lookupValueCache[cacheKey1]; !exists {
			lookupValueCache[cacheKey1] = f.getCellArgOrCalcCache(info.sheet, lookup1Cell, worksheetCache)
		}

		lookup2Cell := info.lookup2Cell
		for _, op := range []string{"-", "+"} {
			if idx := strings.Index(lookup2Cell, op); idx > 0 {
				lookup2Cell = lookup2Cell[:idx]
				break
			}
		}
		lookup2Cell = strings.ReplaceAll(lookup2Cell, "$", "")
		cacheKey2 := info.sheet + "!" + lookup2Cell
		if _, exists := lookupValueCache[cacheKey2]; !exists {
			lookupValueCache[cacheKey2] = f.getCellArgOrCalcCache(info.sheet, lookup2Cell, worksheetCache)
		}
	}

	// Calculate results
	for fullCell, info := range pattern.formulas {
		lookup1Cell := strings.ReplaceAll(info.lookup1Cell, "$", "")
		cacheKey1 := info.sheet + "!" + lookup1Cell
		lookup1Value := lookupValueCache[cacheKey1]

		var lookup2Value formulaArg
		if strings.Contains(info.lookup2Expr, "-") || strings.Contains(info.lookup2Expr, "+") {
			lookup2Cell := info.lookup2Cell
			for _, op := range []string{"-", "+"} {
				if idx := strings.Index(lookup2Cell, op); idx > 0 {
					lookup2Cell = lookup2Cell[:idx]
					break
				}
			}
			lookup2Cell = strings.ReplaceAll(lookup2Cell, "$", "")
			cacheKey2 := info.sheet + "!" + lookup2Cell
			lookup2Value = indexMatchLookupExprValue(lookupValueCache[cacheKey2], info.lookup2Expr)
		} else {
			lookup2Cell := strings.ReplaceAll(info.lookup2Cell, "$", "")
			cacheKey2 := info.sheet + "!" + lookup2Cell
			lookup2Value = lookupValueCache[cacheKey2]
		}

		if rowIdx, ok := rowLookupMap[formulaArgTypedKey(lookup1Value)]; ok {
			if colOffset, ok := colLookupMap[formulaArgTypedKey(lookup2Value)]; ok {
				actualColIdx := startColIdx + colOffset
				if rowIdx < len(rows) && actualColIdx < len(rows[rowIdx]) {
					results[fullCell] = rows[rowIdx][actualColIdx]
				} else {
					results[fullCell] = newNumberFormulaArg(0)
				}
			} else {
				results[fullCell] = newNumberFormulaArg(0)
			}
		} else {
			results[fullCell] = newNumberFormulaArg(0)
		}
	}

	return results
}

// calculateINDEXMATCH1DPatternWithCache calculates INDEX-MATCH 1D using worksheetCache
func (f *File) calculateINDEXMATCH1DPatternWithCache(pattern *indexMatch1DPattern, worksheetCache *WorksheetCache) map[string]string {
	typedResults := f.calculateINDEXMATCH1DPatternArgsWithCache(pattern, worksheetCache)
	results := make(map[string]string, len(typedResults))
	for cell, value := range typedResults {
		results[cell] = value.Value()
	}
	return results
}

func (f *File) calculateINDEXMATCH1DPatternArgsWithCache(pattern *indexMatch1DPattern, worksheetCache *WorksheetCache) map[string]formulaArg {
	results := make(map[string]formulaArg)
	if worksheetCache == nil {
		for cell, value := range f.calculateINDEXMATCH1DPattern(pattern) {
			results[cell] = inferFormulaResultType(value)
		}
		return results
	}

	sourceSheet := extractSheetName(pattern.arrayRange)
	if sourceSheet == "" {
		return results
	}

	arrayParts := strings.Split(pattern.arrayRange, "!")
	if len(arrayParts) != 2 {
		return results
	}

	arrayColPart := strings.ReplaceAll(arrayParts[1], "$", "")
	if strings.Contains(arrayColPart, ":") {
		parts := strings.Split(arrayColPart, ":")
		if len(parts) > 0 {
			arrayColPart = parts[0]
		}
	}
	arrayColIdx, _ := ColumnNameToNumber(arrayColPart)
	arrayColIdx--

	matchCol := extractColumnFromRange(pattern.matchRange)
	matchColIdx, _ := ColumnNameToNumber(matchCol)
	matchColIdx--

	if err := worksheetCache.LoadSheet(f, sourceSheet); err != nil {
		return results
	}
	rows := f.convertCacheToArgRows(worksheetCache.GetSheet(sourceSheet))
	if len(rows) == 0 {
		return results
	}

	// Build lookup map
	lookupMap := make(map[string]int)
	if matchColIdx >= 0 {
		for rowIdx, row := range rows {
			if matchColIdx < len(row) {
				value := row[matchColIdx]
				if value.Value() != "" {
					lookupMap[formulaArgTypedKey(value)] = rowIdx
				}
			}
		}
	}

	// Calculate results
	for fullCell, info := range pattern.formulas {
		lookupCell := strings.ReplaceAll(info.lookupCell, "$", "")
		lookupValue := f.getCellArgOrCalcCache(info.sheet, lookupCell, worksheetCache)

		if rowIdx, ok := lookupMap[formulaArgTypedKey(lookupValue)]; ok {
			if rowIdx < len(rows) && arrayColIdx < len(rows[rowIdx]) {
				results[fullCell] = rows[rowIdx][arrayColIdx]
			} else {
				results[fullCell] = newStringFormulaArg("")
			}
		} else {
			if info.fallbackValue != "" {
				results[fullCell] = f.resolveIndexMatchFallbackArg(info.sheet, info.fallbackValue, worksheetCache)
			} else {
				results[fullCell] = newStringFormulaArg("")
			}
		}
	}

	return results
}
