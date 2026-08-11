package excelize

import (
	"log"
	"regexp"
	"strings"
	"time"
)

// indexMatchMultiCondPattern represents a batch INDEX-MATCH pattern with multiple conditions
// Pattern: INDEX(resultRange, MATCH(1, (cond1)*(cond2)*..., 0))
// Example: IFERROR(INDEX(订单!N:N, MATCH(1, (订单!B:B=ERP!B2)*ISNUMBER(FIND(ERP!F2,订单!P:P)), 0)), "")
type indexMatchMultiCondPattern struct {
	resultRange string // e.g., "订单!N:N" - the column to return values from
	sourceSheet string // e.g., "订单" - the sheet containing the data

	// Condition definitions
	conditions []multiCondition

	// Formula mapping: fullCell -> formula info
	formulas map[string]*indexMatchMultiCondFormula
}

// multiCondition represents a single condition in the MATCH array
type multiCondition struct {
	condType    string // "equal" for (col=val), "find" for ISNUMBER(FIND(val, col))
	sourceRange string // e.g., "订单!B:B" or "订单!P:P"
	sourceCol   int    // column index in source sheet (0-based)
}

// indexMatchMultiCondFormula represents a single formula in the batch
type indexMatchMultiCondFormula struct {
	cell            string
	sheet           string
	lookupCells     []string // e.g., ["ERP!B2", "ERP!F2"] - cells containing lookup values
	originalFormula string
	fallbackValue   string // IFERROR fallback value (e.g., "AE3791" or literal value)
}

// extractIndexMatchMultiCondPattern extracts multi-condition INDEX-MATCH pattern from formula
// Supports patterns like:
//
//	IFERROR(INDEX(订单!N:N, MATCH(1, (订单!B:B=ERP!B2)*ISNUMBER(FIND(ERP!F2,订单!P:P)), 0)), "")
func (f *File) extractIndexMatchMultiCondPattern(sheet, cell, formula string) *indexMatchMultiCondPattern {
	// Check if formula contains INDEX and MATCH
	upperFormula := strings.ToUpper(formula)
	if !strings.Contains(upperFormula, "INDEX(") || !strings.Contains(upperFormula, "MATCH(") {
		return nil
	}

	// Check for MATCH(1, ..., 0) pattern - this is the key indicator
	if !strings.Contains(formula, "MATCH(1,") {
		return nil
	}

	// Extract IFERROR fallback value if present
	var fallbackValue string
	workFormula := strings.TrimPrefix(formula, "=")
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

		// Find the INDEX( position
		idxPos := strings.Index(workFormula, "INDEX(")
		if idxPos > 0 {
			workFormula = workFormula[idxPos:]
		}
	}

	// Extract INDEX arguments
	indexExpr := extractFunctionCall(workFormula, "INDEX")
	if indexExpr == "" {
		return nil
	}

	args := splitFunctionArgs(indexExpr)
	if len(args) != 2 {
		return nil // Must have exactly 2 args: result_range, MATCH(...)
	}

	resultRange := strings.TrimSpace(args[0])
	matchExpr := strings.TrimSpace(args[1])

	// Check if second arg is MATCH(1, ..., 0)
	if !strings.HasPrefix(matchExpr, "MATCH(") {
		return nil
	}

	// Extract MATCH arguments
	matchContent := extractFunctionCall(matchExpr, "MATCH")
	if matchContent == "" {
		return nil
	}

	matchArgs := splitFunctionArgs(matchContent)
	if len(matchArgs) < 3 {
		return nil
	}

	// First arg must be "1"
	if strings.TrimSpace(matchArgs[0]) != "1" {
		return nil
	}

	// Third arg must be "0" (exact match)
	if strings.TrimSpace(matchArgs[2]) != "0" {
		return nil
	}

	// Second arg is the condition array: (cond1)*(cond2)*...
	conditionExpr := strings.TrimSpace(matchArgs[1])

	// Parse conditions
	conditions, lookupCells := parseMultiConditions(conditionExpr)
	if len(conditions) == 0 {
		return nil
	}
	if !isWholeColumnRange(resultRange) {
		return nil
	}
	for _, condition := range conditions {
		if !isWholeColumnRange(condition.sourceRange) {
			return nil
		}
	}

	// Get source sheet from result range
	sourceSheet := extractSheetName(resultRange)
	if sourceSheet == "" {
		return nil
	}

	// Create pattern
	pattern := &indexMatchMultiCondPattern{
		resultRange: resultRange,
		sourceSheet: sourceSheet,
		conditions:  conditions,
		formulas:    make(map[string]*indexMatchMultiCondFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &indexMatchMultiCondFormula{
		cell:            cell,
		sheet:           sheet,
		lookupCells:     lookupCells,
		originalFormula: formula,
		fallbackValue:   fallbackValue,
	}

	return pattern
}

// parseMultiConditions parses condition expression like (订单!B:B=ERP!B2)*ISNUMBER(FIND(ERP!F2,订单!P:P))
// Returns conditions and lookup cells
func parseMultiConditions(expr string) ([]multiCondition, []string) {
	var conditions []multiCondition
	var lookupCells []string

	// Split by * to get individual conditions
	parts := splitByMultiply(expr)

	for _, part := range parts {
		part = strings.TrimSpace(part)

		// Only trim outer parentheses for simple conditions (not function calls)
		// Check if it's a simple parenthesized expression like (col=val)
		if strings.HasPrefix(part, "(") && !strings.Contains(strings.ToUpper(part), "ISNUMBER") && !strings.Contains(strings.ToUpper(part), "FIND(") {
			part = strings.Trim(part, "()")
		}

		// Check for equality condition: col=val or col=cell
		if strings.Contains(part, "=") && !strings.Contains(strings.ToUpper(part), "FIND(") {
			// Pattern: 订单!B:B=ERP!B2
			eqParts := strings.SplitN(part, "=", 2)
			if len(eqParts) == 2 {
				sourceRange := strings.TrimSpace(eqParts[0])
				lookupCell := strings.TrimSpace(eqParts[1])

				// Get column index
				col := getRangeColumnIndex(sourceRange)

				conditions = append(conditions, multiCondition{
					condType:    "equal",
					sourceRange: sourceRange,
					sourceCol:   col,
				})
				lookupCells = append(lookupCells, lookupCell)
			}
			continue
		}

		// Check for FIND condition: ISNUMBER(FIND(val, col))
		upperPart := strings.ToUpper(part)
		if strings.Contains(upperPart, "ISNUMBER(FIND(") || (strings.Contains(upperPart, "ISNUMBER(") && strings.Contains(upperPart, "FIND(")) {
			// Extract FIND arguments - need to find the inner FIND call
			// Pattern: ISNUMBER(FIND(search_text, within_text))
			findIdx := strings.Index(upperPart, "FIND(")
			if findIdx == -1 {
				continue
			}

			// Extract from original case string starting at FIND
			findPart := part[findIdx:]

			// Find matching closing parenthesis for FIND
			depth := 0
			endIdx := -1
			inFind := false
			for i, ch := range findPart {
				if ch == '(' {
					if !inFind {
						inFind = true
					}
					depth++
				} else if ch == ')' {
					depth--
					if depth == 0 {
						endIdx = i
						break
					}
				}
			}

			if endIdx > 0 {
				// Extract FIND arguments: FIND(arg1, arg2)
				findContent := findPart[5:endIdx] // Skip "FIND("
				findArgs := splitFunctionArgs(findContent)

				if len(findArgs) >= 2 {
					lookupCell := strings.TrimSpace(findArgs[0])
					sourceRange := strings.TrimSpace(findArgs[1])

					// Get column index
					col := getRangeColumnIndex(sourceRange)

					conditions = append(conditions, multiCondition{
						condType:    "find",
						sourceRange: sourceRange,
						sourceCol:   col,
					})
					lookupCells = append(lookupCells, lookupCell)
				}
			}
		}
	}

	return conditions, lookupCells
}

// splitByMultiply splits expression by * while respecting parentheses
func splitByMultiply(expr string) []string {
	var parts []string
	var current strings.Builder
	depth := 0

	for _, ch := range expr {
		switch ch {
		case '(':
			depth++
			current.WriteRune(ch)
		case ')':
			depth--
			current.WriteRune(ch)
		case '*':
			if depth == 0 {
				if current.Len() > 0 {
					parts = append(parts, current.String())
					current.Reset()
				}
			} else {
				current.WriteRune(ch)
			}
		default:
			current.WriteRune(ch)
		}
	}

	if current.Len() > 0 {
		parts = append(parts, current.String())
	}

	return parts
}

// getRangeColumnIndex extracts column index from range like "订单!B:B" or "订单!P:P"
func getRangeColumnIndex(rangeRef string) int {
	// Remove sheet name
	if idx := strings.Index(rangeRef, "!"); idx != -1 {
		rangeRef = rangeRef[idx+1:]
	}

	// Remove $ signs
	rangeRef = strings.ReplaceAll(rangeRef, "$", "")

	// Get column letter(s) - handle both "B:B" and "B" formats
	colPart := rangeRef
	if idx := strings.Index(rangeRef, ":"); idx != -1 {
		colPart = rangeRef[:idx]
	}

	// Extract only letters
	re := regexp.MustCompile(`[A-Za-z]+`)
	colLetters := re.FindString(colPart)
	if colLetters == "" {
		return -1
	}

	colNum, _ := ColumnNameToNumber(colLetters)
	return colNum - 1 // 0-based
}

// calculateIndexMatchMultiCondPatternWithCache calculates multi-condition INDEX-MATCH batch
func (f *File) calculateIndexMatchMultiCondPatternWithCache(pattern *indexMatchMultiCondPattern, worksheetCache *WorksheetCache) map[string]string {
	results := make(map[string]string)

	if len(pattern.formulas) == 0 {
		return results
	}

	startTime := time.Now()
	log.Printf("⚡ [MultiCond INDEX-MATCH] Starting batch calculation for %d formulas", len(pattern.formulas))

	// Get source sheet data
	sourceSheet := pattern.sourceSheet
	rows, err := f.GetRows(sourceSheet)
	if err != nil {
		log.Printf("❌ [MultiCond INDEX-MATCH] Failed to get rows from %s: %v", sourceSheet, err)
		return results
	}

	// Get result column index
	resultColIdx := getRangeColumnIndex(pattern.resultRange)
	if resultColIdx < 0 {
		log.Printf("❌ [MultiCond INDEX-MATCH] Invalid result range: %s", pattern.resultRange)
		return results
	}

	// Pre-build FIND index for string contains conditions
	// Map: searchString -> []rowIndex (rows where the string is found)
	findIndexes := make(map[int]map[string][]int) // conditionIndex -> searchStr -> rowIndexes

	for condIdx, cond := range pattern.conditions {
		if cond.condType == "find" && cond.sourceCol >= 0 {
			findIndexes[condIdx] = make(map[string][]int)

			// Build index for this column
			for rowIdx, row := range rows {
				if cond.sourceCol < len(row) {
					cellValue := row[cond.sourceCol]
					if cellValue != "" {
						// Store the cell value for later FIND matching
						// We'll build a reverse index: for each unique search pattern we see,
						// record which rows contain it
						findIndexes[condIdx]["__row_"+cellValue] = append(findIndexes[condIdx]["__row_"+cellValue], rowIdx)
					}
				}
			}
		}
	}

	// Calculate each formula
	for fullCell, info := range pattern.formulas {
		// Get lookup values
		lookupValues := make([]string, len(info.lookupCells))
		for i, lookupCell := range info.lookupCells {
			// Parse cell reference
			lookupSheet := info.sheet
			lookupCellRef := lookupCell
			if strings.Contains(lookupCell, "!") {
				parts := strings.SplitN(lookupCell, "!", 2)
				lookupSheet = strings.Trim(parts[0], "'")
				lookupCellRef = parts[1]
			}
			lookupCellRef = strings.ReplaceAll(lookupCellRef, "$", "")

			// Get value from cache or file
			lookupValues[i] = f.getCellValueOrCalcCache(lookupSheet, lookupCellRef, worksheetCache)
		}

		// Find first matching row
		matchedRow := -1
		for rowIdx, row := range rows {
			allMatch := true

			for condIdx, cond := range pattern.conditions {
				if condIdx >= len(lookupValues) {
					allMatch = false
					break
				}

				lookupVal := lookupValues[condIdx]
				if cond.sourceCol < 0 || cond.sourceCol >= len(row) {
					allMatch = false
					break
				}

				cellValue := row[cond.sourceCol]

				switch cond.condType {
				case "equal":
					// Exact match
					if cellValue != lookupVal {
						allMatch = false
					}
				case "find":
					// String contains (FIND returns position if found, error otherwise)
					// ISNUMBER(FIND(search, text)) is true if search is found in text
					if lookupVal == "" || !strings.Contains(cellValue, lookupVal) {
						allMatch = false
					}
				}

				if !allMatch {
					break
				}
			}

			if allMatch {
				matchedRow = rowIdx
				break // MATCH with 0 returns first match
			}
		}

		// Get result value
		if matchedRow >= 0 && matchedRow < len(rows) && resultColIdx < len(rows[matchedRow]) {
			results[fullCell] = rows[matchedRow][resultColIdx]
		} else {
			// No match found - use fallback value if available
			if info.fallbackValue != "" {
				// Fallback value can be a cell reference or a literal value
				// If it's a cell reference (contains no operators and looks like a cell), get its value
				fallback := info.fallbackValue
				if isCellReference(fallback) {
					// It's a cell reference like "AE3791"
					fallbackVal := f.getCellValueOrCalcCache(info.sheet, fallback, worksheetCache)
					results[fullCell] = fallbackVal
				} else {
					// It's a literal value
					results[fullCell] = fallback
				}
			} else {
				// No fallback specified, return empty string
				results[fullCell] = ""
			}
		}
	}

	log.Printf("✅ [MultiCond INDEX-MATCH] Completed %d formulas in %v", len(pattern.formulas), time.Since(startTime))

	return results
}

// batchCalculateIndexMatchMultiCondWithCache performs batch multi-condition INDEX-MATCH calculation
func (f *File) batchCalculateIndexMatchMultiCondWithCache(formulas map[string]string, worksheetCache *WorksheetCache) map[string]string {
	results := make(map[string]string)

	// Group formulas by pattern (same result range and condition structure)
	patterns := make(map[string]*indexMatchMultiCondPattern)

	for fullCell, formula := range formulas {
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}

		sheet := parts[0]
		cell := parts[1]

		pattern := f.extractIndexMatchMultiCondPattern(sheet, cell, formula)
		if pattern == nil {
			continue
		}

		// Create pattern key based on result range and condition structure
		key := pattern.resultRange
		for _, cond := range pattern.conditions {
			key += "|" + cond.condType + ":" + cond.sourceRange
		}

		if existing, exists := patterns[key]; exists {
			// Merge formulas into existing pattern
			for k, v := range pattern.formulas {
				existing.formulas[k] = v
			}
		} else {
			patterns[key] = pattern
		}
	}

	if len(patterns) > 0 {
		log.Printf("📊 [MultiCond INDEX-MATCH] Found %d unique patterns", len(patterns))
	}

	// Calculate each pattern
	for _, pattern := range patterns {
		patternResults := f.calculateIndexMatchMultiCondPatternWithCache(pattern, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	return results
}

// isMultiCondIndexMatchFormula checks if formula is a multi-condition INDEX-MATCH
func isMultiCondIndexMatchFormula(formula string) bool {
	upperFormula := strings.ToUpper(formula)
	return strings.Contains(upperFormula, "INDEX(") &&
		strings.Contains(formula, "MATCH(1,") &&
		(strings.Contains(formula, "*ISNUMBER(FIND(") || strings.Contains(formula, ")*ISNUMBER(FIND("))
}

// isCellReference checks if a string looks like a cell reference (e.g., "A1", "AE3791", "$B$5")
// Returns true if it's a valid cell reference without operators or functions
func isCellReference(s string) bool {
	// Remove $ signs
	s = strings.ReplaceAll(s, "$", "")

	// Empty or too short
	if len(s) < 2 {
		return false
	}

	// Check if it contains operators or parentheses (indicating it's an expression/function)
	if strings.ContainsAny(s, "+-*/()") {
		return false
	}

	// Basic pattern: letters followed by numbers
	// A simple heuristic: first character should be a letter
	if s[0] < 'A' || s[0] > 'Z' {
		return false
	}

	// Should have at least one digit
	hasDigit := false
	for _, ch := range s {
		if ch >= '0' && ch <= '9' {
			hasDigit = true
			break
		}
	}

	return hasDigit
}
