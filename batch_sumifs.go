package excelize

import (
	"fmt"
	"log"
	"runtime"
	"strconv"
	"strings"
	"sync"
)

func parseFormulaLiteralArg(expr string) (formulaArg, bool) {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return newStringFormulaArg(""), true
	}
	if len(expr) >= 2 && expr[0] == '"' && expr[len(expr)-1] == '"' {
		return newStringFormulaArg(strings.ReplaceAll(expr[1:len(expr)-1], `""`, `"`)), true
	}
	if strings.EqualFold(expr, "TRUE") || strings.EqualFold(expr, "FALSE") {
		return newBoolFormulaArg(strings.EqualFold(expr, "TRUE")), true
	}
	if strings.HasPrefix(expr, "#") {
		return newErrorFormulaArg(expr, expr), true
	}
	if num, err := strconv.ParseFloat(expr, 64); err == nil {
		return newNumberFormulaArg(num), true
	}
	return formulaArg{}, false
}

// getCellArgOrCalcCache retrieves a typed cell value from worksheetCache or the worksheet.
func (f *File) getCellArgOrCalcCache(sheet, cell string, worksheetCache *WorksheetCache) formulaArg {
	if worksheetCache != nil {
		if argValue, ok := worksheetCache.Get(sheet, cell); ok {
			return argValue
		}
	}
	value, _ := f.GetCellValue(sheet, cell, Options{RawCellValue: true})
	cellType, _ := f.GetCellType(sheet, cell)
	return inferCellValueType(value, cellType)
}

// getCellValueOrCalcCache retrieves cell value from the unified worksheetCache
// This ensures all reads get the latest values (both original and calculated)
func (f *File) getCellValueOrCalcCache(sheet, cell string, worksheetCache *WorksheetCache) string {
	argValue := f.getCellArgOrCalcCache(sheet, cell, worksheetCache)
	val := argValue.Value()
	// For ArgNumber, Value() uses %g which may produce scientific
	// notation (e.g. "1.2677910539e+10") for large numbers. The
	// batch result maps are keyed by GetRows raw text (e.g.
	// "12677910539"). Use the full-precision decimal form so that
	// map lookups match.
	if argValue.Type == ArgNumber && !argValue.Boolean &&
		strings.ContainsAny(val, "eE") {
		val = strconv.FormatFloat(argValue.Number, 'f', -1, 64)
	}
	return val
}

// resolveCriteriaValue resolves a SUMIFS criteria argument to its string value.
// The criteria may be a cell reference (e.g. "B2", "$A$1") or a literal value
// (e.g. `"-"`, `"abc"`, `123`). Quoted string literals are unquoted and returned
// directly; numeric literals are returned as-is. Cell references are looked up
// via getCellValueOrCalcCache.
func (f *File) resolveCriteriaValue(sheet, criteria string, worksheetCache *WorksheetCache) string {
	// Quoted string literal: "-", "abc", etc.
	if len(criteria) >= 2 && criteria[0] == '"' && criteria[len(criteria)-1] == '"' {
		return criteria[1 : len(criteria)-1]
	}
	// Numeric literal (doesn't start with a letter, not a cell reference)
	if len(criteria) > 0 && criteria[0] >= '0' && criteria[0] <= '9' {
		if _, err := strconv.ParseFloat(criteria, 64); err == nil {
			return criteria
		}
	}
	// Cell reference: look up the value
	return f.getCellValueOrCalcCache(sheet, criteria, worksheetCache)
}

// sumifs2DPattern represents a batch SUMIFS pattern where formulas form a 2D matrix
type sumifs2DPattern struct {
	// Common ranges (same for all formulas)
	sumRangeRef       string
	criteriaRange1Ref string
	criteriaRange2Ref string

	// Formula mapping: cell -> (criteria1Cell, criteria2Cell)
	formulas map[string]*sumifs2DFormula
}

// sumifs2DFormula represents a single SUMIFS formula in the batch
type sumifs2DFormula struct {
	cell          string
	sheet         string
	criteria1Cell string // e.g., "$A2"
	criteria2Cell string // e.g., "B$1"
}

// sumifs1DPattern represents a batch SUMIFS pattern with only 1 criterion
type sumifs1DPattern struct {
	// Common ranges (same for all formulas)
	sumRangeRef       string
	criteriaRange1Ref string

	// Formula mapping: cell -> criteria1Cell
	formulas map[string]*sumifs1DFormula
}

// sumifs1DFormula represents a single SUMIFS formula with 1 criterion
type sumifs1DFormula struct {
	cell          string
	sheet         string
	criteria1Cell string // e.g., "$A2"
}

// averageifs2DPattern represents a batch AVERAGEIFS pattern
type averageifs2DPattern struct {
	// Common ranges (same for all formulas)
	averageRangeRef   string
	criteriaRange1Ref string
	criteriaRange2Ref string

	// Formula mapping: cell -> (criteria1Cell, criteria2Cell)
	formulas map[string]*averageifs2DFormula
}

// averageifs2DFormula represents a single AVERAGEIFS formula in the batch
type averageifs2DFormula struct {
	cell          string
	sheet         string
	criteria1Cell string
	criteria2Cell string
}

// detectAndCalculateBatchSUMIFS detects and calculates batch SUMIFS patterns
// Returns map of cell -> calculated value for batch-processed formulas
func (f *File) detectAndCalculateBatchSUMIFS() map[string]float64 {
	results := make(map[string]float64)

	// First, detect and calculate SUMPRODUCT patterns
	sumproductResults := f.detectAndCalculateBatchSUMPRODUCT()
	for cell, value := range sumproductResults {
		results[cell] = value
	}

	// Detect and calculate AVERAGE(OFFSET) patterns
	avgOffsetResults := f.detectAndCalculateBatchAverageOffset()
	for cell, value := range avgOffsetResults {
		results[cell] = value
	}

	// Scan all sheets to find SUMIFS formulas
	// Strategy: Sample cells to detect patterns, then batch calculate
	sheetList := f.GetSheetList()

	for _, sheet := range sheetList {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}

		// Collect SUMIFS formulas from this sheet
		sumifsFormulas := make(map[string]string)
		averageifsFormulas := make(map[string]string)

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				if cell.F != nil {
					formula := cell.F.Content
					// Handle shared formulas
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}

					// Extract SUMIFS from formula (even if nested in IF, IFERROR, etc.)
					if sumifsExpr := extractSUMIFSFromFormula(formula); sumifsExpr != "" {
						fullCell := sheet + "!" + cell.R
						sumifsFormulas[fullCell] = sumifsExpr
					}

					// Extract AVERAGEIFS from formula
					if averageifsExpr := extractAVERAGEIFSFromFormula(formula); averageifsExpr != "" {
						fullCell := sheet + "!" + cell.R
						averageifsFormulas[fullCell] = averageifsExpr
					}
				}
			}
		}

		// Group SUMIFS formulas by pattern for this sheet
		if len(sumifsFormulas) >= 10 {
			// Try 1D patterns (1 criterion) first
			patterns1D := f.groupSUMIFS1DByPattern(sumifsFormulas)
			for _, pattern := range patterns1D {
				if len(pattern.formulas) >= 10 {
					batchResults := f.calculateSUMIFS1DPattern(pattern)
					for cell, value := range batchResults {
						results[cell] = value
					}
				}
			}

			// Then try 2D patterns (2 criteria)
			patterns2D := f.groupSUMIFSByPattern(sumifsFormulas)
			for _, pattern := range patterns2D {
				if len(pattern.formulas) >= 10 {
					batchResults := f.calculateSUMIFS2DPattern(pattern)
					for cell, value := range batchResults {
						results[cell] = value
					}
				}
			}
		}

		// Group AVERAGEIFS formulas by pattern for this sheet
		if len(averageifsFormulas) >= 10 {
			patterns := f.groupAVERAGEIFSByPattern(averageifsFormulas)

			// Calculate each pattern
			for _, pattern := range patterns {
				if len(pattern.formulas) >= 10 {
					batchResults := f.calculateAVERAGEIFS2DPattern(pattern)
					for cell, value := range batchResults {
						results[cell] = value
					}
				}
			}
		}
	}

	return results
}

// extractSUMIFSFromFormula extracts SUMIFS expression from a formula (even if nested)
// Examples:
//   - "SUMIFS(...)" -> "SUMIFS(...)"
//   - "=IF(A1=0,"x",SUMIFS(...))" -> "SUMIFS(...)"
//   - "=$E2-G2+SUMIFS(...)" -> "SUMIFS(...)"
func extractSUMIFSFromFormula(formula string) string {
	// Find "SUMIFS(" in the formula
	idx := strings.Index(formula, "SUMIFS(")
	if idx == -1 {
		return ""
	}

	// Extract the complete SUMIFS(...) expression
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

// extractAVERAGEIFSFromFormula extracts AVERAGEIFS expression from a formula (even if nested)
// Examples:
//   - "AVERAGEIFS(...)" -> "AVERAGEIFS(...)"
//   - "=IFERROR(AVERAGEIFS(...))" -> "AVERAGEIFS(...)"
func extractAVERAGEIFSFromFormula(formula string) string {
	// Find "AVERAGEIFS(" in the formula
	idx := strings.Index(formula, "AVERAGEIFS(")
	if idx == -1 {
		return ""
	}

	// Extract the complete AVERAGEIFS(...) expression
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

// ExtractSUMIFSFromFormulaExport is exported for testing
func ExtractSUMIFSFromFormulaExport(formula string) string {
	return extractSUMIFSFromFormula(formula)
}

// batchCalculateSUMIFSWithCache performs batch SUMIFS calculation using worksheetCache
// This is the REAL solution - we modify batch calculation to use unified worksheetCache
// Supports both 1D SUMIFS (1 criterion) and 2D SUMIFS (2 criteria)
func (f *File) batchCalculateSUMIFSWithCache(formulas map[string]string, worksheetCache *WorksheetCache) map[string]string {
	results := make(map[string]string)

	// Separate formulas that can be 2D vs 1D
	remaining := make(map[string]string)

	// Group formulas by pattern (same logic as batchCalculateSUMIFS)
	patterns2D := make(map[string]*sumifs2DPattern)

	for fullCell, formula := range formulas {
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}

		sheet := parts[0]
		cell := parts[1]

		// Try to extract 2D SUMIFS pattern first
		pattern := f.extractSUMIFS2DPattern(sheet, cell, formula)
		if pattern != nil {
			key := pattern.sumRangeRef + "|" + pattern.criteriaRange1Ref + "|" + pattern.criteriaRange2Ref
			if _, exists := patterns2D[key]; !exists {
				patterns2D[key] = pattern
			} else {
				for k, v := range pattern.formulas {
					patterns2D[key].formulas[k] = v
				}
			}
		} else {
			// Not a 2D pattern, try 1D later
			remaining[fullCell] = formula
		}
	}

	// Calculate each 2D pattern using worksheetCache
	for _, pattern := range patterns2D {
		patternResults := f.calculateSUMIFS2DPatternWithCache(pattern, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = fmt.Sprintf("%v", value)
		}
	}

	// Now handle 1D SUMIFS
	if len(remaining) > 0 {
		patterns1D := f.groupSUMIFS1DByPattern(remaining)
		for _, pattern := range patterns1D {
			if len(pattern.formulas) >= 10 {
				patternResults := f.calculateSUMIFS1DPatternWithCache(pattern, worksheetCache)
				for cell, value := range patternResults {
					results[cell] = fmt.Sprintf("%v", value)
				}
			}
		}
	}

	return results
}

// calculateSUMIFS1DPatternWithCache calculates 1D SUMIFS using worksheetCache
func (f *File) calculateSUMIFS1DPatternWithCache(pattern *sumifs1DPattern, worksheetCache *WorksheetCache) map[string]float64 {
	sourceSheet := extractSheetName(pattern.sumRangeRef)
	if sourceSheet == "" {
		return map[string]float64{}
	}

	sumCol := extractColumnFromRange(pattern.sumRangeRef)
	criteria1Col := extractColumnFromRange(pattern.criteriaRange1Ref)

	if sumCol == "" || criteria1Col == "" {
		return map[string]float64{}
	}

	// Read source data directly from file
	rows, err := f.GetRows(sourceSheet, Options{RawCellValue: true})
	if err != nil {
		return map[string]float64{}
	}

	// Build 1D result map: criteria1Value -> sum
	sumColIdx, _ := ColumnNameToNumber(sumCol)
	sumColIdx--
	criteria1ColIdx, _ := ColumnNameToNumber(criteria1Col)
	criteria1ColIdx--

	resultMap := make(map[string]float64)
	for _, row := range rows {
		if criteria1ColIdx >= len(row) || sumColIdx >= len(row) {
			continue
		}

		c1 := row[criteria1ColIdx]
		sumVal := row[sumColIdx]

		if sumVal != "" {
			if v, err := strconv.ParseFloat(sumVal, 64); err == nil {
				resultMap[c1] += v
			}
		}
	}

	// Fill results for all formulas
	results := make(map[string]float64)
	for fullCell, info := range pattern.formulas {
		criteria1Cell := strings.ReplaceAll(info.criteria1Cell, "$", "")

		// 解析 criteria 值：可能是单元格引用（如 B2）或字面量（如 "-"）
		c1 := f.resolveCriteriaValue(info.sheet, criteria1Cell, worksheetCache)

		if val, ok := resultMap[c1]; ok {
			results[fullCell] = val
		} else {
			results[fullCell] = 0
		}
	}

	return results
}

// calculateSUMIFS2DPatternWithCache calculates SUMIFS using worksheetCache
func (f *File) calculateSUMIFS2DPatternWithCache(pattern *sumifs2DPattern, worksheetCache *WorksheetCache) map[string]float64 {
	sourceSheet := extractSheetName(pattern.sumRangeRef)
	if sourceSheet == "" {
		return map[string]float64{}
	}

	sumCol := extractColumnFromRange(pattern.sumRangeRef)
	criteria1Col := extractColumnFromRange(pattern.criteriaRange1Ref)
	criteria2Col := extractColumnFromRange(pattern.criteriaRange2Ref)

	if sumCol == "" || criteria1Col == "" || criteria2Col == "" {
		return map[string]float64{}
	}

	// 直接从文件读取原始数据
	// 注意：worksheetCache 只存储计算结果，不存储原始数据
	rows, err := f.GetRows(sourceSheet, Options{RawCellValue: true})
	if err != nil {
		return map[string]float64{}
	}

	// Build result map by scanning once
	resultMap := f.scanRowsAndBuildResultMap(sourceSheet, rows, sumCol, criteria1Col, criteria2Col)

	// Fill results for all formulas
	results := make(map[string]float64)
	for fullCell, info := range pattern.formulas {
		criteria1Cell := strings.ReplaceAll(info.criteria1Cell, "$", "")
		criteria2Cell := strings.ReplaceAll(info.criteria2Cell, "$", "")

		// 解析 criteria 值：可能是单元格引用（如 B2）或字面量（如 "-"）
		c1 := f.resolveCriteriaValue(info.sheet, criteria1Cell, worksheetCache)
		c2 := f.resolveCriteriaValue(info.sheet, criteria2Cell, worksheetCache)

		if resultMap[c1] != nil {
			if val, ok := resultMap[c1][c2]; ok {
				results[fullCell] = val
			} else {
				results[fullCell] = 0
			}
		} else {
			results[fullCell] = 0
		}
	}

	return results
}

// TestExtractSUMIFS2DPattern is exported for testing
func TestExtractSUMIFS2DPattern(f *File, sheet, cell, formula string) *Sumifs2DPatternExport {
	pattern := f.extractSUMIFS2DPattern(sheet, cell, formula)
	if pattern == nil {
		return nil
	}
	return &Sumifs2DPatternExport{
		SumRangeRef:       pattern.sumRangeRef,
		CriteriaRange1Ref: pattern.criteriaRange1Ref,
		CriteriaRange2Ref: pattern.criteriaRange2Ref,
	}
}

// Sumifs2DPatternExport is exported for testing
type Sumifs2DPatternExport struct {
	SumRangeRef       string
	CriteriaRange1Ref string
	CriteriaRange2Ref string
}

// groupSUMIFS1DByPattern groups SUMIFS formulas with 1 criterion by their pattern
func (f *File) groupSUMIFS1DByPattern(formulas map[string]string) []*sumifs1DPattern {
	patterns := make(map[string]*sumifs1DPattern)

	for fullCell, formula := range formulas {
		// Parse fullCell as "sheet!cell"
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}
		sheet, cell := parts[0], parts[1]

		// Extract 1D pattern (1 criterion)
		pattern := f.extractSUMIFS1DPattern(sheet, cell, formula)
		if pattern == nil {
			continue
		}

		// Group by common ranges
		key := pattern.sumRangeRef + "|" + pattern.criteriaRange1Ref
		if patterns[key] == nil {
			patterns[key] = pattern
		} else {
			// Merge formulas
			for c, info := range pattern.formulas {
				patterns[key].formulas[c] = info
			}
		}
	}

	// Convert to slice
	var result []*sumifs1DPattern
	for _, p := range patterns {
		result = append(result, p)
	}
	return result
}

// extractSUMIFS1DPattern extracts 1D pattern from SUMIFS formula with 1 criterion
func (f *File) extractSUMIFS1DPattern(sheet, cell, formula string) *sumifs1DPattern {
	// SUMIFS(sum_range, criteria_range1, criteria1)

	// Remove "SUMIFS(" and trailing ")"
	if len(formula) < 8 || formula[:7] != "SUMIFS(" {
		return nil
	}

	inner := formula[7 : len(formula)-1]
	parts := splitFormulaArgs(inner)

	// Check if it's a 1-criterion SUMIFS (3 parts)
	if len(parts) != 3 {
		return nil
	}

	sumRange := strings.TrimSpace(parts[0])
	criteriaRange1 := strings.TrimSpace(parts[1])
	criteria1Cell := strings.TrimSpace(parts[2])

	// Check if sum_range and criteria_range are external references (contain '!')
	if !strings.Contains(sumRange, "!") {
		return nil
	}
	if !strings.Contains(criteriaRange1, "!") {
		return nil
	}

	// Check if criteria is a cell reference (not external)
	if strings.Contains(criteria1Cell, "!") {
		return nil
	}

	pattern := &sumifs1DPattern{
		sumRangeRef:       sumRange,
		criteriaRange1Ref: criteriaRange1,
		formulas:          make(map[string]*sumifs1DFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &sumifs1DFormula{
		cell:          cell,
		sheet:         sheet,
		criteria1Cell: criteria1Cell,
	}

	return pattern
}

// calculateSUMIFS1DPattern calculates a batch of SUMIFS formulas with 1 criterion
func (f *File) calculateSUMIFS1DPattern(pattern *sumifs1DPattern) map[string]float64 {
	// Extract sheet from range reference
	sourceSheet := extractSheetName(pattern.sumRangeRef)
	if sourceSheet == "" {
		return map[string]float64{}
	}

	// Extract column letters from range references
	sumCol := extractColumnFromRange(pattern.sumRangeRef)
	criteria1Col := extractColumnFromRange(pattern.criteriaRange1Ref)

	if sumCol == "" || criteria1Col == "" {
		return map[string]float64{}
	}

	// Read all rows from the source sheet
	rows, err := f.GetRows(sourceSheet)
	if err != nil || len(rows) == 0 {
		return map[string]float64{}
	}

	// Build result map by scanning once
	resultMap := f.scanRowsAndBuild1DResultMap(sourceSheet, rows, sumCol, criteria1Col)

	// Fill results for all formulas
	results := make(map[string]float64)
	for fullCell, info := range pattern.formulas {
		// Remove $ from cell references before calling GetCellValue
		criteria1Cell := strings.ReplaceAll(info.criteria1Cell, "$", "")

		// Note: This function doesn't have worksheetCache, so use direct GetCellValue as fallback
		c1, _ := f.GetCellValue(info.sheet, criteria1Cell)

		if val, ok := resultMap[c1]; ok {
			results[fullCell] = val
		} else {
			results[fullCell] = 0
		}
	}

	return results
}

// scanRowsAndBuild1DResultMap scans rows and builds 1D result map (single criterion)
func (f *File) scanRowsAndBuild1DResultMap(
	sheet string,
	rows [][]string,
	sumCol, criteria1Col string,
) map[string]float64 {

	if len(rows) == 0 {
		return nil
	}

	// Convert column letters to indices
	sumColIdx, _ := ColumnNameToNumber(sumCol)
	criteria1ColIdx, _ := ColumnNameToNumber(criteria1Col)

	sumColIdx--       // Convert to 0-based
	criteria1ColIdx-- // Convert to 0-based

	numWorkers := runtime.NumCPU()
	rowCount := len(rows)

	if numWorkers > rowCount {
		numWorkers = rowCount
	}
	if numWorkers < 1 {
		numWorkers = 1
	}

	rowsPerWorker := (rowCount + numWorkers - 1) / numWorkers

	// Worker results
	type workerResult struct {
		data map[string]float64
	}
	results := make([]workerResult, numWorkers)
	var wg sync.WaitGroup

	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()

			start := workerID * rowsPerWorker
			end := start + rowsPerWorker
			if end > rowCount {
				end = rowCount
			}

			localMap := make(map[string]float64)

			for rowIdx := start; rowIdx < end; rowIdx++ {
				row := rows[rowIdx]

				// Extract values from columns
				var c1, sumVal string

				if criteria1ColIdx < len(row) {
					c1 = row[criteria1ColIdx]
				}
				if sumColIdx < len(row) {
					sumVal = row[sumColIdx]
				}

				if c1 == "" || sumVal == "" {
					continue
				}

				// Convert sumVal to number
				var num float64
				_, err := fmt.Sscanf(sumVal, "%f", &num)
				if err != nil {
					continue
				}

				// Accumulate
				localMap[c1] += num
			}

			results[workerID] = workerResult{data: localMap}
		}(i)
	}

	wg.Wait()

	// Merge results
	finalMap := make(map[string]float64)
	for _, r := range results {
		for c1, sum := range r.data {
			finalMap[c1] += sum
		}
	}

	return finalMap
}

// groupSUMIFSByPattern groups SUMIFS formulas by their pattern
func (f *File) groupSUMIFSByPattern(formulas map[string]string) []*sumifs2DPattern {
	patterns := make(map[string]*sumifs2DPattern)

	for fullCell, formula := range formulas {
		// Parse fullCell as "sheet!cell"
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}
		sheet, cell := parts[0], parts[1]

		// Simple pattern extraction:
		// SUMIFS('sheet'!$H:$H,'sheet'!$D:$D,$A2,'sheet'!$A:$A,B$1)
		// Extract: sum_range, criteria_range1, criteria1_cell, criteria_range2, criteria2_cell

		pattern := f.extractSUMIFS2DPattern(sheet, cell, formula)
		if pattern == nil {
			continue
		}

		// Group by common ranges
		key := pattern.sumRangeRef + "|" + pattern.criteriaRange1Ref + "|" + pattern.criteriaRange2Ref
		if patterns[key] == nil {
			patterns[key] = pattern
		} else {
			// Merge formulas
			for c, info := range pattern.formulas {
				patterns[key].formulas[c] = info
			}
		}
	}

	// Convert to slice
	var result []*sumifs2DPattern
	for _, p := range patterns {
		result = append(result, p)
	}
	return result
}

// extractSUMIFS2DPattern extracts 2D pattern from SUMIFS formula
func (f *File) extractSUMIFS2DPattern(sheet, cell, formula string) *sumifs2DPattern {
	// Simple parsing: split by comma (simplified - doesn't handle nested functions)
	// SUMIFS(sum_range,criteria_range1,criteria1,criteria_range2,criteria2,...)

	// Remove "SUMIFS(" and trailing ")"
	if len(formula) < 8 || formula[:7] != "SUMIFS(" {
		return nil
	}

	inner := formula[7 : len(formula)-1]
	parts := splitFormulaArgs(inner)

	if len(parts) != 5 { // We only support exactly 2 criteria for now
		return nil
	}

	sumRange := strings.TrimSpace(parts[0])
	criteriaRange1 := strings.TrimSpace(parts[1])
	criteria1Cell := strings.TrimSpace(parts[2])
	criteriaRange2 := strings.TrimSpace(parts[3])
	criteria2Cell := strings.TrimSpace(parts[4])

	// Check if ranges are external references (contain '!')
	if !strings.Contains(sumRange, "!") {
		return nil
	}
	if !strings.Contains(criteriaRange1, "!") {
		return nil
	}
	if !strings.Contains(criteriaRange2, "!") {
		return nil
	}

	// Check if criteria are cell references (not external)
	if strings.Contains(criteria1Cell, "!") {
		return nil
	}
	if strings.Contains(criteria2Cell, "!") {
		return nil
	}

	pattern := &sumifs2DPattern{
		sumRangeRef:       sumRange,
		criteriaRange1Ref: criteriaRange1,
		criteriaRange2Ref: criteriaRange2,
		formulas:          make(map[string]*sumifs2DFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &sumifs2DFormula{
		cell:          cell,
		sheet:         sheet,
		criteria1Cell: criteria1Cell,
		criteria2Cell: criteria2Cell,
	}

	return pattern
}

// splitFormulaArgs splits formula arguments by comma (simplified version)
func splitFormulaArgs(s string) []string {
	var result []string
	var current strings.Builder
	depth := 0
	inQuote := false

	for i := 0; i < len(s); i++ {
		ch := s[i]

		switch ch {
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
		case '"', '\'':
			inQuote = !inQuote
			current.WriteByte(ch)
		case ',':
			if depth == 0 && !inQuote {
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

// calculateSUMIFS2DPattern calculates a batch of SUMIFS formulas
func (f *File) calculateSUMIFS2DPattern(pattern *sumifs2DPattern) map[string]float64 {
	// Simplified version: directly read Excel data using GetRows
	// Extract sheet from range reference
	sourceSheet := extractSheetName(pattern.sumRangeRef)
	if sourceSheet == "" {
		return map[string]float64{} // Return empty map instead of nil
	}

	// Extract column letters from range references
	// e.g., 'sheet'!$H:$H -> H
	sumCol := extractColumnFromRange(pattern.sumRangeRef)
	criteria1Col := extractColumnFromRange(pattern.criteriaRange1Ref)
	criteria2Col := extractColumnFromRange(pattern.criteriaRange2Ref)

	if sumCol == "" || criteria1Col == "" || criteria2Col == "" {
		return map[string]float64{} // Return empty map instead of nil
	}

	// Read all rows from the source sheet
	rows, err := f.GetRows(sourceSheet)
	if err != nil || len(rows) == 0 {
		return map[string]float64{} // Return empty map instead of nil
	}

	// Build result map by scanning once
	resultMap := f.scanRowsAndBuildResultMap(sourceSheet, rows, sumCol, criteria1Col, criteria2Col)

	// Fill results for all formulas
	results := make(map[string]float64)
	for fullCell, info := range pattern.formulas {
		// Remove $ from cell references before calling GetCellValue
		criteria1Cell := strings.ReplaceAll(info.criteria1Cell, "$", "")
		criteria2Cell := strings.ReplaceAll(info.criteria2Cell, "$", "")

		// Note: This function doesn't have worksheetCache, so use direct GetCellValue as fallback
		c1, _ := f.GetCellValue(info.sheet, criteria1Cell)
		c2, _ := f.GetCellValue(info.sheet, criteria2Cell)

		if resultMap[c1] != nil {
			if val, ok := resultMap[c1][c2]; ok {
				results[fullCell] = val
			} else {
				results[fullCell] = 0 // Add zero result
			}
		} else {
			results[fullCell] = 0 // Add zero result
		}
	}

	return results
}

// extractSheetName extracts sheet name from range reference
// e.g., 'sheet'!$H:$H -> sheet
func extractSheetName(rangeRef string) string {
	parts := strings.Split(rangeRef, "!")
	if len(parts) != 2 {
		return ""
	}
	return strings.Trim(parts[0], "'")
}

// extractColumnFromRange extracts column letter from range reference
// e.g., 'sheet'!$H:$H -> H
func extractColumnFromRange(rangeRef string) string {
	parts := strings.Split(rangeRef, "!")
	if len(parts) != 2 {
		return ""
	}

	ref := parts[1]
	// Remove $ and :$H part
	ref = strings.ReplaceAll(ref, "$", "")
	if idx := strings.Index(ref, ":"); idx != -1 {
		ref = ref[:idx]
	}

	return ref
}

// scanRowsAndBuildResultMap scans rows and builds result map concurrently
func (f *File) scanRowsAndBuildResultMap(
	sheet string,
	rows [][]string,
	sumCol, criteria1Col, criteria2Col string,
) map[string]map[string]float64 {

	if len(rows) == 0 {
		return nil
	}

	// Convert column letters to indices
	sumColIdx, _ := ColumnNameToNumber(sumCol)
	criteria1ColIdx, _ := ColumnNameToNumber(criteria1Col)
	criteria2ColIdx, _ := ColumnNameToNumber(criteria2Col)

	sumColIdx--       // Convert to 0-based
	criteria1ColIdx-- // Convert to 0-based
	criteria2ColIdx-- // Convert to 0-based

	numWorkers := runtime.NumCPU()
	rowCount := len(rows)

	if numWorkers > rowCount {
		numWorkers = rowCount
	}
	if numWorkers < 1 {
		numWorkers = 1
	}

	rowsPerWorker := (rowCount + numWorkers - 1) / numWorkers

	// Worker results
	type workerResult struct {
		data map[string]map[string]float64
	}
	results := make([]workerResult, numWorkers)
	var wg sync.WaitGroup

	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()

			start := workerID * rowsPerWorker
			end := start + rowsPerWorker
			if end > rowCount {
				end = rowCount
			}

			localMap := make(map[string]map[string]float64)

			for rowIdx := start; rowIdx < end; rowIdx++ {
				row := rows[rowIdx]

				// Extract values from columns
				var c1, c2, sumVal string

				if criteria1ColIdx < len(row) {
					c1 = row[criteria1ColIdx]
				}
				if criteria2ColIdx < len(row) {
					c2 = row[criteria2ColIdx]
				}
				if sumColIdx < len(row) {
					sumVal = row[sumColIdx]
				}

				if c1 == "" || c2 == "" || sumVal == "" {
					continue
				}

				// Convert sumVal to number
				var num float64
				_, err := fmt.Sscanf(sumVal, "%f", &num)
				if err != nil {
					continue
				}

				// Accumulate
				if localMap[c1] == nil {
					localMap[c1] = make(map[string]float64)
				}
				localMap[c1][c2] += num
			}

			results[workerID] = workerResult{data: localMap}
		}(i)
	}

	wg.Wait()

	// Merge results
	finalMap := make(map[string]map[string]float64)
	for _, r := range results {
		for c1, m := range r.data {
			if finalMap[c1] == nil {
				finalMap[c1] = make(map[string]float64)
			}
			for c2, sum := range m {
				finalMap[c1][c2] += sum
			}
		}
	}

	return finalMap
}

// detectAndCalculateBatchINDEX detects and batch calculates INDEX formulas
// Pattern: INDEX($K{row}:$AAC{row}, offset) where offset is constant
func (f *File) detectAndCalculateBatchINDEX() map[string]float64 {
	results := make(map[string]float64)

	sheetList := f.GetSheetList()

	for _, sheet := range sheetList {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}

		// Collect INDEX formulas pattern: INDEX($K{row}:${EndCol}{row}, offset)
		indexFormulas := make(map[string]string) // fullCell -> formula

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				if cell.F != nil {
					formula := cell.F.Content
					// Handle shared formulas
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}

					// Check if formula contains INDEX with same-row range pattern
					// Pattern: INDEX($K{row}:$AAC{row}, ...)
					if strings.Contains(formula, "INDEX(") && strings.Contains(formula, ":") {
						fullCell := sheet + "!" + cell.R
						indexFormulas[fullCell] = formula
					}
				}
			}
		}

		// If we have at least 10 INDEX formulas, try batch optimization
		if len(indexFormulas) >= 10 {
			batchResults := f.calculateIndexFormulas(sheet, indexFormulas)
			for cell, value := range batchResults {
				results[cell] = value
			}
		}
	}

	return results
}

// calculateIndexFormulas batch calculates INDEX formulas for a sheet
func (f *File) calculateIndexFormulas(sheet string, formulas map[string]string) map[string]float64 {
	results := make(map[string]float64)

	// For now, just calculate them individually but with row caching
	// The key optimization: cache entire rows to avoid repeated GetRows calls
	rowCache := make(map[int][]string) // rowNum -> row data

	rows, err := f.GetRows(sheet)
	if err != nil {
		return results
	}

	// Pre-cache all rows
	for i, row := range rows {
		rowCache[i+1] = row // Excel rows are 1-indexed
	}

	calculated := 0
	for fullCell := range formulas {
		// Parse cell reference
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}
		cellRef := parts[1]

		// Try to calculate using CalcCellValue (with row cache already loaded)
		// This is faster than repeated GetCellValue calls
		value, err := f.CalcCellValue(sheet, cellRef)
		if err == nil {
			var numValue float64
			_, parseErr := fmt.Sscanf(value, "%f", &numValue)
			if parseErr == nil {
				results[fullCell] = numValue
				calculated++
			}
		}
	}

	if calculated > 0 {
		log.Printf("  ⚡ [INDEX Batch] Calculated %d INDEX formulas in sheet '%s'", calculated, sheet)
	}

	return results
}

// groupAVERAGEIFSByPattern groups AVERAGEIFS formulas by their pattern
func (f *File) groupAVERAGEIFSByPattern(formulas map[string]string) []*averageifs2DPattern {
	patterns := make(map[string]*averageifs2DPattern)

	for fullCell, formula := range formulas {
		// Parse fullCell as "sheet!cell"
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}
		sheet, cell := parts[0], parts[1]

		pattern := f.extractAVERAGEIFS2DPattern(sheet, cell, formula)
		if pattern == nil {
			continue
		}

		// Group by common ranges
		key := pattern.averageRangeRef + "|" + pattern.criteriaRange1Ref + "|" + pattern.criteriaRange2Ref
		if patterns[key] == nil {
			patterns[key] = pattern
		} else {
			// Merge formulas
			for c, info := range pattern.formulas {
				patterns[key].formulas[c] = info
			}
		}
	}

	// Convert to slice
	var result []*averageifs2DPattern
	for _, p := range patterns {
		result = append(result, p)
	}
	return result
}

// extractAVERAGEIFS2DPattern extracts 2D pattern from AVERAGEIFS formula
func (f *File) extractAVERAGEIFS2DPattern(sheet, cell, formula string) *averageifs2DPattern {
	// AVERAGEIFS(average_range,criteria_range1,criteria1,criteria_range2,criteria2,...)

	// Remove "AVERAGEIFS(" and trailing ")"
	if len(formula) < 13 || formula[:11] != "AVERAGEIFS(" {
		return nil
	}

	inner := formula[11 : len(formula)-1]
	parts := splitFormulaArgs(inner)

	if len(parts) != 5 { // We only support exactly 2 criteria for now
		return nil
	}

	averageRange := strings.TrimSpace(parts[0])
	criteriaRange1 := strings.TrimSpace(parts[1])
	criteria1Cell := strings.TrimSpace(parts[2])
	criteriaRange2 := strings.TrimSpace(parts[3])
	criteria2Cell := strings.TrimSpace(parts[4])

	// Check if ranges are external references (contain '!')
	if !strings.Contains(averageRange, "!") {
		return nil
	}
	if !strings.Contains(criteriaRange1, "!") {
		return nil
	}
	if !strings.Contains(criteriaRange2, "!") {
		return nil
	}

	// Check if criteria are cell references (not external)
	if strings.Contains(criteria1Cell, "!") {
		return nil
	}
	if strings.Contains(criteria2Cell, "!") {
		return nil
	}

	pattern := &averageifs2DPattern{
		averageRangeRef:   averageRange,
		criteriaRange1Ref: criteriaRange1,
		criteriaRange2Ref: criteriaRange2,
		formulas:          make(map[string]*averageifs2DFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &averageifs2DFormula{
		cell:          cell,
		sheet:         sheet,
		criteria1Cell: criteria1Cell,
		criteria2Cell: criteria2Cell,
	}

	return pattern
}

// calculateAVERAGEIFS2DPattern calculates a batch of AVERAGEIFS formulas
func (f *File) calculateAVERAGEIFS2DPattern(pattern *averageifs2DPattern) map[string]float64 {
	// Extract sheet from range reference
	sourceSheet := extractSheetName(pattern.averageRangeRef)
	if sourceSheet == "" {
		return map[string]float64{}
	}

	// Extract column letters from range references
	averageCol := extractColumnFromRange(pattern.averageRangeRef)
	criteria1Col := extractColumnFromRange(pattern.criteriaRange1Ref)
	criteria2Col := extractColumnFromRange(pattern.criteriaRange2Ref)

	if averageCol == "" || criteria1Col == "" || criteria2Col == "" {
		return map[string]float64{}
	}

	// Read all rows from the source sheet
	rows, err := f.GetRows(sourceSheet)
	if err != nil || len(rows) == 0 {
		return map[string]float64{}
	}

	// Build result map by scanning once (sum and count for average)
	resultMap := f.scanRowsAndBuildAverageMap(sourceSheet, rows, averageCol, criteria1Col, criteria2Col)

	// Fill results for all formulas
	results := make(map[string]float64)
	for fullCell, info := range pattern.formulas {
		// Remove $ from cell references before calling GetCellValue
		criteria1Cell := strings.ReplaceAll(info.criteria1Cell, "$", "")
		criteria2Cell := strings.ReplaceAll(info.criteria2Cell, "$", "")

		// Note: This function doesn't have worksheetCache, so use direct GetCellValue as fallback
		c1, _ := f.GetCellValue(info.sheet, criteria1Cell)
		c2, _ := f.GetCellValue(info.sheet, criteria2Cell)

		if resultMap[c1] != nil {
			if avgData, ok := resultMap[c1][c2]; ok {
				if avgData.count > 0 {
					results[fullCell] = avgData.sum / float64(avgData.count)
				} else {
					results[fullCell] = 0
				}
			} else {
				results[fullCell] = 0
			}
		} else {
			results[fullCell] = 0
		}
	}

	return results
}

// avgData holds sum and count for calculating average
type avgData struct {
	sum   float64
	count int
}

// scanRowsAndBuildAverageMap scans rows and builds average map concurrently
func (f *File) scanRowsAndBuildAverageMap(
	sheet string,
	rows [][]string,
	averageCol, criteria1Col, criteria2Col string,
) map[string]map[string]*avgData {

	if len(rows) == 0 {
		return nil
	}

	// Convert column letters to indices
	averageColIdx, _ := ColumnNameToNumber(averageCol)
	criteria1ColIdx, _ := ColumnNameToNumber(criteria1Col)
	criteria2ColIdx, _ := ColumnNameToNumber(criteria2Col)

	averageColIdx--   // Convert to 0-based
	criteria1ColIdx-- // Convert to 0-based
	criteria2ColIdx-- // Convert to 0-based

	numWorkers := runtime.NumCPU()
	rowCount := len(rows)

	if numWorkers > rowCount {
		numWorkers = rowCount
	}
	if numWorkers < 1 {
		numWorkers = 1
	}

	rowsPerWorker := (rowCount + numWorkers - 1) / numWorkers

	// Worker results
	type workerResult struct {
		data map[string]map[string]*avgData
	}
	results := make([]workerResult, numWorkers)
	var wg sync.WaitGroup

	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func(workerID int) {
			defer wg.Done()

			start := workerID * rowsPerWorker
			end := start + rowsPerWorker
			if end > rowCount {
				end = rowCount
			}

			localMap := make(map[string]map[string]*avgData)

			for rowIdx := start; rowIdx < end; rowIdx++ {
				row := rows[rowIdx]

				// Extract values from columns
				var c1, c2, avgVal string

				if criteria1ColIdx < len(row) {
					c1 = row[criteria1ColIdx]
				}
				if criteria2ColIdx < len(row) {
					c2 = row[criteria2ColIdx]
				}
				if averageColIdx < len(row) {
					avgVal = row[averageColIdx]
				}

				// Skip if criteria are empty
				if c1 == "" || c2 == "" {
					continue
				}

				// Skip if value is empty or "<>断货" pattern
				if avgVal == "" || avgVal == "断货" {
					continue
				}

				// Convert avgVal to number
				var num float64
				_, err := fmt.Sscanf(avgVal, "%f", &num)
				if err != nil {
					continue
				}

				// Accumulate
				if localMap[c1] == nil {
					localMap[c1] = make(map[string]*avgData)
				}
				if localMap[c1][c2] == nil {
					localMap[c1][c2] = &avgData{}
				}
				localMap[c1][c2].sum += num
				localMap[c1][c2].count++
			}

			results[workerID] = workerResult{data: localMap}
		}(i)
	}

	wg.Wait()

	// Merge results
	finalMap := make(map[string]map[string]*avgData)
	for _, r := range results {
		for c1, m := range r.data {
			if finalMap[c1] == nil {
				finalMap[c1] = make(map[string]*avgData)
			}
			for c2, data := range m {
				if finalMap[c1][c2] == nil {
					finalMap[c1][c2] = &avgData{}
				}
				finalMap[c1][c2].sum += data.sum
				finalMap[c1][c2].count += data.count
			}
		}
	}

	return finalMap
}
