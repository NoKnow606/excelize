package excelize

import (
	"fmt"
	"runtime"
	"strings"
	"sync"
)

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
			patterns := f.groupSUMIFSByPattern(sumifsFormulas)

			// Calculate each pattern
			for _, pattern := range patterns {
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
