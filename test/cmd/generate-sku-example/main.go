package main

import (
	"flag"
	"fmt"
	"log"
	"math/rand"
	"path/filepath"
	"regexp"
	"sort"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

// DataPattern holds patterns extracted from existing data
type DataPattern struct {
	InvOwners     []string
	InvWarehouses []string
	InvLocations  []string
	InvLocTypes   []string
	InvUnits      []string
	InvStatuses   []string
	InvCountries  []string

	TransitWarehouses []string
	TransitOwners     []string
	TransitPhases     []string
	TransitStatuses   []string
	TransitDocTypes   []string

	OutboundStatuses    []string
	OutboundOnlineStats []string
	OutboundShops       []string
	OutboundWarehouses  []string
}

type DateRange struct {
	Start time.Time
	End   time.Time
}

func main() {
	// Parse command line flags
	size := flag.Int("size", 1000, "Number of rows to generate (e.g., 5000 for 5k rows)")
	flag.Parse()

	if *size <= 0 {
		log.Fatal("Size must be a positive number")
	}

	// Calculate unique SKUs (approximately 10% of rows)
	uniqueSKUs := *size / 10
	if uniqueSKUs < 100 {
		uniqueSKUs = 100
	}

	baseTemplate := filepath.Join("test", "real-ecomm", "step3-template-1k-formulas.xlsx")
	outputPath := filepath.Join("test", "real-ecomm", fmt.Sprintf("step3-template-%d-formulas.xlsx", *size))

	log.Printf("=== Generating Template with %d Rows ===\n", *size)
	log.Printf("Base Template: %s\n", baseTemplate)
	log.Printf("Output: %s\n", outputPath)
	log.Printf("Unique SKUs: %d\n", uniqueSKUs)

	// Step 1: Generate mock data
	log.Println("\n=== STEP 1: Generating Mock Data ===")
	if err := generateMockData(baseTemplate, outputPath, *size, uniqueSKUs); err != nil {
		log.Fatalf("Failed Step 1: %v", err)
	}

	// Step 2: Populate formula sheets
	log.Println("\n=== STEP 2: Populating Formula Sheets ===")
	if err := populateFormulaSheets(outputPath); err != nil {
		log.Fatalf("Failed Step 2: %v", err)
	}

	log.Printf("\n=== ✓ Generation Complete ===\n")
	log.Printf("Output file: %s\n", outputPath)
}

func generateMockData(baseTemplate, outputPath string, totalRows, uniqueSKUs int) error {
	f, err := excelize.OpenFile(baseTemplate)
	if err != nil {
		return fmt.Errorf("failed to open base template: %w", err)
	}
	defer f.Close()

	// Extract patterns from base template
	log.Println("Extracting data patterns from base template...")
	patterns, err := extractPatterns(f)
	if err != nil {
		return fmt.Errorf("failed to extract patterns: %w", err)
	}

	// Generate SKU list
	log.Printf("Generating %d unique SKUs...\n", uniqueSKUs)
	skus := generateSKUs(uniqueSKUs)

	// Hardcoded date range: 2025-07-05 to 2025-12-28
	startDate := time.Date(2025, 7, 5, 0, 0, 0, 0, time.UTC)
	endDate := time.Date(2025, 12, 28, 0, 0, 0, 0, time.UTC)
	dateRange := DateRange{Start: startDate, End: endDate}

	rand.Seed(time.Now().UnixNano())

	// Calculate rows per SKU
	rowsPerSKU := totalRows / uniqueSKUs
	if rowsPerSKU < 1 {
		rowsPerSKU = 1
	}

	log.Printf("Generating ~%d rows per SKU\n", rowsPerSKU)

	// Generate data for each sheet concurrently with mutex protection
	var wg sync.WaitGroup
	var mu sync.Mutex
	errChan := make(chan error, 3)

	log.Println("Generating data for all sheets concurrently...")

	wg.Add(1)
	go func() {
		defer wg.Done()
		log.Println("  [Goroutine] Generating 库存台账-all...")
		if err := generateInventoryData(f, &mu, "库存台账-all", skus, totalRows, rowsPerSKU, patterns, dateRange); err != nil {
			errChan <- fmt.Errorf("failed to generate inventory data: %w", err)
		}
	}()

	wg.Add(1)
	go func() {
		defer wg.Done()
		log.Println("  [Goroutine] Generating 在途产品-all...")
		if err := generateTransitData(f, &mu, "在途产品-all", skus, totalRows, rowsPerSKU, patterns, dateRange); err != nil {
			errChan <- fmt.Errorf("failed to generate transit data: %w", err)
		}
	}()

	wg.Add(1)
	go func() {
		defer wg.Done()
		log.Println("  [Goroutine] Generating 出库记录-all...")
		if err := generateOutboundData(f, &mu, "出库记录-all", skus, totalRows, rowsPerSKU, patterns, dateRange); err != nil {
			errChan <- fmt.Errorf("failed to generate outbound data: %w", err)
		}
	}()

	wg.Wait()
	close(errChan)

	// Check for errors
	for err := range errChan {
		if err != nil {
			return err
		}
	}

	// Save file
	log.Println("Saving intermediate file...")
	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	log.Println("✓ Step 1 complete")
	return nil
}

func populateFormulaSheets(filePath string) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer f.Close()

	// Collect all unique SKUs from data sheets concurrently
	log.Println("Collecting unique SKUs from all data sheets...")
	allSKUs := make(map[string]bool)
	var skuMu sync.Mutex

	dataSheets := []struct {
		name   string
		skuCol int // 0-indexed
	}{
		{"库存台账-all", 3},  // Column D
		{"在途产品-all", 10}, // Column K
		{"出库记录-all", 8},  // Column I
	}

	var wg sync.WaitGroup
	errChan := make(chan error, len(dataSheets))

	for _, ds := range dataSheets {
		wg.Add(1)
		go func(sheetName string, skuCol int) {
			defer wg.Done()
			rows, err := f.GetRows(sheetName)
			if err != nil {
				errChan <- fmt.Errorf("failed to read %s: %w", sheetName, err)
				return
			}

			localSKUs := make(map[string]bool)
			for i := 1; i < len(rows); i++ { // Skip header
				if skuCol < len(rows[i]) && rows[i][skuCol] != "" {
					localSKUs[rows[i][skuCol]] = true
				}
			}

			skuMu.Lock()
			for sku := range localSKUs {
				allSKUs[sku] = true
			}
			skuMu.Unlock()

			log.Printf("  %s: collected SKUs\n", sheetName)
		}(ds.name, ds.skuCol)
	}

	wg.Wait()
	close(errChan)

	// Check for errors
	for err := range errChan {
		if err != nil {
			return err
		}
	}

	// Convert to sorted slice
	skus := make([]string, 0, len(allSKUs))
	for sku := range allSKUs {
		skus = append(skus, sku)
	}
	sort.Strings(skus)
	log.Printf("Total unique SKUs: %d\n", len(skus))

	// Formula sheets to populate
	formulaSheets := []string{
		"日库存",
		"日销售",
		"日销预测",
		"补货计划",
		"补货汇总",
	}

	// Process formula sheets concurrently with mutex protection
	var mu sync.Mutex
	errChan = make(chan error, len(formulaSheets))

	log.Println("Processing formula sheets concurrently...")

	for _, sheetName := range formulaSheets {
		wg.Add(1)
		go func(sheet string) {
			defer wg.Done()
			log.Printf("  [Goroutine] Processing %s...\n", sheet)

			// Get rows from template (to preserve header) with lock
			mu.Lock()
			templateRows, err := f.GetRows(sheet)
			mu.Unlock()

			if err != nil {
				errChan <- fmt.Errorf("failed to read sheet %s: %w", sheet, err)
				return
			}

			if len(templateRows) < 2 {
				log.Printf("  Warning: %s has no template row 2, skipping\n", sheet)
				return
			}

			// Preserve header row (Row 1) exactly as-is
			numCols := len(templateRows[0])
			headerRow := make([]string, numCols)
			for i := 0; i < numCols; i++ {
				headerRow[i] = templateRows[0][i]
			}

			// Read formulas and values from Row 2 (the template row)
			row2Formulas := make([]string, numCols)
			row2Values := make([]interface{}, numCols)

			mu.Lock()
			for colIdx := 0; colIdx < numCols; colIdx++ {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, 2)
				formula, _ := f.GetCellFormula(sheet, cell)
				value, _ := f.GetCellValue(sheet, cell)
				row2Formulas[colIdx] = formula
				row2Values[colIdx] = value
			}
			mu.Unlock()

			// Prepare all data WITHOUT lock (this is the heavy computation)
			type cellData struct {
				cell    string
				formula string
				value   interface{}
			}
			allCellData := make([]cellData, 0, len(skus)*numCols)

			for skuIdx, sku := range skus {
				rowNum := skuIdx + 2

				// SKU in column A
				cellA, _ := excelize.CoordinatesToCellName(1, rowNum)
				allCellData = append(allCellData, cellData{cell: cellA, value: sku})

				// Other columns based on Row 2 template
				for colIdx := 1; colIdx < numCols; colIdx++ {
					cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)

					if row2Formulas[colIdx] != "" {
						// Copy formula and adapt row references
						adjustedFormula := adjustFormulaRowReferences(row2Formulas[colIdx], rowNum)
						allCellData = append(allCellData, cellData{cell: cell, formula: adjustedFormula})
					} else if row2Values[colIdx] != nil && row2Values[colIdx] != "" {
						// Copy static value
						allCellData = append(allCellData, cellData{cell: cell, value: row2Values[colIdx]})
					}
				}

				// Progress logging (no lock needed for logging)
				if (skuIdx+1)%1000 == 0 || skuIdx == len(skus)-1 {
					log.Printf("    [%s] Progress: %d/%d SKUs\n", sheet, skuIdx+1, len(skus))
				}
			}

			// NOW do all Excel writes in ONE SINGLE LOCK
			mu.Lock()
			defer mu.Unlock()

			// Clear all data rows (keep header row 1)
			maxRow := len(templateRows)
			for rowIdx := 2; rowIdx <= maxRow; rowIdx++ {
				for colIdx := 1; colIdx <= numCols; colIdx++ {
					cell, _ := excelize.CoordinatesToCellName(colIdx, rowIdx)
					f.SetCellValue(sheet, cell, "")
				}
			}

			// Re-write header row
			for colIdx := 0; colIdx < numCols; colIdx++ {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
				f.SetCellValue(sheet, cell, headerRow[colIdx])
			}

			// Write all prepared data (formulas and values)
			for _, cd := range allCellData {
				if cd.formula != "" {
					f.SetCellFormula(sheet, cd.cell, cd.formula)
				} else {
					f.SetCellValue(sheet, cd.cell, cd.value)
				}
			}

			log.Printf("  ✓ Populated %d rows in %s\n", len(skus), sheet)
		}(sheetName)
	}

	wg.Wait()
	close(errChan)

	// Check for errors
	for err := range errChan {
		if err != nil {
			return err
		}
	}

	// Save file
	log.Println("Saving final file...")
	if err := f.Save(); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	log.Println("✓ Step 2 complete")
	return nil
}

// adjustFormulaRowReferences adjusts row references in formulas from row 2 to targetRow
// This preserves absolute row references (e.g., $A$2) and only changes relative row references
func adjustFormulaRowReferences(formula string, targetRow int) string {
	// Pattern to match cell references with row 2
	// Matches: A2, $A2, A$2, but NOT $A$2
	pattern := regexp.MustCompile(`(\$?)([A-Z]+)(\$?)2\b`)

	result := pattern.ReplaceAllStringFunc(formula, func(match string) string {
		// Check if this is an absolute row reference (e.g., A$2 or $A$2)
		if regexp.MustCompile(`\$2\b`).MatchString(match) {
			return match // Keep absolute row references unchanged
		}

		// Replace "2" with targetRow for relative references
		return regexp.MustCompile(`2\b`).ReplaceAllString(match, fmt.Sprintf("%d", targetRow))
	})

	return result
}

func extractPatterns(f *excelize.File) (*DataPattern, error) {
	patterns := &DataPattern{}
	var mu sync.Mutex
	var wg sync.WaitGroup
	errChan := make(chan error, 3)

	// Extract patterns concurrently from all sheets
	wg.Add(1)
	go func() {
		defer wg.Done()
		invRows, err := f.GetRows("库存台账-all")
		if err != nil {
			errChan <- fmt.Errorf("failed to read 库存台账-all: %w", err)
			return
		}
		if len(invRows) > 1 {
			mu.Lock()
			patterns.InvOwners = extractUniqueValues(invRows, 1)
			patterns.InvWarehouses = extractUniqueValues(invRows, 2)
			patterns.InvLocations = extractUniqueValues(invRows, 5)
			patterns.InvLocTypes = extractUniqueValues(invRows, 6)
			patterns.InvUnits = extractUniqueValues(invRows, 8)
			patterns.InvStatuses = extractUniqueValues(invRows, 27)
			patterns.InvCountries = extractUniqueValues(invRows, 26)
			mu.Unlock()
		}
	}()

	wg.Add(1)
	go func() {
		defer wg.Done()
		transitRows, err := f.GetRows("在途产品-all")
		if err != nil {
			errChan <- fmt.Errorf("failed to read 在途产品-all: %w", err)
			return
		}
		if len(transitRows) > 1 {
			mu.Lock()
			patterns.TransitWarehouses = extractUniqueValues(transitRows, 5)
			patterns.TransitOwners = extractUniqueValues(transitRows, 6)
			patterns.TransitPhases = extractUniqueValues(transitRows, 7)
			patterns.TransitStatuses = extractUniqueValues(transitRows, 8)
			patterns.TransitDocTypes = extractUniqueValues(transitRows, 27)
			mu.Unlock()
		}
	}()

	wg.Add(1)
	go func() {
		defer wg.Done()
		outboundRows, err := f.GetRows("出库记录-all")
		if err != nil {
			errChan <- fmt.Errorf("failed to read 出库记录-all: %w", err)
			return
		}
		if len(outboundRows) > 1 {
			mu.Lock()
			patterns.OutboundStatuses = extractUniqueValues(outboundRows, 1)
			patterns.OutboundOnlineStats = extractUniqueValues(outboundRows, 2)
			patterns.OutboundShops = extractUniqueValues(outboundRows, 6)
			patterns.OutboundWarehouses = extractUniqueValues(outboundRows, 10)
			mu.Unlock()
		}
	}()

	wg.Wait()
	close(errChan)

	// Check for errors
	for err := range errChan {
		if err != nil {
			return nil, err
		}
	}

	return patterns, nil
}

func extractUniqueValues(rows [][]string, colIdx int) []string {
	uniqueMap := make(map[string]bool)
	var result []string

	for i := 1; i < len(rows); i++ {
		if colIdx < len(rows[i]) && rows[i][colIdx] != "" {
			if !uniqueMap[rows[i][colIdx]] {
				uniqueMap[rows[i][colIdx]] = true
				result = append(result, rows[i][colIdx])
			}
		}
	}

	if len(result) == 0 {
		result = append(result, "DEFAULT")
	}

	return result
}

func generateSKUs(count int) []string {
	skus := make([]string, count)
	for i := 0; i < count; i++ {
		skus[i] = fmt.Sprintf("BF1D-74352625-%05d", i+1)
	}
	return skus
}

func generateInventoryData(f *excelize.File, mu *sync.Mutex, sheetName string, skus []string, totalRows, rowsPerSKU int, patterns *DataPattern, dateRange DateRange) error {
	// Build all data in memory first (NO locks during computation)
	cellValues := make(map[string]interface{})
	rowNum := 2

	for _, sku := range skus {
		numRows := rowsPerSKU + rand.Intn(3) - 1
		if numRows < 1 {
			numRows = 1
		}
		if rowNum+numRows > totalRows+2 {
			numRows = totalRows + 2 - rowNum
		}

		for i := 0; i < numRows && rowNum <= totalRows+1; i++ {
			date := randomDate(dateRange)
			owner := randomElement(patterns.InvOwners)
			warehouse := randomElement(patterns.InvWarehouses)
			location := randomElement(patterns.InvLocations)
			locType := randomElement(patterns.InvLocTypes)
			quantity := rand.Intn(300) + 1
			unit := randomElement(patterns.InvUnits)
			status := randomElement(patterns.InvStatuses)
			country := randomElement(patterns.InvCountries)

			// Build cell map for batch insertion
			cellValues[fmt.Sprintf("A%d", rowNum)] = date.Format("2006-01-02")
			cellValues[fmt.Sprintf("B%d", rowNum)] = owner
			cellValues[fmt.Sprintf("C%d", rowNum)] = warehouse
			cellValues[fmt.Sprintf("D%d", rowNum)] = sku
			cellValues[fmt.Sprintf("E%d", rowNum)] = ""
			cellValues[fmt.Sprintf("F%d", rowNum)] = location
			cellValues[fmt.Sprintf("G%d", rowNum)] = locType
			cellValues[fmt.Sprintf("H%d", rowNum)] = quantity
			cellValues[fmt.Sprintf("I%d", rowNum)] = unit
			cellValues[fmt.Sprintf("J%d", rowNum)] = quantity
			cellValues[fmt.Sprintf("AB%d", rowNum)] = status
			cellValues[fmt.Sprintf("AA%d", rowNum)] = country

			rowNum++
		}

		if rowNum > totalRows+1 {
			break
		}
	}

	// Single lock for all Excel operations (delete old + write new)
	mu.Lock()
	defer mu.Unlock()

	// Fast approach: Delete entire sheet and recreate with header
	index, err := f.GetSheetIndex(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}

	// Get header row before deleting
	rows, _ := f.GetRows(sheetName)
	var headerRow []string
	if len(rows) > 0 {
		headerRow = rows[0]
	}

	// Delete and recreate sheet (much faster than RemoveRow loop)
	f.DeleteSheet(sheetName)
	f.NewSheet(sheetName)
	f.SetActiveSheet(index)

	// Restore header
	if len(headerRow) > 0 {
		for colIdx, val := range headerRow {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheetName, cell, val)
		}
	}

	// Write all cells in one batch operation
	if err := f.SetCellValues(sheetName, cellValues); err != nil {
		return fmt.Errorf("failed to set cell values: %w", err)
	}

	log.Printf("  ✓ Generated %d rows\n", rowNum-2)
	return nil
}

func generateTransitData(f *excelize.File, mu *sync.Mutex, sheetName string, skus []string, totalRows, rowsPerSKU int, patterns *DataPattern, dateRange DateRange) error {
	// Build all data in memory first (NO locks during computation)
	cellValues := make(map[string]interface{})
	rowNum := 2

	for _, sku := range skus {
		numRows := rowsPerSKU + rand.Intn(3) - 1
		if numRows < 1 {
			numRows = 1
		}
		if rowNum+numRows > totalRows+2 {
			numRows = totalRows + 2 - rowNum
		}

		for i := 0; i < numRows && rowNum <= totalRows+1; i++ {
			arrivalDate := randomDate(dateRange)
			createDate := arrivalDate.AddDate(0, 0, -rand.Intn(5)-1)
			warehouse := randomElement(patterns.TransitWarehouses)
			owner := randomElement(patterns.TransitOwners)
			phase := randomElement(patterns.TransitPhases)
			status := randomElement(patterns.TransitStatuses)
			docType := randomElement(patterns.TransitDocTypes)
			quantity := rand.Intn(5000) + 100
			receiptNum := fmt.Sprintf("A%d", rand.Intn(900000000)+100000000)
			externalNum := fmt.Sprintf("A%d-%d", rand.Intn(90000000)+10000000, rand.Intn(90000000)+10000000)

			cellValues[fmt.Sprintf("A%d", rowNum)] = arrivalDate.Format("2006-01-02")
			cellValues[fmt.Sprintf("B%d", rowNum)] = createDate.Format("2006-01-02")
			cellValues[fmt.Sprintf("C%d", rowNum)] = ""
			cellValues[fmt.Sprintf("D%d", rowNum)] = receiptNum
			cellValues[fmt.Sprintf("E%d", rowNum)] = externalNum
			cellValues[fmt.Sprintf("F%d", rowNum)] = warehouse
			cellValues[fmt.Sprintf("G%d", rowNum)] = owner
			cellValues[fmt.Sprintf("H%d", rowNum)] = phase
			cellValues[fmt.Sprintf("I%d", rowNum)] = status
			cellValues[fmt.Sprintf("J%d", rowNum)] = ""
			cellValues[fmt.Sprintf("K%d", rowNum)] = sku
			cellValues[fmt.Sprintf("L%d", rowNum)] = ""
			cellValues[fmt.Sprintf("M%d", rowNum)] = quantity
			cellValues[fmt.Sprintf("AC%d", rowNum)] = docType
			cellValues[fmt.Sprintf("AD%d", rowNum)] = createDate.Format("2006-01-02")

			rowNum++
		}

		if rowNum > totalRows+1 {
			break
		}
	}

	// Single lock for all Excel operations
	mu.Lock()
	defer mu.Unlock()

	// Fast approach: Delete entire sheet and recreate with header
	index, err := f.GetSheetIndex(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}

	// Get header row before deleting
	rows, _ := f.GetRows(sheetName)
	var headerRow []string
	if len(rows) > 0 {
		headerRow = rows[0]
	}

	// Delete and recreate sheet
	f.DeleteSheet(sheetName)
	f.NewSheet(sheetName)
	f.SetActiveSheet(index)

	// Restore header
	if len(headerRow) > 0 {
		for colIdx, val := range headerRow {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheetName, cell, val)
		}
	}

	// Write all cells in one batch operation
	if err := f.SetCellValues(sheetName, cellValues); err != nil {
		return fmt.Errorf("failed to set cell values: %w", err)
	}

	log.Printf("  ✓ Generated %d rows\n", rowNum-2)
	return nil
}

func generateOutboundData(f *excelize.File, mu *sync.Mutex, sheetName string, skus []string, totalRows, rowsPerSKU int, patterns *DataPattern, dateRange DateRange) error {
	// Build all data in memory first (NO locks during computation)
	cellValues := make(map[string]interface{})
	rowNum := 2

	for _, sku := range skus {
		numRows := rowsPerSKU + rand.Intn(3) - 1
		if numRows < 1 {
			numRows = 1
		}
		if rowNum+numRows > totalRows+2 {
			numRows = totalRows + 2 - rowNum
		}

		for i := 0; i < numRows && rowNum <= totalRows+1; i++ {
			paymentTime := randomDate(dateRange)
			shipmentTime := paymentTime
			latestShipTime := paymentTime.AddDate(0, 0, rand.Intn(3)+1)
			orderStatus := randomElement(patterns.OutboundStatuses)
			onlineStatus := randomElement(patterns.OutboundOnlineStats)
			shop := randomElement(patterns.OutboundShops)
			warehouse := randomElement(patterns.OutboundWarehouses)
			quantity := rand.Intn(5) + 1
			erpNum := fmt.Sprintf("S%d%02d%08d", paymentTime.Year(), int(paymentTime.Month()), rand.Intn(100000000))
			onlineCode := fmt.Sprintf("BF1D-74352625 (%s)", sku)

			cellValues[fmt.Sprintf("A%d", rowNum)] = erpNum
			cellValues[fmt.Sprintf("B%d", rowNum)] = orderStatus
			cellValues[fmt.Sprintf("C%d", rowNum)] = onlineStatus
			cellValues[fmt.Sprintf("D%d", rowNum)] = paymentTime.Format("2006-01-02 15:04:05") + " UTC+7"
			cellValues[fmt.Sprintf("E%d", rowNum)] = shipmentTime.Format("2006-01-02 15:04:05") + " UTC+7"
			cellValues[fmt.Sprintf("F%d", rowNum)] = latestShipTime.Format("2006-01-02 15:04:05") + " UTC+8"
			cellValues[fmt.Sprintf("G%d", rowNum)] = shop
			cellValues[fmt.Sprintf("H%d", rowNum)] = onlineCode
			cellValues[fmt.Sprintf("I%d", rowNum)] = sku
			cellValues[fmt.Sprintf("J%d", rowNum)] = quantity
			cellValues[fmt.Sprintf("K%d", rowNum)] = warehouse
			cellValues[fmt.Sprintf("L%d", rowNum)] = paymentTime.Format("2006-01-02")

			rowNum++
		}

		if rowNum > totalRows+1 {
			break
		}
	}

	// Single lock for all Excel operations
	mu.Lock()
	defer mu.Unlock()

	// Fast approach: Delete entire sheet and recreate with header
	index, err := f.GetSheetIndex(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}

	// Get header row before deleting
	rows, _ := f.GetRows(sheetName)
	var headerRow []string
	if len(rows) > 0 {
		headerRow = rows[0]
	}

	// Delete and recreate sheet
	f.DeleteSheet(sheetName)
	f.NewSheet(sheetName)
	f.SetActiveSheet(index)

	// Restore header
	if len(headerRow) > 0 {
		for colIdx, val := range headerRow {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheetName, cell, val)
		}
	}

	// Write all cells in one batch operation
	if err := f.SetCellValues(sheetName, cellValues); err != nil {
		return fmt.Errorf("failed to set cell values: %w", err)
	}

	log.Printf("  ✓ Generated %d rows\n", rowNum-2)
	return nil
}

func randomDate(dr DateRange) time.Time {
	if dr.Start.IsZero() || dr.End.IsZero() {
		return time.Now()
	}

	delta := dr.End.Unix() - dr.Start.Unix()
	sec := rand.Int63n(delta)
	return time.Unix(dr.Start.Unix()+sec, 0)
}

func randomElement(arr []string) string {
	if len(arr) == 0 {
		return ""
	}
	return arr[rand.Intn(len(arr))]
}
