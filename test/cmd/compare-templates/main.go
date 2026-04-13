package main

import (
	"log"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Files to compare
	generatedFile := filepath.Join("test", "real-ecomm", "step3-template-5000-formulas.xlsx")
	referenceFile := filepath.Join("test", "real-ecomm", "step3-template-5k-formulas.xlsx")

	log.Println("=== Comparing Generated File with Reference ===")
	log.Printf("Generated: %s\n", generatedFile)
	log.Printf("Reference: %s\n", referenceFile)

	// Open both files
	fGen, err := excelize.OpenFile(generatedFile)
	if err != nil {
		log.Fatalf("Failed to open generated file: %v", err)
	}
	defer fGen.Close()

	fRef, err := excelize.OpenFile(referenceFile)
	if err != nil {
		log.Fatalf("Failed to open reference file: %v", err)
	}
	defer fRef.Close()

	// Formula sheets to compare
	formulaSheets := []string{
		"日库存",
		"日销售",
		"日销预测",
		"补货计划",
		"补货汇总",
	}

	allMatch := true

	for _, sheetName := range formulaSheets {
		log.Printf("\n=== Sheet: %s ===\n", sheetName)

		// Get rows from both files
		rowsGen, err := fGen.GetRows(sheetName)
		if err != nil {
			log.Printf("Error reading generated sheet: %v\n", err)
			continue
		}

		rowsRef, err := fRef.GetRows(sheetName)
		if err != nil {
			log.Printf("Error reading reference sheet: %v\n", err)
			continue
		}

		if len(rowsGen) < 2 || len(rowsRef) < 2 {
			log.Printf("Skipping - insufficient rows\n")
			continue
		}

		// Compare Header Row (Row 1)
		log.Println("Comparing Header Row (Row 1):")
		headerMatches := compareHeaders(rowsGen[0], rowsRef[0])
		if headerMatches {
			log.Println("  ✓ Headers match")
		} else {
			log.Println("  ✗ Headers DO NOT match")
			allMatch = false
		}

		// Compare Row 2 Formulas
		log.Println("\nComparing Row 2 Formulas:")
		formulasMatch := compareFormulas(fGen, fRef, sheetName, 2, len(rowsGen[0]))
		if formulasMatch {
			log.Println("  ✓ Row 2 formulas match")
		} else {
			log.Println("  ✗ Row 2 formulas DO NOT match")
			allMatch = false
		}

		// Show sample of Row 2 values
		log.Println("\nRow 2 Sample (first 5 columns):")
		log.Println("  Generated:")
		for i := 0; i < min(5, len(rowsGen[0])); i++ {
			cell, _ := excelize.CoordinatesToCellName(i+1, 2)
			formula, _ := fGen.GetCellFormula(sheetName, cell)
			value, _ := fGen.GetCellValue(sheetName, cell)
			if formula != "" {
				log.Printf("    %s: %s\n", cell, formula)
			} else {
				log.Printf("    %s: %s (value)\n", cell, value)
			}
		}
		log.Println("  Reference:")
		for i := 0; i < min(5, len(rowsRef[0])); i++ {
			cell, _ := excelize.CoordinatesToCellName(i+1, 2)
			formula, _ := fRef.GetCellFormula(sheetName, cell)
			value, _ := fRef.GetCellValue(sheetName, cell)
			if formula != "" {
				log.Printf("    %s: %s\n", cell, formula)
			} else {
				log.Printf("    %s: %s (value)\n", cell, value)
			}
		}
	}

	log.Println("\n=== Comparison Complete ===")
	if allMatch {
		log.Println("✓ All headers and formulas match!")
	} else {
		log.Println("✗ Some differences found - see details above")
	}
}

func compareHeaders(genHeader, refHeader []string) bool {
	maxCols := min(len(genHeader), len(refHeader))
	if maxCols > 15 {
		maxCols = 15 // Compare first 15 columns
	}

	allMatch := true
	for i := 0; i < maxCols; i++ {
		gen := ""
		ref := ""
		if i < len(genHeader) {
			gen = genHeader[i]
		}
		if i < len(refHeader) {
			ref = refHeader[i]
		}

		if gen != ref {
			log.Printf("  Column %d: Gen='%s' vs Ref='%s' ✗\n", i+1, gen, ref)
			allMatch = false
		}
	}

	return allMatch
}

func compareFormulas(fGen, fRef *excelize.File, sheetName string, rowNum, numCols int) bool {
	if numCols > 15 {
		numCols = 15 // Compare first 15 columns
	}

	allMatch := true
	diffCount := 0

	for colIdx := 0; colIdx < numCols; colIdx++ {
		cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)

		formulaGen, _ := fGen.GetCellFormula(sheetName, cell)
		formulaRef, _ := fRef.GetCellFormula(sheetName, cell)

		// Only compare if at least one has a formula
		if formulaGen != "" || formulaRef != "" {
			if formulaGen != formulaRef {
				if diffCount < 3 { // Show first 3 differences
					log.Printf("  %s: DIFFERENT\n", cell)
					log.Printf("    Gen: %s\n", formulaGen)
					log.Printf("    Ref: %s\n", formulaRef)
				}
				diffCount++
				allMatch = false
			}
		}
	}

	if diffCount > 3 {
		log.Printf("  ... and %d more differences\n", diffCount-3)
	}

	return allMatch
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
