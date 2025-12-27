package excelize

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestWriteNonDestructive tests that WriteNonDestructive preserves internal state
func TestWriteNonDestructive(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	fmt.Println("\n=== Test WriteNonDestructive vs Write ===")

	// Step 1: Create data with empty rows
	fmt.Println("\n--- Step 1: Create data (rows 1-5, 100) ---")
	for row := 1; row <= 5; row++ {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
	}
	f.SetCellValue(sheetName, "A100", "Row100")

	// Check initial state
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)
	ws, _ := f.Sheet.Load(sheetXMLPath)
	worksheet := ws.(*xlsxWorksheet)
	initialRowCount := len(worksheet.SheetData.Row)
	fmt.Printf("Initial row count: %d\n", initialRowCount)

	// Step 2: Call WriteNonDestructive
	fmt.Println("\n--- Step 2: Call WriteNonDestructive ---")
	var buf1 bytes.Buffer
	err := f.WriteNonDestructive(&buf1)
	assert.NoError(t, err)
	fmt.Printf("WriteNonDestructive output: %d bytes\n", buf1.Len())

	// Check state after WriteNonDestructive
	ws, _ = f.Sheet.Load(sheetXMLPath)
	worksheet = ws.(*xlsxWorksheet)
	afterNonDestructiveCount := len(worksheet.SheetData.Row)
	fmt.Printf("After WriteNonDestructive row count: %d\n", afterNonDestructiveCount)

	// 🔥 CRITICAL: Row count should NOT change!
	assert.Equal(t, initialRowCount, afterNonDestructiveCount, "WriteNonDestructive should NOT modify row count")

	// Step 3: Verify we can still write to empty rows
	fmt.Println("\n--- Step 3: Write to empty row 50 ---")
	err = f.SetCellValue(sheetName, "A50", "NewRow50")
	assert.NoError(t, err)

	val, err := f.GetCellValue(sheetName, "A50")
	assert.NoError(t, err)
	fmt.Printf("A50 value: '%s'\n", val)
	assert.Equal(t, "NewRow50", val, "Should be able to write to row 50 after WriteNonDestructive")

	// Step 4: Compare with original Write()
	fmt.Println("\n--- Step 4: Compare with Write() ---")
	f2 := NewFile()
	f2.options = &Options{KeepWorksheetInMemory: true}  // 🔥 Keep worksheet in memory
	defer f2.Close()

	for row := 1; row <= 5; row++ {
		f2.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
	}
	f2.SetCellValue("Sheet1", "A100", "Row100")

	sheetXMLPath2, _ := f2.getSheetXMLPath("Sheet1")
	ws2, _ := f2.Sheet.Load(sheetXMLPath2)
	worksheet2 := ws2.(*xlsxWorksheet)
	beforeWriteCount := len(worksheet2.SheetData.Row)
	fmt.Printf("Before Write() row count: %d\n", beforeWriteCount)

	var buf2 bytes.Buffer
	err = f2.Write(&buf2)
	assert.NoError(t, err)

	var afterWriteCount int
	ws2, loaded := f2.Sheet.Load(sheetXMLPath2)
	if !loaded || ws2 == nil {
		fmt.Println("⚠️  Worksheet was deleted by Write() despite KeepWorksheetInMemory")
		afterWriteCount = 0
	} else {
		worksheet2 = ws2.(*xlsxWorksheet)
		afterWriteCount = len(worksheet2.SheetData.Row)
	}
	fmt.Printf("After Write() row count: %d\n", afterWriteCount)

	// 💥 Write() WILL modify row count (trimRow deletes empty rows)
	assert.Less(t, afterWriteCount, beforeWriteCount, "Write() SHOULD modify row count (trimRow)")

	// Try to write to row 50
	err = f2.SetCellValue("Sheet1", "A50", "NewRow50")
	assert.NoError(t, err)

	val2, err := f2.GetCellValue("Sheet1", "A50")
	assert.NoError(t, err)
	fmt.Printf("A50 value after Write(): '%s'\n", val2)

	// This might still work due to prepareSheetXML auto-expansion,
	// but the internal state has been modified
}

// TestWriteNonDestructiveMultiWorksheet tests multi-worksheet scenario
func TestWriteNonDestructiveMultiWorksheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	fmt.Println("\n=== Test WriteNonDestructive Multi-Worksheet ===")

	// Create Sheet A with 91 columns
	sheetA := "SheetA"
	f.NewSheet(sheetA)
	fmt.Println("\n--- Create Sheet A (91 columns) ---")
	for col := 1; col <= 91; col++ {
		colName, _ := ColumnNumberToName(col)
		f.SetCellValue(sheetA, colName+"1", fmt.Sprintf("A-Col%d", col))
	}

	// Create Sheet B with 30 columns
	sheetB := "SheetB"
	f.NewSheet(sheetB)
	fmt.Println("\n--- Create Sheet B (30 columns) ---")
	for col := 1; col <= 30; col++ {
		colName, _ := ColumnNumberToName(col)
		f.SetCellValue(sheetB, colName+"1", fmt.Sprintf("B-Col%d", col))
	}

	// Call WriteNonDestructive
	fmt.Println("\n--- Call WriteNonDestructive ---")
	var buf bytes.Buffer
	err := f.WriteNonDestructive(&buf)
	assert.NoError(t, err)

	// Switch back to Sheet B and insert rows
	fmt.Println("\n--- Insert rows in Sheet B ---")
	err = f.InsertRows(sheetB, 2, 5)
	assert.NoError(t, err)

	// Write to column A (this should work correctly!)
	fmt.Println("\n--- Write to column A in Sheet B ---")
	for row := 2; row <= 6; row++ {
		cell := fmt.Sprintf("A%d", row)
		value := fmt.Sprintf("SKU-%d", row)
		err = f.SetCellValue(sheetB, cell, value)
		assert.NoError(t, err)

		// Verify immediately
		readBack, _ := f.GetCellValue(sheetB, cell)
		fmt.Printf("%s: wrote='%s', read='%s'\n", cell, value, readBack)
		assert.Equal(t, value, readBack, "Column A should be written correctly")
	}

	fmt.Println("\n✅ WriteNonDestructive preserves multi-worksheet state correctly!")
}

// TestWriteNonDestructiveWithInsertRows tests the exact production bug scenario
func TestWriteNonDestructiveWithInsertRows(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer f.Close()

	sheetName := "Sheet1"

	fmt.Println("\n=== Test WriteNonDestructive with InsertRows (Production Bug) ===")

	// Step 1: Create initial data
	fmt.Println("\n--- Step 1: Create rows 1-5 ---")
	for row := 1; row <= 5; row++ {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
	}

	// Step 2: InsertRows (simulating production)
	fmt.Println("\n--- Step 2: InsertRows(6, 10) ---")
	err := f.InsertRows(sheetName, 6, 10)
	assert.NoError(t, err)

	// Step 3: WriteNonDestructive (simulating save to GridFS)
	fmt.Println("\n--- Step 3: WriteNonDestructive (save to GridFS) ---")
	var buf bytes.Buffer
	err = f.WriteNonDestructive(&buf)
	assert.NoError(t, err)

	// Step 4: Write to inserted rows (THIS SHOULD WORK!)
	fmt.Println("\n--- Step 4: Write to inserted rows 6-15 ---")
	for row := 6; row <= 15; row++ {
		cell := fmt.Sprintf("A%d", row)
		value := fmt.Sprintf("SKU-%d", row)
		err = f.SetCellValue(sheetName, cell, value)
		assert.NoError(t, err)

		readBack, _ := f.GetCellValue(sheetName, cell)
		fmt.Printf("%s: wrote='%s', read='%s'\n", cell, value, readBack)

		// 🔥 CRITICAL: This should work with WriteNonDestructive!
		assert.Equal(t, value, readBack, "Row %d should be written correctly", row)
	}

	fmt.Println("\n✅ WriteNonDestructive fixes the production bug!")
}

// TestWriteVsWriteNonDestructiveComparison compares output size
func TestWriteVsWriteNonDestructiveComparison(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create sparse data
	f.SetCellValue("Sheet1", "A1", "Start")
	f.SetCellValue("Sheet1", "A100", "End")

	// Write()
	var buf1 bytes.Buffer
	err := f.Write(&buf1)
	assert.NoError(t, err)

	// Need to reload because Write() modified state
	f2 := NewFile()
	defer f2.Close()
	f2.SetCellValue("Sheet1", "A1", "Start")
	f2.SetCellValue("Sheet1", "A100", "End")

	// WriteNonDestructive()
	var buf2 bytes.Buffer
	err = f2.WriteNonDestructive(&buf2)
	assert.NoError(t, err)

	fmt.Printf("\nOutput size comparison:\n")
	fmt.Printf("  Write():              %d bytes\n", buf1.Len())
	fmt.Printf("  WriteNonDestructive(): %d bytes\n", buf2.Len())

	// Both should produce valid Excel files with similar sizes
	// The difference is that WriteNonDestructive preserves internal state
	assert.Greater(t, buf1.Len(), 0)
	assert.Greater(t, buf2.Len(), 0)
}

// BenchmarkWriteOriginal benchmarks the original Write()
func BenchmarkWriteOriginal(b *testing.B) {
	f := NewFile()
	defer f.Close()

	for row := 1; row <= 100; row++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var buf bytes.Buffer
		_ = f.Write(&buf)
	}
}

// BenchmarkWriteNonDestructive benchmarks the new WriteNonDestructive()
func BenchmarkWriteNonDestructive(b *testing.B) {
	f := NewFile()
	defer f.Close()

	for row := 1; row <= 100; row++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("Row%d", row))
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var buf bytes.Buffer
		_ = f.WriteNonDestructive(&buf)
	}
}
