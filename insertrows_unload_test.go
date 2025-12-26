package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestInsertRowsDoesNotUnloadWorksheet verifies that InsertRows keeps worksheet in memory
func TestInsertRowsDoesNotUnloadWorksheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheetName := "Sheet1"
	sheetXMLPath, ok := f.getSheetXMLPath(sheetName)
	assert.True(t, ok, "Sheet should exist")

	// Setup: Set initial data
	assert.NoError(t, f.SetCellValue(sheetName, "A1", "Data1"))
	assert.NoError(t, f.SetCellValue(sheetName, "A2", "Data2"))

	// Verify worksheet is loaded in memory
	fmt.Println("\n=== 检查初始状态 ===")
	worksheet, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded, "Worksheet should be loaded after SetCellValue")
	assert.NotNil(t, worksheet, "Worksheet should not be nil")
	fmt.Printf("✅ Worksheet '%s' 在内存中\n", sheetName)

	// Call InsertRows
	fmt.Println("\n=== 调用 InsertRows ===")
	err := f.InsertRows(sheetName, 2, 1)
	assert.NoError(t, err)
	fmt.Printf("✅ InsertRows 成功\n")

	// Check if worksheet is still in memory
	fmt.Println("\n=== 检查 InsertRows 后的状态 ===")
	worksheetAfter, loadedAfter := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loadedAfter, "Worksheet should STILL be loaded after InsertRows")
	assert.NotNil(t, worksheetAfter, "Worksheet should STILL not be nil")
	fmt.Printf("✅ Worksheet '%s' 仍然在内存中\n", sheetName)

	// Verify they are the same instance (not reloaded)
	assert.Equal(t, fmt.Sprintf("%p", worksheet), fmt.Sprintf("%p", worksheetAfter),
		"Should be the same worksheet instance (not reloaded)")
	fmt.Printf("✅ Worksheet 是同一个实例（未重新加载）\n")

	// Verify we can still access the data
	fmt.Println("\n=== 验证数据仍然可访问 ===")
	a1, err := f.GetCellValue(sheetName, "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Data1", a1)
	fmt.Printf("A1 = '%s' ✅\n", a1)

	a3, err := f.GetCellValue(sheetName, "A3")
	assert.NoError(t, err)
	assert.Equal(t, "Data2", a3, "Original A2 should move to A3")
	fmt.Printf("A3 = '%s' (原 A2 移动到这里) ✅\n", a3)
}

// TestMultipleInsertRowsKeepsWorksheet verifies multiple InsertRows don't unload
func TestMultipleInsertRowsKeepsWorksheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup initial data
	for i := 1; i <= 5; i++ {
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", i), fmt.Sprintf("Data%d", i))
	}

	// Get initial worksheet pointer
	wsInitial, _ := f.Sheet.Load(sheetXMLPath)
	initialPtr := fmt.Sprintf("%p", wsInitial)

	fmt.Println("\n=== 多次调用 InsertRows ===")

	// Call InsertRows multiple times
	for i := 1; i <= 3; i++ {
		err := f.InsertRows(sheetName, 3, 1)
		assert.NoError(t, err)

		// Check if still in memory
		ws, loaded := f.Sheet.Load(sheetXMLPath)
		assert.True(t, loaded, "Worksheet should be in memory after InsertRows #%d", i)

		currentPtr := fmt.Sprintf("%p", ws)
		fmt.Printf("InsertRows #%d: Worksheet 在内存中, Ptr=%s\n", i, currentPtr)

		// Verify it's the same instance
		assert.Equal(t, initialPtr, currentPtr, "Should be same instance after InsertRows #%d", i)
	}

	fmt.Println("✅ 多次 InsertRows 后 worksheet 始终在内存中")
}

// TestInsertRowsThenBatchUpdateWorksheetState verifies worksheet state during the full flow
func TestInsertRowsThenBatchUpdateWorksheetState(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheetName := "Sheet1"
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

	// Setup
	f.SetCellValue(sheetName, "A1", "Original")

	fmt.Println("\n=== 完整流程：InsertRows → BatchUpdate → GetValue ===")

	// Step 1: Check initial state
	ws1, loaded1 := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded1)
	ptr1 := fmt.Sprintf("%p", ws1)
	fmt.Printf("1️⃣  初始状态: Worksheet 在内存, Ptr=%s\n", ptr1)

	// Step 2: InsertRows
	err := f.InsertRows(sheetName, 2, 1)
	assert.NoError(t, err)

	ws2, loaded2 := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded2)
	ptr2 := fmt.Sprintf("%p", ws2)
	fmt.Printf("2️⃣  InsertRows 后: Worksheet 在内存, Ptr=%s (相同=%v)\n", ptr2, ptr1 == ptr2)

	// Step 3: BatchUpdateAndRecalculate
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A2", Value: "Inserted"},
		{Sheet: sheetName, Cell: "B2", Value: 123},
	}
	_, err = f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	ws3, loaded3 := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded3)
	ptr3 := fmt.Sprintf("%p", ws3)
	fmt.Printf("3️⃣  BatchUpdate 后: Worksheet 在内存, Ptr=%s (相同=%v)\n", ptr3, ptr1 == ptr3)

	// Step 4: GetCellValue
	a2, err := f.GetCellValue(sheetName, "A2")
	assert.NoError(t, err)
	assert.Equal(t, "Inserted", a2)

	ws4, loaded4 := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded4)
	ptr4 := fmt.Sprintf("%p", ws4)
	fmt.Printf("4️⃣  GetValue 后: Worksheet 在内存, Ptr=%s (相同=%v)\n", ptr4, ptr1 == ptr4)

	fmt.Printf("\n✅ 所有步骤中 worksheet 始终是同一个实例\n")
	fmt.Printf("✅ A2 = '%s' (获取成功)\n", a2)

	// All pointers should be the same
	assert.Equal(t, ptr1, ptr2, "Pointer should be same after InsertRows")
	assert.Equal(t, ptr1, ptr3, "Pointer should be same after BatchUpdate")
	assert.Equal(t, ptr1, ptr4, "Pointer should be same after GetValue")
}

// TestInsertRowsWithWriteAndKeepMemory tests InsertRows with KeepWorksheetInMemory option
func TestInsertRowsWithWriteAndKeepMemory(t *testing.T) {
	t.Run("Without KeepWorksheetInMemory", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		sheetName := "Sheet1"
		sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

		// Setup data
		f.SetCellValue(sheetName, "A1", "Data")

		// InsertRows
		err := f.InsertRows(sheetName, 2, 1)
		assert.NoError(t, err)

		ws1, loaded1 := f.Sheet.Load(sheetXMLPath)
		assert.True(t, loaded1)
		ptr1 := fmt.Sprintf("%p", ws1)
		fmt.Printf("\nInsertRows 后: Worksheet 在内存, Ptr=%s\n", ptr1)

		// Write (without KeepWorksheetInMemory)
		var buf []byte
		_, err = f.WriteTo(&mockWriter{&buf})
		assert.NoError(t, err)

		// Check if unloaded
		_, loaded2 := f.Sheet.Load(sheetXMLPath)
		fmt.Printf("Write 后 (默认): Worksheet 在内存=%v\n", loaded2)
		assert.False(t, loaded2, "Should be unloaded after Write by default")

		// But we can still access it (will reload)
		a1, _ := f.GetCellValue(sheetName, "A1")
		assert.Equal(t, "Data", a1)
		fmt.Printf("重新加载后: A1='%s' ✅\n", a1)
	})

	t.Run("With KeepWorksheetInMemory", func(t *testing.T) {
		f := NewFile()
		f.options = &Options{KeepWorksheetInMemory: true}
		defer f.Close()

		sheetName := "Sheet1"
		sheetXMLPath, _ := f.getSheetXMLPath(sheetName)

		// Setup data
		f.SetCellValue(sheetName, "A1", "Data")

		// InsertRows
		err := f.InsertRows(sheetName, 2, 1)
		assert.NoError(t, err)

		ws1, loaded1 := f.Sheet.Load(sheetXMLPath)
		assert.True(t, loaded1)
		ptr1 := fmt.Sprintf("%p", ws1)
		fmt.Printf("\nInsertRows 后: Worksheet 在内存, Ptr=%s\n", ptr1)

		// Write (with KeepWorksheetInMemory)
		var buf []byte
		_, err = f.WriteTo(&mockWriter{&buf})
		assert.NoError(t, err)

		// Check if still loaded
		ws2, loaded2 := f.Sheet.Load(sheetXMLPath)
		fmt.Printf("Write 后 (KeepMemory): Worksheet 在内存=%v\n", loaded2)
		assert.True(t, loaded2, "Should STILL be loaded with KeepWorksheetInMemory")

		ptr2 := fmt.Sprintf("%p", ws2)
		assert.Equal(t, ptr1, ptr2, "Should be same instance")
		fmt.Printf("Ptr 相同=%v ✅\n", ptr1 == ptr2)
	})
}

// mockWriter for testing Write operations
type mockWriter struct {
	buf *[]byte
}

func (m *mockWriter) Write(p []byte) (n int, err error) {
	*m.buf = append(*m.buf, p...)
	return len(p), nil
}
