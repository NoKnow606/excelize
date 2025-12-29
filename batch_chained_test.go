package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestChainedDependency 测试链式依赖：A1 -> B1 -> C1 -> D1
func TestChainedDependency(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"

	// 设置初始值
	f.SetCellValue(sheet, "A1", 10)

	// 设置链式公式：B1=A1*2, C1=B1+10, D1=C1*2
	formulas := []FormulaUpdate{
		{Sheet: sheet, Cell: "B1", Formula: "A1*2"},
		{Sheet: sheet, Cell: "C1", Formula: "B1+10"},
		{Sheet: sheet, Cell: "D1", Formula: "C1*2"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== 初始状态 ===")
	fmt.Printf("A1 = %s\n", mustGetValue(f, sheet, "A1"))
	fmt.Printf("B1 = A1*2 = %s\n", mustGetValue(f, sheet, "B1"))
	fmt.Printf("C1 = B1+10 = %s\n", mustGetValue(f, sheet, "C1"))
	fmt.Printf("D1 = C1*2 = %s\n", mustGetValue(f, sheet, "D1"))

	// 更新 A1
	updates := []CellUpdate{
		{Sheet: sheet, Cell: "A1", Value: 100},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\n=== 更新 A1=100 后 ===\n")
	fmt.Printf("受影响的单元格数: %d\n", len(affected))

	affectedMap := make(map[string]string)
	for _, cell := range affected {
		affectedMap[cell.Cell] = cell.CachedValue
		fmt.Printf("  %s = %s\n", cell.Cell, cell.CachedValue)
	}

	// 验证所有单元格都被更新
	assert.Contains(t, affectedMap, "B1", "B1 应该被更新")
	assert.Contains(t, affectedMap, "C1", "C1 应该被更新")
	assert.Contains(t, affectedMap, "D1", "D1 应该被更新")

	// 验证值
	assert.Equal(t, "200", affectedMap["B1"], "B1 = 100*2 = 200")
	assert.Equal(t, "210", affectedMap["C1"], "C1 = 200+10 = 210")
	assert.Equal(t, "420", affectedMap["D1"], "D1 = 210*2 = 420")

	fmt.Println("\n✅ 链式依赖测试通过")
}
