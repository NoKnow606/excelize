package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestRealExcelFile 使用真实的 Excel 文件测试
func TestRealExcelFile(t *testing.T) {
	// 打开真实文件
	f, err := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	if err != nil {
		t.Fatalf("打开文件失败: %v", err)
	}
	defer f.Close()

	inventorySheet := "库存台账-all"
	dailySheet := "日库存"

	fmt.Println("\n=== 读取初始状态 ===")

	// 读取日库存表 B2 的公式
	b2Formula, err := f.GetCellFormula(dailySheet, "B2")
	assert.NoError(t, err)
	fmt.Printf("日库存表!B2 公式: %s\n", b2Formula)

	// 读取初始值
	b2Value, _ := f.GetCellValue(dailySheet, "B2")
	fmt.Printf("日库存表!B2 当前值: %s\n", b2Value)

	// 读取库存台账-all 的 A 列数据
	rows, err := f.GetRows(inventorySheet)
	assert.NoError(t, err)
	fmt.Printf("库存台账-all 行数: %d\n", len(rows))

	// 显示前几行
	for i := 0; i < minInt(5, len(rows)); i++ {
		if len(rows[i]) > 0 {
			fmt.Printf("  A%d = %s\n", i+1, rows[i][0])
		}
	}

	// 追加新行到库存台账-all
	newRow := len(rows) + 1
	fmt.Printf("\n=== 追加行到 A%d ===\n", newRow)

	updates := []CellUpdate{
		{Sheet: inventorySheet, Cell: fmt.Sprintf("A%d", newRow), Value: 999},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("受影响的单元格数: %d\n", len(affected))
	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}

	// 检查 B2 是否被更新
	foundB2 := false
	for _, cell := range affected {
		if cell.Sheet == dailySheet && cell.Cell == "B2" {
			foundB2 = true
			fmt.Printf("\n✅ 日库存表!B2 已更新: %s\n", cell.CachedValue)
		}
	}

	if !foundB2 {
		fmt.Printf("\n❌ 日库存表!B2 未在受影响列表中\n")

		// 调试信息
		fmt.Println("\n=== 调试信息 ===")
		fmt.Printf("公式: %s\n", b2Formula)
		fmt.Printf("更新的表: %s\n", inventorySheet)
		fmt.Printf("更新的单元格: A%d\n", newRow)
	}

	assert.True(t, foundB2, "日库存表!B2 应该在受影响列表中")
}

func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}
