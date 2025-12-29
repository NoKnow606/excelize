package excelize

import (
	"fmt"
	"testing"
)

// TestCrossSheetDependency 测试跨表依赖：库存台账-all → 日库存!B1 → 日销售!C1
func TestCrossSheetDependency(t *testing.T) {
	f, _ := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	defer f.Close()

	f.RebuildCalcChain()

	fmt.Println("\n=== 初始状态 ===")

	// 检查公式
	b1Formula, _ := f.GetCellFormula("日库存", "B1")
	c1Formula, _ := f.GetCellFormula("日销售", "C1")

	fmt.Printf("日库存!B1 公式: %s\n", b1Formula)
	fmt.Printf("日销售!C1 公式: %s\n", c1Formula)

	b1Val, _ := f.GetCellValue("日库存", "B1")
	c1Val, _ := f.GetCellValue("日销售", "C1")

	fmt.Printf("日库存!B1 值: %s\n", b1Val)
	fmt.Printf("日销售!C1 值: %s\n", c1Val)

	// 更新库存台账-all
	fmt.Println("\n=== 更新库存台账-all!A4 ===")
	updates := []CellUpdate{
		{Sheet: "库存台账-all", Cell: "A4", Value: "2025-12-30"},
	}

	affected, _ := f.BatchUpdateAndRecalculate(updates)

	fmt.Printf("受影响的单元格数: %d\n", len(affected))

	// 检查日库存和日销售
	foundB1 := false
	foundC1 := false

	for _, cell := range affected {
		if cell.Sheet == "日库存" && cell.Cell == "B1" {
			foundB1 = true
			fmt.Printf("✅ 日库存!B1 = %s\n", cell.CachedValue)
		}
		if cell.Sheet == "日销售" && cell.Cell == "C1" {
			foundC1 = true
			fmt.Printf("✅ 日销售!C1 = %s\n", cell.CachedValue)
		}
	}

	if !foundB1 {
		fmt.Println("❌ 日库存!B1 未在受影响列表中")
	}
	if !foundC1 {
		fmt.Println("❌ 日销售!C1 未在受影响列表中")
	}
}
