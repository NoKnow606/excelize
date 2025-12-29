package excelize

import (
	"fmt"
	"testing"
)

// TestRealFileDebug 调试真实文件
func TestRealFileDebug(t *testing.T) {
	f, err := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	if err != nil {
		t.Fatalf("打开文件失败: %v", err)
	}
	defer f.Close()

	inventorySheet := "库存台账-all"
	dailySheet := "日库存"

	// 读取公式
	b2Formula, _ := f.GetCellFormula(dailySheet, "B2")
	fmt.Printf("\n公式: %s\n", b2Formula)

	// 检查 calcChain
	calcChain, err := f.calcChainReader()
	if err != nil {
		fmt.Printf("读取 calcChain 失败: %v\n", err)
		return
	}

	if calcChain == nil || len(calcChain.C) == 0 {
		fmt.Println("❌ calcChain 为空或没有公式")
		return
	}

	fmt.Printf("\ncalcChain 中有 %d 个公式\n", len(calcChain.C))

	// 查找日库存!B2
	found := false
	for _, c := range calcChain.C {
		sheetID := c.I
		sheetName := f.GetSheetName(sheetID)
		if sheetName == dailySheet && c.R == "B2" {
			found = true
			fmt.Printf("✅ 找到 %s!%s 在 calcChain 中 (SheetID=%d)\n", sheetName, c.R, sheetID)
		}
	}

	if !found {
		fmt.Printf("❌ %s!B2 不在 calcChain 中\n", dailySheet)
		fmt.Println("\ncalcChain 内容:")
		for i, c := range calcChain.C {
			if i < 10 { // 只显示前10个
				sheetName := f.GetSheetName(c.I)
				fmt.Printf("  %d. %s!%s (SheetID=%d)\n", i+1, sheetName, c.R, c.I)
			}
		}
	}

	// 测试更新
	fmt.Println("\n=== 测试更新 ===")
	updates := []CellUpdate{
		{Sheet: inventorySheet, Cell: "A4", Value: 999},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	if err != nil {
		fmt.Printf("更新失败: %v\n", err)
		return
	}

	fmt.Printf("受影响的单元格数: %d\n", len(affected))
	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}
}
