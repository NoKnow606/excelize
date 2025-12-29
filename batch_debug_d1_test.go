package excelize

import (
	"fmt"
	"testing"
)

// TestDebugD1 调试 D1 为什么没有被更新
func TestDebugD1(t *testing.T) {
	f, _ := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	defer f.Close()

	f.RebuildCalcChain()

	// 检查 calcChain 中是否有 D1
	calcChain, _ := f.calcChainReader()
	fmt.Println("\n=== calcChain 中的 日库存 公式 ===")
	for _, c := range calcChain.C {
		sheetName := f.GetSheetName(c.I)
		if sheetName == "日库存" && (c.R == "B1" || c.R == "C1" || c.R == "D1") {
			fmt.Printf("%s (SheetID=%d)\n", c.R, c.I)
		}
	}

	// 更新并查看受影响的单元格
	updates := []CellUpdate{
		{Sheet: "库存台账-all", Cell: "A4", Value: 999},
	}

	affected, _ := f.BatchUpdateAndRecalculate(updates)

	fmt.Println("\n=== 日库存 中受影响的单元格 ===")
	for _, cell := range affected {
		if cell.Sheet == "日库存" && (cell.Cell == "B1" || cell.Cell == "C1" || cell.Cell == "D1") {
			fmt.Printf("%s = %s\n", cell.Cell, cell.CachedValue)
		}
	}
}
