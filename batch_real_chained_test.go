package excelize

import (
	"fmt"
	"testing"
)

// TestRealFileChainedDependency 测试真实文件的链式依赖
func TestRealFileChainedDependency(t *testing.T) {
	f, err := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	if err != nil {
		t.Fatalf("打开文件失败: %v", err)
	}
	defer f.Close()

	// 重建 calcChain
	fmt.Println("重建 calcChain...")
	f.RebuildCalcChain()

	// 检查 B1, C1, D1 的公式
	sheets := []string{"日库存", "日销售", "日销预测", "补货计划"}

	for _, sheet := range sheets {
		fmt.Printf("\n=== %s ===\n", sheet)

		for _, cell := range []string{"B1", "C1", "D1"} {
			formula, _ := f.GetCellFormula(sheet, cell)
			value, _ := f.GetCellValue(sheet, cell)
			if formula != "" {
				fmt.Printf("%s: 公式='%s', 值='%s'\n", cell, formula, value)
			}
		}
	}

	// 测试更新
	fmt.Println("\n=== 测试更新库存台账-all!A4 ===")
	updates := []CellUpdate{
		{Sheet: "库存台账-all", Cell: "A4", Value: 999},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	if err != nil {
		t.Fatalf("更新失败: %v", err)
	}

	fmt.Printf("受影响的单元格数: %d\n", len(affected))

	// 按表分组显示
	sheetMap := make(map[string][]AffectedCell)
	for _, cell := range affected {
		sheetMap[cell.Sheet] = append(sheetMap[cell.Sheet], cell)
	}

	for sheet, cells := range sheetMap {
		fmt.Printf("\n%s: %d 个单元格\n", sheet, len(cells))
		for _, cell := range cells {
			fmt.Printf("  %s = %s\n", cell.Cell, cell.CachedValue)
		}
	}
}
