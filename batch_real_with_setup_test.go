package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestRealFileWithCalcChainSetup 先建立 calcChain 再测试
func TestRealFileWithCalcChainSetup(t *testing.T) {
	f, err := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	if err != nil {
		t.Fatalf("打开文件失败: %v", err)
	}
	defer f.Close()

	inventorySheet := "库存台账-all"
	dailySheet := "日库存"

	fmt.Println("\n=== 步骤1: 读取现有公式 ===")
	b2Formula, _ := f.GetCellFormula(dailySheet, "B2")
	fmt.Printf("日库存!B2 公式: %s\n", b2Formula)

	// 步骤2: 使用 BatchSetFormulasAndRecalculate 重新设置公式以建立 calcChain
	fmt.Println("\n=== 步骤2: 重新设置公式以建立 calcChain ===")
	formulas := []FormulaUpdate{
		{Sheet: dailySheet, Cell: "B2", Formula: b2Formula},
	}
	_, err = f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)
	fmt.Println("✅ calcChain 已建立")

	// 步骤3: 验证 calcChain
	calcChain, _ := f.calcChainReader()
	if calcChain != nil && len(calcChain.C) > 0 {
		fmt.Printf("✅ calcChain 中有 %d 个公式\n", len(calcChain.C))
	}

	// 步骤4: 追加行并测试
	fmt.Println("\n=== 步骤3: 追加行到库存台账-all ===")
	updates := []CellUpdate{
		{Sheet: inventorySheet, Cell: "A4", Value: 999},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("受影响的单元格数: %d\n", len(affected))
	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}

	// 验证 B2 是否被更新
	foundB2 := false
	for _, cell := range affected {
		if cell.Sheet == dailySheet && cell.Cell == "B2" {
			foundB2 = true
			fmt.Printf("\n✅ 日库存!B2 已更新\n")
		}
	}

	if !foundB2 {
		fmt.Printf("\n❌ 日库存!B2 未在受影响列表中\n")
	}

	assert.True(t, foundB2, "日库存!B2 应该在受影响列表中")
}
