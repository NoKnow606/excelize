package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestRebuildCalcChain 测试重建 calcChain
func TestRebuildCalcChain(t *testing.T) {
	f, err := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	assert.NoError(t, err)
	defer f.Close()

	fmt.Println("\n=== 方案1: 使用 RebuildCalcChain ===")

	// 步骤1: 重建 calcChain
	fmt.Println("重建 calcChain...")
	err = f.RebuildCalcChain()
	assert.NoError(t, err)

	calcChain, _ := f.calcChainReader()
	fmt.Printf("✅ calcChain 中有 %d 个公式\n", len(calcChain.C))

	// 步骤2: 追加行
	fmt.Println("\n追加行到库存台账-all...")
	updates := []CellUpdate{
		{Sheet: "库存台账-all", Cell: "A4", Value: 999},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("受影响的单元格数: %d\n", len(affected))

	// 验证
	foundB2 := false
	for _, cell := range affected {
		if cell.Sheet == "日库存" && cell.Cell == "B2" {
			foundB2 = true
			fmt.Printf("✅ 日库存!B2 已更新: %s\n", cell.CachedValue)
		}
	}

	assert.True(t, foundB2, "日库存!B2 应该在受影响列表中")
}
