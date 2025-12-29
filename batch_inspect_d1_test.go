package excelize

import (
	"fmt"
	"testing"
)

// TestInspectD1 检查 D1 的详细信息
func TestInspectD1(t *testing.T) {
	f, _ := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	defer f.Close()

	sheet := "日库存"

	// 读取 D1 的公式
	formula, _ := f.GetCellFormula(sheet, "D1")
	fmt.Printf("D1 公式: '%s'\n", formula)

	// 读取工作表
	ws, _ := f.workSheetReader(sheet)

	// 查找 D1
	col, row, _ := CellNameToCoordinates("D1")
	cellData := f.getCellFromWorksheet(ws, col, row)

	if cellData == nil {
		fmt.Println("❌ D1 cellData 为 nil")
		return
	}

	fmt.Printf("\nD1 详细信息:\n")
	fmt.Printf("  R: %s\n", cellData.R)
	if cellData.F != nil {
		fmt.Printf("  F.Content: '%s'\n", cellData.F.Content)
		fmt.Printf("  F.T: %s\n", cellData.F.T)
		if cellData.F.Si != nil {
			fmt.Printf("  F.Si: %d\n", *cellData.F.Si)
		}
	} else {
		fmt.Println("  F: nil")
	}

	// 扫描所有行查找 D1
	fmt.Println("\n扫描 SheetData.Row:")
	found := false
	for i, row := range ws.SheetData.Row {
		for j, cell := range row.C {
			if cell.R == "D1" {
				found = true
				fmt.Printf("  找到 D1 在 Row[%d].C[%d]\n", i, j)
				if cell.F != nil {
					fmt.Printf("    F.Content: '%s'\n", cell.F.Content)
				}
			}
		}
	}

	if !found {
		fmt.Println("  ❌ 在 SheetData.Row 中没有找到 D1")
	}
}
