package excelize

import (
	"fmt"
	"testing"
)

// TestListSheets 列出文件中的所有表
func TestListSheets(t *testing.T) {
	f, err := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	if err != nil {
		t.Fatalf("打开文件失败: %v", err)
	}
	defer f.Close()

	fmt.Println("\n=== 文件中的所有工作表 ===")
	sheets := f.GetSheetList()
	for i, sheet := range sheets {
		fmt.Printf("%d. '%s'\n", i+1, sheet)
	}
}
