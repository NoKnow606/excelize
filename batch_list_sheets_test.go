package excelize

import (
	"fmt"
	"os"
	"testing"
)

// TestListSheets 列出文件中的所有表
func TestListSheets(t *testing.T) {
	path := os.Getenv("EXCELIZE_LIST_SHEETS_FILE")
	var f *File
	var err error
	if path == "" {
		f = NewFile()
		_, err = f.NewSheet("Sheet2")
	} else {
		f, err = OpenFile(path)
	}
	if err != nil {
		t.Fatalf("创建或打开文件失败: %v", err)
	}
	defer f.Close()

	fmt.Println("\n=== 文件中的所有工作表 ===")
	sheets := f.GetSheetList()
	for i, sheet := range sheets {
		fmt.Printf("%d. '%s'\n", i+1, sheet)
	}
}
