package main

import (
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	path := "test/real-ecomm/跨境电商-补货计划结果表.xlsx"
	f, err := excelize.OpenFile(path)
	if err != nil {
		log.Fatalf("open file failed: %v", err)
	}
	defer f.Close()

	if err := f.RecalculateAllWithDependency(); err != nil {
		log.Fatalf("recalculate failed: %v", err)
	}
	log.Println("recalculate completed successfully")
}
