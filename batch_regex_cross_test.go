package excelize

import (
	"fmt"
	"regexp"
	"testing"
)

// TestRegexCrossSheet 测试跨表引用的正则匹配
func TestRegexCrossSheet(t *testing.T) {
	formulas := []string{
		"日库存!B1",
		"'日库存'!B1",
		"Sheet1!A1",
		"'库存台账-all'!A:A",
	}

	cellRefPattern := regexp.MustCompile(`(?:'([^']+)'!|([A-Za-z_][A-Za-z0-9_.]*!))?(\$?[A-Z]+\$?[0-9]+)`)

	fmt.Println("\n=== 测试单元格引用正则 ===")
	for _, formula := range formulas {
		matches := cellRefPattern.FindAllStringSubmatch(formula, -1)
		fmt.Printf("\n公式: %s\n", formula)
		if len(matches) == 0 {
			fmt.Println("  ❌ 没有匹配")
		} else {
			for _, match := range matches {
				fmt.Printf("  表名(单引号): '%s'\n", match[1])
				fmt.Printf("  表名(普通): '%s'\n", match[2])
				fmt.Printf("  单元格: '%s'\n", match[3])
			}
		}
	}
}
