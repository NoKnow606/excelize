package excelize

import (
	"fmt"
	"regexp"
	"strings"
	"testing"
)

// TestRegexPatternMatching 测试正则表达式匹配
func TestRegexPatternMatching(t *testing.T) {
	formula := "SUMIFS('库存台账-all'!$H:$H,'库存台账-all'!$D:$D,$A2,'库存台账-all'!$A:$A,B$1)"

	fmt.Println("\n=== 测试列引用正则表达式 ===")
	fmt.Printf("公式: %s\n\n", formula)

	// 当前的正则表达式
	colRefPattern := regexp.MustCompile(`(?:'([^']+)'!|([A-Za-z_][A-Za-z0-9_.]*!))?(\$?[A-Z]+):(\$?[A-Z]+)`)
	matches := colRefPattern.FindAllStringSubmatch(formula, -1)

	fmt.Printf("找到 %d 个列引用:\n", len(matches))
	for i, match := range matches {
		fmt.Printf("%d. 完整匹配: '%s'\n", i+1, match[0])
		fmt.Printf("   表名(单引号): '%s'\n", match[1])
		fmt.Printf("   表名(普通): '%s'\n", match[2])
		fmt.Printf("   起始列: '%s'\n", match[3])
		fmt.Printf("   结束列: '%s'\n", match[4])
		fmt.Println()
	}
}

// TestCellReferenceRegex 测试单元格引用正则
func TestCellReferenceRegex(t *testing.T) {
	testFormulas := []string{
		"=Sheet1!B1+10",
		"='Sheet 1'!B1+10",
		"=B1+10",
		"=Sheet1!A1+Sheet1!A2",
	}

	fmt.Println("\n=== 测试单元格引用正则表达式 ===")

	// 当前使用的正则
	currentPattern := regexp.MustCompile(`(?:'([^']+)'!|([^\s\(\)!]+!))?(\$?[A-Z]+\$?[0-9]+)`)

	for _, formula := range testFormulas {
		fmt.Printf("\n公式: %s\n", formula)
		matches := currentPattern.FindAllStringSubmatch(formula, -1)

		for i, match := range matches {
			fmt.Printf("  匹配 %d:\n", i+1)
			fmt.Printf("    完整: '%s'\n", match[0])
			fmt.Printf("    单引号表名(1): '%s'\n", match[1])
			fmt.Printf("    普通表名(2): '%s'\n", match[2])
			fmt.Printf("    单元格(3): '%s'\n", match[3])

			refSheet := "currentSheet"
			if match[1] != "" {
				refSheet = match[1]
			} else if match[2] != "" {
				refSheet = strings.TrimSuffix(match[2], "!")
			}
			fmt.Printf("    => 表名: '%s', 单元格: '%s'\n", refSheet, strings.ReplaceAll(match[3], "$", ""))
		}
	}
}
