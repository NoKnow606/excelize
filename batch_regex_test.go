package excelize

import (
	"fmt"
	"regexp"
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
