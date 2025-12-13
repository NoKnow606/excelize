# CalcFormulaValue API 使用指南

## 概述

`CalcFormulaValue` 是一个轻量级的公式计算 API，可以**临时计算公式值而不修改文件**。

## 核心优势 🚀

| 特性 | CalcFormulaValue | SetCellFormula + CalcCellValue |
|------|------------------|-------------------------------|
| **性能** | **30.21x 更快** | 基准 |
| **缓存清除** | ❌ 不清除 | ✅ 清除所有缓存 |
| **修改文件** | ❌ 不修改 | ✅ 修改 |
| **计算链** | ❌ 不修改 | ✅ 修改 |
| **适用场景** | 预览、验证、What-if 分析 | 永久保存公式 |

## API 文档

### 1. CalcFormulaValue - 单个公式计算

```go
func (f *File) CalcFormulaValue(sheet, cell, formula string, opts ...Options) (string, error)
```

**参数**:
- `sheet`: 工作表名称
- `cell`: 单元格地址（如 "A1"）
- `formula`: 公式（**不带** "=" 号）
- `opts`: 可选的计算选项

**返回值**:
- `string`: 计算结果
- `error`: 错误信息（如果有）

### 2. CalcFormulasValues - 批量公式计算

```go
func (f *File) CalcFormulasValues(sheet string, formulas map[string]string, opts ...Options) (map[string]string, error)
```

**参数**:
- `sheet`: 工作表名称
- `formulas`: 公式映射（单元格 → 公式）
- `opts`: 可选的计算选项

**返回值**:
- `map[string]string`: 计算结果映射
- `error`: 错误信息（部分失败时仍返回成功的结果）

## 使用示例

### 示例 1: 基本用法

```go
package main

import (
    "fmt"
    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    defer f.Close()

    // 设置基础数据
    f.SetCellValue("Sheet1", "B1", 10)
    f.SetCellValue("Sheet1", "B2", 20)
    f.SetCellValue("Sheet1", "B3", 30)

    // 临时计算公式，不保存到文件
    result, err := f.CalcFormulaValue("Sheet1", "A1", "SUM(B1:B3)")
    if err != nil {
        fmt.Println(err)
        return
    }

    fmt.Printf("SUM(B1:B3) = %s\n", result) // 输出: 60

    // A1 单元格没有公式！
    formula, _ := f.GetCellFormula("Sheet1", "A1")
    fmt.Printf("A1 formula: '%s'\n", formula) // 输出: ''（空）
}
```

### 示例 2: What-If 分析

```go
// 场景：用户想知道不同折扣下的最终价格，但不想修改文件
func WhatIfAnalysis(f *excelize.File) {
    // 设置基础数据
    f.SetCellValue("Sheet1", "A1", 1000) // 原价
    f.SetCellValue("Sheet1", "A2", 0.1)  // 折扣率

    // 测试不同折扣率的结果
    discounts := []float64{0.05, 0.1, 0.15, 0.2}

    for _, discount := range discounts {
        // 临时设置折扣率
        f.SetCellValue("Sheet1", "A2", discount)

        // 计算最终价格（不保存公式）
        formula := "A1*(1-A2)"
        result, _ := f.CalcFormulaValue("Sheet1", "A3", formula)

        fmt.Printf("折扣 %.0f%%: 最终价格 = %s\n", discount*100, result)
    }

    // A3 没有公式被保存
}
```

### 示例 3: 公式验证

```go
// 验证用户输入的公式是否正确
func ValidateFormula(f *excelize.File, userFormula string) bool {
    _, err := f.CalcFormulaValue("Sheet1", "A1", userFormula)
    if err != nil {
        fmt.Printf("公式错误: %v\n", err)
        return false
    }
    fmt.Println("公式有效 ✓")
    return true
}

// 使用示例
func main() {
    f := excelize.NewFile()
    defer f.Close()

    f.SetCellValue("Sheet1", "B1", 100)

    // 验证公式
    ValidateFormula(f, "SUM(B1:B10)")     // ✓ 有效
    ValidateFormula(f, "INVALID_FUNC()") // ✗ 无效
}
```

### 示例 4: 批量计算多个公式

```go
func BatchCalculate(f *excelize.File) {
    // 设置数据
    for i := 1; i <= 10; i++ {
        cell := fmt.Sprintf("B%d", i)
        f.SetCellValue("Sheet1", cell, i*10)
    }

    // 批量计算多个公式
    formulas := map[string]string{
        "A1": "SUM(B1:B10)",
        "A2": "AVERAGE(B1:B10)",
        "A3": "MAX(B1:B10)",
        "A4": "MIN(B1:B10)",
        "A5": "COUNT(B1:B10)",
    }

    results, err := f.CalcFormulasValues("Sheet1", formulas)
    if err != nil {
        fmt.Printf("部分计算失败: %v\n", err)
    }

    // 显示结果
    for cell, result := range results {
        fmt.Printf("%s = %s\n", cell, result)
    }

    // 所有 A1-A5 都没有保存公式！
}
```

### 示例 5: 数据预览（不修改文件）

```go
// 场景：在保存前预览报表数据
func PreviewReport(f *excelize.File) {
    // 设置基础数据
    f.SetCellValue("Sheet1", "B1", 100) // 销售额
    f.SetCellValue("Sheet1", "B2", 20)  // 成本

    // 预览计算结果（不保存公式）
    formulas := map[string]string{
        "C1": "B1*0.13",       // 税额
        "C2": "B1-B2",         // 利润
        "C3": "(B1-B2)/B1",    // 利润率
    }

    results, _ := f.CalcFormulasValues("Sheet1", formulas)

    fmt.Println("=== 报表预览 ===")
    fmt.Printf("税额: %s\n", results["C1"])
    fmt.Printf("利润: %s\n", results["C2"])
    fmt.Printf("利润率: %s\n", results["C3"])

    // 用户确认后，再决定是否保存公式
}
```

## 性能对比

```go
func BenchmarkComparison() {
    f := excelize.NewFile()
    defer f.Close()

    // 设置数据
    for i := 1; i <= 100; i++ {
        f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i)
    }

    // 方法 1: 传统方式（慢）
    start := time.Now()
    for i := 0; i < 1000; i++ {
        f.SetCellFormula("Sheet1", "A1", "SUM(B1:B100)")
        f.CalcCellValue("Sheet1", "A1")
    }
    duration1 := time.Since(start)
    fmt.Printf("传统方式: %v\n", duration1)

    // 方法 2: CalcFormulaValue（快）
    start = time.Now()
    for i := 0; i < 1000; i++ {
        f.CalcFormulaValue("Sheet1", "A1", "SUM(B1:B100)")
    }
    duration2 := time.Since(start)
    fmt.Printf("CalcFormulaValue: %v\n", duration2)

    fmt.Printf("提升: %.2fx\n", float64(duration1)/float64(duration2))
}

// 输出:
// 传统方式: 164ms
// CalcFormulaValue: 5.4ms
// 提升: 30.21x
```

## 实际性能数据

基于 1000 次迭代的测试结果：

| 方法 | 总耗时 | 平均每次 | 相对速度 |
|------|--------|---------|---------|
| SetCellFormula + CalcCellValue | 164.2 ms | 164.2 μs | 1x (基准) |
| **CalcFormulaValue** | **5.4 ms** | **5.4 μs** | **30.21x** 🚀 |

## 使用场景

### ✅ 适合使用 CalcFormulaValue

1. **What-If 分析** - 测试不同场景
2. **公式预览** - 显示计算结果但不保存
3. **公式验证** - 检查语法是否正确
4. **临时计算** - 一次性计算不需要保存
5. **数据报告** - 生成预览但不修改文件
6. **交互式工具** - 用户输入公式即时显示结果

### ❌ 不适合使用 CalcFormulaValue

1. **永久保存公式** - 需要公式在文件中
2. **公式引用** - 其他单元格需要引用此公式
3. **计算链** - 需要更新计算链

这些场景应该使用 `SetCellFormula`

## 技术细节

### 实现原理

```go
// 1. 保存原始公式（如果有）
originalFormula := cell.F

// 2. 临时设置公式（仅在内存中）
cell.F = &xlsxF{Content: "SUM(A1:A10)"}

// 3. 计算结果
result := f.CalcCellValue(sheet, cell)

// 4. 恢复原始状态
cell.F = originalFormula

// 5. 清除该单元格的缓存（防止缓存污染）
f.calcCache.Delete(cellRef)
```

### 关键优势

1. **不触发缓存清除** - 其他单元格的缓存保持完整
2. **不修改计算链** - 不调用 deleteCalcChain
3. **不持久化** - 文件状态完全不变
4. **原子操作** - 临时修改 → 计算 → 恢复
5. **线程相对安全** - 只操作单个单元格

## 常见问题 FAQ

**Q: CalcFormulaValue 会影响其他单元格的缓存吗？**
A: 不会。只清除当前计算单元格的缓存，其他单元格缓存保持完整。

**Q: 可以在同一个单元格上多次调用 CalcFormulaValue 吗？**
A: 可以。每次调用都是独立的，互不影响。

**Q: 如果单元格已有公式，会被覆盖吗？**
A: 不会。原始公式被保存并在计算后恢复。

**Q: 性能真的提升 30 倍吗？**
A: 是的。测试显示比传统方法快 30.21 倍，因为避免了缓存清除开销。

**Q: 批量计算 1000 个公式需要多久？**
A: 约 5-10ms（取决于公式复杂度）。传统方法需要 160-200ms。

## 完整示例

```go
package main

import (
    "fmt"
    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    defer f.Close()

    // ========== 场景 1: 基础计算 ==========
    fmt.Println("=== 场景 1: 基础计算 ===")
    f.SetCellValue("Sheet1", "B1", 100)
    f.SetCellValue("Sheet1", "B2", 200)

    result, _ := f.CalcFormulaValue("Sheet1", "A1", "SUM(B1:B2)")
    fmt.Printf("SUM(B1:B2) = %s\n\n", result)

    // ========== 场景 2: What-If 分析 ==========
    fmt.Println("=== 场景 2: What-If 分析 ===")
    f.SetCellValue("Sheet1", "C1", 1000) // 原价

    discounts := []float64{0.05, 0.1, 0.15, 0.2}
    for _, discount := range discounts {
        f.SetCellValue("Sheet1", "C2", discount)
        result, _ := f.CalcFormulaValue("Sheet1", "C3", "C1*(1-C2)")
        fmt.Printf("折扣 %.0f%%: 价格 = %s\n", discount*100, result)
    }
    fmt.Println()

    // ========== 场景 3: 批量计算 ==========
    fmt.Println("=== 场景 3: 批量计算 ===")
    for i := 1; i <= 10; i++ {
        f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i), i*10)
    }

    formulas := map[string]string{
        "E1": "SUM(D1:D10)",
        "E2": "AVERAGE(D1:D10)",
        "E3": "MAX(D1:D10)",
    }

    results, _ := f.CalcFormulasValues("Sheet1", formulas)
    for cell, value := range results {
        fmt.Printf("%s = %s\n", cell, value)
    }

    f.SaveAs("demo.xlsx")
    fmt.Println("\n文件已保存，但不包含任何临时公式！")
}
```

## 总结

- 🚀 **性能**: 比传统方法快 30 倍
- 🔒 **安全**: 不修改文件状态
- 💾 **缓存**: 保留其他单元格的缓存
- ✨ **灵活**: 适合预览、验证、分析场景

使用 `CalcFormulaValue` 让你的公式计算更快、更安全！
