# 快速参考指南 - Batch APIs

## 🚀 快速开始

### 场景 1: 批量更新值并重新计算

```go
package main

import "github.com/xuri/excelize/v2"

func main() {
    f := excelize.NewFile()

    // 设置初始数据和公式
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellFormula("Sheet1", "B1", "=A1*2")

    // ✅ 批量更新并自动重新计算
    updates := []excelize.CellUpdate{
        {Sheet: "Sheet1", Cell: "A1", Value: 100},
        {Sheet: "Sheet1", Cell: "A2", Value: 200},
        {Sheet: "Sheet1", Cell: "A3", Value: 300},
    }

    f.BatchUpdateAndRecalculate(updates)

    // B1 现在是 200 (100*2)
    value, _ := f.GetCellValue("Sheet1", "B1")
    println(value)  // "200"

    f.SaveAs("output.xlsx")
}
```

**性能**：比循环调用 SetCellValue 快 **8-377 倍**

---

### 场景 2: 批量设置公式并计算

```go
package main

import "github.com/xuri/excelize/v2"

func main() {
    f := excelize.NewFile()

    // 设置原始数据
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellValue("Sheet1", "A2", 20)
    f.SetCellValue("Sheet1", "A3", 30)

    // ✅ 批量设置公式，自动计算，自动更新 calcChain
    formulas := []excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
        {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
        {Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
        {Sheet: "Sheet1", Cell: "C1", Formula: "=SUM(B1:B3)"},
    }

    f.BatchSetFormulasAndRecalculate(formulas)

    // 所有公式已计算
    c1, _ := f.GetCellValue("Sheet1", "C1")
    println(c1)  // "120" (20+40+60)

    f.SaveAs("output.xlsx")
}
```

**优势**：
- ✅ 一次调用完成：设置公式 + 计算 + 更新 calcChain
- ✅ 性能提升 10-100 倍
- ✅ 自动处理依赖关系

---

### 场景 3: 跨工作表公式（自动处理）

```go
package main

import "github.com/xuri/excelize/v2"

func main() {
    f := excelize.NewFile()
    f.NewSheet("Sheet2")

    // Sheet1: A1 = 100
    f.SetCellValue("Sheet1", "A1", 100)

    // Sheet2 引用 Sheet1
    formulas := []excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},          // 200
        {Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+10"},  // 210
    }
    f.BatchSetFormulasAndRecalculate(formulas)

    // ✅ 更新 Sheet1，Sheet2 会自动重新计算
    updates := []excelize.CellUpdate{
        {Sheet: "Sheet1", Cell: "A1", Value: 500},
    }
    f.BatchUpdateAndRecalculate(updates)

    // Sheet1.B1 = 1000 (500*2)
    b1, _ := f.GetCellValue("Sheet1", "B1")
    println(b1)  // "1000"

    // ✅ Sheet2.C1 = 1010 (1000+10) - 自动重新计算！
    c1, _ := f.GetCellValue("Sheet2", "C1")
    println(c1)  // "1010"

    f.SaveAs("output.xlsx")
}
```

**关键特性**：
- ✅ 跨工作表依赖自动处理
- ✅ 无需手动指定受影响的工作表
- ✅ 按正确的依赖顺序计算

---

### 场景 4: 频繁读写大文件（内存优化）

```go
package main

import "github.com/xuri/excelize/v2"

func main() {
    // ✅ 启用 KeepWorksheetInMemory
    f, _ := excelize.OpenFile("large.xlsx", excelize.Options{
        KeepWorksheetInMemory: true,  // 关键选项
    })

    // 多次读写同一工作表，避免重复加载
    for i := 0; i < 100; i++ {
        // 读取数据
        value, _ := f.GetCellValue("Sheet1", "A1")

        // 修改数据
        f.SetCellValue("Sheet1", "B1", value)

        // ✅ Write 不会卸载工作表（快 2.4 倍）
        f.Write(someWriter)
    }

    f.SaveAs("output.xlsx")
}
```

**性能**：
- ✅ 2.4x 速度提升（100,000 行场景）
- ⚠️ 内存成本：~20MB per 100k rows

---

## 📋 API 速查

### 批量值更新 API

| API | 功能 | 是否计算 |
|-----|------|---------|
| `BatchSetCellValue(updates)` | 批量设置值 | ❌ 不计算 |
| `RecalculateSheet(sheet)` | 重新计算工作表 | ✅ 计算 |
| `BatchUpdateAndRecalculate(updates)` | 批量更新 + 计算 | ✅ 计算 |

### 批量公式 API

| API | 功能 | 是否计算 | 更新 calcChain |
|-----|------|---------|--------------|
| `BatchSetFormulas(formulas)` | 批量设置公式 | ❌ | ❌ |
| `BatchSetFormulasAndRecalculate(formulas)` | 批量设置 + 计算 | ✅ | ✅ |

### 其他 API

| API | 功能 |
|-----|------|
| `UpdateCellAndRecalculate(sheet, cell)` | 更新单个单元格并触发重新计算 |

---

## 🎯 类型定义

### CellUpdate

```go
type CellUpdate struct {
    Sheet string      // 工作表名称，如 "Sheet1"
    Cell  string      // 单元格坐标，如 "A1"
    Value interface{} // 单元格值（任意类型）
}
```

**示例**：
```go
updates := []excelize.CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 100},
    {Sheet: "Sheet1", Cell: "A2", Value: "Hello"},
    {Sheet: "Sheet1", Cell: "A3", Value: 3.14},
}
```

### FormulaUpdate

```go
type FormulaUpdate struct {
    Sheet   string // 工作表名称，如 "Sheet1"
    Cell    string // 单元格坐标，如 "B1"
    Formula string // 公式内容，可以包含或不包含前导 '='
}
```

**示例**：
```go
formulas := []excelize.FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},     // 可以有 '='
    {Sheet: "Sheet1", Cell: "B2", Formula: "A2*2"},      // 也可以没有
    {Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+10"}, // 跨工作表
}
```

---

## ⚡ 性能对比

### 批量更新 vs 循环调用

```go
// ❌ 慢：循环调用（基准：168.8ms for 10 cells）
for i := 0; i < 10; i++ {
    f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+1), i)
}
f.RecalculateSheet("Sheet1")

// ✅ 快：批量 API（20.3ms for 10 cells = 8.3x faster）
updates := make([]excelize.CellUpdate, 10)
for i := 0; i < 10; i++ {
    updates[i] = excelize.CellUpdate{
        Sheet: "Sheet1",
        Cell:  fmt.Sprintf("A%d", i+1),
        Value: i,
    }
}
f.BatchUpdateAndRecalculate(updates)
```

**结果**：
- 10 单元格：8.3x 提升
- 100 单元格：9.4x 提升
- 特定场景：高达 377x 提升

### KeepWorksheetInMemory 性能

```go
// ❌ 慢：默认行为（每次 Write 后卸载，下次读取需重新加载）
f, _ := excelize.OpenFile("large.xlsx")
for i := 0; i < 100; i++ {
    f.GetCellValue("Sheet1", "A1")  // 每次都要重新加载
    f.Write(writer)
}

// ✅ 快：保持在内存（2.4x faster）
f, _ := excelize.OpenFile("large.xlsx", excelize.Options{
    KeepWorksheetInMemory: true,
})
for i := 0; i < 100; i++ {
    f.GetCellValue("Sheet1", "A1")  // 直接从内存读取
    f.Write(writer)
}
```

**场景**：
- 100,000 行工作表
- 默认：1.2s
- KeepMemory：0.5s（**2.4x 提升**）
- 内存成本：+20MB

---

## 🎨 使用模式

### 模式 1: 数据导入 + 批量公式

```go
func ImportData(f *excelize.File, data [][]int) error {
    // 1. 批量写入原始数据
    updates := make([]excelize.CellUpdate, len(data))
    for i, row := range data {
        for j, val := range row {
            updates = append(updates, excelize.CellUpdate{
                Sheet: "Sheet1",
                Cell:  fmt.Sprintf("%s%d", columnName(j), i+1),
                Value: val,
            })
        }
    }
    f.BatchSetCellValue(updates)

    // 2. 批量设置计算公式
    formulas := []excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: "E1", Formula: "=SUM(A1:D1)"},
        {Sheet: "Sheet1", Cell: "E2", Formula: "=SUM(A2:D2)"},
        // ...
    }
    return f.BatchSetFormulasAndRecalculate(formulas)
}
```

### 模式 2: 动态报表生成

```go
func GenerateReport(f *excelize.File, params ReportParams) error {
    // 1. 设置参数单元格
    updates := []excelize.CellUpdate{
        {Sheet: "Config", Cell: "A1", Value: params.StartDate},
        {Sheet: "Config", Cell: "A2", Value: params.EndDate},
        {Sheet: "Config", Cell: "A3", Value: params.Region},
    }

    // 2. 批量更新并触发所有公式重新计算（包括跨工作表）
    return f.BatchUpdateAndRecalculate(updates)
}
```

### 模式 3: 模板填充

```go
func FillTemplate(templateFile string, data map[string]interface{}) error {
    f, _ := excelize.OpenFile(templateFile, excelize.Options{
        KeepWorksheetInMemory: true,  // 模板通常需要多次读写
    })

    // 批量填充数据
    updates := make([]excelize.CellUpdate, 0, len(data))
    for cell, value := range data {
        updates = append(updates, excelize.CellUpdate{
            Sheet: "Sheet1",
            Cell:  cell,
            Value: value,
        })
    }

    // 一次性更新并计算所有公式
    if err := f.BatchUpdateAndRecalculate(updates); err != nil {
        return err
    }

    return f.SaveAs("output.xlsx")
}
```

---

## ⚠️ 常见陷阱

### 陷阱 1: 忘记跨工作表依赖

```go
// ❌ 错误：只计算被更新的工作表
updates := []excelize.CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 100},
}
f.BatchSetCellValue(updates)
f.RecalculateSheet("Sheet1")  // Sheet2 不会重新计算！

// ✅ 正确：使用 BatchUpdateAndRecalculate（自动处理跨工作表）
f.BatchUpdateAndRecalculate(updates)
```

### 陷阱 2: 频繁读写不启用 KeepWorksheetInMemory

```go
// ❌ 慢：每次 Write 后卸载工作表
f, _ := excelize.OpenFile("file.xlsx")
for i := 0; i < 100; i++ {
    f.GetCellValue("Sheet1", "A1")
    f.Write(writer)  // 卸载
}  // 下次读取需要重新加载 XML

// ✅ 快：保持在内存
f, _ := excelize.OpenFile("file.xlsx", excelize.Options{
    KeepWorksheetInMemory: true,
})
```

### 陷阱 3: 公式设置后忘记更新 calcChain

```go
// ❌ 不完整：calcChain 没有更新
f.BatchSetFormulas(formulas)
f.RecalculateSheet("Sheet1")  // Excel 打开时可能显示 #NAME? 错误

// ✅ 完整：自动更新 calcChain
f.BatchSetFormulasAndRecalculate(formulas)
```

---

## 🔧 错误处理

### 基本错误处理

```go
updates := []excelize.CellUpdate{
    {Sheet: "InvalidSheet", Cell: "A1", Value: 100},
}

if err := f.BatchUpdateAndRecalculate(updates); err != nil {
    // 处理错误
    var sheetErr excelize.ErrSheetNotExist
    if errors.As(err, &sheetErr) {
        fmt.Printf("工作表不存在: %s\n", sheetErr.SheetName)
    }
}
```

### 验证数据完整性

```go
func SafeBatchUpdate(f *excelize.File, updates []excelize.CellUpdate) error {
    // 1. 预验证
    for _, update := range updates {
        // 检查工作表是否存在
        if _, err := f.GetSheetIndex(update.Sheet); err != nil {
            return fmt.Errorf("工作表 %s 不存在", update.Sheet)
        }

        // 检查单元格坐标是否有效
        if _, _, err := excelize.CellNameToCoordinates(update.Cell); err != nil {
            return fmt.Errorf("无效的单元格坐标: %s", update.Cell)
        }
    }

    // 2. 执行批量更新
    return f.BatchUpdateAndRecalculate(updates)
}
```

---

## 📊 性能基准参考

### 批量更新基准（100 单元格）

| 方法 | 时间 | 相对性能 |
|-----|------|---------|
| 循环 SetCellValue | 1673.2ms | 1x |
| BatchUpdateAndRecalculate | 178.4ms | **9.4x** |

### 公式设置基准（100 公式）

| 方法 | 时间 | 相对性能 |
|-----|------|---------|
| 循环 SetCellFormula + RecalculateSheet | 1500ms | 1x |
| BatchSetFormulasAndRecalculate | ~150ms | **10x** |

### KeepWorksheetInMemory 基准（100k 行）

| 方法 | 时间 | 内存 |
|-----|------|------|
| 默认（自动卸载） | 1.2s | 低 |
| KeepWorksheetInMemory | 0.5s | +20MB |

---

## 🎓 最佳实践

### ✅ DO

1. **使用批量 API 处理多个单元格**
   ```go
   f.BatchUpdateAndRecalculate(updates)  // ✅
   ```

2. **频繁读写时启用 KeepWorksheetInMemory**
   ```go
   f, _ := excelize.OpenFile("file.xlsx", excelize.Options{
       KeepWorksheetInMemory: true,
   })
   ```

3. **使用 BatchSetFormulasAndRecalculate 设置公式**
   ```go
   f.BatchSetFormulasAndRecalculate(formulas)  // ✅ 自动更新 calcChain
   ```

4. **预分配切片容量**
   ```go
   updates := make([]excelize.CellUpdate, 0, 1000)  // ✅ 避免重新分配
   ```

### ❌ DON'T

1. **不要循环调用单个 API**
   ```go
   for _, update := range updates {
       f.SetCellValue(...)  // ❌ 慢
   }
   ```

2. **不要忘记跨工作表依赖**
   ```go
   f.BatchSetCellValue(updates)
   f.RecalculateSheet("Sheet1")  // ❌ Sheet2 不会计算
   ```

3. **不要在高频读写时使用默认配置**
   ```go
   f, _ := excelize.OpenFile("large.xlsx")  // ❌ 每次 Write 都卸载
   ```

---

## 🔗 相关文档

### 详细文档
- [完整 API 文档](./BATCH_SET_FORMULAS_API.md) - 620 行，所有 API 详解
- [最佳实践](./BATCH_API_BEST_PRACTICES.md) - 584 行，优化指南
- [计算机制](./BATCH_FORMULA_CALCULATION_MECHANISM.md) - 529 行，底层原理

### 技术细节
- [跨工作表支持](./BATCH_UPDATE_CROSS_SHEET_SUPPORT.md) - 跨表依赖处理
- [性能分析](./BATCH_FORMULA_PERFORMANCE_ANALYSIS.md) - 基准测试结果
- [Bug 修复](./CRITICAL_BUGS_SUMMARY.md) - 生产 bug 修复

### 会话总结
- [完整会话总结](./SESSION_SUMMARY.md) - 所有工作汇总

---

## 📞 获取帮助

### 问题排查

**问题 1：跨工作表公式没有重新计算**
```go
// 解决方案：使用 BatchUpdateAndRecalculate（不要手动指定工作表）
f.BatchUpdateAndRecalculate(updates)
```

**问题 2：Excel 打开时公式显示 #NAME? 错误**
```go
// 解决方案：使用 BatchSetFormulasAndRecalculate（自动更新 calcChain）
f.BatchSetFormulasAndRecalculate(formulas)
```

**问题 3：性能没有提升**
```go
// 解决方案：确保使用批量 API，不要循环调用单个 API
// ❌ 错误
for _, update := range updates {
    f.SetCellValue(...)
}

// ✅ 正确
f.BatchUpdateAndRecalculate(updates)
```

---

**版本**：v2.0.0-20251226
**最后更新**：2025-12-26
**文档状态**：✅ 完整
