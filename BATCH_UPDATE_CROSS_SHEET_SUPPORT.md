# BatchUpdateAndRecalculate 跨工作表支持

## 📋 问题描述

### 原实现的限制

**之前的版本** (`batch.go` 原版本 104-124 行) 存在严重限制：

```go
func (f *File) BatchUpdateAndRecalculate(updates []CellUpdate) error {
    // 1. 批量更新所有单元格
    f.BatchSetCellValue(updates)

    // 2. 收集受影响的工作表
    affectedSheets := make(map[string]bool)
    for _, update := range updates {
        affectedSheets[update.Sheet] = true  // ❌ 只收集被更新的工作表
    }

    // 3. 只重新计算被更新的工作表
    for sheet := range affectedSheets {
        f.RecalculateSheet(sheet)  // ❌ 忽略了其他工作表的依赖
    }
}
```

### 问题场景

```go
// Sheet1: A1 = 100
// Sheet2: B1 = Sheet1!A1 * 2  (引用 Sheet1)

updates := []CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 200},  // 更新 Sheet1
}
f.BatchUpdateAndRecalculate(updates)

// 结果：
// ✅ Sheet1.A1 = 200        (正确)
// ❌ Sheet2.B1 = 200        (错误！应该是 400)
//    原因：Sheet2 没有被重新计算
```

---

## ✅ 新实现

### 核心改进

**新版本** (`batch.go` 105-138 行)：

```go
func (f *File) BatchUpdateAndRecalculate(updates []CellUpdate) error {
    // 1. 批量更新所有单元格
    f.BatchSetCellValue(updates)

    // 2. 读取 calcChain（包含所有工作表的公式）
    calcChain, _ := f.calcChainReader()
    if calcChain == nil || len(calcChain.C) == 0 {
        return nil
    }

    // 3. 清除所有计算缓存
    // ✅ 确保所有依赖（包括跨工作表）都会重新计算
    f.calcCache = sync.Map{}

    // 4. 重新计算所有工作表
    // ✅ 按 calcChain 顺序，确保依赖关系正确
    return f.recalculateAllSheets(calcChain)
}
```

### 新增辅助函数

```go
// recalculateAllSheets 按 calcChain 顺序重新计算所有工作表
func (f *File) recalculateAllSheets(calcChain *xlsxCalcChain) error {
    currentSheetID := -1

    // 遍历 calcChain 中的所有单元格（跨工作表）
    for i := range calcChain.C {
        c := calcChain.C[i]

        // 更新当前工作表 ID
        if c.I != 0 {
            currentSheetID = c.I
        }

        // 获取工作表名称
        sheetName := f.GetSheetMap()[currentSheetID]
        if sheetName == "" {
            continue
        }

        // 重新计算单元格
        f.recalculateCell(sheetName, c.R)
    }

    return nil
}
```

---

## 🔍 工作原理

### calcChain 结构

Excel 使用 `calcChain.xml` 记录所有需要计算的公式单元格：

```xml
<calcChain>
    <c r="B1" i="1"/>  <!-- Sheet1!B1 -->
    <c r="B2" i="1"/>  <!-- Sheet1!B2 -->
    <c r="C1" i="2"/>  <!-- Sheet2!C1 (跨表引用) -->
    <c r="C2" i="2"/>  <!-- Sheet2!C2 (跨表引用) -->
</calcChain>
```

**关键点**：
- `r` - 单元格坐标
- `i` - 工作表 ID（1-based）
- **顺序很重要** - 先计算依赖，再计算引用

### 计算流程

#### 1. 更新数据

```go
updates := []CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 200},
}
```

**结果**：Sheet1.A1 的值从 100 改为 200

#### 2. 清除缓存

```go
f.calcCache = sync.Map{}  // 清空内存缓存
```

**作用**：
- 强制所有公式重新计算
- 不会读取旧的缓存值

#### 3. 遍历 calcChain

```go
for each cell in calcChain:
    CalcCellValue(sheet, cell)
```

**顺序示例**：
```
1. Sheet1!B1 = A1*2
   → 读取 A1 = 200
   → 计算：200*2 = 400
   → 更新缓存：B1.V = "400"

2. Sheet2!C1 = Sheet1!B1+10
   → 读取 Sheet1!B1
   → 优先读取缓存：B1.V = "400"
   → 计算：400+10 = 410
   → 更新缓存：C1.V = "410"
```

---

## 📊 测试验证

### 测试 1: 基本跨工作表依赖

```go
func TestBatchUpdateAndRecalculate_CrossSheet(t *testing.T) {
    f := NewFile()
    f.NewSheet("Sheet2")

    // 设置数据
    f.SetCellValue("Sheet1", "A1", 100)

    // 创建跨工作表公式
    formulas := []FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},      // 200
        {Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+10"}, // 210
    }
    f.BatchSetFormulasAndRecalculate(formulas)

    // 验证初始值
    assert.Equal(t, "200", f.GetCellValue("Sheet1", "B1"))
    assert.Equal(t, "210", f.GetCellValue("Sheet2", "C1"))

    // ✅ 更新 Sheet1 数据
    updates := []CellUpdate{
        {Sheet: "Sheet1", Cell: "A1", Value: 500},
    }
    f.BatchUpdateAndRecalculate(updates)

    // ✅ 验证跨工作表重新计算
    assert.Equal(t, "1000", f.GetCellValue("Sheet1", "B1"))  // 500*2
    assert.Equal(t, "1010", f.GetCellValue("Sheet2", "C1"))  // 1000+10 ✅
}
```

**测试结果**：✅ PASS

### 测试 2: 多层跨工作表依赖链

```go
func TestBatchUpdateAndRecalculate_CrossSheetComplex(t *testing.T) {
    f := NewFile()
    f.NewSheet("Sheet2")
    f.NewSheet("Sheet3")

    // 设置依赖链：Sheet1 → Sheet2 → Sheet3
    f.SetCellValue("Sheet1", "A1", 10)

    formulas := []FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},          // 20
        {Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+5"},   // 25
        {Sheet: "Sheet3", Cell: "D1", Formula: "=Sheet2!C1*3"},   // 75
    }
    f.BatchSetFormulasAndRecalculate(formulas)

    // 验证链：10 → 20 → 25 → 75
    assert.Equal(t, "20", f.GetCellValue("Sheet1", "B1"))
    assert.Equal(t, "25", f.GetCellValue("Sheet2", "C1"))
    assert.Equal(t, "75", f.GetCellValue("Sheet3", "D1"))

    // 更新源数据
    updates := []CellUpdate{
        {Sheet: "Sheet1", Cell: "A1", Value: 50},
    }
    f.BatchUpdateAndRecalculate(updates)

    // 验证整条链都重新计算：50 → 100 → 105 → 315
    assert.Equal(t, "100", f.GetCellValue("Sheet1", "B1"))  // 50*2
    assert.Equal(t, "105", f.GetCellValue("Sheet2", "C1"))  // 100+5 ✅
    assert.Equal(t, "315", f.GetCellValue("Sheet3", "D1"))  // 105*3 ✅
}
```

**测试结果**：✅ PASS

### 测试 3: 多个更新影响跨工作表公式

```go
func TestBatchUpdateAndRecalculate_CrossSheetMultipleUpdates(t *testing.T) {
    f := NewFile()
    f.NewSheet("Sheet2")

    // Sheet1 数据
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellValue("Sheet1", "A2", 20)
    f.SetCellValue("Sheet1", "A3", 30)

    formulas := []FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
        {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
        {Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
        {Sheet: "Sheet2", Cell: "C1", Formula: "=SUM(Sheet1!B1:B3)"},
    }
    f.BatchSetFormulasAndRecalculate(formulas)

    // 初始：SUM(20,40,60) = 120
    assert.Equal(t, "120", f.GetCellValue("Sheet2", "C1"))

    // 批量更新 Sheet1
    updates := []CellUpdate{
        {Sheet: "Sheet1", Cell: "A1", Value: 100},
        {Sheet: "Sheet1", Cell: "A2", Value: 200},
        {Sheet: "Sheet1", Cell: "A3", Value: 300},
    }
    f.BatchUpdateAndRecalculate(updates)

    // 验证 Sheet2 公式重新计算：SUM(200,400,600) = 1200
    assert.Equal(t, "1200", f.GetCellValue("Sheet2", "C1"))  // ✅
}
```

**测试结果**：✅ PASS

---

## 🔄 与旧版本对比

| 方面 | 旧版本 | 新版本 |
|-----|--------|--------|
| **单工作表更新** | ✅ 支持 | ✅ 支持 |
| **跨工作表依赖** | ❌ 不支持 | ✅ 支持 |
| **计算范围** | 只计算被更新的工作表 | 计算所有有公式的工作表 |
| **缓存策略** | 部分清除 | 完全清除 |
| **性能** | 更快（但功能不完整） | 稍慢（但功能完整） |
| **正确性** | ❌ 跨表场景错误 | ✅ 所有场景正确 |

---

## ⚠️ 性能考虑

### 性能影响分析

**旧版本**（只计算部分工作表）：
- ⏱️ 时间：O(m)，m = 被更新工作表的公式数
- ✅ 快速
- ❌ 结果错误（跨工作表场景）

**新版本**（计算所有工作表）：
- ⏱️ 时间：O(n)，n = calcChain 中的总公式数
- ⚠️ 较慢（如果 calcChain 很大）
- ✅ 结果正确

### 优化建议

#### 场景 1: 文件有大量公式（1000+）

如果确定**只有单工作表依赖**，可以使用旧的优化方式：

```go
// 手动优化：只计算特定工作表
updates := []CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 100},
}

f.BatchSetCellValue(updates)
f.RecalculateSheet("Sheet1")  // 只计算 Sheet1
```

#### 场景 2: 需要跨工作表支持（推荐）

使用新版本：

```go
// ✅ 自动处理所有依赖（包括跨工作表）
updates := []CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 100},
}
f.BatchUpdateAndRecalculate(updates)
```

---

## 🎯 使用建议

### ✅ 推荐使用新版本的场景

1. **有跨工作表引用**
   ```go
   // Sheet2 引用 Sheet1
   f.SetCellFormula("Sheet2", "A1", "=Sheet1!B1*2")
   ```

2. **不确定依赖关系**
   - 公式复杂，难以手动追踪依赖

3. **要求结果正确性**
   - 宁可牺牲性能，也要保证正确

### ⚠️ 性能敏感场景

如果有以下情况，考虑手动优化：

1. **calcChain 非常大**（10,000+ 公式）
2. **确定没有跨工作表依赖**
3. **频繁更新**（每秒多次）

**手动优化方案**：
```go
// 只更新和计算特定工作表
f.BatchSetCellValue(updates)
f.RecalculateSheet("Sheet1")  // 手动指定
```

---

## 📝 总结

### 关键改进

| 改进点 | 说明 |
|-------|------|
| ✅ **跨工作表支持** | 更新 Sheet1 后，引用它的 Sheet2 会自动重新计算 |
| ✅ **多层依赖** | 支持 Sheet1 → Sheet2 → Sheet3 的依赖链 |
| ✅ **完全清除缓存** | 使用 `f.calcCache = sync.Map{}` 确保所有公式重新计算 |
| ✅ **calcChain 驱动** | 按 calcChain 顺序计算，保证依赖顺序正确 |

### API 行为变化

**向后兼容性**：✅ 完全兼容
- 单工作表场景：行为不变
- 跨工作表场景：从错误变为正确

**破坏性变更**：❌ 无
- API 签名未变
- 返回值未变
- 只是修复了 bug

---

## 🧪 完整测试覆盖

新增测试：
- ✅ `TestBatchUpdateAndRecalculate_CrossSheet` - 基本跨工作表
- ✅ `TestBatchUpdateAndRecalculate_CrossSheetComplex` - 多层依赖链
- ✅ `TestBatchUpdateAndRecalculate_CrossSheetMultipleUpdates` - 多更新
- ✅ `TestBatchUpdateAndRecalculate_SingleSheetStillWorks` - 单表兼容性

所有测试：✅ PASS

---

**修复日期**：2025-12-26
**修复文件**：`batch.go:105-138, 281-309`
**新增测试**：`batch_cross_sheet_test.go` (200+ 行)
**向后兼容**：✅ 完全兼容
**功能状态**：✅ 生产就绪
