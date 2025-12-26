# BatchSetFormulasAndRecalculate 计算机制详解

## 📋 目录

1. [执行流程](#执行流程)
2. [计算机制](#计算机制)
3. [缓存策略](#缓存策略)
4. [依赖处理](#依赖处理)
5. [完整示例](#完整示例)
6. [性能考虑](#性能考虑)

---

## 执行流程

`BatchSetFormulasAndRecalculate` 的执行分为 **4 个步骤**：

```go
func (f *File) BatchSetFormulasAndRecalculate(formulas []FormulaUpdate) error {
    // 步骤 1: 批量设置公式（写入 XML）
    if err := f.BatchSetFormulas(formulas); err != nil {
        return err
    }

    // 步骤 2: 收集受影响的工作表
    affectedSheets := make(map[string][]string)
    for _, formula := range formulas {
        affectedSheets[formula.Sheet] = append(affectedSheets[formula.Sheet], formula.Cell)
    }

    // 步骤 3: 更新 calcChain（建立依赖关系）
    if err := f.updateCalcChainForFormulas(formulas); err != nil {
        return err
    }

    // 步骤 4: 重新计算每个受影响的工作表
    for sheet := range affectedSheets {
        if err := f.RecalculateSheet(sheet); err != nil {
            return err
        }
    }

    return nil
}
```

---

## 计算机制

### 关键问题解答

**Q: 新设置的公式本身会被计算吗？**
✅ **会**。所有在 `formulas` 列表中的公式都会被计算。

**Q: 是用缓存值给引用的单元格计算吗？**
✅ **是的**。计算后的值会存储在单元格的 `<v>` 标签（缓存值），其他公式引用时直接读取。

### 详细流程

#### 步骤 1: 批量设置公式（不计算）

```go
// BatchSetFormulas 只是写入公式到 XML 结构
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+10"},
}
f.BatchSetFormulas(formulas)
```

**结果**：XML 中写入了公式，但 `<v>` 标签（缓存值）为空：

```xml
<c r="B1">
    <f>A1*2</f>     <!-- 公式已设置 -->
    <v></v>         <!-- 缓存值为空 ❌ -->
</c>
<c r="C1">
    <f>B1+10</f>    <!-- 公式已设置 -->
    <v></v>         <!-- 缓存值为空 ❌ -->
</c>
```

#### 步骤 2: 更新 calcChain

```go
f.updateCalcChainForFormulas(formulas)
```

**作用**：将新公式添加到 calcChain.xml，建立计算顺序：

```xml
<calcChain>
    <c r="B1" i="1"/>  <!-- B1 在 Sheet1（ID=1） -->
    <c r="C1" i="1"/>  <!-- C1 在 Sheet1（ID=1） -->
</calcChain>
```

#### 步骤 3: 重新计算工作表

```go
f.RecalculateSheet("Sheet1")
```

**核心逻辑**：

```go
func (f *File) RecalculateSheet(sheet string) error {
    // 1. 读取 calcChain
    calcChain, _ := f.calcChainReader()

    // 2. 遍历 calcChain 中的所有单元格
    for i := range calcChain.C {
        c := calcChain.C[i]

        // 3. 对每个单元格调用 recalculateCell
        f.recalculateCell(sheetName, c.R)
    }
}
```

#### 步骤 4: 单元格计算

```go
func (f *File) recalculateCell(sheet, cell string) error {
    // 1. 检查单元格是否有公式
    cellRef := findCell(ws, cell)
    if cellRef.F == nil {
        return nil  // 没有公式，跳过
    }

    // 2. 使用 CalcCellValue 计算公式值
    result, err := f.CalcCellValue(sheet, cell)

    // 3. 更新单元格的缓存值（<v> 标签）
    cellRef.V = result
    cellRef.T = "n"  // 数字类型
}
```

### CalcCellValue 的计算机制

**核心实现** (calc.go:854-902):

```go
func (f *File) CalcCellValue(sheet, cell string, opts ...Options) (string, error) {
    // 1. 检查计算缓存
    cacheKey := fmt.Sprintf("%s!%s!raw=%t", sheet, cell, rawCellValue)
    if cachedResult, found := f.calcCache.Load(cacheKey); found {
        return cachedResult.(string), nil  // 🚀 缓存命中
    }

    // 2. 解析公式并计算
    token, err := f.calcCellValue(&calcContext{...}, sheet, cell)

    // 3. 格式化结果
    result := formatValue(token)

    // 4. 存入缓存
    f.calcCache.Store(cacheKey, result)

    return result, nil
}
```

**关键点**：
- ✅ **运行时计算**：每次调用 `CalcCellValue` 都会**实时解析公式**并计算
- ✅ **读取依赖值**：如果公式引用其他单元格（如 `=B1+10`），会递归调用 `GetCellValue` 读取 B1 的值
- ✅ **缓存结果**：计算结果存入 `f.calcCache`（内存缓存）和 `cellRef.V`（XML 缓存）

---

## 缓存策略

Excelize 使用**双层缓存**：

### 1. 内存缓存 (f.calcCache)

```go
// sync.Map 存储计算结果
f.calcCache.Store("Sheet1!B1!raw=false", "20")
```

**作用**：
- 避免重复计算同一个公式
- 加速后续读取

**生命周期**：
- ✅ 进程内有效
- ❌ 不持久化（File 关闭后清空）

### 2. XML 缓存 (cellRef.V)

```xml
<c r="B1">
    <f>A1*2</f>      <!-- 公式 -->
    <v>20</v>        <!-- 缓存值 -->
    <t>n</t>         <!-- 类型：数字 -->
</c>
```

**作用**：
- Excel 打开文件时直接显示缓存值
- 不需要立即重新计算

**生命周期**：
- ✅ 保存到文件
- ✅ 持久化存储

---

## 依赖处理

### 场景 1: 简单依赖

```go
// A1 = 10 (已存在)
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
}
f.BatchSetFormulasAndRecalculate(formulas)
```

**执行过程**：

1. **设置公式**：B1.F = "A1*2"
2. **计算 B1**：
   ```go
   CalcCellValue("Sheet1", "B1")
   → 解析公式 "A1*2"
   → 读取 A1 的值：GetCellValue("Sheet1", "A1") = "10"
   → 计算：10 * 2 = 20
   → 存储：B1.V = "20"
   ```

### 场景 2: 链式依赖

```go
// A1 = 10 (已存在)
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},   // B1 依赖 A1
    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+10"},  // C1 依赖 B1
}
f.BatchSetFormulasAndRecalculate(formulas)
```

**执行过程**：

1. **设置公式**：
   - B1.F = "A1*2"
   - C1.F = "B1+10"

2. **calcChain 顺序**（根据 XML 中的顺序）：
   ```xml
   <calcChain>
       <c r="B1" i="1"/>
       <c r="C1" i="1"/>
   </calcChain>
   ```

3. **计算 B1**：
   ```go
   CalcCellValue("Sheet1", "B1")
   → 读取 A1 = "10"
   → 计算：10 * 2 = 20
   → 存储：B1.V = "20"  ✅
   ```

4. **计算 C1**：
   ```go
   CalcCellValue("Sheet1", "C1")
   → 读取 B1 的值
   → GetCellValue("Sheet1", "B1")
      → 优先返回缓存值 B1.V = "20"  🚀
   → 计算：20 + 10 = 30
   → 存储：C1.V = "30"  ✅
   ```

**关键点**：
- ✅ C1 计算时，B1 已经有缓存值（B1.V = "20"）
- ✅ `GetCellValue` 优先返回缓存值，不会重新计算 B1
- ⚠️ **顺序很重要**：如果 calcChain 中 C1 在 B1 前面，C1 会读取到**空值**

### 场景 3: 循环依赖

```go
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=C1+1"},  // B1 依赖 C1
    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+1"},  // C1 依赖 B1 ❌
}
f.BatchSetFormulasAndRecalculate(formulas)
```

**执行过程**：

1. **计算 B1**：
   ```go
   CalcCellValue("Sheet1", "B1")
   → 解析 "C1+1"
   → 读取 C1：GetCellValue("Sheet1", "C1")
      → C1 有公式 "=B1+1"
      → 递归计算 C1：CalcCellValue("Sheet1", "C1")
         → 解析 "B1+1"
         → 读取 B1：GetCellValue("Sheet1", "B1")
            → 检测到循环引用！ ⚠️
   → 返回错误或默认值
   ```

**处理**：
- Excelize 使用 `maxCalcIterations` 限制递归深度
- 达到限制后返回错误或清空缓存

---

## 完整示例

### 示例 1: 基本计算流程

```go
func Example_BasicCalculation() {
    f := excelize.NewFile()
    defer f.Close()

    // 设置原始数据
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellValue("Sheet1", "A2", 20)

    // 批量设置公式并计算
    formulas := []excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},      // B1 = 10*2 = 20
        {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},      // B2 = 20*2 = 40
        {Sheet: "Sheet1", Cell: "C1", Formula: "=SUM(B1:B2)"}, // C1 = 20+40 = 60
    }

    err := f.BatchSetFormulasAndRecalculate(formulas)
    if err != nil {
        panic(err)
    }

    // 验证计算结果
    b1, _ := f.GetCellValue("Sheet1", "B1")
    fmt.Println("B1 =", b1)  // 输出: B1 = 20

    b2, _ := f.GetCellValue("Sheet1", "B2")
    fmt.Println("B2 =", b2)  // 输出: B2 = 40

    c1, _ := f.GetCellValue("Sheet1", "C1")
    fmt.Println("C1 =", c1)  // 输出: C1 = 60

    // 查看内部 XML 结构
    ws, _ := f.workSheetReader("Sheet1")
    for _, row := range ws.SheetData.Row {
        for _, cell := range row.C {
            if cell.F != nil {
                fmt.Printf("Cell %s: Formula=%s, Cache=%s\n",
                    cell.R, cell.F.Content, cell.V)
            }
        }
    }
    // 输出:
    // Cell B1: Formula=A1*2, Cache=20
    // Cell B2: Formula=A2*2, Cache=40
    // Cell C1: Formula=SUM(B1:B2), Cache=60
}
```

### 示例 2: 依赖链验证

```go
func Example_DependencyChain() {
    f := excelize.NewFile()
    defer f.Close()

    // A1 = 100
    f.SetCellValue("Sheet1", "A1", 100)

    // 创建依赖链：A1 → B1 → C1 → D1
    formulas := []excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},    // 100*2 = 200
        {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+50"},   // 200+50 = 250
        {Sheet: "Sheet1", Cell: "D1", Formula: "=C1/10"},   // 250/10 = 25
    }

    // 一次性计算所有
    f.BatchSetFormulasAndRecalculate(formulas)

    // 验证每一层
    b1, _ := f.GetCellValue("Sheet1", "B1")
    c1, _ := f.GetCellValue("Sheet1", "C1")
    d1, _ := f.GetCellValue("Sheet1", "D1")

    fmt.Printf("A1=100 → B1=%s → C1=%s → D1=%s\n", b1, c1, d1)
    // 输出: A1=100 → B1=200 → C1=250 → D1=25
}
```

### 示例 3: 跨工作表引用

```go
func Example_CrossSheetReference() {
    f := excelize.NewFile()
    defer f.Close()

    f.NewSheet("Sheet2")

    // Sheet1: A1 = 100
    f.SetCellValue("Sheet1", "A1", 100)

    // Sheet2 引用 Sheet1
    formulas := []excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},          // 200
        {Sheet: "Sheet2", Cell: "A1", Formula: "=Sheet1!B1+50"},  // 250
    }

    f.BatchSetFormulasAndRecalculate(formulas)

    // 验证
    sheet1B1, _ := f.GetCellValue("Sheet1", "B1")
    sheet2A1, _ := f.GetCellValue("Sheet2", "A1")

    fmt.Printf("Sheet1!B1 = %s\n", sheet1B1)  // 200
    fmt.Printf("Sheet2!A1 = %s\n", sheet2A1)  // 250
}
```

---

## 性能考虑

### 计算开销分析

| 操作 | 时间复杂度 | 说明 |
|-----|-----------|------|
| 设置公式 | O(n) | n = 公式数量 |
| 更新 calcChain | O(n) | 遍历并添加 |
| 计算公式 | O(m × k) | m = calcChain 中的公式数，k = 平均依赖深度 |

### 缓存命中率

**高命中场景**（性能好）：
```go
// B1 被多次引用，但只计算一次
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+10"},
    {Sheet: "Sheet1", Cell: "C2", Formula: "=B1+20"},
    {Sheet: "Sheet1", Cell: "C3", Formula: "=B1+30"},
}
// B1 计算 1 次，C1/C2/C3 直接读取 B1 缓存
```

**低命中场景**（性能差）：
```go
// 每个公式都依赖不同的源单元格
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
    {Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
    {Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
    // ... 1000 个独立公式
}
// 需要计算 1000 次，无法利用缓存
```

### 优化建议

#### ✅ 推荐：批量设置，一次计算

```go
// ✅ 好：收集所有公式，一次性处理
formulas := make([]excelize.FormulaUpdate, 100)
for i := 0; i < 100; i++ {
    formulas[i] = excelize.FormulaUpdate{
        Sheet:   "Sheet1",
        Cell:    fmt.Sprintf("B%d", i+1),
        Formula: fmt.Sprintf("=A%d*2", i+1),
    }
}
f.BatchSetFormulasAndRecalculate(formulas)
```

#### ❌ 避免：循环调用

```go
// ❌ 差：每次都重新遍历 calcChain
for i := 0; i < 100; i++ {
    f.BatchSetFormulasAndRecalculate([]excelize.FormulaUpdate{
        {Sheet: "Sheet1", Cell: fmt.Sprintf("B%d", i+1), Formula: "=A1*2"},
    })
}
```

---

## 总结

### 计算机制要点

| 方面 | 说明 |
|-----|------|
| **公式本身** | ✅ 会被计算 |
| **计算方式** | 🔄 实时解析公式，递归计算依赖 |
| **缓存使用** | ✅ 优先读取缓存值（XML + 内存） |
| **计算顺序** | 📋 根据 calcChain 顺序计算 |
| **依赖处理** | 🔗 自动递归计算依赖项 |
| **循环检测** | ⚠️ 通过迭代限制防止死循环 |

### API 调用链

```
BatchSetFormulasAndRecalculate
    ├─ BatchSetFormulas (设置公式到 XML)
    ├─ updateCalcChainForFormulas (更新 calcChain)
    └─ RecalculateSheet (重新计算)
        └─ recalculateAllInSheet
            └─ recalculateCell (for each cell in calcChain)
                └─ CalcCellValue (计算公式)
                    ├─ 检查缓存
                    ├─ 解析公式
                    ├─ 递归获取依赖值
                    ├─ 计算结果
                    └─ 更新缓存
```

---

**生成时间**: 2025-12-26
**相关文档**: [BATCH_SET_FORMULAS_API.md](./BATCH_SET_FORMULAS_API.md)
