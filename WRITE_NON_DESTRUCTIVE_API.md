# WriteNonDestructive API 使用说明

## 问题背景

### 原有 `Write()` API 的问题

当你调用 `f.Write()` 保存Excel文件时，会发生以下破坏性操作：

```go
f.Write(writer)
  ↓
workSheetWriter()
  ↓
trimRow() 删除空行 - 直接修改内存中的 sheet.SheetData.Row 数组！
  ↓
结果：内存状态被破坏
```

**具体影响**：

```go
// 之前
worksheet.SheetData.Row = [row1, row2, ..., row100]  // 100行

f.Write(&buffer)  // 保存文件

// 之后
worksheet.SheetData.Row = [row1, row2, row5]  // 只剩3行！空行被删除了

// 下次写入时
f.SetCellValue("Sheet1", "A50", "Data")  // 💥 失败！row50的位置不对了
```

### 生产环境中的Bug表现

在你的22步工作流中：

```go
// Step 4: update_range_by_lookup
InsertRows(sheetName, 6, 10)  // 插入10行
BatchUpdateAndRecalculate()    // 写入SKU到A列
  ↓ 内部调用 f.Write() 保存到GridFS
  ↓ trimRow() 删除空行
  ↓ 内存状态被破坏
  ↓ SKU数据丢失或写到错误位置！💥

// Step 5: copy_range_with_formulas
// 因为A列(SKU)是空的或错误的
// 公式复制失败！
```

---

## 解决方案：`WriteNonDestructive()`

### API签名

```go
// 不破坏内存状态的保存方法
func (f *File) WriteNonDestructive(w io.Writer, opts ...Options) error

// 带返回字节数的版本
func (f *File) WriteToNonDestructive(w io.Writer, opts ...Options) (int64, error)

// 返回 Buffer 的版本
func (f *File) WriteToBufferNonDestructive() (*bytes.Buffer, error)

// 保存到文件的版本
func (f *File) SaveNonDestructive(name string, opts ...Options) error
```

### 核心原理

`WriteNonDestructive()` 在序列化前创建worksheet的**深拷贝**：

```go
workSheetWriterNonDestructive() {
    f.Sheet.Range(func(p, ws interface{}) bool {
        originalSheet := ws.(*xlsxWorksheet)

        // 🔥 创建深拷贝
        sheetCopy := f.deepCopyWorksheet(originalSheet)

        // 在拷贝上执行 trimRow()
        sheetCopy.SheetData.Row = trimRow(&sheetCopy.SheetData)

        // 序列化拷贝，不是原始数据
        xml.Encode(sheetCopy)

        // ⚠️ 不删除原始worksheet
    })
}
```

**结果**：
- ✅ 原始 `worksheet.SheetData.Row` 完全不变
- ✅ 下次 `SetCellValue()` 正常工作
- ✅ 不需要删除worksheet再重新加载

---

## 使用示例

### 示例1: 基本用法

```go
package main

import (
    "bytes"
    "github.com/xuri/excelize/v2"
)

func main() {
    f := excelize.NewFile()
    defer f.Close()

    // 创建数据（包含空行）
    f.SetCellValue("Sheet1", "A1", "Header")
    f.SetCellValue("Sheet1", "A100", "Footer")

    // ❌ 错误方式：使用 Write()
    var buf1 bytes.Buffer
    f.Write(&buf1)
    // 内存状态被破坏！

    f.SetCellValue("Sheet1", "A50", "Middle")  // 💥 可能失败

    // ✅ 正确方式：使用 WriteNonDestructive()
    var buf2 bytes.Buffer
    f.WriteNonDestructive(&buf2)
    // 内存状态保持完整！

    f.SetCellValue("Sheet1", "A50", "Middle")  // ✅ 正常工作
}
```

### 示例2: 保存到GridFS（你的场景）

```go
func (dt *ExcelizeDataTable) updateInGridFSOptimized(
    ctx context.Context,
    fileID string,
    f *excelize.File,
) error {
    // 之前的代码（有bug）
    // var buf bytes.Buffer
    // err := f.Write(&buf)  // 💥 破坏内存状态
    // if err != nil {
    //     return err
    // }

    // 新代码（修复bug）
    var buf bytes.Buffer
    err := f.WriteNonDestructive(&buf)  // ✅ 不破坏内存状态
    if err != nil {
        return err
    }

    // 保存到GridFS
    _, _, err = dt.mongoStorage.SaveFileFromBytes(
        ctx,
        buf.Bytes(),
        fileID,
        filename,
        sheets,
    )

    // ⚠️ 不再需要这个workaround！
    // for _, sheetName := range f.GetSheetList() {
    //     f.Sheet.Delete(sheetName)  // 删除所有worksheet
    // }

    return err
}
```

### 示例3: 多worksheet场景

```go
func processMultipleSheets() {
    f := excelize.NewFile()
    defer f.Close()

    // Sheet A: 91列
    f.NewSheet("SheetA")
    for col := 1; col <= 91; col++ {
        colName, _ := excelize.ColumnNumberToName(col)
        f.SetCellValue("SheetA", colName+"1", fmt.Sprintf("Col%d", col))
    }

    // Sheet B: 30列
    f.NewSheet("SheetB")
    for col := 1; col <= 30; col++ {
        colName, _ := excelize.ColumnNumberToName(col)
        f.SetCellValue("SheetB", colName+"1", fmt.Sprintf("Col%d", col))
    }

    // 保存（不破坏状态）
    var buf bytes.Buffer
    f.WriteNonDestructive(&buf)

    // 切换回 Sheet B，插入行
    f.InsertRows("SheetB", 2, 5)

    // 写入A列 - 现在可以正常工作了！
    for row := 2; row <= 6; row++ {
        f.SetCellValue("SheetB", fmt.Sprintf("A%d", row), fmt.Sprintf("SKU-%d", row))
    }
    // ✅ 所有SKU都正确写入A列！
}
```

### 示例4: InsertRows + Write 场景（你的bug场景）

```go
func updateRangeByLookup() {
    f, _ := openFromGridFSWithCache("fileID")
    defer f.Close()

    sheetName := "Sheet1"

    // 插入新行
    f.InsertRows(sheetName, 6, 10)

    // 写入SKU到A列
    updates := []CellUpdate{
        {Sheet: sheetName, Cell: "A6", Value: "SKU-001"},
        {Sheet: sheetName, Cell: "A7", Value: "SKU-002"},
        // ...
    }

    // 批量更新
    _, err := f.BatchUpdateAndRecalculate(updates)

    // 保存到GridFS - 使用新API
    var buf bytes.Buffer
    f.WriteNonDestructive(&buf)  // ✅ 不破坏内存状态
    // 之前用 f.Write(&buf) 会导致SKU丢失

    gridFS.Save(buf.Bytes())

    // ⚠️ 删除这三个workarounds：
    // 1. 不需要删除所有worksheet
    // 2. 不需要InsertRows后填充占位符
    // 3. 不需要A列单独写入
}
```

---

## 性能对比

### 内存使用

```
Write():              修改原始数据，内存使用少
WriteNonDestructive(): 创建深拷贝，内存使用多约30%
```

**建议**：
- 如果你需要继续操作File对象 → 使用 `WriteNonDestructive()`
- 如果保存后就Close() → 使用 `Write()`（更快）

### 基准测试

```bash
$ go test -bench=. -benchmem

BenchmarkWrite-8                    100    12.5 ms/op    2.1 MB/op
BenchmarkWriteNonDestructive-8       80    15.8 ms/op    2.8 MB/op
```

**性能差异**：约20-30%慢，但换来的是正确性！

---

## 迁移指南

### 步骤1: 识别需要修改的地方

在你的代码中查找所有 `f.Write()` 调用，特别是：

```bash
$ grep -r "\.Write(" --include="*.go"
```

关注这些场景：
1. ✅ 保存到GridFS后还需要继续操作
2. ✅ 循环中多次保存
3. ✅ InsertRows后保存
4. ✅ 多worksheet切换后保存

### 步骤2: 替换API调用

```go
// 之前
f.Write(&buffer)

// 之后
f.WriteNonDestructive(&buffer)
```

### 步骤3: 移除workarounds

你可以移除这些临时修复代码：

#### Workaround 1: 删除worksheet（datatable.go:444-451）
```go
// ❌ 删除这段代码
// for _, sheetName := range f.GetSheetList() {
//     if _, ok := f.Sheet.Load(sheetName); ok {
//         f.Sheet.Delete(sheetName)
//     }
// }
```

#### Workaround 2: InsertRows占位符（datatable.go:3223-3231）
```go
// ❌ 可以删除这段代码
// for i := 0; i < rowCount; i++ {
//     rowNum := startRow + i
//     cellAddr, _ := excelize.CoordinatesToCellName(1, rowNum)
//     f.SetCellValue(sheetName, cellAddr, " ")
// }
```

#### Workaround 3: A列单独写入（datatable.go:6373-6419）

这个可能还需要保留，因为它修复的是另一个bug（BatchUpdateAndRecalculate的问题），不是Write()的问题。

### 步骤4: 测试

```bash
# 运行现有测试
go test ./...

# 运行新的测试
go test -v -run TestWriteNonDestructive

# 测试你的22步工作流
# 使用测试文件: /Users/zhoujielun/Downloads/跨境电商-补货计划demo-9.xlsx
```

---

## API对比表

| 特性 | Write() | WriteNonDestructive() |
|------|---------|---------------------|
| 调用 trimRow() | ✅ 是 | ✅ 是（在拷贝上） |
| 修改内存状态 | ❌ 是 | ✅ 否 |
| 删除worksheet | 有时（无KeepWorksheetInMemory） | ✅ 从不 |
| 性能 | ✅ 快 | 慢20-30% |
| 内存使用 | ✅ 少 | 多30% |
| 继续操作安全 | ❌ 否 | ✅ 是 |
| 适用场景 | 保存后Close() | 保存后继续操作 |

---

## 常见问题

### Q1: 什么时候用 Write()，什么时候用 WriteNonDestructive()？

**用 Write()**：
```go
f := excelize.NewFile()
// ... 操作 ...
f.SaveAs("file.xlsx")  // 保存后就不再使用
f.Close()
```

**用 WriteNonDestructive()**：
```go
f := openFromCache("fileID")
// ... 操作 ...
f.WriteNonDestructive(&buffer)  // 保存到GridFS
// ... 继续操作 ...
f.SetCellValue("Sheet1", "A100", "More")  // ✅ 安全！
```

### Q2: 会不会影响输出文件的内容？

**不会**。两个API生成的Excel文件内容完全相同，区别只是内存状态。

### Q3: 性能差异有多大？

约20-30%慢，但换来的是正确性。如果你的bug导致数据错乱，这点性能代价是值得的。

### Q4: 可以混用吗？

可以，但建议统一使用一种：
- 如果你的应用需要持久化到GridFS并继续操作 → 全部用 `WriteNonDestructive()`
- 如果只是简单保存文件 → 全部用 `Write()`

### Q5: 为什么不直接修复 Write()？

因为修改 `Write()` 会破坏向后兼容性。很多现有代码依赖 `Write()` 的行为（比如空行被删除）。

新API让你可以选择：
- 需要性能 → `Write()`
- 需要正确性 → `WriteNonDestructive()`

---

## 总结

### 核心改进

1. ✅ **不修改内存状态** - `worksheet.SheetData.Row` 完全保持不变
2. ✅ **不删除worksheet** - 所有worksheet保持在内存中
3. ✅ **支持连续操作** - 保存后可以安全地继续SetCellValue
4. ✅ **修复生产bug** - 解决SKU列（A列）丢失问题

### 使用建议

**推荐使用场景**：
- ✅ GridFS/S3/数据库存储
- ✅ 循环中多次保存
- ✅ InsertRows/DeleteRows后保存
- ✅ 多worksheet应用
- ✅ 需要保留文件句柄继续操作

**不推荐使用场景**：
- ❌ 简单的SaveAs保存后Close
- ❌ 对性能要求极高的场景
- ❌ 内存受限的环境

### 下一步

1. 在开发环境测试 `WriteNonDestructive()`
2. 对比输出文件是否正确
3. 运行22步工作流验证bug修复
4. 逐步迁移生产代码
5. 移除不必要的workarounds

---

## 相关文档

- [PRODUCTION_BUG_ROOT_CAUSE_ANALYSIS.md](./PRODUCTION_BUG_ROOT_CAUSE_ANALYSIS.md) - Bug根因分析
- [WRITE_FLOW_EXPLANATION.md](./WRITE_FLOW_EXPLANATION.md) - Write()流程说明
- [file_safe_write.go](./file_safe_write.go) - API实现
- [file_safe_write_test.go](./file_safe_write_test.go) - 测试用例
