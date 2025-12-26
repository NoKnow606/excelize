# CheckRow Index Out of Range 修复

## 问题描述

在生产环境中遇到 `index out of range [23] with length 23` panic，发生在 `rows.go:1091`：

```
panic: runtime error: index out of range [23] with length 23
github.com/xuri/excelize/v2.(*xlsxWorksheet).checkRow(0xc001662908)
    rows.go:1091 +0x63c
```

### 堆栈信息

```
调用链：
workSheetReader → checkRow → panic at line 1091
触发场景：GetCellFormula → getCellStringFunc → workSheetReader → checkRow
问题文件：纪要_1766734050_copy_1766743005.xlsx
```

## Root Cause 分析

### 原始代码（有 bug）

```go
if colCount < lastCol {
    sourceList := rowData.C
    targetList := make([]xlsxC, 0, lastCol)  // ⚠️ 只用 lastCol

    // 创建 targetList
    for colIdx := 0; colIdx < lastCol; colIdx++ {
        targetList = append(targetList, xlsxC{R: cellName})
    }

    rowData.C = targetList

    // 🐛 BUG: 如果 sourceList 中有单元格的列号 > lastCol，这里会 panic
    for colIdx := range sourceList {
        colData := &sourceList[colIdx]
        colNum, _, _ := CellNameToCoordinates(colData.R)
        ws.SheetData.Row[rowIdx].C[colNum-1] = *colData  // panic here!
    }
}
```

### 触发条件

当 `sourceList` 中某个单元格的列号 **大于** `lastCol` 时：

```
示例场景：
- sourceList 有 3 个单元格：A1, B1, X1
- lastCol 从最后一个单元格 (X1) 计算得到 24
- targetList 创建时长度为 24（索引 0-23）
- 但如果 sourceList 中还有 Y1（第25列）：
  - colNum = 25
  - colNum - 1 = 24
  - 访问 targetList[24] → panic!
```

**为什么会出现这种情况？**

1. **trimRow() 修改后数据不一致**
   - `Write()` 调用 `trimRow()` 删除空行
   - 可能导致单元格引用不正确

2. **XML 数据损坏**
   - 文件本身包含无效的单元格引用
   - 列号超出预期范围

3. **并发修改**
   - Write() 修改 worksheet 时，另一个 goroutine 正在读取

## 修复方案

### 方案1：重新计算 maxCol（已实施）

```go
if colCount < lastCol {
    sourceList := rowData.C

    // ✅ FIX: 遍历 sourceList，找到真正的最大列号
    maxCol := lastCol
    for _, cell := range sourceList {
        colNum, _, err := CellNameToCoordinates(cell.R)
        if err != nil {
            continue
        }
        if colNum > maxCol {
            maxCol = colNum
        }
    }

    targetList := make([]xlsxC, 0, maxCol)  // ✅ 使用 maxCol

    // 创建 targetList
    for colIdx := 0; colIdx < maxCol; colIdx++ {
        targetList = append(targetList, xlsxC{R: cellName})
    }

    rowData.C = targetList

    // ✅ FIX: 添加边界检查
    for colIdx := range sourceList {
        colData := &sourceList[colIdx]
        colNum, _, err := CellNameToCoordinates(colData.R)
        if err != nil {
            return err
        }

        // 边界检查，防止 panic
        if colNum-1 < 0 || colNum-1 >= len(ws.SheetData.Row[rowIdx].C) {
            continue  // 跳过无效的单元格引用
        }

        ws.SheetData.Row[rowIdx].C[colNum-1] = *colData
    }
}
```

### 关键改进

1. **重新计算 maxCol**：遍历所有 sourceList 中的单元格，找到真正的最大列号
2. **边界检查**：在访问数组前检查索引是否有效
3. **优雅降级**：遇到无效引用时跳过，而不是 panic

## 测试验证

创建了 6 个测试用例，全部通过 ✅：

1. **TestCheckRowIndexOutOfRange** - 稀疏列数据（A1, X1）
2. **TestCheckRowWithCorruptedData** - 损坏数据（A1, Z1, AA2）
3. **TestCheckRowAfterInsertRows** - InsertRows 后调用 checkRow
4. **TestCheckRowWithWriteAndReload** - Write 后调用 checkRow
5. **TestCheckRowMultipleTimes** - 多次调用 checkRow
6. **TestCheckRow** - 基础功能测试

```bash
go test -run "TestCheckRow" -v
# PASS: 所有 6 个测试通过
```

## 影响范围

### 修改的文件

- `rows.go:1069-1112` - checkRow() 函数

### 向后兼容性

✅ 完全向后兼容：
- 只添加了安全检查，不改变正常流程
- 无效单元格引用会被跳过（原本会 panic）
- 所有现有测试通过

### 性能影响

轻微性能影响（可忽略）：
- 额外遍历一次 sourceList 找 maxCol（通常 sourceList 很小）
- 每次赋值前多一次边界检查（一个 if 判断）

## 部署建议

### 对于你的应用

如果你无法立即更新 Excelize 库，可以在应用层添加 panic 恢复：

```go
func (dt *ExcelizeDataTable) checkCacheHealthQuick(ctx context.Context, f *excelize.File, sheetName string) error {
    defer func() {
        if r := recover(); r != nil {
            log.Printf("⚠️  checkRow panic for sheet=%s: %v", sheetName, r)
            // 可选：保存问题文件用于分析
            // f.SaveAs(fmt.Sprintf("/tmp/corrupted_%s.xlsx", docID))
        }
    }()

    // 你的现有逻辑
    // ...
}
```

### 长期方案

1. **更新 Excelize 库**到包含此修复的版本
2. **添加文件验证**：在上传时检查 Excel 文件的有效性
3. **监控日志**：记录所有触发边界检查的文件，分析根本原因

## 相关问题

这个 bug 可能与以下操作有关：

1. **InsertRows → Write → GetValue** 流程（之前讨论的场景）
2. **trimRow() 修改内部状态**（Write 会调用）
3. **并发访问 worksheet**（多个 goroutine）
4. **损坏的 Excel 文件**（无效的单元格引用）

## 总结

✅ **修复完成**：添加了 maxCol 重新计算和边界检查
✅ **测试通过**：6 个测试用例全部通过
✅ **向后兼容**：不影响现有功能
✅ **防御性编程**：优雅处理无效数据，不会 panic

这个修复应该能解决你遇到的 `index out of range [23] with length 23` 问题。
