# calcCache 缓存清除机制详解

## 缓存清除策略

### 1. 全局清除 (calcCache.Clear())

**触发场景：影响多个单元格或整体结构的操作**

| 操作 | 文件 | 原因 |
|------|------|------|
| `SetCellValue` (删除公式时) | cell.go:187 | 删除公式可能影响所有依赖该单元格的公式 |
| `InsertRows/DeleteRows` | adjust.go:78 | 行列调整会影响所有公式引用 |
| `InsertCols/DeleteCols` | adjust.go:78 | 同上 |
| `MergeCell` | merge.go:118 | 合并单元格改变区域结构 |
| `UnmergeCell` | merge.go:118 | 取消合并改变区域结构 |
| `SetSheetName` | sheet.go:387 | 改名影响跨表引用 |
| `DeleteSheet` | sheet.go:584 | 删除工作表影响所有引用 |
| `CopySheet` | sheet.go:774 | 复制工作表创建新引用 |
| `AddTable` | table.go:121 | 添加表格可能创建新的命名范围 |
| `DeleteTable` | table.go:182 | 删除表格影响表格引用 |
| `AddPivotTable` | pivotTable.go:166 | 数据透视表操作 |
| `DeletePivotTable` | pivotTable.go:1066 | 数据透视表操作 |
| `SetDefinedName` | sheet.go:1779 | 命名范围影响公式 |
| `DeleteDefinedName` | sheet.go:1823 | 命名范围影响公式 |

**特点：**
- ✅ **简单直接** - 一次性清除所有缓存
- ✅ **安全可靠** - 不会遗漏任何依赖
- ❌ **性能开销大** - 后续所有计算都需重新执行
- ❌ **影响范围广** - 即使不相关的单元格也被清除

### 2. 精细清除 (clearCellCache)

**触发场景：单个单元格变化**

```go
// calc.go:2240
func (f *File) clearCellCache(sheet, cell string) {
    ref := fmt.Sprintf("%s!%s", sheet, cell)

    // 1. 清除单元格自身的缓存
    f.calcCache.Delete(ref)

    // 2. 清除包含该单元格的范围缓存
    f.rangeCache.Range(func(key, value interface{}) bool {
        cacheKey := key.(string)
        if strings.HasPrefix(cacheKey, sheet+"!") {
            f.rangeCache.Delete(key)
        }
        return true
    })
}
```

| 操作 | 文件 | 原因 |
|------|------|------|
| `SetCellFormula` | cell.go:802 | 修改单个单元格公式 |
| `SetCellRichText` | cell.go:1379 | 修改单个单元格富文本 |

**特点：**
- ✅ **性能优化** - 只清除相关缓存
- ✅ **影响最小** - 不影响其他单元格
- ⚠️ **实现复杂** - 需要准确判断依赖关系
- ⚠️ **潜在风险** - 如果依赖判断不准确，可能遗漏缓存清除

## 缓存清除对比

| 维度 | 全局清除 | 精细清除 |
|------|---------|---------|
| **清除范围** | 整个 workbook | 单个 cell + 相关 range |
| **性能影响** | 大（13.77x 性能差距） | 小 |
| **实现难度** | 简单 | 复杂 |
| **安全性** | 高（不会遗漏） | 中（需要准确依赖分析） |
| **使用场景** | 结构性变化 | 单元格内容变化 |

## 当前存在的问题

### 1. 过度清除问题

**问题：** `SetCellValue` 时会清除整个缓存

```go
// cell.go:187
func (f *File) removeFormula(c *xlsxC, ws *xlsxWorksheet, sheet string) error {
    // When removing formula due to SetCellValue, clear entire calcCache
    f.calcCache.Clear()  // ⚠️ 过度清除
    f.rangeCache.Clear()
}
```

**影响：**
- 修改一个单元格值 → 清除 400 万单元格的缓存
- 批量导入数据时，每次 `SetCellValue` 都清空缓存
- **性能损失巨大**（13.77x）

### 2. 范围缓存清除不精确

```go
// calc.go:2253
f.rangeCache.Range(func(key, value interface{}) bool {
    if strings.HasPrefix(cacheKey, sheet+"!") {
        f.rangeCache.Delete(key)  // ⚠️ 清除该 sheet 的所有 range 缓存
    }
    return true
})
```

**问题：** 修改 A1 单元格时，B1:Z100 的范围缓存也被清除

## 优化建议

### 优化 1: 改进 SetCellValue 的缓存策略 ⭐⭐⭐⭐⭐

**当前问题：**
```go
// 每次 SetCellValue 都清空整个缓存
for i := 0; i < 40000; i++ {
    f.SetCellValue(sheet, "A"+strconv.Itoa(i), data[i])
    // 每次调用都执行 f.calcCache.Clear()
}
```

**建议方案 1: 延迟清除**
```go
type File struct {
    calcCache       sync.Map
    cacheDirty      bool  // 标记缓存是否需要清除
}

func (f *File) SetCellValue(sheet, cell string, value interface{}) error {
    // 只标记为脏，不立即清除
    f.cacheDirty = true
    // ... 设置值的逻辑
}

func (f *File) CalcCellValue(sheet, cell string) (string, error) {
    // 在计算时才清除缓存
    if f.cacheDirty {
        f.calcCache.Clear()
        f.cacheDirty = false
    }
    // ... 计算逻辑
}
```

**建议方案 2: 依赖追踪**
```go
type File struct {
    calcCache       sync.Map
    cellDependents  sync.Map  // 记录每个单元格被哪些公式依赖
}

func (f *File) SetCellValue(sheet, cell string, value interface{}) error {
    ref := sheet + "!" + cell

    // 只清除依赖该单元格的公式缓存
    if deps, ok := f.cellDependents.Load(ref); ok {
        for _, depCell := range deps.([]string) {
            f.calcCache.Delete(depCell)
        }
    }

    // 清除自身缓存
    f.calcCache.Delete(ref)
}
```

### 优化 2: 批量操作 API ⭐⭐⭐⭐

```go
// 新增批量设置接口，延迟缓存清除
func (f *File) SetCellValues(sheet string, values map[string]interface{}) error {
    // 批量设置所有值
    for cell, value := range values {
        // ... 设置值，不清除缓存
    }

    // 最后一次性清除缓存
    f.calcCache.Clear()
    return nil
}

// 使用方式
values := map[string]interface{}{
    "A1": 100,
    "A2": 200,
    // ... 40000 个单元格
}
f.SetCellValues("Sheet1", values)  // 只清除一次缓存
```

### 优化 3: 精确的范围缓存清除 ⭐⭐⭐

```go
func (f *File) clearCellCache(sheet, cell string) {
    ref := fmt.Sprintf("%s!%s", sheet, cell)
    f.calcCache.Delete(ref)

    col, row, _ := CellNameToCoordinates(cell)

    // 精确检查范围是否包含该单元格
    f.rangeCache.Range(func(key, value interface{}) bool {
        cacheKey := key.(string)
        if !strings.HasPrefix(cacheKey, sheet+"!") {
            return true
        }

        // 解析范围: "Sheet1!A1:Z100"
        rangeStr := strings.TrimPrefix(cacheKey, sheet+"!")
        if coords, err := rangeRefToCoordinates(rangeStr); err == nil {
            // 检查单元格是否在范围内
            if col >= coords[0] && col <= coords[2] &&
               row >= coords[1] && row <= coords[3] {
                f.rangeCache.Delete(key)
            }
        }
        return true
    })
}
```

## 性能影响估算

假设 40k×100 场景，批量导入数据：

| 策略 | 缓存清除次数 | 估算性能 |
|------|------------|---------|
| **当前实现** | 400 万次 (每个 SetCellValue) | 极慢（不可用） |
| **优化 1: 延迟清除** | 1 次 (计算前) | 正常 (~10s) |
| **优化 2: 批量 API** | 1 次 (批量结束) | 正常 (~10s) |
| **优化 3: 依赖追踪** | 按需清除 | 最优 (<10s) |

## 推荐实施顺序

1. **立即实施**: 批量操作 API (SetCellValues)
   - 简单易实现
   - 立即解决批量导入性能问题

2. **中期实施**: 延迟清除策略
   - 兼容现有 API
   - 自动优化性能

3. **长期实施**: 依赖追踪系统
   - 需要重构计算引擎
   - 最优性能，但复杂度高
