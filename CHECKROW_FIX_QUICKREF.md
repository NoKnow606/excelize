# CheckRow Index Out of Range 快速修复指南

## 问题

```
panic: runtime error: index out of range [23] with length 23
at rows.go:1091
```

## 已修复 ✅

**修改文件**: `rows.go:1069-1112`

**关键改动**:
1. 重新计算 `maxCol`（遍历所有 sourceList）
2. 添加边界检查（防止 `colNum-1` 越界）

## 测试验证

```bash
go test -run "TestCheckRow" -v
# ✅ 6/6 测试通过

go test -run "TestInsertRow" -v
# ✅ 所有 InsertRows 测试通过

go test -run "TestBatch" -v
# ✅ 所有批量操作测试通过
```

## 如果你还没更新库

在应用代码中添加 panic 恢复：

```go
defer func() {
    if r := recover(); r != nil {
        log.Printf("⚠️  checkRow panic: %v", r)
        // 可选：保存问题文件
    }
}()
```

## 相关文档

详细分析：[CHECKROW_FIX.md](./CHECKROW_FIX.md)
