# Lock Model Summary

Date: 2026-04-24

## 1. 目标

这份文档从 lock / 并发控制角度总结当前 `excelize` 仓库里的实现现状，重点回答：

- 当前有哪些主要锁和并发保护边界
- 它们分别保护什么
- 现有实现的注意点、风险和缺陷在哪里
- 后续做 block cache / batch calculate / DAG 优化时，锁层面需要遵守什么约束

## 2. 核心结论

当前的并发模型不是“整个 `*File` 完全线程安全”，而是：

- `*File` 整体默认不承诺任意并发安全
- 某些热点路径通过局部锁和并发安全缓存实现“有限并发”
- 重算路径通过 `recalcMu` 串行化
- worksheet 级写入依赖 `ws.mu`
- cache 层依赖 `sync.Map`、`sync.RWMutex`、`lruCache.mu`

因此最重要的认知不是“已经全局线程安全”，而是：

- 当前是一个“局部加锁 + 约定使用方式”的模型

## 3. 主要锁与职责

### 3.1 `File.mu`

位置：

- `excelize.go`

职责：

- 保护 `File` 级别的一部分状态访问
- 常见于 `workSheetReader()` 调用前后
- 用于保护 `formulaChecked` / `inBatchMode` 等状态修改

典型场景：

- `CalcCellValuesDependencyAware()` 初始化 array formula 状态
- `calculateByDAG()` 初始化 `setArrayFormulaCells()`
- `persistFormulaResult()` 中读取 worksheet 对象
- `SetCellValues()` 切换 `inBatchMode`

注意：

- `File.mu` 并没有覆盖整个 `File` 的全部读写生命周期
- 它更像一个局部临界区保护，而不是“整个 `File` 的总锁”

### 3.2 `File.recalcMu`

位置：

- `excelize.go`
- `batch_dependency.go`

职责：

- 防止同一个 `*File` 上同时发生多次依赖图重算

当前覆盖的入口：

- `RecalculateAllWithDependency()`
- `RecalculateSheetWithDependency()`
- `RecalculateAffectedByColumns()`
- `RecalculateAffectedByCellsWithExclusion()`

结论：

- 当前 DAG 重算是“单 `File` 串行重算”模型
- 不是“多个 DAG 重算并发跑在同一个 `File` 上”

### 3.3 `xlsxWorksheet.mu`

位置：

- `xmlWorksheet.go`

类型：

- `sync.RWMutex`

职责：

- 保护 worksheet 对象树本身
- 防止并发修改 `SheetData.Row`、`C`、`V`、`T` 等结构

典型场景：

- `persistScalarFormulaResult()` 写回结果时 `ws.mu.Lock()`
- `persistSQLFormulaResult()` 写 spill range 时 `ws.mu.Lock()`
- `GetRangeValuesConcurrent()` 整段读取时持有 worksheet 锁
- 大量 cell 写操作也会先进入 worksheet 锁

### 3.4 cache 自带锁

#### `WorksheetCache.mu`

位置：

- `worksheet_cache.go`

职责：

- 保护 `map[sheet]map[cellRef]formulaArg`

#### `calcCache`

- `sync.Map`
- 用于公式结果缓存

#### `rangeCache`

- `lruCache` 自带 `mu`
- 用于范围矩阵缓存

#### 其他索引缓存

- `matchIndexCache`
- `ifsMatchCache`
- `rangeIndexCache`

这些缓存大多是并发安全容器，但注意：

- cache 容器线程安全，不等于“整个业务路径线程安全”

## 4. 当前锁模型的真实边界

### 4.1 `File` 默认不是任意并发安全对象

这一点在测试里已经写得很明确：

- `concurrency_test.go` 明确说明 `File` object is NOT designed for concurrent access without external locking

当前更合理的理解是：

- 同一个 `*File` 上，普通读写 API 混合并发仍然有风险
- 某些专项优化路径可以内部并发，但它们依赖额外前提

### 4.2 内部并发主要集中在“读计算”而不是“任意写操作”

例如：

- `CalcCellValuesConcurrent()`
- `CalcCellValuesDependencyAware()`
- `calculateByDAG()`

这些路径可以启 worker，但都有自己的限制：

- 要么是只读计算
- 要么写回经过 worksheet 级锁串行保护
- 要么整个重算被 `recalcMu` 串行化

### 4.3 当前锁更像“分层拼装”，不是统一锁协议

当前能看到的典型模式是：

- `f.mu` 保护 `workSheetReader()` 或初始化状态
- 然后切到 `ws.mu` 保护具体 worksheet 结构
- cache 再由各自容器保护

这意味着：

- 调用路径必须维持一致的锁顺序
- 否则后续很容易引入死锁或锁反转风险

## 5. 关键实现观察

### 5.1 `recalcMu` 是最重要的“粗粒度保护”

它的价值很直接：

- 避免两个重算同时清空 `calcCache`
- 避免两个重算同时写同一批公式结果
- 避免依赖图、level、worksheetCache 生命周期重叠

没有它的话，DAG 重算路径几乎必然出现缓存错乱和结果覆盖。

### 5.2 worksheet 写入保护做得比 `File` 总体保护更明确

例如 `persistScalarFormulaResult()` 的顺序是：

- `f.mu` 下拿到 worksheet
- `ws.mu.Lock()`
- `prepareWorksheetCell()`
- 写 `c.V` 和 `c.T`
- 解锁

这说明当前实现真正认为“需要被严格保护”的对象，是：

- `xlsxWorksheet` 的内部行列结构

### 5.3 有些读路径仍然使用独占锁

例如：

- `GetRangeValuesConcurrent()` 用的是 `ws.mu.Lock()`，不是 `RLock()`

结果是：

- 虽然内部 worker 并发读 range
- 但对外仍然以独占锁方式占住整个 worksheet

这会降低“多读并行”的上限。

## 6. 注意点

### 6.1 不要把 cache 的线程安全误解成 `File` 的线程安全

例如：

- `calcCache` 是 `sync.Map`
- `WorksheetCache` 有 `RWMutex`
- `rangeCache` 有内部锁

但如果并发 goroutine 同时：

- 调 `SetCellValue`
- 调 `RecalculateAllWithDependency`
- 调 `MoveRows` / `InsertCols`

问题不在 cache，而在 worksheet 对象和结构变更时序。

### 6.2 `formulaChecked` 是一个全局状态位

当前多个路径会做：

- 检查 `!f.formulaChecked`
- 然后在 `f.mu` 下执行 `setArrayFormulaCells()`

这能避免重复初始化，但也意味着：

- 这是一个全局一次性状态
- 如果未来 array formula 相关结构支持增量变更，这个位可能不够表达真实状态

### 6.3 `inBatchMode` 只是 cache invalidation 开关，不是事务锁

`SetCellValues()` 会：

- 打开 `inBatchMode`
- 批量调用 `setCellValue`
- 最后统一清理 `calcCache` / `rangeCache`

这很有用，但要明确：

- 它不是“写事务隔离”
- 也不意味着同一个 `File` 可以安全地和其他写操作并发执行

## 7. 风险与缺陷

### 7.1 锁粒度不统一，后续改动容易引入死锁

现状是多层锁并存：

- `recalcMu`
- `f.mu`
- `ws.mu`
- cache locks

这本身没问题，但仓库里没有一个统一的 lock order 文档。风险是：

- 新代码若出现 `ws.mu -> f.mu` 的反向顺序
- 很容易和现有 `f.mu -> ws.mu` 路径形成锁反转

### 7.2 callback 存在重入风险

`persistScalarFormulaResult()` 在写回完成后可能触发：

- `OnCellCalculated`

当前 callback 调用点不持有 `ws.mu`，这是好的；但仍需注意：

- 若 callback 内再次对同一 `File` 发起重算
- 会与外层 `recalcMu` 形成阻塞
- 在复杂调用链中可能表现为“卡住”或自等待

这更像一个使用约束缺失，而不是当前已证实的死锁 bug。

### 7.3 读路径使用独占 worksheet 锁，影响读并发上限

像 `GetRangeValuesConcurrent()` 这类函数：

- 内部并发读
- 但外层是 `ws.mu.Lock()`

风险是：

- 大范围读会阻塞同 sheet 的其他读写
- 在 block cache 落地后，这种独占读可能成为新的瓶颈

### 7.4 `recalcMu` 保护了正确性，也限制了吞吐

它的优点是简单可靠。

它的代价是：

- 同一 `File` 上任何两次依赖重算都不能并行
- 多 session 如果共享一个 `CachedFile`，重算会天然串行

这不一定是 bug，但确实是架构上的吞吐上限。

### 7.5 缺少“锁 + 版本 + overlay”统一设计文档

随着 block cache / source store 引入，未来一定会出现：

- shared store 锁
- session overlay 锁
- worksheet object 锁
- cache invalidation 锁

如果没有统一约束，锁复杂度会明显上升。

## 8. 对 block cache 方案的约束

如果后续引入 `SheetSourceStore` / block cache，锁模型建议遵守：

- block store 锁不要和 `ws.mu` 长时间嵌套
- 优先把 source store 作为独立层管理，不直接复用 worksheet 锁
- session overlay 和 shared source store 分离
- 写入先落 overlay，提交后再更新 shared block
- block dirty / version 更新要和 `rangeCache` / index cache 失效联动

## 9. 建议

### 第一优先级

1. 单独补一份 lock order 约定文档。
2. 标注哪些 API 可以在同一 `File` 上并发读，哪些不行。
3. 给 callback 场景补 usage note，禁止在回调里直接递归重算同一 `File`。

### 第二优先级

1. 评估 `GetRangeValuesConcurrent()` 是否可改为 `ws.mu.RLock()`。
2. 评估 `formulaChecked` 是否需要从布尔值升级为更细粒度状态。
3. 在 block cache 落地前，先定义 shared store / overlay / worksheet object 三层的锁边界。

## 10. 一句话总结

当前 lock 模型的优点是：

- 简单
- 局部有效
- 能支撑当前 batch / DAG 方案落地

但它的核心限制也很明显：

- `File` 不是全局线程安全对象
- 锁粒度不完全统一
- 未来引入 block cache 和多 session 共享后，需要更明确的锁协议与版本边界
