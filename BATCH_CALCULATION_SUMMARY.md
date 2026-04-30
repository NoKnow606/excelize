# Batch Calculation Summary

Date: 2026-04-24

## 1. 目标

这份文档从 batch calculate 角度总结当前实现，重点覆盖：

- 现有 batch 计算 / batch 写入路径有哪些
- 它们各自优化了什么
- 使用时的注意点
- 已知风险、缺陷和适用边界

## 2. 当前主要 batch 路径

### 2.0 需要先说明：`RecalculateAllWithDependency()` 也属于一条 batch recalculation 路径

虽然 `RecalculateAllWithDependency()` 在架构上更适合归类到 DAG 执行体系，但从执行内容看，它本质上也是当前最重要的一条 batch recalculation 路径。

原因是它并不是“只做依赖拓扑调度”，而是会在 `calculateByDAG()` 中组合使用：

- `preCalculateSimpleFormulas()`
- `batchOptimizeLevelWithCache()`
- `PreloadColumnRange()`
- SUMIFS / INDEX-MATCH / AVERAGE(OFFSET) 的批量优化
- 层内 `DAGScheduler`

因此更准确的理解应该是：

- `CalcCellValuesConcurrent()` / `CalcCellValuesDependencyAware()` 更像独立 batch API
- `RecalculateAllWithDependency()` 更像“依赖图驱动的 batch recalculation engine entry”

这也是为什么它既应出现在 DAG 文档里，也应在 batch calculation 文档里被单独点名。

### 2.1 批量写入：`SetCellValues()`

位置：

- `cell_batch.go`

核心思路：

- 批量设置多个 cell
- 期间通过 `inBatchMode` 抑制单次 `setCellValue()` 的 cache 清理
- 最后统一清理 `calcCache` / `rangeCache`

它优化的是：

- 批量写很多 cell 时，避免每次写都清缓存

### 2.2 并发批量算值：`CalcCellValuesConcurrent()`

位置：

- `calc_optimized.go`

核心思路：

- 把一批 cell 分发给 worker goroutine
- 每个 worker 调 `CalcCellValue()`
- SQL 公式的结果先收集，最后单线程写回

它优化的是：

- 一批相互独立或弱相关公式的吞吐

### 2.3 依赖感知批量算值：`CalcCellValuesDependencyAware()`

位置：

- `calc_dependency_aware.go`

核心思路：

- Phase 1：按列取首个 cell 进行 warmup
- Phase 2：其余 cell 并发计算
- 使用缓存预热减少 lookup / range 的重复成本

它优化的是：

- 多公式共享相同依赖表时的批量计算性能

### 2.4 依赖感知批量算值并写回：`CalcAndUpdateCellValuesDependencyAware()`

位置：

- `calc_dependency_aware.go`

核心思路：

- 与上一条类似
- 但会把结果写回 worksheet，并可触发 `OnCellCalculated`

### 2.5 DAG 路径中的“批量优化” 

位置：

- `batch_dependency.go`

核心思路：

- `preCalculateSimpleFormulas()` 先算简单公式
- `batchOptimizeLevelWithCache()` 对 SUMIFS / INDEX-MATCH / AVERAGE(OFFSET) 做批量优化
- 再交给 DAG scheduler 处理剩余计算

这一条严格说属于 DAG 体系，但它本质上也是 batch calculate 的一部分，因此要和普通 batch API 区分开看。

### 2.6 全量依赖重算：`RecalculateAllWithDependency()`

位置：

- `batch_dependency.go`

核心思路：

- 获取 `recalcMu`
- 清理旧的 `calcCache` / `rangeCache`
- 构建 workbook 级依赖图
- 调用 `calculateByDAG()` 逐层执行

它和普通 batch API 的区别在于：

- 输入不是“某一批 cell 列表”
- 而是“整个 workbook 中所有公式”
- batch optimization 是其内部执行手段，而不是外部 API 语义本身

它和 `CalcCellValuesDependencyAware()` 的关系可以概括为：

- 后者是“调用方指定一批 cell”的 batch calculation
- 前者是“系统自己找出所有公式并按依赖批量重算”的 batch recalculation

### 2.7 增量依赖重算：`RecalculateAffectedByColumns()` / `RecalculateAffectedByCellsWithExclusion()`

位置：

- `batch_dependency.go`

核心思路：

- 先基于更新输入定位受影响公式集合
- 再构建过滤后的依赖图或受影响子图
- 只清理受影响公式相关缓存
- 最后复用 `calculateByDAG()` 执行增量重算

主要入口包括：

- `RecalculateAffectedByColumns(updatedColumns map[string]bool)`
- `RecalculateAffectedByCellsWithExclusion(updatedCells, excludeCells map[string]bool)`

它和全量依赖重算的区别在于：

- 全量重算默认处理整个 workbook 的公式集合
- 增量重算只处理受影响子集

它和普通 batch API 的区别在于：

- 输入不是“要算哪些目标 cell”
- 而是“哪些源列 / 源单元格发生了变化”

因此它更准确的定位是：

- dependency-driven incremental batch recalculation

## 3. 当前设计的主要优点

### 3.1 批量写入路径简单直接，收益明确

`SetCellValues()` 的价值很清楚：

- 把多次 cache invalidation 收缩成一次
- 大幅降低大批量写入成本

这类优化风险相对小，收益稳定。

### 3.2 批量计算路径已经分层

当前 batch calculate 不是单一方案，而是：

- 普通并发算值
- 依赖感知 warmup + 并发
- DAG 内分层 batch 优化

这说明代码已经在区分：

- 独立计算任务
- 共享依赖任务
- 有公式依赖图任务

### 3.3 SQL 公式写回做了特殊处理

在 `CalcCellValuesConcurrent()` 中：

- SQL 公式结果先收集
- 最后单线程 `persistFormulaResult()`

这样避免了：

- 多 worker 同时修改 worksheet spill range

这是很重要的实现细节。

## 4. 注意点

### 4.0 不要把 `RecalculateAllWithDependency()` 和普通 batch API 混为一类

这条路径虽然内部大量使用 batch 技术，但它的职责层次更高：

- 普通 batch API 解决的是“给定一批 cell，如何更快算完”
- `RecalculateAllWithDependency()` 解决的是“整本 workbook 如何按依赖顺序完成重算”

因此：

- 它不是 `CalcCellValuesDependencyAware()` 的简单放大版
- 也不是只靠 worker pool 的普通并发计算
- 而是 DAG + batch optimization + cache overlay 的组合执行框架

### 4.1 增量重算也不是普通 batch API

`RecalculateAffectedByColumns()` 和 `RecalculateAffectedByCellsWithExclusion()` 虽然只处理公式子集，但它们仍然属于“重算引擎入口”，而不是简单的批量算值函数。

原因是它们内部仍然会做：

- 受影响范围分析
- 反向依赖传播
- 过滤依赖图构建
- cache 定向失效
- DAG 执行

所以它们应被理解成：

- 增量版的 batch recalculation

而不是：

- 对若干 cell 直接并发调用 `CalcCellValue()`

### 4.1 batch calculate 不是一个统一语义

当前不同 API 的行为差异比较大：

- 有的只返回结果，不写回 worksheet
- 有的会写回并触发 callback
- 有的会清理 `rangeCache`
- 有的会为 SQL 公式做额外持久化

所以调用方不能把这些 API 当作“只是不同性能版本”。

### 4.2 warmup 是启发式，不是精确依赖分析

`CalcCellValuesDependencyAware()` 的 warmup 策略是：

- 每列先算一个 cell

这对很多表格是有效的，但它只是启发式：

- 如果同列公式并不共享依赖
- 或共享依赖不按列分布
- warmup 价值就会下降

### 4.3 `workSheetReader()` 预热有内存代价

依赖感知批量计算会：

- 预热当前 sheet
- 还会启发式预热前 3 个 sheet

优点是：

- 降低后续 namespace conversion / worksheet 首次读取抖动

代价是：

- 会提前把更多 worksheet 对象拉进内存

这对大文档不一定总是划算。

### 4.4 batch 写入只统一清理部分缓存

`SetCellValues()` 统一清理的是：

- `calcCache`
- `rangeCache`

这很合理，但如果后续继续引入：

- block cache
- lookup index caches
- session overlay

那么 batch 写入后的失效链需要重新设计，不能只停留在这两层。

## 5. 风险与缺陷

### 5.1 `SetCellValues()` 只返回第一个错误

当前实现会记录 `firstError` 并返回。

优点：

- 简单

缺点：

- 调用方无法直接知道这批写入里还有多少 cell 失败
- 诊断复杂批量写入问题时信息不足

### 5.2 `inBatchMode` 不是并发隔离机制

`SetCellValues()` 的设计目标是性能，不是并发安全。

如果同一个 `File` 上还有其他 goroutine 同时写：

- `inBatchMode` 不能提供事务边界
- 也不能保证中途读到的一定是完整批次状态

### 5.3 `CalcCellValuesConcurrent()` 适合独立任务，不适合强依赖任务

它没有显式依赖图调度，意味着：

- 如果一批公式之间互相依赖
- 仅靠并发算值不能保证最佳命中和最小重复工作

这不是 bug，但容易被误用。

### 5.4 依赖感知 batch 的预热策略可能放大内存压力

当前路径会提前：

- 读取当前 sheet
- 读取前 3 个 sheet

在小文件上通常没问题。

但在大 sheet / 多热点 sheet 场景中，这可能：

- 加快首轮访问
- 同时增加驻留对象和峰值内存

### 5.5 `rangeCache` 在 batch 结束后被清空，会降低跨批次复用

`CalcCellValuesDependencyAware()` 和相关路径会在结束后清空 `rangeCache`。

好处：

- 释放大矩阵，避免长期堆积

代价：

- 下一批相似请求无法直接复用这些矩阵

这是一个明确的吞吐 vs 内存折中。

### 5.6 回调顺序不保证稳定

在 `CalcAndUpdateCellValuesDependencyAware()` 这类会写回且可触发 callback 的路径上：

- 并发 worker 写回时，回调触发顺序可能不是稳定顺序

因此：

- 调用方不能假设 callback 顺序与输入 cell 顺序一致

### 5.7 SQL 公式和普通公式的处理语义不完全一致

当前代码对 SQL 公式有特殊路径：

- 结果收集
- spill range 清理
- 最终持久化

这很必要，但也说明：

- batch calculate 并不是对所有公式类型都完全同构
- 调优或重构时必须把 SQL 公式单独考虑

### 5.8 增量重算的“精度”依赖于依赖提取粒度

当前增量路径会利用：

- 普通 cell 依赖
- `COLUMN:` 级依赖
- `SHEET:` 级依赖（如 SQL formula 的源 sheet 标记）

这意味着：

- 增量重算不是总能做到精确最小集合
- 某些场景会偏保守，宁可多算一些公式

典型例子：

- `SQL("select * from A")` 在 `A` sheet 有更新时会被视为受影响
- 即使 SQL 实际只关心 `A` 的部分行列

## 6. 当前缺陷中最值得关注的点

### 6.1 batch 路径之间行为分叉较多

现在至少有：

- `CalcCellValuesConcurrent()`
- `CalcCellValuesDependencyAware()`
- `CalcAndUpdateCellValuesDependencyAware()`
- DAG 内 batch 优化路径

这些路径都在做“批量计算”，但：

- 预热方式不同
- 缓存使用方式不同
- 写回方式不同
- SQL 公式处理不同

长期风险是：

- 修一个路径的 bug，另外三个路径语义漂移

### 6.2 batch 优化仍然大量依赖 worksheet object

即使已经有了批量优化，当前很多读取仍然建立在：

- `workSheetReader()`
- 完整 worksheet 对象
- `WorksheetCache` overlay

这意味着：

- batch calculate 的性能优化和内存优化还没有彻底解耦

### 6.3 热点公式模式识别是启发式实现

例如：

- SUMIFS / AVERAGEIFS 抽取
- INDEX-MATCH 模式识别
- AVERAGE(OFFSET) 优先级

这能覆盖大量常见场景，但不是完整公式语义引擎。

风险是：

- 模式越多，维护成本越高
- 很容易出现边缘公式没命中优化，或被错误归类

### 6.4 增量路径容易受“粗粒度依赖”影响而扩大重算面

尤其是以下两类依赖：

- `COLUMN:` 虚拟列依赖
- `SHEET:` SQL source 依赖

它们很适合控制构图复杂度、保证不漏算。

但代价是：

- 一次很小的源数据变动，也可能触发较大一片公式重新进入增量重算

## 7. 对 block cache 方案的影响

如果后续引入 `SheetSourceStore` / block cache，那么 batch calculate 最应该受益的地方是：

- warmup 不再主要靠 `workSheetReader()` 预热对象
- lookup / range 可直接从 source blocks 取值
- SQL 之外的大部分源值访问都不必依赖整表对象展开

这会把当前 batch 计算的重心，从：

- “怎么减少重复解析 worksheet object”

转向：

- “怎么最大化复用共享 source blocks 和派生索引”

## 8. 建议

### 第一优先级

1. 给各类 batch calculate API 补一张语义矩阵。
2. 明确每条路径是否写回、是否触发 callback、是否清缓存。
3. 给大 sheet 场景增加内存提示，避免盲目预热 worksheet。

### 第二优先级

1. 把 warmup heuristic 与 source store 预热解耦。
2. 统一 SQL 公式与普通公式的批量执行约束说明。
3. 评估是否要让 batch 写入返回多错误集合，而不是只返回首错。

## 9. 一句话总结

当前 batch calculate 体系已经能带来显著性能收益，但它的主要问题不是“没有优化”，而是：

- 路径较多
- 语义分叉明显
- 仍然偏依赖 worksheet object 和启发式预热

后续如果要进一步做大 sheet 优化，batch calculate 最需要的是：

- 统一语义
- 统一失效策略
- 接上 block cache / source store 这层新的数据底座
