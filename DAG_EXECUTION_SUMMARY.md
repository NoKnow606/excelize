# DAG Execution Summary

Date: 2026-04-24

## 1. 目标

这份文档从 DAG 角度总结当前依赖图重算方案，重点回答：

- 当前 DAG 重算是怎么工作的
- 分层、调度、批量优化分别承担什么职责
- 当前实现有哪些注意点、风险和缺陷
- DAG 体系与未来 block cache / source store 应该如何衔接

## 2. 当前 DAG 方案概览

主要入口：

- `RecalculateAllWithDependency()`
- `RecalculateSheetWithDependency()`
- `RecalculateAffectedByColumns()`
- `RecalculateAffectedByCellsWithExclusion()`

这里需要特别说明：

- `RecalculateAllWithDependency()` 虽然归类在 DAG 执行体系中
- 但它内部并不是“纯 DAG 调度”
- 而是组合了多段 batch calculation / batch optimization 逻辑

因此它既是：

- DAG execution entry

也是：

- 当前最重要的一条 batch recalculation path

主干流程：

1. 获取 `recalcMu`
2. 清理相关缓存
3. 构建依赖图 `dependencyGraph`
4. 分配 level
5. `calculateByDAG()` 逐层处理
6. 每层先预计算简单公式
7. 再做 SUMIFS / INDEX-MATCH / AVERAGE(OFFSET) 的 batch 优化
8. 最后对当前层使用 `DAGScheduler` 动态调度

这套设计并不是“只靠 DAG 调度器”，而是：

- DAG 分层
- 层内 batch 优化
- 层内动态依赖调度

三者组合。

这也是为什么它不能简单等同于：

- `CalcCellValuesConcurrent()` 这种普通批量算值接口

`RecalculateAllWithDependency()` 的外层职责是“全量依赖重算”，而 batch calculate 是它的内部执行手段。

## 3. 依赖图与分层逻辑

### 3.1 `dependencyGraph`

位置：

- `batch_dependency.go`

关键内容：

- `nodes map[string]*dependencyNode`
- `levels [][]string`

其中每个 node 包含：

- `cell`
- `formula`
- `dependencies`
- `level`

### 3.2 依赖提取

当前实现会从公式里提取：

- 普通 cell 依赖
- cross-sheet 依赖
- 部分范围依赖
- 列级虚拟依赖 `COLUMN:Sheet!Col`

这里有一个非常关键的折中：

- 对很大范围，不一定展开到每个 cell
- 而是退化成列级依赖

这样做的好处是：

- 降低依赖图规模

代价是：

- 精度下降
- 依赖边会更粗

### 3.3 level 分配

`assignLevels()` 使用的是：

- BFS 风格 topological assignment
- reverse dependency index
- unresolved dependency count

此外还会跟踪：

- 每列 unresolved 数量
- 列级虚拟依赖何时整体 resolved

### 3.4 circular dependency 处理

如果 node 最终 level 仍为 `-1`：

- 会被归到最后一个 circular level

后续在某个 level 内若 `NewDAGSchedulerForLevel()` 发现：

- level 有公式
- 但没有任何 ready node

就会：

- 认为 level 内存在循环依赖
- 退回 `parallelCalculateCells()`

这说明当前实现对循环依赖的策略不是完整迭代求解，而是：

- 尽量检测
- 然后做保守 fallback

## 4. `calculateByDAG()` 的实际执行模型

### 4.1 先确保公式元数据准备完毕

在进入并发计算前，会先：

- `setArrayFormulaCells()`
- 并设置 `formulaChecked = true`

目的是：

- 避免 worker 在并行阶段再触发 array formula 初始化

### 4.2 `WorksheetCache` 采用 lazy 模式

当前 `buildWorksheetCache()` 已经不再整表 `LoadSheet()`，而是：

- 只跟踪可能涉及的 sheet
- 让实际读取走后续的 `PreloadColumnRange()` 或按需读取

这是当前实现里比较重要的一步收缩。

### 4.3 每个 level 的顺序

当前 level 内执行顺序是：

1. 检测列范围访问模式
2. 预读取热点列范围
3. `preCalculateSimpleFormulas()`
4. `batchOptimizeLevelWithCache()`
5. `DAGScheduler` 计算剩余项

也就是说：

- DAG 并不是一上来就全量动态执行
- 而是先用批量优化和预计算尽量吃掉可优化部分

## 5. `DAGScheduler` 的职责与实现

位置：

- `batch_dag_scheduler.go`

### 5.1 负责的不是“整图”，而是“当前层内部动态调度”

`NewDAGSchedulerForLevel()` 的关键点是：

- 只关注当前 level 内部依赖
- 跨 level 的依赖视为已满足

所以它更准确的定位是：

- level-local dynamic scheduler

### 5.2 ready queue 模型

当前 scheduler 会维护：

- `readyQueue chan string`
- `dependencyCount map[string]int`
- `dependents map[string][]string`
- `completedCount`
- `inFlightCount`

工作方式：

- 初始无依赖节点入 ready queue
- worker 消费 ready queue
- 一个公式完成后，给 dependents 做 `dependencyCount--`
- 变成 0 的 dependent 入队

### 5.3 worker 的执行顺序

单个公式执行时：

1. 先看 `worksheetCache`
2. 再看 `calcCache`
3. 再从 graph 取公式
4. 调 `CalcCellValueWithSubExprCache()`
5. 结果写入 caches 和 worksheet
6. 通知 dependents

这个顺序很重要，因为它说明：

- DAG scheduler 不是孤立运行
- 它高度依赖 overlay cache、result cache 和 sub-expression cache

## 6. 当前设计的优点

### 6.1 把“全量重算”拆成了可组合的多个优化层

现有 DAG 方案并不是单纯拓扑排序，而是叠加了：

- 依赖图
- level 合并
- simple formula 预计算
- pattern-based batch optimization
- layer-local dynamic scheduling

这在性能工程上是比较务实的路径。

### 6.2 对大范围列依赖做了图规模控制

通过 `COLUMN:` 虚拟依赖，可以避免：

- 一些超大列范围被直接展开成海量边

这对图规模和构图时间非常关键。

### 6.3 对循环依赖有显式 fallback

虽然不是完整 iterative calc，但至少没有假设“图一定是纯 DAG”。

## 7. 注意点

### 7.1 这里的 DAG 不是 Excel 完整 calc engine 的等价物

当前实现更适合：

- 大量普通公式
- 批量 lookup
- 层级相对清晰的依赖图

对以下场景要更谨慎：

- 循环引用
- 大量 volatile / 动态引用
- 极复杂的结构性公式传播

### 7.2 level 是优化单元，不是严格的最小执行单元

由于存在：

- `mergeLevels()`
- 列级虚拟依赖
- 层内 batch 优化

所以 level 不是一个“绝对最细、绝对精确”的依赖边界，而是：

- 为性能工程服务的折中执行单元

### 7.3 `WorksheetCache` 现在更像 overlay，不是完整 source cache

当前 DAG 路径对 `WorksheetCache` 的使用已经更偏向：

- 保存最新计算结果
- 为后续读取提供覆盖层

这与 Memory Summary 的方向是一致的。

## 8. 风险与缺陷

### 8.1 依赖提取精度与图规模之间存在硬折中

列级虚拟依赖有明显收益，但也带来问题：

- 依赖可能被放粗
- 实际只依赖少量 cell 的公式，也可能被整列阻塞
- 增量重算时可能扩大影响面

这不是实现错误，而是精度换规模的选择。

### 8.2 构图和 level 分配本身可能很重

当前在真正计算前，要先完成：

- 扫描公式
- 提取依赖
- 建 reverse index
- assign levels

对于公式量极大的 workbook，这部分：

- 本身就可能成为 CPU 和内存大头

### 8.3 `readyQueue` 溢出时当前实现会丢任务

`notifyDependents()` 中使用的是：

- 非阻塞 `select`
- `default` 分支仅记录日志：`Ready queue full, dropping ...`

这意味着理论上如果 ready queue 真满了，会出现：

- dependent 没有入队
- 公式永久不执行

虽然当前 queue 设得很大，作者预期“不应该发生”，但这仍然是一个真实缺陷点。

### 8.4 circular level 的 fallback 不是完整正确性方案

当前发现 level 内无 ready node 时，会退回：

- `parallelCalculateCells()`

但这并不等于：

- 已经正确实现 Excel 的迭代循环引用求解

因此对循环依赖的支持仍然偏保守。

### 8.5 fallback worker 数固定为 `10`

`parallelCalculateCells()` 当前固定：

- `numWorkers := 10`

这与其他地方按 `runtime.NumCPU()` 自适应不同。

问题是：

- 在小机器上可能偏多
- 在大机器上可能偏少
- 策略不一致，增加调优难度

### 8.6 scheduler 结果写回会持续触发 worksheet 写锁

每个公式执行后会：

- 写 caches
- 写 worksheet

即使 DAG 调度并发度很高，最终写回仍然会在 worksheet 级锁处串行化部分热点。

这会导致：

- 计算本身可能并行
- 但高写密度场景下仍有锁竞争

### 8.7 多层缓存叠加后，一致性复杂度已经很高

当前 DAG 路径会用到：

- `WorksheetCache`
- `calcCache`
- `rangeCache`
- `SubExpressionCache`
- 多类 lookup index cache

优点是命中率更高。

代价是：

- 任意一层失效不完整，都可能产生“算得快但结果旧”的问题

## 9. 对 block cache / source store 的意义

当前 DAG 路径最缺的一层，不是更多调度技巧，而是：

- 一个稳定、轻量、可跨请求复用的大 sheet 源数据底座

如果 `SheetSourceStore` 落地，DAG 最直接的收益会在：

- `PreloadColumnRange()` 不再优先依赖完整 worksheet object
- batch SUMIFS / INDEX-MATCH 可以直接读 source blocks
- 增量重算可以按 block 定位 dirty 源数据

换句话说：

- DAG 现在已经有不错的“执行器”
- 缺的是更合适的“数据底盘”

## 10. 建议

### 第一优先级

1. 修复 `readyQueue` 满时可能丢任务的问题。
2. 给 circular dependency fallback 单独补文档说明，不要让人误以为已完整支持迭代循环求解。
3. 统一 `parallelCalculateCells()` 的 worker 策略。

### 第二优先级

1. 为依赖提取精度 vs 图规模做显式配置或观测。
2. 给 `COLUMN:` 虚拟依赖策略补更多测试与说明。
3. 让 `PreloadColumnRange()` / batch optimizer 优先接入 source store。

## 11. 一句话总结

当前 DAG 体系的优势很明显：

- 已经具备依赖图、层内调度、批量公式优化三层能力

但它的核心问题也同样明确：

- 图构建和缓存一致性复杂
- 循环依赖处理仍然保守
- 依赖精度与规模控制是硬折中
- 缺少 block cache / source store 这层更适合大 sheet 的源数据基础设施
