# Memory Cache Strategy Summary

Date: 2026-04-24

## 1. 调研目标

本次调研围绕两个核心用户场景展开：

1. 场景 1：用户 A 打开某个大 worksheet 后，用户 B 再访问同一个文档时，不应该重新从底层存储和 XML 全量读取，而应该尽量复用缓存。
2. 场景 2：全量重计算时，不应该每个公式都重新从 worksheet XML 读取依赖单元格，而应该从缓存中快速索引到值。

这份总结同时覆盖：

- `excelize` 库本身的内存与计算缓存实现
- `excelize-mcp` 应用层的文件缓存、worksheet 加载与多用户复用路径

## 2. 关键结论

### 2.1 现在的 OOM 风险主因

当前真正的 OOM 风险，不在 `.xlsx` 压缩包大小，而在一张大 sheet 被同时展开成多层内存结构：

1. sheet XML 原始字节
2. `xlsxWorksheet` 完整对象树
3. `WorksheetCache` 中的逐单元格缓存
4. 公式求值过程中的范围矩阵、索引结构与临时对象

最危险的链路是：

- `readBytes(name)` 读整张 sheet XML
- `workSheetReader(sheet)` 解码整张 worksheet
- `WorksheetCache.LoadSheet()` 再把源值复制进 `map[string]formulaArg`

所以风险不是“必须扫描 XML”，而是“扫描后保留了多份重型表示”。

### 2.2 现有缓存很多，但缺少一个关键层

项目里已经有很多缓存，但职责分散：

- worksheet 对象缓存
- shared strings 缓存
- 公式结果缓存
- 范围矩阵缓存
- MATCH / SUMIFS / range 索引缓存
- batch 计算期间的 `WorksheetCache`
- `excelize-mcp` 全局 `FileCache`

这些缓存分别解决了“少重复解析”“少重复算公式”“少重复建索引”等问题。

但是还缺一个真正能同时满足两个目标场景的缓存层：

- 大 sheet 的可复用源数据缓存层
- 既能在单次重算中复用，也能在多个用户请求之间复用

### 2.3 结论不是“每次都重新读 sheet”

不能把优化方向理解成“每次需要值时都重新从 XML 读”。

正确模型应该是：

- 一次重算任务中，某张热点 sheet 第一次被访问时，最多构建一次源数据视图
- 之后本次重算内所有公式都复用它
- 文档未变化时，后续请求也可以复用
- 文档发生变更时，按 sheet 或按 block 失效，而不是整本强制重建

## 3. `excelize` 现有缓存实现与职责

### 3.1 Worksheet XML / Object Cache

主要位置：

- `excelize.go`
- `lib.go`

关键结构：

- `File.Pkg sync.Map`
- `File.Sheet sync.Map`
- `File.tempFiles sync.Map`
- `File.xmlAttr sync.Map`

职责：

- 缓存 workbook / worksheet XML 字节
- 缓存已经解码的 `xlsxWorksheet`
- 避免同一个 `File` 实例内重复解析

解决的问题：

- 小中型 sheet 的重复读取性能
- 普通读写 API 的对象复用

不足：

- 这是单个 `*excelize.File` 实例内的缓存，不是跨请求 / 跨用户共享缓存
- 对大 sheet 来说，缓存的是“重型对象”，本身内存成本很高

### 3.2 Shared Strings Cache

主要位置：

- `rows.go`

关键结构：

- `File.SharedStrings`
- `File.sharedStringsMap`
- `File.sharedStringTemp`

职责：

- 缓存 `sharedStrings.xml`
- 加速字符串单元格解码

解决的问题：

- 避免重复解析 shared strings

不足：

- 不是 worksheet 源数据缓存，只是配套基础缓存

### 3.3 Formula Result Cache

主要位置：

- `excelize.go`
- `calc_subexpr.go`
- `batch_dag_scheduler.go`

关键结构：

- `File.calcCache sync.Map`

职责：

- 缓存公式计算结果
- 缓存 raw / formatted 结果
- 缓存部分公式表达式计算结果

解决的问题：

- 相同公式结果重复计算
- 公式读值兼容性

不足：

- 它缓存的是“结果”，不是“大 sheet 的原始源数据”

### 3.4 WorksheetCache

主要位置：

- `worksheet_cache.go`

关键结构：

- `map[sheet]map[cellRef]formulaArg`

职责：

- batch / DAG 计算期间存储最近可见的单元格值
- 让已经重算出的公式结果优先覆盖 XML 中旧值
- 为 SUMIFS / INDEX-MATCH / AVERAGEIFS 等批量优化提供统一读口

解决的问题：

- 依赖重算时“读取最新值”
- 公式结果在一个重算会话内被后续公式复用

不足：

- `LoadSheet()` 仍然走整表加载
- 对大 sheet 使用 string key + `formulaArg` 的全量 cell map，内存成本非常高
- 它更适合当“重算 overlay cache”，不适合继续当“大 sheet 源数据全量缓存”

### 3.5 Range Matrix Cache

主要位置：

- `calc.go`
- `lru_cache.go`

关键结构：

- `File.rangeCache *lruCache`
- `calcContext.rangeCache`

职责：

- 缓存已解析的范围矩阵，如 `A1:C100`
- 加速公式中重复范围访问

解决的问题：

- 同一公式或多个公式重复读取相同范围

不足：

- 它缓存的是“范围矩阵结果”
- 前提仍然常常依赖完整 worksheet 已加载
- 不是统一的大 sheet 源数据层

### 3.6 Lookup / Index Caches

主要位置：

- `calc.go`

关键结构：

- `File.matchIndexCache`
- `File.ifsMatchCache`
- `File.rangeIndexCache`

职责：

- `MATCH` / `XLOOKUP` 哈希索引
- `SUMIFS` / `COUNTIFS` 条件匹配结果
- 范围值到位置的索引

解决的问题：

- 大量 lookup / 条件匹配时降低复杂度
- 将 O(n) 多次扫描降到 O(1) 或少量过滤

优点：

- 这些缓存与场景 2 高度相关
- 它们应该保留并继续复用

不足：

- 它们是在源数据已经可访问的前提下生效
- 还不能单独解决“大 sheet 如何安全复用源数据”

### 3.7 SubExpressionCache

主要位置：

- `calc_subexpr.go`

职责：

- 缓存 SUMIFS / AVERAGEIFS / INDEX-MATCH 等批量子表达式结果

解决的问题：

- 复合公式中相同子表达式重复求值

不足：

- 这是表达式级缓存，不是源数据级缓存

## 4. 今天实测得到的内存事实

基于今天对大样本文件的调查，确认了这些事实：

### 4.1 `rows * 50` 不是可靠的内存模型

当前 `batch.go` 中有：

```go
estimatedCells = len(currentWs.SheetData.Row) * 50
```

它只是 `cellMap` 预分配容量的经验值，不是内存估算模型。

在 `/Users/zhoujielun/Downloads/Shopee平台数据表.xlsx` 上实测：

- 总行数约 `640,930`
- 总单元格数约 `14,765,057`
- 平均每行约 `23.04` 个单元格
- `rows * 50` 比真实 cell 数高估约 `2.17x`

### 4.2 `WorksheetCache.LoadSheet()` 的每条目成本很高

对中小 sheet 的实际采样显示：

- 每个缓存条目大致增加 `450B ~ 490B` 堆内存

因此，如果对一个千万级单元格的大 sheet 做全量 source cache，额外内存很快会上 GB。

### 4.3 超大 sheet 实际峰值会远高于最终缓存体积

在大 sheet 探针运行时观察到：

- RSS 一度达到约 `12.4 GiB`

说明峰值不仅包括最终缓存，还包括：

- XML 字节
- worksheet 对象树
- `WorksheetCache`
- GC 未回收的中间对象
- 重算期间临时矩阵和索引

## 5. `excelize-mcp` 应用层现状

## 5.1 已经有全局 `FileCache`

主要位置：

- `excelize_tools/third_party/excelize/file_cache.go`
- `excelize_tools/third_party/excelize/datatable.go`

关键事实：

- `GetGlobalFileCache()` 返回全局单例缓存
- key 以 `documentID` 为主，并支持 alias 到 `fileID`
- cache 里缓存的是 `*excelize.File`
- 每个 `CachedFile` 维护：
  - `LoadedSheets`
  - 文件级读写锁
  - 脏标记
  - 估算字节数
  - LRU / memory watermark 管理

这意味着：

- 场景 1 在应用层已经“部分成立”
- 用户 A 打开文档后，用户 B 再访问同文档时，通常可以命中同一个全局 `CachedFile`
- 不需要再次从 COS / GridFS / temp file 完整打开 workbook

### 5.2 但当前共享的是“文件对象 / worksheet 对象”，不是“轻量源数据缓存”

`excelize-mcp` 现在复用的是：

- 同一个 `*excelize.File`
- 同一个内部 worksheet object cache

不是复用：

- block 化的源数据视图
- 面向重算的轻量 source store

这带来一个问题：

- 跨用户共享虽然成立
- 但共享层太重
- 大 sheet 一旦被 `LoadWorksheet()` 进内存，仍然可能非常吃内存

### 5.3 `ensureWorksheetLoaded()` 的行为

在 `datatable.go` 里，很多读写接口都会调用 `ensureWorksheetLoaded()`：

- 若 `excelize` 内部已加载，则同步 `LoadedSheets` 标记
- 否则执行 `f.LoadWorksheet(sheetName)`

这说明当前应用层策略是：

- 先缓存整个文件对象
- 再按 worksheet 懒加载完整 sheet

优点：

- 场景 1 的复用路径很直接
- 避免了重复打开文件

不足：

- 对超大 sheet，复用的是完整 worksheet 对象，成本仍高
- 还没有 “只为重算构建轻量源数据块” 的机制

### 5.4 Recalculate 路径

`excelize-mcp` 在多条路径上调用：

- `RecalculateAllWithDependency()`
- `CalcCellValuesDependencyAware()`
- `CalcCellValues()`

应用层有 `clear_cache` 参数，可强制保存后失效并重新加载文档。

优点：

- 提供了明确的缓存生命周期控制入口

不足：

- 现在 `clear_cache` 是文件级 / workbook 级
- 没有 sheet 级或 block 级源数据失效能力

## 6. 两个目标场景的真实状态

## 6.1 场景 1：用户 A 打开后，用户 B 直接复用

### 当前是否已经部分满足

是，`excelize-mcp` 的全局 `FileCache` 已经部分满足。

### 当前实际复用的内容

复用的是：

- `*excelize.File`
- 已加载的 worksheet 对象
- 文件内已有的 formula / range / index 缓存状态

### 当前存在的问题

- 复用对象过重
- 对大 sheet 来说，跨用户共享的是高内存对象
- 一旦多个热点大文档同时驻留，内存压力会很大

### 结论

场景 1 不是“没有缓存”，而是“缓存层次不对”。

## 6.2 场景 2：全量重算时，值应能快速索引而不是回 XML

### 当前是否已经部分满足

也是部分满足。

已有帮助的缓存包括：

- `WorksheetCache`
- `calcCache`
- `rangeCache`
- `matchIndexCache`
- `ifsMatchCache`
- `rangeIndexCache`
- `SubExpressionCache`

### 当前没有完全满足的地方

- 源数据读取层还不够稳定
- `WorksheetCache.LoadSheet()` 对大 sheet 不安全
- `PreloadColumnRange()` 仍先整表 `workSheetReader()`
- 因此“快速索引”常常建立在“先把大 sheet 全量加载进内存”之上

### 结论

场景 2 不是“没有加速”，而是“加速建立在高内存前提上”。

## 7. 推荐的总体优化方向

最合理的方向不是推翻现有全部缓存，而是分层重构职责。

## 7.1 保留现有缓存，但重新定义职责

### 保留的层

- `FileCache`：继续负责文档级跨请求复用
- `Pkg` / `Sheet` / `SharedStrings`：继续负责 `excelize` 内部对象与 XML 复用
- `calcCache`：继续负责公式结果缓存
- `rangeCache`：继续负责范围矩阵缓存
- `matchIndexCache` / `ifsMatchCache` / `rangeIndexCache`：继续负责 lookup 与条件匹配索引
- `SubExpressionCache`：继续负责 batch 子表达式复用

### 需要收缩职责的层

- `WorksheetCache`

建议改成：

- 主要承载“重算期间的最新公式值 overlay”
- 不再承担“大 sheet 原始数据全量缓存”职责

## 7.2 新增一层：Sheet Source Store

这是今天调研后最明确缺失的一层。

建议引入：

- `SheetSourceStore`

建议 key：

- `(documentID or workbookVersion, sheetName, sourceVersion)`

建议内容：

- block 化的源数据存储
- 以 row block 或 row+column block 为基本单元
- 不用 `map[string]formulaArg` 做全量源值表示
- 尽量用数字坐标 key / 紧凑值表示

建议职责：

- 首次访问热点大 sheet 时，顺序扫描 XML 一次
- 构建本次重算可复用的 source blocks
- 后续范围读取、lookup、批量计算都从 source store 读

这样可以真正满足场景 2。

## 7.3 让 `FileCache` 复用的是“文件 + source store 元数据”

结合 `excelize-mcp` 最自然的落点是：

- `CachedFile` 继续持有 `*excelize.File`
- 额外持有 `sheetSourceStores map[string]*SheetSourceStore`

这样：

- 用户 A 打开并预热 `ERP` sheet
- 用户 B 再访问时，不只复用 `*excelize.File`
- 还可以直接复用 `ERP` 的 source store

这才是场景 1 的理想版本。

## 7.4 Source Store 生命周期

建议生命周期如下：

### 单次重算内

- 某张热点 sheet 第一次被访问
- 若 store 不存在，则构建一次
- 后续本轮重算全部复用

### 多请求 / 多用户间

- 若文档版本未变，则 `CachedFile` 中的 source store 继续有效
- 若文档变更，则只失效相关 sheet 或 block
- 若文档整体被替换，则整个 workbook source store 失效

## 7.5 增量失效，而不是整表清空

未来应支持：

- `update_range`
- `update_range_by_lookup`
- 行列插入删除
- sheet rename / move

对应的 source store 处理策略：

- 标记受影响 row block dirty
- 标记受影响 column block dirty
- 下次访问时懒重建

不要默认：

- 整个文档 clear cache
- 整个 sheet 重新全量 load

## 8. 推荐的缓存分层方案

建议最终形成 4 层：

### Layer A: Document Cache

位置：`excelize-mcp`

载体：`FileCache`

缓存内容：

- `*excelize.File`
- 文档级元数据
- 每个 sheet 的 source store 引用

解决问题：

- 跨请求 / 跨用户复用
- 避免反复从存储层重新打开文档

### Layer B: Worksheet Object Cache

位置：`excelize`

载体：`Pkg` / `Sheet` / `SharedStrings`

缓存内容：

- worksheet XML bytes
- `xlsxWorksheet`
- shared strings

解决问题：

- 普通 API 兼容
- 小 sheet 对象复用
- 写回操作继续依赖完整 worksheet 模型

### Layer C: Source Data Cache

新层

载体：`SheetSourceStore`

缓存内容：

- 大 sheet 的 block 化源数据
- 面向 lookup / range / recalculation 的紧凑值表示

解决问题：

- 大 sheet 重算时不再频繁打 XML
- 为场景 1 和场景 2 提供共同底座

### Layer D: Recalc Overlay And Result Cache

载体：

- `WorksheetCache`
- `calcCache`
- `rangeCache`
- `matchIndexCache`
- `ifsMatchCache`
- `rangeIndexCache`
- `SubExpressionCache`

缓存内容：

- 最新公式结果
- 范围矩阵
- lookup 索引
- batch 子表达式

解决问题：

- 重算期间结果复用
- 大量 lookup / range / criteria 快速命中

## 9. 方案优先级建议

### 第一阶段：低风险重构

1. 明确 `WorksheetCache` 只做重算 overlay，不再鼓励大 sheet `LoadSheet()`。
2. 给大 sheet 增加 hard limit，避免误走整表 source cache。
3. 在 `excelize-mcp` 的 `CachedFile` 中预留 `sheetSourceStores` 结构。
4. 先把热点大 sheet 的“是否已预热 source store”纳入 cache info 观测面板。

### 第二阶段：真正可落地的性能优化

1. 新增基于流式读取的 `SheetSourceStore` 构建逻辑。
2. `PreloadColumnRange` / lookup 批量优化优先从 source store 取值。
3. `CalcCellValuesDependencyAware` / `RecalculateAllWithDependency` 优先读 source store。
4. `WorksheetCache` 只保留计算结果覆盖层。

### 第三阶段：完整多用户共享模型

1. `CachedFile` 复用 source store。
2. 文档保存后为 source store 打版本号。
3. `update_range` / `delete rows` / `insert columns` 等操作支持 block 级 dirty。
4. memory pressure 下可对冷 source block 做淘汰或落临时文件。

## 10. 风险点

### 10.1 如果新增 source store 但不收缩旧缓存，内存可能更糟

必须避免：

- 大 sheet 同时保留完整 worksheet object + 全量 `WorksheetCache` + 全量 source store

### 10.2 缓存一致性会变复杂

需要统一管理以下层的失效：

- `FileCache`
- worksheet object cache
- source store
- `WorksheetCache`
- `calcCache`
- `rangeCache`
- 各类 index cache

### 10.3 写路径仍需要完整 worksheet 对象

因此 source store 不能完全替代 worksheet object cache。

- 写操作、导出、结构调整仍需要 `xlsxWorksheet`
- source store 应是“读优化层”，不是写模型替代层

### 10.4 跨用户共享必须带版本控制

如果场景 1 要彻底成立，必须把 shared source store 与文档版本绑定。

否则会出现：

- A 修改后 B 读到旧 source store
- 或者不同版本 source store 混用

## 11. 最终总结

今天的调研已经说明：

1. `excelize` 侧已经有很多缓存，但主要是对象缓存、结果缓存、范围缓存、索引缓存。
2. `excelize-mcp` 侧已经有全局 `FileCache`，所以“用户 A 打开后用户 B 复用”其实已经部分成立。
3. 当前问题不在“完全没有缓存”，而在“缺少一层专门面向大 sheet 源数据的轻量可复用缓存”。
4. `WorksheetCache` 目前最适合做“重算 overlay cache”，不适合继续当“大 sheet 全量源数据 cache”。
5. 最合适的优化方向是新增 `SheetSourceStore`，并把它挂到 `excelize-mcp` 的 `CachedFile` 上，让：
   - 单次重算内可以复用
   - 多用户请求间也可以复用
   - 文档更新时支持 block 级失效

一句话总结：

- 现有缓存体系不用推翻
- 但必须补上一层“大 sheet source cache”
- 并把 `WorksheetCache` 从“源数据 cache”收缩为“重算结果 overlay”
- 这才是同时满足内存安全、全量重算性能、以及多用户共享场景的正确方向
