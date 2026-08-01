# context.md — Excel-VBA-Libraries 术语表

> 领域词汇的精确定义。AI 在对话中应使用这些术语，避免近义词。
> 链接索引见 [agents.md](../agents.md#参考)。

## 架构术语

**VBA-Core** — 类模块基础层，提供共享数据结构与操作。包含 VariantKit（类型安全包装）、ArrayOps（数组操作）、DictProxy（字典代理）。所有 src/ 模块依赖 VBA-Core，导入顺序固定：VariantKit → ArrayOps → DictProxy。
_Avoid_: 核心层、基础库、Utils

**src/** — 15 个独立标准模块（.bas），每个模块功能内聚，模块间相互独立（仅 RegressUtils 例外：依赖 LinearUtils + StatsUtils）。
_Avoid_: 函数库、模块集

**VariantKit** — VBA-Core 的 Variant 类型安全包装类。提供类型检测（IsNumber/IsText/IsBool/IsDate）、类型转换（ToDouble/ToLong/ToString）、哨兵处理（Null/Empty/Error 统一转对应零值）。
_Avoid_: 类型工具、VariantHelper

**ArrayOps** — VBA-Core 的数组操作类。提供一维/二维数组创建、重塑、转置、展平、排序、过滤、切片。所有操作基于 LBound/UBound 而非硬编码索引。
_Avoid_: 数组工具、ArrayHelper

**DictProxy** — VBA-Core 的字典代理类。封装 Scripting.Dictionary，提供类型安全的键值操作、频率统计、集合运算（交/并/差集）。
_Avoid_: 字典工具、DictHelper

## 数据类型术语

**Variant** — VBA 可变类型。🔴 所有 Public UDF 参数必须声明为 `As Variant`，否则 Range 对象传入时触发 `#VALUE!`。
_Avoid_: 可变类型、动态类型

**Range 对象** — Excel 单元格区域引用。UDF 必须同时处理 Range 对象和 Variant 数组两种输入路径（双路径原则）。
_Avoid_: 区域、单元格范围

**Variant 数组** — 从 Range 或 Python COM 传入的内存数组。通过 `IsArray()` 检测，使用 LBound/UBound 遍历。
_Avoid_: 动态数组、VBA 数组

**哨兵值 (Sentinel)** — 不可转换值的类型零值替代：数值→0、字符串→""、日期→#1/1/1900#、布尔→False。不抛异常，静默返回。
_Avoid_: 默认值、占位值

## 错误处理术语

**CVErr** — 用户定义函数向 Excel 返回错误的机制：`CVErr(xlErrValue)` = `#VALUE!`、`CVErr(xlErrNum)` = `#NUM!`。仅用于 UDF 入口。
_Avoid_: 错误返回、UDF 错误

**Err.Raise** — VBA 内部函数抛异常的标准方式。用于非 UDF 的内部函数，调用方通过 `On Error GoTo` 捕获。
_Avoid_: 抛出错误、错误抛出

**On Error GoTo Cleanup** — VBA 结构化错误处理模式：`On Error GoTo Cleanup` → 正常执行 → `Cleanup:` 标签 → 资源释放。用于文件/数据库等资源操作。
_Avoid_: 错误跳转、异常处理

**裸 On Error Resume Next** — 🔴 反模式：`On Error Resume Next` 后不检查 `Err.Number`，静默吞没所有错误。代码审查必检项。
_Avoid_: 忽略错误、跳过异常

## 测试术语

**双路径原则 (Dual-Path)** — UDF 测试必须覆盖两种输入路径：① 数组路径（Python list → COM Variant 数组，由交叉验证覆盖）；② Range 路径（COM Range → VBA Variant，由集成测试覆盖）。两者都必须通过。
_Avoid_: 双通道、双模式测试

**一致性验证 (Consistency Validation)** — `run_all_validation.py` 执行的离线检查：签名一致性、交叉引用完整性、元数据正确、代码质量扫描。无需 Excel，可 CI 执行。
_Avoid_: 静态检查、代码验证

**交叉验证 (Cross-Validation)** — Python（numpy/scipy）独立计算期望值，与 VBA 实际输出逐项比对。覆盖数组路径，需 Excel COM 调用。
_Avoid_: 数值比对、参考验证

**集成测试 (Integration Test)** — COM Range → VBA Variant → 返回值的端到端测试。覆盖 Range 路径，需 Excel 运行。
_Avoid_: E2E 测试、完整测试

## 同步术语

**6 处同步清单** — 新增 Public 函数必须同步更新的 6 个位置：① 源码实现（.bas）② API Index（签名行 + 计数头）③ CN 手册（中文示例）④ EN 手册（英文示例）⑤ 交叉验证（Python 独立实现比对）⑥ 锚点校验（文档内链有效）。缺一不可。
_Avoid_: 注册清单、同步步骤

**API Index** — api-reference.md 中的函数索引，包含每个函数的签名行和模块头部的函数计数。计数头必须与实际函数数量一致。
_Avoid_: 函数列表、签名索引

**锚点校验** — 验证文档内链（`<a id="...">`）与函数索引中的链接是否一一对应，无死链。
_Avoid_: 链接检查、引用验证

## 开发术语

**VBA-Core 接口冻结** — VBA-Core 类模块（VariantKit/ArrayOps/DictProxy）的 Public 接口禁止修改，除非用户明确要求。15 个 src/ 模块全部依赖这些接口，修改会造成连锁回归。
_Avoid_: 接口锁定、核心层保护

**模块独立性** — 15 个 src/ 模块相互独立设计原则：每个模块可独立测试、独立替换、独立迁移。不共享内部状态，仅通过明确接口通信。这是未来 VBA→Office Scripts 逐模块迁移的基础。
_Avoid_: 模块解耦、独立部署

**导入顺序** — VBA-Core 类模块的固定导入顺序：VariantKit → ArrayOps → DictProxy。ArrayOps 依赖 VariantKit，DictProxy 依赖两者。违反顺序导致编译错误。
_Avoid_: 加载顺序、依赖顺序

## 平台术语

**Excel 2010+ 兼容** — 纯 VBA 实现，零外部依赖（无 DLL、无 .NET），仅使用 Excel 2010 及以上版本支持的 VBA 内置函数和 COM 接口。
_Avoid_: 向后兼容、低版本支持

**Office Scripts 迁移路径** — VBA 语言已停止演进，微软推荐 Office Scripts（TypeScript）作为未来方向。保持 13 个模块独立性 = 保持逐模块迁移的可能性，而非一次性全量重写。
_Avoid_: 迁移策略、语言升级
