# Excel-VBA-Libraries — 项目规格文档

> 版本：v1.0 | 最后更新：2026-07-23 | 状态：功能完备，待重构

## 1. 项目概述

**Excel-VBA-Libraries** 是一个纯 VBA 实现的 Excel 函数增强库，提供 15 个功能模块、530+ 个 Public 函数，覆盖统计、线性代数、回归、物理化学、字符串、日期时间、正则、数组、字典、JSON、XML、透视表、SQL、Range 操作、文件系统等。

### 核心价值

- 零依赖：纯 VBA 实现，无需安装任何运行时或插件
- 双路径：所有 Public 函数同时支持 Range 对象和 Variant 数组输入
- 交叉验证：Python COM 自动化测试，numpy/scipy 独立计算期望值
- 完整文档：中英双语用户手册 + API 参考

### 目标用户

- 无法安装 XLL 插件的企业环境（IT 管控严格）
- 需要 VBA 原生解决方案的 Excel 重度用户
- 需要跨版本兼容（Excel 2010+）的开发者

## 2. 功能规格

### 2.1 模块清单

| 模块 | 文件 | 说明 |
|------|------|------|
| StatsUtils | StatsUtils.bas | 均值/方差/分位数/相关/ANOVA/排名 |
| LinearUtils | LinearUtils.bas | 行列式/求逆/特征值/矩阵运算 |
| RegressUtils | RegressUtils.bas | OLS/WLS/岭回归/R² |
| PhyChemUtils | PhyChemUtils.bas | 分子量/温度/压力/气体定律/单位换算 |
| StringUtils | StringUtils.bas | 反转/提取/编解码/编辑距离/Soundex |
| DateTimeUtils | DateTimeUtils.bas | ISO周/工作日/年龄/时间戳 |
| RegexUtils | RegexUtils.bas | 匹配/替换/计数（VBScript.RegExp） |
| ArrayUtils | ArrayUtils.bas | 排序/筛选/去重/切片/打乱 |
| DictSetUtils | DictSetUtils.bas | 频率/交集/并集/TopN |
| JsonUtils | JsonUtils.bas | JSON 解析/查询（纯 VBA 递归下降） |
| XmlUtils | XmlUtils.bas | XML 解析（MSXML2） |
| PivotUtils | PivotUtils.bas | 透视/逆透视/分组聚合/交叉连接 |
| SqlUtils | SqlUtils.bas | ADODB 对 Excel 区域写 SQL |
| RangeUtils | RangeUtils.bas | Range 导出(HTML/JSON/MD)/区域操作/命名 |
| FileSystemUtils | FileSystemUtils.bas | 读写文件/列目录 |

### 2.2 核心基础设施（VBA-Core）

| 组件 | 文件 | 职责 |
|------|------|------|
| VariantKit | VariantKit.cls | 类型转换 + NormalizeInput 双路径入口 |
| ArrayOps | ArrayOps.cls | 数组基础操作（Sort/Slice/IndexOf/Flatten） |
| DictProxy | DictProxy.cls | Scripting.Dictionary 封装（Late Binding） |

**导入顺序**：VariantKit → ArrayOps → DictProxy → 其余模块

### 2.3 关键技术特性

- **双路径原则**：Public 函数必须同时处理 Range 对象和 Variant 数组
- **As Variant 参数**：所有 UDF 参数声明为 `As Variant`（否则 Range 传入 #VALUE!）
- **Err.Raise 错误处理**：UDF→CVErr，VBA 函数→Err.Raise，资源→Cleanup 标签
- **1-based 数组**：VBA 默认 Option Base 0，但 NormalizeTo2D 统一返回 1-based

## 3. 架构规格

### 3.1 模块依赖

```
VBA-Core (VariantKit → ArrayOps → DictProxy)
    ↑
    ├── StatsUtils ← LinearUtils, RegressUtils 依赖
    ├── LinearUtils
    ├── RegressUtils ← 依赖 LinearUtils, StatsUtils
    ├── PhyChemUtils
    ├── StringUtils
    ├── DateTimeUtils
    ├── RegexUtils
    ├── ArrayUtils
    ├── DictSetUtils
    ├── JsonUtils
    ├── XmlUtils
    ├── PivotUtils
    ├── SqlUtils
    └── FileSystemUtils
```

### 3.2 测试体系

| 层级 | 工具 | 覆盖 |
|------|------|------|
| 一致性验证 | Python 脚本（无需 Excel） | 签名/交叉引用/元数据/代码质量 |
| 交叉验证 | Python COM 自动化 | 18 模块，numpy/scipy 独立计算 |
| 集成测试 | Python COM + 真实数据 | 15 模块，Range 路径 |
| 单元测试 | VBA Test_* | 仅 SqlUtils（ADODB 无法 COM 测试） |

## 4. 质量规格

### 4.1 已知限制

- VBA 无短路求值：`And`/`Or` 两侧均执行，需显式嵌套 If
- ReDim(1 To 0) 崩溃：空数组必须用 `Erase` 或 `Dim` 未初始化
- COM 数组 0-based：Python 传入的数组需 NormalizeTo2D 转 1-based
- VBScript.RegExp 不支持 Unicode 类别：中文匹配需特殊处理
- ADODB SQL 无法 COM 测试：需手动在 Excel 中运行

### 4.2 安全规格

- 文件系统操作无沙箱（VBA 限制）
- SQL 通过 ADODB，存在注入风险（需参数化）
- 正则无超时保护（VBScript.RegExp 限制）

## 5. 历史演化摘要

| 阶段 | commits | 关键事件 |
|------|---------|----------|
| 初始实现 | ~50 | 14 模块 + VBA-Core + 基础测试 |
| 审查修复期 | ~100 | 91 findings 全量审计，Err.Raise 拼写 8 次 |
| 交叉验证期 | ~80 | Python COM 基础设施，双路径覆盖 |
| 文档完善期 | ~50 | 中英双语手册，API 参考，锚点校验 |

**总计**：283 commits，是 5 个项目中 commit 数最多的

## 6. 关键设计决策

| 决策 | 选择 | 理由 |
|------|------|------|
| 纯 VBA vs XLL | 纯 VBA | 零依赖，企业环境兼容 |
| Late Binding vs Early Binding | Late Binding | 无需引用设置，跨版本兼容 |
| Scripting.Dictionary vs Collection | Dictionary | 键值查找 O(1)，Collection 仅 O(N) |
| VBScript.RegExp vs 自实现 | VBScript.RegExp | 性能更好，功能完整 |
| MSXML2 vs 纯 VBA XML | MSXML2 | 性能更好，XPath 支持 |
