# @Description 注解规范 (Annotation Guide)

> 为 Rubberduck VBA 兼容做准备。注解以 VBA 注释形式存在，不影响运行时行为。

## 基本注解

### @Description

为 Public 函数添加描述，Rubberduck 可在 IntelliSense 中显示：

```vba
'@Description("计算数组的算术平均值")
Public Function StatsUtils_Mean(data As Variant) As Variant
```

### @Folder

标注模块所属逻辑文件夹（Rubberduck 用于组织代码浏览器）：

```vba
'@Folder("Statistics")
Option Explicit
```

推荐文件夹结构：

| 文件夹 | 模块 |
|--------|------|
| `Data` | ArrayUtils, DictSetUtils, PivotUtils, SqlUtils |
| `Statistics` | LinearUtils, StatsUtils, RegressUtils |
| `Text` | StringUtils, RegexUtils, JsonUtils, XmlUtils |
| `DateTime` | DateTimeUtils |
| `Excel` | RangeUtils, FileSystemUtils |
| `Science` | PhyChemUtils |
| `Core` | VariantKit, ArrayOps, DictProxy |

### @Ignore / @IgnoreModule

抑制特定检查（仅在确认误报时使用）：

```vba
'@IgnoreModule EmptyStringHandling
' 或针对单行:
result = "" '@Ignore UseMeaningfulName
```

## 参数注解

### @Param

显式描述参数含义：

```vba
'@Description("对数组进行升序排序")
'@Param("data - 输入数组或 Range")
'@Param("stable - 是否使用稳定排序")
Public Function ArrayUtils_SortAsc(data As Variant, Optional stable As Variant) As Variant
```

### @Returns

描述返回值：

```vba
'@Returns 排序后的 1-based Variant 数组
```

## 迁移策略

### 当前状态

注解为**可选推荐**，不作为硬要求。原因：
- 项目当前不依赖 Rubberduck 运行
- 200+ 函数全量添加注解工作量大
- 注解不影响现有测试体系

### 新增函数要求

所有**新增** Public 函数必须包含：
1. `'@Description("...")` — 一句话描述
2. `'@Folder("...")` — 模块级（文件头，已有则跳过）

### 存量迁移优先级

1. **高频使用**：StatsUtils, ArrayUtils, StringUtils（用户最常接触）
2. **复杂参数**：RegressUtils, LinearUtils（参数含义不直观）
3. **其余模块**：按需添加

## 与现有文档的关系

| 信源 | 角色 | 注解的作用 |
|------|------|-----------|
| `rules/api-reference.md` | 签名唯一信源 | 注解不替代，仅补充 |
| `docs/VBA_LIB_User_Manual_EN.md` | 用户教程 | 注解提供 IDE 内快速提示 |
| `scripts/generate_api_docs.py` | 自动提取签名 | 未来可同时提取注解 |

> 原则：注解是**辅助**，不是**信源**。文档冲突时以 `rules/api-reference.md` 为准。
