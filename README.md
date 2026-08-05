# Excel-VBA-Libraries

> 高性能 VBA 函数库：15 个模块，纯 VBA 实现，零外部依赖，兼容 Excel 2010+。

---

## 安装

1. 从 [Releases](https://github.com/zgrwo/Excel-VBA-Libraries/releases) 下载源码
2. 导入 VBA-Core 类模块：`VariantKit.cls` → `ArrayOps.cls` → `DictProxy.cls`（顺序固定）
3. 导入所需的标准模块（.bas）
4. 在 VBA 编辑器中：工具 → 引用 → 勾选 `Microsoft Scripting Runtime`

### 验证安装

在 VBA 立即窗口中运行：
```vba
? Mean(Array(1,2,3,4,5))
' → 3
```

---

## 模块速览

> 完整签名、参数说明见 **[API 参考](rules/api-reference.md)**；每个函数的详细示例见 **[用户手册](rules/user-manual.md)**。

| 模块 | 做什么 | 试一试 |
|------|------|-------|
| `StatsUtils` | 描述统计（均值/方差/分位数/相关/t检验） | `=Mean(A1:A100)` |
| `LinearUtils` | 线性代数（SVD/伪逆/QR/矩阵乘法） | `=UDF_LINALG_DET(A1:D4)` |
| `RegressUtils` | 回归分析（OLS/WLS/岭回归） | `=FitOLS(y_range, x_range)` |
| `PhyChemUtils` | 物理化学（分子量/温度/压力/理想气体） | `=MolecularWeight("H2SO4")` |
| `StringUtils` | 字符串处理（格式化/反转/编码/截断） | `=ReverseString("hello")` |
| `DateTimeUtils` | 日期时间（周/月/季度/工作日） | `=LastDayOfMonth(A1)` |
| `RegexUtils` | 正则表达式（匹配/替换/拆分/捕获组） | `=RegexIsMatch(A1, "\d+")` |
| `ArrayUtils` | 数组操作（排序/去重/过滤/切片） | `=ArrayUnique(B2:B100)` |
| `DictSetUtils` | 字典集合（合并/交集/并集/差集） | `=DictMerge(dict1, dict2)` |
| `JsonUtils` | JSON 处理（解析/查询/验证/美化） | `=JsonParse(A1)` |
| `XmlUtils` | XML 处理（XPath 查询/验证/转表） | `=XmlValidate(A1)` |
| `PivotUtils` | 数据透视（透视/逆透视/分组/交叉连接） | `=GroupBy(A1:D100, 1, 2)` |
| `FileSystemUtils` | 文件系统（FSO 路径/读写/复制/删除） | `=FileExists("C:\Data\test.csv")` |
| `SqlUtils` | SQL 查询（ADODB 单表/双表/三表） | `=UDF_SQL_QUERY("SELECT * FROM data")` |
| `RangeUtils` | 范围导出（HTML/JSON/MD/CSV/转置） | `=LastRow(A1:D10)` |

---

## 使用模式

### 工作表公式

```vb
' 直接在工作表单元格中使用
=ReverseString("hello")
→ "olleh"

' 数组公式（Excel 365 自动溢出）
=ArrayUnique(A1:A100)
```

### VBA 代码调用

```vba
Dim result As Variant
result = ReverseString("hello")
' → "olleh"

' 双路径支持：数组和 Range 均可
result = Mean(Range("A1:A10"))     ' Range 路径
result = Mean(Array(1,2,3,4,5))   ' 数组路径
```

---

## 错误处理

三个错误处理模式：

| 模式 | 适用场景 | 示例 |
|------|---------|------|
| `CVErr(xlErrValue)` | UDF 函数入口 | 参数非法 → `#VALUE!` |
| `Err.Raise` | VBA 内部函数 | 计算异常 → 上层 `On Error GoTo` 捕获 |
| `On Error GoTo Cleanup` | 资源操作 | 文件/数据库 → 确保资源释放 |

> 🔴 禁止裸 `On Error Resume Next` 不检查 `Err.Number`。

---

## 安全

- **完全本地**：纯 VBA 实现，无外部依赖，无网络调用
- **SQL 参数化**：SqlUtils 使用 ADODB 参数化查询
- **Office 2010+**：仅使用 Excel 2010 及以上稳定 COM API

---

## 质量保证

- **双路径测试**：数组路径（Python COM 交叉验证）+ Range 路径（集成测试）
- **交叉验证**：numpy/scipy 独立实现与 VBA 逐项比对
- **四层测试体系**：一致性验证 → 交叉验证 → 集成测试 → 单元测试（SqlUtils）
- **6 处同步**：新增函数强制同步源码/API Index/CN手册/EN手册/交叉验证/锚点

---

## 已知限制

- **VBA 语言**：微软已停止 VBA 语言演进，推荐 Office Scripts（TypeScript）为未来路径
- **Excel COM**：交叉验证和集成测试需安装 Excel（非跨平台）
- **模块依赖**：RegressUtils 依赖 LinearUtils + StatsUtils（其他 13 个模块相互独立）
- **VBA-Core 接口**：接口已冻结，不可随意修改

---

## 长期迁移路径

保持 15 个模块相互独立 = 保持未来逐模块迁移 Office Scripts 的可能性：
- 每个模块可独立测试（不依赖其他模块运行）
- 模块间通过明确接口通信（不共享内部状态）
- 迁移时逐模块替换，而非一次性重写

---

## 贡献

请阅读 [CONTRIBUTING.md](CONTRIBUTING.md) 了解贡献流程（fork → PR → review）。

---

## 许可证

[MIT](LICENSE) © zgrwo

---

## 架构特点

```
VBA-Core (类模块)
  VariantKit → ArrayOps → DictProxy（导入时按此顺序）
      ↑ 依赖
src/ (标准模块)
  15 个 .bas 模块，相互独立（均依赖 VBA-Core）
```

- ✅ 纯 VBA 实现，零外部依赖（仅 Windows 内置 COM 接口）
- ✅ 所有 Public 函数双路径处理（Range 对象 + Variant 数组）
- ✅ 14/15 模块真正独立，支持逐模块迁移至 Office Scripts
- ❌ VBA-Core 接口冻结，禁止修改
- ❌ 禁止裸 `On Error Resume Next` 不检查 Err

---

## 测试命令

```bash
# 快速验证（无需 Excel）
python tests/run_all_validation.py --quick

# 全量验证
python tests/run_all_validation.py

# 交叉验证（需 Excel）
python tests/run_all_crossval.py

# 集成测试（需 Excel）
python tests/utils/integration_test_all_modules.py
```

---

## 文档索引

| 文档 | 角色 | 内容 |
|------|------|------|
| [API 参考](rules/api-reference.md) | 签名唯一信源 | 15 个模块函数签名、参数说明 |
| [用户手册](rules/user-manual.md) | 学习教程 | 每个函数详细示例 + 结果解读 |
| [context.md](rules/context.md) | 术语表 | 所有领域术语唯一定义 |
| [project-structure.md](rules/project-structure.md) | 结构地图 | 文件职责与层级关系 |
| [AGENTS.md](AGENTS.md) | 项目宪法 | 架构分层、红线规则、开发流程 |

---

## 治理体系说明

本项目遵循 [Harmonization 治理规范](https://github.com/zgrwo/Harmonization) 模板体系：

| 文件 | 面向 | 职责 |
|------|------|------|
| `AGENTS.md` | AI 编程助手 | 项目宪法——架构、红线、编码准则、防幻觉铁律 |
| `readme.md` | 人类用户 | 功能指南——安装、模块速览、使用模式（本文件） |
| `rules/` | AI + 人类 | 规范文档——API 参考、用户手册、术语表、审查模板 |
| `skills/` | AI 编码 | 技能定义——语言陷阱、编码模式、重构守则 |

**核心原则**：SSOT（信息只在一处定义）、Skill-first（修改代码前加载技能）、四条核心准则。

<!-- last_updated: 2026-07-24 -->
