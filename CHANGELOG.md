# Changelog

All notable changes to Excel VBA Libraries.

## [2.1.1] — 2026-08-16

### Fixed

- **FileSystemUtils**: 补定义 `ERR_INVALID_INPUT` 常量 — 修复 `ReadBinaryFile` 因未定义标识符导致的过程级编译失败（第二轮发行前审查 FS-01，Critical）；`ValidateSafePath` 新增 UNC 路径拦截（`\\` 与 `//` 两种写法）并在手册声明静默失败语义与符号链接限制（FS-02）
- **JsonUtils**: JSON 对象键改为大小写敏感（RFC 8259，修复 `{"a":1,"A":2}` 静默丢键）；`JsonGet` 数组索引严格整数校验 + 越界检查（拒绝银行家舍入）
- **LinearUtils**: LU 主元容差改为纯相对（修复 `MatrixDeterminant(diag(1e-20))` 静默返回 0）；Cholesky 对齐 LAPACK dpotrf 精确判定（修复高条件数 SPD 矩阵如 `diag(1e10,1e-8)` 被误拒）；12 个矩阵函数支持 Range 输入归一化；补齐 9 处 Err.Raise Source；经济模式 QR 宽矩阵语义文档化
- **VBA-Core**: `VariantKit.Compare` 数组分支修复单行 If 冒号陷阱（比较器不对称）；`DictProxy.Merge` 保留源字典 CompareMode（修复 vbBinaryCompare 字典合并 Error 457）
- **StatsUtils**: `GammaLn`/`BetaReg` 新增域校验（极点/非法参数）；Correlation/RSquare/ZScore/Normalize 除零防护改为尺度感知容差（修复 1e-14 尺度数据误报零方差）；HarmonicMean 补 Kahan 补偿；QuickSortDouble 改三数取中选轴；删除死代码 NormSInv
- **RegressUtils**: ANOVAOneWay 移除死计算 grandSumSq
- **RangeUtils**: ExportRangeToCSV 修复双 BOM（删除手工 BOM，依赖 ADODB.Stream 自动 BOM；bom=False 经二进制复制真正生效）
- **SqlUtils**: SqlRangeQuery 新增多区域 Range 前置检查（置于错误 handler 之后，保持 outOk=False 软失败契约）
- **PhyChemUtils**: 修复多连接符公式系数嵌套（`A·2B·3C` = A+2×B+3×C）；MassToMoles/MolesToMass/Density 保持原参数名（兼容命名参数调用方），单位约定转入注释与手册
- **ArrayUtils**: 删除死代码 CompareAtIndices
- **StringUtils**: 35 个映射类 UDF 支持 1D 数组输入（不再拒绝合法输入）

### Changed

- api-reference.md 计数头修正为 533 Functions / 33 Subs / 566 总计；`generate_counts.py --check` 新增失配门禁并接入 pre_commit_check
- 手册（CN/EN）: JSON 键大小写敏感、SqlEscapeString Error 值行为、UDF_RANGE_FILTER/UDF_PIVOT_CROSSJOIN 已知限制声明、FileSystemUtils 路径安全限制与静默失败语义声明、PhyChem 单位约定说明
- sql-SKILL.md: SqlRangeQuery 实现描述修正（内存 Recordset，非临时工作簿）；EscapeSheetName 示例同步实现
- build_common.py: 新增 `expect_error` 用例类型（断言 VBA 调用必须抛错）；FS 负例实际采用 VBA 探针内部捕获模式（错误不穿透 COM 边界，不触发 Excel 运行时错误弹窗）

### Added

- 交叉验证回归用例: MatrixDeterminant_small_scale、UDF_LINALG_CHOLESKY_high_cond、JsonGetKeys_case_sensitive、ZScore_small_scale、ReadBinaryFile_roundtrip、ReadBinaryFile_unsafe_path（负例）、WriteBinaryFile_roundtrip（补齐零覆盖缺口）

## [2.1.0] — 2026-08-01

### Fixed

- **StatsUtils**: TTest Case 2/3 添加 se=0 除零守卫；ZTest 添加 sigma≤0 校验；HarmonicMean 拒绝负值
- **StatsUtils**: GeometricMean/TrimMean/MeanAbsDev 添加 Kahan 补偿求和
- **StatsUtils**: RankAvg ties 判定改为相对容差
- **SqlUtils**: SqlEscapeString 补充 `]` LIKE 转义；EscapeSheetName 验证方括号配对
- **SqlUtils**: SqlJoin joinType 白名单校验；连接字符串路径追加双引号拦截

### Changed

- api-reference.md 计数同步（33 Subs / 565 总计）
- README 验证示例修正为实际函数名 `Mean`

## [2.0.0] — 2026-06-30

### Core Infrastructure

- **VBA-Core**: `VariantKit.cls` (type normalization), `ArrayOps.cls` (array operations), `DictProxy.cls` (safe dictionary + batch ops)
- **15 source modules** across 6 domains: Data, Stats/Math, Text, Date, Excel/File, PhysChem
- **230+ UDFs** — all parameters `As Variant` for Range compatibility
- **Dual-path architecture**: Array path (VBA functions) + Range path (COM integration)
- **Dual-language documentation**: API Index + User Manuals in Chinese and English

### Modules

| Module | Domain | Highlights |
|--------|--------|------------|
| `ArrayUtils.bas` | Data | Sort, filter, slice, aggregate, search |
| `DictSetUtils.bas` | Data | Dictionary/set merge, intersection, difference, frequency |
| `PivotUtils.bas` | Data | Pivot, unpivot, group-by, cross-join |
| `SqlUtils.bas` | Data | SELECT, JOIN, GROUP BY via ADODB |
| `LinearUtils.bas` | Math | SVD, QR, LU, Cholesky, pseudoinverse |
| `StatsUtils.bas` | Math | Descriptive, inference, distributions, correlation |
| `RegressUtils.bas` | Math | OLS, ANOVA, factor importance, optimization |
| `StringUtils.bas` | Text | Encoding, validation, distance, UUID, URL |
| `RegexUtils.bas` | Text | Match, replace, split, capture groups |
| `JsonUtils.bas` | Text | Pure VBA recursive-descent JSON parser |
| `XmlUtils.bas` | Text | MSXML2 XPath query + worksheet export |
| `DateTimeUtils.bas` | Date | ISO week, workday, age, Easter |
| `RangeUtils.bas` | Excel | Export (HTML/JSON/MD), region ops, naming |
| `FileSystemUtils.bas` | File | UTF-8 read/write, folder traversal, drives |
| `PhyChemUtils.bas` | Science | Molecular weight, unit conversion, gas standard state |

### Testing

- **4-layer test suite**: Signature validation → Cross-validation (numpy/scipy) → COM integration tests → VBA unit tests
- **1,069 VBA assertions + 568 crossval cases + 214 manual examples**
- **Benchmark infrastructure** for performance regression tracking
- **Pre-commit hooks**: Structure validation + test runner

### Key Fixes (since initial release candidate)

- Range→Array normalization centralized in `VariantKit.NormalizeInput`
- `#VALUE!` elimination: all Public params `As Variant`, Range inputs handled
- Numerical stability: Kahan summation, QR vs normal equations, two-pass variance
- Currency precision in `DictKey`
- `ArraySort` negative number ordering
- `MolecularWeight` Ca(OH)₂ parsing
- `JsonUtils` parse null/literal safe content extraction
- `DateDiffParts` month range + key names
- Cross-validation tolerances tightened across all modules
- Integration tests: 0 failures across 82 quantified assertions

### Documentation

- API Index with signatures for all 230+ Public functions
- User Manual (CN) ~168 KB, User Manual (EN) ~178 KB
- VBA coding standards (`skills/vba-SKILL.md`, 17 sections)
- Python test authoring guide (`skills/python-SKILL.md`)
- SQL ADODB guide (`skills/sql-SKILL.md`)
- Manual authoring workflow (`skills/vba-manual-authoring.md`)

---

## Versioning

This project follows [Semantic Versioning](https://semver.org/):

- **MAJOR**: Breaking API changes (signature changes, removed functions)
- **MINOR**: New functions, new modules
- **PATCH**: Bug fixes, performance improvements, doc updates
