# Changelog

All notable changes to Excel VBA Libraries.

## [Unreleased] — 2026-09-13

### Added（SOLVE 工艺参数反解 v1）

- **SolveUtils（第 16 个 `src/` 模块）**: `UDF_SOLVE_INVERSE`（给定输出目标反推可调参数：有界多起点模式搜索 + 可达性判定）、`UDF_SOLVE_PREDICT`（前向预测）、`UDF_SOLVE_QUALITY`（5 折/LOO 交叉验证质量表）、`UDF_SOLVE_EQUATION`（前向方程 + 单可调线性闭式反解）。自动模型选择 `auto`/`linear`/`poly`（poly 岭回归 λ=1e-5）、确定性 XorShift64*（`seed` 默认 42）、角色前缀（`Incoming*`/`来料*`、`Variable*`/`可调*`/`变量*`、`Fixed*`/`固定*`、`Output*`/`输出*`）、独立 request 表与边界表、规模上限（历史 2000 行/请求 100/特征 30/可调 10/输出 10）。依赖 `LinearUtils`（第二个合法依赖例外，已登记）
- **测试**: `tests/crossval/build_SolveUtils.py` 11 个独立 Python 参考用例（numpy lstsq/闭式解/G6 文本）；集成测试 Range 路径 13 例（INVERSE 精确夹具/不可达/QUALITY/EQUATION/PREDICT/负例 CVErr）
- **文档**: `rules/api-reference.md` SolveUtils 表与计数头（16 模块 541 Functions）、README CN/EN 模块速览、AGENTS/context/specification/project-structure 16 模块与依赖例外同步

### Fixed（第三轮全量深度审查修复批）

- **JsonUtils**: 空 JSON 数组改返回已分配 0 长度数组（修复 `JsonStringify(JsonParse("[]"))`/`JsonGetKeys("[]")`/`JsonToRange "[]"`/`JsonGet("[]","[0]")` 全链路 Error 9）；`JsonStringify`/`UDF_JSON_STRINGIFY` 归一化 Range 输入（不再输出字面量 `"Range"`）；`JsonGet` 支持已解析 Dictionary 输入（parse-once/query-many）；字符串内未转义控制字符按 RFC 8259 §7 拒绝；指数溢出改模块错误码（`9e308` 不再裸 Error 6）
- **StringUtils**: `TextJoin`/`UDF_STR_TEXTJOIN` 支持 2D 数组与 Range（修复 Error 9）；`HTMLDecode` 改单遍解码（修复 `&#38;lt;` → `<` 双重解码）；`RemoveChars`/`KeepChars` 主路径与回退路径统一为大小写敏感
- **VariantKit**: `FilterPasses(Null,…,"isblank")` 不再 Error 94（Null 视为空白）；`ValuesEqual` 新增 Array/Object 分支（不再 CStr 崩溃）且数值容差改尺度感知（0 与 1E-13 不再判等）；`Compare` 混合数值/文本改类型秩排序（恢复传递性，修复 ArraySort 错序）；数组比较补第 2 维长度；测试 harness 失败不再被自身错误处理器吞没
- **FileSystemUtils**: `ValidateSafePath` 统一分隔符后判定 UNC（修复 `/\server` 混合写法绕过）、拒绝通配符与 ADS；`FileExists`/`FolderExists`/`ListFiles`/`ListFolders`/`EnsureFolder` 接入安全检查；`ReadBinaryFile` 安全校验先于存在性探测；`WriteUTF8` 修复双 BOM 且 `bom=False` 生效；`UTF8DecodeBytes` 修复 minCp 残留导致 ASCII 变 U+FFFD；`WriteANSI` 失败不再静默
- **SqlUtils**: `SqlRangeQuery` 文本列改 adVarWChar(32767)（修复 33 字符即整表失败）、类型推断改 VarType 白名单 + Date/Error 处理（不再把 `"01234"`/长 ID 转数值）、`SqlEscapeString(forLike)` 改 ACE 方括号转义（KI-009 重开）、Provider 回退链补 ACE 16 与位数提示、大小写重名表头按键去重、WHERE 词边界识别 + ORDER BY 明确拒绝、.xlam 默认改用活动工作簿
- **LinearUtils**: `RangeToMatrix` 数组路径保持 2D 形状（修复 `UDF_LINALG_DET({1,2;3,4})` 等数组常量输入全线错误）；SVD 仅精确零列置零（不再截断 U 正交性）；LU 奇异判据改列范数相对尺度（高动态范围良态矩阵不再误判）；`PolyFit` 混基输入/行向量/负 degree 修复
- **StatsUtils**: `TInv2T` 上界自适应（修复 df 小、alpha 小返回 ≈100）；`GammaLn` 负值反射取绝对值（修复 (-1,0) 区间裸 Error 5）；0 变差判定改偏差尺度（高偏置低变差数据不再误报零方差，同时保留小尺度支持）；`CorrelationMatrix` 无数值列改用 IsEmptyArray 探测
- **RegressUtils**: `FitOLS`/`FitCoefOnly`/`TriangularInverse` 容差去除绝对下限（小尺度良态矩阵不再静默零系数）；`ANOVAOneWay` dfW=0 不再输出假显著；退化分类因子（全空/单水平）不再 Error 9；`FitOLS` 补欠定保护；`StandardizeColumns` 常量列判定改偏差尺度
- **ArrayUtils**: `ArraySum`/`ArrayMin`/`ArrayMax`/`ArrayFind` 标量按单元素处理；`ArraySample(Array())` 空数组返回空；`ArrayFind` Error 值按错误码精确匹配
- **XmlUtils**: `XmlGet` 不再吞解析错误；`XmlToRange` colNames 支持 2D 数组常量与 Range（修复手册示例必失败）
- **RegexUtils**: `RegexFindInRange` 批次地址按 255 字符上限动态分批
- **PhyChemUtils**: `IsUnknown` 支持参数省略（Missing）
- **RangeUtils**: `SafeWriteValue` 合并单元格重定向后复查公式（不再静默覆盖锚点公式）；`RangeToJSON` 转义覆盖全部 <0x20 控制字符（RFC 8259）
- **DateTimeUtils**: `Age` 的 `asOf` 支持日期序列号（不再静默改用今天）
- **StatsUtils**: `TDistCDF` 补 df>0 域校验
- **FileSystemUtils**: `CollectFolders` 枚举错误不再静默吞没（与 CollectFiles 一致）
- **测试基建**: 交叉验证通过率剔除 SKIP 并输出 effective rate/NO RUNS；`run_all_tests` 注入改为守卫式（断言失败写 FAIL 行，不再静默）；`validate_docs` 计数头/AUTO_COUNTS 与锚点检查实做、`analyze_test_gaps` 跳过 skip 用例并将无覆盖模块判 FAIL；集成测试 `FitOLS` 用例修复数据别名；数组比较补形状断言；新增 20+ 条回归用例（JSON 负例探针、TextJoin 2D、HTMLDecode、Controls、TInv2T/GammaLn/高偏置、RangeToMatrix、FS 通配符/混合 UNC、Xml 2D colNames 等）

### Fixed（第四轮全量审查修复批 — SOLVE 移植后，2026-09-13）

> 报告 `.claude/reviews/CODE_REVIEW_2026-09-13_R4.md`（22 项：12 High / 10 Medium，无 Critical；含 1 项驳回、1 项降级）。
> 修复后回归：一致性验证 ALL PASS / lint 0 Warning / 交叉验证 19 模块 0 FAIL / 集成 109P-0F / 新增 30+ 回归用例。

- **JsonUtils**: 慢速路径 `AscW` 有符号比较修复（`And &HFFFF&`）——转义序列后的代理对/高位 BMP（emoji、U+FFFD）不再被误判为控制字符（R4-05，第三轮修复引入的近期回归）；同族站点 `FileSystemUtils`（3 处非 ANSI 检测）与 `StringUtils.URLDecode` 一并修复
- **RegexUtils**: UDF 标量路径错误值不再 `CStr` 字符串化（错误输入返回 `#VALUE!`，此前 `#DIV/0!` 会得到 `TRUE`/计数 1）；核心 `RegexCount` 增加 `IsError` 守卫；新增 `RegexEnsure2D` 使 1D Variant 数组路径可用（此前 Error 9/13）（R4-03/R4-06）
- **RegressUtils**: `InteractionEffects` 2D 首维 `ReDim Preserve` 修复（单水平分类因子不再 Error 9，R4-02）；`LinearModelPredict` 单列 Range 按列内读取（不再越界读取相邻单元格，R4-04）；`OptimizeFactors` 未知 goal 报错（不再静默按 0 分返回，R4-09）；`ANOVAOneWay` 改居中空间累加（修复 `1e13` 量级常数响应假显著 F=12.38/p=4.4e-6，R4-11）；`FitOLS` 新增 `rank`/`rank_deficient` 键暴露秩亏（R4-12）；分类空白行按无效处理（不再静默并入参考水平，R4-14）；分类水平收集改 `vbBinaryCompare`（大小写变体不再错编码，R4-15）；`GetCategoricalLevels`/`EncodePredictRow` 补 Null/Error 守卫（COM 数组传入 Null 不再 Error 94）
- **SolveUtils**: `SvExpression` 缓冲由 `POLY_TERM_LIMIT` 派生（修复 90 项 poly 方程 `#VALUE!`，R4-01）；schema 定长 UDT 写前容量校验（超限返回限额错误而非 Error 9，R4-07）；`SvFitOLS` 新增 `RankDeficient` 标志并在 `UDF_SOLVE_EQUATION` 输出 `备注` 行（R4-12）；模块头计数 4→8 并登记 `add_module_headers.py`（R4-22）
- **LinearUtils**: 经济模式 QR 宽矩阵按标准 reduced QR 返回 `R=m×n`（`Q·R=A` 可重构，R4-10）；`EigenSymmetric` 容差内输入显式取对称部分 `(A+Aᵀ)/2`（R4-13，复核降级 Medium）
- **PivotUtils**: `GroupBy` 空白聚合值跳过（AVG/MIN/MAX 不再按 0 参与，R4-16）
- **DictSetUtils**: 标量/单格入参按单元素集合（`SetIntersect(1,[1,2])` 不再返回空集、`SetIsSubset` 不再恒 True，R4-17）
- **RangeUtils**: `FilterRangeToArray` 运算符白名单校验（未知运算符报错，不再静默只返回标题行，R4-18）
- **DateTimeUtils**: `DaysInMonth` 单数值按 Excel 日期序列号解释（无参仍为当月，越界报错；不再静默返回当前月，R4-19）；`IsHoliday` 识别日期序列号数组/单元格（R4-20）
- **StatsUtils**: `ZScore` 非法 `value` 报错（不再静默返回全量数组，R4-08）
- **SqlUtils**: `.xlsm` 需要 `Excel 12.0 Macro` ISAM 的断言经 ACE 实测**驳回**（ACE12/16 以 `Excel 12.0` 可读 .xlsm，KI-028）
- **测试基建**: SolveUtils 交叉验证新增数组路径/poly/auto/bounds/负例（此前 24 用例全 Range + 全 linear，R4-21）；`analyze_test_gaps` 输出 array/range 路径覆盖并将仅 Range 模块标为 `partial`；负例统一改 **VBA 内直接调用探针**（错误不再穿透 COM，消除 Excel 运行时错误弹窗导致的测试中止）；`ensure_excel` 改晚绑定 `Dispatch`（消除 gencache/makepy 抖动）——全量 19/19 模块 **951P/0F** 稳定通过；新增 Json 转义代理对、Regex 错误值/1D 数组、QR 宽矩阵重构、Eigen 对称化、InteractionEffects 单水平、空白分类/大小写分类、垂直 Range 预测、未知 goal、ANOVA 高偏置、GroupBy 空白、集合标量、未知运算符、DaysInMonth/IsHoliday、ZScore 负例、SolveUtils poly 方程/auto/bounds/限额等 30+ 回归用例

### Changed

- api-reference.md 补齐 `SqlEscapeString [forLike]`、`XmlGet/XmlGetAttr/XmlToRange [namespaces]` 签名；CN/EN 手册修正 PhyChem 单位（ConvertStandard m³/Pa/K、CylinderStdVolume Pa/K、CompressFactorPR Pa）、DilutionSolve/Density 未知项约定、MinMax 保留参数说明、JsonToRange 签名
- CN/EN 手册同步第四轮行为：DaysInMonth 单参数序列号语义、IsHoliday 序列号、FilterRangeToArray 运算符全表/未知报错、GroupBy 空白跳过、集合运算标量=单元素集、ZScore value 校验、OptimizeFactors goal 校验、SolveEquation 秩亏备注行
- sql-SKILL §10 补 ACE LIKE 方括号转义说明；AGENTS.md NormalizeInput 措辞与现状对齐；README.en/specification 版本号同步 2.1.1

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
