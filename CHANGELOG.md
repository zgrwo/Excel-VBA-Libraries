# Changelog

All notable changes to Excel VBA Libraries.

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
