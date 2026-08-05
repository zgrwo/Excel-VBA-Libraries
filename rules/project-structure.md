# 项目结构

> 本文件是项目文件/目录的**唯一定义**。`scripts/check_project_structure.py` 和提交钩子以此文件为准。
> 新增文件后必须同步更新此结构。结构树中直接列出的文件为**精确匹配**；无子项的目录为**通配匹配**（任意内容均允许）。
> **SSOT 主目录**：`ExcelVBA函数库/`（Harmonization 仓库）。`_analysis/Excel-VBA-Libraries/` 仅为历史分析副本。

## 结构树

```
ExcelVBA函数库/
├── VBA-Core/                       # ✅ 公共基础设施 — 可独立复用
│   ├── VariantKit.cls              # 类型归一化
│   ├── ArrayOps.cls                # 通用数组操作
│   └── DictProxy.cls               # 安全字典 + 批量操作
├── src/                            # VBA .bas 模块（15 个）
│   │  数据层
│   ├── ArrayUtils.bas            # 数组 — 排序/筛选/切片/聚合/查找
│   ├── DictSetUtils.bas          # 字典/集合 — 合并/交集/差集/频率
│   ├── PivotUtils.bas            # 重塑 — 透视/逆透视/分组/交叉连接
│   ├── SqlUtils.bas              # SQL — SELECT/JOIN/GROUP BY (ADODB)
│   │  统计/数学层
│   ├── LinearUtils.bas           # 线性代数 — SVD/QR/LU/Cholesky/PINV
│   ├── StatsUtils.bas            # 统计 — 描述/推断/分布/关联
│   ├── RegressUtils.bas          # 回归 — OLS/ANOVA/因子重要性/优化
│   │  文本层
│   ├── StringUtils.bas           # 字符串 — 编码/校验/距离/UUID/URL
│   ├── RegexUtils.bas            # 正则 — 匹配/替换/分割/捕获组
│   ├── JsonUtils.bas             # JSON — 纯 VBA 递归下降解析器
│   ├── XmlUtils.bas              # XML — MSXML2 XPath 查询与转表格
│   │  日期
│   ├── DateTimeUtils.bas         # 日期 — ISO周/工作日/年龄/复活节
│   │  Excel/文件层
│   ├── RangeUtils.bas            # Range — 导出(HTML/JSON/MD)/区域运算/命名
│   ├── FileSystemUtils.bas       # 文件 — UTF-8 读写/文件夹/驱动器
│   │  理化计算
│   └── PhyChemUtils.bas          # 理化 — 分子量/单位换算/气体标准态
├── tests/                          # 4 层测试体系
│   ├── run_all_validation.py     # ① 一致性验证（无需 Excel）
│   ├── run_all_crossval.py       # ② 交叉验证（需 Excel）
│   ├── run_all_tests.py          # 测试入口
│   ├── coverage_report.md        # 测试覆盖率报告（自动生成）
│   ├── test_utils.py             # 测试公共工具
│   ├── retry_decorator.py        # COM 测试重试机制
│   ├── crossval/                 # 各模块交叉验证构建器
│   ├── utils/                    # 集成测试/锚点校验/文档验证
│   ├── data/                     # 测试数据 (.xlsx)
│   └── benchmarks/               # 性能基准
├── docs/                           # 用户文档 + 二进制
│   ├── VBA_LIB_User_Manual_EN.md # 用户手册·英文
│   ├── vba-pitfalls.md           # VBA 陷阱清单（新增函数前必查）
│   ├── annotation-guide.md       # @Description 注解规范
│   ├── review-checklist.md       # 代码审查清单
│   ├── release-checklist.md      # 发布清单
│   ├── security-audit.md         # 安全审计文档
│   ├── VBA_Libraries.xlsm        # 含全部模块 ⚠️ 二进制
│   └── VBA_Libraries.xlam        # 加载项版本 ⚠️ 二进制
├── scripts/                        # 开发工具脚本
│   ├── vba_lint.py               # VBA 静态检查（8 条规则）
│   ├── generate_counts.py        # 函数计数自动生成
│   ├── generate_coverage.py      # 测试覆盖率报告生成
│   ├── generate_api_docs.py      # API 签名提取
│   ├── changelog.py              # 从 git log 生成 CHANGELOG
│   ├── run_all_tests.ps1         # 一键测试（整合 4 层）
│   ├── rebuild.ps1               # 重建 .xlsm/.xlam（需 Excel）
│   ├── pre_commit_check.py       # 提交前门禁（lint + validation）
│   ├── add_module_headers.py     # 模块文档头生成器
│   ├── fix_magic_numbers.py      # 魔法数字→命名常量
│   ├── pre-commit                # Git pre-commit hook
│   ├── check_project_structure.py # 项目结构校验（被 hook 调用）
│   ├── install_hooks.sh          # 安装 hook 到 .git/hooks/
│   └── push.sh                   # 推送前校验 + 文件清单 + 推送
├── rules/                          # 治理规范文档（本目录）
│   ├── project-structure.md      # 📍 本文件 — 结构唯一定义
│   ├── api-reference.md          # 函数签名唯一信源
│   ├── user-manual.md            # 用户手册（治理版）
│   ├── specification.md          # 项目规格文档
│   ├── refactoring-plan.md       # 重构计划
│   ├── context.md                # 术语表
│   ├── documentation.md          # 文档职责
│   └── code-review-prompt.md     # 审查模板
├── skills/                         # AI 编码规范（扁平结构）
│   ├── vba-SKILL.md              # VBA 编码规范
│   ├── vba-manual-authoring.md   # 手册撰写规范
│   ├── python-SKILL.md           # COM 测试规范
│   ├── sql-SKILL.md              # ADODB SQL 规范
│   ├── architecture-reviewer.md  # 架构审查（重构生命周期）
│   ├── refactoring-guardian.md   # 重构守卫（每 Phase 执行）
│   └── project-plan-review.md    # 规划效果审查
├── tools/                          # 辅助工具
│   ├── translate_en.py           # 中→英翻译辅助
│   └── cn_phrases.txt            # 中文短语表
├── .github/                        # GitHub 社区健康文件
│   ├── ISSUE_TEMPLATE/
│   │   ├── bug_report.md        # Bug 报告模板
│   │   └── feature_request.md   # 功能请求模板
│   ├── PULL_REQUEST_TEMPLATE.md # PR 模板
│   └── workflows/
│       └── ci.yml               # GitHub Actions CI
├── .claude/                        # 代码审查工作区
│   └── reviews/
│       ├── CODE_REVIEW_KNOWN_ISSUES.md  # 已知问题注册表
│       └── CODE_REVIEW_VERIFICATION_*.md # 审查复核记录
├── .qoder/                         # Qoder 代理配置
│   └── skills/                     # Harness 可识别的技能注册（skills/ 的镜像）
│       └── */SKILL.md              # 每个技能一个目录，文件名必须为 SKILL.md
├── build/                          # 构建产出（.gitignore）
├── logs/                           # 日志（.gitignore）
├── AGENTS.md                       # 项目宪法（AI 入口）
├── README.md                       # 用户入口·中文
├── README_EN.md                    # 用户入口·英文
├── LICENSE                         # MIT 许可证
├── CONTRIBUTING.md                 # 贡献指南
├── SECURITY.md                     # 安全策略
├── CHANGELOG.md                    # 变更日志
├── CODE_OF_CONDUCT.md              # 行为准则
├── .editorconfig                   # 跨编辑器代码风格
├── .gitattributes                  # Git 行尾规范化
├── .rubberduck                     # Rubberduck 推荐配置（可选）
└── .gitignore                      # Git 忽略规则
```

## 层级划分

### VBA-Core/ — 基础设施层

独立可复用的类模块，为所有 `src/` 模块提供类型安全、数组操作和字典操作的基础能力。

**导入顺序**：`VariantKit.cls` → `ArrayOps.cls` → `DictProxy.cls`

### src/ — 功能模块层

15 个 `.bas` 模块按领域分 6 组：

| 层级 | 模块 | 职责 |
|------|------|------|
| 数据层 | ArrayUtils, DictSetUtils, PivotUtils, SqlUtils | 数组运算、集合操作、数据重塑、SQL 查询 |
| 统计/数学层 | LinearUtils, StatsUtils, RegressUtils | 线性代数、统计分析、回归建模 |
| 文本层 | StringUtils, RegexUtils, JsonUtils, XmlUtils | 字符串编码、正则处理、JSON/XML 解析 |
| 日期 | DateTimeUtils | ISO 周、工作日、年龄、时间戳 |
| Excel/文件层 | RangeUtils, FileSystemUtils | Range 导出、UTF-8 文件读写 |
| 理化计算 | PhyChemUtils | 分子量、单位换算、气体标准态 |

`src/` 层各模块相互独立（均依赖 VBA-Core），可按任意顺序导入。例外：`RegressUtils.bas` 依赖 `LinearUtils.bas` 和 `StatsUtils.bas`（需先加载）。

### tests/ — 测试层

4 级测试体系（按执行速度排序）：

1. **一致性验证** — `run_all_validation.py`（无需 Excel，<2s ~ 10s）
2. **交叉验证** — `run_all_crossval.py`（需 Excel，numpy/scipy 对比，~5-10 min）
3. **集成测试** — `utils/integration_test_all_modules.py`（需 Excel，COM Range 路径，~2-3 min）
4. **单元测试** — VBA `Test_SqlUtils`（仅 SqlUtils 保留，ADODB 无法 COM 测试）

### docs/ — 用户文档层

| 文档 | 受众 | 核心问题 |
|------|------|----------|
| `VBA_LIB_User_Manual_EN.md` | 人类用户 | "I want to do X — what formula?"（English） |
| `review-checklist.md` | 审查者 | 代码审查逐项确认清单 |
| `VBA_Libraries.xlsm` | Excel 用户 | 含全部模块的工作簿（按需导入） |
| `VBA_Libraries.xlam` | Excel 用户 | 加载项版本（全局可用） |

> ℹ️ API 签名速查表和中文用户手册的 SSOT 分别在 `rules/api-reference.md` 和 `rules/user-manual.md`。

### rules/ — 治理规范层

| 文档 | 角色 |
|------|------|
| `project-structure.md` | 📍 本文件 — 结构唯一定义 |
| `api-reference.md` | 函数签名唯一信源（治理版） |
| `user-manual.md` | 用户手册（治理版） |
| `specification.md` | 项目规格 |
| `refactoring-plan.md` | 重构计划 |
| `context.md` | 术语表 |
| `documentation.md` | 文档职责 |
| `code-review-prompt.md` | 审查模板 |

### skills/ — AI 编码规范层

按任务类型按需加载的领域规范（扁平结构）：

| 文档 | 何时加载 |
|------|----------|
| `vba-SKILL.md` | 编辑 .bas / .cls |
| `vba-manual-authoring.md` | 编辑 docs/ 手册 · 新增 Public 函数 |
| `python-SKILL.md` | 编写 tests/ Python |
| `sql-SKILL.md` | 修改 SqlUtils.bas |
| `architecture-reviewer.md` | 新增组件/层级/依赖前 |
| `refactoring-guardian.md` | 每个 Phase 开始/结束时 |
| `project-plan-review.md` | 里程碑复盘/规划评审时 |

### scripts/ — 开发工具层

Git hooks 和自动化脚本，由提交钩子直接调用。

### tools/ — 辅助工具层

翻译辅助等非核心工具。

### 根目录文件

| 文件 | 角色 |
|------|------|
| `AGENTS.md` | 项目宪法（AI 入口），仅保留一级路由 |
| `README.md` | 用户入口·中文 |
| `README_EN.md` | 用户入口·英文 |
| `LICENSE` | MIT 许可证 |
| `CONTRIBUTING.md` | 贡献指南 |
| `SECURITY.md` | 安全策略 |
| `CHANGELOG.md` | 变更日志 |
| `CODE_OF_CONDUCT.md` | 行为准则 |
| `.editorconfig` | 跨编辑器代码风格 |
| `.gitattributes` | Git 行尾规范化 |
| `.gitignore` | Git 忽略规则 |
