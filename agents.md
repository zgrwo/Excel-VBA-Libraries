# agents.md — Excel-VBA-Libraries 项目宪法

> 高性能 VBA 函数库：15 个模块，纯 VBA 实现，零外部依赖。
> 本文件面向 AI 编程助手，编码细节按需加载 Skill。

## 元数据

- **项目名**：Excel-VBA-Libraries
- **GitHub**：https://github.com/zgrwo/Excel-VBA-Libraries
- **语言**：VBA（文档中文，代码注释英文）
- **数字唯一基准**：`rules/api-reference.md` — 函数签名以此为准
- **SSOT**：每个事实只在一处定义，其余仅链接引用

## 四条核心准则

### 1. 先想后写 (Think Before Coding)

- **不确定就提问**。不要猜测业务规则——去查 specification。
- **说出来你做假设了**。
- **发现架构偏离时停下来**。例如：发现自己在非 VBA-Core 模块中重复实现了 ArrayOps 功能 → 停下，调用 VBA-Core。

### 2. 简洁至上 (Simplicity First)

- **最少代码解决问题**。
- **不为一成不变的场景建抽象层**。
- **自检**：一个资深 VBA 开发者看这段代码会觉得过度设计吗？

### 3. 精准修改 (Surgical Changes)

- **只改该改的**。不要顺带重构无关模块。
- **匹配现有风格**。
- **发现无关问题时提出来，不擅自改**。

### 4. 目标驱动 (Goal-Driven Execution)

- **先定义验证方式，再开始写代码**。

| 而不是 | 而是 |
|--------|------|
| "添加函数" | "新函数通过交叉验证 + 集成测试双路径。去验证。" |
| "修复 Bug" | "复现测试 FAILS → 修复后 PASSES + 无回归。去验证。" |

## 技能加载

| 范围 | Skill 文件 | 内容 |
| :--- | :--- | :--- |
| 编辑 .bas / .cls | `skills/vba-SKILL.md` | VBA 陷阱、数组、错误处理、UDF 规范 |
| 编辑手册 / 新增 Public 函数 | `skills/vba-manual-authoring.md` | 6 处同步更新清单 |
| 编写 tests/ Python | `skills/python-SKILL.md` | COM 测试陷阱 |
| 修改 SqlUtils.bas | `skills/sql-SKILL.md` | ADODB 安全 |

### 专家 Skill（重构生命周期）

| 阶段 | Skill | 触发时机 |
|------|-------|----------|
| 决策前 | `skills/architecture-reviewer.md` | 新增组件/层级/依赖前 |
| 执行中 | `skills/refactoring-guardian.md` | 每个 Phase 开始/结束时 |
| 执行后 | `skills/project-plan-review.md` | 里程碑复盘/规划评审时 |

## 架构分层

```
VBA-Core (类模块)
  VariantKit → ArrayOps → DictProxy（导入时按此顺序）
      ↑ 依赖
src/ (标准模块)
  15 个 .bas 模块，相互独立（均依赖 VBA-Core）
  例外：RegressUtils 依赖 LinearUtils + StatsUtils
```

## 仓库目录树

> 路由地图：所有文件路径均以此为基准。详细结构见 [project-structure.md](rules/project-structure.md)。

```
ExcelVBA函数库/
├── VBA-Core/                       # 公共基础设施（VariantKit/ArrayOps/DictProxy）
├── src/                            # 15 个 .bas 标准模块
├── tests/                          # 4 层测试体系（验证/交叉/集成/单元）
├── docs/                           # 用户文档 + 二进制工作簿
├── scripts/                        # 开发工具脚本（hooks/结构校验）
├── rules/                          # 治理规范文档
├── skills/                         # AI 编码规范（扁平结构）
├── tools/                          # 辅助工具（翻译）
├── .github/                        # CI + Issue/PR 模板
├── build/                          # 构建产出
├── logs/                           # 日志
├── agents.md                       # 本文件
├── README.md                       # 用户向功能指南 (CN)
├── README_EN.md                    # 用户向功能指南 (EN)
└── LICENSE / CONTRIBUTING.md / ...  # 社区文件
```

## 红线规则

### 1. 参数类型

- 🔴 所有 Public UDF 参数必须 `As Variant`（否则 Range 传入 #VALUE!）
- 🔴 Public 函数必须同时处理 Range 对象和 Variant 数组（双路径）

### 2. 错误处理

- UDF → `CVErr(xlErrValue)`
- VBA 内部函数 → `Err.Raise`
- 资源操作 → `On Error GoTo Cleanup`
- 🔴 禁止裸 `On Error Resume Next` 不检查 Err

### 3. 新增 Public 函数 6 处同步

源码 + API Index（签名行+计数头）+ CN 手册 + EN 手册 + 交叉验证 + 锚点校验

### 4. VBA-Core 接口冻结

禁止修改 VBA-Core 类的 Public 接口，除非用户明确要求。

### 5. 推送范围

禁止推送 `rules/project-structure.md` 结构树之外的任何文件。

## 测试体系

| 层 | 工具 | 路径覆盖 |
|---|---|---|
| ① 一致性验证 | `run_all_validation.py`（无需 Excel） | 签名/交叉引用/元数据/代码质量 |
| ② 交叉验证 | `run_all_crossval.py`（需 Excel） | **数组路径**：Python list → COM Variant 数组 |
| ③ 集成测试 | `integration_test_all_modules.py`（需 Excel） | **Range 路径**：COM Range → VBA Variant |
| ④ 单元测试 | VBA Test_*（仅 SqlUtils） | ADODB 无法 COM 测试 |

> 🔴 **双路径原则**：交叉验证测数组路径，集成测试测 Range 路径。两者必须都通过。

## 历史经验（从 diff 提炼）

### 高频修复模式

| 模式 | 出现次数 | 根因 |
|------|----------|------|
| Err.Raise 拼写错误（XxxErr.Raise） | 8+ | 复制粘贴时多了模块前缀 |
| Boolean 守卫缺失 | 5+ | RankEq/RankAvg 等未检查输入类型 |
| IsArray 不可靠 | 3 | DictTopN 等使用 IsArray 判断失败 |
| 水合物系数递归 | 2 | MolecularWeight 未跳过系数直接递归 |
| 矩阵奇异未捕获 | 2 | MatrixDeterminant 未处理 ERR_SINGULAR |

### 关键设计决策

- 纯 VBA 零依赖：无需安装运行时，兼容 Excel 2010+
- Python COM 交叉验证：numpy/scipy 独立计算期望值
- 双路径测试：数组路径 + Range 路径全覆盖

## 构建与测试

| 场景 | 命令 |
| :--- | :--- |
| 快速验证 | `python tests/run_all_validation.py --quick` |
| 全量验证 | `python tests/run_all_validation.py` |
| 交叉验证 | `python tests/run_all_crossval.py`（需 Excel） |
| 集成测试 | `python tests/utils/integration_test_all_modules.py`（需 Excel） |

## 开发流程

### 修改前（强制）

1. **Read** 对应 Skill 文件（Skill-first）
2. 检查调用者与影响范围
3. 确认不违反红线规则

### 遇到 Bug 时

1. 写最小复现测试 → confirm: FAILS
2. 修复 → confirm: PASSES + 无回归
3. **保留复现测试**
4. 检查是否需要更新 spec / skill

### 提交前必检

- [ ] 所有 Public UDF 参数为 `As Variant`
- [ ] 错误处理使用正确模式
- [ ] 数组操作使用 LBound/UBound
- [ ] 无裸 `On Error Resume Next` 不检查 Err
- [ ] 新增函数已同步 6 处（源码+API Index+CN手册+EN手册+交叉验证+锚点）
- [ ] `python tests/run_all_validation.py` 通过

## 防幻觉铁律

| 铁律 | 说明 |
|------|------|
| **不靠记忆引用文档** | 先 Read/Grep 确认 |
| **不确定 = 承认** | 去查 spec |
| **写过的 = 读过的** | Read 它再改 |
| **版本号是事实锚点** | 每个结论标注来源文档版本，防止误用过时信息 |

## 会话管理

### 何时自查

- **每完成一个独立功能点** — 对照四条核心准则自检
- **上下文超过 5 个文件 / 20 轮对话** — 提醒用户开新会话

### 跨会话接力

```
上一个会话结束时 → 简述：
  ✅ 已完成 / 🔜 下一步 / ⚠️ 待决策 / 📄 关键上下文
```

### 基本原则

- 新会话先读本文件 + `skills/vba-SKILL.md`
- 跨会话通过 git commit 衔接
- 每个 commit 自包含、可追溯

## 参考

| 文档 | 角色 |
| :--- | :--- |
| [README.md](README.md) | 用户入口、模块速览、使用模式 |
| [context.md](rules/context.md) | 术语表 — 所有领域术语唯一定义 |
| [api-reference.md](rules/api-reference.md) | 签名唯一信源 |
| [user-manual.md](rules/user-manual.md) | 用户手册 |
| [project-structure.md](rules/project-structure.md) | 结构地图 |
| [documentation.md](rules/documentation.md) | 文档职责 |
| [code-review-prompt.md](rules/code-review-prompt.md) | 审查模板 |
| [refactoring-plan.md](rules/refactoring-plan.md) | 重构计划 |

<!-- last_updated: 2026-07-24 -->
