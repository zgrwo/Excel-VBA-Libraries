# SOLVE.* 工艺参数反解 VBA 移植方案

> **来源依据**：ExcelFormulaLabs `docs/plans/2026-09-12-solve-inverse-migration-plan.md`（C# 版，495 行）+ ADR-0007（有界搜索）/ ADR-0008（速率模型）/ ADR-0009（速率演进）
> **目标项目**：Excel-VBA-Libraries v2.1.1+（审修复批 `1a94a64` 之后）
> **文档日期**：2026-09-13
> **文档状态**：待评审（评审通过后按 Phase 逐项执行）
> **定位**：本文件是 **VBA 专项端口方案**。C# 文档中"功能契约/ADR 语义/验收判据"可搬；其工程分解（`dotnet test`、CrossValRunner、`object[,]`、`#VALUE!` 语义）**不适用**，已在本方案中折算为 VBA 等价物。

---

## 1. 结论摘要

| 维度 | 结论 |
|------|------|
| 价值 | **高** —— 目标库 15 模块无"给定输出目标反推可调参数"能力，`RegressUtils.OptimizeFactors` 是网格正向优化，`UDF_LINALG_SOLVE` 是解方程，均不等价 |
| 难度 | **中高** —— C# 核心 2236 行 + UDF 壳 77 行 + 测试 1324 行；约 35~45% 可复用目标库资产，其余需逐段移植（岭回归/RNG/优化器/CV/表解析/可达性） |
| 形态 | **增量新模块**（第 16 个 `src/` 模块），不改 VBA-Core 接口，不改现有 15 模块 |
| 估算 | **v1（linear/poly）7~12 人日**；**全量（rate/rate_poly/共享速率）14~20 人日** |
| 建议 | 分两期：v1 四个 UDF 落地；v2 按 ADR-0008/0009 补速率与共享池化 |
| 关键修订 | 源文档 §8 的 `n=5000 / ≤5s` 与 §5 ASYNC 在 VBA **不可达**（无异步 UDF + 解释执行），上限与性能预算见 §6 |

---

## 2. 范围

### 2.1 v1（本期）

- 4 个 UDF：`UDF_SOLVE_INVERSE` / `UDF_SOLVE_PREDICT` / `UDF_SOLVE_QUALITY` / `UDF_SOLVE_EQUATION`
- `model` = `auto` / `linear` / `poly`（含二次展开与两两交互）
- 表角色解析：`Incoming*`/`Variable*`/`Fixed*`/`Output*` 与中文 `来料*`/`可调*`/`变量*`/`固定*`/`输出*`
- 多输出、多可调参数、边界表、独立 request 表、请求行初值
- 5 折 / LOO 交叉验证 + `auto` 门控（CV R² 择优）
- 有界多起点模式搜索 + 可达性采样 + 推荐表/质量表/方程表
- `seed` 确定性（XorShift64）

### 2.2 v2（后续评估）

- `rate` / `rate_poly`：时间列后缀配对（`Time`/`时间`）、速率列输出（ADR-0008）
- `SharedOutput*`/`共享输出*` 跨输出共享速率池化（ADR-0009）

### 2.3 非目标（明确排除）

- GPR/GBM 等 ML 模型（纯 VBA 无法承载）
- 异步 UDF（VBA 没有 `ExcelAsyncUtil`；只能做规模上限）
- 迁移后的行为不要求与 C# 逐位一致，以"目标达成 + 量级一致 + 确定性可复现"为判据

---

## 3. 目标库资产与差距

### 3.1 可复用（降低移植量）

| 资产 | 用途 | 备注 |
|------|------|------|
| `LinearUtils.QRDecomposition`（Householder，经济模式） | OLS 与岭回归内核 | 岭回归用增广矩阵 `[[Z],[√λ·I]]` 的 QR 实现，无需新增岭回归函数 |
| `LinearUtils` SVD / 矩阵运算 / `MatrixTranspose` | 数据组装与退化回退 | 已通过第三轮审查（SVD/LU 修复在位） |
| `RegressUtils.FitOLS` | v1 线性对照组（可选） | QR 版本；若引用则 SolveUtils→RegressUtils 依赖需登记 |
| `VariantKit.NormalizeTo2D` / `IsEmptyArray` / `IsNumericCell` | Range/数组双路径归一化、空数组探针 | Public 接口冻结，不改 |
| `DictProxy.Create` | 报告行/中间字典 | |
| `ArrayOps.Sort` / `SortIndices` | 数据整理 | |

### 3.2 必须新写

| 组件 | 源实现 | 说明 |
|------|--------|------|
| 表角色解析 | `SolveCore.ParseSchema` | 前缀识别 + 历史/请求行分类 + 双路径入口 |
| 前向模型 | `FitModel` / `Predict` / `BuildPowers` | 列标准化 + 反标准化；poly 展开 `k+k(k+1)/2`，上限 100 项 |
| 岭回归路径 | `RegressionCore.FitRidge`（λ=1e-5） | 增广 QR 实现（见 §3.1） |
| 确定性 RNG | `XorShift64` | 移植为模块级 Private；与 CV 折序、优化器起点、可达性采样共用 |
| 交叉验证 | `CrossValidate` / `FitAuto` | 5 折（n≥20）/ LOO（n<20 且 n≥5）；auto 门控 |
| 有界优化器 | `LocalSearch` + 多起点 | 坐标轮换 + 步长折半；目标函数含目标优先权重与 proximity |
| 可达性 | 2000 点均匀采样 + 相对容差 | 输入/输出区间与 `可达/不可达` 判定 |
| 结果表组装 | `Inverse`/`Quality`/`Equation`/`PredictTable` | 中文表头与状态文字，见 §4.3 |
| 方程文本 | `Expression`/`InverseFormula` | `G6` 系数格式（区域无关） |

---

## 4. 架构落点

### 4.1 模块与命名

| 项 | 决定 |
|----|------|
| 新模块 | `src/SolveUtils.bas`（第 16 个 `src/` 模块） |
| 依赖 | **SolveUtils → LinearUtils**（QR/矩阵）。需在 `project-structure.md` 与 `AGENTS.md` 登记为第二个合法依赖例外；VBA-Core 接口不动 |
| 错误码 | `Private Const ERR_SOLVE_* As Long = vbObjectError + 16xx`（1601–1699，避开现有 1001–3100 段） |
| UDF 命名 | `UDF_SOLVE_INVERSE` / `UDF_SOLVE_PREDICT` / `UDF_SOLVE_QUALITY` / `UDF_SOLVE_EQUATION`（沿用 `UDF_<MODULE>_<NAME>` 全大写约定） |
| 核心命名 | `SolveInverse` / `SolvePredict` / `SolveQuality` / `SolveEquation`（Public VBA 函数，Err.Raise 路径）；内部 `Private` 前缀 `Solve` 或 `SV` 保持一致 |

> 若端口后单文件超过 ~3500 行：优先压缩报告/方程文本段；仍超限时再评审 `SolveCore.bas` 拆分（拆分需重新评估模块独立性红线，默认不拆）。

### 4.2 UDF 契约（参数全 `As Variant`，UDF 层只做归一化 + CVErr）

| UDF | 参数（Excel 名 → 含义） | 默认 |
|-----|--------------------------|------|
| `UDF_SOLVE_INVERSE` | `data`, `[request]`, `[bounds]`, `[model]`, `[seed]`, `[max_starts]` | model=auto, seed=42, max_starts=10 |
| `UDF_SOLVE_PREDICT` | `data`, `values`, `[model]` | model=auto |
| `UDF_SOLVE_QUALITY` | `data`, `[model]`, `[seed]` | model=auto, seed=42 |
| `UDF_SOLVE_EQUATION` | `data`, `[model]` | model=auto |

规则：所有参数 `As Variant`；内部归一化走 `VariantKit.NormalizeTo2D`；失败返回 `CVErr(xlErrValue)`；`Err.Raise` 仅存在于核心 VBA 函数；错误描述中文、代码 `vbObjectError+16xx`。

### 4.3 输出布局（沿用源文档 §3.5，逐字一致）

- `UDF_SOLVE_INVERSE`：`[请求行, <可调列名…>, <输出列名>预测…, 最大偏差σ, 状态]`
  - 请求行：`数据第N行` / `请求N`；状态：`可达` / `不可达`
  - `最大偏差σ = maxⱼ |ŷⱼ−y*ⱼ| / sⱼ`（sⱼ=历史样本 sd，0→1）
- `UDF_SOLVE_QUALITY`：`[输出, 候选, CV方案, CV_R2, CV_MAE, 选用]`；`CV方案`∈{`5折`,`LOO`,`跳过`}；`选用`∈{`是`,`否`}
- `UDF_SOLVE_EQUATION`：`[输出, 类型, 表达式]`；`类型`∈{`前向方程`,`反解公式`}；系数 `G6`、区域无关
- `UDF_SOLVE_PREDICT`：N×M `double[,]`（无表头）

---

## 5. 分阶段任务

> 每个 Task 的完成定义 = **验证锚点全过 + `python tests/run_all_validation.py --quick`**；每个 Phase 收口 = 全量 4 层验证。

### Phase 0 — 预备与治理登记（0.5~1 人日）

- [ ] 决策冻结：v1 范围（linear/poly）、模块名、错误码段、性能上限（§6）、是否引用 `RegressUtils.FitOLS`
- [ ] `rules/project-structure.md`：src 树 + "15 个 .bas" → 16；登记 `SolveUtils → LinearUtils` 依赖
- [ ] `AGENTS.md` / `README.md` / `README.en.md` / `rules/context.md`：模块计数与依赖表述同步
- [ ] 提交：`docs(solve): 移植方案与结构登记`

### Phase 1 — 核心逻辑（3~5 人日）

#### Task 1.1 表解析
- [ ] `ParseSchema` / `IsBlank`：中英前缀、未识别列忽略、行分类、请求行来料缺失报错
- [ ] 锚点：等价 C# `ParseSchema_ClassifiesRowsAndRoles`；中文前缀等价；无 Variable/Output → Err；请求行来料空 → Err
- [ ] 补 crossval：`build_SolveUtils.py` 首批表解析用例（Python dict 参考）

#### Task 1.2 模型拟合与预测
- [ ] `BuildPowers` / `FitModel`（linear/poly，标准化 + 反标准化 + 增广 QR 岭回归 λ=1e-5）/ `Predict`
- [ ] 守卫：行数 < 展开项+1；展开项 > 100（消息含上限）；NonFinite/NaN/Inf 拒绝
- [ ] 锚点：`y=2+0.5a+1.5u` → 系数/截距误差 ≤1e-10；`y=1+u+0.5u²` 预测 ≤1e-6；15 特征 → 报错

#### Task 1.3 RNG + 交叉验证 + auto
- [ ] `XorShift64`（与源实现同参数，逐位可复现）；`CrossValidate`（5 折/LOO）；`FitAuto`（R² 择优，差 <1e-9 取 linear）
- [ ] 锚点：n=25 精确线性 → `5折`、R²≥0.999；n=10 → `LOO`；n<5 → Err；poly 超限 auto 跳过

#### Task 1.4 反解/可达/报告
- [ ] `SolveInverse`（目标函数、多起点、模式搜索、clamp、可达采样、相对容差 1e-9）；`Inverse`/`Quality`/`Equation`/`PredictTable` 报告组装
- [ ] 锚点（源验证锚点逐条折算）：
  - 线性夹具 `a=10, target=13`（u∈[2,20]）→ 推荐 4.0±1e-4、预测 13.0±1e-6、`可达`
  - 目标 1000 → `不可达`；微尺度 y~1e-12 → `不可达`（R-2 回归）
  - 多可调欠定 → 目标达成 ±1e-6；同 seed 两次逐元素一致；请求初值 3.99 → 收敛 4.0
- [ ] 提交：`feat(solve): 核心逻辑（v1）`

### Phase 2 — UDF 层与 6 处同步（1~2 人日）

- [ ] 4 个 `UDF_SOLVE_*`（Variant + 双路径 + CVErr）
- [ ] 锚点：返回布局逐列核对；空 data/无 Variable/无请求行 → `#VALUE!`；QUALITY 表头 6 列；PREDICT 1×1
- [ ] 6 处同步：源码 Module 头 / `rules/api-reference.md`（模块表 + 计数头 `generate_counts.py --check`）/ CN 手册 / EN 手册 / `tests/crossval/build_SolveUtils.py` / 锚点校验
- [ ] 提交：`feat(solve): 四个 UDF 与文档同步`

### Phase 3 — 测试基建（2~4 人日）

- [ ] `tests/crossval/build_SolveUtils.py`：Python **独立**参考（numpy + sklearn LinearRegression/PolynomialFeatures/Ridge + scipy.optimize），**禁止照抄 VBA 算法**
- [ ] 负例探针（VBA 内捕获错误码，模式同 `build_JsonUtils._json_err_probe`）：非法模型名、边界反转、超限、单请求无目标
- [ ] 确定性用例：同 seed 双调用逐元素相等；RNG 若无法与 Python 复刻则改为"行为等价 + 量级"断言并在用例注释说明
- [ ] 集成测试（Range 路径）：INVERSE 精确夹具、不可达、QUALITY 表、EQUATION 文本
- [ ] 手册示例：CN/EN 三步示例数据 + 硬编码期望值（由 Python 复算）
- [ ] `docs/VBA_Libraries.xlsm` 重建（UDF-Range 用例前置新鲜度检查已就绪）
- [ ] 提交：`test(solve): Python 独立对照与双路径用例`

### Phase 4 — 文档收口（1~2 人日）

- [ ] `rules/api-reference.md` 模块表（4 函数签名）与计数头；`rules/user-manual.md` 章节 + 3 个 Recipe；`docs/VBA_LIB_User_Manual_EN.md` 同步
- [ ] `rules/context.md` 术语：反解 / 请求行 / 可达性 / 最大偏差σ
- [ ] `rules/specification.md` 模块清单；README CN/EN 速览行；CHANGELOG `[Unreleased]`
- [ ] 全量验证：`run_all_validation.py` / `vba_lint.py` / `run_all_crossval.py` / `integration_test_all_modules.py` / `run_all_tests.py`
- [ ] 提交：`docs(solve): 文档与门禁收口`

### Phase 5 — v2（可选，按业务评估）

- [ ] rate / rate_poly：时间列识别与后缀配对（ADR-0008）
- [ ] SharedOutput* 池化速率（ADR-0009）
- [ ] 提交：`feat(solve): 速率模型与共享池化`

---

## 6. 规模与性能预算（VBA 修订）

| 项 | C# 原值 | VBA v1 上限 | 理由 |
|----|---------|-------------|------|
| 历史行数 | 5000 | **2000** | QR 拟合可承受；优化器评估次数才是瓶颈 |
| 请求行数 | 200 | 100 | 每请求独立寻优 |
| 特征列 | 50 | 30 | 展开项 ≤100 约束内 |
| 可调参数 | 20 | 10 | 搜索维度成本 |
| 输出数 | 20 | 10 | 报告宽度 |
| poly 项数 | 100 | 100 | 保持 |
| max_starts | 50 | 20（默认 5） | 10×4000 评估在 VBA 可能秒级到十秒级 |
| 单起点评估 | 4000 | 2000 | 同上 |
| 可达采样 | 2000 | 1000 | 只用于区间估计 |
| 性能目标 | n=5000/5 请求 ≤5s | **n=1000/3 可调/5 请求 ≤10s**（Excel 2016 基准机） | 无异步，超限显式 `#VALUE!` |

> 规模上限在 `UDF_SOLVE_*` 入口集中校验并登记错误；`QUALITY`/`EQUATION` 不受优化预算影响。

---

## 7. 测试策略

1. **Python 是裁判**：`build_SolveUtils.py` 用 numpy/sklearn/scipy 独立复算；禁止 `check_close(name, X, X)` 类自校验。
2. **确定性与 RNG**：优先按 `XorShift64` 同参数复刻，使 Python 参考可重放同一洗牌/起点；不可行时降级为"同一 seed 两次 VBA 结果逐元素相等 + 目标达成量级"断言（在用例注释声明）。
3. **错误路径探针**：COM 不传递 VBA 错误描述，负例一律用 VBA 探针返回 `Err.Number` 与期望码比对（`expect_error`/`reconstruct` 机制已具备）。
4. **双路径**：crossval 数组路径 + integration Range 路径；UDF-Range 依赖 `docs/VBA_Libraries.xlsm` 重建。
5. **手册示例**：`build_manual_examples.py` 增 SOLVE 条目，期望值由 Python 复算后硬编码。
6. **性能基线**：可选手工记录（`tests/benchmarks/` 无历史基线，暂不强制）。

---

## 8. 治理影响清单（16 模块）

| 位置 | 变更 |
|------|------|
| `rules/project-structure.md` | src 树 + "15 个 `.bas`" → 16 + 依赖例外登记（已含 `plans/` 目录） |
| `AGENTS.md` | "15 个模块" 表述、模块表/路由、技能触发矩阵（新增 SolveUtils 行） |
| `README.md` / `README.en.md` | 模块计数、功能速览、独立性声明（14/15 → 15/16） |
| `rules/context.md` | "15 个独立标准模块" 与 SolveUtils 依赖例外 |
| `rules/api-reference.md` | 自动计数头（`generate_counts.py --check`）+ 新模块表 |
| `rules/specification.md` | 模块清单新增 SOLVE 行 |
| `docs/VBA_Libraries.xlsm` / `.xlam` | 重建（含新模块） |
| `tests/coverage_report.md` | 重新生成 |
| VBA-Core | **不动**（接口冻结） |

---

## 9. 风险与对策

| 风险 | 对策 |
|------|------|
| VBA 性能不达标（优化器多起点） | §6 上限 + 默认 `max_starts=5` + 文档明示；必要时减少可达采样 |
| 无异步 UDF，Excel UI 阻塞 | 规模上限 + 超限 `#VALUE!`；文档注明大表建议拆分 |
| C#→VBA 数值漂移 | 关键判据用"目标达成/闭式解/量级"，不追求逐位；拟合部分与 numpy QR 对拍 1e-8 |
| RNG 移植偏差 | 单测锁同 seed 双调用一致；跨语言只做量级对照 |
| 模块体积过大 | §4.1 压缩策略；超 3500 行再评审拆分（需治理例外） |
| poly 外推 | 边界默认历史 min/max；触界在手册提示"需实验确认" |
| 文档门禁红 | 先跑 `generate_counts.py --check` 与手册锚点，再一次性同步计数 |
| 二进制陈旧 | `build_udf_range.py` 已加新鲜度告警；发布前 `rebuild.ps1` |

---

## 10. 验收标准（VBA 版）

| 项 | 标准 |
|----|------|
| 功能 | CN/EN 手册三步示例一次成功；精确夹具误差 ≤1e-6；不可达标注正确 |
| 数值 | 线性闭式解 ±1e-4；达成预测 ±1e-6；同 seed 逐元素可复现；微尺度可达判定正确 |
| 测试 | 4 层验证全绿；SOLVE crossval 无 SKIP；负例探针按错误码断言 |
| 门禁 | `run_all_validation.py` ALL PASS；`vba_lint` 0 Warning；锚点/死链 0 |
| 体验 | 表头前缀即唯一约定；返回表无需文档即可读懂 |
| 性能 | n=1000/3 可调/5 请求 ≤10s；超限显式 `#VALUE!` |

---

## 11. 待决策

1. 模块名与拆分：`SolveUtils.bas` 单模块（推荐） vs Core/UDF 拆分（需治理例外）
2. 依赖审批：`SolveUtils → LinearUtils`（QR/矩阵）是否作为第二个合法依赖
3. v1 是否纳入 `rate`（源代码已含；建议放 v2）
4. 默认 `model`：`auto`（更准） vs `linear`（更快）
5. `max_starts` 默认 5 与规模上限（§6）是否接受
6. 是否在本期同时重建 `.xlam` 二进制

---

## 附：C# 源文档章节映射

| 源文档章节 | VBA 处置 |
|------------|----------|
| §3.1~3.4 使用体验/函数清单/数据约定/示例 | **沿用**（参数名不变，UDF 前缀改 `UDF_SOLVE_*`） |
| §3.5 输出布局契约 | **沿用逐字**（§4.3） |
| §3.6 规模上限 | **修订**为 §6（VBA 性能预算） |
| §4 目标项目落点（C# 文件清单） | 替换为 §4 + §8（VBA 文件与治理清单） |
| §5 实施任务（Task 0.1~5.2） | 替换为 §5（Phase 0~5）；Task 5.1 ASYNC 删除 |
| §6 全局约束 | 沿用并叠加 VBA 红线（Variant/CVErr/双路径/6 处同步） |
| §7 风险与对策 | 沿用并叠加 §9 VBA 特有风险 |
| §8 验收标准 | **修订**为 §10（性能项下调） |
