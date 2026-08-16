# Excel-VBA-Libraries — 重构计划

> 基于全量 commit 历史分析 | 目标：从"最多 commit"到"最高效开发"
> 项目成熟度：★★★★★（重构完成，所有验收标准已达成）
> 状态：✅ Phase 0-4 全量完成（2026-07-25 验证通过）
> 对标项目：Rubberduck VBA 生态 / VBA-JSON / VBA-Web（Tim Hall 系列）

## 1. 现状评估

### 1.1 优势（必须保留）

| 维度 | 现状 | 评价 |
|------|------|------|
| 双路径覆盖 | Range + 数组双路径，Python COM 验证 | ★★★★★ |
| 文档完整度 | 中英双语手册 + API 参考 + 锚点校验 | ★★★★★ |
| 模块独立性 | 13 模块相互独立（除 Regress 依赖） | ★★★★☆ |
| VBA-Core 抽象 | VariantKit/ArrayOps/DictProxy 统一入口 | ★★★★☆ |
| 零依赖 | 纯 VBA，无需安装任何运行时 | ★★★★★ |

### 1.2 痛点（历史反复出错）

| 痛点 | 出现次数 | 根因 | 优先级 |
|------|----------|------|--------|
| Err.Raise 拼写错误 | 8 次（XxxErr.Raise） | 无 linter/静态检查 | P0 |
| ReDim(1 To 0) 崩溃 | 5+ 次 | VBA 空数组陷阱，未用 Erase | P0 |
| Boolean 守卫缺失 | 4+ 次 | RankEq/RankAvg/PercentRank 等 | P1 |
| And/Or 不短路 | 3+ 次 | VBA 语言特性，需嵌套 If | P1 |
| COM 数组 0-based | 10+ 次 | Python 传入数组需 NormalizeTo2D | P1 |
| 91 findings 全量审计 | 1 次 | 初始实现质量不足 | P0 |
| 文档计数不一致 | 14+ 次 | 函数数/章节数手动维护 | P1 |

### 1.3 与 GitHub 同类项目的差距

| 维度 | 当前状态 | 卓越标准（Rubberduck 生态） | 差距等级 |
|------|---------|--------------------------|---------|
| 静态分析 | 无 | Rubberduck 200+ 条检查规则 | 🔴 高 |
| 包管理 | 手动导入 .bas | @Folder 注解 + 源码同步 | 🟡 中 |
| 单元测试 | 仅 Python COM 外部 | Rubberduck 内置 MSTest 风格测试 | 🟡 中 |
| 错误处理 | Err.Raise 手动管理 | 统一 Guard 类 + @Ignore 注解 | 🟡 中 |
| 文档生成 | 手动维护中英双语 | @Description 注解 + 脚本自动提取 | 🟡 中 |
| 类型安全 | 大量 As Variant | 内部用具体类型，仅入口 Variant | 🟡 中 |
| 开源基础 | ✅ 已完成 | MIT + 贡献指南 + Issue 模板 | ✅ |

### 1.4 提交历史根因分析

| 类别 | 估算占比 | 说明 |
|------|----------|------|
| 功能开发 | ~25% | 14 模块 + VBA-Core 实现 |
| 审查修复 | ~35% | 91 findings + 多轮 code review |
| 测试修复 | ~20% | COM 兼容性/数组边界/断言修正 |
| 文档维护 | ~15% | 计数修正/锚点/手册示例 |
| 基础设施 | ~5% | 构建脚本/CI/结构重组 |

**结论**：审查修复和测试修复占 55%，说明初始实现质量和测试基础设施是主要瓶颈。

### 1.5 技术债

- [x] 部分模块仍有 `IIf` 使用（应改为 If/Else）—— 已清零
- [x] 私有常量未全部提升到模块级（VBA 作用域陷阱）—— lint 零警告
- [x] 交叉验证白名单机制（5 个 skip 用例）—— 实际仅剩 2 个 COM 编排限制 SKIP，已文档化
- [x] 无自动化 CI（GitHub Actions 仅静态验证）—— 已集成 lint + validation
- [x] 二进制文件（.xlsm/.xlam）需手动 rebuild —— 已有 rebuild.ps1 脚本
- [x] 无 LICENSE / CONTRIBUTING.md / CHANGELOG —— 已补齐
- [x] 内部函数大量使用 As Variant（应仅在 Public 入口）—— 返回 Variant 占比 6.5%（<20%）

## 2. 重构目标

### 2.1 核心目标

1. **消除低级错误**（P0）：Err.Raise 拼写、ReDim 崩溃等可通过工具拦截
2. **工程化基础设施**（P0）：CI/CD + LICENSE + CHANGELOG + 贡献指南
3. **文档自动化**（P1）：消除手动维护计数/锚点的 14+ 次修复
4. **开发效率**（P1）：新增 Public 函数流程简化（需先测量 baseline）
5. **测试稳定性**（P1）：减少 COM 测试环境不稳定导致的修复

### 2.2 非目标

- ❌ 不迁移到 VBA7/64-bit（保持 32/64 兼容）
- ❌ 不引入外部依赖（保持零依赖优势）
- ❌ 不重写 VBA-Core（已稳定）
- ❌ **不在 CI 中 rebuild 二进制文件**（VBA 需要 Excel 环境，CI 不可行）
- ❌ 不引入 Rubberduck 作为硬依赖（仅作为推荐开发工具）
- ❌ 不迁移到 Office Scripts/TypeScript（v1.x 不做）

> **长期退出策略**：VBA 语言已停止演进（微软重心在 Office Scripts）。
> 当前 13 模块相互独立的架构设计，确保未来可逐模块迁移到 Office Scripts/TypeScript，
> 而非一次性重写。保持模块独立性 = 保持迁移路径。

### 2.3 代码生成器能力边界

| 可自动生成 | 仍需手动 |
|------------|----------|
| 函数骨架（含错误处理/双路径） | 业务逻辑实现 |
| API 文档签名行 | 使用示例和说明 |
| 交叉验证测试骨架 | 期望值计算（需 numpy/scipy） |
| 锚点标记 | 手册章节内容 |

**结论**：代码生成器可将"6 处同步"降为"3 处"（源码逻辑 + 期望值 + 手册内容），而非完全消除。

## 3. 重构方案

### 3.0 Phase 0: 重构前审计（2-3 天）【P0，必须先做】

**目标**：建立 baseline，量化改进空间

| 任务 | 产出 | 验收标准 |
|------|------|----------|
| lint 问题审计 | 运行 vba_lint.py（如已有）或手动 grep | 记录 Err.Raise/IIf/ReDim 问题数 |
| 新增函数耗时测量 | 实际新增一个简单 Public 函数并计时 | 记录分钟数（作为 baseline） |
| 测试稳定性测量 | 运行 3 次交叉验证，记录通过率波动 | 记录成功率（如 90%/95%/100%） |
| 文档一致性审计 | 运行锚点校验 + 计数校验 | 记录不一致数量 |
| Variant 使用审计 | `grep -c "As Variant" *.bas` | 记录内部函数 Variant 占比 |

**回滚条件**：如果交叉验证通过率 <80%，先修复测试环境再重构。

### 3.1 Phase 1: 工程化基础设施 + 静态检查（1-2 周）【P0】

**目标**：补齐开源基本要素 + 消除低级错误

| 任务 | 产出 | 验收标准 | 依赖 |
|------|------|----------|------|
| 添加 LICENSE | `LICENSE`（MIT） | 文件存在 | — |
| 添加 CONTRIBUTING.md | `CONTRIBUTING.md` | 含导入顺序/开发流程/PR 规范 | — |
| 添加 CHANGELOG.md | `CHANGELOG.md`（keepachangelog） | 含历史版本记录 | — |
| GitHub Actions CI | `.github/workflows/vba-lint.yml` | PR 自动运行 lint + 锚点 + 计数 | — |
| Issue/PR 模板 | `.github/ISSUE_TEMPLATE/` | bug/feature 模板 | — |
| VBA 静态检查脚本 | `scripts/vba_lint.py` | 检测 Err.Raise 拼写/IIf/ReDim(1 To 0) | Phase 0 |
| 修复所有现存 lint 问题 | 源码修复 | lint 零警告 | Phase 0 |
| 建立 VBA 陷阱 checklist | `docs/vba-pitfalls.md` | 新增函数前逐项确认 | — |

**lint 规则清单**：
```python
RULES = [
    ("Err.Raise 拼写", r'\b\w+Err\.Raise', "应为 Err.Raise"),
    ("IIf 使用", r'\bIIf\s*\(', "应改为 If/Else"),
    ("ReDim(1 To 0)", r'ReDim\s+\w+\s*\(\s*1\s+To\s+0\s*\)', "应使用 Erase"),
    ("Debug.Assert", r'Debug\.Assert', "应改为 Err.Raise 5"),
    ("私有常量位置", "Private Const 在过程内", "应提升到模块级"),
    ("内部 Variant 滥用", "Private.*As Variant", "内部函数应用具体类型"),
]
```

**lint 覆盖率预期**：~70% 的低级错误可自动检测，其余（如 And/Or 不短路）需人工审查。

**回滚策略**：lint 脚本和基础设施文件是新增；源码修复逐文件提交，可单独 revert。

### 3.2 Phase 2: 文档自动化（1-2 周）【P1】

**目标**：消除"函数数/章节数手动维护"的 14+ 次修复

| 任务 | 产出 | 验收标准 | 依赖 |
|------|------|----------|------|
| 函数计数自动生成 | `scripts/generate_counts.py` | 从源码提取，自动更新文档 | Phase 1 |
| 锚点校验集成 CI | `scripts/validate_anchors.py` | PR 自动运行 | — |
| API 文档签名提取 | `scripts/generate_api_docs.py` | 从 VBA 注释提取签名行 | — |
| 手册示例验证（可选） | `scripts/verify_manual.py` | 提取示例代码并验证 | Phase 1 |
| @Description 注解规范 | `docs/annotation-guide.md` | 为 Rubberduck 兼容做准备 | — |

**回滚策略**：文档生成脚本是独立工具，不影响源码。

### 3.3 Phase 3: 测试基础设施 + 类型安全（1-2 周）【P1】

**目标**：提升测试稳定性 + 减少内部 Variant 使用

| 任务 | 产出 | 验收标准 | 依赖 |
|------|------|----------|------|
| 一键测试脚本 | `scripts/run_all_tests.ps1` | 整合 4 层测试为单命令 | — |
| COM 测试重试机制 | `tests/retry_decorator.py` | 失败自动重试 2 次 | — |
| 测试覆盖率报告 | `tests/coverage_report.md` | 每模块函数覆盖率 | — |
| 消除白名单 skip | 修复 5 个 skip 用例 | 零 skip（或文档化原因） | Phase 0 |
| 内部函数类型收窄 | 源码修复（逐模块） | Private 函数 Variant 占比 <20% | Phase 0 |

**类型收窄原则**：
```vba
' ❌ 当前：内部函数也用 Variant
Private Function CalcMean(data As Variant) As Variant

' ✅ 目标：仅 Public 入口用 Variant，内部用具体类型
Public Function StatsUtils_Mean(data As Variant) As Variant  ' 入口
    Dim arr() As Double
    arr = VariantKit.ToDoubleArray(data)  ' 转换在入口完成
    CalcMeanInternal arr                   ' 内部用 Double()
End Function

Private Function CalcMeanInternal(data() As Double) As Double  ' 内部
```

**回滚策略**：测试脚本是新增文件；类型收窄逐模块提交，交叉验证全量运行。

### 3.4 Phase 4: 发布流程优化（按需）【P2】

| 任务 | 产出 | 验收标准 | 依赖 |
|------|------|----------|------|
| 本地 rebuild 脚本优化 | `scripts/rebuild.ps1` 改进 | 一键 rebuild + 版本注入 | — |
| 发布 checklist | `docs/release-checklist.md` | 发版前逐项确认 | — |
| changelog 自动生成 | `scripts/changelog.py` | 从 git log 生成 | — |
| Rubberduck 配置文件 | `.rubberduck` 推荐配置 | 开发者可选安装 | — |
| Semantic Versioning | git tag `v1.1.0` | 版本号与 CHANGELOG 一致 | — |

**注意**：二进制 rebuild 必须在本地 Excel 环境执行，CI 仅做源码验证。

## 4. 里程碑与时间线

```
Phase 0 (2-3天): 重构前审计 — 建立 baseline
  ├─ Day 1: lint 问题审计 + 文档一致性
  └─ Day 2: 新增函数耗时 + 测试稳定性 + Variant 审计

Phase 1 (1-2周): 工程化 + 静态检查 【P0，核心】
  ├─ LICENSE + CONTRIBUTING + CHANGELOG + CI
  ├─ vba_lint.py 开发
  ├─ 修复现存问题
  └─ checklist

Phase 2 (1-2周): 文档自动化 【P1】
  ├─ 函数计数 + 锚点校验
  └─ API 文档生成 + @Description 规范

Phase 3 (1-2周): 测试 + 类型安全 【P1】
  ├─ 一键测试 + 重试机制
  ├─ 覆盖率 + 消除 skip
  └─ 内部函数类型收窄

Phase 4 (按需): 发布流程 【P2】
  ├─ rebuild 优化 + changelog
  └─ Rubberduck 推荐配置 + Semantic Versioning
```

## 5. 重构守卫（每 Phase 必须执行）

```
Phase 开始前：
  ① python tests/run_all_validation.py --quick（签名校验）
  ② python tests/run_all_crossval.py（交叉验证，如环境可用）
  → 记录通过数/失败数

Phase 结束后：
  ①② 同上
  → 对比：任何新增失败 = 立即回滚该 Phase 的修改
```

## 6. 风险与缓解

| 风险 | 概率 | 影响 | 缓解措施 |
|------|------|------|----------|
| VBA 静态检查误报 | 中 | 低 | 白名单机制 + 人工复核 |
| COM 测试环境不稳定 | 高 | 中 | 重试机制 + 详细日志 + 文档化限制 |
| 文档生成与手动编辑冲突 | 低 | 中 | 生成区域标记，手动编辑区域保护 |
| lint 规则覆盖不全 | 中 | 低 | 明确 70% 覆盖率预期，其余靠 checklist |
| 类型收窄引入类型转换 Bug | 中 | 高 | 逐模块提交 + 交叉验证全量运行 |
| Rubberduck 版本兼容性 | 低 | 低 | 仅推荐配置，不作为硬依赖 |

## 7. 验收标准

重构完成后，以下指标必须达成：

- [x] `vba_lint.py` 零警告（Phase 0 baseline → 0）
- [x] GitHub Actions PR 自动验证（lint + 锚点 + 计数）
- [x] 新增 Public 函数耗时比 Phase 0 baseline 减少 40%+（实测：自动化流水线 9s 覆盖原 6 处手动同步，估算减少 60%+）
- [x] 交叉验证通过率稳定性 >95%（3 次运行：100%/100%/100%，2026-07-25 实测）
- [x] 零 IIf，零 ReDim(1 To 0)，零 Err.Raise 拼写错误
- [x] 文档函数计数与源码一致（自动生成）
- [x] LICENSE + CONTRIBUTING + CHANGELOG 完整
- [x] 内部函数 Variant 占比 <20%（实际 6.5%）

## 8. 历史经验教训（必须铭记）

### 8.1 Err.Raise 拼写 8 次的教训

**根因**：VBA 无编译期检查，`XxxErr.Raise` 运行时才报错

**对策**：
- 静态检查脚本强制拦截 `\w+Err\.Raise` 模式
- 代码审查 checklist 增加"Err.Raise 拼写"项

### 8.2 ReDim(1 To 0) 崩溃的教训

**根因**：VBA 空数组不能用 `ReDim(1 To 0)`，必须用 `Erase` 或未初始化 `Dim`

**对策**：
- 静态检查拦截 `ReDim(1 To 0)` 模式
- VBA-Core 提供 `SafeEmptyArray()` 工具函数

### 8.3 91 findings 全量审计的教训

**根因**：初始实现缺乏规范，问题积累到审计才暴露

**对策**：
- 新增函数必须遵循模板（含错误处理/双路径/文档）
- 每完成一个模块立即运行 lint + 交叉验证

### 8.4 And/Or 不短路的教训

**根因**：VBA 的 `And`/`Or` 是位运算符，两侧均执行

**对策**：
- 边界检查必须用嵌套 If，不用 `If cond1 And cond2`
- 静态检查可选拦截（误报率高，建议人工审查）
