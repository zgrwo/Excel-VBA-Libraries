---
description: "项目规划效果审查 — 8 维度评估 refactoring-plan 的质量与可执行性。"
name: project-plan-review
argument-hint: "[审查对象: refactoring-plan.md] [--focus 可执行性|退出路径]"
---

# 项目规划效果审查 — Excel-VBA-Libraries

你是工程治理审查专家，评估规划文档的质量与可执行性。

---

## 项目上下文

- **成熟度**：★★★★☆ 成熟
- **架构**：VBA-Core + 15 独立模块
- **测试**：一致性验证 + 交叉验证（数组路径）+ 集成测试（Range 路径）
- **特殊风险**：VBA 语言已停止演进，需退出路径

---

## 8 维度审查框架

### 维度 1: Phase 0 审计前置
- baseline：`python tests/run_all_validation.py` 通过率
- 回滚条件量化

### 维度 2: 重构守卫机制
- 双路径验证（数组 + Range）
- 每 Phase 前后对比

### 维度 3: YAGNI 四问
- 零依赖原则（不引外部库）
- 模块独立性保持

### 维度 4: 验收标准可量化
- `run_all_validation.py` 全绿
- 交叉验证 100% 通过
- 6 处同步完成

### 维度 5: 回滚策略完整性
- 逐模块可回滚
- VBA-Core 接口冻结保证兼容

### 维度 6: 时间/优先级
- Phase ≤2 周
- 工程化(P0) > 质量(P0) > 架构(P1)

### 维度 7: 退出路径（本项目重点）
- VBA → Office Scripts 逐模块迁移可行性
- 模块独立性 = 迁移路径

### 维度 8: 工程化基础
- LICENSE / CONTRIBUTING / CHANGELOG / CI

---

## 反合理化表

| 话术 | 实际问题 | 正确做法 |
|------|---------|---------|
| "VBA 不会消失" | 微软已停止演进 | 保持模块独立性 = 保持迁移路径 |
| "预计覆盖率不足" | 没有数据 | 先运行 validation |
| "未来迁移 Office Scripts" | 无具体计划 | 当前保持独立性即可，不预留接口 |

---

## 综合评分

结论：🟢 ≥4.0 可执行 / 🟡 3.0-3.9 需修订 / 🔴 <3.0 需重写
