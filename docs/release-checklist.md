# 发布清单 (Release Checklist)

> 每次发版前逐项确认。版本号遵循 [Semantic Versioning](https://semver.org/)。

## 发版前

### 代码质量

- [ ] `python scripts/vba_lint.py` — 零警告
- [ ] `python tests/run_all_validation.py` — 全量通过
- [ ] `python tests/run_all_crossval.py` — 交叉验证通过（需 Excel）
- [ ] `python tests/utils/integration_test_all_modules.py` — 集成测试通过（需 Excel）
- [ ] 无未解决的 `TODO` / `FIXME` 标记

### 文档一致性

- [ ] `rules/api-reference.md` 签名与源码一致
- [ ] `rules/user-manual.md` 示例可运行
- [ ] `docs/VBA_LIB_User_Manual_EN.md` 英文版同步
- [ ] 函数计数正确（`python scripts/generate_counts.py`）

### 版本号

- [ ] 确定版本号（MAJOR.MINOR.PATCH）
  - MAJOR：不兼容的 API 变更
  - MINOR：新增 Public 函数（向后兼容）
  - PATCH：Bug 修复（无 API 变更）
- [ ] CHANGELOG.md 已更新（keepachangelog 格式）
- [ ] 版本号与 CHANGELOG 一致

### 二进制重建

- [ ] `.\scripts\rebuild.ps1 -Version x.y.z` — 重建 .xlsm/.xlam
- [ ] 在 Excel 中手动验证 .xlsm 可正常加载
- [ ] 在 Excel 中手动验证 .xlam 加载项可正常启用

## 发版

- [ ] Git tag: `git tag -a vx.y.z -m "Release vx.y.z"`
- [ ] GitHub Release 创建（附 .xlsm/.xlam 下载）
- [ ] Release Notes 包含：新增函数、修复、已知问题

## 发版后

- [ ] 验证 Release 页面链接可访问
- [ ] 通知相关用户（如适用）
- [ ] 更新 README.md 中的版本号引用（如有）

## 回滚策略

如果发版后发现严重问题：
1. 立即在 GitHub 标记该 Release 为 Pre-release
2. 创建 hotfix 分支修复
3. 发布 PATCH 版本（vx.y.z+1）
4. 不删除已发布的 tag（保持可追溯）
