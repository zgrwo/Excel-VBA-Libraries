# VBA 陷阱清单 (VBA Pitfalls Checklist)

> 新增或修改 Public 函数前，逐项确认。历史提交中 55% 是修复提交（全量统计），本清单旨在将"事后修复"转为"事前拦截"。

## 必检项（每次提交前）

### 1. Err.Raise 拼写

- [ ] 错误抛出使用 `Err.Raise`，**不是** `XxxErr.Raise`
- [ ] 历史教训：8 次复制粘贴导致模块前缀混入

```vba
' ❌ 错误
StatsUtilsErr.Raise vbObjectError + 1, "StatsUtils_Mean", "..."

' ✅ 正确
Err.Raise vbObjectError + 1, "StatsUtils_Mean", "..."
```

### 2. ReDim 空数组

- [ ] 不使用 `ReDim arr(1 To 0)` 表示空数组
- [ ] 空数组使用 `Erase arr` 或保持未初始化状态

```vba
' ❌ 崩溃
ReDim result(1 To 0)

' ✅ 安全
Erase result
' 或: Dim result() As Variant  (未初始化即为空)
```

### 3. IIf 禁用

- [ ] 不使用 `IIf(cond, trueVal, falseVal)`
- [ ] 改用 `If/Else`（IIf 两侧均求值，可能触发溢出或 Null 传播）

```vba
' ❌ 危险：两侧均执行
x = IIf(n > 0, 1 / n, 0)  ' n=0 时仍计算 1/n → 溢出

' ✅ 安全
If n > 0 Then x = 1 / n Else x = 0
```

### 4. And/Or 不短路

- [ ] 边界检查不使用 `If cond1 And cond2`
- [ ] 改用嵌套 If（VBA 的 And/Or 是位运算，两侧均执行）

```vba
' ❌ arr 为空时 LBound(arr) 仍执行 → 崩溃
If IsArray(arr) And UBound(arr) >= 1 Then ...

' ✅ 安全
If IsArray(arr) Then
    If UBound(arr) >= 1 Then ...
End If
```

### 5. 参数类型

- [ ] 所有 Public UDF 参数声明为 `As Variant`
- [ ] 否则 Range 对象传入时返回 `#VALUE!`

```vba
' ❌ Range 传入 #VALUE!
Public Function MyFunc(data As Double) As Double

' ✅ 双路径支持
Public Function MyFunc(data As Variant) As Variant
```

### 6. 双路径处理

- [ ] Public 函数同时处理 Range 对象和 Variant 数组
- [ ] 优先使用 `VariantKit.NormalizeInput(v)` 统一入口

### 7. 错误处理模式

- [ ] UDF 入口 → `CVErr(xlErrValue)`
- [ ] VBA 内部函数 → `Err.Raise`
- [ ] 资源操作 → `On Error GoTo Cleanup` + 确保释放
- [ ] 禁止裸 `On Error Resume Next` 不检查 `Err.Number`

### 8. 数组索引

- [ ] 使用 `LBound`/`UBound` 获取边界，不硬编码
- [ ] COM 传入数组可能是 0-based → 使用 `NormalizeTo2D` 转 1-based
- [ ] 历史教训：10+ 次 Python COM 数组 0-based 问题

### 9. Debug.Assert

- [ ] 生产代码不使用 `Debug.Assert`
- [ ] 改用 `Err.Raise 5` 或条件检查

### 10. 私有常量位置

- [ ] `Private Const` 放在模块顶部，不放在过程内部
- [ ] VBA 作用域陷阱：过程内 Const 每次调用重新初始化

## 推荐项（代码审查时）

### 11. 内部函数类型

- [ ] Private 函数使用具体类型（`As Double`, `As Long`），不用 `As Variant`
- [ ] 仅 Public 入口接受 Variant，入口处完成类型转换

### 12. 文档同步

- [ ] 新增 Public 函数 → 同步 6 处（源码 + API Index + CN手册 + EN手册 + 交叉验证 + 锚点）
- [ ] 修改签名 → 更新 `rules/api-reference.md`

### 13. 模块独立性

- [ ] 不跨模块调用 Private 函数
- [ ] 仅 RegressUtils 可依赖 LinearUtils + StatsUtils
- [ ] 其余 13 模块相互独立

## 自动化拦截

| 工具 | 覆盖规则 | 命令 |
|------|----------|------|
| `scripts/vba_lint.py` | #1 #2 #3 #9 #10 #11 | `python scripts/vba_lint.py` |
| `tests/run_all_validation.py` | #12 (签名/计数一致性) | `python tests/run_all_validation.py` |
| `tests/utils/validate_manual_anchors.py` | #12 (锚点) | 已集成 CI |

> lint 覆盖率约 70%。#4 (And/Or 不短路) 误报率高，需人工审查。
