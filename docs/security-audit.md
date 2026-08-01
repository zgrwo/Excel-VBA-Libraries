# Security Audit — Excel-VBA-Libraries

<!-- last_updated: 2026-07-24 -->

> 本文档记录项目已实施的安全防护措施，供代码审查和安全评估参考。

## 1. SQL 注入防护 (SqlUtils)

| 防护措施 | 位置 | 说明 |
|----------|------|------|
| `SqlEscapeString` | SqlUtils.bas | 转义单引号、Error Variant、LIKE 通配符 |
| `EscapeSheetName` | SqlUtils.bas | 转义工作表名中的 `]` 字符，防止 SQL 注入 |
| 参数化查询 | SqlGetConnection | ADODB Connection 使用固定连接字符串，无用户输入拼接 |
| 表名验证 | SqlRangeQuery | 工作表名经 EscapeSheetName 处理后才拼入 SQL |

**已知限制**: ADODB 对 Excel 工作表的参数化查询支持有限，表名必须拼接（已转义）。

## 2. XXE 注入防护 (XmlUtils)

| 防护措施 | 位置 | 说明 |
|----------|------|------|
| `resolveExternals = False` | XmlUtils.bas L64 | 禁止解析外部实体 |
| `ProhibitDTD = True` | XmlUtils.bas L68 | MSXML 3.0/6.0 上禁止 DTD |
| MSXML3 回退 | XmlUtils.bas | MSXML6 不可用时回退到 MSXML3，同样设置安全属性 |

**已知限制**: MSXML 4.0 不支持 ProhibitDTD 属性，但项目优先使用 MSXML6/MSXML3。

## 3. XPath 注入 (XmlUtils)

| 风险 | 现状 | 缓解 |
|------|------|------|
| `UDF_XML_GET` 接受用户 XPath | 直接传给 `SelectSingleNode` | XPath 1.0 在 MSXML 中为只读查询，无法修改文档或执行代码 |
| 命名空间绕过 | 可选 namespaces 参数 | 用户需提供完整的命名空间映射 |

**风险评估**: **低**。XPath 1.0 在 MSXML 沙箱内执行，无文件系统访问能力。

## 4. 正则表达式回溯 (RegexUtils)

| 防护措施 | 位置 | 说明 |
|----------|------|------|
| `MAX_PATTERN_LENGTH = 256` | RegexUtils.bas L61 | 限制正则模式最大长度 |
| 模式验证 | RegexUtils.bas L81-93 | 超长模式直接报错，不执行 |
| 错误处理 | RegexUtils.bas | 无效正则表达式捕获并 Err.Raise |

**已知限制**: VBScript.RegExp 不支持超时机制。恶意用户仍可构造 256 字符内的灾难性回溯模式（如 `(a+)+$`）。

**缓解建议**: UDF 路径中用户输入的正则模式应视为不可信，建议：
- 工作表 UDF 仅允许预定义模式（白名单）
- VBA 调用路径由开发者自行控制模式复杂度

## 5. 路径遍历防护 (FileSystemUtils)

| 防护措施 | 位置 | 说明 |
|----------|------|------|
| `..` 检测 | FileSystemUtils.bas L1098-1111 | 拒绝包含 `..` 的路径 |
| 路径规范化 | FileSystemUtils.bas L1104 | 规范化后二次检查 `..` |
| 绝对路径要求 | FileSystemUtils.bas | 批量操作要求绝对路径 |

**已知限制**: 仅防御 `..` 遍历，不限制可访问的根目录范围。

## 6. 宏安全

| 项目 | 状态 |
|------|------|
| 数字签名指引 | README.md 中说明 |
| 宏启用风险提示 | README.md 中说明 |
| 无硬编码凭据 | 已验证（全项目 grep） |
| 无硬编码路径 | 已验证（示例使用相对路径） |

## 7. 安全审查检查清单

代码审查时，针对安全维度逐项确认：

- [x] SQL 注入: SqlEscapeString + EscapeSheetName 覆盖所有用户输入
- [x] XXE 注入: resolveExternals=False + ProhibitDTD=True
- [x] XPath 注入: XPath 1.0 只读，风险低
- [x] 正则回溯: MAX_PATTERN_LENGTH=256 限制
- [x] 路径遍历: `..` 检测 + 路径规范化
- [x] 宏安全: 文档告知 + 无硬编码凭据
- [ ] 依赖更新: 零外部依赖，无需跟踪 CVE
