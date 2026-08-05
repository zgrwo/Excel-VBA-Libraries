---
name: vba
description: >
  Excel VBA Libraries development best practices and patterns.
  Use when writing, reviewing, or refactoring VBA code for Excel,
  especially for .xlam add-ins, .bas modules, worksheet functions,
  array/matrix/statistical algorithms, string/date processing,
  RegExp, file I/O, SQL/JSON, or dictionary/set structures.
  Also use for debugging type-safety, performance optimization,
  and ensuring cross-version Excel compatibility (32/64-bit, legacy vs dynamic arrays).
last_updated: 2026-06-17
---

# Excel VBA Libraries Development

Battle-tested patterns from the Excel VBA Libraries project.

## Core Rules Quick Reference

> 🔴 = 静默错误（不崩溃但结果错误）&nbsp;&nbsp;|&nbsp;&nbsp;多 § 引用时，第一个 § 为主规则。

| Scenario | Rule | § |
|----------|------|---|
| Writing UDFs | 🔴 All params `As Variant`; return `CVErr`, never raise; no `Volatile` | §1.3, §1.5 |
| Declaring variables | All `Dim` at top; 🔴 never `As New` | §2.2, §14.1 |
| Assigning objects | Always use `Set`, no exceptions; Dim and Set on separate lines | §3.2 |
| Conditionals | 🔴 `And`/`Or` don't short-circuit; nest guards | §3.3 |
| Arrays | 🔴 0-based internally, 1-based for Range; never `ReDim(m To n)` with m > n | §4 |
| String building | Array + `Join`; never `&` in loops | §5.4 |
| Dictionary usage | `DictProxy.Create()`; use `.Add` not `dict(key)=val` | §5.2, §14 |
| Parsing formulas/expressions | 🔴 Recursive descent or AST; never "evaluate then splice back" | §1.7 |
| Physical-quantity functions | 🔴 Parameter names MUST carry unit suffix | §1.8 |
| Excel Range | Multi-area check; merged-cell safe writes | §8 |
| Error handling | UDF → Template A; VBA function → Template B; resources → Template C | §16 |
| Private VBA-Core copies | 🔴 Forbidden; use module-level `Static` instance or `New` | §17 |

## 1. Core Principles

### 1.1 Function Purity & Parameter Passing
Never mutate input parameters. Work on local copies.

**ByVal vs ByRef conventions:**
- **UDF entry points**: all parameters `ByVal` — prevents accidental mutation of cell data.
- **Internal algorithm helpers**: large arrays `ByRef` (read-only) to avoid copy overhead.
- VBA default is `ByRef`; always write the intent explicitly.

```vba
' UDF: ByVal — safe, no side effects
Public Function ArraySlice(ByVal arr As Variant, ByVal start As Long) As Variant

' Internal sort: ByRef array for performance, treated as read-only
Private Sub QuickSort(ByRef arr() As Variant, ByVal lo As Long, ByVal hi As Long)
```

When refactoring a ByRef parameter that is being reassigned, introduce an explicit local variable (required by Option Explicit).

### 1.2 Guard Clauses
Check all error conditions at the top. Return immediately.
The happy path stays at minimal indentation.

```vba
Public Function ArraySlice(ByVal arr As Variant, ByVal start As Long, _
                           Optional ByVal length As Long) As Variant
    If Not IsArray(arr) Then ArraySlice = CVErr(xlErrValue): Exit Function
    If ArrayDims(arr) <> 1 Then ArraySlice = CVErr(xlErrValue): Exit Function
    If start < 0 Then ArraySlice = CVErr(xlErrValue): Exit Function
    ' ... happy path
End Function
```

### 1.3 Error Values for Worksheet Functions
Functions callable from a cell must return CVErr, never raise.
Reserve Err.Raise for internal VBA-to-VBA functions.
Match the error code to the semantic cause:

| Constant | Use when | Example |
|----------|----------|---------|
| `CVErr(xlErrValue)` | Invalid input | Empty array, wrong type |
| `CVErr(xlErrDiv0)` | Mathematically undefined | Zero variance, division by 0 |
| `CVErr(xlErrNum)` | Invalid parameter value | Negative argument to log |
| `CVErr(xlErrNA)` | Result not available / missing | Mode when all values are unique |

```vba
' Compact wrapper (project standard)
Public Function UDF_NAME(args) As Variant
    On Error GoTo EH: UDF_NAME = CoreImpl(args): Exit Function
EH: UDF_NAME = CVErr(xlErrValue)
End Function
```

#### 1.3.1 UDF Parameter Types — All `As Variant`

🔴 **Every UDF parameter MUST be declared `As Variant`.** Excel passes all
worksheet values through `Variant` — a user passing `=MyUDF(A1:A10)` or
`=MyUDF(42)` will cause a **`#VALUE!`** error if any parameter is typed
`As Long`, `As String`, `As Double`, etc.

```vb
' ❌ WRONG — =MyUDF(A1:A10) → #VALUE!
Public Function MyUDF(ByVal arr As Long) As Variant

' ✅ CORRECT
Public Function MyUDF(ByVal arr As Variant) As Variant
```

This also applies to `Optional` parameters: only `As Variant` supports `IsMissing()`
to detect whether the user omitted the argument.

### 1.4 Never Use Application.WorksheetFunction in a UDF
`Application.WorksheetFunction.VLookup(...)` throws a runtime error on failure,
killing the whole UDF. Use the late-bound Application method instead —
it returns an error variant that you test with IsError.

```vba
Dim res As Variant
res = Application.VLookup(val, rng, 2, False)   ' no "WorksheetFunction"
If IsError(res) Then
    MyFunc = res            ' pass through #N/A or replace with CVErr
    Exit Function
End If
```

### 1.5 Avoid Application.Volatile
Volatile functions recalc on every sheet change, destroying performance.
Only use Application.Volatile when the result genuinely depends on no input
(e.g. RAND(), NOW()) — and even then, document the reason.

### 1.6 Custom Error Codes
All project-specific internal errors use `vbObjectError + 1xxx` to avoid clashes
with built-in error numbers. (UDF functions must return `CVErr`, not raise — see §1.3.)
Error codes are declared as `Private Const` with a distinct range per module
(1000s, 1100s, 1200s...).

### 1.7 Parsing & Formula Engines: No String-Rebuild Loops

When parsing structured text (chemical formulas, JSON, math expressions),
**never** replace sub-strings with their numeric result and re-parse the
spliced string. This technique is fragile and causes silent data corruption
when delimiters get lost (e.g. `(CH3)2CO` → `45.082CO`).

Build an internal representation (AST, token list, or element-count dictionary)
with a dedicated parser and compute from that.

```vba
' WRONG: replacing sub-formulas with numbers in the original string
currentFormula = Left$(currentFormula, idx - 1) & subTotal & rest

' CORRECT: recursively build a parse structure, then evaluate
Private Function ParseFormula(ByVal s As String) As FormulaNode
    ' Returns a tree of Element(name, count) and Group(multiplier, children)
End Function
```

### 1.8 Physical-Quantity Functions: Encode Units in Parameter Names

Functions performing physical calculations (gas laws, unit conversions) **must**
carry the expected unit in the parameter name. Do not rely on comments alone.

```vba
' WRONG: comment says Pa but name doesn't
Public Function IdealGasLaw(ByVal P As Variant, ByVal V As Variant, ...

' CORRECT
Public Function IdealGasLaw_Pa_m3(ByVal pressure_Pa As Variant, _
                                  ByVal volume_m3 As Variant, ...
```

#### UDF Wrapper for Physical-Quantity Functions

Core functions use `Double` parameters with unit suffixes (`_Pa`, `_m3`, `_K`).
UDF wrappers receive `Variant` from Excel, validate with `IsNumeric`, then call
the core with `CDbl` coercion:

```vb
' UDF wrapper — unit suffixes carry through to UDF signature
Public Function UDF_PC_CONVERTVOLUME(ByVal val As Variant, _
    ByVal fromUnit As Variant, Optional ByVal toUnit As Variant) As Variant
    On Error GoTo EH
    ' CDbl will fail on non-numeric → caught by EH → CVErr
    UDF_PC_CONVERTVOLUME = ConvertVolume(CDbl(val), CStr(fromUnit), CStr(toUnit))
    Exit Function
EH: UDF_PC_CONVERTVOLUME = CVErr(xlErrValue)
End Function
```

Rules:
- 🔴 UDF parameter names MUST carry the same unit suffix as the core function
  (e.g. `ByVal pressure_Pa As Variant`, not `ByVal pressure As Variant`).
- Validating `IsNumeric` before `CDbl` prevents type mismatch on non-numeric cell input.
- String-type parameters (`fromUnit`) are coerced with `CStr` — safe since Excel
  always passes strings as Variant/String.

## 2. Naming, Declarations & Scope

### 2.1 Forbidden Names
VBA is case-insensitive. Avoid keywords, operators, built-in identifiers, and
shadowing of existing parameters.

| Category | Examples to avoid |
|----------|-------------------|
| Language keywords | `Sub`, `Function`, `Type`, `For`, `Stop`, `Set` |
| Built-in functions that shadow | `Val`, `Str`, `Date`, `Time`, `String`, `Error`, `Err` |
| Parameter shadowing | `Dim b() As Double` when param is `ByRef b() As …` |

```vba
' BAD                                 ' GOOD
Dim Sub() As Double                   Dim diff() As Double
Dim imp As Variant                    Dim fiResult As Variant
Public Sub Solve(ByRef b() As Double) Public Sub Solve(ByRef b() As Double)
    Dim B() As Double   ' shadow         Dim b2D() As Double
```

### 2.2 Variable Placement
All Dim are procedure-scoped. Never place them inside If or loops — it misleads
the reader into thinking they are block-scoped.

```vba
' GOOD: all Dim at top
Public Sub ProcessData()
    Dim i As Long, temp As Variant
    If condition Then temp = GetValue()
End Sub
```

### 2.3 Option Explicit Refactoring
After any refactoring that introduces a variable, ensure it has a Dim at the top.
Always run **Debug → Compile VBAProject** before committing.

### 2.4 Call Keyword — Forbidden
The `Call` keyword is pure syntactic noise. VBA supports direct invocation without it:

```vba
' WRONG
Call MySub(arg1, arg2)

' RIGHT — Subs: no Call, no parentheses
MySub arg1, arg2
```

For functions, always capture the return value — never `Call` a function.
Parenthesised arguments to a `Sub` without `Call` force ByVal evaluation,
which can cause "ByRef argument type mismatch" — another reason to drop `Call`.

### 2.5 Option Base — Forbidden
Never use `Option Base 1`. It changes the default lower bound of all module-level
arrays from 0 to 1, but `ReDim` always starts at exactly the bound you specify.
The result: arrays with inconsistent bases that silently break `LBound` assumptions.
Always use explicit `LBound`/`UBound` or `0 To N` in ReDim.

## 3. Variant & Data-Type Handling

### 3.1 Safe Comparisons and Sorting
When comparing or deduplicating Variants, you must handle all subtypes:
Empty, Null, Error, Boolean, Object, Array, numeric, and string.
Use a stable key generator (SafeKey) for dedup and a type-aware equality test
(ValuesEqual) for pairwise comparisons.

```vba
' Safe dedup key — built for Dictionary-based deduplication, not sorting
Private Function SafeKey(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeKey = "Null:##NULL##"
    ElseIf IsError(v) Then
        SafeKey = "Error:##ERROR##"
    ElseIf IsEmpty(v) Then
        SafeKey = "Empty:##EMPTY##"
    ElseIf IsObject(v) Then
        Dim objRef As Object: Set objRef = v
        If objRef Is Nothing Then
            SafeKey = "Object:Nothing:0"
        Else
            SafeKey = "Object:" & TypeName(v) & ":" & Hex$(ObjPtr(objRef))
        End If
    ElseIf IsArray(v) Then
        SafeKey = ArrayToKey(v)
    ElseIf IsNumeric(v) Then
        SafeKey = "Numeric:" & CStr(CDbl(v))
    Else
        SafeKey = TypeName(v) & ":" & CStr(v)
    End If
End Function

' Strict equality with epsilon for numbers
Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) And IsNull(b) Then ValuesEqual = True: Exit Function
    If IsNull(a) Or IsNull(b) Then Exit Function
    If IsEmpty(a) And IsEmpty(b) Then ValuesEqual = True: Exit Function
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    If IsError(a) And IsError(b) Then
        ValuesEqual = (CStr(a) = CStr(b)): Exit Function
    End If
    If IsError(a) Or IsError(b) Then Exit Function
    If VarType(a) = vbBoolean Or VarType(b) = vbBoolean Then
        If VarType(a) <> VarType(b) Then Exit Function
        ValuesEqual = (CBool(a) = CBool(b)): Exit Function
    End If
    If IsNumeric(a) And IsNumeric(b) Then
        ValuesEqual = (Abs(CDbl(a) - CDbl(b)) < 1E-12): Exit Function
    End If
    ValuesEqual = (CStr(a) = CStr(b))
End Function
```

**Semantic divergence by design:** `IsEmpty` is checked before `IsNumeric`, so `Empty`
and `0` produce distinct keys (`"Empty:##EMPTY##"` vs `"Numeric:0"`) and never merge
during dedup. However, `IsNumeric(True)=True`, so `True` and `-1` merge into the same
key (`"Numeric:-1"`). ValuesEqual is stricter: when both sides are Boolean, they match
only Booleans; when both are Empty, they match only Empties. Mixed `True` / `-1`
do NOT match — the Boolean VarType guard (line 387) exits early before reaching
the numeric tolerance path. This keeps cross-type values distinct, while dedup's
Compare path tolerates Boolean/numeric mixing for usability.

### 3.2 Object Assignment: Always Use Set

**Use `Set` for ALL object assignments, no exceptions.** Whether the target is `As Object`,
a typed class variable, or `As Variant` — always use `Set`. Relying on implicit default-property
evaluation when `Set` is omitted causes Error 450 on objects without a default property
(e.g. Dictionary) and silent bugs when the default property returns an unexpected type.

**Dim and Set must be on separate lines.** Never combine them with `:` on one line.

```vba
' CORRECT: separate lines, Set used
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

' CORRECT: Variant receiving object via Set
Dim v As Variant
Set v = CreateObject("Scripting.Dictionary")

' WRONG: Dim and Set on same line
Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

' WRONG: no Set on object assignment
Dim dict As Variant: dict = CreateObject("Scripting.Dictionary")
```

**Function return chain trap:** When a function returns `As Variant` and internally uses `Set`
to return an object, callers MUST also use `Set` to capture the result. Omitting `Set` at the
call site triggers default-property evaluation — Error 450 on objects without a default property.

**Variant assignment with Set/Let dispatch:**
When a Variant may contain an object OR a scalar/array, use the inline pattern.
It works everywhere — local variables, function returns, class members — with
zero restrictions:

```vba
' Universal pattern — 2 lines, no caveats
Dim tmp As Variant: tmp = source
If IsObject(tmp) Then Set target = tmp Else target = tmp
```

The legacy `VarLetSet` helper in VariantKit.cls is **deprecated for external callers**.
JsonUtils.bas retains its own private `VarLetSet` for function-return object dispatch
(no alternative in that specific context). For all other modules, use the inline pattern. It cannot be used
for function returns (Error 10/450), cannot be called cross-class, and requires
ByVal discipline for the source parameter. The inline pattern above has none of
these limitations.

**Object vs Variant declarations:** When a variable will ONLY hold object references,
declare it `As Object` rather than `As Variant`. This documents intent, enables
IntelliSense after `TypeOf` checks, and prevents accidental assignment of non-object
values. Reserve `As Variant` for variables that may hold mixed types (objects, arrays,
scalars, or `Null`/`Empty`/`Error`).

### 3.3 Common Type Traps

#### Control Flow (🔴 = silent bug)

- **VBA And/Or do not short-circuit** — all operands evaluate before combining:
  - **Type guards**: `If IsObject(x) Then If TypeOf x Is Range Then ...` (not `And`).
  - **Array bounds in loops**: `Do While j >= LBound(arr) And arr(j) > key` will evaluate
    `arr(j)` even when `j < LBound(arr)`, causing **subscript out of range (Error 9)**.
    Split into nested conditions:
    ```vba
    ' WRONG — arr(j) evaluated when j out of bounds
    Do While j >= low And arr(j) > key

    ' RIGHT — bounds check guards the access
    Do While j >= low
        If arr(j) <= key Then Exit Do
        arr(j + 1) = arr(j): j = j - 1
    Loop
    ```
    Same pattern applies to `vals(idx(j))`, `dict.Keys()(j)`, and any indexed access in a
    compound `And`/`Or` condition.
- **IIf evaluates all arguments**. Never use it with division, function calls, or side effects. Use If/Else.
- **Colon-separated Exit Function on a single-line If** — the colon binds only to the
  branch it appears in. `Exit Function` after `Else` does NOT apply to the `Then` branch:
  ```vba
  ' TRAP: Exit Function only fires in the Else branch
  If v Then StringifyValue = "true" Else StringifyValue = "false": Exit Function
  ' When v = True: assigns "true", then falls through — Exit Function is NOT reached
  ' FIX: move Exit Function after End If
  If v Then StringifyValue = "true" Else StringifyValue = "false"
  Exit Function
  ```
  This is a **silent correctness bug** — no compile error, no runtime error. The function
  continues past the `End If` and may hit downstream logic (e.g. `IsNumeric(True)=True`
  in VBA, overwriting the return value with `"-1"`). **Always put `Exit Function` on its
  own line after `End If`, never inside a single-line If/Else.**

#### Type Detection

- **IsObject(Nothing)** is unreliable on 64-bit. Always pair with `VarType(v) = vbObject`.
- **IsDate(2025) returns True**. When a parameter might be a date or a year number, check `VarType(v) = vbDate` first.
- **VarType constants are not sequential**. Use `Select Case` instead of range checks.
- **IsNumeric(Empty) = True** in VBA. Empty coerces to 0, so `IsNumeric` returns True
  and `CDbl(Empty) = 0`. When filtering non-numeric values, check `IsEmpty` before
  `IsNumeric` if Empty should be excluded. For dedup/mapping, decide whether Empty
  should merge with 0 (SafeKey) or remain distinct (ValuesEqual).

#### Conversion

- **Str$ vs CStr**: `Str$` adds a leading space for the sign of positive numbers
  (negative numbers get `-` at that position). Both use the system locale's decimal
  separator. For locale-independent output, use `Format$` with an explicit format
  string and replace the locale's decimal separator if needed.
- **CLng overflow — string length ≠ value magnitude**: A 10-character string like
  `"3000000000"` far exceeds `Long.MaxValue` (2,147,483,647). Always check against
  numeric boundaries, not string length:
  ```vba
  Dim dVal As Double: dVal = Val(numStr)
  If dVal >= -2147483648# And dVal <= 2147483647# Then
      result = CLng(dVal)
  Else
      result = dVal  ' preserve as Double
  End If
  ```

#### Arrays & Collections

- **`Dim arr() As Variant: arr = Array()` causes Type Mismatch (Error 13)**. `Array()`
  with no arguments returns Empty, not an empty array. Use `Erase arr` to clear, or
  `ReDim arr(0 To 0): Erase arr` for a properly empty typed array.
- **`Split` always returns a 0-based array**, regardless of `Option Base`. This is unlike
  most other VBA array sources (Range.Value is 1-based, RegExp split is 1-based).
- **`Split("")` returns `Array("")`, not an empty array**. `UBound(Split("", ","))` = 0,
  not -1. Guard with `If str = "" Then result = Array() Else result = Split(str, sep)`.
- **Optional non-Variant** parameters cannot use `IsMissing`. Declare optionals
  `As Variant` when you need to detect omission.

#### OOP & Modules

- **Class modules cannot call other Class methods without an instance**. Even if
  VariantKit defines `VarLetSet`, DictProxy cannot call it directly — Classes lack
  the `Implements`-free cross-referencing of standard modules. Use inline `IsObject`
  checks in Classes, or create a local VariantKit instance via `New`.

#### Data Integrity (🔴 = silent corruption)

- **Parser parts-array index consistency**: When building output via a parts array
  (e.g., JSON string decoding), every branch that writes to `parts(pIdx)` must
  increment `pIdx`. Missing increments cause the next character to overwrite the
  decoded value — a silent data corruption bug. Audit all `Select Case` / `If`
  branches: if a branch sets `parts(pIdx)`, it MUST also do `pIdx = pIdx + 1`.

## 4. Array Conventions & Safeguards

### 4.1 Base Consistency
Document the array base of every public function.

| Context | Convention | Reason |
|---------|-----------|--------|
| Internal algorithms | 0-based | Natural loops, easier indexing |
| Excel Range output | 1-based | Matches Range.Value |
| RegExp split results | 1-based | User expectation (1 = first) |

```vba
' Returns: 0-based Variant array (NOT suitable for direct Range write)
Public Function ArraySort(arr As Variant) As Variant
```

### 4.2 Creating Empty Arrays
**`ReDim(m To n)` where `m > n` always raises Error 9** (subscript out of range).
The most common real-world cases are `ReDim(0 To -1)` and `ReDim(1 To 0)` — both
occur naturally from expressions like `ReDim(0 To count - 1)` when `count = 0`.
Never rely on these; behaviour is host-dependent and unreliable across Excel versions.

| Goal | Safe method | Forbidden |
|------|------------|-----------|
| Return empty Variant() | `result = Array()` | `ReDim(0 To -1)`, `ReDim(1 To 0)` |
| Return empty Double()/etc. | `ReDim result(0 To 0): Erase result` | `result = Array()`, `ReDim(1 To 0)` |
| Empty after filtering | Guard size before ReDim | Unguarded `ReDim(0 To n-1)` |
| Test if empty | IsEmptyArray() probe (see §4.3) | Bare `UBound()` on possibly-uninit array |

```vba
' Safe guard for computed bounds
Dim size As Long: size = count - 1
If size < 0 Then result = Array() Else ReDim result(0 To size)
```

### 4.3 IsEmptyArray Probe
Handles unallocated, allocated-empty, and non-array inputs. Always Err.Clear between probes.

```vba
Private Function IsEmptyArray(ByRef v As Variant) As Boolean
    If Not IsArray(v) Then IsEmptyArray = False: Exit Function
    On Error Resume Next
    Dim lb As Long: lb = LBound(v)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: IsEmptyArray = True: Exit Function
    Err.Clear: On Error GoTo 0
    IsEmptyArray = (UBound(v) < lb)
End Function
```

## 5. Performance Patterns

### 5.1 Hybrid Sort
For n > 16: QuickSort. For n ≤ 16: InsertionSort (lower constant factor; threshold 16 is the
project convention).
Always recurse on the smaller partition first to keep stack depth O(log n).

**Pivot selection:** Use median-of-three (first, middle, last) for QuickSort pivot. Avoid using
the middle element alone — this degenerates to O(n²) on partially sorted or patterned data.

```vba
' Median-of-three pivot selection
Private Function MedianOfThree(ByRef arr As Variant, ByVal lo As Long, ByVal hi As Long) As Variant
    Dim mid As Long: mid = (lo + hi) \ 2
    If CompareCmp(arr(lo), arr(mid), "auto") > 0 Then Swap arr, lo, mid
    If CompareCmp(arr(lo), arr(hi), "auto") > 0 Then Swap arr, lo, hi
    If CompareCmp(arr(mid), arr(hi), "auto") > 0 Then Swap arr, mid, hi
    MedianOfThree = arr(mid)
End Function
```

### 5.2 Dictionary for Lookups
Use `Scripting.Dictionary` for grouping and deduplication instead of nested loops.
The project provides `DictProxy` (see §14) as the unified factory.

**Preferred:** Modules that already import DictProxy (via `Private DP As New DictProxy`) should use
`DP.Create()` for all dictionary creation. **Exception:** Modules without a DictProxy dependency
(e.g., PhyChemUtils, SqlUtils) may use direct `CreateObject("Scripting.Dictionary")` — adding a
DictProxy import solely for dictionary creation is unnecessary overhead for lightweight modules.

```vba
Dim dict As Object
Set dict = DictProxy.Create()   ' Unified, error-safe creation
```

#### 5.2.1 DictProxy — Standard Dictionary Factory

`DictProxy` wraps `CreateObject("Scripting.Dictionary")` with error handling and
configurable compare mode. Use it exclusively when creating new dictionaries.

> **Why CompareMode must be set at creation**: VBA Dictionary forbids changing
> `.CompareMode` after any item has been added (Error 5). `DictProxy.Create()` sets
> it on the empty dictionary before returning — you never need to touch `.CompareMode`
> directly.

```vba
' Default: vbTextCompare (case-insensitive)
Set dict = DictProxy.Create()
' Case-sensitive
Set dict = DictProxy.Create(vbBinaryCompare)
```

#### 5.2.2 Dictionary Object Storage — Use .Add, Not dict(key)=val

When `val` contains an object reference, `dict(key) = val` triggers the object's
default property (Error 450). Use `.Add` + `.Remove` instead:

```vba
' WRONG — fails on object-containing Variants
dict(key) = val
' CORRECT
If dict.Exists(key) Then dict.Remove key
dict.Add key, val
```

Critical for JSON parsers and any code building Dictionaries from heterogeneous input.

#### 5.2.3 Locale-Safe Dictionary Keys

`CStr()` on floating-point numbers produces locale-dependent output (e.g., `"1,5"` in
German locale, `"1.5"` in US). The same numeric value generates different dictionary keys
on different machines, causing silent lookup failures in cross-region deployments.

**Project approach:** Use `VariantKit.SafeKey` (§3.1) for dictionary keys. SafeKey
prefixes with a type tag (`"Numeric:1.5"`) which also prevents string/numeric key
collisions in mixed-type data.

```vba
' PREFERRED — type-disambiguated, consistent across locales
dict.Add SafeKey(1.5), value

' ACCEPTABLE — when type-tag prefix is undesirable and deployment is single-locale
dict.Add CStr(1.5), value

' NOTE: Format$(CDbl(v), "0.############") was explored as a locale-independent
' alternative but requires environment-specific verification (behavior varies
' across Excel versions). Test in target locale before adopting.
```

**DictProxy.FromKeys** internally uses `CStr()` for key conversion (verified
against 44 test assertions). If cross-region deployment is expected, test with
non-US locale settings. The locale risk is theoretical for single-locale
environments (the vast majority of Excel VBA deployments).

For ad-hoc dictionary keys in multi-locale deployments, prefer `SafeKey`.

### 5.3 Blocked Matrix Multiplication
For n > 64: tile into 32×32 blocks to fit L1 cache. Implementation: `LinearUtils.bas` → `MatrixMultiply`.

### 5.4 String Concatenation
VBA strings are immutable. `result = result & x` in a loop is O(n²).
Build an array and use Join:

```vba
Dim lb As Long: lb = LBound(items)
Dim n As Long: n = UBound(items) - lb + 1
Dim parts() As String: ReDim parts(0 To n - 1)
For i = 0 To n - 1: parts(i) = items(lb + i): Next
result = Join(parts, ",")
```

### 5.5 Minimize Application.Transpose
`Application.Transpose` converts everything to Variant and has a row limit equal to
the worksheet max (~1M rows). For large or typed arrays, write a dedicated transpose loop.

### 5.6 DoEvents in Loops — Avoid
`DoEvents` yields to the OS message pump, allowing Excel to repaint and respond
to user input. In a tight loop, it can 10x the runtime. If you need a progress
indicator, throttle it:

```vba
If i Mod 100 = 0 Then DoEvents: Application.StatusBar = "Processing " & i & " of " & n
```

Never call `DoEvents` on every iteration.

### 5.7 Application Performance Switches
Disable screen updates, calculation, and events during batch operations.
Always restore in an error-safe path:

```vba
Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False

On Error GoTo Cleanup
' ... batch work ...

Cleanup:
Application.Calculation = prevCalc
Application.ScreenUpdating = True
Application.EnableEvents = True
```

`Calculation` is restored to its previous value (not hardcoded `xlAutomatic`)
in case the user already had it set to `xlCalculationSemiautomatic`.

### 5.8 Dynamic Array Growth — Never ReDim Preserve Inside Loops

🔴 `ReDim Preserve` inside a loop is O(n²) — each call allocates a new array and copies
all existing elements. For wide datasets (e.g., Excel tables with many columns), this
dominates runtime.

**Pattern:** Pre-allocate at maximum possible size → track actual count → single
`ReDim Preserve` (or `Erase`) after the loop.

```vba
' WRONG — O(n²) on every match
For c = 1 To numCols
    If condition Then
        numFound = numFound + 1
        ReDim Preserve colMap(1 To numFound)   ' copies entire array each time
        colMap(numFound) = c
    End If
Next

' CORRECT — pre-allocate, truncate once
ReDim colMap(1 To numCols)
numFound = 0
For c = 1 To numCols
    If condition Then
        numFound = numFound + 1
        colMap(numFound) = c
    End If
Next
If numFound = 0 Then
    Erase colMap
ElseIf numFound < numCols Then
    ReDim Preserve colMap(1 To numFound)
End If
```

Exception: when the loop iteration count is trivially small (≤10) and the array size is
bounded, `ReDim Preserve` inside the loop is acceptable. Document the bound in a comment.

## 6. COM Object Management

- Reuse expensive objects (VBScript.RegExp, ADODB.Stream) — create once, process many.
- ⚠️ **VBScript.RegExp 不支持超时**：复杂正则（如含嵌套量词的回溯爆炸）可能导致 Excel 无响应。对于用户输入的动态正则模式，建议在文档中提示风险；对于内部使用的静态正则，确保模式经过测试且不含指数级回溯。
- Use `Static` with a Nothing guard for objects that survive across calls (FSO).
- `With` blocks reduce COM boundary crossings inside loops.
- Explicit release in reverse creation order, always in an error-handler path (Resume Cleanup).
- ADODB.Stream requires `.Close` before `= Nothing`.
- **When to release**: Must release for loop-created objects, Static/module-level references, and
  circular references (e.g. Excel.Range). Local procedure variables are auto-released on exit —
  explicit `= Nothing` there is optional insurance, not a requirement.
- **Early vs late binding**: This project uses late binding (`CreateObject`) exclusively.
  Early binding (`Tools → References`, then `Dim x As New SpecificType`) gives IntelliSense
  and ~2x faster method dispatch, but locks the project to a specific library version
  (e.g., "Microsoft Scripting Runtime 1.0"). Late binding works across all Windows versions
  without reference configuration. For a redistributable library, late binding is the correct
  default. Early binding is acceptable in end-user workbooks where the reference can be
  managed directly.

```vba
Public Function ReadFile(path As String) As Variant
    Dim fso As Object, stream As Object
    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = CreateObject("ADODB.Stream")
    ' ... processing ...
Cleanup:
    If Not stream Is Nothing Then
        On Error Resume Next: stream.Close: On Error GoTo 0
    End If
    Set stream = Nothing: Set fso = Nothing
    Exit Function
ErrHandler:
    ReadFile = CVErr(xlErrValue)
    Resume Cleanup
End Function
```

## 7. Numerical Stability

- **Kahan compensated summation** for summing many values (especially residuals in regression):
  ```vba
  Dim sum As Double, c As Double, y As Double, t As Double
  sum = 0#: c = 0#
  For i = 1 To n
      y = values(i) - c: t = sum + y: c = (t - sum) - y: sum = t
  Next
  ```
- **Two-pass variance** to avoid catastrophic cancellation.
- **Scale before norm**: divide by the maximum absolute element to prevent overflow.
- **Relative tolerance**: `Abs(pivot) < 1E-12 * columnNorm`, never compare to absolute zero.
- **Log-space for products**: compute geometric means as `Exp(SumOfLogs / n)`.

## 8. Excel Range Interaction

### 8.1 Multi-Area Ranges
`Range.Value` on a multi-area selection (`Areas.Count > 1`) raises Error 1004.
Always check and reject with a clear message.

```vba
If rng.Areas.Count > 1 Then
    Err.Raise vbObjectError + 1003, , "Multi-area ranges are not supported."
End If
```

### 8.2 Last Row
Use `Cells.Find` — handles hidden rows, filtered ranges, and sparse data correctly, unlike `End(xlUp)`.

```vba
Public Function LastRow(ws As Worksheet) As Long
    Dim rng As Range
    Set rng = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious)
    If rng Is Nothing Then LastRow = 0 Else LastRow = rng.Row
End Function
```

### 8.3 Normalise Range.Value
A single cell returns a scalar, a multi-cell returns a 2D array. Normalise:

```vba
If rng.Count = 1 Then
    ReDim data(1 To 1, 1 To 1): data(1,1) = rng.Value
Else
    data = rng.Value
End If
```

### 8.4 Safe Write-Back
- **Check HasFormula** before overwriting — `rng.Value = x` destroys formulas.
- **Merged cells**: writing to a merged area works only through the top-left anchor cell (`MergeArea.Cells(1, 1).Value = x`). Writing to a non-anchor cell raises Error 1004. Check `rng.MergeCells` and redirect writes to `rng.MergeArea.Cells(1, 1)` when true.
- **Array formulas**: partial overwrite of a multi-cell array formula (`Range.CurrentArray`) raises an error. Clear the whole array first or skip.
- **Legacy Excel**: If supporting Excel 2019 or earlier, users must array-enter your function with Ctrl+Shift+Enter. In Excel 365 dynamic arrays, results spill automatically. For empty results in legacy versions, return a single `CVErr(xlErrNA)` instead of `Array()`.

## 9. Input Validation

### 9.1 IsNumeric Guard Before CDbl
VBA's And/Or evaluate all operands. Always check IsNumeric before CDbl.
Fall back to StrComp for text.

```vba
If IsNumeric(a) And IsNumeric(b) Then
    result = (CDbl(a) < CDbl(b))
Else
    result = (StrComp(CStr(a), CStr(b), vbTextCompare) < 0)
End If
```

### 9.2 Err.Clear Before Probes
Stale errors cause false positives under On Error Resume Next.
Always Err.Clear before probing. Also, bare function calls as statements
(e.g. `UBound(data, 2)`) do not compile — capture the return value.

```vba
Dim dummy As Long
Err.Clear: On Error Resume Next
dummy = UBound(data, 2)
If Err.Number <> 0 Then ' truly 1D
```

### 9.3 On Error Resume Next — Scope
`On Error Resume Next` stays active for the **entire procedure** until explicitly
turned off with `On Error GoTo 0`. A common trap: opening it early for one probe
and forgetting to close it, masking every subsequent error in the function.

```vba
' TRAP — OERN still active for all following code
Err.Clear: On Error Resume Next
dummy = UBound(data, 2)
If Err.Number <> 0 Then is1D = True
' BUG: On Error GoTo 0 is MISSING — errors below are silently swallowed

' CORRECT — minimal scope
Err.Clear: On Error Resume Next
dummy = UBound(data, 2)
Dim hasDim2 As Boolean: hasDim2 = (Err.Number = 0)
On Error GoTo 0     ' restore immediately after the probe
If hasDim2 Then ...
```

Rule: the three lines (probe + check + `GoTo 0`) should be adjacent. Never leave
`On Error Resume Next` active across more code than the single probe it protects.

## 10. String & Encoding

For encoding-dependent operations, always provide a pure-VBA fallback:

```vba
Public Function Base64Encode(text As String) As String
    On Error GoTo Fallback
    ' ADODB.Stream fast path
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    ' ... stream.Charset = "UTF-8" ...
    Exit Function
Fallback:
    ' Pure VBA (always works, UTF-8 only)
    Base64Encode = EncodeBase64Bytes(EncodeUTF8Bytes(text))
End Function
```

**Base64 / URL / UTF-8 patterns:**
- ADODB.Stream fast path: `.Type = 2` (text), `.Charset = "UTF-8"`, `.ReadText`/`.WriteText`.
- Pure-VBA fallback: `EncodeBase64Bytes(EncodeUTF8Bytes(text))` — always works, UTF-8 only. (These are project built-in helpers; see `StringUtils.bas`.)
- URL encoding: iterate characters; unreserved pass through, others `%XX` hex-encoded.
- BOM handling: ADODB.Stream UTF-8 output includes BOM (EF BB BF); strip via binary detection.

**Byte array ↔ String conversion:**
- `StrConv(bytes, vbFromUnicode)` for ANSI → Unicode; `vbUnicode` for reverse.
- Always specify the code page explicitly; default behaviour is locale-dependent.

**Surrogate pairs:**
Characters outside the Basic Multilingual Plane (emoji, rare CJK) occupy two UTF-16 code units.
Detect them with `AscW(Mid$(text, i, 1))` and check `&HD800& <= cp <= &HDBFF&`.

## 11. Design Matrix & Regression

- **k-1 dummy encoding** to avoid perfect multicollinearity. Always store the full
  level list for prediction (especially the reference level). Drop the first
  categorical level; encode remaining k-1 levels as binary columns.

- **OLS via QR decomposition**, not normal equations, to preserve numerical condition.
  Normal equations (XᵀX)⁻¹Xᵀy square the condition number. QR solves Rβ = Qᵀy
  via back-substitution — numerically superior.

## 12. Project & File Organisation

List module dependencies in comments:

```
' RegressUtils.bas depends on: LinearUtils.bas, StatsUtils.bas
' SqlUtils.bas depends on: ADODB (Windows built-in)
' JsonUtils.bas — pure VBA, no external dependencies
```

Dependent modules must appear before their dependents in the VBA project.

- Mark internal functions `Private` — only `Public` functions become UDFs.
- Provide thin `Public` wrappers with CVErr for any function that needs cell exposure.
- Testing architecture and run commands → [AGENTS.md](../../../AGENTS.md).
- Module headers must list all public functions, dependencies, and limitations.

## 13. User Manual Authoring

完整手册撰写规范见 [manual-authoring.md](../vba-manual-authoring/SKILL.md)。

## 14. VBA-Core — Shared Infrastructure Classes

The `VBA-Core/` directory contains 3 reusable class modules (`.cls`), parallel to `src/`. Copy the entire directory to any new VBA project.

| Class | Purpose | Key Methods |
|----|------|---------|
| `VariantKit.cls` | Type normalization | `Normalize1D`, `Normalize2D`, `NormalizeTo2D`, `NormalizeInput`, `SafeKey`, `ValuesEqual`, `Compare`, `VarLetSet`, `IsEmptyArray`, `ArrayDims`, `Is1D`, `Is2D`, `IsNumericArray`, `IsNumericCell`, `FilterPasses`, `ToDoubles`, `WrapScalar` |
| `ArrayOps.cls` | Generic array operations | `Sort` (Hybrid: QuickSort > 16, InsertionSort <= 16), `SortIndices`, `Slice`, `IndexOf`, `Flatten`, `CollectNumericColumns` |
| `DictProxy.cls` | Safe Dictionary factory + batch ops | `Create`, `FromKeys`, `ToArray`, `Merge` |

### 14.1 Usage

```vba
Dim VK As VariantKit
Set VK = New VariantKit
Dim AO As ArrayOps
Set AO = New ArrayOps
Dim DP As DictProxy
Set DP = New DictProxy

arr = VK.Normalize1D(input, "R")       ' Any input to 1D array (row-major default)
GD.Require1D arr, "MyFunc"              ' Precondition check
AO.Sort arr, True, "auto"              ' Ascending, type-aware comparison
Set dict = DP.Create()                  ' Safe dictionary (vbTextCompare default)
```

> **Avoid procedure-level `As New`**: `Dim x As New ClassName` inside a procedure is auto-instantiating —
> VBA silently creates the object on first access, so `x Is Nothing` is always `False`. This hides
> initialization bugs (a `Nothing` guard never fires). Always use `Dim` + `Set = New` on separate lines inside procedures. See §3.2.
>
> **Module-level `Private ... As New` is the project standard for VBA-Core singletons.** At module
> level, auto-instantiation is the intended behavior — the object is created once on first use
> and reused for the module's lifetime. This pattern is used across all modules that depend on
> DictProxy (e.g., `Private DP As New DictProxy`). The `Is Nothing` trap does not apply because
> module-level instances are never conditionally initialized.

### 14.2 Technical Notes

- **Dependencies**: `VariantKit` is zero-dependency; others use `Object + New` for late binding
- **Format**: CRLF line endings, no BOM, standard `VERSION 1.0 CLASS` header
- **Import**: Use `VBComponents.Import(path)` (integrated in `test_utils.py`)
- **No Public Const**: Constants use `Public Property Get` (VBA class module limitation)
- **Migration**: Existing `IsNumericCell` Private functions remain deprecated; new code should prefer VBA-Core

## 15. Related Skills

其他 Skill 文件及文档的加载时机见 [AGENTS.md](../../../AGENTS.md) 文档路由表。

## 16. Error Handling Patterns

### 16.0 Decision Tree

```
Is this a worksheet function (UDF)?
├─ Yes → Template A — compact wrapper with CVErr
└─ No  → Does it acquire external resources?
         ├─ Yes → Template C — On Error GoTo Cleanup + finally block
         └─ No  → Template B — Inline validation + Err.Raise
```

### 16.1 Template A — UDF Compact Wrapper

Use for all `UDF_*` worksheet functions. The UDF body is a single-line delegate
to the core implementation, wrapped in `On Error GoTo` → `CVErr`.

```vb
Public Function UDF_XXX(ByVal arg1 As Variant, Optional ByVal arg2 As Variant) As Variant
    On Error GoTo EH: UDF_XXX = CoreImpl(arg1, arg2): Exit Function
EH: UDF_XXX = CVErr(xlErrValue)
End Function
```

Rules (see also §1.3, §1.4, §1.5):
- 🔴 Never `Err.Raise` — it kills the formula chain. Always return `CVErr`.
- 🔴 Never `Application.WorksheetFunction.*` — throws on failure. Use late-bound `Application.*` and `IsError`.
- 🔴 Never `Application.Volatile` — destroys caching.
- All parameters `As Variant` (see §1.3.1).
- Error variant choice: `xlErrValue` (invalid input), `xlErrNum` (invalid parameter), `xlErrNA` (not available), `xlErrDiv0` (math error).

### 16.2 Template B — Public VBA Function

Use for all non-UDF `Public Function` / `Public Sub`. Precondition checks use
inline validation with module-private error codes.
VBA's lack of interface inheritance makes a shared validation class impractical).

```vb
Public Function MyFunction(ByRef data As Variant, ByVal colIdx As Long) As Variant
    ' 1. Precondition validation — inline checks with module-private error codes
    If Not IsArray(data) Then
        Err.Raise ERR_INVALID_INPUT, "MyFunction", "需要数组, 实际为 " & TypeName(data) & "."
    End If
    Dim lb2 As Long
    On Error Resume Next: lb2 = LBound(data, 2): On Error GoTo 0
    If lb2 = 0 Then
        Err.Raise ERR_2D_REQUIRED, "MyFunction", "需要二维数组."
    End If

    ' 2. Business-logic errors — Err.Raise with module-private code
    If colIdx < 1 Or colIdx > UBound(data, 2) Then
        Err.Raise ERR_OUT_OF_BOUNDS, "MyFunction", _
            "colIdx (" & colIdx & ") out of range [1, " & UBound(data, 2) & "]."
    End If

    ' 3. Main logic
    ' ...
End Function
```

#### Error Code Convention

Error codes are `vbObjectError + offset`, **module-private**, defined at the
top of each `.bas` / `.cls` file:

```vb
Private Const ERR_INVALID_INPUT  As Long = vbObjectError + 1001
Private Const ERR_OUT_OF_BOUNDS  As Long = vbObjectError + 1002
Private Const ERR_NOT_AVAIL      As Long = vbObjectError + 1003
```

| Module layer | Recommended offset range |
|-------------|-------------------------|
| Any | `vbObjectError + 1001` is the universal `ERR_INVALID_INPUT` starting point |
| Data (Array, Dict, Pivot, SQL) | 1001–1099 |
| Math (Linear, Stats, Regress) | 1001–1199 (codes may span into custom ranges for domain errors) |
| Text (String, Regex, JSON) | 1001–1299 |
| Excel/File (Range, FileSystem) | 1001–1399 |
| Date (DateTime) | 1001–1499 |
| PhysChem | 1001–1599 |

Rules:
- Use `Private Const`, never `Public` — error codes are module-internal.
- The first code in every module is `ERR_INVALID_INPUT = vbObjectError + 1001` — this is the
  universal convention. Subsequent codes increment arbitrarily from there.
- Offset numbers don't need to be globally unique (modules are independent); the ranges above
  are guidelines, not strict boundaries. Domain-specific error codes (e.g., RegressUtils 3001+)
  may extend beyond them.
- Error `Source` parameter must match the calling function name (see §1.2).
- Error `Description` must be Chinese for user-facing, English for VBA-internal.

#### Err.Raise Message Format

```
"参数名: 中文描述"              ' input validation
"FuncName: 英文描述"             ' internal error (will be caught by caller)
```

### 16.3 Template C — Resource Cleanup (Finally Pattern)

Use when a function acquires external resources (COM objects, file handles,
ADODB connections). VBA has no `Finally` block — simulate with `GoTo Cleanup`.

```vb
Public Function ReadAndProcess(ByVal filePath As String) As Variant
    On Error GoTo ErrHandler
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(filePath, 1)
    ' ... process ...
    stream.Close
    Set stream = Nothing
    Set fso = Nothing
    Exit Function
ErrHandler:
    ' Cleanup — always execute, regardless of error
    If Not stream Is Nothing Then
        On Error Resume Next: stream.Close: On Error GoTo 0
    End If
    Set stream = Nothing
    Set fso = Nothing
    ' Re-raise or return error
    ReadAndProcess = CVErr(xlErrValue)
End Function
```

Rules (see also §6):
- 🔴 `Resume Next` inside cleanup ONLY — re-enable error handling immediately with `On Error GoTo 0`.
- Release in **reverse acquisition order** (stream → fso).
- Check `Is Nothing` before calling `.Close` — the error may have occurred before the `Set`.
- Local variables auto-release on scope exit — explicit `= Nothing` is insurance, not a requirement (see §6).
- If the function IS a UDF, wrap this pattern inside Template A's single-line delegate.

### 16.4 Anti-Patterns

| ❌ Anti-Pattern | Why it's bad | ✅ Use instead |
|----------------|-------------|----------------|
| `On Error Resume Next` as first line | 🔴 Silently swallows ALL errors | `On Error GoTo ErrHandler` |
| `On Error Resume Next` without `GoTo 0` | All subsequent lines ignore errors | Confine to single statement: `On Error Resume Next: riskyCall: On Error GoTo 0` |
| `Err.Raise` inside a UDF | Kills formula chain, user sees #VALUE! crash | `CVErr(xlErrValue)` |
| `Err.Clear` without reading `Err.Number` | Hides root cause, wastes debug time | Log or check `Err.Number` before clearing |
| `Resume` without label | Jumps to the error-causing line → infinite loop | `Resume Next` or `Resume Label` |

## 17. VBA-Core Code Reuse

### 17.1 Private Copies Are Forbidden

🔴 **Do not create private copies of VBA-Core functions in `.bas` modules.**
Use the class directly through a module-level instance:

```vb
' CORRECT: module-level Static instance
Private Function IsNumericCell(ByVal v As Variant) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    IsNumericCell = vk.IsNumericCell(v)
End Function

' WRONG: duplicating VBA-Core logic
Private Function NormalizeToArray(ByRef v As Variant) As Variant
    ' 20-line reimplementation of VariantKit.Normalize1D
End Function
```

### 17.2 When Thin Wrappers Are Acceptable

| Pattern | Acceptable? | Why |
|---------|------------|-----|
| `Private Sub VarLetSet(...)` in JsonUtils | ✅ Kept | Function-return object dispatch (no alternative) |
| `Private Function IsNumericCell(...)` → `vk.IsNumericCell(...)` | ✅ Kept | Static class instance required per module |
| Full reimplementation of VBA-Core logic | ❌ Forbidden | Use VBA-Core directly |

### 17.3 Detection

The validator (Layer 4) automatically flags violations:
```bash
python tests/utils/validate_docs.py --layer 4
```

Known-legitimate patterns are whitelisted in `tests/utils/validate_docs.py` — update the
`KNOWN_LEGITIMATE_WRAPPERS` set when adding new legitimate wrappers.
| Empty error handler (`ErrHandler:`) | 🔴 Error disappears without a trace | At minimum: log the error, set return value
