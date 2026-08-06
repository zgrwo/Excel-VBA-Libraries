Option Explicit

'==============================================================================
' Module:       PhyChemUtils
' Purpose:      Physical chemistry: molecular weight, unit conversion, gas laws
' Layer:        Science
' Dependencies: None (pure VBA, Scripting.Dictionary)
' Public:       31 functions/subs
'==============================================================================


' PhyChemUtils.bas — 理化计算模块
' 依赖: 无外部依赖（纯 VBA，使用 Scripting.Dictionary）
' 分子量解析: 递归下降，支持 ()/[]/{} 嵌套括号、水合物标记 (+/·/.)
'
' === Public VBA 函数 ===
'   MolecularWeight(formula)         — 化学式 → 摩尔质量 (g/mol)
'   ConvertMass(val, from, [to])     — 质量单位换算
'   ConvertVolume(val, from, [to])   — 体积单位换算
'   ConvertPressure(val, from, [to]) — 压力单位换算
'   ConvertTemperature(val, from, [to]) — 温度单位换算
'   ConvertStandard(volume_m3, pressure_Pa, temperature_K, molWeight) — 气体标准态体积与质量
'   MassToMoles(mass, molWeight)     — 质量(g) → 物质的量(mol)
'   MolesToMass(moles, molWeight)    — 物质的量(mol) → 质量(g)
'   DilutionSolve(c1, v1, c2, v2)   — C₁V₁=C₂V₂ 求解器（三缺一）
'   IdealGasLaw(pressure_Pa, volume_m3, moles, temperature_K) — PV=nRT 求解器（三缺一）
'   Density(mass, volume, density_val) — ρ=m/V 求解器（三缺一）
'   PercentYield(actual, theoretical)— (实际产量/理论产量)×100%
'   CompressFactorPR(pressure_Pa, temperature_K, Tc_K, Pc_Pa, omega) — 压缩因子 Z（Peng-Robinson 方程）
'   CylinderStdVolume(cylVolume_L, fillPressure_Pa, fillTemperature_K, gasName) — 钢瓶标态体积（已知充装压力）
'   CylinderStdVolumeFromMass(netWeight_kg, gasFormula) — 钢瓶标态体积（已知净重）
'
' === Public UDF ===
'   UDF_PC_MOLWEIGHT(formula)
'   UDF_PC_CONVERTMASS(val, from, [to])
'   UDF_PC_CONVERTVOLUME(val, from, [to])
'   UDF_PC_CONVERTPRESSURE(val, from, [to])
'   UDF_PC_CONVERTTEMPERATURE(val, from, [to])
'   UDF_PC_CONVERTSTANDARD(volume_m3, pressure_Pa, temperature_K, molWeight)
'   UDF_PC_MASSTOMOLES(mass, molWeight)
'   UDF_PC_MOLESTOMASS(moles, molWeight)
'   UDF_PC_DILUTION(c1, v1, c2, v2)
'   UDF_PC_IDEALGASLAW(pressure_Pa, volume_m3, moles, temperature_K)
'   UDF_PC_DENSITY(mass, volume, density_val)
'   UDF_PC_CONVERTMASS(val, fromUnit, [toUnit])
'   UDF_PC_YIELD(actual, theoretical)
'   UDF_PC_COMPRESS(pressure_Pa, temperature_K, Tc_K, Pc_Pa)
'   UDF_PC_ZFACTOR(P, T, gasName)
'   UDF_PC_STDVOLUME(cylVolume_L, fillPressure_Pa, fillTemperature_K, gasName)
'   UDF_PC_STDVOLMASS(netWeight_kg, gasFormula)
'
' 限制: 温度换算基于绝对零度偏移，不支持兰氏度 (Rankine)。

' ============================================================================
' Private 常量
' ============================================================================

Private Const ERR_INVALID_INPUT        As Long = vbObjectError + 1000
Private Const ERR_UNKNOWN_ELEMENT      As Long = vbObjectError + 1001
Private Const ERR_PARSE_MISSING_CLOSE  As Long = vbObjectError + 1002
Private Const ERR_PARSE_BRACKET_MISMATCH As Long = vbObjectError + 1003
Private Const ERR_INVALID_FORMULA      As Long = vbObjectError + 1004

Private Const ELEMENTS As String = _
    "H=1.008|He=4.003|Li=6.941|Be=9.012|B=10.81|C=12.01|N=14.01|O=16.00|F=19.00|Ne=20.18|" & _
    "Na=22.99|Mg=24.31|Al=26.98|Si=28.09|P=30.97|S=32.07|Cl=35.45|Ar=39.95|K=39.10|Ca=40.08|" & _
    "Sc=44.96|Ti=47.87|V=50.94|Cr=52.00|Mn=54.94|Fe=55.85|Co=58.93|Ni=58.69|Cu=63.55|Zn=65.39|" & _
    "Ga=69.72|Ge=72.61|As=74.92|Se=78.96|Br=79.90|Kr=83.80|Rb=85.47|Sr=87.62|Y=88.91|Zr=91.22|" & _
    "Nb=92.91|Mo=95.94|Tc=98.00|Ru=101.10|Rh=102.90|Pd=106.40|Ag=107.90|Cd=112.40|In=114.80|Sn=118.70|" & _
    "Sb=121.80|Te=127.60|I=126.90|Xe=131.30|Cs=132.90|Ba=137.30|La=138.90|Ce=140.10|Pr=140.90|Nd=144.20|" & _
    "Pm=145.00|Sm=150.40|Eu=151.90|Gd=157.30|Tb=158.90|Dy=162.50|Ho=164.90|Er=167.30|Tm=168.90|Yb=173.00|" & _
    "Lu=175.00|Hf=178.49|Ta=180.95|W=183.85|Re=186.21|Os=190.23|Ir=192.22|Pt=195.08|Au=196.97|Hg=201.97|" & _
    "Tl=204.38|Pb=207.20|Bi=208.98|Po=209.00|At=210.00|Rn=222.00|Fr=223.00|Ra=226.03|Ac=227.03|Th=232.04|" & _
    "Pa=231.04|U=238.03|Np=237.05|Pu=244.06|Am=243.06|Cm=247.07|Bk=247.07|Cf=251.08|Es=252.08|Fm=257.10|" & _
    "Md=258.10|No=259.10|Lr=266.00"

' ============================================================================
' Private 辅助函数
' ============================================================================

' 元素周期表字典缓存 — 首次调用构建，后续复用
Private Function GetElementDict() As Object
    Static elementDict As Object
    If Not elementDict Is Nothing Then
        Set GetElementDict = elementDict
        Exit Function
    End If

    Set elementDict = CreateObject("Scripting.Dictionary")
    Dim element As Variant
    For Each element In Split(ELEMENTS, "|")
        Dim keyValue() As String
        keyValue = Split(element, "=")
        elementDict.Add UCase$(keyValue(0)), CDbl(keyValue(1))
    Next element

    Set GetElementDict = elementDict
End Function

' 递归下降解析化学式 — 遇 '(' / '[' / '{' 递归求子式质量，遇 ')' / ']' / '}' 返回累积值
' pos 在调用间推进，返回从 pos 到匹配闭合括号（或公式末尾）的总质量
Private Function ParseFormulaRecursive(ByVal formula As String, _
    ByVal elementDict As Object, ByRef pos As Long) As Double
    Dim total As Double: total = 0#
    Dim n As Long: n = Len(formula)
    Dim ch As String, elem As String, closeChar As String
    Dim count As Long, subMass As Double

    Do While pos <= n
        ch = Mid$(formula, pos, 1)
        If ch = "(" Or ch = "[" Or ch = "{" Then
            ' 确定匹配的闭合字符
            Select Case ch
                Case "(": closeChar = ")"
                Case "[": closeChar = "]"
                Case "{": closeChar = "}"
            End Select
            pos = pos + 1
            subMass = ParseFormulaRecursive(formula, elementDict, pos)
            ' 期望匹配的闭合字符
            If pos > n Then
                Err.Raise ERR_PARSE_MISSING_CLOSE, "MolecularWeight", _
                    "缺少 '" & closeChar & "'"
            End If
            If Mid$(formula, pos, 1) <> closeChar Then
                Err.Raise ERR_PARSE_BRACKET_MISMATCH, "MolecularWeight", _
                    "期望 '" & closeChar & "'"
            End If
            pos = pos + 1
            ' 解析闭合括号后的数字乘数
            count = ParseNumberAt(formula, pos)
            If count = 0 Then count = 1
            total = total + subMass * count
        ElseIf ch = ")" Or ch = "]" Or ch = "}" Then
            ' 不推进 pos — 由外层调用者处理闭合括号匹配并推进
            ParseFormulaRecursive = total
            Exit Function
        ElseIf ch = "+" Or ch = ChrW$(&HB7) Or ch = "." Then
            ' 水合物/加合物连接符 (+ / · / .) — 系数仅作用于下一个顶层单元,
            ' 随后由循环继续累加后续单元 (避免多连接符时系数嵌套:
            ' 如 A·2B·3C = A + 2×B + 3×C, 而非 A + 2×(B + 3×C))
            pos = pos + 1
            count = ParseNumberAt(formula, pos)
            If count = 0 Then count = 1
            ' 下一单元为括号组或元素序列 — 继续循环解析并乘以系数
            Dim segMass As Double
            If pos <= n Then
                ch = Mid$(formula, pos, 1)
                If ch = "(" Or ch = "[" Or ch = "{" Then
                    Select Case ch
                        Case "(": closeChar = ")"
                        Case "[": closeChar = "]"
                        Case "{": closeChar = "}"
                    End Select
                    pos = pos + 1
                    segMass = ParseFormulaRecursive(formula, elementDict, pos)
                    If pos > n Then
                        Err.Raise ERR_PARSE_MISSING_CLOSE, "MolecularWeight", _
                            "缺少 '" & closeChar & "'"
                    End If
                    If Mid$(formula, pos, 1) <> closeChar Then
                        Err.Raise ERR_PARSE_BRACKET_MISMATCH, "MolecularWeight", _
                            "期望 '" & closeChar & "'"
                    End If
                    pos = pos + 1
                    Dim grpCount As Long: grpCount = ParseNumberAt(formula, pos)
                    If grpCount = 0 Then grpCount = 1
                    segMass = segMass * grpCount
                Else
                    ' 元素序列: 解析到下一个连接符或字符串末尾
                    Dim segStart As Long: segStart = pos
                    Do While pos <= n
                        ch = Mid$(formula, pos, 1)
                        If ch = "+" Or ch = ChrW$(&HB7) Or ch = "." Then Exit Do
                        pos = pos + 1
                    Loop
                    Dim segPos As Long: segPos = segStart
                    segMass = ParseFormulaRange(formula, elementDict, segPos, pos)
                End If
            Else
                segMass = 0#
            End If
            total = total + segMass * count
        ElseIf ch >= "A" And ch <= "Z" Then
            elem = ch
            pos = pos + 1
            If pos <= n Then
                ch = Mid$(formula, pos, 1)
                If ch >= "a" And ch <= "z" Then
                    elem = elem & ch
                    pos = pos + 1
                End If
            End If
            elem = UCase$(elem)  ' 字典键统一大写
            count = ParseNumberAt(formula, pos)
            If count = 0 Then count = 1
            If Not elementDict.Exists(elem) Then
                Err.Raise ERR_UNKNOWN_ELEMENT, "MolecularWeight", _
                    "未知元素: " & elem
            End If
            total = total + CDbl(elementDict(elem)) * count
        Else
            ' 非法字符 — 报错而非静默跳过 (#62)
            Err.Raise ERR_INVALID_FORMULA, "MolecularWeight", _
                "公式中包含非法字符 '" & Mid$(formula, pos, 1) & "' 位置 " & pos & "."
        End If
    Loop
    ParseFormulaRecursive = total
End Function

' 从 s 的 pos 位置解析连续数字字符，返回整数值并推进 pos
Private Function ParseNumberAt(ByVal s As String, ByRef pos As Long) As Long
    Dim nVal As Long: nVal = 0
    Dim ch As String
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch >= "0" And ch <= "9" Then
            nVal = nVal * 10 + CLng(ch)
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    ParseNumberAt = nVal
End Function

' 解析公式片段 [startPos, endPos) 的质量 (用于连接符后的元素序列单元,
' 片段内不含连接符; 括号组仍可嵌套)
Private Function ParseFormulaRange(ByVal formula As String, _
    ByVal elementDict As Object, ByRef startPos As Long, ByVal endPos As Long) As Double
    Dim seg As String
    seg = Mid$(formula, startPos, endPos - startPos)
    Dim segPos As Long: segPos = 1
    ParseFormulaRange = ParseFormulaRecursive(seg, elementDict, segPos)
End Function

' ============================================================================
' Public VBA 函数 — 分子量计算
' ============================================================================

Public Function MolecularWeight(ByVal formula As String) As Variant
    If formula = "" Then Err.Raise ERR_INVALID_FORMULA, "MolecularWeight", "公式不能为空。"
    Dim elementDict As Object: Set elementDict = GetElementDict()
    Dim pos As Long: pos = 1
    On Error GoTo EH
    MolecularWeight = ParseFormulaRecursive(formula, elementDict, pos)
    Exit Function
EH: Err.Raise Err.Number, "MolecularWeight", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 单位换算
' ============================================================================

' Volume unit conversion cache — built once, reused across calls
Private Function GetVolumeConversions() As Object
    Static dict As Object
    If Not dict Is Nothing Then Set GetVolumeConversions = dict: Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    With dict
        .Add "ul", 0.000001
        .Add "ml", 0.001
        .Add "cm3", 0.001
        .Add "l", 1
        .Add "gal", 3.78541
        .Add "galus", 3.78541
        .Add "galuk", 4.54609
        .Add "cuft", 28.3168
        .Add "cum", 1000
        .Add "m3", 1000
    End With
    Set GetVolumeConversions = dict
End Function

Public Function ConvertVolume(ByVal initialVolume As Double, ByVal fromUnit As String, _
                              Optional ByVal toUnit As String = "l") As Variant
    On Error GoTo ErrHandler

    Dim unitConversions As Object
    Set unitConversions = GetVolumeConversions()

    fromUnit = LCase$(fromUnit)
    toUnit = LCase$(toUnit)

    If unitConversions.exists(fromUnit) And unitConversions.exists(toUnit) Then
        ConvertVolume = initialVolume * unitConversions(fromUnit) / unitConversions(toUnit)
    Else
        Err.Raise ERR_INVALID_INPUT, "ConvertVolume", "不支持的单位: " & fromUnit & " 或 " & toUnit
    End If

    Set unitConversions = Nothing
    Exit Function

ErrHandler:
    Set unitConversions = Nothing
    Err.Raise Err.Number, "ConvertVolume", Err.Description
End Function

' Pressure unit conversion cache — built once, reused across calls
Private Function GetPressureConversions() As Object
    Static dict As Object
    If Not dict Is Nothing Then Set GetPressureConversions = dict: Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    With dict
        .Add "pa", 1#
        .Add "kpa", 1000#
        .Add "mpa", 1000000#
        .Add "psi", 6894.76
        .Add "bar", 100000#
        .Add "atm", 101325#
        .Add "kg/cm2", 98066.5
        .Add "torr", 133.3224
        .Add "mmhg", 133.3224
        .Add "cmh2o", 98.0665
    End With
    Set GetPressureConversions = dict
End Function

Public Function ConvertPressure(ByVal initialPressure As Double, ByVal fromUnit As String, _
                                Optional ByVal toUnit As String = "pa") As Variant
    On Error GoTo ErrHandler

    Dim unitConversions As Object
    Set unitConversions = GetPressureConversions()

    fromUnit = LCase$(fromUnit)
    toUnit = LCase$(toUnit)

    If unitConversions.exists(fromUnit) And unitConversions.exists(toUnit) Then
        ConvertPressure = initialPressure * unitConversions(fromUnit) / unitConversions(toUnit)
    Else
        Err.Raise ERR_INVALID_INPUT, "ConvertPressure", "不支持的单位: " & fromUnit & " 或 " & toUnit
    End If

    Set unitConversions = Nothing
    Exit Function

ErrHandler:
    Set unitConversions = Nothing
    Err.Raise Err.Number, "ConvertPressure", Err.Description
End Function

Public Function ConvertTemperature(ByVal initialTemperature As Double, ByVal fromUnit As String, _
                                   Optional ByVal toUnit As String = "k") As Variant
    On Error GoTo ErrHandler

    Dim convertedTemperature As Double
    fromUnit = LCase$(fromUnit)
    toUnit = LCase$(toUnit)

    ' 先统一转换为开尔文
    Select Case fromUnit
        Case "c": convertedTemperature = initialTemperature + 273.15
        Case "f": convertedTemperature = (initialTemperature + 459.67) / 1.8
        Case "k": convertedTemperature = initialTemperature
        Case Else: Err.Raise ERR_INVALID_INPUT, "ConvertTemperature", "不支持的温度单位: " & fromUnit
    End Select

    ' 从开尔文转换到目标单位
    Select Case toUnit
        Case "c": ConvertTemperature = convertedTemperature - 273.15
        Case "f": ConvertTemperature = convertedTemperature * 1.8 - 459.67
        Case "k": ConvertTemperature = convertedTemperature
        Case Else: Err.Raise ERR_INVALID_INPUT, "ConvertTemperature", "不支持的温度单位: " & toUnit
    End Select
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "ConvertTemperature", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 气体标准态换算
' ============================================================================

Public Function ConvertStandard(ByVal volume_m3 As Double, ByVal pressure_Pa As Double, _
                                ByVal temperature_K As Double, ByVal molWeight As Double) As Variant
    On Error GoTo ErrHandler

    Const GasConstant As Double = 8.314462618    ' 通用气体常数 J/(mol·K) (CODATA 2018)
    Const AtmosphericPressure As Double = 101325  ' 标准大气压 Pa

    If volume_m3 <= 0 Or pressure_Pa <= 0 Or temperature_K <= 0 Or molWeight <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "ConvertStandard", "所有参数必须为正数。"
    End If

    Dim numMoles As Double
    numMoles = (volume_m3 * pressure_Pa) / (GasConstant * temperature_K)

    Dim standardVolume As Double
    standardVolume = (numMoles * GasConstant * 273.15) / AtmosphericPressure

    Dim standardWeight As Double
    standardWeight = numMoles * molWeight * 0.001

    ConvertStandard = Array(standardVolume, standardWeight)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "ConvertStandard", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 质量与物质的量换算
' ============================================================================

' 单位约定: mass 单位 g, molWeight 单位 g/mol
Public Function MassToMoles(ByVal mass As Double, ByVal molWeight As Double) As Variant
    If mass < 0 Or molWeight <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "MassToMoles", "质量和摩尔质量必须为非负数。"
    End If
    MassToMoles = mass / molWeight
End Function

' 单位约定: moles 单位 mol, molWeight 单位 g/mol, 结果单位 g
Public Function MolesToMass(ByVal moles As Double, ByVal molWeight As Double) As Variant
    If moles < 0 Or molWeight <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "MolesToMass", "物质的量和摩尔质量必须为非负数。"
    End If
    MolesToMass = moles * molWeight
End Function

' ============================================================================
' Public VBA 函数 — 稀释计算 (C₁V₁ = C₂V₂)
' ============================================================================

' 参数中任意三项传入数值，待求项传入 Empty（或省略），返回解
' 例如：DilutionSolve(2#, 10#, Empty, Empty) → #VALUE!（缺三项，无法求解）
'       DilutionSolve(2#, 10#, Empty, 5#)    → 4（V₂=4 mL，即稀释到 4 mL）
'       DilutionSolve(2#, 10#, 0.5, Empty)   → 40（V₂=40 mL）
' 检测参数是否为待求解项 — Empty (VBA) 或 Null (COM None)
Private Function IsUnknown(ByVal v As Variant) As Boolean
    IsUnknown = IsEmpty(v) Or IsNull(v)
End Function

Public Function DilutionSolve(ByVal c1 As Variant, ByVal v1 As Variant, _
                              Optional ByVal c2 As Variant, _
                              Optional ByVal v2 As Variant) As Variant
    On Error GoTo ErrHandler

    ' 统计待求项（Empty 项数），必须恰好为 1
    Dim unknowns As Long
    unknowns = 0
    If IsUnknown(c1) Then unknowns = unknowns + 1
    If IsUnknown(v1) Then unknowns = unknowns + 1
    If IsUnknown(c2) Then unknowns = unknowns + 1
    If IsUnknown(v2) Then unknowns = unknowns + 1

    If unknowns <> 1 Then
        Err.Raise ERR_INVALID_INPUT, "DilutionSolve", _
            "必须恰好有一个未知参数 (传入 Empty)，当前未知数: " & unknowns & "。"
    End If

    ' 三缺一 → 直接求解
    If IsUnknown(c1) Then
        If CDbl(v1) <= 0 Then Err.Raise ERR_INVALID_INPUT, "DilutionSolve", "v1 必须为正数。"
        DilutionSolve = CDbl(c2) * CDbl(v2) / CDbl(v1)
    ElseIf IsUnknown(v1) Then
        If CDbl(c1) <= 0 Then Err.Raise ERR_INVALID_INPUT, "DilutionSolve", "c1 必须为正数。"
        DilutionSolve = CDbl(c2) * CDbl(v2) / CDbl(c1)
    ElseIf IsUnknown(c2) Then
        If CDbl(v2) <= 0 Then Err.Raise ERR_INVALID_INPUT, "DilutionSolve", "v2 必须为正数。"
        DilutionSolve = CDbl(c1) * CDbl(v1) / CDbl(v2)
    ElseIf IsUnknown(v2) Then
        If CDbl(c2) <= 0 Then Err.Raise ERR_INVALID_INPUT, "DilutionSolve", "c2 必须为正数。"
        DilutionSolve = CDbl(c1) * CDbl(v1) / CDbl(c2)
    End If
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "DilutionSolve", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 理想气体状态方程 (PV = nRT)
' ============================================================================

' 参数中任意三项传入数值，待求项传入 Empty，返回解
' R = 8.314 J/(mol·K)；P 单位 Pa，V 单位 m³，n 单位 mol，T 单位 K
' 例如：IdealGasLaw(101325#, 0.0245, Empty, 298#) → 1.0（约 1 mol）
Public Function IdealGasLaw(ByVal pressure_Pa As Variant, ByVal volume_m3 As Variant, _
                            ByVal moles As Variant, ByVal temperature_K As Variant) As Variant
    On Error GoTo ErrHandler

    Const R As Double = 8.314462618  ' J/(mol·K) (CODATA 2018)

    Dim unknowns As Long
    unknowns = 0
    If IsUnknown(pressure_Pa) Then unknowns = unknowns + 1
    If IsUnknown(volume_m3) Then unknowns = unknowns + 1
    If IsUnknown(moles) Then unknowns = unknowns + 1
    If IsUnknown(temperature_K) Then unknowns = unknowns + 1

    If unknowns <> 1 Then
        Err.Raise ERR_INVALID_INPUT, "IdealGasLaw", _
            "必须恰好有一个未知参数 (传入 Empty)，当前未知数: " & unknowns & "。"
    End If

    If IsUnknown(pressure_Pa) Then
        If CDbl(volume_m3) <= 0 Or CDbl(temperature_K) <= 0 Then Err.Raise ERR_INVALID_INPUT, "IdealGasLaw", "volume_m3 和 temperature_K 必须为正数。"
        IdealGasLaw = CDbl(moles) * R * CDbl(temperature_K) / CDbl(volume_m3)
    ElseIf IsUnknown(volume_m3) Then
        If CDbl(pressure_Pa) <= 0 Or CDbl(temperature_K) <= 0 Then Err.Raise ERR_INVALID_INPUT, "IdealGasLaw", "pressure_Pa 和 temperature_K 必须为正数。"
        IdealGasLaw = CDbl(moles) * R * CDbl(temperature_K) / CDbl(pressure_Pa)
    ElseIf IsUnknown(moles) Then
        If CDbl(temperature_K) <= 0 Then Err.Raise ERR_INVALID_INPUT, "IdealGasLaw", "temperature_K 必须为正数。"
        IdealGasLaw = CDbl(pressure_Pa) * CDbl(volume_m3) / (R * CDbl(temperature_K))
    ElseIf IsUnknown(temperature_K) Then
        If CDbl(moles) <= 0 Then Err.Raise ERR_INVALID_INPUT, "IdealGasLaw", "moles 必须为正数。"
        IdealGasLaw = CDbl(pressure_Pa) * CDbl(volume_m3) / (R * CDbl(moles))
    End If
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "IdealGasLaw", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 质量单位换算
' ============================================================================

' Mass unit conversion cache — built once, reused across calls
Private Function GetMassConversions() As Object
    Static dict As Object
    If Not dict Is Nothing Then Set GetMassConversions = dict: Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    With dict
        .Add "mg", 0.001
        .Add "g", 1
        .Add "kg", 1000#
        .Add "ton", 1000000#
        .Add "lb", 453.59237
        .Add "oz", 28.349523125
    End With
    Set GetMassConversions = dict
End Function

Public Function ConvertMass(ByVal initialMass As Double, ByVal fromUnit As String, _
                            Optional ByVal toUnit As String = "g") As Variant
    On Error GoTo ErrHandler

    Dim unitConversions As Object
    Set unitConversions = GetMassConversions()

    fromUnit = LCase$(fromUnit)
    toUnit = LCase$(toUnit)

    If unitConversions.exists(fromUnit) And unitConversions.exists(toUnit) Then
        ConvertMass = initialMass * unitConversions(fromUnit) / unitConversions(toUnit)
    Else
        Err.Raise ERR_INVALID_INPUT, "ConvertMass", "不支持的单位: " & fromUnit & " 或 " & toUnit
    End If

    Set unitConversions = Nothing
    Exit Function

ErrHandler:
    Set unitConversions = Nothing
    Err.Raise Err.Number, "ConvertMass", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 密度计算 (ρ = m / V)
' ============================================================================

' 参数中任意两项传入数值，待求项传入 Empty，返回解
' 单位约定: mass 单位 g, volume 单位 mL, density_val 单位 g/mL
' 例如：Density(10#, 2#, Empty) → 5（ρ = 10g / 2mL = 5 g/mL）
Public Function Density(ByVal mass As Variant, ByVal volume As Variant, _
                        Optional ByVal density_val As Variant) As Variant
    On Error GoTo ErrHandler

    Dim unknowns As Long
    unknowns = 0
    If IsUnknown(mass) Then unknowns = unknowns + 1
    If IsUnknown(volume) Then unknowns = unknowns + 1
    If IsUnknown(density_val) Then unknowns = unknowns + 1

    If unknowns <> 1 Then
        Err.Raise ERR_INVALID_INPUT, "Density", _
            "必须恰好有一个未知参数 (传入 Empty)，当前未知数: " & unknowns & "。"
    End If

    If IsUnknown(mass) Then
        If CDbl(volume) <= 0 Then Err.Raise ERR_INVALID_INPUT, "Density", "volume 必须为正数。"
        Density = CDbl(density_val) * CDbl(volume)
    ElseIf IsUnknown(volume) Then
        If CDbl(density_val) <= 0 Then Err.Raise ERR_INVALID_INPUT, "Density", "density_val 必须为正数。"
        Density = CDbl(mass) / CDbl(density_val)
    ElseIf IsUnknown(density_val) Then
        If CDbl(volume) <= 0 Then Err.Raise ERR_INVALID_INPUT, "Density", "volume 必须为正数。"
        Density = CDbl(mass) / CDbl(volume)
    End If
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "Density", Err.Description
End Function

' ============================================================================
' Public VBA 函数 — 产率计算
' ============================================================================

Public Function PercentYield(ByVal actual As Double, ByVal theoretical As Double) As Variant
    If theoretical <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "PercentYield", "理论产量必须为正数。"
    End If
    If actual < 0 Then
        Err.Raise ERR_INVALID_INPUT, "PercentYield", "实际产量不能为负数。"
    End If
    PercentYield = (actual / theoretical) * 100#
End Function

' ============================================================================
' Private 常量 — 常用气体的临界参数与偏心因子
' ============================================================================

' 返回气体的 Tc(K), Pc(Pa), ω(偏心因子)
' 数据来源: NIST Chemistry WebBook, CRC Handbook, API Technical Data Book
Private Function GetCriticalConstants(ByVal gasName As String, _
                                      ByRef outTc As Double, ByRef outPc As Double, _
                                      ByRef outOmega As Double) As Boolean
    Dim gas As String
    gas = LCase$(gasName)

    Select Case gas
        ' 惰性气体 / 载气
        Case "he", "helium":     outTc = 5.2:    outPc = 227000#:     outOmega = -0.39
        Case "ne", "neon":       outTc = 44.4:   outPc = 2760000#:    outOmega = -0.04
        Case "ar", "argon":      outTc = 150.9:  outPc = 4900000#:    outOmega = 0#
        Case "kr", "krypton":    outTc = 209.4:  outPc = 5500000#:    outOmega = 0#
        Case "xe", "xenon":      outTc = 289.7:  outPc = 5840000#:    outOmega = 0.01
        ' 双原子气体
        Case "h2", "hydrogen":   outTc = 33.2:   outPc = 1300000#:    outOmega = -0.22
        Case "n2", "nitrogen":   outTc = 126.2:  outPc = 3390000#:    outOmega = 0.037
        Case "o2", "oxygen":     outTc = 154.6:  outPc = 5040000#:    outOmega = 0.022
        Case "co", "carbonmonoxide": outTc = 132.9: outPc = 3500000#: outOmega = 0.049
        ' 电子特气 — 刻蚀/清洗
        Case "cf4":   outTc = 227.5: outPc = 3740000#: outOmega = 0.191
        Case "chf3":  outTc = 299.3: outPc = 4860000#: outOmega = 0.264
        Case "c2f6":  outTc = 293.0: outPc = 3040000#: outOmega = 0.255
        Case "c3f8":  outTc = 345.1: outPc = 2680000#: outOmega = 0.326
        Case "sf6":   outTc = 318.7: outPc = 3760000#: outOmega = 0.21
        Case "nf3":   outTc = 234.0: outPc = 4460000#: outOmega = 0.126
        Case "clf3":  outTc = 427.0: outPc = 5790000#: outOmega = 0.148
        ' 电子特气 — 沉积/掺杂
        Case "sih4":  outTc = 269.7: outPc = 4840000#: outOmega = 0.094
        Case "ph3":   outTc = 324.8: outPc = 6540000#: outOmega = 0.046
        Case "nh3":   outTc = 405.5: outPc = 1.135E+07: outOmega = 0.25
        Case "n2o":   outTc = 309.6: outPc = 7240000#: outOmega = 0.161
        ' 其他常见气体
        Case "co2":   outTc = 304.2: outPc = 7380000#: outOmega = 0.225
        Case "ch4":   outTc = 190.6: outPc = 4600000#: outOmega = 0.011
        Case "c2h6":  outTc = 305.3: outPc = 4870000#: outOmega = 0.099
        Case "c3h8":  outTc = 369.8: outPc = 4250000#: outOmega = 0.152
        Case Else
            GetCriticalConstants = False
            Exit Function
    End Select

    GetCriticalConstants = True
End Function

' ============================================================================
' Private 辅助 — Peng-Robinson 立方型方程求解
' ============================================================================

' Peng-Robinson EoS: P = RT/(V-b) - a·α(T)/(V²+2bV-b²)
' 化为 Z 的三次方程: Z³ - (1-B)Z² + (A-2B-3B²)Z - (AB-B²-B³) = 0
' 返回最大实根（气相压缩因子）
Private Function SolvePRCubic(ByVal A As Double, ByVal B As Double) As Double
    ' Z³ + c2·Z² + c1·Z + c0 = 0
    Dim c2 As Double, c1 As Double, c0 As Double
    c2 = B - 1#
    c1 = A - 2# * B - 3# * B * B
    c0 = -(A * B - B * B - B * B * B)

    ' Cardano: Z = y - c2/3 → y³ + p·y + q = 0
    Dim p As Double, q As Double
    p = c1 - c2 * c2 / 3#
    q = c0 - c2 * c1 / 3# + 2# * c2 * c2 * c2 / 27#

    Dim delta As Double
    delta = (q / 2#) * (q / 2#) + (p / 3#) * (p / 3#) * (p / 3#)
    ' 浮点修正：理论 delta=0 处可能产生极小负值，导致 Sqr 报错
    If delta < 0 And delta > -1E-14 Then delta = 0

    Dim y As Double
    If delta > 0 Then
        ' 一个实根（超临界区）
        Dim u As Double
        u = -q / 2# + Sqr(delta)
        If u < 0 Then y = -(-u) ^ (1# / 3#) Else y = u ^ (1# / 3#)
        Dim v As Double
        v = -q / 2# - Sqr(delta)
        If v < 0 Then y = y - (-v) ^ (1# / 3#) Else y = y + v ^ (1# / 3#)
    Else
        ' 三个实根 — 取最大根（气相）
        Dim phi As Double
        phi = Acos(-q / 2# / Sqr(-(p / 3#) * (p / 3#) * (p / 3#)))
        y = 2# * Sqr(-p / 3#) * Cos(phi / 3#)  ' 最大根
    End If

    SolvePRCubic = y - c2 / 3#
End Function

Private Function Acos(ByVal x As Double) As Double
    If x >= 1# Then Acos = 0#: Exit Function
    If x <= -1# Then Acos = 3.14159265358979: Exit Function
    Acos = Atn(-x / Sqr(-x * x + 1#)) + 2# * Atn(1#)
End Function

' ============================================================================
' Public VBA 函数 — 压缩因子与钢瓶标态体积
' ============================================================================

' Peng-Robinson 压缩因子计算
' pressure_Pa: 压力 (Pa), temperature_K: 温度 (K), Tc_K: 临界温度 (K), Pc_Pa: 临界压力 (Pa), omega: 偏心因子
' 误差: 非极性/弱极性气体 <2%, 近临界区 <5%
Public Function CompressFactorPR(ByVal pressure_Pa As Double, ByVal temperature_K As Double, _
                                  ByVal Tc_K As Double, ByVal Pc_Pa As Double, _
                                  ByVal omega As Double) As Variant
    Const R As Double = 8.314
    If pressure_Pa <= 0 Or temperature_K <= 0 Or Tc_K <= 0 Or Pc_Pa <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "CompressFactorPR", "压力和温度必须为正数。"
    End If

    ' PR 常数
    Dim a As Double, b As Double
    a = 0.457235528921382 * R * R * Tc_K * Tc_K / Pc_Pa
    b = 0.0777960739038885 * R * Tc_K / Pc_Pa

    ' α(T) = (1 + κ(1-√Tr))²
    Dim Tr As Double
    Tr = temperature_K / Tc_K
    Dim kappa As Double
    kappa = 0.37464 + 1.54226 * omega - 0.26992 * omega * omega
    Dim alpha As Double
    alpha = (1# + kappa * (1# - Sqr(Tr)))

    ' A, B 无纲量参数
    Dim Aparam As Double, Bparam As Double
    Aparam = a * alpha * alpha * pressure_Pa / (R * R * temperature_K * temperature_K)
    Bparam = b * pressure_Pa / (R * temperature_K)

    CompressFactorPR = SolvePRCubic(Aparam, Bparam)
End Function

' 钢瓶标态气体体积计算（Peng-Robinson）
' cylVolume_L: 钢瓶水容积 (L, 标准 47L), fillPressure_Pa: 充装压力 (Pa), fillTemperature_K: 充装温度 (K)
' gasName: 气体名称 (大小写不敏感, 如 "NF3", "CF4", "Nitrogen")
' 返回: 标准状态 (0°C, 101325 Pa) 下的气体体积 (m³)
'
' 示例: 47L NF₃ 钢瓶, 12MPa 充装, 25°C —
'   CylinderStdVolume(47, 12e6, 298.15, "NF3") → PR Z≈0.84, V_std≈5.8 m³
Public Function CylinderStdVolume(ByVal cylVolume_L As Double, ByVal fillPressure_Pa As Double, _
                                   ByVal fillTemperature_K As Double, ByVal gasName As String) As Variant
    Const T_STD As Double = 273.15
    Const P_STD As Double = 101325#

    If cylVolume_L <= 0 Or fillPressure_Pa <= 0 Or fillTemperature_K <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "CylinderStdVolume", "所有参数必须为正数。"
    End If

    Dim Tc As Double, Pc As Double, omega As Double
    If Not GetCriticalConstants(gasName, Tc, Pc, omega) Then
        Err.Raise ERR_INVALID_INPUT, "CylinderStdVolume", "未知气体: " & gasName
    End If

    On Error GoTo ErrHandler
    Dim Z As Variant
    Z = CompressFactorPR(fillPressure_Pa, fillTemperature_K, Tc, Pc, omega)
    CylinderStdVolume = (cylVolume_L / 1000#) * (fillPressure_Pa / P_STD) * (T_STD / fillTemperature_K) / CDbl(Z)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CylinderStdVolume", "压缩因子计算失败: " & Err.Description
End Function

' 钢瓶标态气体体积计算（已知净重，无需状态方程）
' netWeight_kg: 钢瓶内气体净重 (kg), gasFormula: 化学式 (如 "NF3", "CF4", "C2F6")
' 返回: 标准状态 (0°C, 101325 Pa) 下的气体体积 (m³)
'
' 原理: 净重 → 摩尔数 → 标态体积（理想气体摩尔体积 22.414 L/mol @ STP）
' 标态下 Z≈1，此方法无需压缩因子修正
'
' 示例: 47L NF₃ 钢瓶, 净重 30 kg →
'   CylinderStdVolumeFromMass(30, "NF3") → ~9.5 m³
Public Function CylinderStdVolumeFromMass(ByVal netWeight_kg As Double, _
                                           ByVal gasFormula As String) As Variant
    Const MOLAR_VOL_STD As Double = 22.414  ' L/mol @ 0°C, 1 atm

    If netWeight_kg <= 0 Then
        Err.Raise ERR_INVALID_INPUT, "CylinderStdVolumeFromMass", "净重必须为正数。"
    End If

    Dim mw As Variant
    mw = MolecularWeight(gasFormula)
    If IsError(mw) Then
        CylinderStdVolumeFromMass = mw
        Exit Function
    End If

    ' moles = netWeight_kg * 1000 / MW(g/mol)
    ' V_std(L) = moles * 22.414, 转为 m³ 除以 1000
    ' 化简: V_std(m³) = netWeight_kg * 22.414 / MW
    CylinderStdVolumeFromMass = netWeight_kg * MOLAR_VOL_STD / CDbl(mw)
End Function

' ============================================================================
' Public UDF Wrapper
' ============================================================================

Public Function UDF_PC_MOLWEIGHT(ByVal formula As Variant) As Variant
    On Error GoTo EH
    If IsObject(formula) Then If TypeName(formula) = "Range" Then formula = formula.Value
    UDF_PC_MOLWEIGHT = MolecularWeight(CStr(formula))
    Exit Function
EH:
    UDF_PC_MOLWEIGHT = CVErr(xlErrValue)
End Function

Public Function UDF_PC_CONVERTVOLUME(ByVal val As Variant, ByVal fromUnit As Variant, _
                                     Optional ByVal toUnit As Variant = "l") As Variant
    On Error GoTo EH
    UDF_PC_CONVERTVOLUME = ConvertVolume(CDbl(val), CStr(fromUnit), CStr(toUnit))
    Exit Function
EH:
    UDF_PC_CONVERTVOLUME = CVErr(xlErrValue)
End Function

Public Function UDF_PC_CONVERTPRESSURE(ByVal val As Variant, ByVal fromUnit As Variant, _
                                       Optional ByVal toUnit As Variant = "pa") As Variant
    On Error GoTo EH
    UDF_PC_CONVERTPRESSURE = ConvertPressure(CDbl(val), CStr(fromUnit), CStr(toUnit))
    Exit Function
EH:
    UDF_PC_CONVERTPRESSURE = CVErr(xlErrValue)
End Function

Public Function UDF_PC_CONVERTTEMPERATURE(ByVal val As Variant, ByVal fromUnit As Variant, _
                                          Optional ByVal toUnit As Variant = "k") As Variant
    On Error GoTo EH
    UDF_PC_CONVERTTEMPERATURE = ConvertTemperature(CDbl(val), CStr(fromUnit), CStr(toUnit))
    Exit Function
EH:
    UDF_PC_CONVERTTEMPERATURE = CVErr(xlErrValue)
End Function

Public Function UDF_PC_CONVERTSTANDARD(ByVal volume_m3 As Variant, ByVal pressure_Pa As Variant, _
                                       ByVal temperature_K As Variant, ByVal molWeight As Variant) As Variant
    On Error GoTo EH
    UDF_PC_CONVERTSTANDARD = ConvertStandard(CDbl(volume_m3), CDbl(pressure_Pa), CDbl(temperature_K), CDbl(molWeight))
    Exit Function
EH:
    UDF_PC_CONVERTSTANDARD = CVErr(xlErrValue)
End Function

Public Function UDF_PC_MASSTOMOLES(ByVal mass As Variant, ByVal molWeight As Variant) As Variant
    On Error GoTo EH
    UDF_PC_MASSTOMOLES = MassToMoles(CDbl(mass), CDbl(molWeight))
    Exit Function
EH:
    UDF_PC_MASSTOMOLES = CVErr(xlErrValue)
End Function

Public Function UDF_PC_MOLESTOMASS(ByVal moles As Variant, ByVal molWeight As Variant) As Variant
    On Error GoTo EH
    UDF_PC_MOLESTOMASS = MolesToMass(CDbl(moles), CDbl(molWeight))
    Exit Function
EH:
    UDF_PC_MOLESTOMASS = CVErr(xlErrValue)
End Function

Public Function UDF_PC_DILUTION(ByVal c1 As Variant, ByVal v1 As Variant, _
                                Optional ByVal c2 As Variant, _
                                Optional ByVal v2 As Variant) As Variant
    On Error GoTo EH
    UDF_PC_DILUTION = DilutionSolve(c1, v1, c2, v2)
    Exit Function
EH:
    UDF_PC_DILUTION = CVErr(xlErrValue)
End Function

Public Function UDF_PC_IDEALGASLAW(ByVal pressure_Pa As Variant, ByVal volume_m3 As Variant, _
                                   ByVal moles As Variant, ByVal temperature_K As Variant) As Variant
    On Error GoTo EH
    UDF_PC_IDEALGASLAW = IdealGasLaw(pressure_Pa, volume_m3, moles, temperature_K)
    Exit Function
EH:
    UDF_PC_IDEALGASLAW = CVErr(xlErrValue)
End Function

Public Function UDF_PC_CONVERTMASS(ByVal val As Variant, ByVal fromUnit As Variant, _
                                   Optional ByVal toUnit As Variant = "g") As Variant
    On Error GoTo EH
    UDF_PC_CONVERTMASS = ConvertMass(CDbl(val), CStr(fromUnit), CStr(toUnit))
    Exit Function
EH:
    UDF_PC_CONVERTMASS = CVErr(xlErrValue)
End Function

Public Function UDF_PC_DENSITY(ByVal mass As Variant, ByVal volume As Variant, _
                               Optional ByVal density_val As Variant) As Variant
    On Error GoTo EH
    ' 直接传递原始 Variant — Density 内部使用 IsUnknown 检测未知参数
    ' CDbl(Empty)=0 会破坏求解逻辑，导致 3 个参数全部"已知"而永远返回 CVErr
    UDF_PC_DENSITY = Density(mass, volume, density_val)
    Exit Function
EH:
    UDF_PC_DENSITY = CVErr(xlErrValue)
End Function

Public Function UDF_PC_YIELD(ByVal actual As Variant, ByVal theoretical As Variant) As Variant
    On Error GoTo EH
    UDF_PC_YIELD = PercentYield(CDbl(actual), CDbl(theoretical))
    Exit Function
EH:
    UDF_PC_YIELD = CVErr(xlErrValue)
End Function

Public Function UDF_PC_COMPRESS(ByVal pressure_Pa As Variant, ByVal temperature_K As Variant, _
                                ByVal Tc_K As Variant, ByVal Pc_Pa As Variant) As Variant
    On Error GoTo EH
    ' ω=0 适用于非极性简单气体 (N₂, O₂, Ar)；精确值请用 UDF_PC_ZFACTOR
    UDF_PC_COMPRESS = CompressFactorPR(CDbl(pressure_Pa), CDbl(temperature_K), CDbl(Tc_K), CDbl(Pc_Pa), 0#)
    Exit Function
EH:
    UDF_PC_COMPRESS = CVErr(xlErrValue)
End Function

' 气体名称查表版 — 自动获取 Tc/Pc/ω，适用于已知气体的精确计算
Public Function UDF_PC_ZFACTOR(ByVal pressure_Pa As Variant, ByVal temperature_K As Variant, _
                                ByVal gasName As Variant) As Variant
    On Error GoTo EH
    If IsObject(pressure_Pa) Then If TypeName(pressure_Pa) = "Range" Then pressure_Pa = pressure_Pa.Value
    If IsObject(temperature_K) Then If TypeName(temperature_K) = "Range" Then temperature_K = temperature_K.Value
    If IsObject(gasName) Then If TypeName(gasName) = "Range" Then gasName = gasName.Value
    Dim Tc As Double, Pc As Double, omega As Double
    If Not GetCriticalConstants(CStr(gasName), Tc, Pc, omega) Then
        UDF_PC_ZFACTOR = CVErr(xlErrValue)
        Exit Function
    End If
    UDF_PC_ZFACTOR = CompressFactorPR(CDbl(pressure_Pa), CDbl(temperature_K), Tc, Pc, omega)
    Exit Function
EH:
    UDF_PC_ZFACTOR = CVErr(xlErrValue)
End Function

Public Function UDF_PC_STDVOLUME(ByVal cylVolume_L As Variant, ByVal fillPressure_Pa As Variant, _
                                  ByVal fillTemperature_K As Variant, ByVal gasName As Variant) As Variant
    On Error GoTo EH
    UDF_PC_STDVOLUME = CylinderStdVolume(CDbl(cylVolume_L), CDbl(fillPressure_Pa), CDbl(fillTemperature_K), CStr(gasName))
    Exit Function
EH:
    UDF_PC_STDVOLUME = CVErr(xlErrValue)
End Function

Public Function UDF_PC_STDVOLMASS(ByVal netWeight_kg As Variant, _
                                   ByVal gasFormula As Variant) As Variant
    On Error GoTo EH
    UDF_PC_STDVOLMASS = CylinderStdVolumeFromMass(CDbl(netWeight_kg), CStr(gasFormula))
    Exit Function
EH:
    UDF_PC_STDVOLMASS = CVErr(xlErrValue)
End Function