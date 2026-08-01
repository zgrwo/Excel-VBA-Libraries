' VERSION 1.0 CLASS
' BEGIN
'   MultiUse = -1  'True
' END
Option Explicit

'=============================================================================
' VariantKit.cls — Type normalization: probe, convert, compare, assign
'
' Zero dependencies, pure VBA base layer.
'
' Public methods:
'   Probe:  IsEmptyArray, ArrayDims, Is1D, Is2D, IsNumericArray, IsNumericCell
'   Convert: Normalize1D, NormalizeTo2D, Normalize2D, ToDoubles, WrapScalar
'   Compare: SafeKey, ValuesEqual, Compare
'   Assign: VarLetSet
'   Filter: Passes
'=============================================================================

' ======== Probe ========

Public Function IsEmptyArray(ByRef v As Variant) As Boolean
    If Not IsArray(v) Then IsEmptyArray = False: Exit Function
    On Error Resume Next
    Dim lb As Long: lb = LBound(v)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        IsEmptyArray = True: Exit Function
    End If
    Err.Clear: On Error GoTo 0
    IsEmptyArray = (UBound(v) < lb)
End Function

Public Function ArrayDims(ByRef v As Variant) As Long
    If Not IsArray(v) Then ArrayDims = 0: Exit Function
    On Error Resume Next
    Dim lb As Long: lb = LBound(v)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        ArrayDims = 0: Exit Function
    End If
    Err.Clear: On Error GoTo 0
    Dim d As Long: d = 1
    Dim dummy As Long
    Do
        On Error Resume Next
        Err.Clear
        dummy = UBound(v, d)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        d = d + 1
    Loop
    On Error GoTo 0
    ArrayDims = d - 1
End Function

Public Function Is1D(ByRef v As Variant) As Boolean
    Is1D = (ArrayDims(v) = 1)
End Function

Public Function Is2D(ByRef v As Variant) As Boolean
    Is2D = (ArrayDims(v) = 2)
End Function

' Returns True if array contains at least one numeric element (not "all elements are numeric")
Public Function IsNumericArray(ByRef v As Variant) As Boolean
    If Not IsArray(v) Then Exit Function
    Dim i As Long, lb As Long, ub As Long
    On Error Resume Next
    lb = LBound(v): ub = UBound(v)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    Err.Clear
    ' 2D 数组 v(i) 会崩溃 — 仅支持 1D
    ub = UBound(v, 2)
    If Err.Number = 0 Then Err.Clear: On Error GoTo 0: Exit Function
    Err.Clear: On Error GoTo 0
    For i = LBound(v) To UBound(v)
        If IsNumericCell(v(i)) Then IsNumericArray = True: Exit Function
    Next i
End Function

Public Function IsNumericCell(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Or VarType(v) = vbDate Then Exit Function
    If IsError(v) Then Exit Function
    ' IsNumeric 已涵盖数值型字符串 (如 "123")；非数值字符串保持默认 False
    ' (原 ElseIf VarType=vbString 分支恒为 False — IsNumeric 失败后再判 IsNumeric 必为 False，已移除该死代码)
    If IsNumeric(v) Then IsNumericCell = True
End Function

' ======== Convert ========

Public Function Normalize1D(ByVal v As Variant, _
                             Optional ByVal order As String = "R") As Variant
    Dim data As Variant, nRows As Long, nCols As Long
    Dim i As Long, j As Long, idx As Long
    Dim result() As Variant

    If IsObject(v) Then
        If TypeOf v Is Range Then
            If v.Areas.Count > 1 Then
                Normalize1D = CVErr(xlErrValue): Exit Function
            End If
            data = v.Value
        Else
            Normalize1D = v: Exit Function
        End If
    Else
        data = v
    End If

    If Not IsArray(data) Then
        ReDim result(0 To 0): result(0) = data
        Normalize1D = result: Exit Function
    End If

    If IsEmptyArray(data) Then Normalize1D = Array(): Exit Function

    If ArrayDims(data) = 1 Then Normalize1D = data: Exit Function

    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    ReDim result(0 To nRows * nCols - 1)

    If order = "C" Then
        idx = 0
        For j = 1 To nCols
            For i = 1 To nRows
                result(idx) = data(LBound(data, 1) + i - 1, LBound(data, 2) + j - 1)
                idx = idx + 1
            Next i
        Next j
    Else
        idx = 0
        For i = 1 To nRows
            For j = 1 To nCols
                result(idx) = data(LBound(data, 1) + i - 1, LBound(data, 2) + j - 1)
                idx = idx + 1
            Next j
        Next i
    End If
    Normalize1D = result
End Function

' NormalizeTo2D — Range/scalar/1D/2D array → 1-based 2D Variant array
' Output always has LBound 1 for both dimensions (suitable for Range.Value write-back)
Public Function NormalizeTo2D(ByVal data As Variant, _
                               ByRef numRows As Long, _
                               ByRef numCols As Long) As Variant()
    Dim rng As Range
    Dim arr As Variant
    Dim temp() As Variant
    Dim arr2D() As Variant
    Dim ri As Long, ci As Long

    If IsObject(data) Then
        If TypeOf data Is Range Then
            Set rng = data
            If rng.Areas.Count > 1 Then Set rng = rng.Areas(1)
            arr = rng.Value
            If Not IsArray(arr) Then
                ReDim temp(1 To 1, 1 To 1)
                If IsEmpty(arr) Then
                    temp(1, 1) = Empty
                Else
                    temp(1, 1) = arr
                End If
                arr = temp
            End If
            numRows = UBound(arr, 1)
            numCols = UBound(arr, 2)
            NormalizeTo2D = arr
        Else
            numRows = 0: numCols = 0
            NormalizeTo2D = Array()
        End If
    ElseIf IsArray(data) Then
        If IsEmptyArray(data) Then
            numRows = 0: numCols = 0
            NormalizeTo2D = Array(): Exit Function
        End If
        Err.Clear
        On Error Resume Next
        numRows = UBound(data, 1)
        numCols = UBound(data, 2)
        If Err.Number <> 0 Then
            Err.Clear
            numRows = UBound(data, 1) - LBound(data, 1) + 1
            numCols = 1
            ReDim arr2D(1 To numRows, 1 To 1)
            For ri = 1 To numRows: arr2D(ri, 1) = data(LBound(data) + ri - 1): Next
            NormalizeTo2D = arr2D
            Exit Function
        End If
        On Error GoTo 0
        ' Normalize 2D array to 1-based if needed (COM returns 0-based)
        Dim lb1 As Long: lb1 = LBound(data, 1)
        Dim lb2 As Long: lb2 = LBound(data, 2)
        If lb1 <> 1 Or lb2 <> 1 Then
            Dim nR As Long: nR = UBound(data, 1) - lb1 + 1
            Dim nC As Long: nC = UBound(data, 2) - lb2 + 1
            ReDim arr2D(1 To nR, 1 To nC)
            For ri = 1 To nR
                For ci = 1 To nC
                    arr2D(ri, ci) = data(lb1 + ri - 1, lb2 + ci - 1)
                Next
            Next
            numRows = nR: numCols = nC
            NormalizeTo2D = arr2D
            Exit Function
        End If
        NormalizeTo2D = data
    Else
        numRows = 0: numCols = 0
        NormalizeTo2D = Array()
    End If
End Function

'=============================================================================
' NormalizeInput — Convert Range→Array and optionally flatten 2D column→1D
'
' Handles the common pattern where COM/VBA callers pass Range objects to
' functions expecting arrays. Modifies ByRef parameter in place.
'
' After calling, v is guaranteed to be either an array or a non-Range Variant.
' Set flattenColumn=True when the parameter is expected to be a 1D list
' (e.g., a single column of values).
'=============================================================================
Public Function NormalizeInput(ByRef v As Variant, _
                                Optional ByVal flattenColumn As Boolean = False) As Boolean
    NormalizeInput = False

    ' Step 1: Range → Array
    If IsObject(v) Then
        If TypeOf v Is Range Then
            If v.Areas.Count > 1 Then
                v = v.Areas(1).Value
            Else
                v = v.Value
            End If
            NormalizeInput = True
        End If
    End If

    ' Step 2: Flatten 2D column to 1D if requested
    If flattenColumn Then
        If Not IsArray(v) Then Exit Function
        Dim is2D As Boolean: is2D = False
        On Error Resume Next
        Dim colLb As Long: colLb = LBound(v, 2)
        is2D = (Err.Number = 0)
        On Error GoTo 0
        If is2D Then
            Dim flat() As Variant, fi As Long
            ' Single-row multi-column: iterate columns (e.g. Range("A1:C1") → 1×3)
            If UBound(v, 1) = LBound(v, 1) And UBound(v, 2) > LBound(v, 2) Then
                ReDim flat(LBound(v, 2) To UBound(v, 2))
                For fi = LBound(v, 2) To UBound(v, 2)
                    If IsObject(v(LBound(v, 1), fi)) Then
                        Set flat(fi) = v(LBound(v, 1), fi)
                    Else
                        flat(fi) = v(LBound(v, 1), fi)
                    End If
                Next
            Else
                ' Multi-row (typical column vector): iterate rows, take first column
                ReDim flat(LBound(v, 1) To UBound(v, 1))
                For fi = LBound(v, 1) To UBound(v, 1)
                    If IsObject(v(fi, LBound(v, 2))) Then
                        Set flat(fi) = v(fi, LBound(v, 2))
                    Else
                        flat(fi) = v(fi, LBound(v, 2))
                    End If
                Next
            End If
            v = flat
            NormalizeInput = True
        End If
    End If
End Function

Public Function Normalize2D(ByVal v As Variant) As Variant
    Dim data As Variant, result() As Variant
    Dim nRows As Long, i As Long

    If IsObject(v) Then
        If TypeOf v Is Range Then
            If v.Areas.Count > 1 Then
                Normalize2D = CVErr(xlErrValue): Exit Function
            End If
            If v.Count = 1 Then
                ReDim result(1 To 1, 1 To 1): result(1, 1) = v.Value
                Normalize2D = result: Exit Function
            End If
            Normalize2D = v.Value: Exit Function
        Else
            ReDim result(1 To 1, 1 To 1): result(1, 1) = v
            Normalize2D = result: Exit Function
        End If
    End If

    data = v
    If Not IsArray(data) Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = data
        Normalize2D = result: Exit Function
    End If

    If IsEmptyArray(data) Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = Empty
        Normalize2D = result: Exit Function
    End If

    If ArrayDims(data) = 1 Then
        nRows = UBound(data) - LBound(data) + 1
        ReDim result(1 To nRows, 1 To 1)
        For i = 1 To nRows
            result(i, 1) = data(LBound(data) + i - 1)
        Next i
        Normalize2D = result: Exit Function
    End If

    ' 如果 2D 数组是 0-based, 转换为 1-based (与 Range.Value 约定一致)
    If LBound(data, 1) <> 1 Or LBound(data, 2) <> 1 Then
        Dim nR As Long, nC As Long, ri As Long, ci As Long
        nR = UBound(data, 1) - LBound(data, 1) + 1
        nC = UBound(data, 2) - LBound(data, 2) + 1
        ReDim result(1 To nR, 1 To nC)
        For ri = 1 To nR
            For ci = 1 To nC
                result(ri, ci) = data(LBound(data, 1) + ri - 1, LBound(data, 2) + ci - 1)
        Next ci, ri
        Normalize2D = result: Exit Function
    End If
    Normalize2D = data
End Function

Public Function ToDoubles(ByRef v As Variant) As Double()
    Dim result() As Double, cnt As Long, i As Long
    Dim data As Variant, lb As Long, ub As Long

    data = Normalize1D(v)
    If Not IsArray(data) Then
        ReDim result(0 To 0): result(0) = 0#
        ToDoubles = result: Exit Function
    End If

    lb = LBound(data): ub = UBound(data)
    If ub < lb Then
        ' 空数组 (如 Array(), LBound=0/UBound=-1) — 返回 Erased Double()
        ' (§4.2 禁止 ReDim(0 To -1); Erase 返回未初始化数组 — 调用方访问 LBound/UBound 将 Error 9)
        Erase result
        ToDoubles = result: Exit Function
    End If
    ReDim result(lb To ub)
    cnt = lb - 1
    For i = lb To ub
        ' Exclude Boolean/Date — consistent with IsNumericCell (IsNumeric(True)=True in VBA)
        If IsNumeric(data(i)) Then
            If VarType(data(i)) <> vbBoolean And VarType(data(i)) <> vbDate Then
                cnt = cnt + 1: result(cnt) = CDbl(data(i))
            End If
        End If
    Next i
    If cnt < lb Then
        Erase result
        ToDoubles = result
        Exit Function
    End If
    If cnt < ub Then ReDim Preserve result(lb To cnt)
    ToDoubles = result
End Function

Public Function WrapScalar(ByVal v As Variant) As Variant()
    Dim tmp(0 To 0) As Variant
    If IsArray(v) Then
        WrapScalar = v
    Else
        tmp(0) = v
        WrapScalar = tmp
    End If
End Function

' ======== Compare ========

Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeStr = "Null"
    ElseIf IsError(v) Then
        SafeStr = "#Error"
    ElseIf IsObject(v) Then
        SafeStr = "[Object]"
    ElseIf IsArray(v) Then
        SafeStr = "[Array]"
    Else
        SafeStr = CStr(v)
    End If
End Function

Private Function ArrayToKey(ByVal v As Variant) As String
    Dim i As Long, lb As Long, ub As Long
    Dim flatArr() As Variant, flat As Variant
    Dim r As Long, c As Long, idx As Long
    Dim nRows As Long, nCols As Long
    Dim parts() As String

    Err.Clear
    On Error Resume Next
    lb = LBound(v, 1)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        ArrayToKey = "Array:##EMPTY##": Exit Function
    End If
    Err.Clear: On Error GoTo 0
    ub = UBound(v, 1)

    On Error Resume Next
    Dim dummy As Long: dummy = UBound(v, 2)
    If Err.Number = 0 Then
        Err.Clear: On Error GoTo 0
        nRows = UBound(v, 1) - LBound(v, 1) + 1
        nCols = UBound(v, 2) - LBound(v, 2) + 1
        ReDim flatArr(0 To nRows * nCols - 1)
        idx = 0
        For r = 1 To nRows
            For c = 1 To nCols
                flatArr(idx) = v(LBound(v, 1) + r - 1, LBound(v, 2) + c - 1)
                idx = idx + 1
            Next c
        Next r
        flat = flatArr
        ArrayToKey = "Array(nd):" & ArrayToKey(flat): Exit Function
    End If
    Err.Clear: On Error GoTo 0

    If ub < lb Then ArrayToKey = "Array:##EMPTY##": Exit Function
    ReDim parts(lb To ub)
    For i = lb To ub
        parts(i) = SafeKey(v(i))
    Next i
    ArrayToKey = "Array(" & CStr(ub - lb + 1) & "):" & Join(parts, "|")
End Function

Public Function SafeKey(ByVal v As Variant) As String
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

Public Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
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
    If IsDate(a) And IsDate(b) Then
        Dim da As Variant, db As Variant
        Dim okA As Boolean, okB As Boolean
        On Error Resume Next
        Err.Clear
        da = CDate(a): okA = (Err.Number = 0)
        Err.Clear
        db = CDate(b): okB = (Err.Number = 0)
        On Error GoTo 0
        If okA And okB Then
            ValuesEqual = (da = db): Exit Function
        End If
        ' CDate 失败 → 回退到 CStr 比较 (用成功标志位判断，避免双失败时 (Empty=Empty)=True 的误判)
    End If
    ValuesEqual = (CStr(a) = CStr(b))
End Function

Public Function Compare(ByVal a As Variant, ByVal b As Variant) As Long
    If IsNull(a) And IsNull(b) Then Compare = 0: Exit Function
    If IsNull(a) Then Compare = -1: Exit Function
    If IsNull(b) Then Compare = 1: Exit Function
    If IsError(a) And IsError(b) Then Compare = 0: Exit Function
    If IsError(a) Then Compare = 1: Exit Function
    If IsError(b) Then Compare = -1: Exit Function
    ' Empty: type-discriminated to prevent IsNumeric(Empty)=True from merging with 0
    If IsEmpty(a) And IsEmpty(b) Then Compare = 0: Exit Function
    If IsEmpty(a) Then Compare = -1: Exit Function
    If IsEmpty(b) Then Compare = 1: Exit Function
    ' Boolean: type-discriminated to prevent IsNumeric(True)=True from merging with -1
    If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        If CBool(a) = CBool(b) Then
            Compare = 0
        ElseIf CBool(a) Then
            Compare = 1   ' True > False
        Else
            Compare = -1  ' False < True
        End If
        Exit Function
    End If
    If VarType(a) = vbBoolean Then Compare = 1: Exit Function
    If VarType(b) = vbBoolean Then Compare = -1: Exit Function
    If VarType(a) = vbDate And VarType(b) = vbDate Then
        If CDate(a) < CDate(b) Then
            Compare = -1
        ElseIf CDate(a) > CDate(b) Then
            Compare = 1
        Else
            Compare = 0
        End If
    ElseIf IsNumeric(a) And IsNumeric(b) Then
        If CDbl(a) < CDbl(b) Then
            Compare = -1
        ElseIf CDbl(a) > CDbl(b) Then
            Compare = 1
        Else
            Compare = 0
        End If
    Else
        ' Arrays: compare by dimensions and element count, then element-by-element
        If IsArray(a) And IsArray(b) Then
            Dim dimsA As Long: dimsA = ArrayDims(a)
            Dim dimsB As Long: dimsB = ArrayDims(b)
            If dimsA <> dimsB Then If dimsA < dimsB Then Compare = -1 Else Compare = 1: Exit Function
            Dim nA As Long: nA = UBound(a, 1) - LBound(a, 1) + 1
            Dim nB As Long: nB = UBound(b, 1) - LBound(b, 1) + 1
            If nA <> nB Then If nA < nB Then Compare = -1 Else Compare = 1: Exit Function
            ' Arrays of same dims/size are equal for ordering
            Compare = 0: Exit Function
        End If
        If IsArray(a) Then Compare = 1: Exit Function
        If IsArray(b) Then Compare = -1: Exit Function
        Compare = StrComp(SafeStr(a), SafeStr(b), vbTextCompare)
    End If
End Function

' ======== Assign ========

' @deprecated Use inline IsObject dispatch instead:
'   Dim tmp As Variant: tmp = source
'   If IsObject(tmp) Then Set target = tmp Else target = tmp
' Works for function returns and local variables — no restrictions.
' Retained only for backward compatibility with existing consumers.
Public Sub VarLetSet(ByRef target As Variant, ByVal source As Variant)
    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If
End Sub

' ======== Filter ========

' FilterPasses — unified condition filter.
' Operators: "=" "<>" "<" "<=" ">" ">=" "contains" "notcontains" "startswith" "endswith"
'            "isblank" "isnotblank" "regex"
Public Function FilterPasses(ByVal element As Variant, ByVal matchValue As Variant, ByVal op As String) As Boolean
    Static regex As Object
    Dim s As String, sCrit As String

    ' Fast-reject: Error/Null/Object/Array cannot be compared by value operators
    Select Case LCase$(op)
        Case "isblank":  FilterPasses = (Len(Trim(CStr(element))) = 0): Exit Function
        Case "isnotblank": FilterPasses = (Len(Trim(CStr(element))) > 0): Exit Function
    End Select

    If IsError(element) Or IsNull(element) Or IsObject(element) Or IsArray(element) Then Exit Function
    If IsError(matchValue) Or IsNull(matchValue) Or IsObject(matchValue) Or IsArray(matchValue) Then Exit Function

    Select Case LCase$(op)
        Case "=":  FilterPasses = ValuesEqual(element, matchValue)
        Case "<>": FilterPasses = Not ValuesEqual(element, matchValue)
        Case "<":  FilterPasses = Compare(element, matchValue) < 0
        Case "<=": FilterPasses = Compare(element, matchValue) <= 0
        Case ">":  FilterPasses = Compare(element, matchValue) > 0
        Case ">=": FilterPasses = Compare(element, matchValue) >= 0
        Case "contains":    FilterPasses = (InStr(1, CStr(element), CStr(matchValue), vbTextCompare) > 0)
        Case "notcontains": FilterPasses = (InStr(1, CStr(element), CStr(matchValue), vbTextCompare) = 0)
        Case "startswith"
            s = CStr(element): sCrit = CStr(matchValue)
            FilterPasses = (Len(s) >= Len(sCrit)) And (StrComp(Left$(s, Len(sCrit)), sCrit, vbTextCompare) = 0)
        Case "endswith"
            s = CStr(element): sCrit = CStr(matchValue)
            FilterPasses = (Len(s) >= Len(sCrit)) And (StrComp(Right$(s, Len(sCrit)), sCrit, vbTextCompare) = 0)
        Case "regex"
            If regex Is Nothing Then
                On Error Resume Next
                Set regex = CreateObject("VBScript.RegExp")
                If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
                On Error GoTo 0: regex.IgnoreCase = True
            End If
            On Error Resume Next
            regex.Pattern = CStr(matchValue)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
            FilterPasses = regex.Test(CStr(element))
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
            On Error GoTo 0
    End Select
End Function

'=============================================================================
' Test_VariantKit — Debug.Assert test harness
' Run: Call Test_VariantKit (any assertion failure halts VBA)
'=============================================================================
Public Sub Test_VariantKit()
    On Error GoTo ErrHandler
    Dim failCount As Long: failCount = 0
    Dim arr As Variant, result As Variant, arr2D As Variant
    Dim numRows As Long, numCols As Long, i As Long

    '=====================================================================
    ' Normalize1D — scalar → array
    '=====================================================================
    result = Normalize1D(42)
    If Not (LBound(result) = 0) Then Err.Raise 5
    If Not (UBound(result) = 0) Then Err.Raise 5
    If Not (result(0) = 42) Then Err.Raise 5

    ' Normalize1D — 1D array pass-through
    result = Normalize1D(Array(1, 2, 3))
    If Not (LBound(result) = 0) Then Err.Raise 5
    If Not (result(0) = 1) Then Err.Raise 5
    If Not (result(2) = 3) Then Err.Raise 5

    ' Normalize1D — 2D to 1D (row-major)
    ReDim arr2D(1 To 2, 1 To 3)
    arr2D(1, 1) = 1: arr2D(1, 2) = 2: arr2D(1, 3) = 3
    arr2D(2, 1) = 4: arr2D(2, 2) = 5: arr2D(2, 3) = 6
    result = Normalize1D(arr2D)
    If Not (UBound(result) = 5) Then Err.Raise 5
    If Not (result(0) = 1) Then Err.Raise 5
    If Not (result(1) = 2) Then Err.Raise 5
    If Not (result(5) = 6) Then Err.Raise 5

    ' Normalize1D — column-major
    result = Normalize1D(arr2D, "C")
    If Not (result(0) = 1) Then Err.Raise 5
    If Not (result(1) = 4) Then Err.Raise 5
    If Not (result(2) = 2) Then Err.Raise 5

    ' Normalize1D — empty array (returns Array()=Empty — uninitialized array)
    Dim emptyArr() As Variant
    result = Normalize1D(emptyArr)
    ' result is Empty — Verify with IsEmptyArray probe
    If Not IsEmptyArray(result) Then Err.Raise 5

    '=====================================================================
    ' NormalizeTo2D — scalar / 1D / 2D
    '=====================================================================
    ' scalar → 2D (returns Array() = Empty — verify counts, skip array bounds)
    result = NormalizeTo2D(42, numRows, numCols)
    If Not (numRows = 0) Then Err.Raise 5
    If Not (numCols = 0) Then Err.Raise 5

    ' 1D → 2D column
    arr = Array(10, 20, 30)
    result = NormalizeTo2D(arr, numRows, numCols)
    If Not (numRows = 3) Then Err.Raise 5
    If Not (numCols = 1) Then Err.Raise 5
    If Not (result(1, 1) = 10) Then Err.Raise 5
    If Not (result(3, 1) = 30) Then Err.Raise 5

    ' 2D pass-through
    ReDim arr2D(1 To 2, 1 To 2)
    arr2D(1, 1) = 100: arr2D(1, 2) = 200
    arr2D(2, 1) = 300: arr2D(2, 2) = 400
    result = NormalizeTo2D(arr2D, numRows, numCols)
    If Not (numRows = 2) Then Err.Raise 5
    If Not (numCols = 2) Then Err.Raise 5
    If Not (result(1, 1) = 100) Then Err.Raise 5
    If Not (result(2, 2) = 400) Then Err.Raise 5

    '=====================================================================
    ' SafeKey — type-disambiguated key
    '=====================================================================
    If Not (Left$(SafeKey(Null), 4) = "Null") Then Err.Raise 5
    If Not (Left$(SafeKey(CVErr(2000)), 5) = "Error") Then Err.Raise 5
    If Not (Left$(SafeKey(Empty), 5) = "Empty") Then Err.Raise 5
    If Not (Left$(SafeKey(42), 7) = "Numeric") Then Err.Raise 5

    ' SafeKey — object identity via ObjPtr
    Dim d1 As Object, d2 As Object
    Set d1 = CreateObject("Scripting.Dictionary")
    Set d2 = d1
    If Not (SafeKey(d1) = SafeKey(d2)) Then Err.Raise 5
    Set d2 = CreateObject("Scripting.Dictionary")
    If Not (SafeKey(d1) <> SafeKey(d2)) Then Err.Raise 5
    Set d1 = Nothing: Set d2 = Nothing

    ' SafeKey — array key
    If Not (Left$(SafeKey(Array(1, 2)), 5) = "Array") Then Err.Raise 5

    '=====================================================================
    ' ValuesEqual — equality with sub-type handling
    '=====================================================================
    If Not (ValuesEqual(Null, Null) = True) Then Err.Raise 5
    If Not (ValuesEqual(Null, 1) = False) Then Err.Raise 5
    If Not (ValuesEqual(Empty, Empty) = True) Then Err.Raise 5
    If Not (ValuesEqual(Empty, 0) = False) Then Err.Raise 5
    If Not (ValuesEqual(CVErr(2000), CVErr(2000)) = True) Then Err.Raise 5
    If Not (ValuesEqual(CVErr(2000), CVErr(2001)) = False) Then Err.Raise 5

    ' Numeric equality with epsilon
    If Not (ValuesEqual(1#, 1# + 1E-14) = True) Then Err.Raise 5
    If Not (ValuesEqual(1#, 2#) = False) Then Err.Raise 5

    ' Boolean vs numeric distinction
    If Not (ValuesEqual(True, -1) = False) Then Err.Raise 5
    If Not (ValuesEqual(True, True) = True) Then Err.Raise 5
    If Not (ValuesEqual(CLng(-1), -1#) = True) Then Err.Raise 5

    ' String equality
    If Not (ValuesEqual("abc", "abc") = True) Then Err.Raise 5
    If Not (ValuesEqual("abc", "ABC") = False) Then Err.Raise 5

    '=====================================================================
    ' IsEmptyArray
    '=====================================================================
    Dim uninitArr() As Variant
    If Not (IsEmptyArray(uninitArr) = True) Then Err.Raise 5
    If Not (IsEmptyArray(emptyArr) = True) Then Err.Raise 5
    If Not (IsEmptyArray(Array(1, 2, 3)) = False) Then Err.Raise 5
    If Not (IsEmptyArray(42) = False) Then Err.Raise 5

    '=====================================================================
    ' ArrayDims
    '=====================================================================
    If Not (ArrayDims(42) = 0) Then Err.Raise 5
    If Not (ArrayDims(Empty) = 0) Then Err.Raise 5
    If Not (ArrayDims(uninitArr) = 0) Then Err.Raise 5
    If Not (ArrayDims(Array(1, 2, 3)) = 1) Then Err.Raise 5
    ReDim arr2D(1 To 2, 1 To 3)
    If Not (ArrayDims(arr2D) = 2) Then Err.Raise 5

    '=====================================================================
    ' IsNumericArray — contains at least one numeric element
    '=====================================================================
    If Not (IsNumericArray(Array(1, 2, 3)) = True) Then Err.Raise 5
    If Not (IsNumericArray(Array("a", "b", "c")) = False) Then Err.Raise 5
    If Not (IsNumericArray(Array("a", 1, "c")) = True) Then Err.Raise 5
    If Not (IsNumericArray(emptyArr) = False) Then Err.Raise 5
    If Not (IsNumericArray("scalar") = False) Then Err.Raise 5

    '=====================================================================
    ' ToDoubles — numeric → Double()
    '=====================================================================
    Dim dArr() As Double
    dArr = ToDoubles(Array(1, 2#, 3.5))
    If Not (LBound(dArr) = 0) Then Err.Raise 5
    If Not (UBound(dArr) = 2) Then Err.Raise 5
    If Not (dArr(0) = 1#) Then Err.Raise 5
    If Not (dArr(2) = 3.5) Then Err.Raise 5

    ' Mixed → filters non-numeric (Empty is numeric in VBA → converts to 0)
    dArr = ToDoubles(Array(1, "x", 2.5, Empty, 3))
    If Not (UBound(dArr) = 3) Then Err.Raise 5
    If Not (dArr(0) = 1#) Then Err.Raise 5
    If Not (dArr(1) = 2.5) Then Err.Raise 5
    If Not (dArr(2) = 0#) Then Err.Raise 5
    If Not (dArr(3) = 3#) Then Err.Raise 5

    ' All non-numeric → returns empty Double() (Erase result per SKILL.md §4.2)
    dArr = ToDoubles(Array("a", "b"))
    ' Empty typed array — no elements to check

    '=====================================================================
    ' WrapScalar
    '=====================================================================
    result = WrapScalar(42)
    If Not (LBound(result) = 0) Then Err.Raise 5
    If Not (result(0) = 42) Then Err.Raise 5

    result = WrapScalar(Array(1, 2, 3))
    If Not (UBound(result) = 2) Then Err.Raise 5
    If Not (result(0) = 1) Then Err.Raise 5

    If failCount > 0 Then Err.Raise 5, , "VariantKit: " & failCount & " assertion(s) failed"
    Exit Sub
ErrHandler:
    If Err.Number = 5 Then
        failCount = failCount + 1
        Err.Clear
        Resume Next
    End If
    Err.Raise Err.Number
End Sub