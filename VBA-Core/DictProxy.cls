' VERSION 1.0 CLASS
' BEGIN
'   MultiUse = -1  'True
' END
' Attribute VB_GlobalNameSpace = False
' Attribute VB_Creatable = False
' Attribute VB_PredeclaredId = False
' Attribute VB_Exposed = False
Option Explicit

' DictProxy.cls — Safe Dictionary Factory + Batch Operations
' Zero-dependency (pure CreateObject). Default CompareMode = vbTextCompare.
' Object assignment: inline IsObject check per SKILL.md §3 (class can't use VarLetSet directly)
'
'   Public:
'     Create([compareMode]) — create a new Scripting.Dictionary
'     FromKeys(keys, [defaultValue], [compareMode]) — create dict from key array
'     ToArray(dict) — export dict keys+items to 2D Variant array
'     Merge(dict1, dict2, [overwrite]) — merge two dictionaries

Private Const ERR_NOT_AVAIL As Long = vbObjectError + 1401

Public Function Create(Optional ByVal compareMode As Long = vbTextCompare) As Object
    On Error Resume Next
    Set Create = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Err.Raise ERR_NOT_AVAIL, "DictProxy", "Scripting.Dictionary unavailable."
    End If
    On Error GoTo 0
    Select Case compareMode
        Case vbBinaryCompare, vbTextCompare, vbDatabaseCompare
            Create.CompareMode = compareMode
        Case Else
            Err.Raise ERR_NOT_AVAIL, "DictProxy", "Unsupported CompareMode: " & compareMode
    End Select
End Function

Public Function FromKeys(ByRef keys As Variant, _
                         Optional ByVal defaultValue As Variant, _
                         Optional ByVal compareMode As Long = vbTextCompare) As Object
    Dim dict As Object, i As Long, lb As Long, ub As Long, k As Variant
    Dim keyStr As String
    Set dict = Create(compareMode)
    If Not IsArray(keys) Then Set FromKeys = dict: Exit Function
    ' 防御: 探测数组是否可安全访问 (未初始化/2D 数组会导致 LBound/UBound 崩溃)
    On Error Resume Next
    lb = LBound(keys): ub = UBound(keys)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Set FromKeys = dict: Exit Function
    ub = UBound(keys, 2)
    If Err.Number = 0 Then Err.Clear: On Error GoTo 0: Set FromKeys = dict: Exit Function
    Err.Clear: On Error GoTo 0
    For i = lb To ub
        If VarType(keys(i)) = vbObject Then Set k = keys(i) Else k = keys(i)
        If IsEmpty(k) Or IsNull(k) Or IsError(k) Or IsArray(k) Or VarType(k) = vbObject Then GoTo NextKey
        keyStr = CStr(k)
        If Not dict.Exists(keyStr) Then dict.Add keyStr, defaultValue
NextKey:
    Next i
    Set FromKeys = dict
End Function

Public Function ToArray(ByRef dict As Object) As Variant
    Dim result() As Variant, keys As Variant, items As Variant, i As Long, n As Long
    If dict Is Nothing Then ToArray = Empty: Exit Function
    n = dict.Count
    If n = 0 Then ToArray = Empty: Exit Function
    keys = dict.Keys: items = dict.Items: ReDim result(1 To n, 1 To 2)
    For i = 0 To n - 1
        result(i + 1, 1) = keys(i)
        ' 对象类型必须使用 Set, 否则对象引用丢失
        If IsObject(items(i)) Then
            Set result(i + 1, 2) = items(i)
        Else
            result(i + 1, 2) = items(i)
        End If
    Next i
    ToArray = result
End Function

Public Function Merge(ByRef dict1 As Object, _
                      ByRef dict2 As Object, _
                      Optional ByVal overwrite As Boolean = False) As Object
    Dim result As Object, key As Variant
    Dim item1 As Variant, item2 As Variant, item3 As Variant
    ' 保留源字典的 CompareMode (避免 vbBinaryCompare 字典合并时大小写键冲突 Error 457)
    Dim cm As Long: cm = vbTextCompare
    If Not dict1 Is Nothing Then
        cm = dict1.CompareMode
    ElseIf Not dict2 Is Nothing Then
        cm = dict2.CompareMode
    End If
    Set result = Create(cm)
    If Not dict1 Is Nothing Then
        For Each key In dict1.Keys
            If IsObject(dict1(key)) Then Set item1 = dict1(key) Else item1 = dict1(key)
            result.Add key, item1
        Next key
    End If
    If Not dict2 Is Nothing Then
        For Each key In dict2.Keys
            If result.Exists(key) Then
                If overwrite Then
                    result.Remove key
                    If IsObject(dict2(key)) Then Set item2 = dict2(key) Else item2 = dict2(key)
                    result.Add key, item2
                End If
            Else
                If IsObject(dict2(key)) Then Set item3 = dict2(key) Else item3 = dict2(key)
                result.Add key, item3
            End If
        Next key
    End If
    Set Merge = result
End Function

'=============================================================================
' Test_DictProxy — Debug.Assert test harness
' Run: Call Test_DictProxy (any assertion failure halts VBA)
'=============================================================================
Public Sub Test_DictProxy()
    On Error GoTo ErrHandler
    Dim failCount As Long: failCount = 0
    Dim dp As DictProxy: Set dp = New DictProxy
    Dim dict As Object, dict1 As Object, dict2 As Object
    Dim result As Variant, i As Long

    '=====================================================================
    ' Create — basic creation
    '=====================================================================
    Set dict = dp.Create()
    If Not (Not dict Is Nothing) Then Err.Raise 5
    If Not (dict.Count = 0) Then Err.Raise 5
    If Not (dict.CompareMode = vbTextCompare) Then Err.Raise 5

    ' vbBinaryCompare
    Set dict = dp.Create(vbBinaryCompare)
    If Not (dict.CompareMode = vbBinaryCompare) Then Err.Raise 5

    ' vbTextCompare
    Set dict = dp.Create(vbTextCompare)
    If Not (dict.CompareMode = vbTextCompare) Then Err.Raise 5

    '=====================================================================
    ' FromKeys — string keys
    '=====================================================================
    Set dict = dp.FromKeys(Array("a", "b", "c"), 0)
    If Not (dict.Count = 3) Then Err.Raise 5
    If Not (dict.Exists("a")) Then Err.Raise 5
    If Not (dict.Exists("B")) Then Err.Raise 5
    If Not (dict("b") = 0) Then Err.Raise 5

    ' FromKeys — numeric keys
    Set dict = dp.FromKeys(Array(1, 2.5, -3), Empty)
    If Not (dict.Count = 3) Then Err.Raise 5
    If Not (dict.Exists("1")) Then Err.Raise 5
    If Not (dict.Exists("2.5")) Then Err.Raise 5

    ' FromKeys — Null/Error/Empty keys should be skipped (H2 fix verification)
    Set dict = dp.FromKeys(Array(Null, "valid", CVErr(2000), Empty, "also_valid"), "test")
    If Not (dict.Count = 2) Then Err.Raise 5
    If Not (dict.Exists("valid")) Then Err.Raise 5
    If Not (dict.Exists("also_valid")) Then Err.Raise 5
    If Not (Not dict.Exists("Null")) Then Err.Raise 5

    ' FromKeys — Object/Array keys should be skipped
    Set dict = dp.FromKeys(Array(CreateObject("Scripting.Dictionary"), "ok"), 1)
    If Not (dict.Count = 1) Then Err.Raise 5
    If Not (dict.Exists("ok")) Then Err.Raise 5

    ' FromKeys — duplicate keys
    Set dict = dp.FromKeys(Array("x", "X", "x"), 1)
    If Not (dict.Count = 1) Then Err.Raise 5

    ' FromKeys — scalar input (not array)
    Set dict = dp.FromKeys("hello", 42)
    If Not (dict.Count = 0) Then Err.Raise 5

    '=====================================================================
    ' ToArray — export to 2D array
    '=====================================================================
    ' Empty dict → empty array (Erase-d, UBound errors)
    Set dict = dp.Create()
    result = dp.ToArray(dict)
    On Error Resume Next
    i = UBound(result, 1)
    If Err.Number = 0 Then Err.Raise 5  ' should error if truly empty
    Err.Clear: On Error GoTo 0

    ' Nothing → empty array
    result = dp.ToArray(Nothing)
    On Error Resume Next
    i = UBound(result, 1)
    If Err.Number = 0 Then Err.Raise 5
    Err.Clear: On Error GoTo 0

    ' Single item
    Set dict = dp.Create()
    dict.Add "key1", "value1"
    result = dp.ToArray(dict)
    If Not (result(1, 1) = "key1") Then Err.Raise 5
    If Not (result(1, 2) = "value1") Then Err.Raise 5

    ' Multiple items
    Set dict = dp.Create()
    dict.Add "a", 1: dict.Add "b", 2: dict.Add "c", 3
    result = dp.ToArray(dict)
    If Not (UBound(result, 1) = 3) Then Err.Raise 5
    If Not (result(1, 1) = "a") Then Err.Raise 5
    If Not (result(1, 2) = 1) Then Err.Raise 5

    '=====================================================================
    ' Merge — two dicts
    '=====================================================================
    Set dict1 = dp.Create()
    dict1.Add "a", 1: dict1.Add "b", 2

    Set dict2 = dp.Create()
    dict2.Add "b", 99: dict2.Add "c", 3

    ' Merge without overwrite → first dict wins for shared keys
    Set dict = dp.Merge(dict1, dict2, False)
    If Not (dict.Count = 3) Then Err.Raise 5
    If Not (dict("a") = 1) Then Err.Raise 5
    If Not (dict("b") = 2) Then Err.Raise 5
    If Not (dict("c") = 3) Then Err.Raise 5

    ' Merge with overwrite → second dict wins for shared keys
    Set dict = dp.Merge(dict1, dict2, True)
    If Not (dict.Count = 3) Then Err.Raise 5
    If Not (dict("a") = 1) Then Err.Raise 5
    If Not (dict("b") = 99) Then Err.Raise 5
    If Not (dict("c") = 3) Then Err.Raise 5

    ' Merge with Nothing
    Set dict = dp.Merge(dict1, Nothing)
    If Not (dict.Count = 2) Then Err.Raise 5
    If Not (dict("a") = 1) Then Err.Raise 5

    ' Merge object-containing items (M2 fix verification)
    Set dict1 = dp.Create()
    Dim inner As Object: Set inner = CreateObject("Scripting.Dictionary")
    inner.Add "nested", 42
    dict1.Add "obj", inner
    dict1.Add "num", 100

    Set dict2 = dp.Create()
    dict2.Add "obj", "replacement"
    dict2.Add "extra", 200

    Set dict = dp.Merge(dict1, dict2, False)
    If Not (dict.Count = 3) Then Err.Raise 5
    If Not (dict("num") = 100) Then Err.Raise 5
    If Not (dict("extra") = 200) Then Err.Raise 5
    If Not (IsObject(dict("obj"))) Then Err.Raise 5  ' from dict1 (not overwritten)

    Set dict = dp.Merge(dict1, dict2, True)
    If Not (dict("obj") = "replacement") Then Err.Raise 5
    If Not (dict("num") = 100) Then Err.Raise 5

    Set inner = Nothing: Set dict = Nothing
    Set dict1 = Nothing: Set dict2 = Nothing

    If failCount > 0 Then Err.Raise 5, , "DictProxy: " & failCount & " assertion(s) failed"
    Exit Sub
ErrHandler:
    If Err.Number = 5 Then
        failCount = failCount + 1
        Err.Clear
        Resume Next
    End If
    Err.Raise Err.Number
End Sub