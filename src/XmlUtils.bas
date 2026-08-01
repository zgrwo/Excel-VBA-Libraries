Option Explicit

'==============================================================================
' Module:       XmlUtils
' Purpose:      XML: MSXML2 XPath query and table conversion
' Layer:        Text
' Dependencies: MSXML2 (Windows built-in)
' Public:       7 functions/subs
' Notes:        Requires MSXML2 (Windows built-in). XXE protection enabled.
'==============================================================================


'=====================================================================
' XmlUtils.bas — XML to Excel (MSXML2)
'
' 依赖: MSXML2.DOMDocument (Windows 预装)
' 功能: XML 校验 / XPath 取值 / XML 转 2D 数组 / XML 命名空间支持
'
' 命名空间: 所有 XPath 函数接受可选 namespaces 参数
'   格式: "xmlns:prefix1='uri1' xmlns:prefix2='uri2'" (空格分隔)
'   示例: "xmlns:ns='http://www.w3.org/2005/Atom'"
'
' 工作表函数 (UDF_XML_*):
'   UDF_XML_VALIDATE — 校验 XML 格式
'   UDF_XML_GET      — XPath 查询取值
'   UDF_XML_TABLE    — XML 转表格
'
' VBA-only:
'   XmlValidate — 校验 + 错误详情
'   XmlGet      — XPath → Variant
'   XmlToRange  — XML + XPath → 2D 数组 (可直接写入 Range)
'   XmlGetAttr  — XPath + 属性名 → 属性值
'=====================================================================

Private Const ERR_XML_EMPTY As Long = vbObjectError + 1401
Private Const ERR_XML_COM   As Long = vbObjectError + 1402
Private Const ERR_XML_PARSE  As Long = vbObjectError + 1403
Private Const ERR_XML_NOT_FOUND As Long = vbObjectError + 1404

'=====================================================================
' MSXML2 Engine
'=====================================================================

Private Function CreateDoc() As Object
    Dim doc As Object
    On Error Resume Next
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    If Err.Number <> 0 Then
        Err.Clear
        Set doc = CreateObject("MSXML2.DOMDocument.3.0")
    End If
    On Error GoTo 0
    If doc Is Nothing Then Err.Raise ERR_XML_COM, "XmlUtils", "MSXML2 不可用。"
    Set CreateDoc = doc
End Function

Private Function GetDoc(ByVal xml As String, _
                        Optional ByVal namespaces As String) As Object
    If Len(xml) = 0 Then Err.Raise ERR_XML_EMPTY, "XmlUtils", "XML 字符串为空。"
    Dim doc As Object
    Set doc = CreateDoc()
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    ' 禁用 DTD 处理 — 防御 XXE (Billion Laughs / 外部实体注入)
    ' ProhibitDTD 在 MSXML 3.0/6.0 上可用; 4.0 不支持但本项目不使用
    On Error Resume Next
    doc.setProperty "ProhibitDTD", True
    On Error GoTo 0
    doc.setProperty "SelectionLanguage", "XPath"
    ' 设置 XML 命名空间 (MSXML2 SelectionNamespaces)
    If Len(namespaces) > 0 Then
        doc.setProperty "SelectionNamespaces", namespaces
    End If
    If Not doc.LoadXML(xml) Then
        Err.Raise ERR_XML_PARSE, "XmlUtils", _
            doc.parseError.reason & " (行 " & doc.parseError.Line & ")"
    End If
    Set GetDoc = doc
End Function

Private Function DetectColumns(ByRef firstRow As Object, _
                                ByRef cols() As String) As Long
    Dim child As Object
    Dim cnt As Long
    ' 防护: childNodes.Length=0 时 ReDim(1 To 0) 触发 Error 9
    If firstRow.childNodes.Length = 0 Then
        DetectColumns = 0
        Exit Function
    End If
    ReDim cols(1 To firstRow.childNodes.Length)  ' pre-allocate max
    cnt = 0
    For Each child In firstRow.childNodes
        If child.nodeType = 1 Then  ' NODE_ELEMENT
            cnt = cnt + 1
            cols(cnt) = child.nodeName
        End If
    Next
    If cnt > 0 And cnt < UBound(cols) Then
        ReDim Preserve cols(1 To cnt)
    End If
    DetectColumns = cnt
End Function


'=====================================================================
' Public VBA Functions
'=====================================================================

Public Function XmlValidate(ByVal xml As String, _
                             Optional ByRef errDetail As Variant) As Boolean
    If Len(xml) = 0 Then
        errDetail = "XML 字符串为空。"
        XmlValidate = False: Exit Function
    End If
    Dim doc As Object
    On Error Resume Next
    Set doc = CreateDoc()
    On Error GoTo 0
    If doc Is Nothing Then
        errDetail = "MSXML2 不可用。"
        XmlValidate = False: Exit Function
    End If
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    ' 禁用 DTD 处理 — 防御 XXE 攻击 (与 GetDoc 一致)
    On Error Resume Next
    doc.setProperty "ProhibitDTD", True
    On Error GoTo 0
    If doc.LoadXML(xml) Then
        errDetail = Empty
        XmlValidate = True
    Else
        errDetail = doc.parseError.reason & " (行 " & doc.parseError.Line & ")"
        XmlValidate = False
    End If
End Function

Public Function XmlGet(ByVal xml As Variant, ByVal xpath As String, _
                      Optional ByVal namespaces As String) As Variant
    If IsObject(xml) Then
        If TypeName(xml) = "Range" Then xml = xml.Value
    End If
    If IsArray(xml) Then
        On Error Resume Next
        ' 尝试 2D 访问 (Range.Value 典型场景)
        xml = xml(LBound(xml, 1), LBound(xml, 2))
        If Err.Number <> 0 Then
            ' 1D 数组 — 取第一个元素 (#84)
            Err.Clear
            xml = xml(LBound(xml))
        End If
        On Error GoTo 0
    End If
    If Len(CStr(xml)) = 0 Then XmlGet = Empty: Exit Function
    Dim doc As Object
    On Error Resume Next
    Set doc = GetDoc(CStr(xml), namespaces)
    If doc Is Nothing Then
        XmlGet = Empty
        Exit Function
    End If
    On Error GoTo 0
    Dim node As Object
    On Error Resume Next
    Set node = doc.SelectSingleNode(xpath)
    On Error GoTo 0
    If node Is Nothing Then
        Err.Raise ERR_XML_NOT_FOUND, "XmlGet", "XPath 未匹配任何节点: " & xpath
    Else
        XmlGet = node.Text
    End If
End Function

Public Function XmlGetAttr(ByVal xml As String, ByVal xpath As String, _
                            ByVal attrName As String, _
                            Optional ByVal namespaces As String) As Variant
    Dim doc As Object, node As Object
    Dim val As Variant
    Set doc = GetDoc(xml, namespaces)
    On Error Resume Next
    Set node = doc.SelectSingleNode(xpath)
    On Error GoTo 0
    If node Is Nothing Then
        Err.Raise ERR_XML_NOT_FOUND, "XmlGetAttr", "XPath 未匹配任何节点: " & xpath
    End If
    On Error Resume Next
    val = node.getAttribute(attrName)
    On Error GoTo 0
    If IsNull(val) Or IsEmpty(val) Or IsError(val) Then
        Err.Raise ERR_XML_NOT_FOUND, "XmlGetAttr", "属性 '" & attrName & "' 不存在或无法读取"
    Else
        XmlGetAttr = val
    End If
End Function

Public Function XmlToRange(ByVal xml As String, _
                            ByVal rowXPath As String, _
                            Optional ByVal colNames As Variant, _
                            Optional ByVal namespaces As String) As Variant
    Dim doc As Object, rows As Object, cel As Object
    Dim cols() As String, out() As Variant
    Dim nRows As Long, nCols As Long, r As Long, c As Long
    Dim i As Long, lb As Long

    Set doc = GetDoc(xml, namespaces)

    ' 选取行节点
    On Error Resume Next
    Set rows = doc.SelectNodes(rowXPath)
    On Error GoTo 0
    If rows Is Nothing Then
        Err.Raise ERR_XML_NOT_FOUND, "XmlToRange", "XPath 未匹配任何行节点: " & rowXPath
    End If
    nRows = rows.Length
    If nRows = 0 Then
        Err.Raise ERR_XML_NOT_FOUND, "XmlToRange", "XPath 返回空节点集: " & rowXPath
    End If

    ' 确定列名
    If IsMissing(colNames) Then
        nCols = DetectColumns(rows(0), cols)
    ElseIf Not IsArray(colNames) Then
        nCols = 1
        ReDim cols(1 To 1): cols(1) = CStr(colNames)
    Else
        nCols = UBound(colNames) - LBound(colNames) + 1
        ReDim cols(1 To nCols)
        lb = LBound(colNames)
        For i = 1 To nCols: cols(i) = CStr(colNames(lb + i - 1)): Next
    End If
    If nCols = 0 Then
        Err.Raise ERR_XML_NOT_FOUND, "XmlToRange", "未检测到任何列 — 行节点可能无子元素"
    End If

    ' 构建输出数组 (1-based, Range-ready)
    ReDim out(1 To nRows, 1 To nCols)
    For r = 0 To nRows - 1
        For c = 1 To nCols
            Set cel = Nothing
            On Error Resume Next
            Set cel = rows(r).SelectSingleNode(cols(c))
            On Error GoTo 0
            If cel Is Nothing Then
                out(r + 1, c) = Empty
            Else
                out(r + 1, c) = cel.Text
            End If
        Next
    Next
    XmlToRange = out
End Function


'=====================================================================
' UDF Wrappers
'=====================================================================

Public Function UDF_XML_VALIDATE(ByVal xml As Variant) As Variant
    On Error GoTo EH
    If XmlValidate(CStr(xml)) Then UDF_XML_VALIDATE = True Else UDF_XML_VALIDATE = False
    Exit Function
EH: UDF_XML_VALIDATE = CVErr(xlErrValue)
End Function

Public Function UDF_XML_GET(ByVal xml As Variant, ByVal xpath As Variant) As Variant
    On Error GoTo EH
    UDF_XML_GET = XmlGet(CStr(xml), CStr(xpath))
    Exit Function
EH: UDF_XML_GET = CVErr(xlErrValue)
End Function

Public Function UDF_XML_TABLE(ByVal xml As Variant, _
                               ByVal rowXPath As Variant, _
                               Optional ByVal colNames As Variant) As Variant
    On Error GoTo EH
    If IsMissing(colNames) Then
        UDF_XML_TABLE = XmlToRange(CStr(xml), CStr(rowXPath))
    Else
        UDF_XML_TABLE = XmlToRange(CStr(xml), CStr(rowXPath), colNames)
    End If
    Exit Function
EH: UDF_XML_TABLE = CVErr(xlErrValue)
End Function