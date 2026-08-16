Option Explicit

'==============================================================================
' Module:       FileSystemUtils
' Purpose:      File system: UTF-8 read/write, folder ops, drive info
' Layer:        Excel
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       51 functions/subs
'==============================================================================

Private DP As New DictProxy

'=====================================================================
' FileSystemUtils.bas — 文件系统工具集
'
' 工作表函数 (UDF_FS_*):
'   UDF_FS_NORMALIZEPATH   — 统一路径分隔符
'   UDF_FS_PATHCOMBINE     — 拼接路径
'   UDF_FS_FILEEXISTS      — 文件是否存在
'   UDF_FS_FOLDEREXISTS    — 文件夹是否存在
'   UDF_FS_READTEXT        — 读取文本文件
'   UDF_FS_FILENAME        — 提取文件名
'   UDF_FS_BASENAME        — 提取文件名(不含扩展名)
'   UDF_FS_EXTENSION       — 提取扩展名
'   UDF_FS_FOLDERPATH      — 提取父文件夹路径
'   UDF_FS_FILESIZE        — 文件大小(字节)
'   UDF_FS_FILESIZEFMT     — 格式化文件大小
'   UDF_FS_FILEMODIFIED    — 文件最后修改时间
'   UDF_FS_ENSUREFOLDER    — 创建目录(递归)
'   UDF_FS_TEMPFILENAME    — 生成临时文件名
'   UDF_FS_DELETEFILE      — 删除文件
'   UDF_FS_COPYFILE        — 安全复制文件
'   UDF_FS_TEMPFOLDER      — 系统临时文件夹
'   UDF_FS_SPECIALFOLDER   — 特殊文件夹路径
'   UDF_FS_ISPATHVALID     — 路径语法检查
'   UDF_FS_LISTFILES       — 列出文件
'   UDF_FS_LISTFOLDERS     — 列出子文件夹
'   UDF_FS_COPYFOLDER      — 递归复制文件夹
'   UDF_FS_DELETEFOLDER    — 递归删除文件夹
'
' 依赖: Scripting.FileSystemObject (Windows 内置)
'       ADODB.Stream (用于 UTF-8 读写，Windows 内置)
'
' 路径/文件信息:
'   GetFileName     — 提取文件名 (不含路径)
'   GetBaseName     — 提取文件名 (不含扩展名)
'   GetExtension    — 提取扩展名
'   GetFolderPath   — 提取父文件夹路径
'   GetFileSize     — 获取文件大小 (字节)
'   GetFileSizeFmt  — 格式化文件大小 (KB/MB/GB)
'   FileModified    — 文件最后修改时间
'   NormalizePath   — 统一路径分隔符
'   PathCombine     — 拼接路径
'   IsPathValid     — 路径语法是否合法
'   GetTempFolder   — 系统临时文件夹
'   GetSpecialFolder — 获取特殊文件夹 (桌面/文档等)
'   GetDriveInfo    — 驱动器信息
'
' 文件操作:
'   ReadTextFile    — 读取文本文件
'   WriteTextFile   — 写入文本文件
'   AppendTextFile  — 追加文本
'   ReadBinaryFile  — 读取二进制文件
'   WriteBinaryFile — 写入二进制文件
'   FileExists      — 文件是否存在
'   DeleteFile      — 删除文件
'   CopyFileSafe    — 安全复制 (自动创建目标文件夹)
'
' 文件夹操作:
'   FolderExists    — 文件夹是否存在
'   ListFiles       — 列出文件
'   ListFolders     — 列出子文件夹
'   EnsureFolder    — 创建目录 (递归)
'   CopyFolder      — 递归复制文件夹
'   DeleteFolder    — 递归删除文件夹
'
' 实用工具:
'   TempFileName    — 生成临时文件名
'=====================================================================

Private Const FSO_PROGID As String = "Scripting.FileSystemObject"
Private Const ERR_NOT_AVAIL   As Long = vbObjectError + 1001  ' Reserved
Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1002
Private Const ERR_FILE_NOT_FOUND As Long = vbObjectError + 1010
Private Const ERR_BINARY_READ As Long = vbObjectError + 1011
Private Const ERR_BINARY_WRITE As Long = vbObjectError + 1012


'=============================================================================
' GetFSO — 获取 FileSystemObject 实例 (带缓存)
'=============================================================================
Private Function GetFSO() As Object
    Static fso As Object
    Dim test As Boolean
    If fso Is Nothing Then
        Err.Clear
        On Error Resume Next
        Set fso = CreateObject(FSO_PROGID)
        On Error GoTo 0
    Else
        ' 防御性检查: 验证缓存对象仍然有效
        Err.Clear
        On Error Resume Next
        test = fso.DriveExists(Environ$("SystemDrive") & "\")
        If Err.Number <> 0 Then
            Err.Clear
            Set fso = Nothing
            Set fso = CreateObject(FSO_PROGID)
        End If
        On Error GoTo 0
    End If
    Set GetFSO = fso
End Function

'=============================================================================
' NormalizePath — 统一路径分隔符为反斜杠
'=============================================================================
Public Function NormalizePath(ByVal path As String) As String
    Dim uncPrefix As Boolean
    path = Replace(path, "/", "\")
    ' 保留 UNC 路径前缀 (\\server\share)
    uncPrefix = (Left$(path, 2) = "\\")
    If uncPrefix Then path = Mid$(path, 3)
    Do While InStr(path, "\\") > 0
        path = Replace(path, "\\", "\")
    Loop
    If uncPrefix Then path = "\\" & path
    NormalizePath = path
End Function

'=============================================================================
' PathCombine — 安全拼接路径
'=============================================================================
Public Function PathCombine( _
    ByVal folderPath As String, _
    ByVal fileName As String) As String

    folderPath = NormalizePath(folderPath)
    If Len(folderPath) = 0 Then PathCombine = fileName: Exit Function
    If Right$(folderPath, 1) = "\" Then
        PathCombine = folderPath & fileName
    Else
        PathCombine = folderPath & "\" & fileName
    End If
End Function

'=============================================================================
' FileExists / FolderExists
'=============================================================================
Public Function FileExists(ByVal path As String) As Boolean
    Dim fso As Object
    If Len(path) = 0 Then
        FileExists = False
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        FileExists = (Len(Dir(path)) > 0)
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    FileExists = fso.FileExists(path)
    On Error GoTo 0
End Function

Public Function FolderExists(ByVal path As String) As Boolean
    Dim fso As Object
    If Len(path) = 0 Then
        FolderExists = False
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        FolderExists = (Len(Dir(path, vbDirectory)) > 0)
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    FolderExists = fso.FolderExists(path)
    On Error GoTo 0
End Function

'=============================================================================
' ReadTextFile — 读取文本文件
'
' 参数:
'   filePath - 文件路径
'   encoding - "UTF-8" (默认), "ANSI", "Unicode" (UTF-16LE)
'   lines    - 若 > 0 只读取前 n 行
'=============================================================================
Public Function ReadTextFile( _
    ByVal filePath As String, _
    Optional ByVal encoding As String = "UTF-8", _
    Optional ByVal lines As Long = -1) As String

    Dim content As String
    Dim arr() As String
    Dim slice() As String
    Dim i As Long

    If Len(filePath) = 0 Then
        ReadTextFile = ""
        Exit Function
    End If
    If Not ValidateSafePath(filePath) Then
        ReadTextFile = ""
        Exit Function
    End If
    If Not FileExists(filePath) Then
        ReadTextFile = ""
        Exit Function
    End If

    Select Case UCase$(encoding)
        Case "UTF-8"
            content = ReadUTF8(filePath)
        Case "UNICODE", "UTF-16", "UTF-16LE"
            content = ReadUnicode(filePath)
        Case Else ' ANSI / ASCII
            content = ReadANSI(filePath)
    End Select

    If lines > 0 Then
        ' 统一换行符: CRLF, LF, CR → vbCrLf
        content = Replace(content, vbCrLf, vbLf)
        content = Replace(content, vbCr, vbLf)
        content = Replace(content, vbLf, vbCrLf)
        arr = Split(content, vbCrLf)
        If lines > UBound(arr) - LBound(arr) + 1 Then lines = UBound(arr) - LBound(arr) + 1
        ' 使用 Join 实现 O(n) 性能，代替 O(n²) 字符串拼接
        ReDim slice(0 To lines - 1)
        For i = 0 To lines - 1
            slice(i) = arr(LBound(arr) + i)
        Next i
        ReadTextFile = Join(slice, vbCrLf)
    Else
        ReadTextFile = content
    End If
End Function

Private Function ReadUTF8(ByVal filePath As String) As String
    Dim stream As Object
    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUTF8 = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function

Fallback:
    Err.Clear
    If Not stream Is Nothing Then
        Err.Clear: On Error Resume Next: stream.Close: Set stream = Nothing: On Error GoTo 0
    End If
    ReadUTF8 = UTF8FallbackRead(filePath)
End Function

' -- UTF8FallbackRead — 纯 VBA UTF-8 读取（ADODB 不可用时的回退路径）--
Private Function UTF8FallbackRead(ByVal filePath As String) As String
    Dim fNum As Integer, rawBytes() As Byte, fileSz As Long
    Dim startPos As Long, byteCount As Long
    On Error GoTo ErrHandler

    fNum = FreeFile
    Open filePath For Binary As #fNum
    fileSz = LOF(fNum)
    If fileSz = 0 Then
        Close #fNum
        UTF8FallbackRead = ""
        Exit Function
    End If

    ' 跳过 BOM（EF BB BF）如果存在
    If fileSz >= 3 Then
        Dim b0 As Byte, b1 As Byte, b2 As Byte
        Get #fNum, 1, b0: Get #fNum, 2, b1: Get #fNum, 3, b2
        If b0 = &HEF And b1 = &HBB And b2 = &HBF Then
            startPos = 4
            byteCount = fileSz - 3
        Else
            startPos = 1
            byteCount = fileSz
        End If
    Else
        startPos = 1
        byteCount = fileSz
    End If

    If byteCount > 0 Then
        ReDim rawBytes(0 To byteCount - 1)
        Get #fNum, startPos, rawBytes
    End If
    Close #fNum

    If byteCount = 0 Then
        UTF8FallbackRead = ""
        Exit Function
    End If

    UTF8FallbackRead = UTF8DecodeBytes(rawBytes)
    Exit Function

ErrHandler:
    If fNum > 0 Then Close #fNum
    UTF8FallbackRead = ""
End Function

' -- UTF8DecodeBytes — 纯 VBA UTF-8 字节序列 → Unicode 字符串 --
Private Function UTF8DecodeBytes(ByRef rawBytes() As Byte) As String
    Dim i As Long, n As Long, cp As Long
    Dim b As Byte, trail As Long, minCp As Long
    n = UBound(rawBytes) - LBound(rawBytes) + 1
    Dim chars() As Long  ' 动态数组存储解码后的码点
    ReDim chars(0 To n - 1)
    Dim outCnt As Long: outCnt = 0

    i = LBound(rawBytes)
    Do While i <= UBound(rawBytes)
        b = rawBytes(i)
        If b < &H80 Then
            ' 1-byte: U+0000–U+007F
            cp = b
            trail = 0
        ElseIf (b And &HE0) = &HC0 Then
            ' 2-byte: U+0080–U+07FF
            cp = b And &H1F
            trail = 1
            minCp = &H80
        ElseIf (b And &HF0) = &HE0 Then
            ' 3-byte: U+0800–U+FFFF
            cp = b And &HF
            trail = 2
            minCp = &H800
        ElseIf (b And &HF8) = &HF0 Then
            ' 4-byte: U+10000–U+10FFFF
            cp = b And &H7
            trail = 3
            minCp = &H10000
        Else
            ' 无效首字节 → 替换为 U+FFFD
            chars(outCnt) = &HFFFD: outCnt = outCnt + 1
            i = i + 1
            GoTo Continue
        End If

        ' 检查是否有足够的续字节
        If i + trail > UBound(rawBytes) Then
            chars(outCnt) = &HFFFD: outCnt = outCnt + 1
            i = i + 1
            GoTo Continue
        End If

        ' 读取续字节（格式: 10xxxxxx）
        Dim j As Long
        For j = 1 To trail
            Dim tb As Byte: tb = rawBytes(i + j)
            If (tb And &HC0) <> &H80 Then
                ' 续字节格式无效
                chars(outCnt) = &HFFFD: outCnt = outCnt + 1
                i = i + 1
                GoTo Continue
            End If
            cp = (cp * &H40) Or (tb And &H3F)
        Next j

        ' 验证：避免过长编码；拒绝代理对范围（U+D800–U+DFFF）；拒绝超出 U+10FFFF 的码点
        If cp < minCp Or (cp >= &HD800 And cp <= &HDFFF) Or cp > &H10FFFF Then
            chars(outCnt) = &HFFFD
        Else
            chars(outCnt) = cp
        End If
        outCnt = outCnt + 1
        i = i + trail + 1
Continue:
    Loop

    ' 将码点数组转换为 VBA 字符串（通过 ChrW$ 处理 BMP 字符，代理对处理补充平面）
    Dim sb() As String, sbIdx As Long
    ReDim sb(0 To outCnt - 1)
    For i = 0 To outCnt - 1
        cp = chars(i)
        If cp <= &HFFFF Then
            sb(i) = ChrW$(cp)
        Else
            ' 补充平面 → UTF-16 代理对
            cp = cp - &H10000
            sb(i) = ChrW$(&HD800 Or (cp \ &H400)) & ChrW$(&HDC00 Or (cp And &H3FF))
        End If
    Next i
    UTF8DecodeBytes = Join(sb, "")
End Function

' -- StripUTF8BOM — 字节级 BOM 检测 (Binary I/O)，在所有代码页上可靠 --
Private Sub StripUTF8BOM(ByRef content As String, ByVal filePath As String)
    Dim fNum As Integer
    Dim hasBOM As Boolean
    Dim b0 As Byte, b1 As Byte, b2 As Byte
    Dim fNum2 As Long
    Dim fileSz As Long
    Dim rawBytes() As Byte

    ' 字节级 BOM 检测 (EF BB BF) — 独立于系统代码页
    fNum = FreeFile
    Err.Clear
    On Error Resume Next
    Open filePath For Binary As #fNum
    If LOF(fNum) >= 3 Then
        Get #fNum, 1, b0: Get #fNum, 2, b1: Get #fNum, 3, b2
        hasBOM = (b0 = &HEF And b1 = &HBB And b2 = &HBF)
    End If
    Close #fNum
    On Error GoTo 0
    If Not hasBOM Then Exit Sub

    ' BOM 已在字节级确认。通过二进制重读取跳过前 3 字节,
    ' 避免 DBCS 代码页中字符级剥离导致的损坏。
    fNum2 = FreeFile
    On Error GoTo BOMReadFail  ' 确保读取失败时关闭 fNum2，避免句柄泄漏 (§资源管理)
    Open filePath For Binary As #fNum2
    fileSz = LOF(fNum2)
    If fileSz > 3 Then
        ReDim rawBytes(fileSz - 4)  ' 0-based: fileSz-3 bytes after BOM
        Get #fNum2, 4, rawBytes
        content = StrConv(rawBytes, vbUnicode)  ' ANSI → Unicode, 所有代码页均正确
    Else
        content = ""
    End If
    Close #fNum2
    Exit Sub

BOMReadFail:
    ' 读取失败 — 关闭句柄防止泄漏；content 保持调用方传入的原值 (BOM 剥离为尽力而为)
    ' ⚠ 此错误处理器仅覆盖 fNum2；fNum 在此作用域设置前已关闭。
    '    若代码重构在 On Error GoTo 作用域内重新使用 fNum，需同步更新此处。
    On Error Resume Next
    Close #fNum2
    On Error GoTo 0
End Sub

Private Function ReadUnicode(ByVal filePath As String) As String
    Dim stream As Object
    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "Unicode"
    stream.Open
    stream.LoadFromFile filePath
    ReadUnicode = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function

Fallback:
    Err.Clear
    If Not stream Is Nothing Then
        Err.Clear: On Error Resume Next: stream.Close: Set stream = Nothing: On Error GoTo 0
    End If
    ReadUnicode = ReadANSI(filePath)
    ' 去除 UTF-16LE BOM (FF FE) — 跨代码页匹配渲染字符
    If Left$(ReadUnicode, 1) = ChrW$(&HFEFF) Then ReadUnicode = Mid$(ReadUnicode, 2)
End Function

Private Function ReadANSI(ByVal filePath As String) As String
    Dim fNum As Integer
    Dim lines() As String
    Dim cnt As Long
    Dim line As String

    On Error GoTo ErrHandler
    fNum = FreeFile
    Open filePath For Input As #fNum
    ReDim lines(0 To 127)
    cnt = 0
    Do Until EOF(fNum)
        Line Input #fNum, line
        If cnt > UBound(lines) Then ReDim Preserve lines(0 To UBound(lines) * 2 + 1)
        lines(cnt) = line
        cnt = cnt + 1
    Loop
    Close #fNum
    If cnt > 0 Then
        ReDim Preserve lines(0 To cnt - 1)
        ReadANSI = Join(lines, vbCrLf)
    Else
        ReadANSI = ""
    End If
    Exit Function

ErrHandler:
    If fNum > 0 Then
        Err.Clear: On Error Resume Next: Close #fNum: On Error GoTo 0
    End If
    ReadANSI = ""
End Function

'=============================================================================
' WriteTextFile — 写入文本文件
'
' 参数:
'   filePath  - 文件路径
'   content   - 文本内容
'   encoding  - "UTF-8" (默认), "ANSI", "Unicode"
'   append    - 是否追加模式
'   bom       - UTF-8 时是否加 BOM (默认 True)
'=============================================================================
Public Sub WriteTextFile( _
    ByVal filePath As String, _
    ByVal content As String, _
    Optional ByVal encoding As String = "UTF-8", _
    Optional ByVal append As Boolean = False, _
    Optional ByVal bom As Boolean = True)

    Dim folderPath As String

    If Len(filePath) = 0 Then Exit Sub
    If Not ValidateSafePath(filePath) Then Exit Sub

    ' 自动创建父文件夹
    folderPath = GetFolderPath(filePath)
    If Len(folderPath) > 0 Then EnsureFolder folderPath

    Select Case UCase$(encoding)
        Case "UTF-8"
            If append And FileExists(filePath) Then
                WriteUTF8Append filePath, content
            Else
                WriteUTF8 filePath, content, bom
            End If
        Case "UNICODE", "UTF-16", "UTF-16LE"
            WriteUnicode filePath, content, append
        Case Else
            WriteANSI filePath, content, append
    End Select
End Sub

Private Sub WriteUTF8(ByVal filePath As String, ByVal content As String, ByVal bom As Boolean)
    Dim stream As Object
    Dim errNum As Long, errSrc As String, errDesc As String
    Dim i As Long
    On Error GoTo ErrHandler
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    If bom Then stream.WriteText ChrW$(&HFEFF)  ' 写入 UTF-8 BOM
    If Len(content) > 0 Then stream.WriteText content
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
    Exit Sub

ErrHandler:
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    If Not stream Is Nothing Then
        Err.Clear: On Error Resume Next: stream.Close: Set stream = Nothing: On Error GoTo 0
    End If
    ' 如果内容包含非 ANSI 字符则重新抛出 — WriteANSI 回退会损坏它们
    For i = 1 To Len(content)
        If AscW(Mid$(content, i, 1)) > 255 Then
            Err.Raise errNum, "WriteTextFile", errDesc & " (content contains Unicode characters; ADODB unavailable)"
        End If
    Next i
    ' 内容为 ANSI 安全 — 回退到 ANSI (BOM 对 ANSI 文件无关紧要)
    WriteANSI filePath, content, False
End Sub

' 注: ADODB.Stream 追加需加载整个文件再重写 (读取-修改-写入)
' 大文件 (>100MB) 会消耗大量内存，建议用 WriteANSI 追加或分批写入
Private Sub WriteUTF8Append(ByVal filePath As String, ByVal content As String)
    Dim stream As Object
    Dim errNum As Long, errSrc As String, errDesc As String
    Dim i As Long
    On Error GoTo ErrHandler
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    If FileExists(filePath) Then
        stream.LoadFromFile filePath
        stream.Position = stream.Size
    End If
    If Len(content) > 0 Then stream.WriteText content, 0  ' adWriteChar
    stream.SaveToFile filePath, 2  ' adSaveCreateOverWrite (覆盖写入)
    stream.Close
    Set stream = Nothing
    Exit Sub

ErrHandler:
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    If Not stream Is Nothing Then
        Err.Clear: On Error Resume Next: stream.Close: Set stream = Nothing: On Error GoTo 0
    End If
    ' 如果内容包含非 ANSI 字符则重新抛出 — ANSI 回退会损坏它们
    For i = 1 To Len(content)
        If AscW(Mid$(content, i, 1)) > 255 Then
            Err.Raise errNum, "WriteTextFileAppend", "Unicode characters in content; ADODB unavailable"
        End If
    Next i
    ' 内容为 ANSI 安全 — 回退到 ANSI
    WriteANSI filePath, content, True
End Sub

Private Sub WriteUnicode(ByVal filePath As String, ByVal content As String, ByVal append As Boolean)
    Dim stream As Object
    Dim errNum As Long, errSrc As String, errDesc As String
    Dim i As Long
    On Error GoTo ErrHandler
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "Unicode"
    stream.Open
    If append And FileExists(filePath) Then
        stream.LoadFromFile filePath
        stream.Position = stream.Size
    End If
    If Len(content) > 0 Then stream.WriteText content
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
    Exit Sub

ErrHandler:
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.Description
    If Not stream Is Nothing Then
        Err.Clear: On Error Resume Next: stream.Close: Set stream = Nothing: On Error GoTo 0
    End If
    ' 如果内容包含非 ANSI 字符则重新抛出 — ANSI 回退会损坏它们
    For i = 1 To Len(content)
        If AscW(Mid$(content, i, 1)) > 255 Then
            Err.Raise errNum, errSrc, "WriteUnicode: Unicode characters in content; ADODB unavailable"
        End If
    Next i
    ' 内容为 ANSI 安全 — 回退到 ANSI
    WriteANSI filePath, content, append
End Sub

Private Sub WriteANSI(ByVal filePath As String, ByVal content As String, ByVal append As Boolean)
    Dim fNum As Integer
    On Error GoTo ErrHandler
    fNum = FreeFile
    If append Then
        Open filePath For Append As #fNum
    Else
        Open filePath For Output As #fNum
    End If
    Print #fNum, content;
    Close #fNum
    Exit Sub

ErrHandler:
    If fNum > 0 Then
        Err.Clear: On Error Resume Next: Close #fNum: On Error GoTo 0
    End If
End Sub

'=============================================================================
' ListFiles — 列出文件
'
' 参数:
'   folder  - 文件夹路径
'   pattern - 通配符 (默认 "*.*")
'   recurse - 是否递归子文件夹
'
' 返回: Variant 字符串数组
'=============================================================================
Public Function ListFiles( _
    ByVal folder As String, _
    Optional ByVal pattern As String = "*.*", _
    Optional ByVal recurse As Boolean = False) As Variant

    Dim fso As Object
    Dim result() As Variant
    Dim cnt As Long
    Dim files As Object
    Dim file As Object

    If Len(folder) = 0 Or Len(pattern) = 0 Then
        result = Array()
        ListFiles = result
        Exit Function
    End If

    folder = NormalizePath(folder)
    Set fso = GetFSO()
    If fso Is Nothing Then
        result = Array()
        ListFiles = result
        Exit Function
    End If

    If Not fso.FolderExists(folder) Then
        result = Array()
        ListFiles = result
        Exit Function
    End If

    If recurse Then
        ReDim result(0 To 127)
        CollectFiles fso.GetFolder(folder), pattern, result, cnt
        If cnt > 0 Then ReDim Preserve result(0 To cnt - 1) Else result = Array()
    Else
        Set files = fso.GetFolder(folder).Files

        If files.Count = 0 Then
            result = Array()
        Else
            ReDim result(0 To files.Count - 1)
            Err.Clear
            On Error Resume Next
            For Each file In files
                If MatchPattern(file.Name, pattern) Then
                    result(cnt) = file.path
                    cnt = cnt + 1
                End If
            Next file
            On Error GoTo 0
            If cnt > 0 Then
                ReDim Preserve result(0 To cnt - 1)
            Else
                result = Array()
            End If
        End If
    End If

    ListFiles = result
End Function

Private Sub CollectFiles(ByRef folderObj As Object, ByVal pattern As String, ByRef result() As Variant, ByRef cnt As Long)
    Dim file As Object
    Dim subFolder As Object

    ' 空集合/不可枚举惯用法：Resume Next 探针 + 循环前 Err.Clear（Err.Number 未查为有意——空集合直接结束循环，见 vba-pitfalls）
    Err.Clear
    On Error Resume Next
    For Each file In folderObj.Files
        If MatchPattern(file.Name, pattern) Then
            If cnt >= UBound(result) Then ReDim Preserve result(0 To UBound(result) * 2 + 1)
            result(cnt) = file.path
            cnt = cnt + 1
        End If
    Next file

    For Each subFolder In folderObj.SubFolders
        CollectFiles subFolder, pattern, result, cnt
    Next subFolder
    On Error GoTo 0
End Sub

Private Function MatchPattern(ByVal fileName As String, ByVal pattern As String) As Boolean
    If pattern = "*.*" Or pattern = "*" Then
        MatchPattern = True
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    MatchPattern = (UCase$(fileName) Like UCase$(pattern))
    On Error GoTo 0
End Function

'=============================================================================
' ListFolders — 列出子文件夹
'=============================================================================
Public Function ListFolders( _
    ByVal folder As String, _
    Optional ByVal recurse As Boolean = False) As Variant

    Dim fso As Object
    Dim result() As Variant
    Dim cnt As Long
    Dim subFolders As Object
    Dim subFolder As Object

    If Len(folder) = 0 Then
        result = Array()
        ListFolders = result
        Exit Function
    End If

    folder = NormalizePath(folder)
    Set fso = GetFSO()
    If fso Is Nothing Then
        result = Array()
        ListFolders = result
        Exit Function
    End If

    If Not fso.FolderExists(folder) Then
        result = Array()
        ListFolders = result
        Exit Function
    End If

    If recurse Then
        ReDim result(0 To 127)
        CollectFolders fso.GetFolder(folder), result, cnt
        If cnt > 0 Then ReDim Preserve result(0 To cnt - 1) Else result = Array()
    Else
        Set subFolders = fso.GetFolder(folder).SubFolders

        If subFolders.Count = 0 Then
            result = Array()
        Else
            ReDim result(0 To subFolders.Count - 1)
            Err.Clear
            On Error Resume Next
            For Each subFolder In subFolders
                result(cnt) = subFolder.path
                cnt = cnt + 1
            Next subFolder
            On Error GoTo 0
        End If
    End If

    ListFolders = result
End Function

Private Sub CollectFolders(ByRef folderObj As Object, ByRef result() As Variant, ByRef cnt As Long)
    Dim subFolder As Object
    Err.Clear
    On Error Resume Next
    For Each subFolder In folderObj.SubFolders
        If cnt >= UBound(result) Then ReDim Preserve result(0 To UBound(result) * 2 + 1)
        result(cnt) = subFolder.path
        cnt = cnt + 1
        CollectFolders subFolder, result, cnt
    Next subFolder
    On Error GoTo 0
End Sub

'=============================================================================
' GetFileName / GetBaseName / GetExtension / GetFolderPath
'=============================================================================
Public Function GetFileName(ByVal path As String) As String
    Dim fso As Object
    Dim pos As Long
    If Len(path) = 0 Or Right$(path, 1) = "\" Then
        GetFileName = ""
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        path = NormalizePath(path)
        pos = InStrRev(path, "\")
        If pos > 0 Then GetFileName = Mid$(path, pos + 1) Else GetFileName = path
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    GetFileName = fso.GetFileName(path)
    On Error GoTo 0
End Function

Public Function GetBaseName(ByVal path As String) As String
    Dim fso As Object
    Dim fn As String
    Dim dot As Long
    If Len(path) = 0 Then
        GetBaseName = ""
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        fn = GetFileName(path)
        dot = InStrRev(fn, ".")
        ' 点号在首位 (dot=1) 时为隐藏文件 (.gitignore → ".gitignore");
        ' dot>1 时视为扩展名分隔符 (file.tar.gz → "file.tar", .. → "..")
        If dot > 1 Then GetBaseName = Left$(fn, dot - 1) Else GetBaseName = fn
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    GetBaseName = fso.GetBaseName(path)
    On Error GoTo 0
End Function

Public Function GetExtension(ByVal path As String) As String
    Dim fso As Object
    Dim fn As String
    Dim dot As Long
    Dim ext As String
    If Len(path) = 0 Then
        GetExtension = ""
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        fn = GetFileName(path)
        dot = InStrRev(fn, ".")
        If dot > 0 Then GetExtension = Mid$(fn, dot) Else GetExtension = ""
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    ext = fso.GetExtensionName(path)
    If Len(ext) > 0 Then GetExtension = "." & ext Else GetExtension = ""
    On Error GoTo 0
End Function

Public Function GetFolderPath(ByVal path As String) As String
    Dim fso As Object
    Dim pos As Long
    If Len(path) = 0 Then
        GetFolderPath = ""
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        path = NormalizePath(path)
        pos = InStrRev(path, "\")
        If pos > 0 Then
            If pos = 3 And Mid$(path, 2, 1) = ":" Then
                GetFolderPath = Left$(path, 3) ' "C:\"
            Else
                GetFolderPath = Left$(path, pos - 1)
            End If
        Else
            GetFolderPath = ""
        End If
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    GetFolderPath = fso.GetParentFolderName(path)
    On Error GoTo 0
End Function

'=============================================================================
' GetFileSize / GetFileSizeFmt
'=============================================================================
Public Function GetFileSize(ByVal filePath As String) As Double
    Dim fso As Object
    If Not FileExists(filePath) Then
        GetFileSize = -1#
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        Err.Clear
        On Error Resume Next
        GetFileSize = CDbl(FileLen(filePath))
        If Err.Number <> 0 Or GetFileSize < 0# Then GetFileSize = -1#
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    GetFileSize = CDbl(fso.GetFile(filePath).Size)
    If Err.Number <> 0 Then GetFileSize = -1#
    On Error GoTo 0
End Function

Public Function GetFileSizeFmt(ByVal filePath As String) As String
    Dim size As Double
    size = GetFileSize(filePath)

    If size < 0 Then
        GetFileSizeFmt = "(not found)"
    ElseIf size < 1024# Then
        GetFileSizeFmt = Format(size, "0") & " B"
    ElseIf size < 1048576# Then
        GetFileSizeFmt = Format(size / 1024#, "0.0") & " KB"
    ElseIf size < 1073741824# Then
        GetFileSizeFmt = Format(size / 1048576#, "0.00") & " MB"
    Else
        GetFileSizeFmt = Format(size / 1073741824#, "0.00") & " GB"
    End If
End Function

'=============================================================================
' FileModified — 文件最后修改时间
'=============================================================================
Public Function FileModified(ByVal filePath As String) As Date
    Dim fso As Object
    If Not FileExists(filePath) Then
        FileModified = CDate(0)
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        Err.Clear
        On Error Resume Next
        FileModified = FileDateTime(filePath)
        If Err.Number <> 0 Then FileModified = CDate(0)
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    FileModified = fso.GetFile(filePath).DateLastModified
    If Err.Number <> 0 Then FileModified = CDate(0)
    On Error GoTo 0
End Function

'=============================================================================
' EnsureFolder — 自动创建不存在的目录 (支持多层嵌套)
'=============================================================================
Public Function EnsureFolder(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Dim parentPath As String

    If Len(folderPath) = 0 Then
        EnsureFolder = False
        Exit Function
    End If

    folderPath = NormalizePath(folderPath)
    If Right$(folderPath, 1) = "\" And Not (Len(folderPath) = 3 And Mid$(folderPath, 2, 1) = ":") Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    ' 防止去尾反斜杠后路径为空
    If Len(folderPath) = 0 Then
        EnsureFolder = False
        Exit Function
    End If

    If FolderExists(folderPath) Then
        EnsureFolder = True
        Exit Function
    End If

    Set fso = GetFSO()
    If fso Is Nothing Then
        EnsureFolder = False
        Exit Function
    End If

    ' 递归创建父级
    parentPath = GetFolderPath(folderPath)
    If Len(parentPath) > 0 And Not FolderExists(parentPath) Then
        If Not EnsureFolder(parentPath) Then
            EnsureFolder = False
            Exit Function
        End If
    End If
    Err.Clear
    On Error Resume Next
    fso.CreateFolder folderPath
    EnsureFolder = FolderExists(folderPath)
    On Error GoTo 0
End Function

'=============================================================================
' TempFileName — 生成临时文件路径
'
' 参数:
'   prefix - 前缀 (默认 "tmp")
'   ext    - 扩展名 (默认 "tmp")
'   folder - 临时目录 (默认系统 Temp)
'=============================================================================
Public Function TempFileName( _
    Optional ByVal prefix As String = "tmp", _
    Optional ByVal ext As String = "tmp", _
    Optional ByVal folder As String = "") As String

    Static counter As Long
    Dim fso As Object
    Dim candidate As String
    Dim i As Long

    If Len(folder) = 0 Then
        folder = Environ$("TEMP")
        If Len(folder) = 0 Then folder = Environ$("TMP")
        If Len(folder) = 0 Then folder = "."
    End If

    If Len(prefix) = 0 Then prefix = "tmp"
    If Len(ext) = 0 Then ext = "tmp"

    folder = NormalizePath(folder)

    Set fso = GetFSO()

    counter = counter + 1

    Do
        candidate = folder & "\" & prefix & "_" & _
                   Format(Timer * 1000# + counter, "00000000") & "_" & _
                   Format(Int(Rnd(-Timer) * 1000000), "000000")  ' 注: Rnd 未设种子, 唯一性由 Timer+counter 保证
        If Len(ext) > 0 Then candidate = candidate & "." & ext
        If Not FileExists(candidate) Then
            TempFileName = candidate
            Exit Function
        End If
        i = i + 1
        If i > 1000 Then
            TempFileName = ""
            Exit Function
        End If
    Loop
End Function

'=============================================================================
' ContainsTraversal — 判断 ".." 是否作为完整路径段出现 (目录穿越攻击)
'
' 仅段级匹配: 文件名内的连续点 (如 "data..v2.txt") 不构成穿越, 不误杀.
'=============================================================================
Private Function ContainsTraversal(ByVal path As String) As Boolean
    Dim norm As String
    norm = Replace(path, "/", "\")
    norm = Replace(norm, "\\", "\")
    ContainsTraversal = (InStr("\" & norm & "\", "\..\") > 0)
End Function

'=============================================================================
' ValidateSafePath — 防目录穿越攻击
'
' 拒绝 ".." 路径段 (包括规范化后), 防止沙箱外文件访问.
' 允许 Windows 绝对路径 (含 :), 但会规范化检查是否含 .. 段.
' 拒绝 UNC 路径 (\\server\share) — 远程目标不受本地沙箱约束.
' 已知限制: 不解析符号链接/junction (GetAbsolutePathName 不展开).
'=============================================================================
Private Function ValidateSafePath(ByVal path As String) As Boolean
    ' 拒绝含 ".." 路径段的路径 (目录穿越攻击; 文件名内连续点不受影响)
    If ContainsTraversal(path) Then Exit Function
    ' 拒绝 UNC 路径 (远程共享不受本地安全策略约束; 兼容 \\ 与 // 两种写法)
    If Left$(path, 2) = "\\" Or Left$(path, 2) = "//" Then Exit Function
    ' 规范化路径后再检查一次 (防止规范化展开 .. 的攻击)
    Dim fso As Object: Set fso = GetFSO()
    If Not fso Is Nothing Then
        Dim normPath As String: normPath = ""  ' 先置空, 避免探测失败时残留旧值
        On Error Resume Next
        normPath = fso.GetAbsolutePathName(path)
        On Error GoTo 0
        If Len(normPath) > 0 Then
            If ContainsTraversal(normPath) Then Exit Function
        End If
    End If
    ValidateSafePath = True
End Function

'=============================================================================
' DeleteFile — 删除文件
'=============================================================================
Public Function DeleteFile(ByVal filePath As String) As Boolean
    If Not ValidateSafePath(filePath) Then Exit Function
    Dim fso As Object
    If Not FileExists(filePath) Then
        DeleteFile = True ' 不存在也算成功 (与 DeleteFolder 保持一致)
        Exit Function
    End If
    Set fso = GetFSO()
    If fso Is Nothing Then
        Err.Clear
        On Error Resume Next
        Kill filePath
        DeleteFile = (Err.Number = 0)
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    fso.DeleteFile filePath, True
    DeleteFile = (Err.Number = 0)
    On Error GoTo 0
End Function

'=============================================================================
' CopyFileSafe — 安全复制 (自动创建目标文件夹)
'=============================================================================
Public Function CopyFileSafe( _
    ByVal sourcePath As String, _
    ByVal destPath As String, _
    Optional ByVal overwrite As Boolean = True) As Boolean

    If Not ValidateSafePath(sourcePath) Or Not ValidateSafePath(destPath) Then Exit Function
    Dim fso As Object

    If Not FileExists(sourcePath) Then
        CopyFileSafe = False
        Exit Function
    End If

    EnsureFolder GetFolderPath(destPath)

    Set fso = GetFSO()
    If fso Is Nothing Then
        Err.Clear
        On Error Resume Next
        FileCopy sourcePath, destPath
        CopyFileSafe = (Err.Number = 0)
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    fso.CopyFile sourcePath, destPath, overwrite
    CopyFileSafe = (Err.Number = 0)
    On Error GoTo 0
End Function

'=============================================================================
' AppendTextFile — 追加文本到文件
'=============================================================================
Public Sub AppendTextFile( _
    ByVal filePath As String, _
    ByVal content As String, _
    Optional ByVal encoding As String = "UTF-8")

    WriteTextFile filePath, content, encoding, True, True
End Sub

'=============================================================================
' ReadBinaryFile — 读取二进制文件为 Byte 数组
'
' outOk 参数保留用于向后兼容，成功时设为 True。
' 错误通过 Err.Raise 向上传播（ADODB 不可用时有 Open For Binary 纯 VBA 降级）。
'=============================================================================
Public Function ReadBinaryFile(ByVal filePath As String, Optional ByRef outOk As Boolean) As Byte()
    Dim result() As Byte
    Dim stream As Object
    Dim fNum As Integer
    Dim fileLen As Long

    If Not FileExists(filePath) Then
        Err.Raise ERR_FILE_NOT_FOUND, "ReadBinaryFile", _
            "文件不存在: " & filePath
    End If
    If Not ValidateSafePath(filePath) Then
        Dim rejReason As String
        If Left$(filePath, 2) = "\\" Or Left$(filePath, 2) = "//" Then
            rejReason = "UNC 路径不受支持"
        Else
            rejReason = "路径含 '..' 目录穿越段"
        End If
        Err.Raise ERR_INVALID_INPUT, "ReadBinaryFile", _
            "路径被安全检查拒绝 (" & rejReason & "): " & filePath
    End If

    ' 首选: ADODB.Stream (最快, 可处理大文件)
    Err.Clear
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        On Error GoTo StreamErrHandler
        stream.Type = 1 ' binary
        stream.Open
        stream.LoadFromFile filePath
        result = stream.Read
        stream.Close
        Set stream = Nothing
        On Error GoTo 0
        outOk = True
        ReadBinaryFile = result
        Exit Function
StreamErrHandler:
        If Not stream Is Nothing Then
            On Error Resume Next: stream.Close: On Error GoTo 0
            Set stream = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' 备用方案: Open For Binary (纯 VBA, 始终可用)
    fNum = FreeFile
    On Error GoTo BinaryErrHandler
    Open filePath For Binary As #fNum
    fileLen = LOF(fNum)
    If fileLen > 0 Then
        ReDim result(0 To fileLen - 1)
        Get #fNum, 1, result
    End If
    Close #fNum
    outOk = True
    ReadBinaryFile = result
    Exit Function

BinaryErrHandler:
    If fNum > 0 Then
        Err.Clear
        On Error Resume Next: Close #fNum: On Error GoTo 0
    End If
    Err.Raise ERR_BINARY_READ, "ReadBinaryFile", _
        "无法读取文件 (ADODB 不可用且 Binary I/O 失败): " & Err.Description
End Function

'=============================================================================
' WriteBinaryFile — 写入 Byte 数组到文件
'=============================================================================
Public Sub WriteBinaryFile(ByVal filePath As String, ByRef bytes() As Byte)
    Dim ub As Long
    Dim stream As Object
    Dim fNum As Integer

    If Len(filePath) = 0 Then Exit Sub
    If Not ValidateSafePath(filePath) Then Exit Sub

    ' 检查未初始化的数组
    Err.Clear
    On Error Resume Next
    ub = UBound(bytes)
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    On Error GoTo 0

    EnsureFolder GetFolderPath(filePath)

    ' 首选: ADODB.Stream — 带 On Error 保护，失败时降级到 Binary I/O
    Err.Clear
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        stream.Type = 1 ' binary
        stream.Open
        If Err.Number = 0 Then
            stream.Write bytes
            If Err.Number = 0 Then
                stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
            End If
        End If
        Dim adoErr As Long: adoErr = Err.Number
        stream.Close
        Set stream = Nothing
        If adoErr = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' 备用方案: Open For Binary (纯 VBA, 始终可用)
    fNum = FreeFile
    On Error GoTo BinaryErrHandler
    Open filePath For Binary As #fNum
    Put #fNum, 1, bytes
    Close #fNum
    Exit Sub

BinaryErrHandler:
    If fNum > 0 Then
        Err.Clear
        On Error Resume Next: Close #fNum: On Error GoTo 0
    End If
    Err.Raise ERR_BINARY_WRITE, "WriteBinaryFile", _
        "无法写入文件 (ADODB 不可用且 Binary I/O 失败): " & Err.Description
End Sub

'=============================================================================
' GetTempFolder — 获取系统临时文件夹路径
'=============================================================================
Public Function GetTempFolder() As String
    GetTempFolder = Environ$("TEMP")
    If Len(GetTempFolder) = 0 Then GetTempFolder = Environ$("TMP")
    If Len(GetTempFolder) = 0 Then GetTempFolder = "C:\Windows\Temp"
    GetTempFolder = NormalizePath(GetTempFolder)
End Function

'=============================================================================
' GetSpecialFolder — 获取特殊文件夹路径
'
' folderType: "desktop", "documents", "downloads", "appdata", "programfiles", "windows"
'=============================================================================
Public Function GetSpecialFolder(ByVal folderType As String) As String
    ' 使用 Shell.Application 获取特殊文件夹
    Dim shell As Object
    Dim folderConstant As Long

    Err.Clear
    On Error Resume Next
    Set shell = CreateObject("Shell.Application")
    On Error GoTo 0

    Select Case LCase$(folderType)
        Case "desktop":     folderConstant = &H10
        Case "documents":   folderConstant = &H5
        Case "downloads":   folderConstant = &H28
        Case "appdata":     folderConstant = &H1A
        Case "programfiles": folderConstant = &H26
        Case "windows":     folderConstant = &H24
        Case "system":      folderConstant = &H25
        Case "favorites":   folderConstant = &H6
        Case "music":       folderConstant = &Hd
        Case "pictures":    folderConstant = &H27
        Case "videos":      folderConstant = &He
        Case Else
            GetSpecialFolder = ""
            Exit Function
    End Select

    Err.Clear
    On Error Resume Next
    GetSpecialFolder = NormalizePath(shell.Namespace(folderConstant).Self.path)
    On Error GoTo 0

    If Len(GetSpecialFolder) = 0 Then
        ' 备用方案: 使用环境变量
        Select Case LCase$(folderType)
            Case "desktop"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Desktop")
            Case "documents"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Documents")
            Case "downloads"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Downloads")
            Case "music"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Music")
            Case "pictures"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Pictures")
            Case "videos"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Videos")
            Case "favorites"
                GetSpecialFolder = NormalizePath(Environ$("USERPROFILE") & "\Favorites")
            Case "appdata"
                GetSpecialFolder = NormalizePath(Environ$("APPDATA"))
            Case "programfiles"
                GetSpecialFolder = NormalizePath(Environ$("ProgramFiles"))
            Case "windows"
                GetSpecialFolder = NormalizePath(Environ$("SystemRoot"))
            Case Else
                GetSpecialFolder = ""
        End Select
    End If
End Function

'=============================================================================
' IsPathValid — 检查路径语法是否合法
'
' 注意: 不检查文件/文件夹是否真实存在
'=============================================================================
Public Function IsPathValid(ByVal path As String) As Boolean
    Dim fso As Object
    Dim illegalChars As Variant
    Dim j As Long

    If Len(path) = 0 Then
        IsPathValid = False
        Exit Function
    End If

    Set fso = GetFSO()
    If fso Is Nothing Then
        ' 简单检查
        IsPathValid = (InStr(path, Chr$(0)) = 0) ' 无 NULL 字符
        Exit Function
    End If

    ' 检查非法字符: < > " | ? * 以及空字符
    illegalChars = Array("<", ">", Chr$(34), "|", "?", "*", Chr$(0))
    For j = LBound(illegalChars) To UBound(illegalChars)
        If InStr(path, CStr(illegalChars(j))) > 0 Then
            IsPathValid = False
            Exit Function
        End If
    Next j

    IsPathValid = True
End Function

'=============================================================================
' GetDriveInfo — 获取驱动器信息
'
' 参数: driveLetter — "C" 或 "C:" 或 "C:\"
' 返回: Dictionary ("total", "free", "available", "type", "filesystem")
'=============================================================================
Public Function GetDriveInfo(ByVal driveLetter As String) As Object
    Dim result As Object
    Dim d As String
    Dim fso As Object
    Dim drv As Object
    Dim driveType As String
    Dim driveReady As Boolean
    Set result = DP.Create()

    If Len(driveLetter) = 0 Then
        Set GetDriveInfo = result
        Exit Function
    End If

    ' 规范化
    d = Left$(NormalizePath(driveLetter) & "\", 1)
    If Len(d) = 0 Then
        Set GetDriveInfo = result
        Exit Function
    End If

    Set fso = GetFSO()
    If fso Is Nothing Then
        Set GetDriveInfo = result
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    Set drv = fso.GetDrive(d)
    On Error GoTo 0

    If drv Is Nothing Then
        Set GetDriveInfo = result
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    driveReady = drv.IsReady
    On Error GoTo 0

    If driveReady Then
        Err.Clear
        On Error Resume Next
        result("total") = CDbl(drv.TotalSize)
        If Err.Number <> 0 Then result("total") = 0#: Err.Clear
        result("free") = CDbl(drv.FreeSpace)
        If Err.Number <> 0 Then result("free") = 0#: Err.Clear
        result("available") = CDbl(drv.AvailableSpace)
        If Err.Number <> 0 Then result("available") = 0#: Err.Clear
        result("filesystem") = drv.FileSystem
        If Err.Number <> 0 Then result("filesystem") = "": Err.Clear
        On Error GoTo 0
    Else
        result("total") = 0#
        result("free") = 0#
        result("available") = 0#
        result("filesystem") = ""
    End If

    Err.Clear
    On Error Resume Next
    Select Case drv.DriveType
        Case 0: driveType = "Unknown"
        Case 1: driveType = "Removable"
        Case 2: driveType = "Fixed"
        Case 3: driveType = "Network"
        Case 4: driveType = "CD-ROM"
        Case 5: driveType = "RAM Disk"
        Case Else: driveType = "Unknown"
    End Select
    If Err.Number <> 0 Then driveType = "Unknown": Err.Clear
    On Error GoTo 0
    result("type") = driveType

    Set fso = Nothing
    Set drv = Nothing
    Set GetDriveInfo = result
End Function

'=============================================================================
' CopyFolder — 递归复制文件夹
'=============================================================================
Public Function CopyFolder( _
    ByVal sourcePath As String, _
    ByVal destPath As String, _
    Optional ByVal overwrite As Boolean = True) As Boolean

    If Not ValidateSafePath(sourcePath) Or Not ValidateSafePath(destPath) Then Exit Function
    Dim fso As Object

    If Not FolderExists(sourcePath) Then
        CopyFolder = False
        Exit Function
    End If

    sourcePath = NormalizePath(sourcePath)
    destPath = NormalizePath(destPath)

    EnsureFolder destPath

    Set fso = GetFSO()
    If fso Is Nothing Then
        CopyFolder = False
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    fso.CopyFolder sourcePath, destPath, overwrite
    CopyFolder = (Err.Number = 0)
    On Error GoTo 0
    Set fso = Nothing
End Function

'=============================================================================
' DeleteFolder — 递归删除文件夹
'=============================================================================
Public Function DeleteFolder(ByVal folderPath As String) As Boolean
    If Not ValidateSafePath(folderPath) Then Exit Function
    Dim fso As Object

    If Not FolderExists(folderPath) Then
        DeleteFolder = True ' 不存在也算成功
        Exit Function
    End If

    folderPath = NormalizePath(folderPath)

    Set fso = GetFSO()
    If fso Is Nothing Then
        DeleteFolder = False
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    fso.DeleteFolder folderPath, True
    DeleteFolder = (Err.Number = 0)
    On Error GoTo 0
    Set fso = Nothing
End Function

'=============================================================================
' 工作表函数 (UDF_FS_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_FS_NORMALIZEPATH(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_NORMALIZEPATH = NormalizePath(path): Exit Function
EH: UDF_FS_NORMALIZEPATH = CVErr(xlErrValue)
End Function

Public Function UDF_FS_PATHCOMBINE(ByVal folderPath As Variant, ByVal fileName As Variant) As Variant
    On Error GoTo EH: UDF_FS_PATHCOMBINE = PathCombine(folderPath, fileName): Exit Function
EH: UDF_FS_PATHCOMBINE = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FILEEXISTS(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_FILEEXISTS = FileExists(path): Exit Function
EH: UDF_FS_FILEEXISTS = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FOLDEREXISTS(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_FOLDEREXISTS = FolderExists(path): Exit Function
EH: UDF_FS_FOLDEREXISTS = CVErr(xlErrValue)
End Function

Public Function UDF_FS_READTEXT(ByVal filePath As Variant, _
    Optional ByVal encoding As Variant = "UTF-8", Optional ByVal lines As Variant = -1) As Variant
    On Error GoTo EH: UDF_FS_READTEXT = ReadTextFile(filePath, encoding, lines): Exit Function
EH: UDF_FS_READTEXT = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FILENAME(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_FILENAME = GetFileName(path): Exit Function
EH: UDF_FS_FILENAME = CVErr(xlErrValue)
End Function

Public Function UDF_FS_BASENAME(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_BASENAME = GetBaseName(path): Exit Function
EH: UDF_FS_BASENAME = CVErr(xlErrValue)
End Function

Public Function UDF_FS_EXTENSION(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_EXTENSION = GetExtension(path): Exit Function
EH: UDF_FS_EXTENSION = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FOLDERPATH(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_FOLDERPATH = GetFolderPath(path): Exit Function
EH: UDF_FS_FOLDERPATH = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FILESIZE(ByVal filePath As Variant) As Variant
    On Error GoTo EH
    Dim r As Double: r = GetFileSize(filePath)
    If r < 0 Then UDF_FS_FILESIZE = CVErr(xlErrValue) Else UDF_FS_FILESIZE = r
    Exit Function
EH: UDF_FS_FILESIZE = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FILESIZEFMT(ByVal filePath As Variant) As Variant
    On Error GoTo EH: UDF_FS_FILESIZEFMT = GetFileSizeFmt(filePath): Exit Function
EH: UDF_FS_FILESIZEFMT = CVErr(xlErrValue)
End Function

Public Function UDF_FS_FILEMODIFIED(ByVal filePath As Variant) As Variant
    On Error GoTo EH
    Dim r As Date: r = FileModified(filePath)
    If r = CDate(0) Then UDF_FS_FILEMODIFIED = CVErr(xlErrValue) Else UDF_FS_FILEMODIFIED = r
    Exit Function
EH: UDF_FS_FILEMODIFIED = CVErr(xlErrValue)
End Function

Public Function UDF_FS_ENSUREFOLDER(ByVal folderPath As Variant) As Variant
    On Error GoTo EH: UDF_FS_ENSUREFOLDER = EnsureFolder(folderPath): Exit Function
EH: UDF_FS_ENSUREFOLDER = CVErr(xlErrValue)
End Function

Public Function UDF_FS_TEMPFILENAME(Optional ByVal prefix As Variant = "tmp", _
    Optional ByVal ext As Variant = "tmp", Optional ByVal folder As Variant = "") As Variant
    On Error GoTo EH: UDF_FS_TEMPFILENAME = TempFileName(prefix, ext, folder): Exit Function
EH: UDF_FS_TEMPFILENAME = CVErr(xlErrValue)
End Function

Public Function UDF_FS_DELETEFILE(ByVal filePath As Variant) As Variant
    On Error GoTo EH: UDF_FS_DELETEFILE = DeleteFile(filePath): Exit Function
EH: UDF_FS_DELETEFILE = CVErr(xlErrValue)
End Function

Public Function UDF_FS_COPYFILE(ByVal sourcePath As Variant, _
    ByVal destPath As Variant, Optional ByVal overwrite As Variant = True) As Variant
    On Error GoTo EH: UDF_FS_COPYFILE = CopyFileSafe(sourcePath, destPath, overwrite): Exit Function
EH: UDF_FS_COPYFILE = CVErr(xlErrValue)
End Function

Public Function UDF_FS_TEMPFOLDER() As Variant
    On Error GoTo EH: UDF_FS_TEMPFOLDER = GetTempFolder(): Exit Function
EH: UDF_FS_TEMPFOLDER = CVErr(xlErrValue)
End Function

Public Function UDF_FS_SPECIALFOLDER(ByVal folderType As Variant) As Variant
    On Error GoTo EH: UDF_FS_SPECIALFOLDER = GetSpecialFolder(folderType): Exit Function
EH: UDF_FS_SPECIALFOLDER = CVErr(xlErrValue)
End Function

Public Function UDF_FS_ISPATHVALID(ByVal path As Variant) As Variant
    On Error GoTo EH: UDF_FS_ISPATHVALID = IsPathValid(path): Exit Function
EH: UDF_FS_ISPATHVALID = CVErr(xlErrValue)
End Function

Public Function UDF_FS_LISTFILES(ByVal folder As Variant, _
    Optional ByVal pattern As Variant = "*.*", Optional ByVal recurse As Variant = False) As Variant
    On Error GoTo EH: UDF_FS_LISTFILES = ListFiles(folder, pattern, recurse): Exit Function
EH: UDF_FS_LISTFILES = CVErr(xlErrValue)
End Function

Public Function UDF_FS_LISTFOLDERS(ByVal folder As Variant, Optional ByVal recurse As Variant = False) As Variant
    On Error GoTo EH: UDF_FS_LISTFOLDERS = ListFolders(folder, recurse): Exit Function
EH: UDF_FS_LISTFOLDERS = CVErr(xlErrValue)
End Function

Public Function UDF_FS_COPYFOLDER(ByVal sourcePath As Variant, _
    ByVal destPath As Variant, Optional ByVal overwrite As Variant = True) As Variant
    On Error GoTo EH: UDF_FS_COPYFOLDER = CopyFolder(sourcePath, destPath, overwrite): Exit Function
EH: UDF_FS_COPYFOLDER = CVErr(xlErrValue)
End Function

Public Function UDF_FS_DELETEFOLDER(ByVal folderPath As Variant) As Variant
    On Error GoTo EH: UDF_FS_DELETEFOLDER = DeleteFolder(folderPath): Exit Function
EH: UDF_FS_DELETEFOLDER = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例
'=====================================================================
' FileExists("C:\Windows\notepad.exe")                     → True
' GetFileName("C:\Data\report.xlsx")                       → "report.xlsx"
' GetFileSizeFmt("C:\Windows\notepad.exe")                 → "237.50 KB"
' GetTempFolder()                                          → "C:\Users\...\AppData\Local\Temp"
' GetSpecialFolder("desktop")                              → "C:\Users\...\Desktop"
' EnsureFolder "C:\MyApp\Logs\2025"                        → True
' AppendTextFile("C:\log.txt", "new line")
' Set info = GetDriveInfo("C")                             → info("total"), info("free")

'=====================================================================