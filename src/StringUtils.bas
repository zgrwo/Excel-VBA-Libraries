Option Explicit

'==============================================================================
' Module:       StringUtils
' Purpose:      String processing: encode, decode, distance, UUID, URL
' Layer:        Text
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       70 functions/subs
'==============================================================================

Private DP As New DictProxy

'=====================================================================
' StringUtils.bas — VBA 字符串工具集 (编码/校验/距离/生成/UUID)
'
' 工作表函数 (UDF_STR_*):
'   UDF_STR_EXTRACTBETWEEN    — 取两个分隔符之间的文本
'   UDF_STR_REVERSESTRING     — 字符串反转
'   UDF_STR_COUNTSUBSTRING    — 统计子串出现次数
'   UDF_STR_STARTSWITH        — 判断前缀
'   UDF_STR_ENDSWITH          — 判断后缀
'   UDF_STR_LEFTOF            — 取分隔符左侧文本
'   UDF_STR_RIGHTOF           — 取分隔符右侧文本
'   UDF_STR_TEXTJOIN          — 带分隔符连接 (兼容旧版 Excel)
'   UDF_STR_NTHWORD           — 提取第 n 个词
'   UDF_STR_COMMONPREFIX      — 最长公共前缀
'   UDF_STR_PADLEFT           — 左侧填充对齐
'   UDF_STR_PADRIGHT          — 右侧填充对齐
'   UDF_STR_TRUNCATE          — 截断加后缀
'   UDF_STR_NORMALIZEWHITESPACE — 折叠多余空白
'   UDF_STR_REMOVECHARS       — 移除指定字符
'   UDF_STR_KEEPCHARS         — 保留指定字符
'   UDF_STR_TOTITLECASE       — 智能首字母大写 (Mc/Mac/O')
'   UDF_STR_REMOVEDIACRITICS  — 移除变音符号
'   UDF_STR_SLUGIFY           — URL 友好化
'   UDF_STR_ISNULLOREMPTY     — 空值/空串判断
'   UDF_STR_ISNULLORWHITESPACE — 空值/空白判断
'   UDF_STR_LEVENSHTEIN       — 编辑距离
'   UDF_STR_SOUNDEX           — 发音哈希
'   UDF_STR_URLENCODE         — URL 编码
'   UDF_STR_URLDECODE         — URL 解码
'   UDF_STR_BASE64ENCODE      — Base64 编码
'   UDF_STR_BASE64DECODE      — Base64 解码
'   UDF_STR_HTMLENCODE        — HTML 实体编码
'   UDF_STR_HTMLDECODE        — HTML 实体解码
'   UDF_STR_RANDOMSTRING      — 生成随机字符串
'   UDF_STR_ISEMAIL           — 验证邮箱格式
'   UDF_STR_ISURL             — 验证 URL 格式
'   UDF_STR_UUID              — 生成 UUID
'   UDF_STR_COALESCE          — 返回第一个非空值
'   UDF_STR_REPEAT            — 重复字符串
'
' 内部函数 (PascalCase，供 VBA 代码调用):
'   ExtractBetween    — 取两个分隔符之间的文本
'   ReverseString     — 字符串反转
'   CountSubstring    — 统计子串出现次数
'   StartsWith        — 判断前缀
'   EndsWith          — 判断后缀
'   LeftOf            — 取分隔符左侧文本
'   RightOf           — 取分隔符右侧文本
'   TextJoin          — 带分隔符连接 (兼容旧版 Excel)
'   NthWord           — 提取第 n 个词
'   CommonPrefix      — 最长公共前缀
'   PadLeft / PadRight — 填充对齐
'   Truncate          — 截断加后缀
'   NormalizeWhitespace — 折叠多余空白
'   RemoveChars        — 移除指定字符
'   KeepChars          — 保留指定字符
'   ToTitleCase        — 智能首字母大写 (Mc/Mac/O')
'   RemoveDiacritics   — 移除变音符号
'   Slugify            — URL 友好化
'   IsNullOrEmpty      — 空值/空串判断
'   IsNullOrWhitespace — 空值/空白判断
'   LevenshteinDistance — 编辑距离
'   Soundex            — 发音哈希
'   URLEncode / URLDecode — URL 编解码
'   Base64Encode / Base64Decode — Base64 编解码
'   HTMLEncode / HTMLDecode — HTML 实体编解码
'   IsEmail           — 校验邮箱格式
'   IsUrl             — 校验 URL 格式
'   Coalesce          — 返回第一个非空值 (类似 SQL COALESCE)
'   UUID              — 生成 RFC 4122 v4 UUID
'   Repeat            — 重复字符串 n 次
'   RandomString       — 生成随机字符串
'=====================================================================
Private Const ERR_NOT_AVAIL As Long = vbObjectError + 1001  ' Reserved
Private Const B64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"


'=============================================================================
' ExtractBetween — 取两个分隔符之间的文本
'
' 参数:
'   text        - 源文本
'   leftDelim   - 左侧分隔符
'   rightDelim  - 右侧分隔符
'   nth         - 第 n 次出现的 leftDelim (1-based, 默认 1)
'   includeDelim - 结果是否包含分隔符 (默认 False)
'=============================================================================
Public Function ExtractBetween( _
    ByVal text As String, _
    ByVal leftDelim As String, _
    ByVal rightDelim As String, _
    Optional ByVal nth As Long = 1, _
    Optional ByVal includeDelim As Boolean = False) As String

    Dim posL As Long, posR As Long
    Dim i As Long, startPos As Long

    If Len(text) = 0 Or Len(leftDelim) = 0 Or Len(rightDelim) = 0 Then
        ExtractBetween = ""
        Exit Function
    End If
    If nth < 1 Then nth = 1

    startPos = 1
    For i = 1 To nth
        posL = InStr(startPos, text, leftDelim, vbTextCompare)
        If posL = 0 Then
            ExtractBetween = ""
            Exit Function
        End If
        If i < nth Then startPos = posL + Len(leftDelim)
    Next i

    posR = InStr(posL + Len(leftDelim), text, rightDelim, vbTextCompare)
    If posR = 0 Then
        ExtractBetween = ""
        Exit Function
    End If

    If includeDelim Then
        ExtractBetween = Mid$(text, posL, posR - posL + Len(rightDelim))
    Else
        ExtractBetween = Mid$(text, posL + Len(leftDelim), posR - posL - Len(leftDelim))
    End If
End Function

'=============================================================================

Private Function IsAsciiUpper(ByVal ch As String) As Boolean
    Dim code As Long: code = AscW(ch)
    IsAsciiUpper = (code >= 65 And code <= 90)
End Function

Private Function IsAsciiLower(ByVal ch As String) As Boolean
    Dim code As Long: code = AscW(ch)
    IsAsciiLower = (code >= 97 And code <= 122)
End Function


'=============================================================================
' RemoveDiacritics — 移除变音符号
'
' 例: "café résumé naïve" → "cafe resume naive"
' 覆盖: Latin-1 Supplement (U+00C0–U+00FF)
'=============================================================================
Public Function RemoveDiacritics(ByVal text As String) As String
    Dim i As Long, ch As String
    Dim code As Long

    If Len(text) = 0 Then
        RemoveDiacritics = ""
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(0 To Len(text) - 1) As String
    Dim pIdx As Long: pIdx = 0
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch)
        If code >= 192 And code <= 383 Then  ' Latin-1 Supplement + Latin Extended-A (#11)
            parts(pIdx) = DiacriticMap(code)
        Else
            parts(pIdx) = ch
        End If
        pIdx = pIdx + 1
    Next i
    RemoveDiacritics = Join(parts, "")
End Function

Private Function DiacriticMap(ByVal code As Long) As String
    Select Case code
        ' Latin-1 补充区 (U+00C0 - U+00FF)
        Case 192 To 197: DiacriticMap = "A"
        Case 198:         DiacriticMap = "AE"
        Case 199:         DiacriticMap = "C"
        Case 200 To 203: DiacriticMap = "E"
        Case 204 To 207: DiacriticMap = "I"
        Case 208:         DiacriticMap = "D"
        Case 209:         DiacriticMap = "N"
        Case 210 To 214: DiacriticMap = "O"
        Case 216:         DiacriticMap = "O"
        Case 217 To 220: DiacriticMap = "U"
        Case 221:         DiacriticMap = "Y"
        Case 222:         DiacriticMap = "Th"
        Case 223:         DiacriticMap = "ss"
        Case 224 To 229: DiacriticMap = "a"
        Case 230:         DiacriticMap = "ae"
        Case 231:         DiacriticMap = "c"
        Case 232 To 235: DiacriticMap = "e"
        Case 236 To 239: DiacriticMap = "i"
        Case 240:         DiacriticMap = "d"
        Case 241:         DiacriticMap = "n"
        Case 242 To 246: DiacriticMap = "o"
        Case 248:         DiacriticMap = "o"
        Case 249 To 252: DiacriticMap = "u"
        Case 253, 255:    DiacriticMap = "y"
        Case 254:         DiacriticMap = "th"
        ' Latin 扩展A区 (U+0100 - U+017F)
        Case 256, 258:    DiacriticMap = "A"  ' A-长音符, A-短音符
        Case 257, 259:    DiacriticMap = "a"
        Case 260:         DiacriticMap = "A"  ' A-反尾形符 (波兰语)
        Case 261:         DiacriticMap = "a"
        Case 262, 264, 266, 268: DiacriticMap = "C"  ' C-锐音符, C-抑扬符, C-上点, C-楔形符
        Case 263, 265, 267, 269: DiacriticMap = "c"
        Case 270:         DiacriticMap = "D"  ' D-楔形符
        Case 271:         DiacriticMap = "d"
        Case 272:         DiacriticMap = "D"  ' D-横杠
        Case 273:         DiacriticMap = "d"
        Case 274, 276, 278: DiacriticMap = "E"  ' E-长音符, E-短音符, E-反尾形符
        Case 275, 277, 279: DiacriticMap = "e"
        Case 280:         DiacriticMap = "E"  ' E-反尾形符 (波兰语)
        Case 281:         DiacriticMap = "e"
        Case 282:         DiacriticMap = "E"  ' E-楔形符
        Case 283:         DiacriticMap = "e"  ' e-楔形符
        Case 284, 286, 288, 290: DiacriticMap = "G"  ' G-抑扬符, G-短音符, G-上点, G-逗号
        Case 285, 287, 289, 291: DiacriticMap = "g"
        Case 292:         DiacriticMap = "H"  ' H-抑扬符
        Case 293:         DiacriticMap = "h"
        Case 296, 298, 300: DiacriticMap = "I"  ' I-波浪符, I-长音符, I-短音符
        Case 297, 299, 301: DiacriticMap = "i"
        Case 302:         DiacriticMap = "I"  ' I-反尾形符
        Case 303:         DiacriticMap = "i"
        Case 304:         DiacriticMap = "I"  ' I-上点 (土耳其语)
        Case 305:         DiacriticMap = "i"  ' 无点 i (土耳其语)
        Case 306:         DiacriticMap = "IJ" ' IJ 合字
        Case 307:         DiacriticMap = "ij"
        Case 308:         DiacriticMap = "J"  ' J-抑扬符
        Case 309:         DiacriticMap = "j"
        Case 310:         DiacriticMap = "K"  ' K-逗号
        Case 311:         DiacriticMap = "k"
        Case 313, 317:    DiacriticMap = "L"  ' L-锐音符, L-楔形符
        Case 314, 318:    DiacriticMap = "l"
        Case 321:         DiacriticMap = "L"  ' L-横杠 (波兰语)
        Case 322:         DiacriticMap = "l"
        Case 323, 325, 327: DiacriticMap = "N"  ' N-锐音符, N-逗号, N-楔形符
        Case 324, 326, 328: DiacriticMap = "n"
        Case 332, 336:    DiacriticMap = "O"  ' O-长音符, O-双锐音符 (匈牙利语)
        Case 333, 337:    DiacriticMap = "o"
        Case 338:         DiacriticMap = "OE" ' OE 合字
        Case 339:         DiacriticMap = "oe"
        Case 340, 344:    DiacriticMap = "R"  ' R-锐音符, R-楔形符
        Case 341, 345:    DiacriticMap = "r"
        Case 346, 348, 350, 352: DiacriticMap = "S"  ' S-锐音符, S-抑扬符, S-逗号, S-楔形符
        Case 347, 349, 351, 353: DiacriticMap = "s"
        Case 354, 356:    DiacriticMap = "T"  ' T-逗号 (罗马尼亚语), T-楔形符
        Case 355, 357:    DiacriticMap = "t"
        Case 360, 362, 364: DiacriticMap = "U"  ' U-波浪符, U-长音符, U-短音符
        Case 361, 363, 365: DiacriticMap = "u"
        Case 366:         DiacriticMap = "U"  ' U-圆圈
        Case 367:         DiacriticMap = "u"
        Case 368, 370:    DiacriticMap = "U"  ' U-双锐音符, U-反尾形符
        Case 369, 371:    DiacriticMap = "u"
        Case 372:         DiacriticMap = "W"  ' W-抑扬符
        Case 373:         DiacriticMap = "w"
        Case 374, 376:    DiacriticMap = "Y"  ' Y-抑扬符, Y-分音符
        Case 375:         DiacriticMap = "y"
        Case 377, 379, 381: DiacriticMap = "Z"  ' Z-锐音符, Z-上点, Z-楔形符 (波兰语)
        Case 378, 380, 382: DiacriticMap = "z"
        ' Latin 扩展B区精选 (罗马尼亚语 S/T 带逗号)
        Case 536:         DiacriticMap = "S"  ' S-逗号 (罗马尼亚语)
        Case 537:         DiacriticMap = "s"
        Case 538:         DiacriticMap = "T"  ' T-逗号 (罗马尼亚语)
        Case 539:         DiacriticMap = "t"
        Case Else:        DiacriticMap = ChrW$(code)
    End Select
End Function

'=============================================================================
' ReverseString — 字符串反转
'=============================================================================
Public Function ReverseString(ByVal text As String) As String
    Dim i As Long, n As Long
    Dim srcPos As Long, dstPos As Long
    Dim code As Long, lo As Long
    Dim result As String

    n = Len(text)
    If n = 0 Then
        ReverseString = ""
        Exit Function
    End If

    result = Space$(n)
    srcPos = n
    dstPos = 1
    Do While srcPos >= 1
        code = AscW(Mid$(text, srcPos, 1)) And &HFFFF&
        ' 检测代理对：低代理前有高代理
        If code >= &HDC00& And code <= &HDFFF& And srcPos > 1 Then
            lo = AscW(Mid$(text, srcPos - 1, 1)) And &HFFFF&
            If lo >= &HD800& And lo <= &HDBFF& Then
                ' 反转时保持代理对顺序：先写入高位，再写入低位
                Mid$(result, dstPos, 1) = Mid$(text, srcPos - 1, 1)
                Mid$(result, dstPos + 1, 1) = Mid$(text, srcPos, 1)
                srcPos = srcPos - 2
                dstPos = dstPos + 2
            Else
                Mid$(result, dstPos, 1) = Mid$(text, srcPos, 1)
                srcPos = srcPos - 1
                dstPos = dstPos + 1
            End If
        Else
            Mid$(result, dstPos, 1) = Mid$(text, srcPos, 1)
            srcPos = srcPos - 1
            dstPos = dstPos + 1
        End If
    Loop
    ReverseString = result
End Function

'=============================================================================
' CountSubstring — 统计子串出现次数
'=============================================================================
Public Function CountSubstring( _
    ByVal text As String, _
    ByVal search As String, _
    Optional ByVal caseSensitive As Boolean = False) As Long

    Dim compare As VbCompareMethod
    Dim pos As Long, cnt As Long

    If Len(text) = 0 Or Len(search) = 0 Then
        CountSubstring = 0
        Exit Function
    End If

    If caseSensitive Then compare = vbBinaryCompare Else compare = vbTextCompare

    pos = InStr(1, text, search, compare)
    Do While pos > 0
        cnt = cnt + 1
        pos = InStr(pos + Len(search), text, search, compare)
    Loop
    CountSubstring = cnt
End Function

'=============================================================================
' StartsWith / EndsWith — 判断前缀 / 后缀
'=============================================================================
Public Function StartsWith( _
    ByVal text As String, _
    ByVal prefix As String, _
    Optional ByVal caseSensitive As Boolean = False) As Boolean

    If Len(text) = 0 Or Len(prefix) = 0 Then
        StartsWith = False
        Exit Function
    End If
    If Len(prefix) > Len(text) Then
        StartsWith = False
        Exit Function
    End If

    If caseSensitive Then
        StartsWith = (Left$(text, Len(prefix)) = prefix)
    Else
        StartsWith = (StrComp(Left$(text, Len(prefix)), prefix, vbTextCompare) = 0)
    End If
End Function

Public Function EndsWith( _
    ByVal text As String, _
    ByVal suffix As String, _
    Optional ByVal caseSensitive As Boolean = False) As Boolean

    If Len(text) = 0 Or Len(suffix) = 0 Then
        EndsWith = False
        Exit Function
    End If
    If Len(suffix) > Len(text) Then
        EndsWith = False
        Exit Function
    End If

    If caseSensitive Then
        EndsWith = (Right$(text, Len(suffix)) = suffix)
    Else
        EndsWith = (StrComp(Right$(text, Len(suffix)), suffix, vbTextCompare) = 0)
    End If
End Function

'=============================================================================
' LeftOf / RightOf — 取分隔符左侧 / 右侧文本
'
' 参数:
'   text      - 源文本
'   delimiter - 分隔符
'   nth       - 第 n 次出现 (1-based, 默认 1)
'=============================================================================
Public Function LeftOf( _
    ByVal text As String, _
    ByVal delimiter As String, _
    Optional ByVal nth As Long = 1) As String

    Dim pos As Long, startPos As Long, i As Long

    If Len(text) = 0 Or Len(delimiter) = 0 Then
        LeftOf = ""
        Exit Function
    End If
    If nth < 1 Then nth = 1

    startPos = 1
    For i = 1 To nth
        pos = InStr(startPos, text, delimiter, vbTextCompare)
        If pos = 0 Then
            If i = nth Then
                LeftOf = text
            Else
                LeftOf = ""
            End If
            Exit Function
        End If
        If i < nth Then startPos = pos + Len(delimiter)
    Next i

    If pos > 0 Then
        LeftOf = Left$(text, pos - 1)
    Else
        LeftOf = text
    End If
End Function

Public Function RightOf( _
    ByVal text As String, _
    ByVal delimiter As String, _
    Optional ByVal nth As Long = 1, _
    Optional ByVal fromRight As Boolean = False) As String

    Dim pos As Long, startPos As Long, i As Long, dl As Long

    If Len(text) = 0 Or Len(delimiter) = 0 Then
        RightOf = ""
        Exit Function
    End If
    If nth < 1 Then nth = 1

    dl = Len(delimiter)

    If fromRight Then
        startPos = Len(text)
        For i = 1 To nth
            pos = InStrRev(text, delimiter, startPos, vbTextCompare)
            If pos = 0 Then
                RightOf = ""
                Exit Function
            End If
            If i < nth Then startPos = pos - 1
        Next i
        RightOf = Mid$(text, pos + dl)
    Else
        startPos = 1
        For i = 1 To nth
            pos = InStr(startPos, text, delimiter, vbTextCompare)
            If pos = 0 Then
                RightOf = ""
                Exit Function
            End If
            If i < nth Then startPos = pos + dl
        Next i
        RightOf = Mid$(text, pos + dl)
    End If
End Function

'=============================================================================
' TextJoin — 带分隔符连接
'
' 兼容 Excel 2016 以下没有 TEXTJOIN 的版本
'=============================================================================
Public Function TextJoin( _
    ByVal delimiter As String, _
    ParamArray args() As Variant) As String

    Dim i As Long, j As Long
    Dim item As Variant
    Dim first As Boolean
    Dim parts() As String
    Dim pIdx As Long: pIdx = 0

    Dim cap As Long: cap = 256
    ReDim parts(0 To cap - 1) As String
    first = True
    For i = LBound(args) To UBound(args)
        If IsArray(args(i)) Then
            For j = LBound(args(i)) To UBound(args(i))
                If IsError(args(i)(j)) Or IsNull(args(i)(j)) Or IsEmpty(args(i)(j)) Then
                    ' 跳过错误值、Null 和空值
                Else
                    If pIdx + 2 > cap Then
                        cap = cap * 2
                        ReDim Preserve parts(0 To cap - 1)
                    End If
                    If Not first Then
                        parts(pIdx) = delimiter: pIdx = pIdx + 1
                    End If
                    parts(pIdx) = CStr(args(i)(j)): pIdx = pIdx + 1
                    first = False
                End If
            Next j
        Else
            If IsError(args(i)) Or IsNull(args(i)) Or IsEmpty(args(i)) Then
                ' 跳过
            Else
                If pIdx + 2 > cap Then
                    cap = cap * 2
                    ReDim Preserve parts(0 To cap - 1)
                End If
                If Not first Then
                    parts(pIdx) = delimiter: pIdx = pIdx + 1
                End If
                parts(pIdx) = CStr(args(i)): pIdx = pIdx + 1
                first = False
            End If
        End If
    Next i
    If pIdx > 0 Then
        ReDim Preserve parts(0 To pIdx - 1)
        TextJoin = Join(parts, "")
    End If
End Function

'=============================================================================
' NthWord — 提取第 n 个词
'
' 参数:
'   text      - 源文本
'   n         - 第 n 个词 (1-based), -1 表示最后一个
'   delimiter - 分隔符 (默认 " ")
'
' 已知限制: VBA Split 函数仅支持单字符分隔符。若需多字符分隔符，
' 请使用 RegexUtils.RegexSplit 替代。
'=============================================================================
Public Function NthWord( _
    ByVal text As String, _
    ByVal n As Long, _
    Optional ByVal delimiter As String = " ") As String

    Dim words() As String
    Dim cnt As Long

    If Len(text) = 0 Then
        NthWord = ""
        Exit Function
    End If

    words = Split(text, delimiter)

    If n > 0 Then
        cnt = UBound(words) - LBound(words) + 1
        If n <= cnt Then
            NthWord = words(LBound(words) + n - 1)
        Else
            NthWord = ""
        End If
    ElseIf n < 0 Then
        cnt = UBound(words) - LBound(words) + 1
        If -n <= cnt Then
            NthWord = words(UBound(words) + n + 1)
        Else
            NthWord = ""
        End If
    Else
        NthWord = ""
    End If
End Function

'=============================================================================
' CommonPrefix — 最长公共前缀
'=============================================================================
Public Function CommonPrefix( _
    ByVal a As String, _
    ByVal b As String, _
    Optional ByVal caseSensitive As Boolean = False) As String

    Dim i As Long, maxLen As Long

    If Len(a) = 0 Or Len(b) = 0 Then
        CommonPrefix = ""
        Exit Function
    End If

    Dim originalA As String
    If Not caseSensitive Then
        originalA = a
        a = LCase$(a)
        b = LCase$(b)
    End If

    If Len(a) < Len(b) Then maxLen = Len(a) Else maxLen = Len(b)

    For i = 1 To maxLen
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then Exit For
    Next i
    If caseSensitive Then
        CommonPrefix = Left$(a, i - 1)
    Else
        CommonPrefix = Left$(originalA, i - 1)
    End If
End Function

'=============================================================================

'=============================================================================
' PadLeft / PadRight — 填充对齐
'=============================================================================
Public Function PadLeft( _
    ByVal text As String, _
    ByVal totalWidth As Long, _
    Optional ByVal padChar As String = " ") As String

    If Len(padChar) = 0 Then padChar = " "
    If Len(padChar) > 1 Then padChar = Left$(padChar, 1)
    If Len(text) >= totalWidth Then
        PadLeft = text
        Exit Function
    End If
    PadLeft = String$(totalWidth - Len(text), padChar) & text
End Function

Public Function PadRight( _
    ByVal text As String, _
    ByVal totalWidth As Long, _
    Optional ByVal padChar As String = " ") As String

    If Len(padChar) = 0 Then padChar = " "
    If Len(padChar) > 1 Then padChar = Left$(padChar, 1)
    If Len(text) >= totalWidth Then
        PadRight = text
        Exit Function
    End If
    PadRight = text & String$(totalWidth - Len(text), padChar)
End Function

'=============================================================================
' Repeat — 重复字符串 n 次
'
' 与 VBA String$(n, char) 不同，Repeat 接受完整字符串。
' 例: Repeat("ab", 3) → "ababab"
'=============================================================================
Public Function Repeat(ByVal text As String, ByVal n As Long) As String
    If n <= 0 Then Repeat = "": Exit Function
    If Len(text) = 0 Then Repeat = "": Exit Function
    If n = 1 Then Repeat = text: Exit Function
    ' 使用 Join 避免 O(n²) 字符串拼接
    Dim parts() As String, i As Long
    ReDim parts(0 To n - 1)
    For i = 0 To n - 1: parts(i) = text: Next i
    Repeat = Join(parts, "")
End Function

'=============================================================================
' NormalizeWhitespace — 将连续空白字符折叠为单个空格
'
' 包括: 空格、制表符、回车、换行 → 统一替换为 " "
'=============================================================================
Public Function NormalizeWhitespace(ByVal text As String) As String
    Dim i As Long, n As Long
    Dim ch As String, result As String
    Dim prevSpace As Boolean

    n = Len(text)
    If n = 0 Then
        NormalizeWhitespace = ""
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(0 To n - 1) As String
    Dim pIdx As Long: pIdx = 0
    For i = 1 To n
        ch = Mid$(text, i, 1)
        If IsWhitespace(ch) Then
            If Not prevSpace Then
                parts(pIdx) = " ": pIdx = pIdx + 1
                prevSpace = True
            End If
        Else
            parts(pIdx) = ch: pIdx = pIdx + 1
            prevSpace = False
        End If
    Next i
    If pIdx > 0 Then result = Join(parts, "")

    ' 去除首尾空白
    If Len(result) > 0 Then
        If IsWhitespace(Left$(result, 1)) Then result = Mid$(result, 2)
        If Len(result) > 0 Then
            If IsWhitespace(Right$(result, 1)) Then result = Left$(result, Len(result) - 1)
        End If
    End If

    NormalizeWhitespace = result
End Function

Private Function IsWhitespace(ByVal ch As String) As Boolean
    Select Case AscW(ch)
        Case 9, 10, 13, 32, 160: IsWhitespace = True
        Case Else: IsWhitespace = False
    End Select
End Function

'=============================================================================
' RemoveChars — 移除指定字符集中的字符
'=============================================================================
Public Function RemoveChars( _
    ByVal text As String, _
    ByVal charsToRemove As String) As String

    Dim i As Long, nText As Long
    Dim ch As String
    Dim charDict As Object

    nText = Len(text)
    If nText = 0 Then
        RemoveChars = ""
        Exit Function
    End If
    If Len(charsToRemove) = 0 Then
        RemoveChars = text
        Exit Function
    End If

    ' 尝试创建 Dictionary，失败则使用纯 VBA 回退
    On Error Resume Next
    Set charDict = DP.Create()
    On Error GoTo 0

    Dim parts() As String
    ReDim parts(0 To nText - 1) As String
    Dim pIdx As Long: pIdx = 0
    If charDict Is Nothing Then
        ' 回退: 逐字符 InStr 检查
        For i = 1 To nText
            ch = Mid$(text, i, 1)
            If InStr(1, charsToRemove, ch, vbBinaryCompare) = 0 Then
                parts(pIdx) = ch: pIdx = pIdx + 1
            End If
        Next i
    Else
        For i = 1 To Len(charsToRemove)
            ch = Mid$(charsToRemove, i, 1)
            If Not charDict.Exists(ch) Then charDict.Add ch, True
        Next i
        For i = 1 To nText
            ch = Mid$(text, i, 1)
            If Not charDict.Exists(ch) Then parts(pIdx) = ch: pIdx = pIdx + 1
        Next i
    End If
    If pIdx > 0 Then
        RemoveChars = Join(parts, "")
    End If
    Set charDict = Nothing
End Function

'=============================================================================
' KeepChars — 只保留允许的字符，其他全部移除
'=============================================================================
Public Function KeepChars( _
    ByVal text As String, _
    ByVal allowedChars As String) As String

    Dim i As Long, nText As Long
    Dim ch As String
    Dim charDict As Object

    nText = Len(text)
    If nText = 0 Then
        KeepChars = ""
        Exit Function
    End If
    If Len(allowedChars) = 0 Then
        KeepChars = ""
        Exit Function
    End If

    ' 尝试创建 Dictionary，失败则使用纯 VBA 回退
    On Error Resume Next
    Set charDict = DP.Create()
    On Error GoTo 0

    Dim parts() As String
    ReDim parts(0 To nText - 1) As String
    Dim pIdx As Long: pIdx = 0
    If charDict Is Nothing Then
        For i = 1 To nText
            ch = Mid$(text, i, 1)
            If InStr(1, allowedChars, ch, vbBinaryCompare) > 0 Then
                parts(pIdx) = ch: pIdx = pIdx + 1
            End If
        Next i
    Else
        For i = 1 To Len(allowedChars)
            ch = Mid$(allowedChars, i, 1)
            If Not charDict.Exists(ch) Then charDict.Add ch, True
        Next i
        For i = 1 To nText
            ch = Mid$(text, i, 1)
            If charDict.Exists(ch) Then parts(pIdx) = ch: pIdx = pIdx + 1
        Next i
    End If
    If pIdx > 0 Then
        KeepChars = Join(parts, "")
    End If
    Set charDict = Nothing
End Function

'=============================================================================
' ToTitleCase — 智能首字母大写
'
' 处理 Mc/Mac/O' 等特殊情况，比 Excel PROPER 更准确
'=============================================================================
Public Function ToTitleCase(ByVal text As String) As String
    Dim i As Long, n As Long
    Dim ch As String, prev As String
    Dim capitalize As Boolean

    If Len(text) = 0 Then
        ToTitleCase = ""
        Exit Function
    End If

    text = LCase$(text)
    n = Len(text)
    capitalize = True
    Dim result As String
    Dim parts() As String
    ReDim parts(0 To n - 1) As String
    Dim pIdx As Long: pIdx = 0
    For i = 1 To n
        ch = Mid$(text, i, 1)
        If capitalize Then
            parts(pIdx) = UCase$(ch): pIdx = pIdx + 1
            capitalize = False
        Else
            parts(pIdx) = ch: pIdx = pIdx + 1
        End If

        ' 判断下一个字母是否需要大写
        Select Case AscW(ch)
            Case 32, 9, 10, 13, 45, 46, 47, 38 ' 空格、制表、换行、-、.、/、&
                capitalize = True
            Case 39 ' 单引号: O'Brien → O'Brien; 同时避免 don't → Don'T
                If i >= 2 Then
                    If StrComp(Mid$(text, i - 1, 1), "O", vbTextCompare) = 0 Then
                        capitalize = True
                    End If
                End If
            Case Else
                capitalize = False
        End Select
    Next i

    ' 修复 Mac/Mc 模式: "Macdonald" → "MacDonald", "Mcdonald" → "McDonald"
    result = FixNamePrefix(Join(parts, ""))

    ToTitleCase = result
End Function

' FixNamePrefix: 修复 Mac/Mc 姓氏大写 (Macdonald→MacDonald, Mcdonald→McDonald)
' 注: 输入来自 ToTitleCase 的 Title Case 文本（非小写），通过 vbTextCompare
' 比较实现大小写不敏感匹配，不依赖输入大小写假设
Private Function FixNamePrefix(ByVal text As String) As String
    Dim i As Long, n As Long, adv As Long, ch As String, prev As String
    n = Len(text): i = 1
    Do While i <= n
        adv = 1
        ' 检查前一个字符是否为非字母 (词边界), 防止 "machine" → "MacHine" (#12)
        If i > 1 Then prev = Mid$(text, i - 1, 1) Else prev = " "
        If i + 2 <= n Then
            If (prev < "A" Or prev > "z" Or (prev > "Z" And prev < "a")) Then
                If StrComp(Mid$(text, i, 3), "Mac", vbTextCompare) = 0 Then
                    If i + 3 <= n Then
                        ch = Mid$(text, i + 3, 1)
                        If ch >= "a" And ch <= "z" Then Mid$(text, i + 3, 1) = UCase$(ch)
                    End If
                    adv = 3
                End If
            End If
        End If
        If adv = 1 And i + 1 <= n Then
            If (prev < "A" Or prev > "z" Or (prev > "Z" And prev < "a")) Then
                If StrComp(Mid$(text, i, 2), "Mc", vbTextCompare) = 0 Then
                    If i + 2 <= n Then
                        ch = Mid$(text, i + 2, 1)
                        If ch >= "a" And ch <= "z" Then Mid$(text, i + 2, 1) = UCase$(ch)
                    End If
                    adv = 2
                End If
            End If
        End If
        i = i + adv
    Loop
    FixNamePrefix = text
End Function

'=============================================================================
' IsNullOrEmpty / IsNullOrWhitespace
'=============================================================================
Public Function IsNullOrEmpty(ByVal text As Variant) As Boolean
    If IsObject(text) Then
        If TypeOf text Is Range Then text = text.Value
    End If
    ' Range.Value may return a 2D array — extract scalar if single cell
    If IsArray(text) Then
        On Error Resume Next
        text = text(LBound(text, 1), LBound(text, 2))
        On Error GoTo 0
    End If

    If IsNull(text) Then
        IsNullOrEmpty = True
    ElseIf IsEmpty(text) Then
        IsNullOrEmpty = True
    ElseIf VarType(text) = vbError Then
        IsNullOrEmpty = True
    ElseIf Len(CStr(text)) = 0 Then
        IsNullOrEmpty = True
    Else
        IsNullOrEmpty = False
    End If
End Function

Public Function IsNullOrWhitespace(ByVal text As Variant) As Boolean
    Dim i As Long, s As String

    If IsNullOrEmpty(text) Then
        IsNullOrWhitespace = True
        Exit Function
    End If
    s = CStr(text)
    For i = 1 To Len(s)
        If Not IsWhitespace(Mid$(s, i, 1)) Then
            IsNullOrWhitespace = False
            Exit Function
        End If
    Next i
    IsNullOrWhitespace = True
End Function

'=============================================================================
' IsEmail — 校验是否为合法的电子邮件地址格式
'=============================================================================
Public Function IsEmail(ByVal text As Variant) As Boolean
    If IsNull(text) Then Exit Function
    Dim s As String: s = CStr(text)
    Dim atPos As Long, dotPos As Long
    If Len(s) < 5 Then Exit Function
    atPos = InStr(s, "@")
    If atPos <= 1 Or atPos = Len(s) Then Exit Function
    dotPos = InStrRev(s, ".")
    If dotPos <= atPos + 1 Or dotPos = Len(s) Then Exit Function
    If InStr(atPos + 1, s, "@") > 0 Then Exit Function  ' 多个 @ → 无效
    IsEmail = True
End Function

'=============================================================================
' IsUrl — 校验是否为合法的 URL 格式
'=============================================================================
Public Function IsUrl(ByVal text As Variant) As Boolean
    If IsNull(text) Then Exit Function
    Dim s As String: s = CStr(text)
    If Len(s) < 8 Then Exit Function
    If Left$(LCase$(s), 7) = "http://" Then IsUrl = True: Exit Function
    If Left$(LCase$(s), 8) = "https://" Then IsUrl = True: Exit Function
End Function

'=============================================================================
' Coalesce — 返回参数列表中第一个非空值 (类似 SQL COALESCE)
'=============================================================================
Public Function Coalesce(ParamArray values() As Variant) As Variant
    Dim i As Long, v As Variant
    For i = LBound(values) To UBound(values)
        v = values(i)
        If Not IsNull(v) And Not IsEmpty(v) And Not IsError(v) Then
            If VarType(v) <> vbString Or Len(CStr(v)) > 0 Then
                Coalesce = v: Exit Function
            End If
        End If
    Next i
    Coalesce = Empty
End Function

'=============================================================================
' Truncate — 截断加后缀
'=============================================================================
Public Function Truncate( _
    ByVal text As String, _
    ByVal maxLength As Long, _
    Optional ByVal suffix As String = "...") As String

    If maxLength < 1 Then
        Truncate = ""
        Exit Function
    End If
    If Len(text) <= maxLength Then
        Truncate = text
        Exit Function
    End If
    If maxLength <= Len(suffix) Then
        Truncate = Left$(text, maxLength)
        Exit Function
    End If
    Truncate = Left$(text, maxLength - Len(suffix)) & suffix
End Function

'=============================================================================
' UUID — 生成符合 RFC 4122 的版本 4 UUID (36 字符)
'=============================================================================
Public Function UUID() As String
    Static seeded As Boolean
    If Not seeded Then
        Randomize
        seeded = True
    End If
    ' RFC 4122 version 4 UUID: 8-4-4-4-12, version=4, variant=8/9/a/b
    ' And mask clears nibble bits before Or sets version/variant
    UUID = LCase$(Hex4(Rnd * &H10000) & Hex4(Rnd * &H10000) & "-" & _
           Hex4(Rnd * &H10000) & "-" & _
           Hex4(Rnd * &H10000 And &H0FFF Or &H4000) & "-" & _
           Hex4(Rnd * &H10000 And &H3FFF Or &H8000) & "-" & _
           Hex4(Rnd * &H10000) & Hex4(Rnd * &H10000) & Hex4(Rnd * &H10000))
End Function

' Int() 向下取整 — 避免 CLng 银行家舍入导致 65535.5→65536 掩码归零偏差
Private Function Hex4(ByVal n As Double) As String
    Hex4 = Right$("0000" & Hex$((Int(n) And &HFFFF&)), 4)
End Function

'=============================================================================
' RandomString — 生成随机字符串
'=============================================================================
Public Function RandomString( _
    Optional ByVal length As Long = 8, _
    Optional ByVal charset As String = "") As String

    Const DEFAULT_ALPHANUM As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    Const DEFAULT_ALPHA As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Const DEFAULT_NUM As String = "0123456789"
    Const DEFAULT_UPPER As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Const DEFAULT_HEX As String = "0123456789ABCDEF"

    Dim chars As String
    Dim i As Long, idx As Long

    Static seeded As Boolean
    If Not seeded Then
        Randomize
        seeded = True
    End If

    If length < 0 Then length = 0

    If Len(charset) = 0 Then
        chars = DEFAULT_ALPHANUM
    ElseIf LCase$(charset) = "alpha" Then
        ' 含大小写字母; 仅大写用 "upper"
        chars = DEFAULT_ALPHA
    ElseIf LCase$(charset) = "num" Then
        chars = DEFAULT_NUM
    ElseIf LCase$(charset) = "upper" Then
        chars = DEFAULT_UPPER
    ElseIf LCase$(charset) = "hex" Then
        chars = DEFAULT_HEX
    Else
        chars = charset
    End If

    If length > 0 Then
        Dim parts() As String
        ReDim parts(0 To length - 1) As String
        For i = 1 To length
            idx = Int(Rnd() * Len(chars)) + 1
            parts(i - 1) = Mid$(chars, idx, 1)
        Next i
        RandomString = Join(parts, "")
    End If
End Function

'=============================================================================
' Slugify — URL 友好化
'
' 例: "Hello World! (v2.0)" → "hello-world-v20"
'=============================================================================
Public Function Slugify( _
    ByVal text As String, _
    Optional ByVal separator As String = "-") As String

    Dim i As Long, ch As String
    Dim code As Long
    Dim prevSep As Boolean
    Dim result As String

    If Len(separator) = 0 Then separator = "-"

    If Len(text) = 0 Then
        Slugify = ""
        Exit Function
    End If

    text = RemoveDiacritics(LCase$(Trim$(text)))
    Dim parts() As String
    ReDim parts(0 To Len(text) - 1) As String
    Dim pIdx As Long: pIdx = 0
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch)

        If (code >= 97 And code <= 122) Or (code >= 48 And code <= 57) Then
            parts(pIdx) = ch: pIdx = pIdx + 1
            prevSep = False
        ElseIf code = 45 Or code = 95 Or IsWhitespace(ch) Then
            ' 已有的连字符/下划线/空白 → 统一替换为分隔符
            If Not prevSep And pIdx > 0 Then
                parts(pIdx) = separator: pIdx = pIdx + 1
                prevSep = True
            End If
        Else
            ' 其他标点符号 → 直接丢弃
        End If
    Next i
    If pIdx > 0 Then result = Join(parts, "")

    ' 去除尾部连字符
    If Len(result) > 0 Then
        While Right$(result, Len(separator)) = separator
            result = Left$(result, Len(result) - Len(separator))
        Wend
    End If

    Slugify = result
End Function

'=============================================================================
' URLEncode / URLDecode — URL 编解码
'
' 空格编码为 "%20" (RFC 3986 标准)，解码同时兼容 "%20" 和 "+"
'=============================================================================
Public Function URLEncode(ByVal text As String) As String
    Dim i As Long, code As Long, lo As Long
    Dim parts() As String
    Dim pIdx As Long: pIdx = 0
    ReDim parts(0 To Len(text) * 6)

    i = 1
    Do While i <= Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& And i < Len(text) Then
            lo = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00& And lo <= &HDFFF& Then
                code = &H10000 + ((code - &HD800&) * &H400&) + (lo - &HDC00&)
                i = i + 1
            End If
        End If

        ' 字母数字和不安全列表中未包含的字符直接保留
        If (code >= 48 And code <= 57) Or _
           (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or _
           code = 45 Or code = 95 Or code = 46 Or code = 126 Then
            parts(pIdx) = ChrW$(code): pIdx = pIdx + 1
        ElseIf code = 32 Then
            parts(pIdx) = "%20": pIdx = pIdx + 1
        ElseIf code < &H80& Then
            parts(pIdx) = "%" & Right$("0" & Hex$(code), 2): pIdx = pIdx + 1
        ElseIf code < &H800& Then
            parts(pIdx) = "%" & Right$("0" & Hex$(&HC0& Or (code \ &H40&)), 2) & _
                                 "%" & Right$("0" & Hex$(&H80& Or (code And &H3F&)), 2)
            pIdx = pIdx + 1
        ElseIf code < &H10000 Then
            parts(pIdx) = "%" & Right$("0" & Hex$(&HE0& Or (code \ &H1000&)), 2) & _
                                 "%" & Right$("0" & Hex$(&H80& Or ((code \ &H40&) And &H3F&)), 2) & _
                                 "%" & Right$("0" & Hex$(&H80& Or (code And &H3F&)), 2)
            pIdx = pIdx + 1
        Else
            parts(pIdx) = "%" & Right$("0" & Hex$(&HF0& Or (code \ &H40000)), 2) & _
                                 "%" & Right$("0" & Hex$(&H80& Or ((code \ &H1000&) And &H3F&)), 2) & _
                                 "%" & Right$("0" & Hex$(&H80& Or ((code \ &H40&) And &H3F&)), 2) & _
                                 "%" & Right$("0" & Hex$(&H80& Or (code And &H3F&)), 2)
            pIdx = pIdx + 1
        End If
        i = i + 1
    Loop
    If pIdx > 0 Then
        ReDim Preserve parts(0 To pIdx - 1)
        URLEncode = Join(parts, "")
    End If
End Function

Public Function URLDecode(ByVal text As String) As String
    If Len(text) = 0 Then
        URLDecode = ""
        Exit Function
    End If

    Dim bytes() As Byte
    ReDim bytes(0 To Len(text) - 1)
    Dim byteIdx As Long: byteIdx = 0
    Dim i As Long, n As Long
    Dim ch As String, hexStr As String

    n = Len(text): i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If ch = "+" Then
            bytes(byteIdx) = 32
            byteIdx = byteIdx + 1
            i = i + 1
        ElseIf ch = "%" And i + 2 <= n Then
            hexStr = Mid$(text, i + 1, 2)
            If IsHexString(hexStr) Then
                ' 自定义 hex→byte 防止 CByte 符号溢出 (#13)
                Dim hi As Long: hi = InStr(1, "0123456789ABCDEF", UCase$(Mid$(hexStr, 1, 1)), vbBinaryCompare) - 1
                Dim lo As Long: lo = InStr(1, "0123456789ABCDEF", UCase$(Mid$(hexStr, 2, 1)), vbBinaryCompare) - 1
                If hi >= 0 And lo >= 0 Then
                    bytes(byteIdx) = CByte(hi * 16 + lo)
                Else
                    bytes(byteIdx) = CByte(AscW(ch))
                End If
                byteIdx = byteIdx + 1
                i = i + 3
            Else
                bytes(byteIdx) = CByte(AscW(ch))
                byteIdx = byteIdx + 1
                i = i + 1
            End If
        Else
            If AscW(ch) <= 255 Then
                bytes(byteIdx) = CByte(AscW(ch))
                byteIdx = byteIdx + 1
            End If
            i = i + 1
        End If
    Loop

    If byteIdx = 0 Then URLDecode = "": Exit Function
    ReDim Preserve bytes(0 To byteIdx - 1)

    ' 使用 ADODB.Stream 将 UTF-8 字节解码为字符串
    On Error Resume Next
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    If stream Is Nothing Then
        URLDecode = DecodeUTF8Bytes(bytes): Exit Function
    End If
    On Error GoTo 0

    On Error Resume Next
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "UTF-8"
    URLDecode = stream.ReadText
    stream.Close
    Set stream = Nothing
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        URLDecode = DecodeUTF8Bytes(bytes): Exit Function
    End If
    On Error GoTo 0
End Function

Private Function IsHexString(ByVal s As String) As Boolean
    Dim i As Long, code As Long
    For i = 1 To Len(s)
        code = AscW(UCase$(Mid$(s, i, 1)))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 70)) Then
            IsHexString = False
            Exit Function
        End If
    Next i
    IsHexString = True
End Function

'=============================================================================
' Base64Encode / Base64Decode — Base64 编解码
'
' 使用 ADODB.Stream 实现，依赖 Windows 内置组件
'=============================================================================

Public Function Base64Encode(ByVal text As String, _
                             Optional ByVal encoding As String = "UTF-8") As String

    If Len(text) = 0 Then
        Base64Encode = ""
        Exit Function
    End If

    ' 将文本转为字节数组
    Dim bytes() As Byte
    Dim stream As Object
    Dim bomIdx As Long

    On Error GoTo Fallback
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' 文本类型
    stream.Charset = encoding
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1 ' 二进制
    bytes = stream.Read
    stream.Close
    Set stream = Nothing

    ' 移除 UTF-8 BOM (EF BB BF)，如果存在
    If UBound(bytes) - LBound(bytes) > 2 Then
        If bytes(LBound(bytes)) = &HEF And _
           bytes(LBound(bytes) + 1) = &HBB And _
           bytes(LBound(bytes) + 2) = &HBF Then
            For bomIdx = LBound(bytes) To UBound(bytes) - 3
                bytes(bomIdx) = bytes(bomIdx + 3)
            Next bomIdx
            ReDim Preserve bytes(LBound(bytes) To UBound(bytes) - 3)
        End If
    End If

    Base64Encode = EncodeBase64Bytes(bytes)
    Exit Function

Fallback:
    If Not stream Is Nothing Then
        On Error Resume Next
        stream.Close
        On Error GoTo 0
    End If
    Base64Encode = EncodeBase64Bytes(EncodeUTF8Bytes(text))
End Function

Public Function Base64Decode(ByVal base64 As String, _
                             Optional ByVal encoding As String = "UTF-8") As String

    If Len(base64) = 0 Then
        Base64Decode = ""
        Exit Function
    End If

    Dim bytes() As Byte
    bytes = DecodeBase64Bytes(base64)

    ' 安全检测空/未分配数组（On Error Resume Next + LBound 探针）
    Dim isEmptyArr As Boolean
    On Error Resume Next
    Dim lbTest As Long: lbTest = LBound(bytes)
    If Err.Number <> 0 Then
        Err.Clear
        isEmptyArr = True
    Else
        Err.Clear
        isEmptyArr = (UBound(bytes) < LBound(bytes))
    End If
    On Error GoTo 0
    If isEmptyArr Then
        Base64Decode = ""
        Exit Function
    End If

    On Error GoTo Fallback
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' 二进制
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2 ' 文本类型
    stream.Charset = encoding
    Base64Decode = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function

Fallback:
    ' 确保流在回退前已关闭
    If Not stream Is Nothing Then
        On Error Resume Next
        stream.Close
        Set stream = Nothing
        On Error GoTo 0
    End If
    Base64Decode = DecodeUTF8Bytes(bytes)
End Function

Private Function EncodeBase64Bytes(ByRef bytes() As Byte) As String
    Dim i As Long, n As Long
    Dim b1 As Long, b2 As Long, b3 As Long
    Dim parts() As String
    Dim pIdx As Long: pIdx = 0

    n = UBound(bytes) - LBound(bytes)
    If n < 0 Then
        EncodeBase64Bytes = ""
        Exit Function
    End If
    ReDim parts(0 To (n \ 3 + 1) * 4)

    For i = LBound(bytes) To UBound(bytes) Step 3
        b1 = bytes(i)
        If i + 1 <= UBound(bytes) Then b2 = bytes(i + 1) Else b2 = -1
        If i + 2 <= UBound(bytes) Then b3 = bytes(i + 2) Else b3 = -1

        parts(pIdx) = Mid$(B64_CHARS, (b1 \ 4) + 1, 1): pIdx = pIdx + 1
        If b2 >= 0 Then
            parts(pIdx) = Mid$(B64_CHARS, ((b1 And 3) * 16 + b2 \ 16) + 1, 1): pIdx = pIdx + 1
        Else
            parts(pIdx) = Mid$(B64_CHARS, ((b1 And 3) * 16) + 1, 1): pIdx = pIdx + 1
            parts(pIdx) = "=": pIdx = pIdx + 1
            parts(pIdx) = "=": pIdx = pIdx + 1
            Exit For
        End If
        If b3 >= 0 Then
            parts(pIdx) = Mid$(B64_CHARS, ((b2 And 15) * 4 + b3 \ 64) + 1, 1): pIdx = pIdx + 1
            parts(pIdx) = Mid$(B64_CHARS, (b3 And 63) + 1, 1): pIdx = pIdx + 1
        Else
            parts(pIdx) = Mid$(B64_CHARS, ((b2 And 15) * 4) + 1, 1): pIdx = pIdx + 1
            parts(pIdx) = "=": pIdx = pIdx + 1
            Exit For
        End If
    Next i
    If pIdx > 0 Then
        ReDim Preserve parts(0 To pIdx - 1)
        EncodeBase64Bytes = Join(parts, "")
    End If
End Function

Private Function DecodeBase64Bytes(ByVal base64 As String) As Byte()
    Dim result() As Byte
    Dim n As Long, i As Long
    Dim padding As Long, outLen As Long
    Dim ch As String, val As Long

    ' 清理空白和换行
    base64 = RemoveChars(base64, vbCr & vbLf & " " & vbTab)
    n = Len(base64)

    If n = 0 Then
        Erase result
        DecodeBase64Bytes = result
        Exit Function
    End If

    If n < 2 Then
        Erase result
        DecodeBase64Bytes = result
        Exit Function
    End If

    padding = 0
    If Mid$(base64, n, 1) = "=" Then padding = padding + 1
    If n >= 2 Then
        If Mid$(base64, n - 1, 1) = "=" Then padding = padding + 1
    End If

    ' 计算输出长度：考虑末尾非4倍数字符（无填充 Base64 变体）
    Dim remChars As Long: remChars = n Mod 4
    outLen = (n \ 4) * 3 - padding
    If remChars > 0 Then outLen = outLen + remChars - 1
    If outLen <= 0 Then
        Erase result
        DecodeBase64Bytes = result
        Exit Function
    End If
    ReDim result(0 To outLen - 1)

    Dim outIdx As Long: outIdx = 0
    Dim sextet(0 To 3) As Long
    Dim si As Long

    For i = 1 To n
        ch = Mid$(base64, i, 1)
        If ch = "=" Then
            ' 遇到填充字符，输出已累积的部分字节
            If si >= 2 Then
                result(outIdx) = (sextet(0) * 4) Or (sextet(1) \ 16)
                If si >= 3 And outIdx + 1 < outLen Then
                    result(outIdx + 1) = ((sextet(1) And 15) * 16) Or (sextet(2) \ 4)
                End If
            End If
            Exit For
        End If

        val = InStr(1, B64_CHARS, ch, vbBinaryCompare) - 1
        ' 非法字符映射为 0 (等同 'A')，宽容解析，不中断解码
        If val < 0 Then val = 0

        sextet(si) = val
        si = si + 1

        If si = 4 Then
            result(outIdx) = (sextet(0) * 4) Or (sextet(1) \ 16)
            If outIdx + 1 < outLen Then
                result(outIdx + 1) = ((sextet(1) And 15) * 16) Or (sextet(2) \ 4)
            End If
            If outIdx + 2 < outLen Then
                result(outIdx + 2) = ((sextet(2) And 3) * 64) Or sextet(3)
            End If
            outIdx = outIdx + 3
            si = 0
        End If
    Next i

    ' 处理末尾未满 4 个 sextet 的部分组（无填充 Base64）
    If si >= 2 Then
        result(outIdx) = (sextet(0) * 4) Or (sextet(1) \ 16)
        If si >= 3 And outIdx + 1 < outLen Then
            result(outIdx + 1) = ((sextet(1) And 15) * 16) Or (sextet(2) \ 4)
        End If
    End If

    DecodeBase64Bytes = result
End Function

' UTF-8 编码辅助函数: 将 Unicode 字符串编码为 UTF-8 字节数组
Private Function EncodeUTF8Bytes(ByVal text As String) As Byte()
    Dim result() As Byte
    Dim n As Long: n = 0
    Dim i As Long, code As Long, ch As Long, lo As Long
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code >= &HD800 And code <= &HDBFF And i < Len(text) Then
            lo = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00 And lo <= &HDFFF Then
                code = &H10000 + ((code - &HD800) * &H400) + (lo - &HDC00)
                i = i + 1
            End If
        End If
        If code < &H80 Then
            n = n + 1
        ElseIf code < &H800 Then
            n = n + 2
        ElseIf code < &H10000 Then
            n = n + 3
        Else
            n = n + 4
        End If
    Next i
    If n = 0 Then Erase result: EncodeUTF8Bytes = result: Exit Function
    ReDim result(0 To n - 1)
    Dim idx As Long: idx = 0
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code >= &HD800 And code <= &HDBFF And i < Len(text) Then
            lo = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
            If lo >= &HDC00 And lo <= &HDFFF Then
                code = &H10000 + ((code - &HD800) * &H400) + (lo - &HDC00)
                i = i + 1
            End If
        End If
        If code < &H80 Then
            result(idx) = CByte(code): idx = idx + 1
        ElseIf code < &H800 Then
            result(idx) = CByte(&HC0 Or (code \ &H40)): idx = idx + 1
            result(idx) = CByte(&H80 Or (code And &H3F)): idx = idx + 1
        ElseIf code < &H10000 Then
            result(idx) = CByte(&HE0 Or (code \ &H1000)): idx = idx + 1
            result(idx) = CByte(&H80 Or ((code \ &H40) And &H3F)): idx = idx + 1
            result(idx) = CByte(&H80 Or (code And &H3F)): idx = idx + 1
        Else
            result(idx) = CByte(&HF0 Or (code \ &H40000)): idx = idx + 1
            result(idx) = CByte(&H80 Or ((code \ &H1000) And &H3F)): idx = idx + 1
            result(idx) = CByte(&H80 Or ((code \ &H40) And &H3F)): idx = idx + 1
            result(idx) = CByte(&H80 Or (code And &H3F)): idx = idx + 1
        End If
    Next i
    EncodeUTF8Bytes = result
End Function

Private Function DecodeUTF8Bytes(ByRef bytes() As Byte) As String
    Dim i As Long, code As Long, b As Long
    Dim lb As Long, ub As Long
    lb = LBound(bytes): ub = UBound(bytes)
    If ub < lb Then DecodeUTF8Bytes = "": Exit Function

    Dim parts() As String
    ReDim parts(0 To ub - lb + 1) As String
    Dim pIdx As Long: pIdx = 0
    i = lb
    Do While i <= ub
        b = bytes(i)
        If b < &H80& Then
            ' 1-byte: U+0000 - U+007F
            parts(pIdx) = ChrW$(b): pIdx = pIdx + 1
            i = i + 1
        ElseIf b < &HE0& Then
            ' 2-byte: U+0080 - U+07FF
            If i + 1 <= ub Then
                code = (CLng(b) And &H1F&) * &H40& + (CLng(bytes(i + 1)) And &H3F&)
                parts(pIdx) = ChrW$(code): pIdx = pIdx + 1
            End If
            i = i + 2
        ElseIf b < &HF0& Then
            ' 3-byte: U+0800 - U+FFFF
            If i + 2 <= ub Then
                code = (CLng(b) And &HF&) * &H1000& + _
                       (CLng(bytes(i + 1)) And &H3F&) * &H40& + _
                       (CLng(bytes(i + 2)) And &H3F&)
                parts(pIdx) = ChrW$(code): pIdx = pIdx + 1
            End If
            i = i + 3
        ElseIf b < &HF8& Then
            ' 4-byte: U+10000 - U+10FFFF (VBA 中的代理对)
            If i + 3 <= ub Then
                code = (CLng(b) And &H7&) * &H40000 + _
                       (CLng(bytes(i + 1)) And &H3F&) * &H1000& + _
                       (CLng(bytes(i + 2)) And &H3F&) * &H40& + _
                       (CLng(bytes(i + 3)) And &H3F&)
                code = code - &H10000
                parts(pIdx) = ChrW$(&HD800& + (code \ &H400&)): pIdx = pIdx + 1
                parts(pIdx) = ChrW$(&HDC00& + (code And &H3FF&)): pIdx = pIdx + 1
            End If
            i = i + 4
        Else
            i = i + 1  ' 跳过无效的后续字节
        End If
    Loop
    If pIdx > 0 Then DecodeUTF8Bytes = Join(parts, "")
End Function

'=============================================================================
' HTMLEncode / HTMLDecode — HTML 实体编解码
'=============================================================================
Public Function HTMLEncode(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    text = Replace(text, "'", "&#39;")
    HTMLEncode = text
End Function

Public Function HTMLDecode(ByVal text As String) As String
    ' Numeric entities first: prevents double-decoding of &amp;#NNN; sequences.
    ' Per HTML spec, entity-decoded output must not be re-scanned for further entities.
    ' If named entities were processed first, &amp;#65; → &#65; → A (wrong).
    ' With numeric first, &amp;#65; stays as &#65; because &amp; hasn't been decoded yet.

    ' 解析数字实体 &#nnn; &#xHH;
    Dim pos As Long, endPos As Long
    Dim entity As String, code As Long

    pos = InStr(text, "&#")
    Do While pos > 0
        code = 0
        endPos = InStr(pos + 2, text, ";")
        If endPos > 0 Then
            entity = Mid$(text, pos + 2, endPos - pos - 2)
            Err.Clear
            On Error Resume Next
            If Left$(entity, 1) = "x" Or Left$(entity, 1) = "X" Then
                If IsHexString(Mid$(entity, 2)) Then
                    code = CLng("&H" & Mid$(entity, 2))
                End If
            ElseIf IsNumeric(entity) Then
                code = CLng(entity)
            End If
            On Error GoTo 0

            If code > 0 Then
                If code <= 65535 Then
                    text = Left$(text, pos - 1) & ChrW$(code) & Mid$(text, endPos + 1)
                ElseIf code <= &H10FFFF Then
                    ' 代理对编码 (U+10000 - U+10FFFF)
                    code = code - &H10000
                    text = Left$(text, pos - 1) & _
                           ChrW$(&HD800& + (code \ &H400&)) & _
                           ChrW$(&HDC00& + (code And &H3FF&)) & _
                           Mid$(text, endPos + 1)
                End If
            End If
        End If
        pos = InStr(pos + 1, text, "&#")
    Loop

    ' 解码命名实体。&amp; 必须最后处理，避免 &amp;lt; → &lt; → <
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&amp;", "&")

    HTMLDecode = text
End Function

'=============================================================================
' Soundex — Soundex 发音哈希 (美式)
'
' 用于模糊匹配人名，返回 4 字符哈希
' 例: "Smith" → "S530", "Smythe" → "S530" (相同)
'=============================================================================
Public Function Soundex(ByVal text As String) As String
    Dim first As String
    Dim prevCode As Long
    Dim i As Long, code As Long
    Dim ch As String

    If Len(text) = 0 Then
        Soundex = ""
        Exit Function
    End If

    first = UCase$(Left$(text, 1))
    ReDim parts(0 To 3) As String
    Dim pIdx As Long: pIdx = 0
    parts(pIdx) = first: pIdx = pIdx + 1

    prevCode = GetSoundexCode(AscW(first))
    For i = 2 To Len(text)
        If pIdx >= 4 Then Exit For
        ch = UCase$(Mid$(text, i, 1))
        code = GetSoundexCode(AscW(ch))
        If code > 0 Then
            If code <> prevCode Then
                parts(pIdx) = CStr(code): pIdx = pIdx + 1
                prevCode = code
            End If
        Else
            ' 元音 (AEIOUY) 分隔相同编码的辅音 — 重置 prevCode
            ' 以便下次出现相同编码时能再次输出。
            ' H 和 W 在 Soundex 规范中为连接符：不重置 prevCode
            ' (例如 "Ashcraft" → prevCode 越过 H 保持为 2 → A261)。
            If InStr("AEIOUY", ch) > 0 Then prevCode = 0
        End If
    Next i

    ' 补齐到 4 位
    While pIdx < 4
        parts(pIdx) = "0": pIdx = pIdx + 1
    Wend

    Soundex = Left$(Join(parts, ""), 4)
End Function

Private Function GetSoundexCode(ByVal charCode As Long) As Long
    Select Case charCode
        Case AscW("B"), AscW("F"), AscW("P"), AscW("V"): GetSoundexCode = 1
        Case AscW("C"), AscW("G"), AscW("J"), AscW("K"), AscW("Q"), AscW("S"), AscW("X"), AscW("Z"): GetSoundexCode = 2
        Case AscW("D"), AscW("T"): GetSoundexCode = 3
        Case AscW("L"): GetSoundexCode = 4
        Case AscW("M"), AscW("N"): GetSoundexCode = 5
        Case AscW("R"): GetSoundexCode = 6
        Case Else: GetSoundexCode = 0
    End Select
End Function
'=============================================================================
' LevenshteinDistance — 编辑距离 (Levenshtein 算法)
'
' 用两个一维滚动数组避免 O(n*m) 空间开销
'=============================================================================
Public Function LevenshteinDistance( _
    ByVal a As String, _
    ByVal b As String, _
    Optional ByVal caseSensitive As Boolean = False) As Long

    Dim nA As Long, nB As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim prevRow() As Long, curRow() As Long

    If Not caseSensitive Then
        a = LCase$(a)
        b = LCase$(b)
    End If

    nA = Len(a)
    nB = Len(b)

    If nA = 0 Then
        LevenshteinDistance = nB
        Exit Function
    End If
    If nB = 0 Then
        LevenshteinDistance = nA
        Exit Function
    End If

    ReDim prevRow(0 To nB)
    ReDim curRow(0 To nB)

    For j = 0 To nB
        prevRow(j) = j
    Next j

    For i = 1 To nA
        curRow(0) = i
        For j = 1 To nB
            If mid$(a, i, 1) = mid$(b, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            curRow(j) = Min3(prevRow(j) + 1, curRow(j - 1) + 1, prevRow(j - 1) + cost)
        Next j
        ' 将 curRow 复制到 prevRow 以备下次迭代（VBA 深层复制数组）
        prevRow = curRow
    Next i

    LevenshteinDistance = prevRow(nB)
End Function

Private Function Min3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Min3 = a
    If b < Min3 Then Min3 = b
    If c < Min3 Then Min3 = c
End Function

'=============================================================================
' 工作表函数 (UDF_STR_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_STR_EXTRACTBETWEEN( _
    ByVal text As Variant, _
    ByVal leftDelim As Variant, _
    ByVal rightDelim As Variant, _
    Optional ByVal nth As Variant = 1, _
    Optional ByVal includeDelim As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = ExtractBetween(CStr(v), leftDelim, rightDelim, nth, includeDelim)
        Next: Next
        UDF_STR_EXTRACTBETWEEN = resultArr: Exit Function
    End If
    UDF_STR_EXTRACTBETWEEN = ExtractBetween(CStr(text), leftDelim, rightDelim, nth, includeDelim): Exit Function
EH: UDF_STR_EXTRACTBETWEEN = CVErr(xlErrValue)
End Function


Public Function UDF_STR_REVERSESTRING(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = ReverseString(CStr(v))
        Next: Next
        UDF_STR_REVERSESTRING = resultArr: Exit Function
    End If
    UDF_STR_REVERSESTRING = ReverseString(CStr(text)): Exit Function
EH: UDF_STR_REVERSESTRING = CVErr(xlErrValue)
End Function

Public Function UDF_STR_COUNTSUBSTRING( _
    ByVal text As Variant, _
    ByVal search As Variant, _
    Optional ByVal caseSensitive As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = CountSubstring(CStr(v), search, caseSensitive)
        Next: Next
        UDF_STR_COUNTSUBSTRING = resultArr: Exit Function
    End If
    UDF_STR_COUNTSUBSTRING = CountSubstring(CStr(text), search, caseSensitive): Exit Function
EH: UDF_STR_COUNTSUBSTRING = CVErr(xlErrValue)
End Function

Public Function UDF_STR_STARTSWITH( _
    ByVal text As Variant, _
    ByVal prefix As Variant, _
    Optional ByVal caseSensitive As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = StartsWith(CStr(v), prefix, caseSensitive)
        Next: Next
        UDF_STR_STARTSWITH = resultArr: Exit Function
    End If
    UDF_STR_STARTSWITH = StartsWith(CStr(text), prefix, caseSensitive): Exit Function
EH: UDF_STR_STARTSWITH = CVErr(xlErrValue)
End Function

Public Function UDF_STR_ENDSWITH( _
    ByVal text As Variant, _
    ByVal suffix As Variant, _
    Optional ByVal caseSensitive As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = EndsWith(CStr(v), suffix, caseSensitive)
        Next: Next
        UDF_STR_ENDSWITH = resultArr: Exit Function
    End If
    UDF_STR_ENDSWITH = EndsWith(CStr(text), suffix, caseSensitive): Exit Function
EH: UDF_STR_ENDSWITH = CVErr(xlErrValue)
End Function

Public Function UDF_STR_LEFTOF( _
    ByVal text As Variant, _
    ByVal delimiter As Variant, _
    Optional ByVal nth As Variant = 1) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = LeftOf(CStr(v), delimiter, nth)
        Next: Next
        UDF_STR_LEFTOF = resultArr: Exit Function
    End If
    UDF_STR_LEFTOF = LeftOf(CStr(text), delimiter, nth): Exit Function
EH: UDF_STR_LEFTOF = CVErr(xlErrValue)
End Function

Public Function UDF_STR_RIGHTOF( _
    ByVal text As Variant, _
    ByVal delimiter As Variant, _
    Optional ByVal nth As Variant = 1, _
    Optional ByVal fromRight As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RightOf(CStr(v), delimiter, nth, fromRight)
        Next: Next
        UDF_STR_RIGHTOF = resultArr: Exit Function
    End If
    UDF_STR_RIGHTOF = RightOf(CStr(text), delimiter, nth, fromRight): Exit Function
EH: UDF_STR_RIGHTOF = CVErr(xlErrValue)
End Function

Public Function UDF_STR_TEXTJOIN( _
    ByVal delimiter As Variant, _
    ParamArray args() As Variant) As Variant
    On Error GoTo EH: UDF_STR_TEXTJOIN = TextJoin(delimiter, args): Exit Function
EH: UDF_STR_TEXTJOIN = CVErr(xlErrValue)
End Function

Public Function UDF_STR_NTHWORD( _
    ByVal text As Variant, _
    ByVal n As Variant, _
    Optional ByVal delimiter As Variant = " ") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = NthWord(CStr(v), n, delimiter)
        Next: Next
        UDF_STR_NTHWORD = resultArr: Exit Function
    End If
    UDF_STR_NTHWORD = NthWord(CStr(text), n, delimiter): Exit Function
EH: UDF_STR_NTHWORD = CVErr(xlErrValue)
End Function

Public Function UDF_STR_COMMONPREFIX( _
    ByVal a As Variant, _
    ByVal b As Variant, _
    Optional ByVal caseSensitive As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(a) Then If TypeOf a Is Range Then a = a.Value
    If IsObject(b) Then If TypeOf b Is Range Then b = b.Value
    If IsArray(a) Then
        ReDim resultArr(LBound(a,1) To UBound(a,1), LBound(a,2) To UBound(a,2))
        For i = LBound(a,1) To UBound(a,1)
            For j = LBound(a,2) To UBound(a,2)
                v = a(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = CommonPrefix(CStr(v), CStr(b), caseSensitive)
        Next: Next
        UDF_STR_COMMONPREFIX = resultArr: Exit Function
    End If
    UDF_STR_COMMONPREFIX = CommonPrefix(CStr(a), CStr(b), caseSensitive): Exit Function
EH: UDF_STR_COMMONPREFIX = CVErr(xlErrValue)
End Function


Public Function UDF_STR_PADLEFT( _
    ByVal text As Variant, _
    ByVal totalWidth As Variant, _
    Optional ByVal padChar As Variant = " ") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = PadLeft(CStr(v), totalWidth, padChar)
        Next: Next
        UDF_STR_PADLEFT = resultArr: Exit Function
    End If
    UDF_STR_PADLEFT = PadLeft(CStr(text), totalWidth, padChar): Exit Function
EH: UDF_STR_PADLEFT = CVErr(xlErrValue)
End Function

Public Function UDF_STR_PADRIGHT( _
    ByVal text As Variant, _
    ByVal totalWidth As Variant, _
    Optional ByVal padChar As Variant = " ") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = PadRight(CStr(v), totalWidth, padChar)
        Next: Next
        UDF_STR_PADRIGHT = resultArr: Exit Function
    End If
    UDF_STR_PADRIGHT = PadRight(CStr(text), totalWidth, padChar): Exit Function
EH: UDF_STR_PADRIGHT = CVErr(xlErrValue)
End Function

Public Function UDF_STR_TRUNCATE( _
    ByVal text As Variant, _
    ByVal maxLength As Variant, _
    Optional ByVal suffix As Variant = "...") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = Truncate(CStr(v), maxLength, suffix)
        Next: Next
        UDF_STR_TRUNCATE = resultArr: Exit Function
    End If
    UDF_STR_TRUNCATE = Truncate(CStr(text), maxLength, suffix): Exit Function
EH: UDF_STR_TRUNCATE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_NORMALIZEWHITESPACE(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = NormalizeWhitespace(CStr(v))
        Next: Next
        UDF_STR_NORMALIZEWHITESPACE = resultArr: Exit Function
    End If
    UDF_STR_NORMALIZEWHITESPACE = NormalizeWhitespace(CStr(text)): Exit Function
EH: UDF_STR_NORMALIZEWHITESPACE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_REMOVECHARS( _
    ByVal text As Variant, _
    ByVal charsToRemove As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RemoveChars(CStr(v), charsToRemove)
        Next: Next
        UDF_STR_REMOVECHARS = resultArr: Exit Function
    End If
    UDF_STR_REMOVECHARS = RemoveChars(CStr(text), charsToRemove): Exit Function
EH: UDF_STR_REMOVECHARS = CVErr(xlErrValue)
End Function

Public Function UDF_STR_KEEPCHARS( _
    ByVal text As Variant, _
    ByVal allowedChars As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = KeepChars(CStr(v), allowedChars)
        Next: Next
        UDF_STR_KEEPCHARS = resultArr: Exit Function
    End If
    UDF_STR_KEEPCHARS = KeepChars(CStr(text), allowedChars): Exit Function
EH: UDF_STR_KEEPCHARS = CVErr(xlErrValue)
End Function

Public Function UDF_STR_TOTITLECASE(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = ToTitleCase(CStr(v))
        Next: Next
        UDF_STR_TOTITLECASE = resultArr: Exit Function
    End If
    UDF_STR_TOTITLECASE = ToTitleCase(CStr(text)): Exit Function
EH: UDF_STR_TOTITLECASE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_REMOVEDIACRITICS(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RemoveDiacritics(CStr(v))
        Next: Next
        UDF_STR_REMOVEDIACRITICS = resultArr: Exit Function
    End If
    UDF_STR_REMOVEDIACRITICS = RemoveDiacritics(CStr(text)): Exit Function
EH: UDF_STR_REMOVEDIACRITICS = CVErr(xlErrValue)
End Function

Public Function UDF_STR_SLUGIFY( _
    ByVal text As Variant, _
    Optional ByVal separator As Variant = "-") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = Slugify(CStr(v), separator)
        Next: Next
        UDF_STR_SLUGIFY = resultArr: Exit Function
    End If
    UDF_STR_SLUGIFY = Slugify(CStr(text), separator): Exit Function
EH: UDF_STR_SLUGIFY = CVErr(xlErrValue)
End Function

Public Function UDF_STR_ISNULLOREMPTY(ByVal text As Variant) As Variant
    On Error GoTo EH:     UDF_STR_ISNULLOREMPTY = IsNullOrEmpty(text): Exit Function
EH: UDF_STR_ISNULLOREMPTY = CVErr(xlErrValue)
End Function

Public Function UDF_STR_ISNULLORWHITESPACE(ByVal text As Variant) As Variant
    On Error GoTo EH: UDF_STR_ISNULLORWHITESPACE = IsNullOrWhitespace(text): Exit Function
EH: UDF_STR_ISNULLORWHITESPACE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_ISEMAIL(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = IsEmail(CStr(v))
        Next: Next
        UDF_STR_ISEMAIL = resultArr: Exit Function
    End If
    UDF_STR_ISEMAIL = IsEmail(CStr(text)): Exit Function
EH: UDF_STR_ISEMAIL = CVErr(xlErrValue)
End Function

Public Function UDF_STR_ISURL(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = IsUrl(CStr(v))
        Next: Next
        UDF_STR_ISURL = resultArr: Exit Function
    End If
    UDF_STR_ISURL = IsUrl(CStr(text)): Exit Function
EH: UDF_STR_ISURL = CVErr(xlErrValue)
End Function

Public Function UDF_STR_LEVENSHTEIN( _
    ByVal a As Variant, _
    ByVal b As Variant, _
    Optional ByVal caseSensitive As Variant = False) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(a) Then If TypeOf a Is Range Then a = a.Value
    If IsObject(b) Then If TypeOf b Is Range Then b = b.Value
    If IsArray(a) Then
        ReDim resultArr(LBound(a,1) To UBound(a,1), LBound(a,2) To UBound(a,2))
        For i = LBound(a,1) To UBound(a,1)
            For j = LBound(a,2) To UBound(a,2)
                v = a(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = LevenshteinDistance(CStr(v), CStr(b), caseSensitive)
        Next: Next
        UDF_STR_LEVENSHTEIN = resultArr: Exit Function
    End If
    UDF_STR_LEVENSHTEIN = LevenshteinDistance(CStr(a), CStr(b), caseSensitive): Exit Function
EH: UDF_STR_LEVENSHTEIN = CVErr(xlErrValue)
End Function

Public Function UDF_STR_SOUNDEX(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = Soundex(CStr(v))
        Next: Next
        UDF_STR_SOUNDEX = resultArr: Exit Function
    End If
    UDF_STR_SOUNDEX = Soundex(CStr(text)): Exit Function
EH: UDF_STR_SOUNDEX = CVErr(xlErrValue)
End Function

Public Function UDF_STR_URLENCODE(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = URLEncode(CStr(v))
        Next: Next
        UDF_STR_URLENCODE = resultArr: Exit Function
    End If
    UDF_STR_URLENCODE = URLEncode(CStr(text)): Exit Function
EH: UDF_STR_URLENCODE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_URLDECODE(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = URLDecode(CStr(v))
        Next: Next
        UDF_STR_URLDECODE = resultArr: Exit Function
    End If
    UDF_STR_URLDECODE = URLDecode(CStr(text)): Exit Function
EH: UDF_STR_URLDECODE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_BASE64ENCODE( _
    ByVal text As Variant, _
    Optional ByVal encoding As Variant = "UTF-8") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = Base64Encode(CStr(v), encoding)
        Next: Next
        UDF_STR_BASE64ENCODE = resultArr: Exit Function
    End If
    UDF_STR_BASE64ENCODE = Base64Encode(CStr(text), encoding): Exit Function
EH: UDF_STR_BASE64ENCODE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_BASE64DECODE( _
    ByVal base64 As Variant, _
    Optional ByVal encoding As Variant = "UTF-8") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(base64) Then If TypeOf base64 Is Range Then base64 = base64.Value
    If IsArray(base64) Then
        ReDim resultArr(LBound(base64,1) To UBound(base64,1), LBound(base64,2) To UBound(base64,2))
        For i = LBound(base64,1) To UBound(base64,1)
            For j = LBound(base64,2) To UBound(base64,2)
                v = base64(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = Base64Decode(CStr(v), encoding)
        Next: Next
        UDF_STR_BASE64DECODE = resultArr: Exit Function
    End If
    UDF_STR_BASE64DECODE = Base64Decode(CStr(base64), encoding): Exit Function
EH: UDF_STR_BASE64DECODE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_HTMLENCODE(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = HTMLEncode(CStr(v))
        Next: Next
        UDF_STR_HTMLENCODE = resultArr: Exit Function
    End If
    UDF_STR_HTMLENCODE = HTMLEncode(CStr(text)): Exit Function
EH: UDF_STR_HTMLENCODE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_HTMLDECODE(ByVal text As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = HTMLDecode(CStr(v))
        Next: Next
        UDF_STR_HTMLDECODE = resultArr: Exit Function
    End If
    UDF_STR_HTMLDECODE = HTMLDecode(CStr(text)): Exit Function
EH: UDF_STR_HTMLDECODE = CVErr(xlErrValue)
End Function

' 注意: ParamArray UDF 不显示在 Excel 函数向导中, 但可直接在单元格中键入调用
' 参数限制: 最多 255 个, 区域引用不会自动展开为逐元素
Public Function UDF_STR_COALESCE(ParamArray values() As Variant) As Variant
    On Error GoTo EH
    ' VBA cannot forward ParamArray directly; Coalesce receives the whole array as one element
    Dim i As Long, v As Variant
    For i = LBound(values) To UBound(values)
        v = values(i)
        If Not IsNull(v) And Not IsEmpty(v) And Not IsError(v) Then
            If VarType(v) <> vbString Or Len(CStr(v)) > 0 Then
                UDF_STR_COALESCE = v: Exit Function
            End If
        End If
    Next i
    UDF_STR_COALESCE = Empty
    Exit Function
EH: UDF_STR_COALESCE = CVErr(xlErrValue)
End Function

Public Function UDF_STR_UUID() As Variant
    On Error GoTo EH: UDF_STR_UUID = UUID(): Exit Function
EH: UDF_STR_UUID = CVErr(xlErrValue)
End Function

Public Function UDF_STR_RANDOMSTRING( _
    Optional ByVal length As Variant = 8, _
    Optional ByVal charset As Variant = "") As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(charset) Then If TypeOf charset Is Range Then charset = charset.Value
    If IsArray(charset) Then
        ReDim resultArr(LBound(charset,1) To UBound(charset,1), LBound(charset,2) To UBound(charset,2))
        For i = LBound(charset,1) To UBound(charset,1)
            For j = LBound(charset,2) To UBound(charset,2)
                v = charset(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RandomString(length, CStr(v))
        Next: Next
        UDF_STR_RANDOMSTRING = resultArr: Exit Function
    End If
    UDF_STR_RANDOMSTRING = RandomString(length, CStr(charset)): Exit Function
EH: UDF_STR_RANDOMSTRING = CVErr(xlErrValue)
End Function

Public Function UDF_STR_REPEAT(ByVal text As Variant, ByVal n As Variant) As Variant
    Dim i As Long, j As Long, v As Variant, resultArr() As Variant
    On Error GoTo EH
    If IsObject(text) Then If TypeOf text Is Range Then text = text.Value
    If IsArray(text) Then
        ' 1D 数组支持: 探测维度 2, 若为 1D 则提升为单行 2D 后走统一映射 (避免 Error 9 拒绝合法输入)
        Err.Clear: On Error Resume Next
        Dim d2p As Long: d2p = UBound(text, 2)
        If Err.Number <> 0 Then
            Dim tx2() As Variant: ReDim tx2(1 To 1, LBound(text) To UBound(text))
            Dim kk As Long
            For kk = LBound(text) To UBound(text): tx2(1, kk) = text(kk): Next kk
            text = tx2
        End If
        On Error GoTo EH
        ReDim resultArr(LBound(text,1) To UBound(text,1), LBound(text,2) To UBound(text,2))
        For i = LBound(text,1) To UBound(text,1)
            For j = LBound(text,2) To UBound(text,2)
                v = text(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = Repeat(CStr(v), n)
        Next: Next
        UDF_STR_REPEAT = resultArr: Exit Function
    End If
    UDF_STR_REPEAT = Repeat(CStr(text), n): Exit Function
EH: UDF_STR_REPEAT = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例
'=====================================================================
' ExtractBetween("abc[123]def", "[", "]")                    → "123"