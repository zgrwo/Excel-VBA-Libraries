Option Explicit

'==============================================================================
' Module:       DateTimeUtils
' Purpose:      Date/Time: ISO week, workdays, age, Easter, timestamps
' Layer:        DateTime
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       52 functions/subs
'==============================================================================

Private DP As New DictProxy

' -- 错误代码常量 -----------------------------------------------------
Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1001
Private Const ERR_OUT_OF_BOUNDS As Long = vbObjectError + 1002

'=====================================================================
' DateTimeUtils.bas — 日期时间工具集
'
' 工作表函数 (UDF_DT_*):
'   UDF_DT_ISOWEEKNUM      — ISO 8601 周数
'   UDF_DT_FIRSTDAYOFMONTH — 月初
'   UDF_DT_LASTDAYOFMONTH  — 月末
'   UDF_DT_FIRSTDAYOFQUARTER — 季度首日
'   UDF_DT_LASTDAYOFQUARTER  — 季度末日
'   UDF_DT_DAYSINMONTH     — 当月天数
'   UDF_DT_DAYSINYEAR      — 当年天数
'   UDF_DT_DAYOFYEAR       — 一年中的第几天
'   UDF_DT_ISLEAPYEAR      — 闰年判断
'   UDF_DT_QUARTER         — 季度 (1-4)
'   UDF_DT_FISCALYEAR      — 财年
'   UDF_DT_DATETOUNIX      — 日期 → Unix 时间戳
'   UDF_DT_UNIXTODATE      — Unix 时间戳 → 日期
'   UDF_DT_ISWEEKEND       — 判断是否为周末
'   UDF_DT_ISHOLIDAY       — 判断是否为节假日
'   UDF_DT_ADDMONTHSSAFE   — 安全添加月份 (月末处理)
'   UDF_DT_WORKDAYSBETWEEN — 工作日数
'   UDF_DT_NEXTWORKDAY     — n 个工作日后的日期
'   UDF_DT_WEEKOFMONTH     — 当月第几周
'   UDF_DT_STARTOFWEEK     — 周首日
'   UDF_DT_ENDOFWEEK       — 周末日
'   UDF_DT_DATERANGE       — 生成日期序列
'   UDF_DT_AGEYEARS         — 简单整数年龄
'   UDF_DT_NTHWEEKDAY       — 某月第 n 个指定星期几
'   UDF_DT_EASTER          — 复活节日期
'
' 内部函数 (PascalCase，供 VBA 代码调用):
'   ISOWeekNum          — ISO 8601 周数
'   FirstDayOfMonth     — 月初
'   LastDayOfMonth      — 月末
'   FirstDayOfQuarter   — 季度首日
'   LastDayOfQuarter    — 季度末日
'   DaysInMonth         — 当月天数
'   DaysInYear          — 一年中的总天数
'   DayOfYear           — 一年中的第几天
'   IsLeapYear          — 闰年判断
'   Age                 — 精确年龄 (VBA-only, 返回 Dictionary)
'   Quarter             — 季度 (1-4)
'   FiscalYear          — 财年
'   WorkdaysBetween     — 工作日数
'   NextWorkday         — n 个工作日后的日期
'   UnixToDate          — Unix 时间戳 → 日期
'   DateToUnix          — 日期 → Unix 时间戳
'   IsWeekend           — 判断是否为周末
'   IsHoliday           — 判断是否为节假日
'   AddMonthsSafe       — 安全添加月份 (月末处理)
'   DateDiffParts       — 精确日期差 (VBA-only, 返回 Dictionary)
'   WeekOfMonth         — 当月第几周
'   StartOfWeek         — 周首日 (可指定起始日)
'   EndOfWeek           — 周末日
'   DateRange           — 生成日期序列 (d/w/m/y)
'   NthWeekday          — 某月第 n 个指定星期几
'   AgeYears            — 简单整数年龄
'   Easter              — 复活节日期
'
' 注意事项:
'   - DaysInMonth 单参数调用时将参数视为日期而非年份（如 DaysInMonth(2024)
'     解读为 1905-07-17 的序列号，而非 2024 年）。请使用双参数 (year, month)。
'   - VBA-only 函数 (Age, DateDiffParts) 返回 Dictionary，仅限 VBA 调用。
'=====================================================================


'=============================================================================
' ExtractHolidays — 将节假日输入标准化为 Variant（数组或 Range.Value）
'
' 处理 Range（.Value 提取）、数组（直传）、标量（包装为单元素数组）。
' 重要: VBA And 不会短路，因此 IsObject / TypeOf 必须嵌套
' 以避免数组/标量上出现 "Object required" 错误。
'=============================================================================
Private Function ExtractHolidays(ByRef holidays As Variant) As Variant
    If IsObject(holidays) Then
        If holidays Is Nothing Then
            ExtractHolidays = Array()
        ElseIf TypeOf holidays Is Range Then
            If holidays.Cells.Count = 1 Then
                ExtractHolidays = Array(holidays.Value)
            Else
                ExtractHolidays = holidays.Value
            End If
        Else
            ExtractHolidays = holidays
        End If
    ElseIf IsArray(holidays) Then
        ExtractHolidays = holidays
    Else
        ExtractHolidays = Array(holidays)
    End If
End Function

'=============================================================================
' ISOWeekNum — ISO 8601 周数
'
' 规则: 包含当年第一个星期四的那一周为第 1 周
' 结果: 1-53
'=============================================================================
Public Function ISOWeekNum(ByVal d As Variant) As Long
    If IsObject(d) Then
        If TypeOf d Is Range Then d = d.Value
    End If
    If VarType(d) <> vbDate Then If Not IsDate(d) Then ISOWeekNum = 0: Exit Function
    d = Int(CDate(d))
    Dim y As Long
    Dim jan4 As Date
    Dim firstMonday As Date
    Dim prevJan4 As Date
    Dim prevFirstMonday As Date
    Dim nextYearFirstMonday As Date
    y = Year(d)

    jan4 = DateSerial(y, 1, 4)
    firstMonday = jan4 - Weekday(jan4, vbMonday) + 1

    If d < firstMonday Then
        If y <= 100 Then
            ' 守卫: DateSerial(99, ...) 低于 VBA 最小年份 100;
            ' 年份 100 的 ISO 周一之前的日期属于年份 99 的最后一周。
            ISOWeekNum = 52
        Else
            prevJan4 = DateSerial(y - 1, 1, 4)
            prevFirstMonday = prevJan4 - Weekday(prevJan4, vbMonday) + 1
            ISOWeekNum = CLng((DateSerial(y - 1, 12, 31) - prevFirstMonday) \ 7) + 1
        End If
    ElseIf y >= 9999 Then
        ' 守卫: DateSerial(10000, ...) 会抛出错误 5
        ' 注: 年 9999-01-01 为周四 → 该年为 53 周年份, 上限设为 53 而非 52
        ISOWeekNum = CLng((d - firstMonday) \ 7) + 1
        If ISOWeekNum > 53 Then ISOWeekNum = 53
    Else
        nextYearFirstMonday = DateSerial(y + 1, 1, 4)
        nextYearFirstMonday = nextYearFirstMonday - Weekday(nextYearFirstMonday, vbMonday) + 1
        If d >= nextYearFirstMonday Then
            ISOWeekNum = 1
        Else
            ISOWeekNum = CLng((d - firstMonday) \ 7) + 1
        End If
    End If
End Function

'=============================================================================
' FirstDayOfMonth / LastDayOfMonth
'=============================================================================
Public Function FirstDayOfMonth(Optional ByVal d As Variant) As Date
    If IsMissing(d) Then d = Date
    If IsNull(d) Then d = Date
    If VarType(d) <> vbDate Then
        If Not IsDate(d) Then d = Date
    End If
    d = CDate(d)
    FirstDayOfMonth = DateSerial(Year(d), Month(d), 1)
End Function

Public Function LastDayOfMonth(Optional ByVal d As Variant) As Date
    If IsMissing(d) Then d = Date
    If IsNull(d) Then d = Date
    If VarType(d) <> vbDate Then
        If Not IsDate(d) Then d = Date
    End If
    d = CDate(d)
    If Year(d) >= 9999 And Month(d) = 12 Then
        LastDayOfMonth = DateSerial(9999, 12, 31)
    Else
        LastDayOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
    End If
End Function

'=============================================================================
' FirstDayOfQuarter / LastDayOfQuarter
'=============================================================================
Public Function FirstDayOfQuarter(Optional ByVal d As Variant) As Date
    If IsMissing(d) Then d = Date
    If IsNull(d) Then d = Date
    If VarType(d) <> vbDate Then
        If Not IsDate(d) Then d = Date
    End If
    d = CDate(d)
    Dim q As Long: q = Quarter(d)
    FirstDayOfQuarter = DateSerial(Year(d), (q - 1) * 3 + 1, 1)
End Function

Public Function LastDayOfQuarter(Optional ByVal d As Variant) As Date
    If IsMissing(d) Then d = Date
    If IsNull(d) Then d = Date
    If VarType(d) <> vbDate Then
        If Not IsDate(d) Then d = Date
    End If
    d = CDate(d)
    Dim q As Long: q = Quarter(d)
    If Year(d) >= 9999 And q = 4 Then
        LastDayOfQuarter = DateSerial(9999, 12, 31)
    Else
        LastDayOfQuarter = DateSerial(Year(d), q * 3 + 1, 0)
    End If
End Function

'=============================================================================
' DaysInMonth / DaysInYear / DayOfYear
'=============================================================================
Public Function DaysInMonth(Optional ByVal y As Variant, Optional ByVal m As Variant) As Long
    If IsMissing(y) Then y = Year(Date)
    If IsMissing(m) Then
        ' 单参数: 如果看起来像日期，同时提取年份和月份。
        ' IsDate 同时捕获 vbDate 和工作表 Double 序列号。
        ' 注意: 纯数字（如 2024）在 VBA 中是日期序列号;
        ' 调用者应传递两个参数 (year, month) 来表示纯数值年份。
        If VarType(y) = vbDate Or IsDate(y) Then
            m = Month(CDate(y))
            y = Year(CDate(y))
        Else
            m = Month(Date)
            If Not IsNumeric(y) Then
                Err.Raise ERR_OUT_OF_BOUNDS, "DateTimeUtils", _
                    "DaysInMonth: 单参数必须是日期或数值年份，但传入了非数值型值。"
            End If
        End If
    Else
        ' Two arguments: treat as year and month numbers
        If IsNull(y) Or Not IsNumeric(y) Then y = Year(Date)
        If IsNull(m) Or Not IsNumeric(m) Then m = Month(Date)
    End If

    Dim yy As Long: yy = CLng(y)
    Dim mm As Long: mm = CLng(m)

    If yy < 100 Or yy > 9999 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "DateTimeUtils", _
            "DaysInMonth: year " & CStr(yy) & " is out of range (100..9999)"
    End If
    If mm < 1 Or mm > 12 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "DateTimeUtils", _
            "DaysInMonth: month " & CStr(mm) & " is out of range (1..12)"
    End If
    DaysInMonth = Day(DateSerial(yy, mm + 1, 0))
End Function

Public Function DaysInYear(Optional ByVal y As Variant) As Long
    If IsMissing(y) Then y = Year(Date)
    If IsNull(y) Or IsEmpty(y) Then y = Year(Date)
    If Not IsNumeric(y) Then y = Year(Date)
    Dim yy As Long: yy = CLng(Val(CStr(y)))
    If yy < 100 Or yy > 9999 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "DateTimeUtils", _
            "DaysInYear: year " & CStr(yy) & " is out of range (100..9999)"
    End If
    If IsLeapYear(yy) Then DaysInYear = 366 Else DaysInYear = 365
End Function

Public Function DayOfYear(ByVal d As Variant) As Long
    If VarType(d) <> vbDate Then If Not IsDate(d) Then DayOfYear = 0: Exit Function
    d = CDate(d)
    DayOfYear = DateDiff("d", DateSerial(Year(d), 1, 0), d)
End Function

'=============================================================================
' IsLeapYear — 闰年判断
'=============================================================================
Public Function IsLeapYear(ByVal y As Variant) As Boolean
    ' Range→Value conversion
    If IsObject(y) Then
        If TypeName(y) = "Range" Then y = y.Value
    End If
    ' Extract scalar from single-cell Range result
    If IsArray(y) Then
        On Error Resume Next
        y = y(LBound(y, 1), LBound(y, 2))
        On Error GoTo 0
    End If

    If IsNull(y) Or IsEmpty(y) Then
        IsLeapYear = False
        Exit Function
    End If
    ' Accept both numeric years and date values (Date, Date string, date serial)
    Dim yy As Long
    If IsDate(y) Then
        yy = Year(CDate(y))
    ElseIf IsNumeric(y) Then
        ' Val() handles strings like "1,234" that IsNumeric accepts but CLng rejects
        yy = CLng(Val(CStr(y)))
    Else
        IsLeapYear = False
        Exit Function
    End If
    IsLeapYear = (yy Mod 4 = 0 And yy Mod 100 <> 0) Or (yy Mod 400 = 0)
End Function

'=============================================================================
' Quarter — 返回季度 1-4
'=============================================================================
Public Function Quarter(Optional ByVal d As Variant) As Long
    If IsMissing(d) Then d = Date
    If IsNull(d) Then d = Date
    If VarType(d) <> vbDate Then
        If Not IsDate(d) Then d = Date
    End If
    Quarter = (Month(CDate(d)) - 1) \ 3 + 1
End Function

'=============================================================================
' Age — 精确年龄
'
' 返回: 包含 "years", "months", "days", "totalYears", "totalMonths", "totalDays" 的字典
'=============================================================================
Public Function Age( _
    ByVal birthDate As Date, _
    Optional ByVal asOf As Variant) As Object

    Dim result As Object
    Dim refDate As Date
    Dim y As Long, m As Long, d As Long
    Dim lastBirthday As Date
    Dim monthAnniv As Date

    If IsMissing(asOf) Then
        refDate = Date
    ElseIf VarType(asOf) = vbDate Then
        refDate = CDate(asOf)
    Else
        refDate = Date
    End If

    Set result = DP.Create()

    If birthDate > refDate Then
        result("years") = 0: result("months") = 0: result("days") = 0
        result("totalYears") = 0#: result("totalMonths") = 0: result("totalDays") = 0
        Set Age = result
        Exit Function
    End If

    ' 年月日分量
    y = Year(refDate) - Year(birthDate)
    If Month(refDate) < Month(birthDate) Or _
       (Month(refDate) = Month(birthDate) And Day(refDate) < Day(birthDate)) Then
        y = y - 1
    End If

    Dim bY As Long, bM As Long, bD As Long, capD As Long
    bY = Year(birthDate): bM = Month(birthDate): bD = Day(birthDate)
    ' 手动计算周年避免 VBA DateAdd("yyyy") 闰年 bug（2月29日）
    capD = DaysInMonth(bY + y, bM)
    If bD > capD Then bD = capD
    lastBirthday = DateSerial(bY + y, bM, bD)

    ' 计算月份: 向 lastBirthday 添加月份，如果超出则回退
    m = DateDiff("m", lastBirthday, refDate)
    monthAnniv = DateAdd("m", m, lastBirthday)
    If monthAnniv > refDate Then
        m = m - 1
        monthAnniv = DateAdd("m", m, lastBirthday)
    End If
    d = DateDiff("d", monthAnniv, refDate)

    ' Normalize: roll excess months into years
    y = y + m \ 12
    m = m Mod 12

    result("years") = y
    result("months") = m
    result("days") = d
    ' 公历平均年（365.2425 天）— 近似值; 个体寿命可能偏差约 0.001 年
    result("totalYears") = CDbl(DateDiff("d", birthDate, refDate)) / 365.2425#
    result("totalMonths") = DateDiff("m", birthDate, refDate)
    result("totalDays") = DateDiff("d", birthDate, refDate)

    Set Age = result
End Function

'=============================================================================
' AgeYears — 简单整数年龄 (比 Age() 更轻量，适合高频场景)
'
' 参数: birthDate, Optional asOf (默认今天)
'=============================================================================
Public Function AgeYears(ByVal birthDate As Date, Optional ByVal asOf As Variant) As Long
    Dim refDate As Date
    If IsMissing(asOf) Then refDate = Date Else refDate = CDate(asOf)
    AgeYears = Year(refDate) - Year(birthDate)
    If Month(refDate) < Month(birthDate) Or _
       (Month(refDate) = Month(birthDate) And Day(refDate) < Day(birthDate)) Then
        AgeYears = AgeYears - 1
    End If
End Function

'=============================================================================
' WorkdaysBetween — 两个日期之间的工作日数（纯 VBA 实现，不依赖 Excel）
'
' 按 SKILL.md 约定，错误由 Err.Raise 向上传播（VBA-to-VBA 函数），而非吞没。
' outOk 参数保留用于向后兼容，函数成功返回时始终设为 True。
'=============================================================================
Public Function WorkdaysBetween( _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    Optional ByRef holidays As Variant, _
    Optional ByRef outOk As Boolean) As Long
    ' Note: outOk (non-Variant Optional) does not support IsMissing

    Dim d As Date, count As Long, hasH As Boolean
    Dim swapTmp As Date
    Dim hols As Variant

    hasH = Not IsMissing(holidays)
    If hasH Then hols = ExtractHolidays(holidays)

    If startDate > endDate Then
        swapTmp = startDate
        startDate = endDate
        endDate = swapTmp
    End If

    count = 0
    d = Int(startDate)
    Do While d <= Int(endDate)
        If Not IsWeekend(d) Then
            If Not hasH Then
                count = count + 1
            ElseIf Not IsHoliday(d, hols) Then
                count = count + 1
            End If
        End If
        d = d + 1
    Loop

    WorkdaysBetween = count
    outOk = True
End Function

'=============================================================================
' NextWorkday — n 个工作日后的日期（纯 VBA 实现，不依赖 Excel）
'
' 按 SKILL.md 约定，错误由 Err.Raise 向上传播（VBA-to-VBA 函数），而非吞没。
' outOk 参数保留用于向后兼容，函数成功返回时始终设为 True。
'=============================================================================
Public Function NextWorkday( _
    ByVal startDate As Date, _
    ByVal n As Long, _
    Optional ByRef holidays As Variant, _
    Optional ByRef outOk As Boolean) As Date
    ' Note: outOk (non-Variant Optional) does not support IsMissing

    Dim d As Date
    Dim direction As Long
    Dim remaining As Long
    Dim hasH As Boolean
    Dim hols As Variant

    hasH = Not IsMissing(holidays)
    If hasH Then hols = ExtractHolidays(holidays)

    direction = 1
    If n < 0 Then direction = -1
    remaining = Abs(n)
    d = Int(startDate)

    Do While remaining > 0
        d = d + direction
        If Not IsWeekend(d) Then
            If hasH Then
                If Not IsHoliday(d, hols) Then
                    remaining = remaining - 1
                End If
            Else
                remaining = remaining - 1
            End If
        End If
    Loop

    NextWorkday = d
    outOk = True
End Function

'=============================================================================
' DateRange — 生成日期序列 (d=日, w=周, m=月, y=年)
'=============================================================================
Public Function DateRange(ByVal start As Date, ByVal endDate As Date, _
    Optional ByVal interval As String = "d") As Variant
    Dim result() As Date, cnt As Long, d As Date, maxCnt As Long
    If start > endDate Then
        ReDim result(0 To 0): result(0) = start
        DateRange = result: Exit Function
    End If

    ' 预计算最大可能元素数，避免循环内 ReDim Preserve (O(n²) → O(n))
    Select Case LCase$(interval)
        Case "d":  maxCnt = DateDiff("d", start, endDate) + 2
        Case "w":  maxCnt = DateDiff("ww", start, endDate) + 2
        Case "m":  maxCnt = DateDiff("m", start, endDate) + 2
        Case "y":  maxCnt = DateDiff("yyyy", start, endDate) + 2
        Case Else: Err.Raise ERR_INVALID_INPUT, "DateRange", "不支持的间隔类型: " & interval
    End Select
    ReDim result(0 To maxCnt - 1)

    cnt = 0: d = start: result(cnt) = d: cnt = cnt + 1
    Do While d < endDate
        Select Case LCase$(interval)
            Case "d": d = d + 1
            Case "w": d = d + 7
            Case "m": d = DateAdd("m", 1, d)
            Case "y": d = DateAdd("yyyy", 1, d)
        End Select
        If d > endDate Then Exit Do
        result(cnt) = d: cnt = cnt + 1
    Loop
    If cnt < maxCnt Then ReDim Preserve result(0 To cnt - 1)
    DateRange = result
End Function

'=============================================================================
' NthWeekday — 返回某月第 n 个指定星期几的日期
'
' 例: NthWeekday(2026, 6, 2, vbMonday) → 2026年6月第2个周一
'=============================================================================
Public Function NthWeekday(ByVal y As Long, ByVal m As Long, _
    ByVal n As Long, Optional ByVal dayOfWeek As Long = vbMonday) As Variant
    Dim first As Date, firstWD As Long, diff As Long
    first = DateSerial(y, m, 1)
    firstWD = Weekday(first, dayOfWeek)
    If firstWD <= 1 Then diff = 1 - firstWD Else diff = 8 - firstWD
    NthWeekday = DateAdd("d", diff, first) + (n - 1) * 7
    If Month(NthWeekday) <> m Then
        Err.Raise ERR_OUT_OF_BOUNDS, "NthWeekday", _
            "该月不存在第 " & n & " 个指定星期几。"
    End If
End Function

'=============================================================================
' FiscalYear — 财年
'
' 参数:
'   d          - 日期
'   startMonth - 财年开始月份 (默认 1 月)
'
' 返回: 财年年份 (财年开始时的日历年)
'   例: 若 7 月开始, 2024/8 → 2024 财年, 2025/3 → 2024 财年
'=============================================================================
Public Function FiscalYear( _
    Optional ByVal d As Variant, _
    Optional ByVal startMonth As Long = 1) As Long

    If IsMissing(d) Then d = Date
    If IsNull(d) Then d = Date
    If VarType(d) <> vbDate Then
        If Not IsDate(d) Then d = Date
    End If
    d = CDate(d)

    If startMonth < 1 Or startMonth > 12 Then startMonth = 1

    If Month(d) >= startMonth Then
        FiscalYear = Year(d)
    Else
        FiscalYear = Year(d) - 1
    End If
End Function

'=============================================================================
' UnixToDate / DateToUnix — Unix 时间戳与日期的互转
'
' UnixToDate 始终按 Unix 时间戳（1970-01-01 以来的秒数）处理。如需处理
' Excel 日期序列号，直接使用 CDate() 即可。此前基于 3M 阈值的启发式判断
' 会导致 1970-01-01～1970-02-04 之间的 Unix 时间戳被误判为 Excel 序列号，
' 造成 DateToUnix → UnixToDate 往返失败。
'=============================================================================
Public Function UnixToDate(ByVal ts As Variant) As Date
    If IsNull(ts) Or Not IsNumeric(ts) Then
        UnixToDate = CDate(0)
        Exit Function
    End If
    ' Unix 纪元: 1970-01-01 00:00:00 UTC
    UnixToDate = DateAdd("s", CDbl(ts), #1/1/1970#)
End Function

Public Function DateToUnix(ByVal d As Variant) As Double
    If IsNull(d) Or Not IsDate(d) Then
        DateToUnix = 0#
        Exit Function
    End If
    ' 使用 DateValue 去除时间分量（仅保留日期，丢弃时分秒信息）
    ' 设计限制: 时间分量丢失；若需保留时间精度，请使用 UnixTimestamp 函数
    ' 使用 Double 运算避免 DateDiff("s") Long 溢出（2038-01-19 限制）
    DateToUnix = CDbl(CDbl(DateValue(CDate(d))) - CDbl(DateValue(#1/1/1970#))) * 86400#
End Function

'=============================================================================
' IsWeekend — 判断是否为周末 (周六或周日)
'=============================================================================
Public Function IsWeekend(ByVal d As Variant) As Boolean
    If VarType(d) <> vbDate Then If Not IsDate(d) Then Exit Function
    d = CDate(d)
    Dim wd As Long
    wd = Weekday(d, vbSunday)
    IsWeekend = (wd = vbSaturday Or wd = vbSunday)
End Function

'=============================================================================
' IsHoliday — 判断日期是否在节假日列表中
'
' holidays 可以是:
'   - Range (单列或单行)
'   - Variant 数组
'   - Dictionary (key = 日期, value 任意)
'=============================================================================
Public Function IsHoliday(ByVal d As Date, ByVal holidays As Variant) As Boolean
    Dim i As Long
    Dim holidayData As Variant
    Dim r As Long, c As Long
    Dim is2D As Boolean
    Dim dummy2D As Long
    Dim r2 As Long, c2 As Long

    d = Int(d)

    If IsObject(holidays) Then
        If holidays Is Nothing Then
            IsHoliday = False
            Exit Function
        ElseIf TypeOf holidays Is Range Then
            holidayData = holidays.Value
            If Not IsArray(holidayData) Then
                If IsDate(holidayData) Then
                    If Int(CDate(holidayData)) = d Then
                        IsHoliday = True
                        Exit Function
                    End If
                End If
                Exit Function
            End If
            For r = LBound(holidayData, 1) To UBound(holidayData, 1)
                For c = LBound(holidayData, 2) To UBound(holidayData, 2)
                    If IsDate(holidayData(r, c)) Then
                        If Int(CDate(holidayData(r, c))) = d Then
                            IsHoliday = True
                            Exit Function
                        End If
                    End If
                Next
            Next
        ElseIf TypeName(holidays) = "Dictionary" Then
            If holidays.Exists(d) Then
                IsHoliday = True
                Exit Function
            End If
        End If
    ElseIf IsArray(holidays) Then
        Err.Clear
        On Error Resume Next
        i = UBound(holidays)
        If Err.Number = 0 Then
            On Error GoTo 0
            Err.Clear
            On Error Resume Next
            dummy2D = UBound(holidays, 2)
            is2D = (Err.Number = 0)
            On Error GoTo 0
            If is2D Then
                For r2 = LBound(holidays, 1) To UBound(holidays, 1)
                    For c2 = LBound(holidays, 2) To UBound(holidays, 2)
                        If IsDate(holidays(r2, c2)) Then
                            If Int(CDate(holidays(r2, c2))) = d Then
                                IsHoliday = True
                                Exit Function
                            End If
                        End If
                    Next
                Next
            Else
                For i = LBound(holidays) To UBound(holidays)
                    If IsDate(holidays(i)) Then
                        If Int(CDate(holidays(i))) = d Then
                            IsHoliday = True
                            Exit Function
                        End If
                    End If
                Next
            End If
        Else
            On Error GoTo 0
        End If
    Else
        Err.Raise ERR_INVALID_INPUT, "DateTimeUtils", _
            "Unsupported holidays type: " & TypeName(holidays) & ". Expected Range, Dictionary, or Array."
    End If
End Function

'=============================================================================
' AddMonthsSafe — 安全添加月份 (处理月末溢出)
'
' 示例: 2025-01-31 + 1个月 → 2025-02-28 (而非 2025-03-03)
'=============================================================================
Public Function AddMonthsSafe(ByVal d As Date, ByVal n As Long) As Date
    Dim y As Long, m As Long, dayOfMonth As Long
    Dim maxDay As Long

    y = Year(d)
    m = Month(d)
    dayOfMonth = Day(d)

    m = m + n
    Do While m > 12
        m = m - 12
        y = y + 1
    Loop
    Do While m < 1
        m = m + 12
        y = y - 1
    Loop

    If y < 100 Then y = 100
    If y > 9999 Then y = 9999

    maxDay = DaysInMonth(y, m)
    If dayOfMonth > maxDay Then
        AddMonthsSafe = DateSerial(y, m, maxDay)
    Else
        AddMonthsSafe = DateSerial(y, m, dayOfMonth)
    End If
End Function

'=============================================================================
' DateDiffParts — 精确日期差，返回年/月/日分量
'
' 返回 Dictionary: {"Years":n, "Months":n, "Days":n}
'=============================================================================
Public Function DateDiffParts(ByVal startDate As Date, ByVal endDate As Date) As Object
    Dim result As Object
    Dim y1 As Long, m1 As Long, d1 As Long
    Dim y2 As Long, m2 As Long, d2 As Long
    Dim years As Long, months As Long, days As Long
    Dim tmp As Date
    Dim prevMonth As Long, prevYear As Long
    Dim capDay As Long

    Set result = DP.Create()

    If startDate > endDate Then
        tmp = startDate
        startDate = endDate
        endDate = tmp
    End If

    y1 = Year(startDate): m1 = Month(startDate): d1 = Day(startDate)
    y2 = Year(endDate):   m2 = Month(endDate):   d2 = Day(endDate)

    years = y2 - y1
    If m2 < m1 Or (m2 = m1 And d2 < d1) Then
        years = years - 1
    End If

    If d2 >= d1 Then
        months = m2 - m1
        If months < 0 Then months = months + 12
        days = d2 - d1
    Else
        months = m2 - m1 - 1
        If months < 0 Then months = months + 12
        prevMonth = m2 - 1
        prevYear = y2
        If prevMonth < 1 Then
            prevMonth = 12
            prevYear = prevYear - 1
        End If
        capDay = Day(DateSerial(prevYear, prevMonth + 1, 0))
        If d1 > capDay Then
            days = d2
        Else
            days = capDay - d1 + d2
        End If
    End If

    result.Add "years", years
    result.Add "months", months
    result.Add "days", days
    result.Add "totalDays", DateDiff("d", startDate, endDate)
    result.Add "totalMonths", DateDiff("m", startDate, endDate)
    Set DateDiffParts = result
End Function

'=============================================================================
' WeekOfMonth — 当月第几周 (1-5)
'
' startDay: 周起始日 (vbSunday=1, vbMonday=2, ..., 默认 vbMonday)
'=============================================================================
Public Function WeekOfMonth(ByVal d As Date, Optional ByVal startDay As Long = vbMonday) As Long
    Dim firstOfMonth As Date
    Dim firstWeekday As Long
    Dim dayOfMonth As Long

    d = Int(d)
    firstOfMonth = DateSerial(Year(d), Month(d), 1)
    firstWeekday = Weekday(firstOfMonth, startDay)
    dayOfMonth = Day(d)

    WeekOfMonth = ((dayOfMonth + firstWeekday - 2) \ 7) + 1
End Function

'=============================================================================
' StartOfWeek — 周首日
'
' startDay: 周起始日 (vbSunday=1, vbMonday=2, ..., 默认 vbMonday)
'=============================================================================
Public Function StartOfWeek(ByVal d As Date, Optional ByVal startDay As Long = vbMonday) As Date
    StartOfWeek = Int(d) - Weekday(d, startDay) + 1
End Function

'=============================================================================
' EndOfWeek — 周末日
'
' startDay: 周起始日 (vbSunday=1, vbMonday=2, ..., 默认 vbMonday)
'=============================================================================
Public Function EndOfWeek(ByVal d As Date, Optional ByVal startDay As Long = vbMonday) As Date
    EndOfWeek = Int(d) - Weekday(d, startDay) + 7
End Function

'=============================================================================
' Easter — 计算复活节日期 (Gauss 算法)
'
' 返回给定年份的复活节星期日 (公历)
' 适用范围: 1900-2099
'=============================================================================
Public Function Easter(ByVal y As Long) As Date
    If y < 1900 Or y > 2099 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "DateTimeUtils", "Easter: Year must be between 1900 and 2099"
    End If
    Dim a As Long, b As Long, c As Long
    Dim d As Long, e As Long
    Dim easterMonth As Long, easterDay As Long

    a = y Mod 19
    b = y Mod 4
    c = y Mod 7
    d = (19 * a + 24) Mod 30
    e = (2 * b + 4 * c + 6 * d + 5) Mod 7

    easterDay = 22 + d + e
    easterMonth = 3

    If easterDay > 31 Then
        easterDay = easterDay - 31
        easterMonth = 4
    End If

    If easterDay = 26 And easterMonth = 4 Then easterDay = 19
    If easterDay = 25 And easterMonth = 4 And d = 28 And e = 6 And a > 10 Then easterDay = 18

    Easter = DateSerial(y, easterMonth, easterDay)
End Function

'=============================================================================

'=============================================================================
' 工作表函数 (UDF_DT_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_DT_ISOWEEKNUM(ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_ISOWEEKNUM = ISOWeekNum(CDate(d))
    Exit Function
ErrHandler:
    UDF_DT_ISOWEEKNUM = CVErr(xlErrValue)
End Function

Public Function UDF_DT_FIRSTDAYOFMONTH(Optional ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_FIRSTDAYOFMONTH = FirstDayOfMonth(d)
    Exit Function
ErrHandler:
    UDF_DT_FIRSTDAYOFMONTH = CVErr(xlErrValue)
End Function

Public Function UDF_DT_LASTDAYOFMONTH(Optional ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_LASTDAYOFMONTH = LastDayOfMonth(d)
    Exit Function
ErrHandler:
    UDF_DT_LASTDAYOFMONTH = CVErr(xlErrValue)
End Function

Public Function UDF_DT_FIRSTDAYOFQUARTER(Optional ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_FIRSTDAYOFQUARTER = FirstDayOfQuarter(d)
    Exit Function
ErrHandler:
    UDF_DT_FIRSTDAYOFQUARTER = CVErr(xlErrValue)
End Function

Public Function UDF_DT_LASTDAYOFQUARTER(Optional ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_LASTDAYOFQUARTER = LastDayOfQuarter(d)
    Exit Function
ErrHandler:
    UDF_DT_LASTDAYOFQUARTER = CVErr(xlErrValue)
End Function

Public Function UDF_DT_DAYSINMONTH(Optional ByVal y As Variant, Optional ByVal m As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_DAYSINMONTH = DaysInMonth(y, m)
    Exit Function
ErrHandler:
    UDF_DT_DAYSINMONTH = CVErr(xlErrValue)
End Function

Public Function UDF_DT_DAYSINYEAR(Optional ByVal y As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_DAYSINYEAR = DaysInYear(y)
    Exit Function
ErrHandler:
    UDF_DT_DAYSINYEAR = CVErr(xlErrValue)
End Function

Public Function UDF_DT_DAYOFYEAR(ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_DAYOFYEAR = DayOfYear(CDate(d))
    Exit Function
ErrHandler:
    UDF_DT_DAYOFYEAR = CVErr(xlErrValue)
End Function

Public Function UDF_DT_ISLEAPYEAR(ByVal y As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_ISLEAPYEAR = IsLeapYear(y)
    Exit Function
ErrHandler:
    UDF_DT_ISLEAPYEAR = CVErr(xlErrValue)
End Function

Public Function UDF_DT_QUARTER(Optional ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_QUARTER = Quarter(d)
    Exit Function
ErrHandler:
    UDF_DT_QUARTER = CVErr(xlErrValue)
End Function

Public Function UDF_DT_FISCALYEAR(Optional ByVal d As Variant, Optional ByVal startMonth As Variant = 1) As Variant
    On Error GoTo ErrHandler
    UDF_DT_FISCALYEAR = FiscalYear(d, startMonth)
    Exit Function
ErrHandler:
    UDF_DT_FISCALYEAR = CVErr(xlErrValue)
End Function

Public Function UDF_DT_DATETOUNIX(ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_DATETOUNIX = DateToUnix(d)
    Exit Function
ErrHandler:
    UDF_DT_DATETOUNIX = CVErr(xlErrValue)
End Function

Public Function UDF_DT_UNIXTODATE(ByVal ts As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_UNIXTODATE = UnixToDate(ts)
    Exit Function
ErrHandler:
    UDF_DT_UNIXTODATE = CVErr(xlErrValue)
End Function

Public Function UDF_DT_ISWEEKEND(ByVal d As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_ISWEEKEND = IsWeekend(CDate(d))
    Exit Function
ErrHandler:
    UDF_DT_ISWEEKEND = CVErr(xlErrValue)
End Function

Public Function UDF_DT_ISHOLIDAY(ByVal d As Variant, ByVal holidays As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_ISHOLIDAY = IsHoliday(CDate(d), holidays)
    Exit Function
ErrHandler:
    UDF_DT_ISHOLIDAY = CVErr(xlErrValue)
End Function

Public Function UDF_DT_ADDMONTHSSAFE(ByVal d As Variant, ByVal n As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_DT_ADDMONTHSSAFE = AddMonthsSafe(CDate(d), n)
    Exit Function
ErrHandler:
    UDF_DT_ADDMONTHSSAFE = CVErr(xlErrValue)
End Function


Public Function UDF_DT_DATERANGE(ByVal start As Variant, ByVal endDate As Variant, Optional ByVal interval As Variant = "d") As Variant
    On Error GoTo ErrHandler
    UDF_DT_DATERANGE = DateRange(CDate(start), CDate(endDate), interval)
    Exit Function
ErrHandler:
    UDF_DT_DATERANGE = CVErr(xlErrValue)
End Function
Public Function UDF_DT_WORKDAYSBETWEEN( _
    ByVal startDate As Variant, _
    ByVal endDate As Variant, _
    Optional ByVal holidays As Variant) As Variant
    On Error GoTo ErrHandler
    Dim ok As Boolean
    UDF_DT_WORKDAYSBETWEEN = WorkdaysBetween(CDate(startDate), CDate(endDate), holidays, ok)
    Exit Function
ErrHandler:
    UDF_DT_WORKDAYSBETWEEN = CVErr(xlErrValue)
End Function

Public Function UDF_DT_NEXTWORKDAY( _
    ByVal startDate As Variant, _
    ByVal n As Variant, _
    Optional ByVal holidays As Variant) As Variant
    On Error GoTo ErrHandler
    Dim ok As Boolean
    UDF_DT_NEXTWORKDAY = NextWorkday(CDate(startDate), n, holidays, ok)
    Exit Function
ErrHandler:
    UDF_DT_NEXTWORKDAY = CVErr(xlErrValue)
End Function

Public Function UDF_DT_WEEKOFMONTH(ByVal d As Variant, Optional ByVal startDay As Variant) As Variant
    On Error GoTo ErrHandler
    If IsMissing(startDay) Then
        UDF_DT_WEEKOFMONTH = WeekOfMonth(CDate(d))
    Else
        UDF_DT_WEEKOFMONTH = WeekOfMonth(CDate(d), CLng(startDay))
    End If
    Exit Function
ErrHandler:
    UDF_DT_WEEKOFMONTH = CVErr(xlErrValue)
End Function

Public Function UDF_DT_STARTOFWEEK(ByVal d As Variant, Optional ByVal startDay As Variant) As Variant
    On Error GoTo ErrHandler
    If IsMissing(startDay) Then
        UDF_DT_STARTOFWEEK = StartOfWeek(CDate(d))
    Else
        UDF_DT_STARTOFWEEK = StartOfWeek(CDate(d), CLng(startDay))
    End If
    Exit Function
ErrHandler:
    UDF_DT_STARTOFWEEK = CVErr(xlErrValue)
End Function

Public Function UDF_DT_ENDOFWEEK(ByVal d As Variant, Optional ByVal startDay As Variant) As Variant
    On Error GoTo ErrHandler
    If IsMissing(startDay) Then
        UDF_DT_ENDOFWEEK = EndOfWeek(CDate(d))
    Else
        UDF_DT_ENDOFWEEK = EndOfWeek(CDate(d), CLng(startDay))
    End If
    Exit Function
ErrHandler:
    UDF_DT_ENDOFWEEK = CVErr(xlErrValue)
End Function

Public Function UDF_DT_EASTER(ByVal y As Variant) As Variant
    On Error GoTo ErrHandler
    If IsObject(y) Then If TypeName(y) = "Range" Then y = y.Value
    UDF_DT_EASTER = Easter(CLng(y))
    Exit Function
ErrHandler:
    UDF_DT_EASTER = CVErr(xlErrValue)
End Function

Public Function UDF_DT_AGEYEARS(ByVal birthDate As Variant, Optional ByVal asOf As Variant) As Variant
    On Error GoTo ErrHandler
    If IsMissing(asOf) Then
        UDF_DT_AGEYEARS = AgeYears(CDate(birthDate))
    Else
        UDF_DT_AGEYEARS = AgeYears(CDate(birthDate), CDate(asOf))
    End If
    Exit Function
ErrHandler:
    UDF_DT_AGEYEARS = CVErr(xlErrValue)
End Function

Public Function UDF_DT_NTHWEEKDAY(ByVal y As Variant, ByVal m As Variant, ByVal n As Variant, Optional ByVal dayOfWeek As Variant) As Variant
    On Error GoTo ErrHandler
    If IsObject(y) Then If TypeName(y) = "Range" Then y = y.Value
    If IsObject(m) Then If TypeName(m) = "Range" Then m = m.Value
    If IsObject(n) Then If TypeName(n) = "Range" Then n = n.Value
    If IsObject(dayOfWeek) Then If TypeName(dayOfWeek) = "Range" Then dayOfWeek = dayOfWeek.Value
    Dim dow As Long: If IsMissing(dayOfWeek) Then dow = vbMonday Else dow = CLng(dayOfWeek)
    UDF_DT_NTHWEEKDAY = NthWeekday(CLng(y), CLng(m), CLng(n), dow)
    Exit Function
ErrHandler:
    UDF_DT_NTHWEEKDAY = CVErr(xlErrValue)
End Function


'=====================================================================
' 使用示例（期望返回值写在注释中）
'=====================================================================
' ISOWeekNum(#2025-01-01#)                              → 1
' FirstDayOfMonth(#2025-03-15#)                         → 2025-03-01
' LastDayOfMonth(#2025-02-15#)                          → 2025-02-28
' DaysInMonth(2024, 2)                                  → 29
' IsLeapYear(2024)                                      → True
' Quarter(#2025-08-15#)                                 → 3
' DateToUnix(#2025-01-01#)                              → 1735689600
' UnixToDate(1735689600#)                               → 2025-01-01
' IsWeekend(#2025-05-18#)                               → True (周日)
' Set diff = DateDiffParts(#2025-01-15#, #2025-03-20#)  → Years=0, Months=2, Days=5
' AddMonthsSafe(#2025-01-31#, 1)                        → 2025-02-28
' WeekOfMonth(#2025-03-15#, vbMonday)                   → 3
' StartOfWeek(#2025-05-19#, vbMonday)                   → 2025-05-19
' EndOfWeek(#2025-05-19#, vbMonday)                     → 2025-05-25
' Easter(2025)                                          → 2025-04-20
