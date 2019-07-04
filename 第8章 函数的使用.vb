' 范例119 使用时间和日期函数

' 119-1 计算程序运行时间
' timer函数
Option Explicit
Sub MyTime()
    Dim i As Integer
    Dim StartTime As Single
    Dim EndTime As Single
    StartTime = Timer
    For i = 1 To 10000
        Cells(1, 1) = i
    Next
    EndTime = Timer - StartTime
    MsgBox "程序运行时间：" & Format(EndTime, "0.00") & "秒"
End Sub

' 119-2 获得当月的最后一天
Option Explicit
Sub Endday()
    Dim Endday As Byte
    '范例中的day参数设置为0，则被解释成month参数指定月的前一天，
    '即表达式Month（Date）+1指定的下一个月的前一天，也就是本月的最后1天。
    Endday = Day(DateSerial(Year(Date), Month(Date) + 1, 0))
    ' 这么好的函数竟然只能在VBA中用
    MsgBox "当月最后一天是" & Month(Date) & "月" & Endday & "号"
End Sub

' 119-3 计算某个日期为星期几
' 直接用excel函数就欧克了
Option Explicit
Sub Myweekday()
    Dim StrDate As String
    Dim Myweekday As String
    StrDate = InputBox("请输入日期：")
    If Len(StrDate) = 0 Then Exit Sub
    If IsDate(StrDate) Then
        Select Case Weekday(StrDate, vbSunday)
            Case vbSunday
                Myweekday = "星期日"
            Case vbMonday
                Myweekday = "星期一"
            Case vbTuesday
                Myweekday = "星期二"
            Case vbWednesday
                Myweekday = "星期三"
            Case vbThursday
                Myweekday = "星期四"
            Case vbFriday
                Myweekday = "星期五"
            Case vbSaturday
                Myweekday = "星期六"
        End Select
        MsgBox DateValue(StrDate) & " " & Myweekday
    Else
        MsgBox "请输入正确格式的日期!"
    End If
End Sub

' 119-4 计算两个日期的时间间隔
' DateDiff(interval, date1, date2, [ firstdayofweek, [ firstweekofyear ]] )
Option Explicit
Sub DateInterval()
    Dim StrDate As String
    StrDate = InputBox("请输入日期：")
    If Len(StrDate) = 0 Then Exit Sub
    If IsDate(StrDate) Then
        MsgBox DateValue(StrDate) & Chr(13) & "距离今天有" _
            & Abs(DateDiff("d", Date, StrDate)) & "天"
    Else
        MsgBox "请输入正确格式的日期!"
    End If
End Sub

' 119-5 获得指定时间间隔的日期
' dateadd
Option Explicit
Sub MyDateAdd()
    Dim StrDate As String
    StrDate = Application.InputBox(Prompt:="请输入间隔的天数：", Type:=1)
    If StrDate = False Then Exit Sub
    MsgBox StrDate & "天后的日期是" & DateAdd("d", StrDate, Date)
End Sub

' 119-6 格式化时间和日期
' Format
Option Explicit
Sub TimeDateFormat()
    Dim Str As String
    Str = Format(Now, "Medium Time") & Chr(13) _
        & Format(Now, "Long Time") & Chr(13) _
        & Format(Now, "Short Time") & Chr(13) _
        & Format(Now, "General Date") & Chr(13) _
        & Format(Now, "Long Date") & Chr(13) _
        & Format(Now, "Medium Date") & Chr(13) _
        & Format(Now, "Short Date")
    MsgBox Str
End Sub

' 范例120 使用字符串处理函数
Option Explicit
Sub StrFunctions()
    Dim Str As String
    Str = "Use String Functions"
    MsgBox "原始字符串：" & Str & Chr(13) _
        & "字符串长度：" & Len(Str) & Chr(13) _
        & "左边8个字符：" & Left(Str, 8) & Chr(13) _
        & "右边6个字符：" & Right(Str, 6) & Chr(13) _
        & """Str""出现在字符串的第" & InStr(Str, "Str") & "位" & Chr(13) _
        & "从左边第5个开始取6个字符：" & Mid(Str, 5, 6) & Chr(13) _
        & "转换为大写：" & UCase(Str) & Chr(13) _
        & "转换为小写：" & LCase(Str) & Chr(13)
End Sub

' 范例121 判断表达式是否为数值
' IsNumeric
Option Explicit
Sub MyNumeric()
    Dim r As Integer
    Dim rng As Range
    Dim Ynumber As String
    Dim Nnumber As String
    r = Cells(Rows.Count, 1).End(xlUp).Row
    For Each rng In Range("A1:A" & r)
        If IsNumeric(rng) Then
            Ynumber = Ynumber & rng.Address(0, 0) & vbTab & rng & vbCrLf
        Else
            Nnumber = Nnumber & rng.Address(0, 0) & vbTab & rng & vbCrLf
        End If
    Next
    MsgBox "数值单元格：" & vbCrLf & Ynumber & vbCrLf _
        & "非数值单元格：" & vbCrLf & Nnumber
End Sub

' 范例122 自定义数值格式
Option Explicit
Sub CustomDigitalFormat()
    Dim MyNumeric As Double
    Dim Str As String
    MyNumeric = 123456789
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications
    Str = Format(MyNumeric, "0.00") & vbCrLf _
        & Format(MyNumeric, "0%") & vbCrLf _
        & Format(MyNumeric, "#,##0.00") & vbCrLf _
        & Format(MyNumeric, "$#,##0.00") & vbCrLf _
        & Format(-(MyNumeric), "￥#,##0.00;(￥#,##0.00)")
        'Four sections	The first section applies to positive values, the second to negative values, 
        'the third to zeros, and the fourth to Null values.
   MsgBox Str
End Sub

' 范例123 使用Round函数进行四舍五入运算
Option Explicit
Sub Rounding()
    MsgBox Round(4.56789, 2)
End Sub
'VBA 内置的Round函数不是算术舍入,round(2.5)会变成2
Sub AmendmentsRound()
    MsgBox Round(2.5 + 0.0000001)
End Sub
Sub SheetsRound()
    MsgBox Application.Round(2.5, 0)
End Sub

' 范例124 使用Array函数创建数组
Option Explicit
Option Base 1
Sub Myarr()
    Dim arr As Variant
    Dim i As Integer
    arr = Array("王晓明", "吴胜玉", "周志国", "曹武伟", "张新发", "卓雪梅", "沈煜婷", "丁林平")
    For i = LBound(arr) To UBound(arr)
        Cells(i, 1) = arr(i)
    Next
End Sub

' 范例125 将字符串按指定的分隔符分开
' split
Option Explicit
Sub Splitarr()
    Dim Arr As Variant
    Arr = Split(Cells(1, 2), ",")
    Cells(1, 1).Resize(UBound(Arr) + 1, 1) = Application.Transpose(Arr) '一个range填入一个array
End Sub

' 范例126 使用动态数组去除重复值
Option Explicit
Sub Splitarr()
    Dim Splarr() As String
    Dim Arr() As String
    Dim Temp() As String
    Dim r As Integer
    Dim i As Integer
    On Error Resume Next
    Splarr = Split(Range("B1"), ",")
    For i = 0 To UBound(Splarr)
        Temp = Filter(Arr, Splarr(i)) 
        If UBound(Temp) < 0 Then ' 判断上届
            r = r + 1
            ReDim Preserve Arr(1 To r)
            Arr(r) = Splarr(i)
        End If
    Next
    Range("A1").Resize(r, 1) = Application.Transpose(Arr)
    'Range("a1").Resize(1, r) = Arr
    'Range("A1") = Join(Arr, ",")
End Sub

' 范例127 调用工作表函数

' 127-1 使用Sum函数求和
Option Explicit
Sub SumCell()
    Dim r As Integer
    Dim rng As Range
    Dim Dsum As Double
    r = Cells(Rows.Count, 1).End(xlUp).Row
    Set rng = Range("A1:F" & r)
    Dsum = Application.WorksheetFunction.Sum(rng)
    MsgBox rng.Address(0, 0) & "单元格的和为" & Dsum
End Sub

' 127-2 查找工作表中最大、最小值
' 如果是极值就标注颜色，并记录极值的数量
Option Explicit
Sub FindMaxAndMin()
    Dim r As Integer
    Dim Rng As Range, MyRng As Range
    Dim MaxCount As Integer, MainCount As Integer
    Dim Mymax As Double, Mymin As Double
    r = Cells(Rows.Count, 1).End(xlUp).Row
    Set MyRng = Range("A1:J" & r)
    For Each Rng In MyRng
        If Rng.Value = WorksheetFunction.max(MyRng) Then
            Rng.Interior.ColorIndex = 3
            MaxCount = MaxCount + 1
            Mymax = Rng.Value
        ElseIf Rng.Value = WorksheetFunction.min(MyRng) Then
            Rng.Interior.ColorIndex = 5
            MainCount = MainCount + 1
            Mymin = Rng.Value
        Else
            Rng.Interior.ColorIndex = 0
        End If
    Next
    MsgBox "最大值是：" & Mymax & "，共有" & MaxCount & "个。" _
        & Chr(13) & "最小值是：" & Mymin & "，共有" & MainCount & "个。"
End Sub

' 127-3 不重复值的录入
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
    With Target
        If .Column <> 1 Or .Count > 1 Then Exit Sub '第一列不能有重复值
        If WorksheetFunction.CountIf(Range("A:A"), .Value) > 1 Then
            .Select
            MsgBox "不能输入重复的数据!", 64
            Application.EnableEvents = False 
            '如果不禁用事件，那么在清除重复值的过程中会不断地触发工作表的Change事件，从而造成代码运行的死循环。
            .Value = ""
            Application.EnableEvents = True
        End If
    End With
End Sub

' 范例128 使用个人所得税自定义函数
Option Explicit
Public Function PITax(Income, Optional Threshold) As Double
    Dim Rate As Double
    Dim Deduction As Double
    Dim Taxliability As Double
    If IsMissing(Threshold) Then Threshold = 2000
    Taxliability = Income - Threshold
    Select Case Taxliability
        Case 0 To 500
            Rate = 0.05
            Deduction = 0
        Case 500.01 To 2000
            Rate = 0.1
            Deduction = 25
        Case 2000.01 To 5000
            Rate = 0.15
            Deduction = 125
        Case 5000.01 To 20000
            Rate = 0.2
            Deduction = 375
        Case 20000.01 To 40000
            Rate = 0.25
            Deduction = 1375
        Case 40000.01 To 60000
            Rate = 0.3
            Deduction = 3375
        Case 60000.01 To 80000
            Rate = 0.35
            Deduction = 6375
        Case 80000.01 To 10000
            Rate = 0.4
            Deduction = 10375
        Case Else
            Rate = 0.45
            Deduction = 15375
    End Select
    If Taxliability <= 0 Then
        PITax = 0
    Else
        PITax = Application.Round(Taxliability * Rate - Deduction, 2)
    End If
End Function

' 范例129 使用人民币大写函数
Option Explicit
Public Function YuanCapital(Amountin)
'重点在[DBnum2]，转换为中文大写
    '把.替换为元
    YuanCapital = Replace(Application.Text(Round(Amountin + 0.00000001, 2), "[DBnum2]"), ".", "元")
    '
    YuanCapital = IIf(Left(Right(YuanCapital, 3), 1) = "元", Left(YuanCapital, Len(YuanCapital) - 1) & "角" & Right(YuanCapital, 1) & "分", _IIf(Left(Right(YuanCapital, 2), 1) = "元", YuanCapital & "角整", IIf(YuanCapital = "零", "", YuanCapital & "元整")))
    YuanCapital = Replace(Replace(Replace(Replace(YuanCapital, "零元零角", ""), "零元", ""), "零角", "零"), "-", "负")
End Function

' 范例130 判断工作表是否为空表
Option Explicit
Function IsBlankSht(Sh As Variant) As Boolean
    If TypeName(Sh) = "String" Then Set Sh = Worksheets(Sh)
    If Application.CountA(Sh.UsedRange.Cells) = 0 Then '是否有cell被使用过
        IsBlankSht = True
    End If
End Function
Sub DelBlankSht()
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Sheets
        If IsBlankSht(Sh) Then
            Application.DisplayAlerts = False
            Sh.Delete
            Application.DisplayAlerts = True
        End If
    Next
    Set Sh = Nothing
End Sub

' 范例131 查找指定工作表
Option Explicit
Function ExistSh(Sh As String) As Boolean
    Dim Sht As Worksheet
    On Error Resume Next
    Set Sht = Sheets(Sh)
    If Err = 0 Then ExistSh = True
    Set Sht = Nothing
End Function
Sub NotSht()
    Dim Sh As String
    Sh = InputBox("请输入工作表名称：")
    If Len(Sh) > 0 Then
        If Not ExistSh(Sh) Then
            MsgBox "对不起," & Sh & "工作表不存在!"
        Else
            Sheets(Sh).Select
        End If
    End If
End Sub

' 范例132 查找指定工作簿
Option Explicit
Function ExistWorkbook(WbName As String) As Boolean
    Dim Wb As Workbook
    On Error Resume Next
    Set Wb = Workbooks(WbName)
    If Err = 0 Then ExistWorkbook = True '如果没有错误就打开了
    Set Wb = Nothing
End Function
Sub NotWorkbook()
    Dim Wb As String
    Wb = InputBox("请输入工作簿名称：")
    If Len(Wb) > 0 Then
        If Not (ExistWorkbook(Wb)) Then
            MsgBox Wb & "工作簿没有打开!"
        End If
    End If
End Sub

' 范例133 取得应用程序的安装路径
Option Explicit
Function GetSetupPath(AppName As String)
    Dim Wsh As Object
    Set Wsh = CreateObject("Wscript.Shell")
    GetSetupPath = Wsh.RegRead("HKEY_LOCAL_MACHINE\Software" _
        & "\Microsoft\Windows\CurrentVersion\App Paths\" _
        & AppName & "\Path")
    Set Wsh = Nothing
End Function
Sub WinRARPath()
    MsgBox GetSetupPath("WinRAR.exe")
End Sub





