' 范例57 取得Excel版本信息
Option Explicit
Sub ExcelVersion()
    Dim MyVersion As String
    Select Case Application.Version
        Case "8.0"
            MyVersion = "97"
        Case "9.0"
            MyVersion = "2000"
        Case "10.0"
            MyVersion = "2002"
        Case "11.0"
            MyVersion = "2003"
        Case "12.0"
            MyVersion = "2007"
        Case Else
            MyVersion = "未知版本"
    End Select
    MsgBox "Excel 的版本是： " & MyVersion
End Sub

' 范例58 取得当前用户名称
Option Explicit
Sub ToUserName()
    MsgBox "当前用户是： " & Application.UserName
End Sub

' 范例59 实现简单的计时器功能
Option Explicit
Sub StartTimer()
    Sheet1.Cells(1, 2) = Sheet1.Cells(1, 2) + 1
    'https://docs.microsoft.com/en-us/office/vba/api/excel.application.ontime
    Application.OnTime Now + TimeValue("00:00:01"), "StartTimer"
End Sub
Sub EndTimer()
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "StartTimer", , False
    Sheet1.Cells(1, 2) = 0
End Sub

' 范例60 屏蔽、更改组合键功能
Option Explicit
Private Sub Workbook_Activate()
    Application.OnKey "^{c}", "MyOnKey"
End Sub
Private Sub Workbook_Deactivate()
    Application.OnKey "^{c}"
End Sub

' 范例61 设置Excel标题栏
Option Explicit
Sub ModifyTheTitle()
    Application.Caption = "修改标题栏"
End Sub
Sub RemoveTheTitle()
    Application.Caption = vbNullChar
End Sub
Sub RrestoreTheTitle()
    Application.Caption = Empty
End Sub

' 范例62 自定义Excel状态栏
' 左下角那个
Option Explicit
Sub MyStatusBar()
    Dim r As Long
    Dim i As Long
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 1 To r
            .Cells(i, 3) = .Cells(i, 1) + .Cells(i, 2)
            Application.StatusBar = "正在计算" & .Cells(i, 3).Address(0, 0) & " 的数据..."
        Next
    End With
    Application.StatusBar = False
End Sub

' 范例63 灵活关闭Excel
Option Explicit
Sub FlexibleClose()
    If Workbooks.Count > 1 Then
        ThisWorkbook.Close
    Else
        Application.Quit
    End If
End Sub

' 范例64 暂停代码的运行
Option Explicit
Private Sub UserForm_Activate()
    Dim i As Integer
    For i = 1 To 10
        Label1.Caption = "这是个演示窗体,将在" & 11 - i & "秒后自动关闭!"
        Application.Wait Now() + TimeValue("00:00:01")
        DoEvents
    Next
    Unload Me
End Sub

Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub UserForm_Activate()
    Dim Str As String
    Dim i As Integer
    Str = "这是一个模拟打字效果的演示。"
    For i = 1 To Len(Str)
        TextBox1 = Left(Str, i)
        Sleep 400
        DoEvents
    Next
End Sub

' 范例65 防止用户中断代码运行
Option Explicit
Sub ProhibitionEsc()
    Dim i As Integer
    Application.EnableCancelKey = xlDisabled
    For i = 1 To 10000
        Cells(1, 1) = i
    Next
End Sub

' 范例66 隐藏Excel主窗口

' 66-1 设置Visible属性为False
Option Explicit
Private Sub Workbook_Open()
    Application.Visible = False
    UserForm1.Show
End Sub

Option Explicit
Private Sub CommandButton1_Click()
    Application.Visible = True
    Unload Me
End Sub

' 66-2 将窗口移出屏幕
Option Explicit
Private Sub Workbook_Open()
    Application.WindowState = xlNormal
    Application.Left = 10000
    UserForm1.StartUpPosition = 2
    UserForm1.Show
End Sub






