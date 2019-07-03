' 范例113 使用Msgbox函数显示消息框
Option Explicit
Sub Mymsg()
    Dim Mymsg As Integer
    Mymsg = MsgBox("文件即将关闭,是否保存所作的修改?", vbYesNoCancel + vbQuestion)
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
    Select Case Mymsg
        Case vbYes
            ThisWorkbook.Save
        Case vbNo
            ThisWorkbook.Saved = True
        Case vbCancel
            Exit Sub
    End Select
    ThisWorkbook.Close
End Sub


' 范例114 自动关闭的消息框

' 114-1 使用WshShell.Popup方法关闭消息框
Option Explicit
Sub AutoClose()
    Dim MyShell As Object
    Set MyShell = CreateObject("Wscript.Shell")
    MyShell.Popup "程序已执行完毕!", 2, "运行提示", 64
    'WshShell.Popup（strText，[natSecondsToWait], [strTitle], [natType]） = intButton
    Set MyShell = Nothing
End Sub

' 114-2 使用API函数关闭消息框
Option Explicit
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElaspe As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Dim MyTimer As Long
Sub AutoClose()
    MyTimer = SetTimer(0, 0, 2000, AddressOf CloseMsg)
    MsgBox "程序已执行完毕!", 64
End Sub
Sub CloseMsg(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)
    Application.SendKeys "~", True
    KillTimer 0, MyTimer
End Sub

' 范例115 使用InputBox函数输入数据
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const EM_SETPASSWORDCHAR = &HCC
Public lTimeID As Long
Sub TimeProc(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
    Dim hwd As Long
    hwd = FindWindow("#32770", "Microsoft Excel")
    If hwd <> 0 Then
        hwd = FindWindowEx(hwd, 0, "edit", vbNullString)
        SendMessage hwd, EM_SETPASSWORDCHAR, 42, 0
        timeKillEvent lTimeID
    End If
  End Sub
Sub PassInput()
    Dim Str As String
    lTimeID = timeSetEvent(10, 0, AddressOf TimeProc, 1, 1)
    Str = InputBox("请输入密码：", "Microsoft Excel")
    If Str = "12345678" Then
        MsgBox "密码输入正确!"
    Else
        MsgBox "密码输入错误!"
    End If
End Sub

' 范例116 使用InputBox方法

' 116-1 输入指定类型的数据
Option Explicit
Sub EnterNumbers()
    Dim myInput As Long
    Dim r As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        myInput = Application.InputBox(Prompt:="输入数字：", Type:=1) 'type指示输入的类型
        If myInput <> False Then
            .Cells(r + 1, 1).Value = myInput '如果是正确的，那么就输入到单元格里
        End If
    End With
End Sub

' 116-2 获得选定的单元格区域
Option Explicit
Sub SelecteRange()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox(Prompt:="请选择单元格区域：", Type:=8) 'type为8返回一个range
    'The following table lists the values that can be passed in the Type argument. Can be one or a sum of the values. 
    'For example, for an input box that can accept both text and numbers, set Type to 1 + 2.
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.application.inputbox
    rng.Interior.ColorIndex = 15
    Set rng = Nothing
End Sub

' 范例117 使用内置对话框

' 117-1 调用Excel内置对话框
' 就是平时打开某些选项（如选择字体段落）时的对话框
Option Explicit
Sub MyFont()
    If TypeName(Selection) = "Range" Then
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.xlbuiltindialog
        Application.Dialogs(xlDialogActiveCellFont).Show _
            arg1:="黑体", arg2:="加粗 倾斜", arg3:=30, _
            arg4:=True, arg10:=3, arg11:=False
    End If
End Sub

' 117-2 获取所选文件的文件名和路径
Option Explicit
Sub FileNameAndPath()
    Dim FilterList As String
    Dim FileName As Variant
    Dim i As Integer
    Dim Str As String
    FilterList = "All Files (*.*),*.*,Excel Files(*.xlsm),*.xlsm"
    FileName = Application.GetOpenFilename(FileFilter:=FilterList, _
                Title:="请选择文件", MultiSelect:=True)
    If IsArray(FileName) Then
        For i = 1 To UBound(FileName)
            Str = Str & FileName(i) & Chr(10)
        Next
        MsgBox Str
    End If
End Sub

' 117-3 使用“另存为”对话框备份文件
Option Explicit
Sub FileBackup()
    Dim FileName As String
    Dim FilePath As String
    Dim FilterList As String
    On Error GoTo line
    FilePath = "D:\" & Format(Date, "yyyymmdd") & "备份文件.xlsx"
    FilterList = "Excel Files(*.xlsx),*.xlsx,All Files (*.*),*.*" '显示在对话框中的文件类型
    FileName = Application.GetSaveAsFilename(InitialFileName:=FilePath, FileFilter:=FilterList, Title:="文件备份")
    If FileName <> "False" Then
        Sheet2.Copy
        ActiveWorkbook.Close SaveChanges:=True, FileName:=FileName
    End If
    Exit Sub
line:
    ActiveWorkbook.Close False
End Sub

' 范例118 调用操作系统的“关于”对话框
Option Explicit
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Sub SystemDialogBox()
    Dim ApphWnd As Long
    ApphWnd = FindWindow("XLMAIN", Application.Caption)
    ShellAbout ApphWnd, "财务处理系统", "yuanzhuping@yeah.net  0513-86XXXX30", 0
End Sub









