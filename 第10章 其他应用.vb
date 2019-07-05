' 第10章 其他应用

' 范例144 检查电脑名称
Option Explicit
Private Sub Workbook_Open()
    Dim myName As String
    myName = Environ("Computername")
    If myName <> "YUANZHUPING" Then
        MsgBox "对不起，您不是合法用户，文件将关闭!"
        ThisWorkbook.Close
    End If
End Sub

' 范例144 检查电脑名称
Option Explicit
Sub TimingOff()
    Shell ("at 20:09 Shutdown.exe -s")
End Sub

' 范例146 保护VBA代码

' 范例147 使用数字签名

' 范例148 打开指定网页
Option Explicit
Sub OpenTheWeb()
    ActiveWorkbook.FollowHyperlink _
        Address:="http://www.microsoft.com/zh/cn/default.aspx", _
        NewWindow:=True
End Sub

' 范例149 自定义“加载项”选项卡
Option Explicit
Sub Addinstab()
    Dim myBarPopup As CommandBarPopup
    Dim myBar As CommandBar
    Dim ArrOne As Variant
    Dim ArrTwo As Variant
    Dim ArrThree As Variant
    Dim ArrFour As Variant
    Dim i As Byte
    On Error Resume Next
    ArrOne = Array("凭证打印", "账簿打印", "报表打印")
    ArrThree = Array("会计凭证", "会计账簿", "会计报表")
    ArrTwo = Array(281, 283, 285)
    ArrFour = Array(9893, 284, 9590)
    With Application.CommandBars("Worksheet menu bar")
        .Reset
        Set myBarPopup = .Controls.Add(msoControlPopup)
        With myBarPopup
            .Caption = "打印"
            For i = 0 To UBound(ArrOne)
                With .Controls.Add(msoControlButton)
                    .Caption = ArrOne(i)
                    .FaceId = ArrTwo(i)
                    .OnAction = "myOnAction"
                End With
            Next
        End With
    End With
    Application.CommandBars("MyToolbar").Delete
    Set myBar = Application.CommandBars.Add("MyToolbar")
    With myBar
        .Visible = True
        For i = 0 To UBound(ArrThree)
            With .Controls.Add(msoControlButton)
                .Caption = ArrThree(i)
                .FaceId = ArrFour(i)
                .OnAction = "myOnAction"
                .Style = msoButtonIconAndCaptionBelow
            End With
        Next
    End With
    Set myBarPopup = Nothing
    Set myBar = Nothing
End Sub
Public Sub myOnAction()
    MsgBox "您选择了：" & Application.CommandBars.ActionControl.Caption
End Sub
Sub DeleteToolbar()
    On Error Resume Next
    Application.CommandBars("MyToolbar").Delete
    Application.CommandBars("Worksheet menu bar").Reset
End Sub

' 范例150 使用右键快捷菜单

' 150-1 使用右键快捷菜单增加菜单项
Option Explicit
Sub MyCmb()
    Dim MyCmb As CommandBarButton
    With Application.CommandBars("Cell")
        .Reset
        Set MyCmb = .Controls.Add(Type:=msoControlButton, _
            ID:=2521, Temporary:=True)
    End With
    MyCmb.BeginGroup = True
    Set MyCmb = Nothing
End Sub

' 150-2 自定义右键快捷菜单
Option Explicit
Sub Mycell()
    With Application.CommandBars.Add("Mycell", msoBarPopup)
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "会计凭证"
            .FaceId = 9893
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "会计账簿"
            .FaceId = 284
        End With
        With .Controls.Add(Type:=msoControlPopup)
            .Caption = "会计报表"
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "月报"
                .FaceId = 9590
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "季报"
                .FaceId = 9591
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "年报"
                .FaceId = 9592
            End With
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "凭证打印"
            .FaceId = 9614
            .BeginGroup = True
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "账簿打印"
            .FaceId = 707
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "报表打印"
            .FaceId = 986
        End With
    End With
End Sub

' 150-3 使用快捷菜单输入数据
Option Explicit
Sub Mycell()
    Dim arr As Variant
    Dim i As Integer
    Dim Mycell As CommandBar
    On Error Resume Next
    Application.CommandBars("Mycell").Delete
    arr = Array("经理室", "办公室", "生技科", "财务科", "营业部")
    Set Mycell = Application.CommandBars.Add("Mycell", msoBarPopup)
    For i = 0 To 4
        With Mycell.Controls.Add(1)
            .Caption = arr(i)
            .OnAction = "MyOnAction"
        End With
    Next
End Sub
Sub MyOnAction()
    ActiveCell = Application.CommandBars.ActionControl.Caption
End Sub

' 150-4 禁用右键快捷菜单
Option Explicit
Sub DisableMenu()
    Dim myBar As CommandBar
    For Each myBar In CommandBars
        If myBar.Type = msoBarTypePopup Then
            myBar.Enabled = False
        End If
    Next
End Sub
Sub EnableMenu()
    Dim myBar As CommandBar
    For Each myBar In CommandBars
        If myBar.Type = msoBarTypePopup Then
            myBar.Enabled = True
        End If
    Next
End Sub

' 范例151 VBE相关操作

' 151-1 添加模块和过程
Option Explicit
Sub NowModule()
    Dim VBC As VBComponent
    Set VBC = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    VBC.Name = "NowModule"
    With VBC.CodeModule
        If .Lines(1, 1) <> "Option Explicit" Then
           .InsertLines 1, "Option Explicit"
        End If
        .InsertLines 2, "Sub ProcessOne()"
        .InsertLines 3, vbTab & "MsgBox ""这是第一个过程!"""
        .InsertLines 4, "End Sub"
        .AddFromString "Sub ProcessTwo()" & Chr(13) & vbTab _
            & "MsgBox ""这是第二个过程!""" & Chr(13) & "End Sub"
    End With
    Set VBC = Nothing
End Sub

' 151-2 建立事件过程
Option Explicit
Sub AddMatter()
    Dim Sh As Worksheet
    Dim r As Integer
    For Each Sh In Worksheets
        If Sh.Name = "Matter" Then Exit Sub
    Next
    Set Sh = Sheets.Add(After:=Sheets(Sheets.Count))
    Sh.Name = "Matter"
    Application.VBE.MainWindow.Visible = True
    With ThisWorkbook.VBProject.VBComponents(Sh.CodeName).CodeModule
        r = .CreateEventProc("SelectionChange", "Worksheet")
        .ReplaceLine r + 1, vbTab & "If Target.Count = 1 Then" _
            & Chr(13) & Space(8) & "MsgBox ""你选择了"" & Target.Address(0, 0) & ""单元格!""" _
            & Chr(13) & vbTab & "End If"
    End With
    Application.VBE.MainWindow.Visible = False
    Set Sh = Nothing
End Sub

' 151-3 模块的导入与导出
Option Explicit
Sub CopyModule()
    Dim Nowbook As Workbook
    Dim MyTxt As String
    MyTxt = ThisWorkbook.Path & "\AddMatter.txt"
    ThisWorkbook.VBProject.VBComponents("AddMatter").Export MyTxt
    Set Nowbook = Workbooks.Add
    With Nowbook
        .SaveAs Filename:=ThisWorkbook.Path & "\CopyModule.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        .VBProject.VBComponents.Import MyTxt
        .Close Savechanges:=True
    End With
    Kill MyTxt
End Sub

' 151-4 删除VBA代码
Option Explicit
Sub DelMacro()
    Dim Wb As Workbook
    Dim Vbc As VBComponent
    Set Wb = Workbooks.Open(ThisWorkbook.Path & "\DelMacro.xlsm")
    With Wb
        For Each Vbc In .VBProject.VBComponents
            If Vbc.Type <> vbext_ct_Document Then
                Select Case Vbc.Name
                    Case "ShowForm"
                        Vbc.CodeModule.DeleteLines 3, 3
                    Case "MyTreeView"
                    Case Else
                        .VBProject.VBComponents.Remove Vbc
                End Select
            End If
        Next
        .SaveAs FileName:=ThisWorkbook.Path & "\Backup.xlsm", _
                FileFormat:=xlOpenXMLWorkbookMacroEnabled
        .Close False
    End With
    Set Wb = Nothing
    Set Vbc = Nothing
End Sub

' 范例152 优化代码

' 152-1 关闭屏幕刷新
' Application.ScreenUpdating = False
Option Explicit
Sub CloseScreen()
    Dim i As Integer
    Dim StartTime As Single
    Dim TimeOne As String
    Dim TimeTwo As String
    StartTime = Timer
    For i = 1 To 30000
        Cells(1, 1) = i
    Next
    TimeOne = Format(Timer - StartTime, "0.00000") & "秒"
    Application.ScreenUpdating = False
    StartTime = Timer
    For i = 1 To 30000
        Cells(1, 1) = i
    Next
    TimeTwo = Format(Timer - StartTime, "0.00000") & "秒"
    Application.ScreenUpdating = True
    MsgBox "第一次运行时间：" & TimeOne & vbCrLf & "第二次运行时间：" & TimeTwo
End Sub

' 152-2 使用工作表函数
Option Explicit
Sub ShtFunctions()
    Dim i As Long
    Dim StartTime As Single
    Dim MySum As Double
    Dim TimeOne As String
    Dim TimeTwo As String
    StartTime = Timer
    For i = 1 To 40000
        MySum = MySum + Cells(i, 1)
    Next
    Cells(1, 2) = MySum
    TimeOne = Format(Timer - StartTime, "0.00000") & "秒"
    StartTime = Timer
    Cells(2, 2) = Application.Sum(Range("A1:A40000"))
    TimeTwo = Format(Timer - StartTime, "0.00000") & "秒"
    MsgBox "第一次运行时间：" & TimeOne & vbCrLf & "第二次运行时间：" & TimeTwo
End Sub

' 152-3 使用更快的VBA方法
Option Explicit
Sub UseMethods()
    Dim MyArr As Variant
    Dim i As Integer
    Dim StartTime As Single
    Dim TimeOne As String
    Dim TimeTwo As String
    MyArr = Range("A1:A20000").Value
    StartTime = Timer
    For i = 20000 To 1 Step -1
        If Cells(i, 1) = "VBA方法" Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next
    TimeOne = Format(Timer - StartTime, "0.00000") & "秒"
    Range("A1:A20000").Value = MyArr
    StartTime = Timer
    Range("A1:A20000").Replace "VBA方法", ""
    Range("A1:A20000").SpecialCells(4).EntireRow.Delete
    TimeTwo = Format(Timer - StartTime, "0.00000") & "秒"
    Range("A1:A20000").Value = MyArr
    MsgBox "第一次运行时间：" & TimeOne & Chr(13) & "第二次运行时间：" & TimeTwo
End Sub

' 152-4 使用With语句引用对象
Option Explicit
Sub ReferenceObject()
    Dim i As Integer
    Dim StartTime As Single
    Dim TimeOne As String
    Dim TimeTwo As String
    StartTime = Timer
    For i = 1 To 10000
        Worksheets("Sheet1").Range("A1").FormulaR1C1 = "=RAND()"
        Worksheets("Sheet1").Range("A1").Interior.ColorIndex = Int(56 * Rnd() + 1)
    Next
    TimeOne = Format(Timer - StartTime, "0.00000") & "秒"
    StartTime = Timer
    With Worksheets("Sheet1").Range("A1")
        For i = 1 To 10000
            .FormulaR1C1 = "=RAND()"
            .Interior.ColorIndex = Int(56 * Rnd() + 1)
        Next
    End With
    TimeTwo = Format(Timer - StartTime, "0.00000") & "秒"
    MsgBox "第一次运行时间：" & TimeOne & vbCrLf & "第二次运行时间：" & TimeTwo
End Sub

' 152-5 简化代码
' 删除录制宏时的冗余代码
Option Explicit
Sub Simplification()
    Dim i As Integer
    Dim StartTime As Single
    Dim TimeOne As String
    Dim TimeTwo As String
    StartTime = Timer
    For i = 1 To 5000
        Sheets("Sheet2").Select
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "100"
    Next
    TimeOne = Format(Timer - StartTime, "0.00000") & "秒"
    StartTime = Timer
    For i = 1 To 5000
        Sheets("Sheet2").Range("A1") = 100
    Next
    TimeTwo = Format(Timer - StartTime, "0.00000") & "秒"
    MsgBox "第一次运行时间：" & TimeOne & vbCrLf & "第二次运行时间：" & TimeTwo
End Sub






