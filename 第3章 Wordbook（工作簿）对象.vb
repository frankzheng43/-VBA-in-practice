' 范例38 引用工作簿的方法

' 38-1 使用工作簿名称引用工作簿
Option Explicit
Sub WbPath()
    MsgBox Workbooks("38-1 使用工作簿名称.xlsm").Path
End Sub

' 38-2 使用工作簿索引号引用工作簿
Option Explicit
Sub WbName()
    MsgBox "第一个打开的工作簿名字为：" & Workbooks(1).Name
End Sub
Sub WbFullName()
    MsgBox "包括完整路径的工作簿名称为：" & Workbooks(1).FullName
End Sub

' 38-3 使用ThisWorkbook属性引用工作簿
Option Explicit
Sub WbClose()
    ThisWorkbook.Close SaveChanges:=False
End Sub

' 38-4 使用ActiveWorkbook属性引用工作簿
Option Explicit
Sub WbActive()
    MsgBox "当前活动工作簿名字为：" & ActiveWorkbook.Name
End Sub

' 范例39 新建工作簿
Option Explicit
Sub AddNowwb()
    Dim AddNowwb As Workbook
    Dim ShtName As Variant
    Dim Arr As Variant
    Dim i As Integer
    Dim MyInNewWb As Integer
    MyInNewWb = Application.SheetsInNewWorkbook '保存Excel自动插入到新工作簿中的工作表数目
    Arr = Array("品名", "单价", "数量", "金额")
    ShtName = Array("01月", "02月", "03月", "04月", "05月", "06月", "07月", "08月", "09月", "10月", "11月", "12月")
    Application.SheetsInNewWorkbook = 12
    Set AddNowwb = Workbooks.Add
    With AddNowwb
        For i = 1 To 12
            With .Sheets(i)
                .Name = ShtName(i - 1)
                .Range("A1").Resize(1, UBound(Arr) + 1) = Arr
            End With
        Next
        .SaveAs Filename:=ThisWorkbook.Path & "\" & "存货明细.xlsx"
        .Close Savechanges:=True
    End With
    Application.SheetsInNewWorkbook = MyInNewWb '还原Application.SheetsInNewWorkbook的值
    Set AddNowwb = Nothing
End Sub

' 范例40 打开指定的工作簿
Option Explicit
Sub Openwb()
    Workbooks.Open ThisWorkbook.Path & "\123.xlsx"
End Sub

' 范例41 判断指定工作簿是否被打开

' 41-1 遍历Workbooks集合方法
Option Explicit
Sub WbIsOpenOne()
    Dim Wb As Workbook
    Dim WbName As String
    WbName = "abc.xlsx"
    For Each Wb In Workbooks
        If Wb.Name = WbName Then
            MsgBox "工作簿" & WbName & "已经被打开!"
            Exit Sub
        End If
    Next
    MsgBox "工作簿" & WbName & "没有被打开!"
End Sub

' 41-2 使用错误处理方法
Option Explicit
Sub WbIsOpenTwo()
    Dim Wb As Workbook
    Dim WbName As String
    WbName = "abc.xlsx"
    On Error GoTo line
    '使用Set语句将Workbook对象引用赋予变量Wb，如果“abc.xlsx”工作簿没有被打开，将发生下标越界错误的运行时错误。
    Set Wb = Application.Workbooks(WbName)
    MsgBox "工作簿" & WbName & "已经被打开!"
    Exit Sub
line:
    MsgBox "工作簿" & WbName & "没有被打开!"
End Sub

' 范例42 关闭工作簿不弹出保存对话框
Option Explicit
Sub wbCloseOne()
    ThisWorkbook.Close SaveChanges:=False
End Sub
Sub wbCloseTwo()
    ThisWorkbook.Saved = True
    ThisWorkbook.Close
End Sub
Sub wbCloseThree()
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub

' 范例43 禁用工作簿的关闭按钮
Option Explicit
Dim WbClose As Boolean
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If WbClose = False Then
        Cancel = True
        MsgBox "请使用""关闭""按钮关闭工作簿!", 48, "提示"
    End If
End Sub
Public Sub CloseWb()
    WbClose = True
    ThisWorkbook.Close
End Sub

' 范例44 保存工作簿的方法
Option Explicit
Sub SaveWb()
    ThisWorkbook.Save
End Sub
Sub SaveAsWb()
    On Error Resume Next
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\SaveAsWb.xlsm"
End Sub
Sub SaveCopyWb()
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\SaveCopyWb.xlsm"
End Sub

' 范例45 保存指定工作表为工作簿
Option Explicit
Sub ShtCopy()
    On Error GoTo line '有可能已存在，所以需要on error
    Sheet2.Copy
    ActiveWorkbook.Close SaveChanges:=True, _
        Filename:=ThisWorkbook.Path & "\ShtCopy.xlsx"
    Exit Sub
line:
    ActiveWorkbook.Close False '关闭新建的工作簿并且不予保存。
End Sub
Sub ArrShtCopy()
    On Error GoTo line
    Worksheets(Array("Sheet2", "Sheet3")).Copy
    ActiveWorkbook.Close SaveChanges:=True, _
        Filename:=ThisWorkbook.Path & "\ShtCopy.xlsx"
    Exit Sub
line:
    ActiveWorkbook.Close False
End Sub

' 范例46 不打开工作簿取得其他工作簿数据

' 46-1 使用公式取得数据
Option Explicit
Sub UsingTheFormula()
    Dim Temp As String
    Temp = "'" & ThisWorkbook.Path & "\[数据.xlsx]Sheet1'!" '将引用工作表的完整路径赋予变量Temp
    With Sheet1.Range("A1:F22")
        .FormulaR1C1 = "=" & Temp & "RC"
        .Value = .Value
    End With
End Sub

' 46-2 使用GetObject函数取得数据
Option Explicit
Sub UseGetObject()
    Dim Wb As Workbook
    Dim Temp As String
    Temp = ThisWorkbook.Path & "\数据.xlsx"
    Set Wb = GetObject(Temp)
    ' 当GetObject函数指定的对象被激活后，就可以在代码中使用对象变量Wb来访问指定对象的属性和方法。
    With Wb.Sheets(1).Range("A1").CurrentRegion
        Range("A1").Resize(.Rows.Count, .Columns.Count) = .Value
    End With
    ' 使用GetObject函数返回对象的引用时，虽然在窗口中看不到对象的实例，但实际上是打开的，所以需要用Close语句将其关闭。
    Wb.Close False
    Set Wb = Nothing
End Sub

' 46-3 隐藏Application对象取得数据
Option Explicit
Sub HideApplication()
    Dim MyApp As New Application
    Dim Sht As Worksheet
    Dim Temp As String
    Temp = ThisWorkbook.Path & "\数据.xlsx"
    MyApp.Visible = False
    Set Sht = MyApp.Workbooks.Open(Temp).Sheets(1)
    With Sht.Range("A1").CurrentRegion
        Range("A1").Resize(.Rows.Count, .Columns.Count) = .Value
    End With
    MyApp.Quit
    Set MyApp = Nothing
    Set Sht = Nothing
End Sub

' 46-4 使用ExecuteExcel4Macro方法取得数据
Option Explicit
Sub UsingMacroFunction()
    Dim RCount As Long
    Dim CCount As Long
    Dim Temp As String
    Dim Temp1 As String
    Dim Temp2 As String
    Dim Temp3 As String
    Dim r As Long
    Dim c As Long
    Dim arr() As Variant
    Temp = "'" & ThisWorkbook.Path & "\[数据.xlsx]Sheet1'!"
    Temp1 = "Counta(" & Temp & Rows(1).Address(, , xlR1C1) & ")"
    CCount = Application.ExecuteExcel4Macro(Temp1)
    Temp2 = "Counta(" & Temp & Columns("A").Address(, , xlR1C1) & ")"
    RCount = Application.ExecuteExcel4Macro(Temp2)
    ReDim arr(1 To RCount, 1 To CCount)
    For r = 1 To RCount
        For c = 1 To CCount
            Temp3 = Temp & Cells(r, c).Address(, , xlR1C1)
            arr(r, c) = Application.ExecuteExcel4Macro(Temp3)
        Next
    Next
    Range("A1").Resize(RCount, CCount).Value = arr
End Sub

' 46-5 使用SQL连接取得数据
Option Explicit
Sub UsingSQL()
    Dim Sql As String
    Dim j As Integer
    Dim r As Integer
    Dim Cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    With Sheet1
        .Cells.Clear
        Set Cnn = New ADODB.Connection
        With Cnn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .ConnectionString = "Extended Properties=Excel 12.0;" _
                & "Data Source=" & ThisWorkbook.Path & "\数据.xlsx"
            .Open
        End With
        Set rs = New ADODB.Recordset
        Sql = "Select * From [Sheet1$]"
        rs.Open Sql, Cnn, adOpenKeyset, adLockOptimistic
        For j = 0 To rs.Fields.Count - 1
            .Cells(1, j + 1) = rs.Fields(j).Name
        Next
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A" & r + 1).CopyFromRecordset rs
    End With
    rs.Close
    Cnn.Close
    Set rs = Nothing
    Set Cnn = Nothing
End Sub





