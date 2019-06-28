' 范例18 引用工作表的方法

' 18-1 使用工作表名称
Option Explicit
Sub ShtName()
    Worksheets("Sheet2").Range("A1") = "Excel 2007"
End Sub

' 18-2 使用工作表索引号
Option Explicit
Sub ShtIndex()
    Worksheets(Worksheets.Count).Select
End Sub
' 获得索引
MsgBox Worksheets("Sheet3").Index

' 18-3 使用工作表代码名称
Option Explicit
Sub ShtCodeName()
    Sheet3.Select
End Sub

' 范例19 选择工作表的方法
Option Explicit
Sub ShtSelect()
    MsgBox "下面将选择" & Sheet2.Name & "工作表"
    Sheet2.Select
    MsgBox "下面将激活" & Sheet3.Name & "工作表"
    Sheet3.Activate
End Sub
Sub SelectSht()
    Dim Sht As Worksheet
    For Each Sht In Worksheets
        Sht.Select False 'replace 参数为false，全选。
    Next
End Sub
Sub SelectSheets()
    Worksheets.Select
End Sub
Sub ArraySheets()
    Worksheets(Array(1, 3)).Select
End Sub

' 范例20 遍历工作表的方法
Option Explicit
Sub TraversalShtOne()
    Dim i As Integer
    Dim Str As String
    For i = 1 To Worksheets.Count
        Str = Str & Worksheets(i).Name & vbCrLf
    Next
    MsgBox "工作簿中含有以下工作表：" & vbCrLf & Str
End Sub
Sub TraversalShtTwo()
    Dim Sht As Worksheet
    Dim Str As String
    For Each Sht In Worksheets
        Str = Str & Sht.Name & vbCrLf
    Next
    MsgBox "工作簿中含有以下工作表：" & vbCrLf & Str
End Sub

' 范例21 工作表的添加与删除
Option Explicit
Sub ShtAddOne()
    Worksheets.Add.Name = "数据"
End Sub
Sub ShtAddTwo()
    Dim i As Integer
    Dim Sht As Worksheet
    With Worksheets
        For i = 1 To 6
            Set Sht = .Add(after:=Worksheets(.Count)) '在已有工作表后加
            Sht.Name = i
        Next
    End With
    Set Sht = Nothing
End Sub
Sub ShtDel()
    Dim Sht As Worksheet
    Application.DisplayAlerts = False
    For Each Sht In Worksheets
        If Sht.Name <> "工作表的添加与删除" Then
            Sht.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Set Sht = Nothing
End Sub
Sub ShtAddThree()
    Dim Sht As Worksheet
    For Each Sht In Worksheets
        If Sht.Name = "数据" Then
            If MsgBox("工作簿中已有""数据""工作表,是否删除后添加?", 36) = 6 Then
                Application.DisplayAlerts = False
                Sht.Delete
                Application.DisplayAlerts = True
            Else
                Exit Sub
            End If
        End If
    Next
    Worksheets.Add.Name = "数据"
    Set Sht = Nothing
End Sub
Sub ShtAddFour()
    Dim arr As Variant
    Dim i As Integer
    Dim Sht As Worksheet
    On Error Resume Next
    arr = Array(1, 2, 3, 4, 5, 6)
    With Worksheets
        For i = 0 To UBound(arr)
            Set Sht = .Add(after:=Worksheets(.Count))
            Sht.Name = arr(i)
        Next
    End With
    Application.DisplayAlerts = False
    For Each Sht In Worksheets
        If Sht.Name Like "Sheet*" Then Sht.Delete
    Next
    Application.DisplayAlerts = True
    Set Sht = Nothing
End Sub

' 范例22 禁止删除指定工作表
Option Explicit
Private Sub Workbook_Activate()
    Application.CommandBars.FindControl(ID:=847).OnAction = "MyDelSht"
End Sub

Option Explicit
Sub MyDelSht()
    If ActiveSheet.CodeName = "Sheet2" Then 'codename 真实的代码名称
        MsgBox ActiveSheet.Name & "工作表禁止删除!", 48
    Else
        ActiveSheet.Delete
    End If
End Sub

Private Sub Workbook_Deactivate()
    Application.CommandBars.FindControl(ID:=847).OnAction = ""
End Sub

' 范例23 禁止更改工作表名称
Option Explicit
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Sheet1.Name <> "Important" Then Sheet1.Name = "Important"
    ThisWorkbook.Save
End Sub

' 范例24 判断是否存在指定工作表
Option Explicit
Sub ShtExists()
    Dim Sht As Worksheet
    On Error GoTo line
    ' 使用Set语句将工作表对象赋予变量Sht，如果工作簿中没有“abc”工作表，此行代码将发生运行错误。
    Set Sht = Worksheets("abc")
    MsgBox "工作簿中已有""abc""工作表!"
    Exit Sub
line:
    MsgBox "工作簿中没有""abc""工作表!"
End Sub

' 范例25 工作表的深度隐藏
Option Explicit
Public sht As Worksheet
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Sheet1.Visible = True
    For Each sht In ThisWorkbook.Sheets
        If sht.CodeName <> "Sheet1" Then
            sht.Visible = xlSheetVeryHidden
        End If
    Next
    ThisWorkbook.Save
End Sub
Private Sub Workbook_Open()
    For Each sht In ThisWorkbook.Sheets
        If sht.CodeName <> "Sheet1" Then
            sht.Visible = xlSheetVisible
        End If
    Next
    Sheet1.Visible = xlSheetVeryHidden
End Sub

' 范例26 工作表的保护与取消保护
Option Explicit
Sub ShProtect()
    With Sheet1
        .Unprotect Password:="123"
        .Cells(1, 1) = .Cells(1, 1) + 100
        .Protect Password:="123"
    End With
End Sub
Sub RemoveShProtect()
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    Dim i7 As Integer, i8 As Integer, i9 As Integer
    Dim i10 As Integer, i11 As Integer, i12 As Integer
    Dim t As String
    On Error Resume Next ' 忽略错误
    If ActiveSheet.ProtectContents = False Then
        MsgBox "该工作表没有保护密码!"
        Exit Sub
    End If
    t = Timer
    For i1 = 65 To 66: For i2 = 65 To 66: For i3 = 65 To 66
    For i4 = 65 To 66: For i5 = 65 To 66: For i6 = 65 To 66
    For i7 = 65 To 66: For i8 = 65 To 66: For i9 = 65 To 66
    For i10 = 65 To 66: For i11 = 65 To 66: For i12 = 32 To 126
        ActiveSheet.Unprotect Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) _
        & Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(i10) & Chr(i11) & Chr(i12)
        If ActiveSheet.ProtectContents = False Then
            MsgBox "解除工作表保护!用时" & Format(Timer - t, "0.00") & "秒"
            Exit Sub
        End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Sub

' 范例27 自动建立工作表目录
Option Explicit
Private Sub Worksheet_Activate()
    Dim Sht As Worksheet
    Dim a As Integer
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    a = 2 '第一行是“目录”
    If r > 1 Then Range("A2:A" & r).ClearContents
    For Each Sht In Worksheets
        If Sht.CodeName <> "Sheet1" Then
            Cells(a, 1).Value = Sht.Name
            a = a + 1
        End If
    Next
    Set Sht = Nothing
End Sub
' 使用工作表的SelectionChange事件，建立到各工作表的链接
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    On Error Resume Next
    If Not Application.Intersect(Target, Range("A2:A" & r)) Is Nothing Then
        Sheets(Target.Text).Select '选到哪个target就select到对应的表上
    End If
End Sub

' 范例28 循环选择工作表
Option Explicit
Sub ShtNext()
    If ActiveSheet.Index < Worksheets.Count Then
        ActiveSheet.Next.Activate
    Else
        Worksheets(1).Activate
    End If
End Sub
Sub ShtPrevious()
    If ActiveSheet.Index > 1 Then
        ActiveSheet.Previous.Activate
    Else
        Worksheets(Worksheets.Count).Activate
    End If
End Sub

' 范例29 在工作表中一次插入多行
Option Explicit
Sub InSertRow()
    Dim i As Integer
    '如果需要一次插入多行，可以使用For...Next语句重复插入
    For i = 1 To 3
        Sheet1.Rows(3).Insert
    Next
    'Sheet1.Range("A3").EntireRow.Resize(3).Insert '用resize引用三行
    'Sheet1.Rows(3).Resize(3).Insert
End Sub

' 范例30 删除工作表中的空行
Option Explicit
Sub DelBlankRow()
    Dim r As Long
    Dim i As Long
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet.usedrange
    r = Sheet1.UsedRange.Rows.Count
    For i = r To 1 Step -1
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.range.find
        If Rows(i).Find("*", , xlValues, , , 2) Is Nothing Then
            Rows(i).Delete
        End If
    Next
End Sub

' 范例31 删除工作表的重复行
Option Explicit
Sub DeleteRow()
    Dim r As Integer
    Dim i As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = r To 1 Step -1
            If WorksheetFunction.CountIf(.Columns(1), .Cells(i, 1)) > 1 Then
                .Rows(i).Delete '遇到重复的删掉，往上搜索后还有重复的会被继续删掉
            End If
        Next
    End With
End Sub

' 范例32 定位删除特定内容所在的行
Option Explicit
Sub SpecialDelete()
    Dim r As Long
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A2:A" & r).Replace "VT248PA", "", 2 '将所有的值替代为空白
        ' https://docs.microsoft.com/en-us/office/vba/api/excel.xlcelltype
        .Columns(1).SpecialCells(4).EntireRow.Delete '再把空行删了
    End With
End Sub

' 范例33 判断是否选中整行
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Rows.Count = 1 Then
        If Target.Columns.Count = 16384 Then
            MsgBox "您选中了整行,当前行号" & Target.Row
        End If
    End If
End Sub

' 范例34 限制工作表的滚动区域
Option Explicit
Private Sub Workbook_Open()
    Sheet1.ScrollArea = "B4:H12"
End Sub

' 范例35 复制自动筛选后的数据区域
Option Explicit
Sub CopyFilter()
    Sheet2.Cells.Clear
    With Sheet1
        If .FilterMode Then
            .AutoFilter.Range.SpecialCells(12).Copy Sheet2.Cells(1, 1)
        End If
    End With
End Sub

' 范例36 使用高级筛选功能获得不重复记录
Option Explicit
Sub Filter()
' Filters or copies data from a list based on a criteria range. 
' If the initial selection is a single cell, that cell's current region is used.
    Sheet1.Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        Unique:=True, _
        CopyToRange:=Sheet2.Range("A1")
End Sub

' 范例37 获得工作表打印页数
Option Explicit
Sub PrintPage()
    Dim Page As Integer
    ' Excel4 Macro
    ' These types of Macros were superseded when VBA was introduced in Excel version 5, 
    ' hence why any macros before that are referred to as Excel 4 Macros.
    ' https://exceloffthegrid.com/using-excel-4-macro-functions/
    Page = ExecuteExcel4Macro("GET.DOCUMENT(50)")
    MsgBox "工作表打印页数共" & Page & "页!"
End Sub
