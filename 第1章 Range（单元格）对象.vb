' 范例1 单元格的引用方法

' 1-1 使用Range属性引用单元格区域
' 主要需要学习Range的取法
' https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.range
Option Explicit
Sub MyRng()
    Range("A1:B4, D5:E8").Select '合并
    Range("A1").Formula = "=Rand()"
    Range("A1:B4 B2:C6").Value = 10 '交叉
    Range("A1", "B4").Font.Italic = True
End Sub

' 1-2 使用Cells属性引用单元格区域
' Cells按照行列取一个格
' https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.cells
Option Explicit
Sub MyCell()
    Dim i As Byte
    For i = 1 To 10
        Sheets("Sheet1").Cells(i, 1).Value = i
    Next
End Sub

' 1-3 使用快捷记号实现快速输入
' 直接赋值
Option Explicit
Sub FastMark()
    [A1] = "Excel 2007"
End Sub

' 1-4 使用Offset属性返回单元格区域
Option Explicit
Sub RngOffset()
    Sheets("Sheet1").Range("A1:B2").Offset(2, 2).Select
End Sub

' 1-5 使用Resize属性返回调整后的单元格区域
' https://docs.microsoft.com/zh-cn/office/vba/api/excel.range.resize
Option Explicit
Sub RngResize()
    Sheets("Sheet1").Range("A1").Resize(4, 4).Select
End Sub

' 范例2 选定单元格区域的方法

' 2-1 使用Select方法选定单元格区域
Option Explicit
Sub RngSelect()
    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("A1:B10").Select
End Sub

' 2-2 使用Activate方法选定单元格区域
Option Explicit
Sub RngActivate()
    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("A1:B10").Activate
End Sub

' 2-3 使用Goto方法选定单元格区域
' https://docs.microsoft.com/zh-cn/office/vba/api/excel.application.goto
Option Explicit
Sub RngGoto()
    Application.Goto Reference := Sheets("Sheet2").Range("A1：B10"), Scroll:=True '滚动到此处
End Sub

' 范例3 获得指定行的最后一个非空单元格
Option Explicit
Sub LastCell()
    Dim rng As Range
    Set rng = Cells(Rows.Count, 1).End(xlUp) 
    ' Rows.Count 先定位到最后一行，xlUp向上遍历 xlDown xlToLeft xlToRight
    ' https://stackoverflow.com/a/27067705/5063930
    ' The End function starts at a cell and then, 
    ' depending on the direction you tell it, goes that direction until 
    ' it reaches the edge of a group of cells that have text.
    MsgBox "A列的最后一个非空单元格是" & rng.Address(0, 0)_ & ",行号" & rng.Row & ",数值" & rng.Value
    Set rng = Nothing
End Sub

' 范例4 使用SpecialCells方法定位单元格
' 返回含有特殊格式的单元格，比如含有公式等
Sub SpecialAddress()
    Dim rng As Range
    Set rng = Sheet1.UsedRange.SpecialCells(xlCellTypeFormulas)'含有公式的单元格
    'UsedRange 所有用过的单元格
    rng.Select
    MsgBox "工作表中有公式的单元格为： " & rng.Address
    Set rng = Nothing
End Sub

' 范例5 查找特定内容的单元格

' 5-1 使用Find方法查找特定信息
' https://docs.microsoft.com/zh-cn/office/vba/api/excel.range.find
Option Explicit
Sub FindCell()
    Dim StrFind As String
    Dim rng As Range
    StrFind = InputBox("请输入要查找的值：")
    If Len(Trim(StrFind)) > 0 Then
        With Sheet1.Range("A:A")
            Set rng = .Find(What:=StrFind, _
                After:=.Cells(.Cells.Count), _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False)
            If Not rng Is Nothing Then
                Application.Goto rng, True
            Else
                MsgBox "没有找到匹配单元格!"
            End If
        End With
    End If
    Set rng = Nothing
End Sub

' 延伸1：With 语句
' https://docs.microsoft.com/zh-cn/office/vba/language/reference/user-interface-help/with-statement
' The With statement allows you to perform a series of statements on a specified object without 
' requalifying the name of the object.
With MyLabel 
 .Height = 2000 
 .Width = 2000 
 .Caption = "This is MyLabel" 
End With

' 延伸2：Understanding named arguments and optional arguments
' https://docs.microsoft.com/en-us/office/vba/Language/Concepts/Getting-Started/understanding-named-arguments-and-optional-arguments
PassArgs "Mary", 29, #2-21-69#
PassArgs intAge:=29, dteBirth:=#2/21/69#, strName:="Mary"

' https://docs.microsoft.com/zh-cn/office/vba/api/excel.range.findnext
Sub FindNextCell()
    Dim StrFind As String
    Dim rng As Range
    Dim FindAddress As String
    StrFind = InputBox("请输入要查找的值：")
    If Len(Trim(StrFind)) > 0 Then
        With Sheet1.Range("A:A")' 第A列
            .Interior.ColorIndex = 0
            Set rng = .Find(What:=StrFind, _ '_ 是换行符
                After: = .Cells(.Cells.Count), _ '定位到最后一格，然后返回到第一格开始搜索
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, _
                MatchCase:=False)
            If Not rng Is Nothing Then
                FindAddress = rng.Address
                Do
                    rng.Interior.ColorIndex = 6 '设成黄色
                    Set rng = .FindNext(rng)
                Loop While Not rng Is Nothing _
                    And rng.Address <> FindAddress
            End If
        End With
    End If
    Set rng = Nothing
End Sub

' 5-2 使用Like运算符进行模式匹配查找
' 正则模糊匹配
Option Explicit
Sub RngLike()
    Dim rng As Range
    Dim r As Integer
    r = 1
    Sheet1.Range("A:A").ClearContents '清理区域中的公式和值
    For Each rng In Sheet2.Range("A1:A40")
        If rng.Text Like "*a*" Then
            Cells(r, 1) = rng.Text
            r = r + 1
        End If
    Next
    Set rng = Nothing
End Sub

' 范例6 替换单元格内字符串
Option Explicit
Sub Replacement()
    Range("A:A").Replace _
        What: = "市", Replacement: = "区", _
        LookAt: = xlPart, SearchOrder: = xlByRows, _
        MatchCase: = True
End Sub

' 范例7 复制单元格
Option Explicit
Sub RangeCopy()
    Sheet1.Range("A1:G7").Copy _
    destination := Sheet2.Range("A1") '这是什么怪异的语法，为什么没有括号
End Sub
Sub Copyalltheforms()
    Dim i As Integer
    Sheet1.Range("A1:G7").Copy
    With Sheet3.Range("A1")
        .PasteSpecial(xlPasteAll) '相当于sheet3.range("A1").pastespecial(xlpasteall)
        .PasteSpecial(xlPasteColumnWidths)
    End With
    Application.CutCopyMode = False
    For i = 1 To 7
        Sheet3.Rows(i).RowHeight = Sheet1.Rows(i).RowHeight
    Next
End Sub

'延伸1：PasteSpecial
' https://docs.microsoft.com/en-us/office/vba/api/excel.range.pastespecial
' https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastetype
' https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastespecialoperation

' 7-2 仅复制数值到另一区域
Option Explicit
Sub CopyValue()
    Sheet1.Range("A1:G7").Copy
    Sheet2.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub
Sub GetValueResize()
    With Sheet1.Range("A1").CurrentRegion '选取整个区域
         Sheet3.Range("A1").Resize(.Rows.Count, .Columns.Count).Value = .Value '用resize拓展出相同大小的区域
    End With
End Sub

' 范例8 禁用单元格拖放功能
' 好像没啥用？
Option Explicit
Private Sub Worksheet_Deactivate()
    Application.CellDragAndDrop = True
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column = 1 Then
        Application.CellDragAndDrop = False
    Else
        Application.CellDragAndDrop = True
    End If
End Sub

' 9-1 设置单元格字体格式
' https://docs.microsoft.com/en-us/office/vba/api/excel.font(object)
Option Explicit
Sub CellFont()
    With Range("A1").Font 'font object
        .Name = "华文彩云"
        .FontStyle = "Bold"
        .Size = 22
        .ColorIndex = 3
        .Underline = 2
    End With
End Sub

' 9-2 设置单元格内部格式
' https://docs.microsoft.com/en-us/office/vba/api/excel.interior(object)
Option Explicit
Sub CellInternalFormat()
    With Range("A1").Interior 'interior object
        .ColorIndex = 3
        .Pattern = xlPatternGrid
        .PatternColorIndex = 6
    End With
End Sub

' 9-3 为单元格区域添加边框
Option Explicit
Sub CellBorder()
     Dim rng As Range
     Set rng = Range("B2:E8")
     With rng.Borders(xlInsideHorizontal) 'borders object 里面的参数代表是哪条边
     ' https://docs.microsoft.com/en-us/office/vba/api/excel.xlbordersindex
         .LineStyle = xlDot
         .Weight = xlThin
         .ColorIndex = xlColorIndexAutomatic
     End With
     With rng.Borders(xlInsideVertical)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlColorIndexAutomatic
     End With
     rng.BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
     Set rng = Nothing
End Sub
Sub QuickBorder()
    Range("B12:E18").Borders.LineStyle = xlContinuous
End Sub

' 范例10 单元格的数据有效性

' 10-1 添加数据有效性
' https://docs.microsoft.com/en-us/office/vba/api/excel.validation
' https://docs.microsoft.com/en-us/office/vba/api/excel.xldvtype  
Option Explicit
Sub AddValidation()
    With Range("A1:A10").Validation 'Use the Validation property of the Range object to return the Validation object.
        ' Methods Add Delete Modify
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="1,2,3,4,5,6,7,8"
        .ErrorMessage = "只能输入1-8的数值,请重新输入!"
    End With
End Sub

' 10-2 判断是否存在数据有效性
Option Explicit
Sub ErrValidation()
    On Error GoTo Line
    If Range("A1").Validation.Type >= 0 Then
        MsgBox "有数据有效性!"
        Exit Sub
    End If
Line:
    MsgBox "没有数据有效性!"
End Sub

' 10-3 动态的数据有效性
' 之前有做过，根据上一条件决定下一条件的下拉菜单
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column = 1 And Target.Count = 1 And Target.Row > 1 Then
        With Target.Validation
            .Delete
            .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="主机,显示器"
        End With
    End If
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 1 And Target.Row > 1 And Target.Count = 1 Then
        With Target.Offset(0, 1).Validation
            .Delete
            Select Case Target
                Case "主机"
                    .Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:="Z286,Z386,Z486,Z586"
                Case "显示器"
                    .Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, _
                        Formula1:="15,17,21,25"
            End Select
        End With
    End If
End Sub

' 范例11 单元格中的公式

' 11-1 在单元格中写入公式
Option Explicit
Sub rngFormula()
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row '最后一行
    Range("C2").Formula = "=A2*B2"
    Range("C2").Copy Range("C3:C" & r) '向下填充
    Range("A" & r + 1) = "合计"
    Range("C" & r + 1).Formula = "=SUM(C2:C" & r & ")"
End Sub

Sub rngFormulaRC()
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C2:C" & r).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("A" & r + 1) = "合计"
    Range("C" & r + 1).FormulaR1C1 = "=SUM(R[-" & r - 1 & "]C:R[-1]C)"
End Sub

Sub RngFormulaArray()
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    Range("C2:C" & r).FormulaR1C1 = "=RC[-2]*RC[-1]"
    Range("A" & r + 1) = "合计"
    Range("C" & r + 1).FormulaArray = "=SUM(R[-" & r - 1 & "]C[-2]:R[-1]C[-2]*R[-" & r - 1 & "]C[-1]:R[-1]C[-1])"
    'Range("C" & r + 1).FormulaArray = "=SUM(A2:A" & r & "*B2:B" & r & ")"
End Sub

' 11-2 判断单元格是否包含公式
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/select-case-statement
' https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-select-case-statements
Option Explicit
Sub rngIsHasFormula()
    Select Case Selection.HasFormula
        Case True
            MsgBox "单元格包含公式!"
        Case False
            MsgBox "单元格没有公式!"
        Case Else
            MsgBox "公式区域：" & Selection.SpecialCells(-4123, 23).Address(0, 0)
    End Select
End Sub

' 11-3 判断单元格公式是否存在错误
Option Explicit
Sub CellFormulaIsWrong()
    If IsError(Range("A1").Value) = True Then
        MsgBox "A1单元格错误类型为:" & Range("A1").Text
    Else
        MsgBox "A1单元格公式结果为" & Range("A1").Value
    End If
End Sub
'MsgBox "Text:" & Range("A1").Text & vbCrLf & "Value:" & Range("A1").Value
    
' 11-4 取得公式的引用单元格
Option Explicit
Sub RngPrecedent()
    Dim rng As Range
    Set rng = Sheet1.Range("C10").Precedents 'precedents property
    MsgBox "公式所引用的单元格是：" & rng.Address
    Set rng = Nothing
End Sub

' 11-5 将公式转换为数值
Option Explicit
Sub SpecialPaste()
    With Range("A1:A10")
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With
    Application.CutCopyMode = False
    ' 延伸1：CutCopyMode
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.application.cutcopymode
End Sub

'Range("A1:A10").Value = Range("A1:A10").Value

' 范例12 为单元格添加批注
Option Explicit
Sub AddComment()
    With Range("A1")
        If Not .Comment Is Nothing Then .Comment.Delete
        .AddComment Text:=Date & vbCrLf & .Text
        ' vbCrLf 回车加换行
        .Comment.Visible = True
    End With
End Sub

' 范例13 合并单元格操作

' 13-1 判断单元格区域是否存在合并单元格
Option Explicit
Sub IsMergeCell()
    If Range("A1").MergeCells Then 'True if the range contains merged cells
        MsgBox "合并单元格", vbInformation '控制弹出的窗格是什么类型的
    Else
        MsgBox "非合并单元格", vbInformation
    End If
End Sub
Sub IsMergeCells()
    If IsNull(Range("A1:D10").MergeCells) Then
        MsgBox "包含合并单元格", vbInformation
    Else
        MsgBox "没有包含合并单元格", vbInformation
    End If
End Sub

' 13-2 合并单元格时连接每个单元格的文本
Option Explicit
Sub MergeCells()
    Dim MergeStr As String
    Dim MergeRng As Range
    Dim rng As Range
    Set MergeRng = Range("A1:B2")
    For Each rng In MergeRng
        MergeStr = MergeStr & rng & " "
    Next
    Application.DisplayAlerts = False
    MergeRng.Merge
    MergeRng.Value = MergeStr
    Application.DisplayAlerts = True
    Set MergeRng = Nothing
    Set rng = Nothing
End Sub
'MergeRng.MergeCells = True

' 13-3 合并内容相同的连续单元格
Sub MergeLinkedCell()
    Dim r As Integer
    Dim i As Integer
    Application.DisplayAlerts = False
    With Sheet1
        r = .Cells(Rows.Count, 1).End(xlUp).Row
        For i = r To 2 Step -1
            If .Cells(i, 2).Value = .Cells(i - 1, 2).Value Then
                .Range(.Cells(i - 1, 2), .Cells(i, 2)).Merge
            End If
        Next
    End With
    Application.DisplayAlerts = True
End Sub

' 13-4 取消合并单元格时在每个单元格中保留内容
Option Explicit
Sub CancelMergeCells()
    Dim r As Integer
    Dim MergeStr As String
    Dim MergeCot As Integer
    Dim i As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 2 To r
            MergeStr = .Cells(i, 2).Value
            MergeCot = .Cells(i, 2).MergeArea.Count
            .Cells(i, 2).UnMerge
            .Range(.Cells(i, 2), .Cells(i + MergeCot - 1, 2)).Value = MergeStr
            i = i + MergeCot - 1
        Next
        .Range("B1:B" & r).Borders.LineStyle = xlContinuous
    End With
End Sub

' 范例14 高亮显示选定单元格区域
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells.Interior.ColorIndex = xlColorIndexNone
    Target.Interior.ColorIndex = Int(56 * Rnd() + 1)
End Sub

Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
' It means that a variable named Target is passed in to the
' procedure. The Target variable is a Range type variable, meaning
' that it refers to (points to) a cell or range of cells.
    Dim rng As Range
    Cells.Interior.ColorIndex = xlColorIndexNone
    Set rng = Application.Union(Target.EntireColumn, Target.EntireRow)
    rng.Interior.ColorIndex = Int(56 * Rnd() + 1)
    Set rng = Nothing
End Sub

' 范例15 双击被保护单元格时不弹出提示消息框
Option Explicit
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Locked = True Then
        MsgBox "此单元格已保护，不能编辑!"
        Cancel = True
    End If
End Sub

'延伸：保护工作表
' https://support.office.com/en-us/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6?ui=en-US&rs=en-US&ad=US
' ctrl+1 保护某些格，然后选择“保护工作表选项”

' 范例16 单元格录入数据后的自动保护
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim msg As Byte
    With Target
        If Not Application.Intersect(Target, Range("A2:F6")) Is Nothing Then
            If .Count > 1 Then
                Range("A1").Select
                Exit Sub
            End If
            ActiveSheet.Unprotect
            If Len(Trim(.Value)) > 0 Then
                msg = MsgBox("当前单元格已录入数据,是否修改?", 32 + 4)
                .Locked = IIf(msg = 6, False, True)
            End If
            ActiveSheet.Protect
            ActiveSheet.EnableSelection = 0
        End If
    End With
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Application.Intersect(Target, Range("A2:F6")) Is Nothing Then
        If Len(Trim(Target.Value)) > 0 Then
            ActiveSheet.Unprotect
            Target.Locked = True
            ActiveSheet.Protect
            ActiveSheet.EnableSelection = 0
        End If
    End If
End Sub

' 范例17 Target参数的使用方法

' 17-1 使用Address 属性17-1 使用Address 属性
Option Explicit
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'https://docs.microsoft.com/en-us/office/vba/api/excel.range.address
    Select Case Target.Address(0, 0)
        Case "A1"
            Sh.Unprotect
        Case "A2"
            Sh.Protect
        Case Else
    End Select
End Sub

'17-2 使用Column属性和Row属性
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column < 3 And Target.Row < 11 Then
        MsgBox "你选择了" & Target.Address(RowAbsolute := 0, ColumnAbsolute := 0) & "单元格" '相对引用值
    End If
End Sub

' 17-3 使用Intersect属性
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not Application.Intersect(Target, Union(Range("A1:B10"), Range("E1:F10"))) Is Nothing Then 
        If Target.Count = 1 Then '如果重叠，Intersect方法返回一个Range对象
            MsgBox "你选择了" & Target.Address(0, 0) & "单元格"
        End If
    End If
End Sub






