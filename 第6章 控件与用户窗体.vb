' 范例67 使用文本框输入数值
Option Explicit
Private Sub CommandButton1_Click()
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    If Len(Trim(TextBox1.Text)) > 0 Then
        Cells(r + 1, 1) = Round(TextBox1.Text, 2)
        TextBox1.Text = ""
        TextBox1.SetFocus
    End If
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)
    With TextBox1
        Select Case KeyANSI
            Case Asc("0") To Asc("9")
            Case Asc("-")
                If InStr(1, .Text, "-") > 0 Or .SelStart > 0 Then
                    KeyANSI = 0
                End If
            Case Asc(".")
                If InStr(1, .Text, ".") > 0 Then KeyANSI = 0
            Case Else
                KeyANSI = 0
        End Select
    End With
End Sub
Private Sub TextBox1_Change()
    Dim i As Integer
    Dim Str As String
    With TextBox1
        For i = 1 To Len(.Text)
            Str = Mid(.Text, i, 1)
            Select Case Str
                Case ".", "-", "0" To "9"
                Case Else
                    .Text = Replace(.Text, Str, "")
            End Select
        Next
    End With
End Sub

' 范例68 限制文本框的输入长度
Option Explicit
Private Sub TextBox1_Change()
    TextBox1.MaxLength = 6
End Sub

' 范例69 验证文本框输入的数据
Option Explicit
Private Sub CommandButton1_Click()
    With TextBox1
        If (Len(Trim(.Text))) = 15 Or (Len(Trim(.Text))) = 18 Then
            Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = .Text '输入到下一格
        Else
            MsgBox "身份证号码错误，请重新输入!"
        End If
        .Text = ""
        .SetFocus
    End With
End Sub

' 范例70 文本框回车后自动输入数据
Option Explicit
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim r As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    With TextBox1
        If Len(Trim(.Text)) > 0 And KeyCode = vbKeyReturn Then '回车键
            Cells(r + 1, 1) = .Text
            .Text = ""
        End If
    End With
End Sub

' 范例71 文本框的自动换行
Option Explicit
Private Sub UserForm_Initialize()
    With TextBox1
        .WordWrap = True '反正记住这两行开起来
        .MultiLine = True
        .Text = "文本框是一个灵活的控件，受下列属性的影响：Text、" _
                & "MultiLine、WordWrap和AutoSize。" & vbCrLf _
                & "Text 包含显示在文本框中的文本。" & vbCrLf _
                & "MultiLine 控制文本框是单行还是多行显示文本。" _
                & "换行字符用于标识在何处结束一行并开始新的一行。" _
                & "如果 MultiLine 的值为False，则文本将被截断，" _
                & "而不会换行。如果文本的长度大于文本框的宽度，" _
                & "WordWrap允许文本框根据其宽度自动换行。" & vbCrLf _
                & "如果不使用 WordWrap，当文本框在文本中遇到换行字符时，" _
                & "开始一个新行。如果关闭WordWrap，TextBox中可以有不能" _
                & "完全适合其宽度的文本行。文本框根据该宽度，显示宽度以" _
                & "内的文本部分，截断宽度以外的那文本部分。只有当" _
                & "MultiLine为True时，WordWrap才起作用。" & vbCrLf _
                & "AutoSize 控制是否调节文本框的大小，以便显示所有文本。" _
                & "当文本框使用AutoSize 时，文本框的宽度按照文本框中的" _
                & "文字量以及显示该文本的字体大小收缩或扩大。"
    End With
End Sub

' 范例72 格式化文本框数据
Option Explicit
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox1 = Format(TextBox1, "##,#0.00")
End Sub
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox2 = Format(TextBox2, "##,#0.00")
End Sub

' 范例73 使控件始终位于可视区域
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rng As Range
    Set rng = ActiveWindow.VisibleRange.Cells(1)
    With CommandButton1
        .Top = rng.Top
        .Left = rng.Left
    End With
    With CommandButton2
        .Top = rng.Top
        .Left = rng.Left + CommandButton1.Width
    End With
    Set rng = Nothing
End Sub

' 范例74 高亮显示按钮控件
Option Explicit
Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.CommandButton1
        .BackColor = &HFFFF00
        .Width = 62
        .Height = 62
        .Top = 69
        .Left = 31
    End With
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With Me.CommandButton1
        .BackColor = Me.BackColor
        .Width = 60
        .Height = 60
        .Top = 70
        .Left = 32
    End With
End Sub

' 范例75 为列表框添加列表项的方法

' 75-1 使用RowSource属性添加列表项
Option Explicit
Private Sub UserForm_Initialize()
    Dim r As Integer
    r = Sheet3.Range("A1048576").End(xlUp).Row
    ListBox1.RowSource = "Sheet3!a1:a" & r
End Sub
    'ListBox1.RowSource = Sheet3.Range("A1:A" & r).Address(External:=True)
    'ListBox1.RowSource = "规格"

' 75-2 使用ListFillRange属性添加列表项
Option Explicit
Sub ListFillRange()
    Dim r As Integer
    r = Sheet3.Range("A1048576").End(xlUp).Row
    Sheet1.ListBox1.ListFillRange = "Sheet3!a1:a" & r
    Sheet1.Shapes("列表框").ControlFormat.ListFillRange = "Sheet3!a1:a" & r
End Sub

' 75-3 使用List属性添加列表项
Option Explicit
Private Sub UserForm_Initialize()
    Dim arr As Variant
    Dim r As Integer
    r = Sheet3.Range("A1048576").End(xlUp).Row
    arr = Sheet3.Range("A1:A" & r)
    ListBox1.List = arr
End Sub
    'ListBox1.List = Range("规格").Value

' 75-4 使用AddItem方法添加列表项
Option Explicit
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim i As Integer
    r = Sheet3.Range("A1048576").End(xlUp).Row
    For i = 1 To r
        ListBox1.AddItem (Sheet3.Cells(i, 1))
    Next
End Sub

' 范例76 去除列表项的空行和重复项
Option Explicit
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim i As Integer
    Dim MyCol As New Collection
    Dim arr() As Variant
    On Error Resume Next
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 1 To r
            If Trim(.Cells(i, 1)) <> "" Then
                MyCol.Add Item:=Cells(i, 1), key:=CStr(.Cells(i, 1))
            End If
        Next
    End With
    ReDim arr(1 To MyCol.Count)
    For i = 1 To MyCol.Count
        arr(i) = MyCol(i)
    Next
    ListBox1.List = arr
End Sub

' 范例77 移动列表框的列表项
Option Explicit
Private Sub CommandButton1_Click()
    Dim Ind As Integer
    Dim Str As String
    With Me.ListBox1
        Ind = .ListIndex
        Select Case Ind
            Case -1
                MsgBox "请选择一行后再移动!"
            Case 0
                MsgBox "已经是第一行了!"
            Case Is > 0
                Str = .List(Ind)
                .List(Ind) = .List(Ind - 1)
                .List(Ind - 1) = Str
                .ListIndex = Ind - 1
        End Select
    End With
End Sub
Private Sub CommandButton2_Click()
    Dim Ind As Integer
    Dim Str As String
    With ListBox1
        Ind = .ListIndex
        Select Case Ind
            Case -1
                MsgBox "请选择一行后再移动!"
            Case .ListCount - 1
                MsgBox "已经是最后一行了!"
            Case Is < .ListCount - 1
                Str = .List(Ind)
                .List(Ind) = .List(Ind + 1)
                .List(Ind + 1) = Str
                .ListIndex = Ind + 1
        End Select
    End With
End Sub
Private Sub CommandButton3_Click()
    Dim i As Integer
    For i = 1 To ListBox1.ListCount
        Cells(i, 1) = ListBox1.List(i - 1)
    Next
End Sub
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim arr As Variant
    r = Cells(Rows.Count, 1).End(xlUp).Row
    arr = Range("A1:A" & r)
    ListBox1.List = arr
End Sub

' 范例78 允许多项选择的列表框
Option Explicit
Private Sub UserForm_Initialize()
    Dim arr As Variant
    arr = Array("经理室", "办公室", "生技科", "财务科", "营业部", "制水车间", "污水厂", "其他")
    With Me.ListBox1
        .List = arr
        .MultiSelect = 1 '允许多项选择
        .ListStyle = 1
    End With
End Sub
Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim Str As String
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            Str = Str & ListBox1.List(i) & Chr(13)
        End If
    Next
    If Str <> "" Then
        MsgBox Str
    Else
        MsgBox "至少需要选择一个部门!"
    End If
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub

' 范例79 多列列表框的设置
Option Explicit
Private Sub UserForm_Initialize()
    Dim r As Integer
    With Sheet3
        r = .Cells(.Rows.Count, 1).End(xlUp).Row - 1
    End With
    With ListBox1
        .ColumnCount = 7
        .ColumnWidths = "35,45,45,45,45,40,50"
        .BoundColumn = 1
        .ColumnHeads = True
        .TextAlign = 3
        .RowSource = Sheet3.Range("A2:G" & r).Address(External:=True)
    End With
End Sub
Private Sub ListBox1_Click()
    Dim r As Integer
    Dim i As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        For i = 1 To ListBox1.ColumnCount
            .Cells(r, i) = ListBox1.Column(i - 1)
        Next
    End With
End Sub

' 范例80 加载二级组合框
Option Explicit 
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim MyCol As New Collection
    Dim arr() As Variant
    Dim rng As Range
    Dim i As Integer
    On Error Resume Next
    r = Cells(Rows.Count, 1).End(xlUp).Row
    For Each rng In Range("A2:A" & r)
        MyCol.Add rng, CStr(rng)
    Next
    ReDim arr(1 To MyCol.Count)
    For i = 1 To MyCol.Count
        arr(i) = MyCol(i)
    Next
    ComboBox1.List = arr
    ComboBox1.ListIndex = 0
    Set MyCol = Nothing
    Set rng = Nothing
End Sub

Private Sub ComboBox1_Change()
    Dim MyAddress As String
    Dim rng As Range
    ComboBox2.Clear
    With Sheet1.Range("A:A")
        Set rng = .Find(What:=ComboBox1.Text) '使用Find 方法将所有属于ComboBox1所选省份的市县名称加载到ComboBox2中
        If Not rng Is Nothing Then
            MyAddress = rng.Address
            Do
                ComboBox2.AddItem rng.Offset(, 1)
                Set rng = .FindNext(rng)
            Loop While Not rng Is Nothing And rng.Address <> MyAddress
        End If
    End With
    ComboBox2.ListIndex = 0
    Set rng = Nothing
End Sub

' 范例81 使用RefEdit控件选择区域
Option Explicit
Private Sub CommandButton1_Click()
    Dim rng As Range
    On Error Resume Next
    Set rng = Range(RefEdit1.Value)
    rng.Interior.ColorIndex = 16
    Set rng = Nothing
End Sub

' 范例82 使用多页控件
Option Explicit
Private Sub UserForm_Initialize()
    MultiPage1.Value = 0
End Sub
Private Sub MultiPage1_Change()
    If MultiPage1.SelectedItem.Index > 0 Then
        MsgBox "您选择的是" & MultiPage1.SelectedItem.Caption & "页面!"
    End If
End Sub

' 范例83 使用TabStrip控件
Option Explicit
Private Sub TabStrip1_Change()
    Dim str As String
    Dim FilPath As String
    str = TabStrip1.SelectedItem.Caption
    FilPath = ThisWorkbook.Path & "\" & str & ".jpg"
    Image1.Picture = LoadPicture(FilPath)
    Label1.Caption = str & "欢迎您!"
End Sub
Private Sub UserForm_Initialize()
    TabStrip1.Value = 0
    TabStrip1.Style = 0
End Sub

' 范例84 在框架中使用滚动条
Option Explicit
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    With Frame1
        .ScrollBars = 3 '显示横竖的滚动条
        .ScrollHeight = Image1.Height
        .ScrollWidth = Image1.Width
    End With
End Sub

' 范例85 制作进度条
Option Explicit
Sub myProgressBar()
    Dim r As Integer
    Dim i As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        UserForm1.Show 0
        With UserForm1.ProgressBar1
            .Min = 1
            .Max = r
            .Scrolling = 0
        End With
        For i = 1 To r
            .Cells(i, 3) = Round(.Cells(i, 1) * .Cells(i, 2), 2)
            Application.Goto Reference:=.Cells(i, 1), Scroll:=True
            UserForm1.ProgressBar1.Value = i
            UserForm1.Caption = "程序正在运行,已完成" & Format((i / r) * 100, "0.00") & "%,请稍候!"
        Next
    End With
    Unload UserForm1
End Sub

' 范例86 使用DTP控件输入日期
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    With Me.DTPicker1
        If Target.Count = 1 And Target.Column = 1 And Not Target.Row = 1 Or Target.MergeCells Then
            .Visible = True
            .Top = Selection.Top
            .Left = Selection.Left
            .Height = Selection.Height
            .Width = Selection.Width
            If Target.Cells(1, 1) <> "" Then
                .Value = Target.Cells(1, 1).Value
            Else
                .Value = Date
            End If
        Else
            .Visible = False
        End If
    End With
End Sub
Private Sub DTPicker1_CloseUp()
    ActiveCell.Value = Me.DTPicker1.Value
    Me.DTPicker1.Visible = False
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Count = 1 And Target.Column = 1 Or Target.MergeCells Then
        If Target.Cells(1, 1).Value = "" Then
            DTPicker1.Visible = False
        End If
    End If
End Sub

' 范例87 使用spreadsheet控件
Option Explicit
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim arr As Variant
    Dim i As Integer
    With Sheet3
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        arr = .Range("A1:G" & r)
    End With
    With Me.Spreadsheet1
        .DisplayToolbar = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = True
        .Rows.RowHeight = 15
        .Columns.ColumnWidth = 8
        With .Range("A1:G" & r)
            .Value = arr
            .HorizontalAlignment = -4108
            .Borders.LineStyle = xlContinuous
            .Borders.ColorIndex = 10
            .NumberFormat = "0.00"
        End With
    End With
End Sub
Private Sub CommandButton1_Click()
    Dim r As Integer
    Dim arr As Variant
    With Me.Spreadsheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        arr = .Range("A1:G" & r)
        Sheet1.Range("A1:G" & r) = arr
    End With
    Unload Me
End Sub

' 范例88 使用TreeView控件显示层次
Option Explicit
Private Sub TreeView1_DblClick()
    Dim r As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        If TreeView1.SelectedItem.Children = 0 Then
            .Range("A" & r) = TreeView1.SelectedItem.Text
        Else
            MsgBox "您所选择的不是末级科目,请重新选择!"
        End If
    End With
End Sub
Private Sub UserForm_Initialize()
    Dim c As Integer
    Dim r As Integer
    Dim rng As Variant
    rng = Sheet2.UsedRange
    With TreeView1
        .Style = tvwTreelinesPlusMinusPictureText
        .LineStyle = tvwRootLines
        .CheckBoxes = False
        With .Nodes
            .Clear
            .Add Key:="科目", Text:="科目名称" '添加科目
            For c = 1 To Sheet2.UsedRange.Columns.Count
                For r = 2 To Sheet2.UsedRange.Rows.Count
                    If Not IsEmpty(rng(r, c)) Then
                        If c = 1 Then
                            .Add relative:="科目", Relationship:=tvwChild, Key:=rng(r, c), Text:=rng(r, c)
                        ElseIf Not IsEmpty(rng(r, c - 1)) Then
                            .Add relative:=rng(r, c - 1), Relationship:=tvwChild, Key:=rng(r, c), Text:=rng(r, c)
                        Else
                            .Add relative:=CStr(Sheet2.Cells(r, c - 1).End(xlUp)), Relationship:=tvwChild, Key:=rng(r, c), Text:=rng(r, c)
                        End If
                    End If
                Next
            Next
        End With
    End With
End Sub

' 范例89 使用Listview控件

' 89-1 使用Listview控件显示数据列表
Option Explicit
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
End Sub
Private Sub UserForm_Initialize()
    Dim Itm As ListItem
    Dim r As Integer
    Dim i As Integer
    Dim c As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    With ListView1
        .ColumnHeaders.Add , , "人员编号 ", 50, 0
        .ColumnHeaders.Add , , "技能工资 ", 50, 1
        .ColumnHeaders.Add , , "岗位工资 ", 50, 1
        .ColumnHeaders.Add , , "工龄工资 ", 50, 1
        .ColumnHeaders.Add , , "浮动工资 ", 50, 1
        .ColumnHeaders.Add , , "其他  ", 50, 1
        .ColumnHeaders.Add , , "应发合计", 50, 1
        .View = lvwReport
        .Gridlines = True
        For i = 2 To r
            Set Itm = .ListItems.Add()
            Itm.Text = Space(2) & Cells(i, 1)
            For c = 1 To 6
                Itm.SubItems(c) = Format(Cells(i, c + 1), "##,#,0.00")
            Next
        Next
    End With
    Set Itm = Nothing
End Sub

' 89-2 在Listview控件中使用复选框
Option Explicit
Private Sub CommandButton1_Click()
    Dim r As Integer
    Dim i As Integer
    Dim c As Integer
    r = Cells(Rows.Count, 1).End(xlUp).Row
    If r > 1 Then Range("A2:G" & r).ClearContents
    With ListView1
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = .ListItems(i)
                For c = 1 To 6
                    Cells(Rows.Count, c + 1).End(xlUp).Offset(1, 0) = .ListItems(i).SubItems(c)
                Next
            End If
        Next
    End With
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Dim Itm As ListItem
    Dim r As Integer
    Dim i As Integer
    Dim c As Integer
    r = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row
    With ListView1
        .ColumnHeaders.Add , , "人员编号 ", 50, 0
        .ColumnHeaders.Add , , "技能工资 ", 50, 1
        .ColumnHeaders.Add , , "岗位工资 ", 50, 1
        .ColumnHeaders.Add , , "工龄工资 ", 50, 1
        .ColumnHeaders.Add , , "浮动工资 ", 50, 1
        .ColumnHeaders.Add , , "其他  ", 50, 1
        .ColumnHeaders.Add , , "应发合计", 50, 1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .CheckBoxes = True '复选框
        For i = 2 To r - 1
            Set Itm = .ListItems.Add()
            Itm.Text = Sheet2.Cells(i, 1)
            For c = 1 To 6
                Itm.SubItems(c) = Format(Sheet2.Cells(i, c + 1), "##,#,0.00")
            Next
        Next
    End With
    Set Itm = Nothing
End Sub

' 89-3 调整Listview控件的行距
Option Explicit
Private Sub UserForm_Initialize()
    Dim Itm As ListItem
    Dim i As Integer
    Dim c As Integer
    Dim Img As ListImage
    With ListView1
        .ColumnHeaders.Add , , "人员编号 ", 50, 0
        .ColumnHeaders.Add , , "技能工资 ", 50, 1
        .ColumnHeaders.Add , , "岗位工资 ", 50, 1
        .ColumnHeaders.Add , , "工龄工资 ", 50, 1
        .ColumnHeaders.Add , , "浮动工资 ", 50, 1
        .ColumnHeaders.Add , , "其他  ", 50, 1
        .ColumnHeaders.Add , , "应发合计", 50, 1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            Set Itm = .ListItems.Add()
            Itm.Text = Space(2) & Cells(i, 1)
            For c = 1 To 6
                Itm.SubItems(c) = Format(Cells(i, c + 1), "##,#,0.00")
            Next
        Next
        Set Img = ImageList1.ListImages.Add _
            (Picture:=LoadPicture(ThisWorkbook.Path & "\" & "1×25.bmp"))
        .SmallIcons = ImageList1
    End With
    Set Itm = Nothing
    Set Img = Nothing
End Sub

















