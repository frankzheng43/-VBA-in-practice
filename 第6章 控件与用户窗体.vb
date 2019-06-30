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

' 89-4 在Listview控件中排序
Option Explicit
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
        .Sorted = True
        .SortOrder = (.SortOrder + 1) Mod 2 '在设置SortOrder属性值时，使用Mod运算符以达到第一次排序以降序排序，再次排序时以升序排序，交替进行的效果。
        .SortKey = ColumnHeader.Index - 1
    End With
End Sub
Private Sub UserForm_Initialize()
    Dim Itm As ListItem
    Dim i As Integer
    Dim c As Integer
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
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row - 1
            Set Itm = .ListItems.Add()
            Itm.Text = Space(2) & Cells(i, 1)
            For c = 1 To 6
                Itm.SubItems(c) = Format(Cells(i, c + 1), "##,#,0.00")
            Next
        Next
    End With
    Set Itm = Nothing
End Sub

' 89-5 Listview控件的图标设置
Option Explicit
Private Sub UserForm_Initialize()
    Dim ITM As ListItem
    Dim i As Integer
    With ListView1
        .View = lvwIcon '视图样式
        .Icons = ImageList1
        For i = 2 To 6
            Set ITM = .ListItems.Add()
            ITM.Text = Cells(i, 1)
            ITM.Icon = i - 1
        Next
    End With
    Set ITM = Nothing
End Sub

Option Explicit
Private Sub UserForm_Initialize()
    Dim ITM As ListItem
    Dim i As Integer
    With ListView1
        .View = lvwSmallIcon
        .SmallIcons = ImageList1
        For i = 2 To 6
            Set ITM = .ListItems.Add()
            ITM.Text = Cells(i, 1)
            ITM.SmallIcon = i - 1
        Next
    End With
    Set ITM = Nothing
End Sub

Option Explicit
Private Sub UserForm_Initialize()
    Dim ITM As ListItem
    Dim i As Integer
    With ListView1
        .View = lvwList
        .SmallIcons = ImageList1
        For i = 2 To 6
            Set ITM = .ListItems.Add()
            ITM.Text = Cells(i, 1)
            ITM.SmallIcon = i - 1
        Next
    End With
    Set ITM = Nothing
End Sub

Option Explicit
Private Sub UserForm_Initialize()
    Dim ITM As ListItem
    Dim i As Integer
    Dim c As Integer
    With ListView1
        .ColumnHeaders.Add , , "人员编号 ", 70, 0
        .ColumnHeaders.Add , , "技能工资 ", 70, 1
        .ColumnHeaders.Add , , "岗位工资 ", 70, 1
        .View = lvwReport
        .Gridlines = True
        .SmallIcons = ImageList1
        For i = 2 To 6
            Set ITM = .ListItems.Add()
            ITM.Text = Cells(i, 1)
            ITM.SmallIcon = i - 1
            For c = 1 To 2
                ITM.SubItems(c) = Format(Cells(i, c + 1), "##,#,0.00")
            Next
        Next
    End With
    Set ITM = Nothing
End Sub

' 范例90 使用Toolbar控件添加工具栏
Option Explicit
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    MsgBox Button.Caption
End Sub
Private Sub UserForm_Initialize()
    Dim arr As Variant
    Dim i As Byte
    arr = Array(" 录入 ", " 审核", " 记账 ", " 结账 ", "负债表", "损益表")
    With Toolbar1
        .ImageList = ImageList1
        .Appearance = ccFlat
        .BorderStyle = ccNone
        .TextAlignment = tbrTextAlignBottom
        With .Buttons
            .Add(1, , "").Style = tbrPlaceholder
            For i = 0 To UBound(arr)
                .Add(i + 2, , , , i + 1).Caption = arr(i)
            Next
        End With
    End With
End Sub

' 范例91 使用StatusBar控件添加状态栏
Option Explicit
Private Sub TextBox1_Change()
    StatusBar1.Panels(1).Text = "正在输入：" & TextBox1.Text
End Sub
Private Sub UserForm_Initialize()
    Dim Pal As Panel
    Dim arr1 As Variant
    Dim arr2 As Variant
    Dim i As Integer
    arr1 = Array(0, 6, 5)
    arr2 = Array(180, 60, 54)
    StatusBar1.Width = 294
    For i = 1 To 3
        Set Pal = StatusBar1.Panels.Add()
        With Pal
            .Style = arr1(i - 1)
            .Width = arr2(i - 1)
            .Alignment = i - 1
        End With
    Next
    StatusBar1.Panels(1).Text = "准备就绪!"
End Sub

' 范例92 使用AniGif控件显示GIF图片
Option Explicit
Private Sub CommandButton1_Click()
    AniGif1.Stretch = True
    AniGif1.Filename = ThisWorkbook.Path & "\001.gif"
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub

' 范例93 使用ShockwaveFlash控件播放Flash文件
Option Explicit
Private Sub CommandButton1_Click()
    With ShockwaveFlash1
        .Movie = ThisWorkbook.Path & "\001.swf"
        .EmbedMovie = False
        .Menu = False
        .ScaleMode = 2
    End With
End Sub
Private Sub CommandButton2_Click()
    ShockwaveFlash1.Play
End Sub
Private Sub CommandButton3_Click()
    ShockwaveFlash1.Forward
End Sub
Private Sub CommandButton4_Click()
    ShockwaveFlash1.Stop
End Sub
Private Sub CommandButton5_Click()
    ShockwaveFlash1.Back
End Sub
Private Sub CommandButton6_Click()
    ShockwaveFlash1.Movie = " "
End Sub
Private Sub CommandButton7_Click()
    Unload Me
End Sub

' 范例94 注册自定义控件
Option Explicit
Sub Regsvrs()
    Dim SouFile As String
    Dim DesFile As String
    On Error Resume Next
    SouFile = ThisWorkbook.Path & "\VBAniGIF.OCX"
    DesFile = "C:\Windows\system32\VBAniGIF.OCX"
    FileCopy SouFile, DesFile
    Shell "REGSVR32 /s " & DesFile
    MsgBox "AniGif控件已成功注册，现在可以使用了!"
End Sub
Sub Regsvru()
    Shell "REGSVR32 /u C:\Windows\system32\VBAniGIF.OCX"
End Sub

' 范例95 不打印工作表中的控件

' 范例96 遍历控件的方法
Option Explicit
Private Sub CommandButton1_Click()
    Dim i As Integer
    For i = 1 To 3
        Me.Controls("TextBox" & i) = ""
    Next
End Sub

Option Explicit
Sub ClearText()
    Dim i As Integer
    For i = 1 To 4
        Sheet1.OLEObjects("TextBox" & i).Object.Text = ""
    Next
End Sub
Sub FormShow()
    UserForm1.Show
End Sub

' 96-2 使用对象类型
Option Explicit
Sub ClearText()
    Dim Obj As OLEObject
    For Each Obj In Sheet1.OLEObjects '遍历
        If TypeName(Obj.Object) = "TextBox" Then
            Obj.Object.Text = ""
        End If
    Next
    Set Obj = Nothing
End Sub
Option Explicit
Private Sub CommandButton1_Click()
    Dim Ctr As Control
    For Each Ctr In Me.Controls
        If TypeName(Ctr) = "TextBox" Then
            Ctr = ""
        End If
    Next
    Set Ctr = Nothing
End Sub

' 96-3 使用程序标识符
Option Explicit
Sub ClearText()
    Dim Obj As OLEObject
    For Each Obj In Sheet1.OLEObjects
        If Obj.progID = "Forms.TextBox.1" Then '程序标识符
            Obj.Object.Text = ""
        End If
    Next
    Set Obj = Nothing
End Sub

' 96-4 使用FormControlType属性
Option Explicit
Sub ControlType()
    Dim MyShape As Shape
    For Each MyShape In Sheet1.Shapes
        If MyShape.Type = msoFormControl Then
            If MyShape.FormControlType = xlCheckBox Then
                MyShape.ControlFormat.Value = 1
            End If
        End If
    Next
    Set MyShape = Nothing
End Sub

' 范例97 使用程序代码添加控件
Option Explicit
Sub AddButton()
    Dim MyButton As Button
    On Error Resume Next
    Sheet1.Shapes("MyButton").Delete
    Set MyButton = Sheet1.Buttons.Add(60, 40, 100, 30) '直接添加
    With MyButton
        .Name = "MyButton"
        .Font.Size = 12
        .Font.ColorIndex = 5
        .Characters.Text = "新建的按钮"
        .OnAction = "MyButton"
    End With
    Set MyButton = Nothing
End Sub
Sub MyButton()
    MsgBox "这是使用Add方法新建的按钮!"
End Sub

' 97-2 使用AddFormControl方法添加表单控件
Option Explicit
Sub AddButton()
    Dim MyShape As Shape
    On Error Resume Next
    Sheet1.Shapes("MyButton").Delete
    Set MyShape = Sheet1.Shapes.AddFormControl(0, 60, 40, 100, 30)
    With MyShape
        .Name = "MyButton"
        With .TextFrame.Characters
            .Font.ColorIndex = 3
            .Font.Size = 12
            .Text = "新建的按钮"
        End With
        .OnAction = "MyButton"
    End With
    Set MyShape = Nothing
End Sub
Sub MyButton()
    MsgBox "这是使用AddFormControl方法新建的按钮!"
End Sub

' 97-3 使用Add方法添加ActiveX控件
Option Explicit
Sub AddButton()
    Dim Obj As New OLEObject
    On Error Resume Next
    Sheet1.OLEObjects("MyButton").Delete
    Set Obj = Sheet1.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
            Left:=60, Top:=40, Width:=100, Height:=30)
    With Obj
        .Name = "MyButton"
        .Object.Caption = "新建的按钮"
        .Object.Font.Size = 12
        .Object.ForeColor = &HFF&
    End With
    With ActiveWorkbook.VBProject.VBComponents(Sheet1.CodeName).CodeModule
        If .Lines(1, 1) <> "Option Explicit" Then
            .InsertLines 1, "Option Explicit"
        End If
        If .Lines(2, 1) = "Private Sub MyButton_Click()" Then Exit Sub
        .InsertLines 2, "Private Sub MyButton_Click()"
        .InsertLines 3, vbTab & "MsgBox ""这是使用Add方法新建的按钮!"""
        .InsertLines 4, "End Sub"
    End With
    Set Obj = Nothing
End Sub

' 97-4 使用AddOLEObject方法添加ActiveX控件
Option Explicit
Sub AddButton()
    Dim MyButton As Shape
    On Error Resume Next
    Sheet1.Shapes("MyButton").Delete
    Set MyButton = Sheet1.Shapes.AddOLEObject( _
        ClassType:="Forms.CommandButton.1", _
        Left:=60, Top:=40, Width:=100, Height:=30)
    MyButton.Name = "MyButton"
    With ActiveWorkbook.VBProject.VBComponents(Sheet1.CodeName).CodeModule
        If .Lines(1, 1) <> "Option Explicit" Then
            .InsertLines 1, "Option Explicit"
        End If
        If .Lines(2, 1) = "Private Sub MyButton_Click()" Then Exit Sub
        .InsertLines 2, "Private Sub MyButton_Click()"
        .InsertLines 3, vbTab & "MsgBox ""这是使用AddOLEObject方法新建的按钮!"""
        .InsertLines 4, "End Sub"
    End With
    Set MyButton = Nothing
End Sub

' 范例98 禁用用户窗体的关闭按钮
Option Explicit
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then '如果是从界面上关闭的，则取消这个操作
        Cancel = True
        MsgBox "请点击""关闭""按钮关闭用户窗体!"
    End If
End Sub

' 范例99 屏蔽用户窗体的关闭按钮
' 只能在32位上运行
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Hwnd As Long
Private Sub UserForm_Initialize()
    Dim Istype As Long
    Hwnd = FindWindow("ThunderDFrame", Me.Caption)
    Istype = GetWindowLong(Hwnd, GWL_STYLE)
    Istype = Istype And Not WS_SYSMENU
    SetWindowLong Hwnd, GWL_STYLE, Istype
    DrawMenuBar Hwnd
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub

' 范例100 为用户窗体添加图标
' 只能在32位上运行
Option Explicit
Dim hwnd As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
Sub ChangeIcon(ByVal hwnd As Long, Optional ByVal hIcon As Long = 0&) '改变窗体图标
    SendMessage hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon
    SendMessage hwnd, WM_SETICON, ICON_BIG, ByVal hIcon
    DrawMenuBar hwnd
End Sub
Private Sub UserForm_Initialize()
    hwnd = FindWindow(vbNullString, Me.Caption) '获得窗口句柄
    Call ChangeIcon(hwnd, Image1.Picture.Handle)
End Sub

' 范例101 为用户窗体添加最大最小化按纽
' 只能在32位上运行
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)
Private Sub UserForm_Initialize()
    Dim hWndForm As Long
    Dim iStyle As Long
    hWndForm = FindWindow("ThunderDFrame", Me.Caption)
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_MINIMIZEBOX
    iStyle = iStyle Or WS_MAXIMIZEBOX
    SetWindowLong hWndForm, GWL_STYLE, iStyle
End Sub

' 范例102 屏蔽用户窗体的标题栏和边框
' 只能在32位上运行
Option Explicit
Private Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Sub UserForm_Initialize()
    Dim IStyle As Long
    Dim Hwnd As Long
    If Val(Application.Version) < 9 Then
        Hwnd = FindWindow("ThunderXFrame", Me.Caption)
    Else
        Hwnd = FindWindow("ThunderDFrame", Me.Caption)
    End If
    IStyle = GetWindowLong(Hwnd, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    SetWindowLong Hwnd, GWL_STYLE, IStyle
    DrawMenuBar Hwnd
    IStyle = GetWindowLong(Hwnd, GWL_EXSTYLE) And Not WS_EX_DLGMODALFRAME
    SetWindowLong Hwnd, GWL_EXSTYLE, IStyle
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub

' 范例103 显示透明的用户窗体
' 只能在32位上运行
Option Explicit
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWndForm As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWndForm As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWndForm As Long, ByVal crKey As Integer, ByVal bAlpha As Integer, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = &HFFEC
Dim hWndForm As Long
Private Sub UserForm_Activate()
    Dim nIndex As Long
    hWndForm = GetActiveWindow
    nIndex = GetWindowLong(hWndForm, GWL_EXSTYLE)
    SetWindowLong hWndForm, GWL_EXSTYLE, nIndex Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWndForm, 0, (255 * 60) / 100, LWA_ALPHA
End Sub

' 范例104 为用户窗体添加菜单
' 只能在32位上运行
Option Explicit
Public PreWinProc As Long, hwnd As Long
Public Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_ENABLED = &H0&
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Const MF_BYCOMMAND = &H0&
Public Function MsgProcess(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim SubMenu_hWnd As Long
    Select Case wParam
        Case 100
            MsgBox "你选择的是""保存""按钮!"
        Case 101
            MsgBox "你选择的是""备份""按钮!"
        Case 102
            Unload UserForm1
        Case 110
            MsgBox "你选择的是""录入""按钮!"
        Case 111
            MsgBox "你选择的是""审核""按钮!"
        Case 112
            MsgBox "你选择的是""记账""按钮!"
        Case 113
            MsgBox "你选择的是""结账""按钮!"
        Case 114
            MsgBox "你选择的是""资产负债表""按钮!"
        Case 115
            MsgBox "你选择的是""损益表""按钮!"
        Case Else
            MsgProcess = CallWindowProc(PreWinProc, hwnd, Msg, wParam, lParam)
    End Select
End Function

' 范例105 自定义用户窗体的鼠标指针类型
Option Explicit
Private Sub UserForm_Initialize()
    Me.MousePointer = 99
    Me.MouseIcon = LoadPicture(ThisWorkbook.Path & "\myMouse.ico")
End Sub

' 范例106 用户窗体的打印
Option Explicit
Private Sub CommandButton6_Click() '退出按钮
    Unload Me
End Sub
Private Sub CommandButton7_Click() '打印按钮
    Dim myHeight As Integer
    With UserForm1
        myHeight = .Height
        .Frame1.Visible = False
        .Height = myHeight - 30
        .PrintForm '打印
        .Height = myHeight
        .Frame1.Visible = True
    End With
End Sub

' 范例107 设置用户窗体的显示位置

' 107-1 调整用户窗体的显示位置
Option Explicit
Private Sub UserForm_Initialize()
' The Me keyword behaves like an implicitly declared variable. 
' It is automatically available to every procedure in a class module.
    With Me 
        .StartUpPosition = 0
        .Left = 25
        .Top = 75
    End With
End Sub

' 107-2 由活动单元格确定显示位置
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim CellX As Integer
    Dim CellY As Integer
    CellX = ExecuteExcel4Macro("GET.CELL(42)")
    CellY = ExecuteExcel4Macro("GET.CELL(43)")
    With UserForm1
        .Show 0
        .Left = CellX
        .Top = CellY + 60
    End With
End Sub

' 范例108 用户窗体的全屏显示

' 108-1 设置用户窗体的大小为应用程序的大小
Option Explicit
Private Sub UserForm_Initialize()
    With Application
        .WindowState = xlMaximized
        Width = .Width
        Height = .Height
        Left = .Left
        Top = .Top
    End With
End Sub

' 108-2 根据屏幕分辨率设置
' 只能在32位上运行
Option Explicit
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN As Long = 0
Const SM_CYSCREEN As Long = 1
Private Sub UserForm_Initialize()
    With Me
        .Height = GetSystemMetrics(SM_CYSCREEN) * 0.75
        .Width = GetSystemMetrics(SM_CXSCREEN) * 0.75
        .Left = 0
        .Top = 0
    End With
End Sub

' 范例109 在用户窗体中显示图表

' 109-1 使用Export方法显示图表
Option Explicit
Private Sub UserForm_Initialize()
    Dim myChart As Chart
    Dim str As String
    Set myChart = Sheets("Sheet2").ChartObjects(1).Chart
    str = ThisWorkbook.Path & "\Temp.gif"
    myChart.Export Filename:=str, FilterName:="GIF"
    Image1.Picture = LoadPicture(str)
  End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Kill ThisWorkbook.Path & "\Temp.gif"
End Sub

' 109-2 使用API函数显示图表
Option Explicit
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Function LoadShapePicture(shp As Object) As IPictureDisp
    Dim nClipsize As Long
    Dim hMem As Long
    Dim lpData As Long
    Dim sdata() As Byte
    Dim fmt As Long
    Dim fmtName As String
    Dim iClipBoardFormatNumber As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown
    If TypeName(shp) = "ChartObject" Then
        shp.CopyPicture xlPrinter
        Sheet1.Paste
        Selection.Cut
    Else
        shp.Copy
    End If
    OpenClipboard 0&
    If iClipBoardFormatNumber = 0 Then
        fmt = EnumClipboardFormats(0)
        Do While fmt <> 0
            fmtName = Space(255)
            GetClipboardFormatName fmt, fmtName, 255
            fmtName = Trim(fmtName)
            If fmtName <> "" Then
                fmtName = Left(fmtName, Len(fmtName) - 1)
                If fmtName = "GIF" Then
                    iClipBoardFormatNumber = fmt
                    Exit Do
                End If
            End If
            fmt = EnumClipboardFormats(fmt)
         Loop
    End If
    hMem = GetClipboardData(iClipBoardFormatNumber)
    If CBool(hMem) Then
        nClipsize = GlobalSize(hMem)
        lpData = GlobalLock(hMem)
        GlobalUnlock hMem
        If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
            If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                Call OleLoadPicture(ByVal ObjPtr(istm), nClipsize, 0, IID_IPicture(0), LoadShapePicture)
            End If
        End If
    End If
    EmptyClipboard
    CloseClipboard
End Function
Private Sub UserForm_Initialize()
    Image1.Picture = LoadShapePicture(Sheet2.ChartObjects(1))
End Sub

' 范例110 用户窗体运行时调整控件大小
Option Explicit
Dim Abscissa As Single
Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Abscissa = x
End Sub
Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then
        If Abscissa - x > Frame1.Width Or x > Frame2.Width Then Exit Sub
        Frame1.Width = Frame1.Width - Abscissa + x
        Image1.Left = Image1.Left - Abscissa + x
        Frame2.Left = Frame2.Left - Abscissa + x
        Frame2.Width = Frame2.Width + Abscissa - x
    End If
End Sub

' 范例111 使用代码添加用户窗体及控件
Option Explicit
Sub CreatingForms()
    Dim MyForm As VBComponent
    Dim MyTextBox As Control
    Dim MyButton As Control
    Dim i As Integer
    On Error Resume Next
    Set MyForm = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    With MyForm
        .Properties("Name") = "Formtest"
        .Properties("Caption") = "演示窗体"
        .Properties("Height") = "180"
        .Properties("Width") = "240"
        Set MyTextBox = .Designer.Controls.Add("Forms.CommandButton.1")
        With MyTextBox
            .Name = "MyTextBox"
            .Caption = "新建文本框"
            .Top = 40
            .Left = 138
            .Height = 20
            .Width = 70
        End With
        Set MyButton = .Designer.Controls.Add("Forms.CommandButton.1")
        With MyButton
            .Name = "MyButton"
            .Caption = "删除文本框"
            .Top = 70
            .Left = 138
            .Height = 20
            .Width = 70
        End With
        With .CodeModule
            i = .CreateEventProc("Click", "MyTextBox")
            .ReplaceLine i + 1, Space(4) & "Dim MyTextBox As Control" & vbCrLf & Space(4) & "Dim i As Integer" & vbCrLf & Space(4) & "Dim k As Integer" _
                & vbCrLf & Space(4) & "k = 10" & vbCrLf & Space(4) & "For i = 1 To 5" & vbCrLf & Space(8) & "Set MyTextBox = Me.Controls.Add(bstrprogid:=""Forms.TextBox.1"")" _
                & vbCrLf & Space(8) & "With MyTextBox" & vbCrLf & Space(12) & ".Name = ""MyTextBox"" & i" & vbCrLf & Space(12) & ".Left = 20" _
                & vbCrLf & Space(12) & ".Top = k" & vbCrLf & Space(12) & ".Height = 18" & vbCrLf & Space(12) & ".Width = 80" _
                & vbCrLf & Space(12) & "k = .Top + 28" & vbCrLf & Space(8) & "End With" & vbCrLf & Space(4) & "Next"
            i = .CreateEventProc("Click", "MyButton")
            .ReplaceLine i + 1, Space(4) & "Dim i As Integer" & vbCrLf & Space(4) & "On Error Resume Next" & vbCrLf & Space(4) & "For i = 1 To 5" & vbCrLf & Space(8) & "Me.Controls.Remove ""MyTextBox"" & i" & vbCrLf & Space(4) & "Next"
        End With
    End With
    Set MyForm = Nothing
    Set MyTextBox = Nothing
    Set MyButton = Nothing
End Sub

' 范例112 以非模式显示用户窗体
Option Explicit
Sub vbModelessProgressBar()
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
Sub vbModelProgressBar()
    Dim r As Integer
    Dim i As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row
        UserForm1.Show
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




    
    
    




























