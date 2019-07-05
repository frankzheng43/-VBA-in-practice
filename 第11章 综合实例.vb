' 范例153 制作员工花名册
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
    Sheet1.Unprotect
    With Target
        If .Count = 1 And .Row > 5 Then
            If .Column = 2 And Len(.Text) <> 0 Then
                Range(Cells(.Row, 1), Cells(.Row, 10)).Borders.LineStyle = xlContinuous
                .Offset(0, -1).FormulaR1C1 = "=ROW()-5"
            End If
            If .Column = 6 Then
                Application.EnableEvents = False
                Select Case Len(.Text)
                    Case 18
                        .Offset(0, -5).FormulaR1C1 = "=ROW()-5"
                        .Offset(0, -3) = IIf(Mid(.Text, 17, 1) Mod 2 = 0, "女", "男")
                        .Offset(0, -2) = Format(Mid(.Text, 7, 8), "#-00-00")
                        .Offset(0, -1).FormulaR1C1 = "=DATEDIF(TEXT(MID(RC[1],7,8),""#-00-00""),TODAY(),""y"")"
                    Case 0
                        Range(Cells(.Row, 3), Cells(.Row, 6)) = ""
                    Case Else
                        .Select
                        MsgBox "身份证号码不正确,请重新输入!", 64, "提示"
                End Select
                Application.EnableEvents = True
            End If
        End If
    End With
    Sheet1.Protect
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim r As Integer
    r = Sheet1.Cells(Sheet1.Rows.Count, 2).End(xlUp).Row
    With Target
        If .Count = 1 And .Row > 5 And .Row <= r Then
            Sheet1.Unprotect
            Select Case .Column
                Case 7
                    With .Validation
                        .Delete
                        .Add 3, 1, 1, "高级会计师,会计师,助理会计师,会计员," _
                                & "高级工程师,工程师,助理工程师,技术员," _
                                & "高级经济师,经济师,助理经济师,无"
                    End With
                Case 8
                    With .Validation
                        .Delete
                        .Add 3, 1, 1, "经理室,办公室,行政科,生技科," _
                                & "财务科,营业部,其他,退休"
                    End With
                Case 9
                    With .Validation
                        .Delete
                        .Add 3, 1, 1, "经理,副经理,中层正职,中层副职," _
                                & "总账会计,辅助会计,出纳会计,驾驶员," _
                                & "办事员,收费员,发货员,采购员,化验员," _
                                & "班组长,电工,中控值班,制水工,外借,内退"
                    End With
                Case 10
                    With .Validation
                        .Delete
                        .Add 3, 1, 1, "在职,内退,退休"
                    End With
                Case Else
            End Select
            Sheet1.Protect
        End If
    End With
End Sub

Option Explicit
Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim r1 As Integer
    Dim r2 As Integer
    Application.ScreenUpdating = False
    With Sheet2
        .Unprotect
        r2 = .Cells(.Rows.Count, 2).End(xlUp).Row
        If r2 > 5 Then
            With .Range("A6:J" & r2)
                .ClearContents
                .Borders.LineStyle = xlNone
            End With
        End If
        Sheet1.Unprotect
        r1 = Sheet1.Cells(Sheet1.Rows.Count, 2).End(xlUp).Row
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
                Sheet1.Range("A5:J" & r1).AutoFilter Field:=8, Criteria1:="=" & ListBox1.List(i)
                Sheet1.Range("A6:J" & r1).SpecialCells(12).Copy
                r2 = .Cells(.Rows.Count, 2).End(xlUp).Row
                .Cells(r2 + 1, 1).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                .Range("A6:A" & .Cells(.Rows.Count, 2).End(xlUp).Row).FormulaR1C1 = "=ROW()-5"
            End If
        Next
        Sheet1.Range("A1:J" & r1).AutoFilter
        Sheet1.Protect
        r2 = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Range("A6:J" & r2).Borders.LineStyle = xlContinuous
        Application.Goto Reference:=.Range("A2"), Scroll:=True
        .Protect
    End With
    Unload Me
    Application.ScreenUpdating = True
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim Col As New Collection
    Dim rng As Range
    Dim arr As Variant
    Dim i As Integer
    On Error Resume Next
    With Sheet1
        r = .Cells(.Rows.Count, 2).End(xlUp).Row
        For Each rng In .Range("H6:H" & r)
            Col.Add rng, key:=CStr(rng)
        Next
        ReDim arr(1 To Col.Count)
        For i = 1 To Col.Count
            arr(i) = Col(i)
        Next
    End With
    With Me.ListBox1
        .List = arr
        .ListStyle = 1
        .MultiSelect = 1
    End With
    Set rng = Nothing
End Sub

Option Explicit
Private Sub CommandButton1_Click()
    Dim r1 As Integer
    Dim r2 As Integer
    Application.ScreenUpdating = False
    If Frame1.ComboBox1.Value = "" Or Frame1.ComboBox2.Value = "" Then
        MsgBox "请选择需要筛选的年龄!"
        Exit Sub
    End If
    If Frame1.ComboBox1.Value > Frame1.ComboBox2.Value Then
        MsgBox "开始年龄不能大于结束年龄!"
        Frame1.ComboBox1.ListIndex = -1
        Frame1.ComboBox2.ListIndex = -1
        Exit Sub
    End If
    With Sheet1
        r1 = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Unprotect
        .Range("A5:J" & r1).AutoFilter Field:=5, Criteria1:=">=" _
            & Frame1.ComboBox1.Value, Operator:=xlAnd, _
            Criteria2:="<=" & Frame1.ComboBox2.Value
        With Sheet2
            .Unprotect
            r2 = .Cells(.Rows.Count, 2).End(xlUp).Row
            If r2 > 5 Then
                With .Range("A6:J" & r2)
                    .ClearContents
                    .Borders.LineStyle = xlNone
                End With
            End If
            Sheet1.Range("A6:J" & r1).SpecialCells(12).Copy
            .Cells(6, 1).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            r2 = .Cells(.Rows.Count, 2).End(xlUp).Row
            .Range("A6:A" & r2).FormulaR1C1 = "=ROW()-5"
            .Range("A6:J" & r2).Borders.LineStyle = xlContinuous
            Application.Goto Reference:=.Range("A3"), Scroll:=True
            .Protect
        End With
        .Range("A1:J" & r1).AutoFilter
        .Protect
    End With
    Unload Me
    Application.ScreenUpdating = True
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Dim r As Integer
    Dim Col As New Collection
    Dim rng As Range
    Dim arr As Variant
    Dim i As Integer
    On Error Resume Next
    With Sheet1
        r = .Cells(.Rows.Count, 2).End(xlUp).Row
        For Each rng In .Range("E6:E" & r)
            Col.Add rng, key:=CStr(rng)
        Next
        ReDim arr(1 To Col.Count)
        For i = 1 To Col.Count
            arr(i) = Col(i)
        Next
    End With
    Me.Frame1.ComboBox1.List = arr
    Me.Frame1.ComboBox2.List = arr
    Set rng = Nothing
End Sub

Option Explicit
Sub AgeSort()
    Dim r As Integer
    Dim Mymsg As Integer
    With Sheet1
        r = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Unprotect
        Mymsg = MsgBox("选择""是""按升序排序,选择""否""按降序排序!", vbYesNoCancel)
        Select Case Mymsg
            Case 6
                .Range("A6:J" & r).Sort Key1:=.Range("E6"), _
                    Order1:=xlAscending, Key2:=.Range("D6")
            Case 7
                .Range("A6:J" & r).Sort Key1:=.Range("E6"), _
                    Order1:=xlDescending, Key2:=.Range("D6")
            Case Else
        End Select
        .Protect
    End With
End Sub
Sub SectorSort()
    Dim r As Integer
    With Sheet1
        .Unprotect
        r = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Range("A6:J" & r).Sort Key1:=.Range("H6"), _
            Order1:=xlAscending, Key2:=Range("D6"), _
            OrderCustom:=13
        .Protect
    End With
End Sub
Sub Forshow()
    Dim r As Integer
    With Sheet1
        .Unprotect
        r = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Range("A6:J" & r).Sort Key1:=.Range("H6"), _
            Order1:=xlAscending, Key2:=Range("D6"), _
            OrderCustom:=13
        .Protect
    End With
    按部门筛选.Show
End Sub
Sub AgeSortForshow()
    Dim r As Integer
    With Sheet1
        .Unprotect
        r = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Range("A6:J" & r).Sort Key1:=.Range("E6"), _
            Order1:=xlAscending, Key2:=.Range("D6")
        .Protect
    End With
    按年龄筛选.Show
End Sub

' 范例154 制作收据打印系统
