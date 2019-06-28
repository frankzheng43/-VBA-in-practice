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






