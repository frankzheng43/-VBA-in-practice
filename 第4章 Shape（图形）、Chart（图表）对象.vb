' 范例47 在工作表中添加图形
Option Explicit
Sub AddingGraphics()
    Dim MyShape As Shape
    On Error Resume Next
    Sheet1.Shapes("MyShape").Delete
    Set MyShape = Sheet1.Shapes.AddShape(msoShapeRectangle, 40, 120, 280, 30)
    With MyShape
        .Name = "MyShape"
        With .TextFrame.Characters
            .Text = "单击将选择Sheet2!"
            With .Font
                .Size = 20
                .ColorIndex = 5
            End With
        End With
        With .Line
            .Weight = 1
            .Style = msoLineSingle
            .Transparency = 0.5
            .ForeColor.SchemeColor = 40
            .BackColor.RGB = RGB(255, 255, 255)
        End With
        With .Fill
            .Transparency = 0.5
            .ForeColor.SchemeColor = 41
            .OneColorGradient 1, 4, 0.23
        End With
        .Placement = 3
    End With
    Sheet1.Hyperlinks.Add Anchor:=MyShape, Address:="", _
        SubAddress:="Sheet2!A1", ScreenTip:="选择Sheet2!"
    Set MyShape = Nothing
End Sub

' 范例48 导出工作表中的图片
Option Explicit
Sub ExportPictures()
    Dim MyShp As Shape
    Dim Filename As String
    For Each MyShp In Sheet1.Shapes
        If MyShp.Type = msoPicture Then
            Filename = ThisWorkbook.Path & "\" & MyShp.Name & ".gif"
            MyShp.Copy
            With Sheet1.ChartObjects.Add(0, 0, MyShp.Width, MyShp.Height).Chart
                .Paste
                .Export Filename
                .Parent.Delete
            End With
        End If
    Next
    Set MyShp = Nothing
End Sub

' 范例49 在工作表中添加艺术字
Option Explicit
Sub AddingWordArt()
    On Error Resume Next
    Sheet1.Shapes("MyShape").Delete
    Sheet1.Shapes.AddTextEffect _
        (PresetTextEffect:=msoTextEffect16, _
        Text:="Excel 2007", FontName:="宋体", _
        FontSize:=50, FontBold:=True, _
        FontItalic:=True, Left:=60, Top:=60).Name = "MyShape"
End Sub

' 范例50 遍历工作表中的形状
Option Explicit
Sub TraversalShapeOne()
    Dim i As Integer
    For i = 1 To 4
        Sheet1.Shapes("文本框 " & i).TextFrame.Characters.Text = ""
    Next
End Sub
Sub TraversalShapeTwo()
    Dim MyShape As Shape
    Dim MyCount As Integer
    MyCount = 1
    For Each MyShape In Sheet1.Shapes
        If MyShape.Type = msoTextBox Then
            MyShape.TextFrame.Characters.Text = "第" & MyCount & "个文本框"
            MyCount = MyCount + 1
        End If
    Next
    Set MyShape = Nothing
End Sub

' 范例51 移动、旋转图形
Option Explicit
Sub MoveAndRotate()
    Dim i As Long
    Dim j As Long
    With Sheet1.Shapes(1)
        For i = 1 To 3000 Step 5
           .Top = Sin(i * (3.1415926535 / 180)) * 100 + 100
           .Left = Cos(i * (3.1415926535 / 180)) * 100 + 100
           .Fill.ForeColor.RGB = i * 100
            For j = 1 To 20
                .IncrementRotation -2
                DoEvents
            Next
        Next
    End With
End Sub

' 范例52 自动插入图片
Option Explicit
Sub InsertPicture()
    Dim MyShape As Shape
    Dim r As Integer
    Dim c As Integer
    Dim PicPath As String
    Dim Picrng As Range
    With Sheet1
        For Each MyShape In .Shapes
            If MyShape.Type = 13 Then
                MyShape.Delete
            End If
        Next
        For r = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            For c = 1 To 8 Step 2
                PicPath = ThisWorkbook.Path & "\" & .Cells(r, c).Text & ".jpg"
                If Dir(PicPath) <> "" Then '查找是否有这张图
                    Set MyShape = .Shapes.AddPicture(PicPath, False, True, 6, 6, 6, 6)
                    Set Picrng = .Cells(r, c + 1)
                    With MyShape
                        .LockAspectRatio = msoFalse
                        .Top = Picrng.Top + 1
                        .Left = Picrng.Left + 1
                        .Width = Picrng.Width - 1.5
                        .Height = Picrng.Height - 1.5
                        .TopLeftCell = ""
                    End With
                Else
                    .Cells(r, c + 1) = "暂无照片"
                End If
            Next
        Next
    End With
    Set MyShape = Nothing
    Set Picrng = Nothing
End Sub

' 范例53 固定图片的尺寸和位置
Option Explicit
Sub FixedPicture()
    Dim Picrng As Range
    Set Picrng = Range("B4:E22")
    With Sheet1.Shapes("Picture 1")
        .Rotation = 0 '旋转
        .Top = Picrng.Top - 1
        .Left = Picrng.Left - 1
        .Width = Picrng.Width + 1
        .Height = Picrng.Height + 1
    End With
    Set Picrng = Nothing
End Sub

' 范例54 使用VBA自动生成图表
Option Explicit
Sub ProductionCharts()
    Dim r As Integer
    Dim rng As Range
    Dim MyChart As ChartObject
    On Error Resume Next
    With Sheet1
       .ChartObjects("MyChart").Delete '先删掉之前存在的表
        r = .Cells(.Rows.Count, 1).End(xlUp).Row '选择最后一行
        Set rng = .Range(.Cells(1, 1), .Cells(r, 2)) '数据区域
        '表达式.Add （Left， Top， width， Height）
        Set MyChart = .ChartObjects.Add(120, 40, 400, 250) 
        MyChart.Name = "MyChart"
        With MyChart.Chart
        'https://docs.microsoft.com/en-us/office/vba/api/excel.xlcharttype
            '这个是attribute
            .ChartType = xlLineMarkers
            '下面两个是method
            .SetSourceData Source:=rng, PlotBy:=xlColumns
            .ApplyDataLabels ShowValue:=True
            'https://docs.microsoft.com/en-us/office/vba/api/excel.chart.applycharttemplate
            '这个是属性
            .HasTitle = True
            ' chart里还有一个charttitle的子类 
            With .ChartTitle
                .Text = "图表制作示例"
                .Font.Name = "宋体"
                .Font.Size = 14
            End With
        End With
    End With
    Set rng = Nothing
    Set MyChart = Nothing
End Sub

' 范例55 批量制作图表
Option Explicit
Sub ProductionCharts()
    Dim MyChart As ChartObject
    Dim i As Integer
    Dim r As Integer
    Dim m As Integer
    On Error Resume Next
    Sheet2.ChartObjects.Delete
    With Sheet1
        r = .Cells(.Rows.Count, 1).End(xlUp).Row - 1
        m = Abs(Int(-(r / 4)))
        For i = 1 To r
            Set MyChart = Sheet2.ChartObjects.Add _
                (Left:=(((i - 1) Mod m) + 1) * 350 - 340, _
                Top:=((i - 1) \ m + 1) * 220 - 210, _
                Width:=300, Height:=200)
            MyChart.Name = .Range("A2").Offset(i - 1)
            With MyChart.Chart
                .ChartType = xl3DColumnStacked
                .SetSourceData Source:=Sheet1.Range("B2:M2").Offset(i - 1), _
                    PlotBy:=xlRows
                .HasTitle = True
                .HasLegend = False
                With .ChartTitle
                    .Text = Sheet1.Range("A2").Offset(i - 1)
                    .Font.Name = "微软雅黑"
                    .Font.Size = 12
                End With
            End With
        Next
    End With
    Sheet2.Select
    Set MyChart = Nothing
End Sub

' 范例56 导出工作表中的图表
Option Explicit
Sub ExportChart()
    Dim ChartPath As String
    ChartPath = ThisWorkbook.Path & "\" & "MyChart.jpg"
    On Error Resume Next
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/kill-statement
    'Deletes files from a disk
    Kill ChartPath
    Sheet1.ChartObjects(1).Chart.Export FileName:=ChartPath, Filtername:="JPG"
    MsgBox "图表已保存在""" & ThisWorkbook.Path & """文件夹中!"
End Sub










