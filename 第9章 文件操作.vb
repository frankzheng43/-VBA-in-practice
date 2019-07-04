' 范例134 导入文本文件

' 134-1 使用查询表导入
Option Explicit
Sub AddQuery()
    With Sheet2
        .UsedRange.ClearContents
        'Returns the QueryTables collection that represents all the query tables on the specified worksheet. Read-only.
        With .QueryTables.Add(Connection:="TEXT;" & ThisWorkbook.Path & "\工资表.txt", Destination:=.Range("A1"))
            .TextFileCommaDelimiter = True
            .Refresh
        End With
        .Select
    End With
End Sub

' 延伸 QueryTables
For Each qt in Worksheets(1).QueryTables 
 qt.Refresh 
Next


' 134-2 使用Open 语句导入
Option Explicit
Sub OpenText()
    Dim MyText As String
    Dim MyArr() As String
    Dim c As Integer
    Dim r As Integer
    r = 1
    With Sheet2
        .UsedRange.ClearContents
        Open ThisWorkbook.Path & "\工资表.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, MyText
            MyArr = Split(MyText, ",")
            For c = 0 To UBound(MyArr)
                .Cells(r, c + 1) = MyArr(c)
            Next
            r = r + 1
        Loop
        Close #1
        .Select
    End With
End Sub

' 134-3 使用OpenText方法导入
Option Explicit
Sub OpenText()
    Sheet2.UsedRange.ClearContents
    Workbooks.OpenText Filename:=ThisWorkbook.Path & "\" & "工资表.txt", StartRow:=1, DataType:=xlDelimited, Comma:=True
    With ActiveWorkbook
        With .Sheets("工资表").Range("A1").CurrentRegion
            ThisWorkbook.Sheets("Sheet2").Range("A1").Resize(.Rows.Count, .Columns.Count).Value = .Value
        End With
        .Close False
    End With
    Sheet2.Select
End Sub

' 范例135 创建文本文件

' 135-1 使用Print # 语句将数据写入文本文件
Option Explicit
Sub PrintText()
    Dim File As String
    Dim Arr() As Variant
    Dim Str As String
    Dim r As Integer
    Dim c As Integer
    Dim i As Integer
    Dim j As Integer
    On Error Resume Next
    File = ThisWorkbook.Path & "\" & "工资表.txt"
    Kill File
    With Sheet2
        r = .UsedRange.Rows.Count
        c = .UsedRange.Columns.Count
        ReDim Arr(1 To r, 1 To c)
        For i = 1 To r
            For j = 1 To c
                Arr(i, j) = .Cells(i, j).Value '将工作表数据赋予数组Arr
            Next
        Next
    End With
    Open File For Output As #1
    For i = 1 To UBound(Arr, 1)
        Str = ""
        For j = 1 To UBound(Arr, 2)
            Str = Str & CStr(Arr(i, j)) & ","
        Next
        Str = Left(Str, (Len(Str) - 1))
        Print #1, Str
    Next
    Close #1
    MsgBox "文件保存成功!"
End Sub

' 135-2 使用SaveAs方法.将数据另存为文本文件
Option Explicit
Sub SaveText()
    Dim File As String
    File = ThisWorkbook.Path & "\工资表.txt"
    On Error Resume Next
    Kill File
    Sheet2.Copy
    ActiveWorkbook.SaveAs FileName:=File, FileFormat:=xlCSV
    ActiveWorkbook.Close SaveChanges:=False
    MsgBox "文件保存成功!"
End Sub

' 范例136 获得文件修改的日期和时间
' FileDateTime
Option Explicit
Sub MyDateTime()
    Dim Str As String
    Str = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    MsgBox Str & "的最后修改时间是：" & Chr(13) & FileDateTime(Str)
End Sub

' 范例137 查找文件和文件夹
Option Explicit
Sub MyName()
    Dim MyName As String
    Dim r As Integer
    r = 1
    Columns("A").ClearContents
    'vbDirectory 16 Specifies directories or folders in addition to files with no attributes.
    'Dir 函数会返回匹配pathname参数的第一个文件名，如果已没有合乎条件的文件，则会返回一个零长度字符串。
    MyName = Dir(ThisWorkbook.Path & "\", vbDirectory)
    Do While MyName <> "" '排除空
        If MyName <> "." And MyName <> ".." Then '排除本目录和上级目录
            Cells(r, 1) = MyName
            r = r + 1
        End If
        MyName = Dir '当Dir函数返回匹配pathname参数的第一个文件名后，若要得到其他匹配pathname参数的文件名，需再一次调用Dir 函数，且不要使用参数。
    Loop
End Sub

' 范例138 获得当前文件夹
Option Explicit
Sub CurFolder()
    MsgBox CurDir("D")
End Sub

' 范例139 创建和删除文件夹
' mkdir rmdir
Option Explicit
Sub CreateFolder()
    On Error Resume Next
    MkDir ThisWorkbook.Path & "\Temp"
End Sub
Sub DeleteFolder()
    On Error Resume Next
    RmDir ThisWorkbook.Path & "\Temp"
End Sub

' 范例140 重命名文件或文件夹
' name 
Option Explicit
Sub RenameFiles()
    Dim MyPath As String
    On Error Resume Next
    MyPath = ThisWorkbook.Path
    Name MyPath & "\123" As MyPath & "\ABC"
    Name MyPath & "\123.xlsx" As MyPath & "\ABC\ABC.xlsx"
End Sub

' 范例141 复制指定的文件
' FileCopy
Option Explicit
Sub CopyingFiles()
    Dim SourceFile As String
    Dim DestinationFile As String
    SourceFile = ThisWorkbook.Path & "\123.xlsx"
    DestinationFile = ThisWorkbook.Path & "\ABC\abc.xlsx"
    FileCopy SourceFile, DestinationFile
End Sub

' 范例142 删除指定的文件
' kill 
Option Explicit
Sub DeleteFiles()
    Dim myFile As String
    myFile = ThisWorkbook.Path & "\123.xlsx"
    If Dir(myFile) <> "" Then Kill myFile
End Sub

' 范例143 使用WSH处理文件

' 143-1 获取文件信息
Option Explicit
Sub FileInformation()
    Dim MyFile As Object
    Dim Str As String
    Dim StrMsg As String
    Str = ThisWorkbook.Path & "\123.xlsx"
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object
    Set MyFile = CreateObject("Scripting.FileSystemObject") '获得一个object，从而获得各种属性
    With MyFile.Getfile(Str)
        StrMsg = StrMsg & "文件名称：" & .Name & Chr(13) _
            & "文件创建日期：" & .DateCreated & Chr(13) _
            & "文件修改日期：" & .DateLastModified & Chr(13) _
            & "文件访问日期：" & .DateLastAccessed & Chr(13) _
            & "文件保存路径：" & .ParentFolder
    End With
    MsgBox StrMsg
    Set MyFile = Nothing
End Sub

' 143-2 取得文件基本名
' GetBaseName
Option Explicit
Sub FileBaseName()
    Dim MyFile As Object
    Dim FileName As Variant
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    FileName = Application.GetOpenFilename
    If FileName <> "False" Then
        MsgBox MyFile.GetBaseName(FileName)
    End If
    Set MyFile = Nothing
End Sub

' 143-3 查找文件
' FileExists
Option Explicit
Sub FindFiles()
    Dim MyFile As Object
    Dim Str As String
    Str = ThisWorkbook.Path & "\123.xlsx"
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    If Not MyFile.FileExists(Str) Then
        MsgBox "文件不存在!"
    Else
        MsgBox "文件已找到!"
    End If
    Set MyFile = Nothing
End Sub

' 143-4 搜索文件
Option Explicit
Sub SearchFiles()
    Dim MyFile As Object
    Dim MyFiles As Object
    Dim MyStr As String
    Set MyFile = CreateObject("Scripting.FileSystemObject") _
        .Getfolder(ThisWorkbook.Path)
    For Each MyFiles In MyFile.Files
        If InStr(MyFiles.Name, ".xlsx") <> 0 Then
            MyStr = MyStr & MyFiles.Name & Chr(13)
        End If
    Next
    MsgBox MyStr
    Set MyFile = Nothing
    Set MyFiles = Nothing
End Sub

' 143-5 移动文件
' MoveFile
Option Explicit
Sub MovingFiles()
    Dim MyFile As Object
    On Error Resume Next
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.MoveFile ThisWorkbook.Path & "\123.xlsx", ThisWorkbook.Path & "\abc\"
    Set MyFile = Nothing
End Sub

' 143-6 复制文件
' CopyFile
Option Explicit
Sub CopyingFiles()
    Dim MyFile As Object
    On Error Resume Next
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.CopyFile ThisWorkbook.Path & "\123.xlsx", ThisWorkbook.Path & "\abc\"
    Set MyFile = Nothing
End Sub

' 143-7 删除文件
' DeleteFile
Option Explicit
Sub DeleteFiles()
    Dim MyFile As Object
    On Error Resume Next
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.DeleteFile ThisWorkbook.Path & "\123.xlsx"
    Set MyFile = Nothing
End Sub

' 143-8 创建文件夹
' CreateFolder
Option Explicit
Sub CreateFolder()
    Dim MyFile As Object
    On Error Resume Next
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.CreateFolder (ThisWorkbook.Path & "\abc")
    Set MyFile = Nothing
End Sub

' 143-9 复制文件夹
' CopyFolder
Option Explicit
Sub CopyFolder()
    Dim MyFile As Object
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.CopyFolder ThisWorkbook.Path & "\ABC", ThisWorkbook.Path & "\123"
    Set MyFile = Nothing
End Sub

' 143-10 移动文件夹
' MoveFolder
Option Explicit
Sub MoveFolders()
    Dim MyFile As Object
    On Error Resume Next
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.MoveFolder ThisWorkbook.Path & "\123", ThisWorkbook.Path & "\abc\"
    Set MyFile = Nothing
End Sub

' 143-11 删除文件夹
' DeleteFolder
Option Explicit
Sub DeleteFolders()
    Dim MyFile As Object
    On Error Resume Next
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    MyFile.DeleteFolder ThisWorkbook.Path & "\123"
    Set MyFile = Nothing
End Sub

' 143-12 导入文本文件
Option Explicit
Sub ImportingText()
    Dim MyFile As Object
    Dim Arr() As String
    Dim r As Integer
    Dim i As Integer
    r = 1
    Sheet2.UsedRange.ClearContents
    Set MyFile = CreateObject("Scripting.FileSystemObject") _
            .OpenTextFile(ThisWorkbook.Path & "\" & "工资表.txt")
    Do While Not MyFile.AtEndOfStream
        Arr = Split(MyFile.ReadLine, ",")
        For i = 0 To UBound(Arr)
            Sheet2.Cells(r, i + 1) = Arr(i) '在这
        Next
        r = r + 1
    Loop
    MyFile.Close
    Sheet2.Select
    Set MyFile = Nothing
End Sub

' 143-13 创建文本文件
Option Explicit
Sub CreateTtextFile()
    Dim MyFile As Object
    Dim MyStr As String
    Dim r As Integer
    Dim c As Integer
    With Sheet2
        Set MyFile = CreateObject("Scripting.FileSystemObject") _
            .CreateTextFile(ThisWorkbook.Path & "\工资表.txt", True)
        For r = 1 To .UsedRange.Rows.Count
            MyStr = ""
            For c = 1 To .UsedRange.Columns.Count
                MyStr = MyStr & .Cells(r, c) & ","
            Next
            MyStr = Left(MyStr, (Len(MyStr) - 1))
            MyFile.WriteLine (MyStr)
        Next
        MyFile.Close
    End With
    Set MyFile = Nothing
End Sub
Sub CreateTtextFiles()
    Dim MyFile As Object
    Dim MyStr As String
    Dim r As Integer
    Dim c As Integer
    With Sheet2
        Set MyFile = CreateObject("Scripting.FileSystemObject") _
            .OpenTextFile(ThisWorkbook.Path & "\" & "工资表.txt", 2, True)
            For r = 1 To .UsedRange.Rows.Count
                MyStr = ""
                For c = 1 To .UsedRange.Columns.Count
                    MyStr = MyStr & .Cells(r, c) & ","
                Next
                MyStr = Left(MyStr, (Len(MyStr) - 1))
                MyFile.WriteLine (MyStr)
            Next
        MyFile.Close
    End With
    Set MyFile = Nothing
End Sub

' 143-14 取得驱动器信息
Option Explicit
Sub DiskInformation()
    Dim MyFile As Object
    Dim MyDisk As Object
    Dim MyStr As String
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    With MyFile.GetDrive("C")
        MyStr = MyStr & "驱动器盘符:" & UCase(.DriveLetter) & vbCrLf _
            & "驱动器类型:" & .DriveType & vbCrLf _
            & "驱动器文件系统:" & .FileSystem & vbCrLf _
            & "驱动器系列号:" & .SerialNumber & vbCrLf _
            & "驱动器大小:" & FormatNumber(.TotalSize / 1024, 0) & "KB" & vbCrLf _
            & "驱动器剩余空间:" & FormatNumber(.FreeSpace / 1024, 0) & "KB" & vbCrLf
        End With
    MsgBox MyStr, 64, "驱动器信息"
    Set MyFile = Nothing
    Set MyDisk = Nothing
End Sub

