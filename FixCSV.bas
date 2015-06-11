Attribute VB_Name = "Module1"
Sub Main()
Attribute Main.VB_Description = "Import Amazon Data Stripping Out LF"
Attribute Main.VB_ProcData.VB_Invoke_Func = "I\n14"
    
    Dim fn As String
    Dim tempFile As String
    
    tempFile = "amazon_temp_file.txt"
    fn = GetFileName(ActiveWorkbook.Path)
    StripFile fn, tempFile
    'ImportFile tempFile
    Sheets("MageData").Select
    For Each myQT In ActiveSheet.QueryTables
        myQT.Refresh
    Next
    
End Sub

Function GetFileName(PathName) As String

sPath = PathName
'Path = sPatch

If sPath = vbNullString Then
        sPath = "the path to documents folder"
Else
        sPath = " alias """ & sPath & """"
End If
    sMacScript = "set sFile to (choose file of type ({" & _
        """com.microsoft.Excel.xls"", ""org.openxmlformats.spreadsheetml.sheet"",""public.comma-separated-values-text"", ""public.text"",""public.csv"",""org.openxmlformats.spreadsheetml.sheet.macroenabled""}) with prompt " & _
        """Select a file to import"" default location " & sPath & ") as string" _
        & vbLf & _
        "return sFile"
     'Debug.Print sMacScript
    GetFileName = MacScript(sMacScript)
End Function

Sub StripFile(Filename As String, Temp As String)
    
Dim Indata As String * 1
 
 If Filename <> "" And Temp <> "" Then

        Inside = False
        
        Open Filename For Binary As #1
        
        Open Temp For Binary Access Write As #2
        
        Do
            Get #1, , Indata
            'Debug.Print Indata
            If Asc(Indata) = 34 Then
                Inside = Not Inside
            End If
            If Not (Inside And (Asc(Indata) = 10 Or Asc(Indata) = 13)) Then
                Put #2, , Indata
            End If
        Loop While Not EOF(1)
        Close #1
        Close #2
        
        End If
    
End Sub

Sub ReImportFile(Filename)
Attribute ReImportFile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
' Only need this if data connection is broken?


    Sheets("MageData").Select
    Cells.ClearContents
    
    For Each myQT In ActiveSheet.QueryTables
        myQT.Delete
    Next
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & Filename _
        , Destination:=Range("A1"))
        .Name = "temp_file"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, 2, _
        2, 2, 2)
        .Refresh BackgroundQuery:=False
        .UseListObject = False
    End With
End Sub

Function Parse(ParseString)
        Select Case Asc(ParseString)
            Case 226 ' smart quote
                Parse = "'"
        
            Case Else
                Parse = ParseString
        End Select

End Function
