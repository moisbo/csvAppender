Public LastDir As String
Public LastDirDocProp As DocumentProperty

Sub Init()
    On Error Resume Next
    Set LastDirDocProp = ActiveWorkbook.CustomDocumentProperties("ModelOutputDirectory")
    If Err.Number > 0 Then
        LastDir = ""
        ActiveWorkbook.CustomDocumentProperties.Add Name:="ModelOutputDirectory", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=LastDir
        Set LastDirDocProp = ActiveWorkbook.CustomDocumentProperties("ModelOutputDirectory")
    Else
        LastDir = LastDirDocProp.Value
    End If
End Sub

Sub LoadFile(ByVal control As IRibbonControl)
    LoadCSVForm.Show
End Sub

Sub DeleteSheets(ByVal control As IRibbonControl)
    Application.DisplayAlerts = False
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
         Worksheets(i).Cells.Delete
   Next i
    Application.DisplayAlerts = True
End Sub

Function SheetExists(SheetsObject As Sheets, Sheetname As String) As Boolean
    SheetExists = False
    For Each ws In SheetsObject
        If Sheetname = ws.Name Then
            SheetExists = True
            Exit For
        End If
    Next
End Function

Function FileNameToDate(FileName As String) As String
    FileName = Mid(FileName, InStrRev(FileName, "_") + 1)
    FileName = Left(FileName, InStr(FileName, ".") - 1)
    strYear = Mid(FileName, 5)
    strMonth = Mid(FileName, 3, 2)
    strDay = Left(FileName, 2)
    FileDate = DateSerial(Val(strYear), Val(strMonth), Val(strDay))
    FileNameToDate = Format(FileDate, "dd-mm-yyyy")
    Debug.Print FileNameToDate
End Function

Sub ImportFileToSheet(FileName As String, FilePath As String, ByRef DestWb As Workbook, Sheetname As String)
    Dim wbtemp As Workbook
    Dim ws As Worksheet

    If SheetExists(DestWb.Sheets, Sheetname) Then
        Set ws = DestWb.Sheets(Sheetname)
    Else
        Set ws = DestWb.Sheets.Add
        ws.Name = Sheetname
    End If
            
    Set DestCell = ws.Cells(Rows.Count, "B").End(xlUp).Offset(1)
    Set FirstCol = ws.Cells(Rows.Count, "A").End(xlUp).Offset(1)
    
    With ws.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=DestCell)
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
    End With
    
    'Debug.Print ws.UsedRange.Rows.Count
    'Debug.Print FirstCol.Row
    
    ws.Cells.Range("A" & FirstCol.Row & ":A" & ws.UsedRange.Rows.Count + 1).Value = FileNameToDate(FileName)
    ws.Columns("A").NumberFormat = "mm-dd-yy"
    
End Sub

Sub ImportData()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim FileName As String
    FileName = Dir(LastDir & "\ ")
    
    Do Until FileName = ""
        ImportFileToSheet FileName, LastDir & "\" & FileName, wb, "Sheet"
        Debug.Print FileName
        FileName = Dir()
    Loop
    
    'LastDirDocProp = LastDir 'Save last dir info for later

    LoadCSVForm.Hide
    
End Sub
