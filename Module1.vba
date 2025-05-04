Option Explicit

Public Sub ExportEstimateTableToExcel()
    Const FILE_PATH As String = "C:\yourspreadsheetpath.xlsm"
    
    Dim xlApp        As Object
    Dim xlBook       As Object
    Dim ws           As Object
    Dim tbl          As Table
    Dim cc           As ContentControl
    Dim r            As Long, c As Long
    Dim customerName As String, custDate As String
    Dim custAddress  As String, custCity As String
    Dim safeName     As String, finalName As String
    Dim outputFolder As String, outputPath As String
    
    '— Read CCs from the active document —
    For Each cc In ActiveDocument.ContentControls
        Select Case LCase(cc.Title)
            Case "customername": customerName = Trim(cc.Range.Text)
            Case "date":         custDate = Trim(cc.Range.Text)
            Case "address":      custAddress = Trim(cc.Range.Text)
            Case "city":         custCity = Trim(cc.Range.Text)
        End Select
    Next cc
    If Len(customerName) = 0 Then
        MsgBox "Missing Customer Name.", vbCritical
        Exit Sub
    End If
    
    '— Save a copy as .docm —
    outputFolder = "C:\yourdocumentfolder"
    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder
    safeName = SafeSheetName(customerName)
    outputPath = outputFolder & safeName & ".docm"
    ActiveDocument.SaveAs2 _
      FileName:=outputPath, _
      FileFormat:=wdFormatXMLDocumentMacroEnabled

    '— Export into Excel —
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(FILE_PATH, ReadOnly:=False)

    xlApp.DisplayAlerts = False
      On Error Resume Next
        xlBook.Sheets(safeName).Delete
        xlBook.Names(safeName).Delete
        Err.Clear
      On Error GoTo 0
    xlApp.DisplayAlerts = True

    Set ws = xlBook.Sheets.Add(After:=xlBook.Sheets(xlBook.Sheets.Count))
    On Error Resume Next
      ws.Name = safeName
      If Err.Number <> 0 Then
          Err.Clear
          finalName = Left(safeName & "_" & Format(Now, "yyyymmdd_hhnnss"), 31)
          ws.Name = finalName
      Else
          finalName = safeName
      End If
    On Error GoTo 0

    '— Write headers —
    ws.Cells(1, 1).Value = "Date: " & custDate
    ws.Cells(1, 2).Value = "Customer: " & customerName
    ws.Cells(1, 3).Value = "Address: " & custAddress
    ws.Cells(1, 4).Value = "City: " & custCity

    '— Copy the first table —
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "No table found.", vbExclamation
        GoTo WrapUp
    End If
    Set tbl = ActiveDocument.Tables(1)
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            ws.Cells(r + 1, c).Value = Clean(tbl.Cell(r, c).Range.Text)
        Next c
    Next r
    ws.UsedRange.Columns.AutoFit

WrapUp:
    xlBook.Save
    xlBook.Close False
    xlApp.Quit

    MsgBox "Saved copy to " & outputPath & vbCrLf & _
           "Exported to sheet: " & finalName, vbInformation
End Sub

'————————————————————
' Make Clean PUBLIC so it’s visible everywhere
'————————————————————
Public Function Clean(txt As String) As String
    txt = Replace(txt, Chr(13) & Chr(7), "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(7), "")
    Clean = Trim(txt)
End Function

'————————————————————
' Make SafeSheetName PUBLIC too
'————————————————————
Public Function SafeSheetName(rawName As String) As String
    Dim badChars As Variant, i As Long, s As String
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    s = rawName
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "")
    Next i
    If Len(s) > 31 Then s = Left(s, 31)
    SafeSheetName = s
End Function





