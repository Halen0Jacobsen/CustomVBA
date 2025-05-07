Sub InsertKeywordNotices()
    Dim tbl As Table
    Dim cel As Cell
    Dim keyword As Variant
    Dim foundKeywords As Collection
    Dim keywordMessages As Object
    Dim messageText As String
    Dim shp As Shape
    Dim i As Integer
    Dim doc As Document
    Dim anchorRange As Range

    Set doc = ActiveDocument

    ' Define warranty messages
    Set keywordMessages = CreateObject("Scripting.Dictionary")
    keywordMessages.Add "plant", "Plant Warranty: All plant materials are warranted for 90 days from installation, provided proper care is maintained."
    keywordMessages.Add "cement", "Concrete Warranty: Cement work is covered for 1 year against cracks caused by workmanship (not ground movement or weather)."
    keywordMessages.Add "concrete", "Concrete Warranty: Cement work is covered for 1 year against cracks caused by workmanship (not ground movement or weather)."
    keywordMessages.Add "Flatwork", "Concrete Warranty: Cement work is covered for 1 year against cracks caused by workmanship (not ground movement or weather)."
    
    

    Set foundKeywords = New Collection

    ' Look for keywords in tables
    For Each tbl In doc.Tables
        For Each cel In tbl.Range.Cells
            For Each keyword In keywordMessages.Keys
                If InStr(1, LCase(cel.Range.Text), LCase(keyword)) > 0 Then
                    On Error Resume Next
                    foundKeywords.Add keyword, CStr(keyword) ' avoid duplicates
                    On Error GoTo 0
                End If
            Next keyword
        Next cel
    Next tbl

    ' If no keywords found, exit
    If foundKeywords.Count = 0 Then
        MsgBox "No keywords found in tables.", vbInformation
        Exit Sub
    End If

    ' Get anchor point on the last page
    Set anchorRange = doc.Range(doc.Content.End - 2, doc.Content.End - 1)

    ' Position boxes side-by-side just above the footer
    i = 0
    For Each keyword In foundKeywords
        messageText = keywordMessages(keyword)

        Set shp = doc.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=InchesToPoints(1 + i * 3.2), _
            Top:=doc.PageSetup.PageHeight - doc.PageSetup.BottomMargin - InchesToPoints(1), _
            Width:=InchesToPoints(3), _
            Height:=InchesToPoints(0.5), _
            Anchor:=anchorRange)

        With shp
            .TextFrame.TextRange.Text = messageText
            .TextFrame.TextRange.Font.Size = 8
            .TextFrame.TextRange.Font.Name = "Cambria"
            .Line.Visible = msoFalse
            .Fill.ForeColor.RGB = RGB(253, 245, 230)                        ' RGB Color for Text box
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            .LockAnchor = True
            .WrapFormat.Type = wdWrapNone
        End With

        i = i + 1
    Next keyword

    MsgBox "Warranty notice(s) added above the footer.", vbInformation
End Sub




