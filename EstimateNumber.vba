Sub InsertEstimateNumber()
    Dim folderPath As String
    Dim filePath As String
    Dim fNum As Integer
    Dim estNum As Long
    Dim cc As ContentControl
    Dim shp As Shape
    Dim found As Boolean

    ' Tracker path
    folderPath = Environ("USERPROFILE") & "\Documents\EstimateTracker"
    filePath = folderPath & "\estimate_number.txt"

    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    If Dir(filePath) = "" Then
        fNum = FreeFile
        Open filePath For Output As #fNum
        Print #fNum, 1000
        Close #fNum
    End If

    ' Read current estimate number
    fNum = FreeFile
    Open filePath For Input As #fNum
    Input #fNum, estNum
    Close #fNum

    ' === Search document content controls ===
    found = False
    For Each cc In ActiveDocument.ContentControls
        If LCase(cc.Title) = "estimatenumber" Then
            cc.LockContents = False
            cc.Range.Text = estNum
            found = True
            Exit For
        End If
    Next cc

    ' === Also search inside shapes (text boxes etc.) ===
    If Not found Then
        For Each shp In ActiveDocument.Shapes
            If shp.Type = msoTextBox Then
                For Each cc In shp.TextFrame.TextRange.ContentControls
                    If LCase(cc.Title) = "estimatenumber" Then
                        cc.LockContents = False
                        cc.Range.Text = estNum
                        found = True
                        Exit For
                    End If
                Next cc
            End If
            If found Then Exit For
        Next shp
    End If

    If Not found Then
        MsgBox "Content Control titled 'EstimateNumber' not found.", vbExclamation
    End If

    ' Save incremented number
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, estNum + 1
    Close #fNum
End Sub
