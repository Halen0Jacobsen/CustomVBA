Option Explicit

Public Sub CalculateTotalsInWord()
    Dim tbl As Table
    Dim r As Long
    Dim unitCost As Double, quantity As Double, lineTotal As Double
    Dim grandTotal As Double
    Dim cc As ContentControl

    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "No tables found.", vbExclamation
        Exit Sub
    End If

    Set tbl = ActiveDocument.Tables(1)

    grandTotal = 0

    For r = 2 To tbl.Rows.Count ' Skip header row
        With tbl.Rows(r)
            unitCost = Val(Clean(.Cells(2).Range.Text))
            quantity = Val(Clean(.Cells(3).Range.Text))
            lineTotal = unitCost * quantity
            .Cells(4).Range.Text = Format(lineTotal, "0.00")
            grandTotal = grandTotal + lineTotal
        End With
    Next r

    ' Write the grand total to a content control
    For Each cc In ActiveDocument.ContentControls
        If LCase(cc.Title) = "totalestimate" Then
            cc.Range.Text = Format(grandTotal, "$#,##0.00")
            Exit For
        End If
    Next cc

    MsgBox "Totals calculated and inserted.", vbInformation
End Sub

Private Function Clean(txt As String) As String
    txt = Replace(txt, Chr(13) & Chr(7), "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(7), "")
    Clean = Trim(txt)
End Function
