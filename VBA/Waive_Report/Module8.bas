Attribute VB_Name = "Module8"
Sub FillInColumnD_VisibleOnly_Corrected1()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fillValue As Variant
    Dim sourceRow As Long

    Set ws = ThisWorkbook.Sheets("7.MONTH.AC")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For sourceRow = 2 To lastRow
        If Not ws.Rows(sourceRow).Hidden Then
            fillValue = ws.Cells(sourceRow, "D").Value
            Exit For
        End If
    Next sourceRow

    If sourceRow > lastRow Then
        MsgBox "?????????????????????????????? D2:D" & lastRow
        Exit Sub
    End If

    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, "D").Value = fillValue
        End If
    Next i

End Sub
