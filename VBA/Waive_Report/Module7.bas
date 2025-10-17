Attribute VB_Name = "Module7"
Sub FilterByReview3000()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("8.p3k")
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 12 To lastRow
        If Trim(ws.Cells(i, "F").Value) = "Review3000" Then
            ws.Rows(i).Hidden = False
        Else
            ws.Rows(i).Hidden = True
        End If
    Next i

    Application.ScreenUpdating = True
End Sub
