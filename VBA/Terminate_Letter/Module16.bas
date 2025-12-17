Attribute VB_Name = "Module16"
Sub InsertToday_ColV_16()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim visibleRange As Range
    Dim cell As Range
    
    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set reference to the "assign repo" worksheet
    Set ws = ThisWorkbook.Sheets("assign repo")

    ' Find the last used row based on Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there is data below the header
    If lastRow >= 2 Then
        ' Get only visible cells in Column A (from row 2 to last row)
        On Error Resume Next
        Set visibleRange = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        ' Check if any visible cells were found
        If Not visibleRange Is Nothing Then
            For Each cell In visibleRange
                ' Reference Column V (Column 22) on the same row as the visible cell
                With ws.Cells(cell.Row, 22)
                    .Value = Date ' Insert current system date
                    .NumberFormat = "dd/mm/yyyy" ' Set display format to dd/mm/yyyy
                End With
            Next cell
        End If
    End If

    ' Restore original application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
