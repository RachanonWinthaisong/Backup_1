Attribute VB_Name = "Module4"
Sub Copy_LookupRepo_4()
    ' Declare variables
    Dim ws As Worksheet
    Dim reportWs As Worksheet ' Declared but not used in the current sub
    Dim lastRow As Long
    Dim cell As Range ' Added declaration for the 'cell' variable used in loops
    
    ' Optimize performance by turning off screen updates and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set reference to the "assign repo" worksheet
    Set ws = ThisWorkbook.Sheets("assign repo")

    ' Ensure AutoFilter is enabled on the sheet
    ' Note: This code enables autofilter but does not apply a filter condition.
    ' It assumes a filter is set elsewhere or applied manually by the user before running this macro.
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Find the last used row in Column A to define the data range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through all cells in Column U (starting from row 2) and copy their values
    ' to Column Q only if the entire row is visible (not hidden by a filter)
    For Each cell In ws.Range("U2:U" & lastRow)
        If Not cell.EntireRow.Hidden Then
            ' Column Q (17)
            ws.Cells(cell.Row, 17).Value = cell.Value
        End If
    Next cell

    ' *** Start of the new code section to populate Column P ***
    Dim visibleRange As Range
    
    ' Use SpecialCells(xlCellTypeVisible) to get a range object containing only visible cells in Column A
    On Error Resume Next ' Use error handling in case no cells are visible (e.g., filter hides all data rows)
    Set visibleRange = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0 ' Resume normal error handling

    ' Check if there are any visible cells found
    If Not visibleRange Is Nothing Then
        ' Loop through each visible cell and set the corresponding cell in Column P to "repossessed"
        For Each cell In visibleRange
            ' Column P (16)
            ws.Cells(cell.Row, 16).Value = "repossessed"
        Next cell
    End If
    ' *** End of the new code section ***

    ' Restore original application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
