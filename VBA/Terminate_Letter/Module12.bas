Attribute VB_Name = "Module12"
Sub ChangeCol_P_12()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim visibleRange As Range
    Dim cell As Range ' Declare 'cell' as a Range object for the loop
    
    ' Optimize performance by turning off screen updates and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True ' Set to True to show status bar messages, use False to hide

    ' Set reference to the "assign repo" worksheet
    Set ws = ThisWorkbook.Sheets("assign repo")

    ' *** Modification: Assign a value to lastRow ***
    ' Find the last used row with data in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' *** Start of the new code section to populate Column P ***
    
    ' Use SpecialCells(xlCellTypeVisible) to get a range object containing only visible cells in Column A
    On Error Resume Next ' Use error handling in case no cells are visible (e.g., filter hides all data rows)
    ' Adjust the range correctly, using the calculated lastRow
    Set visibleRange = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0 ' Resume normal error handling

    ' Check if there are any visible cells found
    If Not visibleRange Is Nothing Then
        ' Loop through each visible cell and set the corresponding cell in Column P to "repossessed"
        For Each cell In visibleRange
            ' Column P (16)
            ws.Cells(cell.Row, 16).Value = "without repossession"
        Next cell
    End If
    ' *** End of the new code section ***

    ' Restore original application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
