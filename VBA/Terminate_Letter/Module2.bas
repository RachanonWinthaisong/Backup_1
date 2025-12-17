Attribute VB_Name = "Module2"
Option Explicit

Sub Filter_ColumnF_ErrorNA_ThenDeleteVisible_1()
    ' Declares variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    ' Variable to hold the range of visible cells after filtering
    Dim visibleRange As Range
    Dim dataRange As Range

    ' Turn off screen updating for performance optimization
    Application.ScreenUpdating = False

    ' Reference the target sheet named "assign repo"
    Set ws = ThisWorkbook.Sheets("assign repo")

    ' Find the last used row in Column A and last used column in Row 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Define the entire data range including the header row
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Clear existing filters (if any) to ensure a clean slate
    If ws.AutoFilterMode Then
        On Error Resume Next
        If ws.FilterMode Then ws.ShowAllData
        On Error GoTo 0
    End If

    ' Apply AutoFilter capability to the data range
    dataRange.AutoFilter

    ' Apply filter on column F (Field 6) to show only #N/A errors
    dataRange.AutoFilter Field:=6, Criteria1:="=#N/A"

    ' --- New Steps: Delete Filtered Rows (Excluding Header) ---

    ' 1. Define the range of visible cells starting from Row 2 (excluding header A1)
    '    SpecialCells(xlCellTypeVisible) selects only the rows currently displayed by the filter
    On Error Resume Next ' Handle potential error if no #N/A rows are found (visibleRange remains Nothing)
    Set visibleRange = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0 ' Resume normal error handling

    ' 2. Check if any visible cells were found
    If Not visibleRange Is Nothing Then
        ' Delete the entire rows of the visible cells selected
        visibleRange.EntireRow.Delete Shift:=xlUp
    Else
    End If
    
    ' 3. Clear the AutoFilter after deletion is complete
    If ws.AutoFilterMode Then
        ws.ShowAllData
    End If

    ' Restore screen updating
    Application.ScreenUpdating = True
End Sub
