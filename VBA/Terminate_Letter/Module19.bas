Attribute VB_Name = "Module19"
Option Explicit

Sub DeleteRow_19()
    ' Declares variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    ' Variable to hold the range of visible cells after filtering
    Dim visibleRange As Range
    Dim dataRange As Range

    ' Turn off screen updating for performance optimization
    Application.ScreenUpdating = False

    ' Reference the target sheet named "㺵ͺ�Ѻ"
    Set ws = ThisWorkbook.Sheets("㺵ͺ�Ѻ")

    ' Find the last used row in Column A and last used column in Row 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Define the entire data range including the header row
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    'Delete Rows (Excluding Header) ---

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

    ' Restore screen updating
    Application.ScreenUpdating = True
End Sub
