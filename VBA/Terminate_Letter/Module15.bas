Attribute VB_Name = "Module15"
Sub DeleteVisibleColumnsV_W_15()
    ' // This subroutine deletes only visible data in Columns V and W (excluding the header row)
    ' // It respects active filters and does not affect hidden rows.

    Dim ws As Worksheet
    Dim lastRowV As Long
    Dim lastRowW As Long
    Dim maxLastRow As Long
    Dim targetRange As Range
    
    ' Reference the target sheet named "assign repo"
    Set ws = ThisWorkbook.Sheets("assign repo")
    
    ' Find the last used row in both Column V and Column W to ensure all data is covered
    lastRowV = ws.Cells(ws.Rows.Count, "V").End(xlUp).Row
    lastRowW = ws.Cells(ws.Rows.Count, "W").End(xlUp).Row
    
    ' Determine the furthest row containing data between the two columns
    If lastRowV > lastRowW Then maxLastRow = lastRowV Else maxLastRow = lastRowW
    
    ' Check if there is data to delete starting from Row 2
    If maxLastRow >= 2 Then
        ' Define the range covering both Column V and Column W
        Set targetRange = ws.Range("V2:W" & maxLastRow)
        
        ' Ignore error if no visible cells are found within the filtered range
        On Error Resume Next
        
        ' Clear contents of ONLY the visible cells within the defined range
        targetRange.SpecialCells(xlCellTypeVisible).ClearContents
        
        ' Reset error handling
        On Error GoTo 0
    End If
End Sub
