Attribute VB_Name = "Module3"
Option Explicit

Sub Filter_ColumnU_Q_3()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    ' Turn off screen updating for performance
    Application.ScreenUpdating = False

    ' Reference the target sheet
    Set ws = ThisWorkbook.Sheets("assign repo")

    ' Find last used row and column to define the filter range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Define the entire data range including the header row
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Clear existing filters (if any)
    If ws.AutoFilterMode Then
        On Error Resume Next
        If ws.FilterMode Then ws.ShowAllData
        On Error GoTo 0
    End If

    ' Ensure AutoFilter is applied to the header row
    dataRange.AutoFilter

    ' STEP 1: Apply filter on column U (Field:=21) to show criteria NOT equal to the text "#N/A"
    ' Criteria1:="<>" (Not Equal To)
    dataRange.AutoFilter Field:=21, Criteria1:="<>" & "#N/A"
    
        ' STEP 2: Apply a secondary filter on column Q (Field:=17) for Blank values
    ' "" (Empty string) is used to filter for blanks when the data is text/general format
    dataRange.AutoFilter Field:=17, Criteria1:="="

    ' Restore screen updating
    Application.ScreenUpdating = True
    
End Sub

