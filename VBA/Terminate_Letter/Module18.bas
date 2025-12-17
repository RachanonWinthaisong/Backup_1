Attribute VB_Name = "Module18"
Option Explicit

Sub Filter_ColumnE_18()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' 1. Set worksheet and handle potential naming errors
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("㺵ͺ�Ѻ")
    On Error GoTo 0
    
    ' Check if worksheet exists
    If ws Is Nothing Then
        MsgBox "Worksheet not found. Please check the sheet name.", vbCritical
        Exit Sub
    End If

    ' 2. Remove existing filters to prevent conflicts
    ws.AutoFilterMode = False

    ' 3. Find the last row based on column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 4. Define range from A to E and apply filter
    ' Field:=5 refers to Column E. Criteria "<>#N/A" excludes N/A errors.
    With ws.Range("$A$1:$E" & lastRow)
        .AutoFilter Field:=5, Criteria1:="<>#N/A"
    End With

    ' Restore screen updating
    Application.ScreenUpdating = True
End Sub
