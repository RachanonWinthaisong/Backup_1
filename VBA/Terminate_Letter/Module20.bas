Attribute VB_Name = "Module20"
Option Explicit

Sub ClearFilterLast_20()
    ' Declares variables
    Dim ws As Worksheet

    ' Reference the target sheet named "assign repo"
    Set ws = ThisWorkbook.Sheets("㺵ͺ�Ѻ")
    
    ' Clear the AutoFilter after deletion is complete
    If ws.AutoFilterMode Then
        ws.ShowAllData
    End If

    ' Restore screen updating
    Application.ScreenUpdating = True
End Sub
