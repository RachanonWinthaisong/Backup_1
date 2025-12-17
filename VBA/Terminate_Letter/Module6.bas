Attribute VB_Name = "Module6"
Option Explicit

Sub ClearFilter_6()
    ' Declares variables
    Dim ws As Worksheet

    ' Reference the target sheet named "assign repo"
    Set ws = ThisWorkbook.Sheets("assign repo")
    
    ' Clear the AutoFilter after deletion is complete
    If ws.AutoFilterMode Then
        ws.ShowAllData
    End If

    ' Restore screen updating
    Application.ScreenUpdating = True
End Sub
