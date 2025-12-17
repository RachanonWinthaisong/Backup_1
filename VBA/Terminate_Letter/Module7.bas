Attribute VB_Name = "Module7"
Sub DeleteColumnU_7()
    ' // This Subroutine deletes all data in Column U (excluding the header row)

    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Reference the target sheet named "assign repo"
    Set ws = ThisWorkbook.Sheets("assign repo")
    
    ' Find the last used row in Column U (moving up from the bottom)
    lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    
    ' Check if there is data to delete (i.e., if lastRow is 2 or greater)
    If lastRow >= 2 Then
        ' Delete contents of the range from cell U2 down to the last used cell in U
        ws.Range("U2:U" & lastRow).ClearContents
        
        ' Note: Use .Clear if you want to remove formatting as well
        ' ws.Range("U2:U" & lastRow).Clear
    End If
    
    ' Optional: Restore screen updating if it was turned off elsewhere
    ' Application.ScreenUpdating = True
End Sub
