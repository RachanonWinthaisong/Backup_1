Attribute VB_Name = "Module21"
Sub SaveSheetNT()
    Dim ws As Worksheet
    Dim wbNew As Workbook
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    
    ' 1. Define the destination folder path (Must end with \)
    folderPath = "Z:\8.Collection\`work_MIS\Terminate\Backup\NT File\"
    
    ' 2. Define the file name with current date (Format: NT_dd.mm.yyyy)
    fileName = "NT_" & Format(Date, "dd.mm.yyyy") & ".xlsx"
    fullPath = folderPath & fileName
    
    ' 3. Attempt to set the worksheet "NT" and handle error if not found
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("NT")
    On Error GoTo 0
    
    ' Check if worksheet exists
    If ws Is Nothing Then
        MsgBox "Worksheet 'NT' not found!", vbCritical
        Exit Sub
    End If
    
    ' 4. Copy the "NT" sheet to a new Workbook
    ws.Copy
    Set wbNew = ActiveWorkbook
    
    ' 5. Save the new workbook and close it
    ' Disable alerts to overwrite existing files without confirmation
    Application.DisplayAlerts = False
    
    ' Save as standard Excel workbook (.xlsx)
    wbNew.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    
    ' Close the new workbook without saving further changes
    wbNew.Close SaveChanges:=False
    
    ' Re-enable alerts
    Application.DisplayAlerts = True
    
    ' 6. Display success message
    MsgBox "เซฟไฟล์เรียบร้อย: " & fullPath, vbInformation
End Sub
