Attribute VB_Name = "Module1"
Sub Export_DetailSheets_WithLabel()
    Dim ws As Worksheet
    Dim wbNew As Workbook
    Dim folderPath As String
    Dim a1Value As String
    Dim shortName As String
    Dim todayStr As String
    Dim safeName As String
    Dim fileCode As String
    Dim filePassword As String

    todayStr = Format(Date, "yyyymmdd")
    filePassword = "1234"  ' ????????????

    ' Choose folder for saving files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choose folder to save files"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1) & "\"
    End With

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 6) = "Detail" Then
            ' Clean spaces from A1 value
            a1Value = CleanSpaces(ws.Range("A1").Value)

            ' Get last 3 characters
            If Len(a1Value) >= 3 Then
                shortName = Right(a1Value, 3)
            ElseIf Len(a1Value) > 0 Then
                shortName = a1Value
            Else
                shortName = ws.Name
            End If

            ' Clean invalid filename characters
            safeName = shortName
            safeName = Replace(safeName, "\", "")
            safeName = Replace(safeName, "/", "")
            safeName = Replace(safeName, ":", "")
            safeName = Replace(safeName, "*", "")
            safeName = Replace(safeName, "?", "")
            safeName = Replace(safeName, """", "")
            safeName = Replace(safeName, "<", "")
            safeName = Replace(safeName, ">", "")
            safeName = Replace(safeName, "|", "")

            ' Copy the sheet to a new workbook
            ws.Copy
            Set wbNew = ActiveWorkbook

            ' Insert label in cell A1
            wbNew.Sheets(1).Range("A1").Value = "Nissan Confidential C"

            ' Save the workbook with the additional file code (1234) and password
            wbNew.SaveAs Filename:=folderPath & safeName & "_" & todayStr & ".xlsx", _
                FileFormat:=xlOpenXMLWorkbook, Password:=filePassword
            wbNew.Close SaveChanges:=False
        End If
    Next ws

    MsgBox "Done! Exported all 'Detail' sheets with label in cell A1 and password protection.", vbInformation
End Sub

' Function to clean extra spaces
Function CleanSpaces(text As String) As String
    Dim result As String
    result = text

    result = Replace(result, Chr(160), "") ' Non-breaking space
    result = Replace(result, Chr(9), "")   ' Tab
    result = Replace(result, Chr(10), "")  ' Line feed
    result = Replace(result, Chr(13), "")  ' Carriage return
    result = Replace(result, " ", "")      ' Normal space

    CleanSpaces = result
End Function

