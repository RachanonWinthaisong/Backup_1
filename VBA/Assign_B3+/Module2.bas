Attribute VB_Name = "Module2"
Sub SendAutoEmails()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim subject As String
    Dim body As String
    Dim attachments As String
    Dim recipient As String
    Dim ccList As String
    Dim fixedAttachmentPath As String

    ' Set worksheet to Auto email sheet
    Set ws = ThisWorkbook.Sheets("Auto email")

    ' Create Outlook Application object
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Get the last row of data in the Auto email sheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Loop through all rows in the Auto email sheet and send emails
    For i = 2 To lastRow ' Assuming data starts at row 2
        ' Read values from the sheet
        subject = ws.Cells(i, 2).Value ' Column B: Subject
        body = ws.Cells(i, 3).Value ' Column C: Body
        attachments = ws.Cells(i, 4).Value ' Column D: Attachment path
        recipient = ws.Cells(i, 5).Value ' Column E: To
        ccList = ws.Cells(i, 6).Value ' Column F: CC

        ' Ensure the email recipient is not empty
        If recipient <> "" Then
            ' Fix the attachment path
            fixedAttachmentPath = Replace(attachments, "\", "\\")

            ' Check if the file exists
            If fixedAttachmentPath <> "" Then
                If Dir(fixedAttachmentPath) = "" Then
                    MsgBox "‰¡Ëæ∫‰ø≈Ï·π∫: " & fixedAttachmentPath & vbCrLf & "®–‰¡Ë Ëß‡¡≈„π·∂«∑’Ë " & i, vbExclamation
                    GoTo SkipEmail
                End If
            End If

            ' Create a new email item
            Set OutlookMail = OutlookApp.CreateItem(0)

            With OutlookMail
                .To = recipient
                If ccList <> "" Then .CC = ccList ' Add CC if provided
                .subject = subject
                .body = body

                If fixedAttachmentPath <> "" Then
                    .attachments.Add fixedAttachmentPath
                End If

                .Send
            End With

            Set OutlookMail = Nothing
        End If

SkipEmail:
    Next i

    Set OutlookApp = Nothing
End Sub

