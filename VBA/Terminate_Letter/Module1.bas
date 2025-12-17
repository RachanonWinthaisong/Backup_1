Attribute VB_Name = "Module1"
Sub Lookup_Team_Repo_2()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim lastRow As Long
    Dim rngVisible As Range
    Dim c As Range

    ' Turn off screen updating and calculation for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set references to the sheets
    Set ws = ThisWorkbook.Sheets("assign repo")
    Set reportWs = ThisWorkbook.Sheets("AllQuery")

    ' Enable AutoFilter if it is not already enabled
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Find the last row in column B (adjust based on actual data)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Loop through only visible rows
    On Error Resume Next
    Set rngVisible = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not rngVisible Is Nothing Then
        For Each c In rngVisible
            ' Insert XLOOKUP formula for exact match
            ws.Cells(c.Row, 21).Formula = _
                "=XLOOKUP(A" & c.Row & ",AllQuery!$A$2:$A$1000,AllQuery!$B$2:$B$1000,,,1)"
        Next c

        ' Convert formulas to values
        With ws.Range("U2:U" & lastRow)
            .Value = .Value
        End With
    End If

    ' Turn screen updating and calculation back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
