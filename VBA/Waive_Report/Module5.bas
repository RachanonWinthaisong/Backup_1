Attribute VB_Name = "Module5"
Sub FilterColumnE_NotDashOrZero()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Pivot")

    ' Find the last row in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Define the range starting from E7 to the last row
    Set rng = ws.Range("E7:E" & lastRow)

    ' Apply AutoFilter to exclude "-" and 0
    With ws
        .AutoFilterMode = False ' Clear any existing filters
        rng.AutoFilter Field:=1, Criteria1:="<>-", Operator:=xlAnd, Criteria2:="<>0"
    End With

End Sub
