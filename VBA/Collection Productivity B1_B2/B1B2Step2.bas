Attribute VB_Name = "B1B2Step2"
Sub B1B2Step2()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim colNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim formulaRange As Range
    Dim lookupCell As Range
    Dim yesterday As String  ' Declare yesterday as a String to use it in Excel formulas
    Dim colNumFromCell As Integer     ' Column number obtained from cell EE1 (e.g., 1)

    ' Turn off screen updating and automatic calculation to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set references to the Data and Report sheets
    Set ws = ThisWorkbook.Sheets("Data")
    Set reportWs = ThisWorkbook.Sheets("Report")
    
    ' Get the numerical value from cell EE1
    ' This value will be used as the number of days to subtract from TODAY()
    colNumFromCell = ws.Range("EE1").Value

    ' Enable AutoFilter if it is not already on
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Filter only the rows where the DC column (Column 107) is blank
    colNum = 107
    ws.Range("A1").AutoFilter Field:=colNum, Criteria1:="="

    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Get yesterday's date in dd/mm/yyyy format to use in the Excel formula
    yesterday = "TODAY()-" & colNumFromCell   ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' Insert IF formulas into column DX (Column 128)
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 128).Formula = "=IF(X" & i & "="""",DB" & i & "," & yesterday & ")"
        End If
    Next i

    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("DX2:DX" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Copy values from DX to DB only for visible rows
    For Each cell In ws.Range("DX2:DX" & lastRow)
        If Not cell.EntireRow.Hidden Then
            ws.Cells(cell.Row, 106).Value = cell.Value
        End If
    Next cell
    
    ' ____________________________________________________________________________________________
    
    ' Insert IF formulas into column CX (Column 102) , DC (Column 107)
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 102).Formula = "=IF(OR(AND(DB" & i & "=" & yesterday & ",F" & i & "=1,R" & i & "=""""),AND(DB" & i & "=" & yesterday & ",M" & i & ">=59,R" & i & "="""")),""Trf"","""")"
            ws.Cells(i, 107).Formula = "=IF(OR(X" & i & "=""Yes"",CX" & i & "=""Trf"",AND(DB" & i & "=" & yesterday & ",R" & i & "=1)),""Stop"","""")"
        End If
    Next i

    ' Convert formulas to values
    Set formulaRange = ws.Range("CX2:CX" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Convert formulas to values
    Set formulaRange = ws.Range("DC2:DC" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Turn screen updating and automatic calculation back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
