Attribute VB_Name = "OB1Step2"
Sub OB1Step2()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim colNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim formulaRange As Range
    Dim lookupCell As Range
    Dim yesterday As String  ' Declare yesterday as a String to use it in Excel formulas
    Dim colNumFromCell As Integer     ' Column number obtained from cell DE1 (e.g., 63 or 1)

    ' Turn off screen updating and automatic calculation to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set references to the Data and Report sheets
    Set ws = ThisWorkbook.Sheets("Data")
    Set reportWs = ThisWorkbook.Sheets("Report")
    
    ' Get the numerical value from cell DE1
    ' This value will be used as the number of days to subtract from TODAY()
    colNumFromCell = ws.Range("DE1").Value

    ' Enable AutoFilter if it is not already on
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Filter only the rows where the CR column (Column 96) is blank
    colNum = 96
    ws.Range("A1").AutoFilter Field:=colNum, Criteria1:="="

    ' Find the last row with data in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Get yesterday's date in dd/mm/yyyy format to use in the Excel formula
    yesterday = "TODAY()-" & colNumFromCell   ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    ' Insert IF formulas into column DA (Column 105)
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 105).Formula = "=IF(S" & i & "="""",CS" & i & "," & yesterday & ")"
        End If
    Next i

    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("DA2:DA" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Copy values from DA to CS only for visible rows
    For Each cell In ws.Range("DA2:DA" & lastRow)
        If Not cell.EntireRow.Hidden Then
            ws.Cells(cell.Row, 97).Value = cell.Value
        End If
    Next cell
    
    ' ____________________________________________________________________________________________
    
    ' Insert IF formulas into column CO (Column 93) , CR (Column 96)
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 93).Formula = "=IF(AND(CS" & i & "=" & yesterday & ",S" & i & "=""""),""Trf"","""")"
            ws.Cells(i, 96).Formula = "=IF(OR(S" & i & "=""Yes"",CO" & i & "=""Trf"",AND(CS" & i & "=" & yesterday & ",P" & i & "=1)),""Stop"","""")"
        End If
    Next i

    ' Convert formulas to values
    Set formulaRange = ws.Range("CO2:CO" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Convert formulas to values
    Set formulaRange = ws.Range("CR2:CR" & lastRow)
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
