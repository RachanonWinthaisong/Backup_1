Attribute VB_Name = "OB1Step1"
Sub OB1Step1()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim colNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim formulaRange As Range
    Dim lookupCell As Range
    Dim colNumFromCell As Integer   ' Column number obtained from cell DC1 (e.g., 63)
    Dim colLetterFromCell As String ' Column letter obtained from cell DD1 (e.g., "BK")
    Dim colNum2 As Integer          ' Variable for the second filter column (Column P, 16)
    Dim colNum3 As Integer          ' Variable for the third filter column (Column CF, 84)

    ' Turn off screen updating and automatic calculation to speed up the code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set the Data and Report sheets
    Set ws = ThisWorkbook.Sheets("Data")
    Set reportWs = ThisWorkbook.Sheets("Report")
    
    ' Get column number and letter from specific cells on the Data sheet
    colNumFromCell = ws.Range("DC1").Value     ' Example: DC1 = 63 (for column BK)
    colLetterFromCell = ws.Range("DD1").Value  ' Example: DD1 = "BK"

    ' Turn on AutoFilter if it is not already on
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Filter only the rows where CR (Column 96) is blank
    colNum = 96
    ws.Range("A1").AutoFilter Field:=colNum, Criteria1:="="

    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Loop through only the visible rows
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then

            ' Insert XLOOKUP formula
            ws.Cells(i, 14).Formula = "=XLOOKUP(B" & i & ",Report!$A$2:$A$20000,Report!$D$2:$D$20000,,,1)"
            ws.Cells(i, colNumFromCell).Formula = "=XLOOKUP(B" & i & ",Report!$H2:$H$20000,Report!$I$2:$I$20000,,,1)"  ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        End If
    Next i

    ' Convert formulas to values
    Set formulaRange = ws.Range("N2:N" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then
                lookupCell.Value = lookupCell.Value
            End If
            ' Convert #N/A to blank
            If IsError(lookupCell.Value) Then
                lookupCell.Value = ""
            End If
        Next lookupCell
    End If
    
    ' Convert formulas to values
    Set formulaRange = ws.Range(colLetterFromCell & "2:" & colLetterFromCell & lastRow)   ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then
                lookupCell.Value = lookupCell.Value
            End If
            ' Convert #N/A to blank
            If IsError(lookupCell.Value) Then
                lookupCell.Value = ""
            End If
        Next lookupCell
    End If
    
    ' _________________________________________________________________________________________________'

    ' Insert IF formulas in the specified columns
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 16).Formula = "=IF(OR(N" & i & "=1,N" & i & "=""""),1,"""")"
            ws.Cells(i, 19).Formula = "=IF(OR(N" & i & "=1,N" & i & "=""""),""Yes"","""")"
        End If
    Next i

    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("P2:P" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If
    
        Set formulaRange = ws.Range("S2:S" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If
    
    ' Filter only the rows where P (Column 16) is 1
    colNum2 = 16
    ws.Range("A1").AutoFilter Field:=colNum2, Criteria1:="1"
    
        ' Filter only the rows where CF (Column 84) is 0.00
    colNum3 = 84
    ws.Range("A1").AutoFilter Field:=colNum3, Criteria1:="0.00"

    ' Turn on screen updating and automatic calculation again
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
