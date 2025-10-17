Attribute VB_Name = "Step1"
Sub Step1()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim colNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim formulaRange As Range
    Dim lookupCell As Range
    Dim colNumFromCell As Integer   ' Column number obtained from cell EC1 (e.g., 69)
    Dim colLetterFromCell As String ' Column letter obtained from cell ED1 (e.g., "BQ")
    Dim colNum2 As Integer
    Dim colNum3 As Integer
    

    ' Turn off screen updating and automatic calculation to speed up the code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set the Data and Report sheets
    Set ws = ThisWorkbook.Sheets("Data")
    Set reportWs = ThisWorkbook.Sheets("Report")
    
    ' Get column number and letter from specific cells on the Data sheet
    colNumFromCell = ws.Range("EC1").Value     ' Example: EC1 = 69 (for column BQ)
    colLetterFromCell = ws.Range("ED1").Value  ' Example: ED1 = "BQ"

    ' Turn on AutoFilter if it is not already on
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Filter only the rows where DC (Column 107) is blank
    colNum = 107
    ws.Range("A1").AutoFilter Field:=colNum, Criteria1:="="

    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Loop through only the visible rows
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ' Copy value to Column N
            If ws.Cells(i, 13).Value <> "" Then
                ws.Cells(i, 14).Value = ws.Cells(i, 13).Value
            Else
                ws.Cells(i, 14).Value = ws.Cells(i, 8).Value
            End If

            ' Insert XLOOKUP formula
            ws.Cells(i, 13).Formula = "=XLOOKUP(B" & i & ",Report!$A$2:$A$20000,Report!$D$2:$D$20000,,,1)"
            ws.Cells(i, colNumFromCell).Formula = "=XLOOKUP(B" & i & ",Report!$H$2:$H$20000,Report!$I$2:$I$20000,,,1)"  ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ws.Cells(i, 23).Formula = "=IF(ISERROR(XLOOKUP(B" & i & ",Report!$S$2:$S$1000,Report!$S$2:$S$1000,,,1))=TRUE,NA(),""Defer"")"
        End If
    Next i

    ' Convert formulas to values
    Set formulaRange = ws.Range("M2:M" & lastRow)
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
    Set formulaRange = ws.Range(colLetterFromCell & "2:" & colLetterFromCell & lastRow)   ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
    Set formulaRange = ws.Range("W2:W" & lastRow)
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
            ws.Cells(i, 130).Formula = "=IF(R" & i & "="""",IF(O" & i & ">=0,1,NA()),R" & i & ")"
            ws.Cells(i, 19).Formula = "=IF(OR(M" & i & "=1,M" & i & "=""""),1,"""")"
            ws.Cells(i, 20).Formula = "=IF(AND(OR(M" & i & "=1,M" & i & "=""""),Q" & i & ">=3),1,"""")"
            ws.Cells(i, 24).Formula = "=IF(OR(M" & i & "=1,M" & i & "=""""),""Yes"","""")"
        End If
    Next i

    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("DZ2:DZ" & lastRow)
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
    
        Set formulaRange = ws.Range("S2:S" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If
    
            Set formulaRange = ws.Range("T2:T" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If
    
            Set formulaRange = ws.Range("X2:X" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Copy values from DZ to R only for visible rows
    For Each cell In ws.Range("DZ2:DZ" & lastRow)
        If Not cell.EntireRow.Hidden Then
            ws.Cells(cell.Row, 18).Value = cell.Value
        End If
    Next cell
    
    ' Filter only the rows where R (Column 18) is 1
    colNum2 = 18
    ws.Range("A1").AutoFilter Field:=colNum2, Criteria1:="1"
    
        ' Filter only the rows where CK (Column 89) is 0.00
    colNum3 = 89
    ws.Range("A1").AutoFilter Field:=colNum3, Criteria1:="0.00"

    ' Turn on screen updating and automatic calculation again
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
