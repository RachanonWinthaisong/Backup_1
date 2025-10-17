Attribute VB_Name = "Module3456"
Sub B3up()

    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim colNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim formulaRange As Range
    Dim lookupCell As Range
    Dim colNumFromCell As Integer   ' Column number obtained from cell CT1 (e.g., 41)
    Dim colLetterFromCell As String ' Column letter obtained from cell CU1 (e.g., "AO")

    ' Turn off screen updating and automatic calculation to speed up the code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set the Data and Report sheets
    Set ws = ThisWorkbook.Sheets("Data")
    Set reportWs = ThisWorkbook.Sheets("Report")
    
    ' Get column number and letter from specific cells on the Data sheet
    colNumFromCell = ws.Range("CT1").Value     ' Example: DC1 = 41 (for column AO)
    colLetterFromCell = ws.Range("CU1").Value  ' Example: DD1 = "AO"

    ' Turn on AutoFilter if it is not already on
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through only the visible rows
    For i = 10 To lastRow
        If Not ws.Rows(i).Hidden Then
        
            ' Insert formula
            ws.Cells(i, 14).Formula = "=XLOOKUP(B" & i & ",Report!$A$2:$A$20000,Report!$D$2:$D$20000,,,1)"
            ws.Cells(i, 24).Formula = "=IF(ISERROR(XLOOKUP(B" & i & ",Report!$M$2:$M$2000,Report!$M$2:$M$2000,,,1))=TRUE,NA(),""Repo"")"
            ws.Cells(i, colNumFromCell).Formula = "=XLOOKUP(B" & i & ",Report!$H$2:$H$20000,Report!$I$2:$I$20000,,,1)"  ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            ws.Cells(i, 71).Formula = "=XLOOKUP(B" & i & ",Report!$A$2:$A$20000,Report!$B$2:$B$20000,,,1)"
        End If
    Next i


    ' Convert formulas to values
    Set formulaRange = ws.Range("N10:N" & lastRow)
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
    Set formulaRange = ws.Range("X10:X" & lastRow)
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
    Set formulaRange = ws.Range(colLetterFromCell & "10:" & colLetterFromCell & lastRow)   ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
    Set formulaRange = ws.Range("BS10:BS" & lastRow)
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

' ________________________________________________________________________________________________________________'

    ' Insert IF formulas in the specified columns  *** Close_ ***
    For i = 10 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 93).Formula = "=IF(BP" & i & "="""",XLOOKUP(B" & i & ",Report!$P$2:$P$2000,Report!$Q$2:$Q$2000,,,1),BP" & i & ")"
        End If
    Next i

    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("CO10:CO" & lastRow)
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

    ' Copy values from CO to BP only for visible rows
    For Each cell In ws.Range("CO10:CO" & lastRow)
        If Not cell.EntireRow.Hidden Then
            ws.Cells(cell.Row, 68).Value = cell.Value
        End If
    Next cell
' ________________________________________________________________________________________________________________'


' ________________________________________________________________________________________________________________'

    ' Insert IF formulas in the specified columns  ***  finish ***
    For i = 10 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 25).Formula = "=IF(AND(F" & i & ">=5,P" & i & "="""",BE" & i & ">0,BS" & i & ">=2),""N"", NA())"
        End If
    Next i
    
    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("Y10:Y" & lastRow)
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

' ________________________________________________________________________________________________________________'

' ________________________________________________________________________________________________________________'

    ' Insert IF formulas in the specified columns  ***  kept ***
    For i = 10 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 95).Formula = "=IF(P" & i & "="""",IFS(AND(O" & i & ">=0,Y" & i & "="""",BP" & i & "<>""Total Loss""),1,X" & i & "=""Repo"",1),P" & i & ")"
        End If
    Next i
    
    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("CQ10:CQ" & lastRow)
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

    ' Copy values from CQ to P only for visible rows
    For Each cell In ws.Range("CQ10:CQ" & lastRow)
        If Not cell.EntireRow.Hidden Then
            ws.Cells(cell.Row, 16).Value = cell.Value
        End If
    Next cell
' ________________________________________________________________________________________________________________'

' ________________________________________________________________________________________________________________'

    ' Insert IF formulas in the specified columns  *** Team kept2 ***
    For i = 10 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 76).Formula = "=IF(P" & i & "=1,BW" & i & ",NA())"
        End If
    Next i
    
    ' Force Excel to recalculate all formulas before converting them to values
    Application.Calculate

    ' Convert formulas to values
    Set formulaRange = ws.Range("BX10:BX" & lastRow)
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
' ________________________________________________________________________________________________________________'


    ' Turn on screen updating and automatic calculation again
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
