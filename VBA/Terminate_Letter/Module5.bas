Attribute VB_Name = "Module5"
Sub Xlook_Repo_5()
    ' Declare variables
    Dim ws As Worksheet
    Dim reportWs As Worksheet ' Declared but not used in the current sub
    Dim lastRow As Long
    Dim cell As Range ' Added declaration for the 'cell' variable used in loops
    
    ' Optimize performance by turning off screen updates and automatic calculations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    ' Set reference to the "assign repo" worksheet
    Set ws = ThisWorkbook.Sheets("assign repo")

    ' Ensure AutoFilter is enabled on the sheet
    ' Note: This code enables autofilter but does not apply a filter condition.
    ' It assumes a filter is set elsewhere or applied manually by the user before running this macro.
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter

    ' Find the last used row in Column A to define the data range
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Insert formulas into column R (Column 18) ,  T (Column 20) , X (Column 24)
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            ws.Cells(i, 18).Formula = "=XLOOKUP(A" & i & ",AllQuery!$A$2:$A$1000,AllQuery!$C$2:$C$1000,,,1)"
            ws.Cells(i, 20).Formula = "=XLOOKUP(A" & i & ",AllQuery!$A$2:$A$1000,AllQuery!$D$2:$D$1000,,,1)"
            ws.Cells(i, 24).Formula = "=XLOOKUP(A" & i & ",AllQuery!$A$2:$A$1000,AllQuery!$E$2:$E$1000,,,1)"
        End If
    Next i

    ' Convert formulas to values
    Set formulaRange = ws.Range("R2:R" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If
    
    ' Convert formulas to values
    Set formulaRange = ws.Range("T2:T" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If
    
    ' Convert formulas to values
    Set formulaRange = ws.Range("X2:X" & lastRow)
    If Not formulaRange Is Nothing Then
        For Each lookupCell In formulaRange
            If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        Next lookupCell
    End If

    ' Restore original application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
End Sub
