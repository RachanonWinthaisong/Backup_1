Attribute VB_Name = "OB1Step3"
Sub OB1Step3()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim formulaStr As String
    
    ' Set the "Data" worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Find the last row in column B that contains data
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Define the XLOOKUP formula with absolute references
    formulaStr = "=XLOOKUP(B2,Report!$N$2:$N$20000,Report!$O$2:$O$20000,,,1)"
    
    ' Insert the formula into column CU from row 2 to the last row
    ws.Range("CU2:CU" & lastRow).Formula = formulaStr
    
    ' Notify the user
    MsgBox "XLOOKUP formula has been inserted in column CU.", vbInformation, "Task Completed"
End Sub
