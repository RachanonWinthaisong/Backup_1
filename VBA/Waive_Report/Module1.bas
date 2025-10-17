Attribute VB_Name = "Module1"
Sub UpdateAllPivotTablesIn2Pivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pt2 As PivotTable, pt3 As PivotTable
    Dim pt5 As PivotTable, pt8 As PivotTable
    Dim lastRowCreation As Long
    Dim lastRowDate As Long
    Dim latestCreation As String
    Dim latestDay As String

    ' Access the data sheet
    Set wsData = ThisWorkbook.Sheets("data")

    ' Get last row with data in column D (Creation)
    lastRowCreation = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row
    latestCreation = Format(wsData.Cells(lastRowCreation, "D").Value, "dd/mm/yyyy")

    ' Get last row with data in column L (Date)
    lastRowDate = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
    latestDay = CStr(wsData.Cells(lastRowDate, "L").Value)

    ' Access the sheet that contains all PivotTables
    Set wsPivot = ThisWorkbook.Sheets("2.pivot")
    Set pt2 = wsPivot.PivotTables("PivotTable2")
    Set pt3 = wsPivot.PivotTables("PivotTable3")
    Set pt5 = wsPivot.PivotTables("PivotTable5")
    Set pt8 = wsPivot.PivotTables("PivotTable8")

    ' Update PivotTable2
    With pt2.PivotFields("Date")
        .ClearAllFilters
        .CurrentPage = latestDay
    End With
    With pt2.PivotFields("Creation")
        .ClearAllFilters
        .CurrentPage = latestCreation
    End With

    ' Update PivotTable3
    With pt3.PivotFields("Date")
        .ClearAllFilters
        .CurrentPage = latestDay
    End With
    With pt3.PivotFields("Creation")
        .ClearAllFilters
        .CurrentPage = latestCreation
    End With

    ' Update PivotTable5 to show all items
    With pt5.PivotFields("Date")
        .ClearAllFilters
    End With
    With pt5.PivotFields("Creation")
        .ClearAllFilters
    End With

    ' Update PivotTable8 to show all items
    With pt8.PivotFields("Date")
        .ClearAllFilters
    End With
    With pt8.PivotFields("Creation")
        .ClearAllFilters
    End With

End Sub
