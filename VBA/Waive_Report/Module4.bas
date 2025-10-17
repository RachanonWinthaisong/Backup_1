Attribute VB_Name = "Module4"
Sub FilterPivotTable1ByLatestDate()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim lastRowDate As Long
    Dim latestDay As String

    ' Access the data sheet
    Set wsData = ThisWorkbook.Sheets("data")

    ' Get the latest Date from column L
    lastRowDate = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
    latestDay = CStr(wsData.Cells(lastRowDate, "L").Value)

    ' Access the pivot sheet and PivotTable1
    Set wsPivot = ThisWorkbook.Sheets("Pivot")
    Set pt = wsPivot.PivotTables("PivotTable1")

    ' Update filter for Date field
    With pt.PivotFields("Date")
        .ClearAllFilters
        .CurrentPage = latestDay
    End With

    ' Refresh the pivot table
    pt.RefreshTable

End Sub
