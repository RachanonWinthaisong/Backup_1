Sub UpdatePivotTable5In8p3k()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim lastRowDate As Long, lastRowCreation As Long
    Dim latestDay As String, latestCreation As String

    ' Access the data sheet
    Set wsData = ThisWorkbook.Sheets("data")

    ' Get the latest Date from column L
    lastRowDate = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
    latestDay = CStr(wsData.Cells(lastRowDate, "L").Value)

    ' Get the latest Creation date from column D
    lastRowCreation = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row
    latestCreation = Format(wsData.Cells(lastRowCreation, "D").Value, "dd/mm/yyyy")

    ' Access the pivot sheet and PivotTable5
    Set wsPivot = ThisWorkbook.Sheets("8.p3k")
    Set pt = wsPivot.PivotTables("PivotTable5")

    ' Update filters
    With pt.PivotFields("Date")
        .ClearAllFilters
        .CurrentPage = latestDay
    End With

    With pt.PivotFields("Creation")
        .ClearAllFilters
        .CurrentPage = latestCreation
    End With


    ' Refresh the pivot table
    pt.RefreshTable
    
    ' Copy data from PivotTable5 excluding header
    Dim rngSource As Range
    Dim rngTarget As Range
    Dim headerRow As Long

    ' Identify the range of the PivotTable
    With pt.TableRange1
        headerRow = .Row
        Set rngSource = wsPivot.Range(wsPivot.Cells(headerRow + 1, .Column), wsPivot.Cells(.Rows.Count + headerRow - 1, .Columns(.Columns.Count).Column))
    End With

    ' Set target range in "9.Review3000"
    Set rngTarget = ThisWorkbook.Sheets("9.Review3000").Range("A3")

    ' Clear existing data in target range
    ThisWorkbook.Sheets("9.Review3000").Range("A3:E200").ClearContents

    ' Copy and paste values only
    rngSource.Copy
    rngTarget.PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

End Sub
