Attribute VB_Name = "Module2"
Sub UpdatePivotTablesIn5acPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsDaily As Worksheet
    Dim wsMonth As Worksheet
    Dim pt1 As PivotTable, pt2 As PivotTable
    Dim lastRowDate As Long
    Dim latestDay As String
    Dim pt1Full As Range, pt2Full As Range
    Dim pt1Data As Range, pt2Data As Range
    Dim rowCount1 As Long, rowCount2 As Long
    Dim colCount1 As Long, colCount2 As Long
    Dim cell As Range
    Dim isGrandTotalRow As Boolean
    Dim pasteRow As Long
    Dim r As Long

    ' Access the data sheet
    Set wsData = ThisWorkbook.Sheets("data")

    ' Get the last row with data in column L (Date)
    lastRowDate = wsData.Cells(wsData.Rows.Count, "L").End(xlUp).Row
    latestDay = CStr(wsData.Cells(lastRowDate, "L").Value)

    ' Access the pivot sheet and pivot tables
    Set wsPivot = ThisWorkbook.Sheets("5.ac.pivot")
    Set pt1 = wsPivot.PivotTables("PivotTable1")
    Set pt2 = wsPivot.PivotTables("PivotTable2")

    ' Update filters to show only the latest date
    With pt1.PivotFields("Date")
        .ClearAllFilters
        .CurrentPage = latestDay
    End With
    With pt2.PivotFields("Date")
        .ClearAllFilters
        .CurrentPage = latestDay
    End With

    ' Refresh pivot tables
    pt1.RefreshTable
    pt2.RefreshTable

    ' Access the destination sheets
    Set wsDaily = ThisWorkbook.Sheets("6.Daily.ac")
    Set wsMonth = ThisWorkbook.Sheets("7.MONTH.AC")

    ' --- Handle PivotTable2 for 6.Daily.ac ---
    Set pt2Full = pt2.TableRange1
    rowCount2 = pt2Full.Rows.Count
    colCount2 = pt2Full.Columns.Count

    If LCase(Trim(pt2Full.Cells(rowCount2, 1).Text)) = "grand total" Then
        rowCount2 = rowCount2 - 1
    End If

    If rowCount2 > 1 Then
        Set pt2Data = pt2Full.Offset(1, 0).Resize(rowCount2 - 1, colCount2)

        wsDaily.Range("A2:C1000").ClearContents
        wsDaily.Range("A2").Resize(pt2Data.Rows.Count, pt2Data.Columns.Count).Value = pt2Data.Value
    End If

    ' --- Handle PivotTable1 for 6.Daily.ac ---
    Set pt1Full = pt1.TableRange1
    rowCount1 = pt1Full.Rows.Count
    colCount1 = pt1Full.Columns.Count

    isGrandTotalRow = False
    For Each cell In pt1Full.Rows(rowCount1).Cells
        If LCase(Trim(cell.Text)) = "grand total" Then
            isGrandTotalRow = True
            Exit For
        End If
    Next cell

    If isGrandTotalRow Then
        rowCount1 = rowCount1 - 1
    End If

    If rowCount1 > 1 Then
        Set pt1Data = pt1Full.Offset(1, 0).Resize(rowCount1 - 1, colCount1)

        wsDaily.Range("J2:L1000").ClearContents
        wsDaily.Range("J2").Resize(pt1Data.Rows.Count, pt1Data.Columns.Count).Value = pt1Data.Value
    End If

    ' --- Copy PivotTable2 (all data) to 7.MONTH.AC ---
    pt2.PivotFields("Date").ClearAllFilters
    pt2.RefreshTable

    Set pt2Full = pt2.TableRange1
    rowCount2 = pt2Full.Rows.Count
    colCount2 = pt2Full.Columns.Count

    If LCase(Trim(pt2Full.Cells(rowCount2, 1).Text)) = "grand total" Then
        rowCount2 = rowCount2 - 1
    End If

    If rowCount2 > 1 Then
        Set pt2Data = pt2Full.Offset(1, 0).Resize(rowCount2 - 1, colCount2)

        pasteRow = 2
        For r = 1 To pt2Data.Rows.Count
            Do While wsMonth.Rows(pasteRow).Hidden
                pasteRow = pasteRow + 1
            Loop
            wsMonth.Range("A" & pasteRow & ":C" & pasteRow).Value = pt2Data.Rows(r).Value
            pasteRow = pasteRow + 1
        Next r
    End If

    ' --- Copy PivotTable1 (all data) to 7.MONTH.AC (Charge) ---
    Dim wsMonthCharge As Worksheet
    Set wsMonthCharge = ThisWorkbook.Sheets("7.MONTH.AC (Charge)")

    pt1.PivotFields("Date").ClearAllFilters
    pt1.RefreshTable

    Set pt1Full = pt1.TableRange1
    rowCount1 = pt1Full.Rows.Count
    colCount1 = pt1Full.Columns.Count

    isGrandTotalRow = False
    For Each cell In pt1Full.Rows(rowCount1).Cells
        If LCase(Trim(cell.Text)) = "grand total" Then
            isGrandTotalRow = True
            Exit For
        End If
    Next cell

    If isGrandTotalRow Then
        rowCount1 = rowCount1 - 1
    End If

    If rowCount1 > 1 Then
        Set pt1Data = pt1Full.Offset(1, 0).Resize(rowCount1 - 1, colCount1)

        pasteRow = 2
        For r = 1 To pt1Data.Rows.Count
            Do While wsMonthCharge.Rows(pasteRow).Hidden
                pasteRow = pasteRow + 1
            Loop
            wsMonthCharge.Range("A" & pasteRow & ":C" & pasteRow).Value = pt1Data.Rows(r).Value
            pasteRow = pasteRow + 1
        Next r
    End If

End Sub
