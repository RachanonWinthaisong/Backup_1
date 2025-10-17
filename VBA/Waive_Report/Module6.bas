Attribute VB_Name = "Module6"
Sub RunAllStepsInOrder()

    Call UpdateAllPivotTablesIn2Pivot
    Call UpdatePivotTablesIn5acPivot
    Call UpdatePivotTable5In8p3k
    Call FilterPivotTable1ByLatestDate
    Call FilterColumnE_NotDashOrZero
    Call FillInColumnD_VisibleOnly_Corrected1
    Call FillInColumnD_VisibleOnly_Corrected2
    

    MsgBox "All Done!"
End Sub
