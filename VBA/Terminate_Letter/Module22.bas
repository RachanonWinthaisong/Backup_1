Attribute VB_Name = "Module22"
Sub RunAllStepsInOrder_2()

    Call Lookup_NT_14
    Call Filter_ColumnU_9
    Call DeleteVisibleColumnsV_W_15
    Call InsertToday_ColV_16
    Call ClearFilter_6
    Call DeleteColumnU_7
    Call Lookup_Lastone_17
    Call Filter_ColumnE_18
    Call DeleteRow_19
    Call ClearFilterLast_20

    MsgBox "เรียบร้อย!"
End Sub
