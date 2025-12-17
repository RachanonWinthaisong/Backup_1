Attribute VB_Name = "Module13"
Sub RunAllStepsInOrder()

    Call Filter_ColumnF_ErrorNA_ThenDeleteVisible_1
    Call Lookup_Team_Repo_2
    Call Filter_ColumnU_Q_3
    Call Copy_LookupRepo_4
    Call Xlook_Repo_5
    Call ClearFilter_6
    Call DeleteColumnU_7
    Call Lookup_6Repo_8
    Call Filter_ColumnU_9
    Call ChangeCol_P_10
    Call Filter_ColumnU_11
    Call ChangeCol_P_12
    Call ClearFilter_6
    Call DeleteColumnU_7
    

    MsgBox "เรียบร้อย!"
End Sub
