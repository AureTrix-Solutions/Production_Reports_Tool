Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub RunSelectedBtn_Click()
DoCmd.Requery
On Error GoTo Err_Command1_Click
    Dim str As String
    Dim dbpath As String
    Dim fso As Object
    Dim finalPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    dbpath = CurrentProject.Path
    If Not (dbpath Like "*Production Reports*") Then
    finalPath = fso.BuildPath(dbpath, "Production Reports")
    Else: finalPath = dbpath
    End If

            DoCmd.SetWarnings False


            Call modNalcomis.VerifyNALCOMIS
        DoCmd.RunSQL "DELETE * FROM [Maf_Notes]"
        DoCmd.RunSQL "DELETE * FROM [combined]"
        DoCmd.RunSQL "DELETE * FROM [Requisitions]"
        DoCmd.RunSQL "DELETE * FROM [Inventory]"
        DoCmd.RunSQL "DELETE * FROM [Hof_Niin]"
        str = "4_5_6_7 historical.csv"
               DoCmd.TransferText acImportDelim, "4hist", "combined", finalPath & "\reports new adhocs\4_5_6_7 historical.csv", False
        str = "8_9 historical.csv"
               DoCmd.TransferText acImportDelim, "4hist", "combined", finalPath & "\reports new adhocs\8_9 historical.csv", False
        str = "reqs_hist.csv"
               DoCmd.TransferText acImportDelim, "reqs1", "requisitions", finalPath & "\reports new adhocs\reqs_hist.csv", False
        str = "active.csv"
               DoCmd.TransferText acImportDelim, "4hist", "combined", finalPath & "\reports new adhocs\active.csv", False
         str = "reqs1.csv"
               DoCmd.TransferText acImportDelim, "reqs1", "requisitions", finalPath & "\reports new adhocs\reqs1.csv", False
         str = "reqs2.csv"
               DoCmd.TransferText acImportDelim, "reqs1", "requisitions", finalPath & "\reports new adhocs\reqs2.csv", False
         str = "mafnotes.csv"
               DoCmd.TransferText acImportDelim, "mafnotes", "Maf_Notes", finalPath & "\reports new adhocs\mafnotes.csv", False
         str = "inventory.csv"
               DoCmd.TransferText acImportDelim, "inventory", "Inventory", finalPath & "\reports new adhocs\inventory.csv", False
        str = "hof_niin.csv"
               DoCmd.TransferText acImportDelim, "hof_niin", "Hof_niin", finalPath & "\reports new adhocs\hof_niin.csv", False

        DoCmd.OpenQuery ("Requisitions Query1")
        DoCmd.OpenQuery ("Maf_Notes Query1")
        DoCmd.OpenQuery ("combined query1")
        DoCmd.OpenQuery ("maf_notes query")
        DoCmd.OpenQuery ("requisitions query")
        DoCmd.OpenQuery ("maf_age_q")
        DoCmd.OpenQuery ("cost_key")
        DoCmd.OpenQuery ("p8_ddsn_list_clean_up")
        DoCmd.OpenQuery ("h60_ddsn_list_clean_up")
        DoCmd.OpenQuery ("growler_ddsn_list_clean_up")
        Call DeleteImportErrors
        Call DeleteAutoCorrect
        Call Form_Gen("500_600_Priority_List", "Priority_List_W/Mafs & Notes", "500_600_DIFM_trimmed_notes", "fgc", "fgc")
        Call Form_Gen("500_600_Priority_List", "Priority_List_W/Mafs & Parts", "500_600_DIFM_trimmed_parts", "fgc", "fgc")
        Call Form_Gen("500_600_Priority_List", "Priority_List_W/Mafs & Support", "500_600_DIFM_trimmed_support", "fgc", "fgc")
        Call Form_Gen("500_600_Pri_Report", "Priority_List_Report", "", "", "")
        Call UpdateFileModifiedDates
        Set fso = Nothing

                MsgBox "Adhocs Are Done Running And Tables Are Updated!"


    Exit Sub

Exit_Command1_Click:
    Exit Sub

Err_Command1_Click:
    MsgBox Err.Description
    Resume Exit_Command1_Click
End Sub

Private Sub select_all_Click()
CurrentDb.Execute ("update [configtbl_adhocengine] set runadhoc=true")
DoCmd.Requery
CurrentDb.Execute ("update [configtbl_adhocengine] set runadhoc=true")
DoCmd.Requery
End Sub

Private Sub UNSELECT_ALL_Click()
CurrentDb.Execute ("update [configtbl_adhocengine] set runadhoc=false")
DoCmd.Requery
CurrentDb.Execute ("update [configtbl_adhocengine] set runadhoc=false")
DoCmd.Requery

End Sub
