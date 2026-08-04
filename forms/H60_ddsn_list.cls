Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
DoCmd.SetWarnings False
DoCmd.OpenQuery "h60_ddsn_list_clean_up"
DoCmd.Requery
DoCmd.SetWarnings True
End Sub
