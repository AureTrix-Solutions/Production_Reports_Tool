Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Load()
DoCmd.NavigateTo "production report tool"
DoCmd.Minimize
If CommandBars("ribbon").Height > 100 Then
    CommandBars.ExecuteMso "MinimizeRibbon"
End If
Call SetupProjectFoldersAndFiles_FromTable
Call UpdateFileModifiedDates
End Sub

Private Sub NavigationButton120_Click()
DoCmd.OpenForm "BfferRedFrm"
End Sub

Private Sub NavigationButton122_Click()
DoCmd.OpenForm "bfferyellowfrm"
End Sub

Private Sub NavigationButton137_Click()
DoCmd.OpenForm "sepmsfrm"
End Sub

Private Sub NavigationButton193_Click()
DoCmd.OpenForm "mocfrm"
End Sub

Private Sub NavigationButton370_Click()
    Dim FilePath As String
    Dim queryName As String
    Dim xlApp As Object
    Dim xlWB As Object
    Dim dbpath As String
    Dim fso As Object
    Dim finalPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    dbpath = CurrentProject.Path
    If Not (dbpath Like "*Production Reports*") Then
    finalPath = fso.BuildPath(dbpath, "Production Reports")
    Else: finalPath = dbpath
    End If

    queryName = "AK0_AK7_BSR"
    FilePath = finalPath & "\reports new adhocs\AK0_AK7_BSR.xlsx"

    If Dir(FilePath) <> "" Then
        Kill FilePath
    End If

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, queryName, FilePath, True

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Open(FilePath)

    Set xlWB = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
End Sub

Private Sub NavigationButton22_Click()
DoCmd.OpenForm "exrep_frm", acNormal
End Sub

Private Sub NavigationButton418_Click()
    Dim FilePath As String
    Dim queryName As String
    Dim xlApp As Object
    Dim xlWB As Object
    Dim strPath As String
    Dim dbpath As String
    Dim fso As Object
    Dim finalPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    dbpath = CurrentProject.Path
    If Not (dbpath Like "*Production Reports*") Then
    finalPath = fso.BuildPath(dbpath, "Production Reports")
    Else: finalPath = dbpath
    End If

    strPath = finalPath & "\Cost_savings Charts.xlsx"
    queryName = "cost_savings"
    FilePath = finalPath & "\reports new adhocs\cost_savings.xlsx"

    If Dir(FilePath) <> "" Then
        Kill FilePath
    End If

    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, queryName, FilePath, True

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Open(strPath)

    Set xlWB = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
End Sub

Private Sub NavigationButton672_Click()
DoCmd.OpenReport "priority_report", acViewPreview

End Sub

Private Sub NavigationButton696_Click()
DoCmd.OpenReport "500_600_IW_Report", acViewPreview
End Sub

Private Sub NavigationButton925_Click()
DoCmd.OpenReport "doghouse", acViewPreview

End Sub
