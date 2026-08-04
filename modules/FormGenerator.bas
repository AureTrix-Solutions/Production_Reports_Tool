Attribute VB_Name = "FormGenerator"
Public Sub Form_Gen(queryName As String, formName As String, subformSourceName As String, linkMasterField As String, linkChildField As String)
    Dim frm As Form
    Dim oldName As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim fld As DAO.Field
    Dim ctl As Control
    Dim subCtl As Control

    On Error Resume Next
    DoCmd.DeleteObject acForm, formName

    Set db = CurrentDb
    Set qdf = db.QueryDefs(queryName)

    Set frm = CreateForm()
    frm.RecordSource = queryName

    For Each fld In qdf.Fields
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail)
        ctl.ControlSource = fld.Name
        ctl.Name = fld.Name
    Next fld

    Set subCtl = CreateControl(frm.Name, acSubform, acDetail, "", "", 100, 1000)

    With subCtl
        .SourceObject = subformSourceName
        .LinkMasterFields = linkMasterField
        .LinkChildFields = linkChildField
        .Name = "MySubform"
    End With

    frm.DefaultView = 2
    frm.Caption = formName
    oldName = frm.Name

    DoCmd.Save acForm, oldName
    DoCmd.Close acForm, oldName, acSaveYes
    DoCmd.Rename formName, acForm, oldName
    DoCmd.OpenForm formName, acFormDS, , , , acHidden

    Set frm = Forms(formName)

    For Each ctl In frm.Controls
        On Error Resume Next
        ctl.ColumnWidth = -2
    Next ctl

    DoCmd.Close acForm, formName, acSaveYes




End Sub
