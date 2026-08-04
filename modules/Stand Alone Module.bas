Attribute VB_Name = "Stand Alone Module"
Option Compare Database
'ADHOC Engine Stand Alone Module was created by Travis Lee Toulson
'Created July 20, 2006
'ADEN v3.1 updated as of August 1, 2006

Public Function CloseForm(Name)
    DoCmd.Close acForm, Name
End Function

Public Function OpenForm(Name)
    DoCmd.OpenForm Name
End Function
Public Function DeleteImportErrors()
    Dim n As Integer
    Dim db As DAO.Database

    Set db = CurrentDb
    For n = db.TableDefs.Count - 1 To 0 Step -1
    ' loop through all tables
        If InStr(1, db.TableDefs(n).Name, "ImportError") > 0 Then
        ' if table is import errors table
            DoCmd.DeleteObject acTable, db.TableDefs(n).Name
            ' delete table
        End If
    Next n
End Function

Public Function DeleteAutoCorrect()
    Dim n As Integer
    Dim db As DAO.Database

    Set db = CurrentDb
    For n = db.TableDefs.Count - 1 To 0 Step -1
    ' loop through all tables
        If InStr(1, db.TableDefs(n).Name, "AutoCorrect") > 0 Then
        ' if table is import errors table
            DoCmd.DeleteObject acTable, db.TableDefs(n).Name
            ' delete table
        End If
    Next n
End Function
