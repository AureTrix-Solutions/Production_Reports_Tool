Attribute VB_Name = "FirstRun"
Option Compare Database
Option Explicit

Public Sub SetupProjectFoldersAndFiles_FromTable()

    Const MAIN_FOLDER_NAME As String = "Production Reports"
    Const SCRIPTS_SUBFOLDER As String = "reports new adhocs"

    Dim fso As Object
    Dim db As DAO.Database
    Dim rsAh As DAO.Recordset
    Dim rsCsv As DAO.Recordset

    Dim basePath As String
    Dim mainFolderPath As String
    Dim scriptsFolderPath As String

    Dim setupWasRequired As Boolean ' <-- The new flag variable
    setupWasRequired = False        ' <-- Initialize to False

    On Error GoTo Err_Handler

    ' 1. INITIALIZE OBJECTS
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb

    ' 2. ENSURE FOLDER STRUCTURE EXISTS
    basePath = CurrentProject.Path
    mainFolderPath = fso.BuildPath(basePath, MAIN_FOLDER_NAME)
    If Not fso.FolderExists(mainFolderPath) Then
        fso.CreateFolder mainFolderPath
        setupWasRequired = True ' <-- Set flag
    End If

    scriptsFolderPath = fso.BuildPath(mainFolderPath, SCRIPTS_SUBFOLDER)
    If Not fso.FolderExists(scriptsFolderPath) Then
        fso.CreateFolder scriptsFolderPath
        Debug.Print " -> Created subfolder: " & scriptsFolderPath
        setupWasRequired = True ' <-- Set flag
    Else
        Debug.Print "Scripts folder already exists: " & scriptsFolderPath
    End If

    ' 3. WRITE .AH FILES FROM TABLE
    Debug.Print "Processing .ah files..."
    Set rsAh = db.OpenRecordset("SELECT FileName, FileContent FROM tblAhTemplates", dbOpenSnapshot)
    If rsAh.RecordCount > 0 Then
        rsAh.MoveFirst
        Do While Not rsAh.EOF
            ' Check the return value of the function
            If WriteTextFile(fso, scriptsFolderPath, rsAh!fileName, rsAh!FileContent) Then
                setupWasRequired = True ' <-- Set flag if function returns True
            End If
            rsAh.MoveNext
        Loop
    Else
        Debug.Print "No .ah templates found in 'tblAhTemplates'. Skipping."
    End If

    ' 4. CREATE EMPTY CSV FILES FROM TABLE
    Debug.Print "Processing empty CSV files..."
    Set rsCsv = db.OpenRecordset("SELECT FileName FROM tblCsvFileNames", dbOpenSnapshot)
    If rsCsv.RecordCount > 0 Then
        rsCsv.MoveFirst
        Do While Not rsCsv.EOF
            ' Check the return value of the function
            If CreateEmptyFile(fso, scriptsFolderPath, rsCsv!fileName) Then
                setupWasRequired = True ' <-- Set flag if function returns True
            End If
            rsCsv.MoveNext
        Loop
    Else
        Debug.Print "No CSV filenames found in 'tblCsvFileNames'. Skipping."
    End If

    Debug.Print "Setup complete!"

    ' --- FINAL CHECK ---
    ' Only show the MsgBox if the flag was tripped
If setupWasRequired Then
    MsgBox "Project setup is complete." & vbCrLf & vbCrLf & _
           "Recommend running all AdHocs for first use", vbInformation, "Success"
End If



Exit_Sub:
    If Not rsAh Is Nothing Then rsAh.Close
    If Not rsCsv Is Nothing Then rsCsv.Close
    Set rsAh = Nothing
    Set rsCsv = Nothing
    Set db = Nothing
    Set fso = Nothing
    Exit Sub

Err_Handler:
    MsgBox "An error occurred:" & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume Exit_Sub

End Sub

' --- CONVERTED TO A FUNCTION ---
Private Function WriteTextFile(fso As Object, ByVal folderPath As String, ByVal fileName As String, ByVal content As String) As Boolean

    Dim fullPath As String
    fullPath = fso.BuildPath(folderPath, fileName)

    If Not fso.FileExists(fullPath) Then
        Dim file As Object
        Set file = fso.CreateTextFile(fullPath, True)
        file.Write content
        file.Close
        Debug.Print "   -> Wrote file: " & fullPath
        WriteTextFile = True ' <-- Return True because we took action
    Else
        Debug.Print "   -> Skipped (already exists): " & fullPath
        WriteTextFile = False ' <-- Return False because we did nothing
    End If

End Function

' --- CONVERTED TO A FUNCTION ---
Private Function CreateEmptyFile(fso As Object, ByVal folderPath As String, ByVal fileName As String) As Boolean

    Dim fullPath As String
    fullPath = fso.BuildPath(folderPath, fileName)

    If Not fso.FileExists(fullPath) Then
        fso.CreateTextFile(fullPath, True).Close
        Debug.Print "   -> Created empty file: " & fullPath
        CreateEmptyFile = True ' <-- Return True because we took action
    Else
        Debug.Print "   -> Skipped (already exists): " & fullPath
        CreateEmptyFile = False ' <-- Return False because we did nothing
    End If

End Function
