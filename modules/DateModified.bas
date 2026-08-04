Attribute VB_Name = "DateModified"
Option Compare Database
Option Explicit

' Purpose: Checks the 'Date Last Modified' for a list of CSV files and records it in a table.

Public Sub UpdateFileModifiedDates()

    ' --- CONFIGURATION ---
    Const MAIN_FOLDER_NAME As String = "Production Reports"
    Const SCRIPTS_SUBFOLDER As String = "reports new adhocs"
    ' --- END CONFIGURATION ---

    Dim fso As Object
    Dim db As DAO.Database
    Dim rsSourceFiles As DAO.Recordset ' Reads from tblCsvFileNames

    Dim folderPath As String
    Dim fullFilePath As String
    Dim fileObj As Object
    Dim lastModDate As Date
    Dim dbpath As String

    On Error GoTo Err_Handler

    ' 1. INITIALIZE OBJECTS AND PATHS
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set db = CurrentDb

    ' Build the path to the folder containing the CSV files
    dbpath = CurrentProject.Path
    If Not (dbpath Like "*Production Reports*") Then
    folderPath = fso.BuildPath(fso.BuildPath(CurrentProject.Path, MAIN_FOLDER_NAME), SCRIPTS_SUBFOLDER)
    Else: folderPath = fso.BuildPath(dbpath, SCRIPTS_SUBFOLDER)
    End If

    ' Check if the target folder actually exists before proceeding
    If Not fso.FolderExists(folderPath) Then
        MsgBox "The folder was not found:" & vbCrLf & folderPath, vbCritical, "Folder Not Found"
        GoTo Exit_Sub
    End If

    ' 2. CLEAR THE DESTINATION TABLE
    ' This ensures we always have the freshest data and no old entries.
    Debug.Print "Clearing old data from [tblFileLastModified]..."
    db.Execute "DELETE * FROM tblFileLastModified", dbFailOnError

    ' 3. LOOP THROUGH THE SOURCE FILES AND GET TIMESTAMPS
    Debug.Print "Reading source file list from [tblCsvFileNames]..."
    Set rsSourceFiles = db.OpenRecordset("SELECT FileName FROM tblCsvFileNames", dbOpenSnapshot)

    If rsSourceFiles.RecordCount = 0 Then
        MsgBox "No CSV filenames were found in 'tblCsvFileNames'.", vbExclamation
        GoTo Exit_Sub
    End If

    rsSourceFiles.MoveFirst
    Do While Not rsSourceFiles.EOF
        fullFilePath = fso.BuildPath(folderPath, rsSourceFiles!fileName)

        ' Check if the specific CSV file exists
        If fso.FileExists(fullFilePath) Then
            ' Get the file object and its last modified date
            Set fileObj = fso.GetFile(fullFilePath)
            lastModDate = fileObj.DateLastModified

            ' Write the information to our destination table using a SQL INSERT statement
            db.Execute "INSERT INTO tblFileLastModified ( FileName, LastModified ) " & _
                       "VALUES ('" & Replace(rsSourceFiles!fileName, "'", "''") & "', #" & Format(lastModDate, "yyyy-mm-dd hh:nn:ss") & "#)", dbFailOnError

            Debug.Print " -> Logged: " & rsSourceFiles!fileName & " (" & lastModDate & ")"
        Else
            ' Log a warning if a file listed in the table is missing from the folder
            Debug.Print " -> WARNING: File not found in folder: " & rsSourceFiles!fileName
        End If

        rsSourceFiles.MoveNext
    Loop

Exit_Sub:
    ' Clean up all objects
    If Not rsSourceFiles Is Nothing Then rsSourceFiles.Close
    Set rsSourceFiles = Nothing
    Set db = Nothing
    Set fso = Nothing
    Exit Sub

Err_Handler:
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume Exit_Sub

End Sub
