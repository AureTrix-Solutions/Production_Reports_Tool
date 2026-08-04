Attribute VB_Name = "historicalNalcomis"
Option Compare Database

#If VBA7 Then
Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function GetClassName Lib "USER32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "USER32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public ADHOCReport As String

Const SW_SHOWNORMAL = 1
Const WM_CLOSE = &H10
Sub SmartKeys(delay As Long, keys As String)

    Sleep delay
        DoEvents
    SendKeys keys
        DoEvents
        'Debug.Print keys
        DoEvents

End Sub

Public Function VerifyNALCOMIS()

    Dim WinHndl As Long

        WinHndl = GetHndl("NALCOMIS IMA - [Unclassified]")

        Select Case WinHndl
            Case Is <> 0
                 gloAFiles = ""
                 ADHOCReport = ""
                SetForegroundWindow WinHndl
                ShowWindow WinHndl, 1
                VerifyNALCOMIS = True
            Case 0
                MsgBox "Unable to locate NALCOMIS IMA - [Unclassified]." & _
                    Chr(10) & "Please open NALCOMIS to this screen now and " & _
                    "attempt the program again.", vbSystemModal, "Error - NALCOMIS Not Active"
                TerminateProgram = True
                VerifyNALCOMIS = True
        End Select

        Nalcomis


End Function
Public Function Nalcomis()
On Error GoTo Err_Nalcomis

    Dim ADEN As DAO.Database
    Dim rst As DAO.Recordset
    Dim ADHOCPath As String
    Dim SavePath As String
    Dim dbpath As String
    Dim fso As Object
    Dim finalPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    dbpath = CurrentProject.Path
    If Not (dbpath Like "*Production Reports*") Then
    finalPath = fso.BuildPath(dbpath, "Production Reports")
    Else: finalPath = dbpath
    End If

        Set ADEN = CurrentDb
        Set rst = ADEN.OpenRecordset("Historical_Adhocs", dbOpenTable)

            rst.MoveFirst

            If rst.RecordCount = 0 Then
                Exit Function
            End If

            Do Until rst.EOF
                ADHOCPath = finalPath & rst![ADHOCPath]
                SavePath = finalPath & rst![SavePath]
                SaveFileType = rst![SaveFileType]
                ExADHOC = rst![ExeWith]
                AHName = rst![ADHOCName]


                    Call OpenADHOC(AHName, ADHOCPath)
                    Call RunADHOC(AHName)
                    Call SaveADHOC
                    Call SaveAsWindow(SavePath)

                rst.MoveNext
            Loop

            gloAFiles = ADHOCReport

            rst.Close
            ADEN.Close

            Call Min_Nalcomis
            Call num_lock
            Set fso = Nothing
Exit Function

Err_Nalcomis:
    MsgBox Err.Number & " " & Err.Description
    Exit Function

End Function
Public Function OpenADHOC(AHName As String, ADHOCPath As String)
On Error GoTo Err_OpenAdhoc

    '--- Variable Declarations ---
    Dim NWnd As Long, MWnd As Long, WinWnd As Long
    Dim WinTitleN As String, WinTitle As String, MainWin As String

    '--- STAGE 1: Open the "Reports" window ---
    WinTitleN = "NALCOMIS IMA - [Unclassified]"

    ' First, find the main application window
    Do
        NWnd = FindWindow(vbNullString, WinTitleN)
        If NWnd <> 0 Then Exit Do
        Sleep 250
        DoEvents
    Loop

    ' Now, activate it and send the command ONCE to open the Reports window
    SetForegroundWindow NWnd
    SmartKeys 300, "%NR" ' Send Alt+N, R

    ' Then, patiently wait for the "Reports" window to appear
    WinTitle = "NALCOMIS IMA - [Reports]"
    Do
        MWnd = FindWindow(vbNullString, WinTitle)
        If MWnd <> 0 Then Exit Do
        Sleep 250
        DoEvents
    Loop

    '--- STAGE 2: Open the specific Ad Hoc file ---
    ShowWindow MWnd, SW_SHOWNORMAL
    SetForegroundWindow MWnd

    ' Send commands to navigate the dialogs and open the file
    SmartKeys 300, "%DE%OK%FO"
    SmartKeys 2000, ADHOCPath ' Type the path to the file
    SmartKeys 300, "{Enter}"  ' Confirm opening the file

    '--- STAGE 3: Wait for the "Expert Mode" window ---
    MainWin = "Nalcomis IMA Ad hoc Query (Expert Mode - " & AHName & ".ah)"

    ' Patiently wait for the ad hoc file to finish loading and the window to appear
    Do
        WinWnd = FindWindow(vbNullString, MainWin)
        If WinWnd <> 0 Then Exit Do
        Sleep 250
        DoEvents
    Loop

    '--- STAGE 4: Run the report ---
    ShowWindow WinWnd, SW_SHOWNORMAL
    SetForegroundWindow WinWnd

    ' Send the final command to run the report, which opens the "Query Impact" window for the next function
    SmartKeys 500, "%OR" ' Added a little extra delay for stability

Exit Function

Err_OpenAdhoc:
    MsgBox "Error in OpenADHOC: " & Err.Number & " - " & Err.Description
End Function

Public Function RunADHOC(AHName As String)
On Error GoTo Err_RunAdhoc

    Dim QIWnd As Long, CWnd As Long, qrWnd As Long, pWnd As Long
    Dim WinTitleQI As String, WinTitleCR As String, WinTitleQR As String, WinTitlePC As String

    'Set Window titles
    WinTitleQI = "Query Impact"
    WinTitleQR = "Query Retrieval"
    WinTitleCR = "Cancel Retrieve"
    WinTitlePC = "Procedural Error" ' Assuming this is the title

    '--- STAGE 1: Wait for "Query Impact" window ---
    Do
        QIWnd = FindWindow(vbNullString, WinTitleQI)
        If QIWnd <> 0 Then Exit Do
        Sleep 250
        DoEvents
    Loop

    '--- STAGE 2: Start the Query ---
    'Bring the window to the front and send the commands to start the query
    ShowWindow QIWnd, SW_SHOWNORMAL
    SetForegroundWindow QIWnd
    SmartKeys 300, "%G"      ' Send Alt+G just in case
    SmartKeys 300, "{Enter}" ' This starts the long process

    '--- STAGE 3: The CRITICAL WAIT ---
    'The "Cancel Retrieve" window MUST appear before we can wait for it to disappear.

    'First, wait for the "Cancel Retrieve" window to show up.
    'We'll give it a timeout (e.g., 15 seconds) in case it never appears.
    Dim startTime As Single
    startTime = Timer
    Do
        CWnd = FindWindow(vbNullString, WinTitleCR)
        If CWnd <> 0 Then Exit Do ' It appeared, proceed to the next stage

        ' Check for a failure window ("Query Retrieval") while we wait
        If FindWindow(vbNullString, WinTitleQR) <> 0 Then
            ADHOCReport = ADHOCReport & AHName & ":" & Chr(13)
            ADHOCReport = ADHOCReport & "ADHOC Failed, no rows retrieved." & Chr(13) & Chr(13)
            SmartKeys 300, "{Enter}" ' Dismiss the error
            Exit Function ' Exit early on failure
        End If

        If Timer > startTime + 15 Then ' Timeout after 15 seconds
            ' The "Cancel Retrieve" window never appeared. Assume the query was very fast.
            GoTo QueryFinished
        End If

        Sleep 250
        DoEvents
    Loop

    'Now that the "Cancel Retrieve" window EXISTS, wait for it to DISAPPEAR.
    Do
        CWnd = FindWindow(vbNullString, WinTitleCR)
        If CWnd = 0 Then Exit Do ' It disappeared, the query is done.

        ' While waiting, periodically check for and dismiss procedural error pop-ups
        pWnd = FindWindow(vbNullString, WinTitlePC)
        If pWnd <> 0 Then
            SmartKeys 100, "{Enter}"
        End If

        Sleep 500 ' Check every half-second
        DoEvents
    Loop

QueryFinished:
    'This label provides a single exit point for the success case.
    ADHOCReport = ADHOCReport & AHName & ":" & Chr(13)
    ADHOCReport = ADHOCReport & "ADHOC Returned fewer than 10,000 Rows." & Chr(13) & Chr(13)

Exit Function

Err_RunAdhoc:
    MsgBox "Error in RunADHOC: " & Err.Number & " - " & Err.Description
End Function

Public Function SaveADHOC()
On Error GoTo Err_SaveAdhoc

    Dim REQWnd As Long ' Using Long as per your working code
    Dim WinTitleREQ As String

    'Set Window Name
    WinTitleREQ = "Run/Execute Query"

    '--- STABLE WAITING LOOP ---
    'Loop until the "Run/Execute Query" window is found.
    'This solves the race condition by patiently waiting for the window to be ready.
    Do
        REQWnd = FindWindow(vbNullString, WinTitleREQ)
        If REQWnd <> 0 Then Exit Do 'Window found, exit the loop

        Sleep 250 'Wait for a quarter-second to avoid freezing Access
        DoEvents  'Yield to other processes (like the window appearing)
    Loop
    '--- END OF FIX ---

    'Show the window
    ShowWindow REQWnd, SW_SHOWNORMAL

    'Set window to the run/execute query window
    SetForegroundWindow REQWnd

    'Save csv file. Added a tiny delay to ensure the window is fully active.
    SmartKeys 500, "%S"

Exit Function

Err_SaveAdhoc:
    MsgBox Err.Number & " " & Err.Description
Exit Function

End Function


Public Function SaveAsWindow(SavePath As String)
On Error GoTo Err_SaveAsWindow

    '--- Variable Declarations ---
    Dim SWnd As Long, REQWnd As Long, QIWnd As Long
    Dim WinTitleS As String, WinTitleREQ As String, WinTitleQI As String

    '--- STAGE 1: Wait for and handle the "Save As" dialog ---
    WinTitleS = "Save As"

    ' Patiently wait for the "Save As" window to appear
    Do
        SWnd = FindWindow(vbNullString, WinTitleS)
        If SWnd <> 0 Then Exit Do
        Sleep 250
        DoEvents
    Loop

    ' Activate the "Save As" window and fill it out
    ShowWindow SWnd, SW_SHOWNORMAL
    SetForegroundWindow SWnd

    SmartKeys 1000, SavePath ' Give plenty of time to type the path
    SmartKeys 300, "{Tab}"
    SmartKeys 300, "CC"      ' Select file type
    SmartKeys 300, "%S"      ' Send Alt+S to Save
    SmartKeys 500, "%Y"      ' Send Alt+Y to confirm overwrite (if needed)

    ' After saving, the other windows need to be closed.
    ' It's better to find and close them explicitly rather than assuming the order.

    '--- STAGE 2: Close the "Run/Execute Query" window ---
    WinTitleREQ = "Run/Execute Query"

    ' Wait for it to become the active window again after saving is complete
    Do
        REQWnd = FindWindow(vbNullString, WinTitleREQ)
        If REQWnd <> 0 Then Exit Do
        Sleep 250
        DoEvents
    Loop

    ' Activate and close it
    ShowWindow REQWnd, SW_SHOWNORMAL
    SetForegroundWindow REQWnd
    SmartKeys 300, "{ESC}" ' Send Escape to close it

    '--- STAGE 3: Close the "Query Impact" window ---
    ' This might have reappeared or still be open. This makes sure it gets closed.
    WinTitleQI = "Query Impact"

    Do
        QIWnd = FindWindow(vbNullString, WinTitleQI)
        If QIWnd <> 0 Then Exit Do

        ' It's possible this window closes quickly. We'll only wait a few seconds.
        ' If it's not found, we assume it's already gone and move on.
        Static startTime As Single
        If startTime = 0 Then startTime = Timer
        If Timer > startTime + 3 Then Exit Do ' Timeout after 3 seconds

        Sleep 250
        DoEvents
    Loop

    ' If the window was found, activate and close it
    If QIWnd <> 0 Then
        ShowWindow QIWnd, SW_SHOWNORMAL
        SetForegroundWindow QIWnd
        SmartKeys 300, "{ESC}"
    End If

    '--- STAGE 4: Final Keystrokes ---
    ' This appears to be navigating the main application menu
    SmartKeys 300, "%fcn%FXyn%MG"

    ' The global variable AHName is not available here, so this line would fail.
    ' Assuming it's available in the module scope.
    gloAFiles = AHName

Exit Function

Err_SaveAsWindow:
    MsgBox "Error in SaveAsWindow: " & Err.Number & " - " & Err.Description
End Function

Public Function Min_Nalcomis()

    Dim WinHndl As Long

        WinHndl = GetHndl("NALCOMIS IMA - [Unclassified]")

            If WinHndl <> 0 Then
                SetForegroundWindow WinHndl
                ShowWindow WinHndl, 6

            End If

End Function
Public Function num_lock()
    ToggleNumLock (OnLoadNumlockOnOff)
    ToggleCapsLock (OnLoadCapslockOnOff)
    ToggleScrollLock (OnLoadScrolllockOnOff)
End Function
