Attribute VB_Name = "API"
'*****************************************************
'*****************************************************
'**     This program was designed as part of a      **
'**       personal project to improve AIMD PC       **
'**      production research.                       **
'**                                                 **
'**    - Written and developed by Rick Correia -    **
'*****************************************************
'*****************************************************





'This module controls Nalcomis.
'It relies on the use of sendmessage
'and postmessage Windows API commands
'among others. These are vastly superior
'and more reliable than sendkeys

Option Explicit

'******** WINDOWS API *************
#If VBA7 Then
Public Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare PtrSafe Function SendMessageSTRING Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare PtrSafe Function GetMenuItemID Lib "USER32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare PtrSafe Function GetMenu Lib "USER32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function GetSubMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare PtrSafe Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare PtrSafe Function GetDlgCtrlID Lib "USER32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Sub SetWindowPos Lib "USER32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#Else
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageSTRING Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function GetDlgCtrlID Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#End If
Const WM_COMMAND As Integer = &H111
Const WM_SETTEXT As Long = &HC
Const BM_CLICK = &HF5
Const WM_SYSCOMMAND = &H112
Const SC_CLOSE = &HF060
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2
Const Nalc_prog As String = "F:\Program Files (x86)\ntcss\oimacl\bin\nalcomis.exe"


    Dim hwnd As Long
    Dim hMainMenu As Long
    Dim hMenu As Long
    Dim itemID As Long
    Dim dtStartTime As Date
    Dim Dlg_ChildWIN  As Long
