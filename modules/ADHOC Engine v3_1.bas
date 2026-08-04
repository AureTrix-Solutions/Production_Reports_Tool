Attribute VB_Name = "ADHOC Engine v3_1"
Option Compare Database
'Created by Travis Lee Toulson June 2006
'ADHOC Engine is designed to be used in conjunction with NALCOMIS ADHOC,
    'NALCOMIS SQL, RSupply ADHOC or RSupply SQL
    'to automatically execute and save ADHOC and SQL queries to specified locations
'ADEN v3.1 updated as of August 1, 2006
Public Declare PtrSafe Function FindWindowA Lib "user32.dll" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Declare PtrSafe Function SetForegroundWindow Lib "user32.dll" _
    (ByVal hwnd As Long) As Long

Declare PtrSafe Function ShowWindow Lib "user32.dll" _
    (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As Long

Public Declare PtrSafe Function GetWindowTextA Lib "user32.dll" _
    (ByVal hwnd As Long, ByVal lpString As _
    String, ByVal cch As Long) As Long

Public Declare PtrSafe Function GetWindowTextLengthA Lib "user32.dll" _
    (ByVal hwnd As Long) As Long

Public Declare PtrSafe Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer

Public Const SW_MINIMIZE = 6, SW_NORMAL = 1, SW_MAXIMIZE = 3, SW_RESTORE = 9

Private Declare PtrSafe Function GetWindowPlacement Lib "USER32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function SetWindowPlacement Lib "USER32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long


Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type


Public TerminateProgram As Boolean
Public ADHOCPath As String
Public SavePath As String
Public SaveFileType As String
Public ExADHOC As String
Public AHName As String
Public WinName As String
Public ADHOCReport As String
Public gloName As String


Public Function GetHndl(WindowName As String) As Long
    'Obtain the Handle of a window with the window title text specified by WindowName
    GetHndl = FindWindowA(vbNullString, WindowName)
End Function
