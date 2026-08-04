Attribute VB_Name = "FixCapsNumScroll"
Option Compare Database
Option Explicit
Private Const VK_CAPITAL = &H14
Private Const VK_NUMLOCK = &H90
Private Const VK_SCROLL = &H91
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare PtrSafe Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function GetKeyboardState Lib "USER32" (pbKeyState As Byte) As Long
Private Declare PtrSafe Function SetKeyboardState Lib "USER32" (lppbKeyState As Byte) As Long
Public Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal nVirtKey As Long) As Integer
Public OnLoadNumlockOnOff, OnLoadCapslockOnOff, OnLoadScrolllockOnOff As Boolean
Public Function GetCapslock() As Boolean
    GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
    OnLoadCapslockOnOff = GetCapslock
End Function
Public Function GetNumlock() As Boolean
    GetNumlock = CBool(GetKeyState(vbKeyNumlock) And 1)
    OnLoadNumlockOnOff = GetNumlock
End Function
Public Function GetScrollLock() As Boolean
    GetScrollLock = CBool(GetKeyState(&H91) And 1)
    OnLoadScrolllockOnOff = GetScrollLock
End Function
Public Sub ToggleNumLock(TurnOn As Boolean)
    Dim bytKeys(255) As Byte
    Dim bNumLockOn As Boolean
    GetKeyboardState bytKeys(0)
    bNumLockOn = bytKeys(VK_NUMLOCK)
    If OnLoadNumlockOnOff = True Then
        If bNumLockOn <> TurnOn Then
            keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
            keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub
Public Sub ToggleCapsLock(TurnOn As Boolean)
    Dim bytKeys(255) As Byte
    Dim bCapsLockOn As Boolean
    GetKeyboardState bytKeys(0)
    bCapsLockOn = bytKeys(VK_CAPITAL)
    If OnLoadCapslockOnOff = True Then
        If bCapsLockOn <> TurnOn Then
            keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
            keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub
Public Sub ToggleScrollLock(TurnOn As Boolean)
    Dim bytKeys(255) As Byte
    Dim bScrollLockOn As Boolean
    GetKeyboardState bytKeys(0)
    bScrollLockOn = bytKeys(VK_SCROLL)
    If OnLoadScrolllockOnOff = True Then
        If bScrollLockOn <> TurnOn Then
            keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
            keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub
'Above procedures are used to get the key state for CAPS/NUM/SCROLL and toggle them back to that state if turned off by SendKeys
