Attribute VB_Name = "mWinAPI_Windows"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit
Option Private Module

Public Const WM_QUIT = &H12
Public Const WM_USER = &H400

' see - http://msdn.microsoft.com/en-us/library/ms633574%28VS.85%29.aspx#class_elements
Public Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

' see - http://www.nirsoft.net/vb/wndclass.html
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As Long
End Type

' GetLastError
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Public Const LANG_NEUTRAL = &H0

' CreateWindowExA
Public Const HWND_MESSAGE As Long = -3
Public Const HWND_BROADCAST As Long = &HFFFF&

Public Const WS_VISIBLE As Long = &H10000000
Public Const WS_VSCROLL As Long = &H200000
Public Const WS_TABSTOP As Long = &H10000
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_BORDER As Long = &H800000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_CHILDWINDOW As Long = (WS_CHILD)
Public Const WS_CLIPCHILDREN As Long = &H2000000
Public Const WS_CLIPSIBLINGS As Long = &H4000000
Public Const WS_DISABLED As Long = &H8000000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_EX_ACCEPTFILES As Long = &H10&
Public Const WS_EX_DLGMODALFRAME As Long = &H1&
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Public Const WS_EX_TOPMOST As Long = &H8&
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const WS_GROUP As Long = &H20000
Public Const WS_HSCROLL As Long = &H100000
Public Const WS_ICONIC As Long = WS_MINIMIZE
Public Const WS_OVERLAPPED As Long = &H0&
Public Const WS_OVERLAPPEDWINDOW As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP As Long = &H80000000
Public Const WS_POPUPWINDOW As Long = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX As Long = WS_THICKFRAME
Public Const WS_TILED As Long = WS_OVERLAPPED
Public Const WS_TILEDWINDOW As Long = WS_OVERLAPPEDWINDOW

Public Const WM_GETTEXT = &HD

Public Const ICC_USEREX_CLASSES = &H200

Public Const CW_USEDEFAULT As Long = &H80000000

Public Const CS_HREDRAW As Long = &H2
Public Const CS_VREDRAW As Long = &H1
Public Const CS_OWNDC As Long = &H20
Public Const CS_GLOBALCLASS = &H4000

Public Const COLOR_APPWORKSPACE = 12

Public Const IDI_APPLICATION As Long = 32512&
Public Const IDC_ARROW As Long = 32512&

Public Const GWL_WNDPROC As Long = -4&

Public Const GW_Child As Long = 5
Public Const GW_HWNDFIRST As Long = 0
Public Const GW_HWNDLAST As Long = 1
Public Const GW_HWNDNEXT As Long = 2
Public Const GW_HWNDPREV As Long = 3
Public Const GW_OWNER As Long = 4


' Error handling
Declare Function apiGetLastError Lib "kernel32" Alias "GetLastError" () As Long
Declare Function apiFormatMessageA Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long

' Windows
Declare Function apiCreateWindowExA Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function apiDestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hwnd As Long) As Long
Declare Function apiFindWindowExA Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function apiFindWindowA Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function apiEnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lparam As Long) As Long
Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
Declare Function apiGetWindowTextA Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function apiIsWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long

' Messages
Declare Function apiRegisterWindowMessageA Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Declare Function apiDefWindowProcA Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Declare Function apiPostMessageA Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Declare Function apiCallWindowProcA Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Any, ByVal hwnd As Long, ByVal msg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Declare Function apiSendMessageA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Declare Function apiSetWindowLongA Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function apiGetWindowLongA Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' Classes
Declare Function apiRegisterClassA Lib "user32" Alias "RegisterClassA" (lpWndClass As WNDCLASS) As Long
Declare Function apiRegisterClassExA Lib "user32" Alias "RegisterClassExA" (lpwcx As WNDCLASSEX) As Long
Declare Function apiUnregisterClassA Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As Any, ByVal hInstance As Long) As Long
Declare Function apiGetClassNameA Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' Atoms
Declare Function apiGlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Long
Declare Function apiGlobalGetAtomNameA Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Long, ByVal lpString As String, ByVal nSize As Long) As Long
Declare Function apiGlobalDeleteAtom Lib "kernel32" Alias "GlobalDeleteAtom" (ByVal nAtom As Long) As Long

Function lastDLLErrorDescription(lastDLLError As Long) As String
    Dim buffer As String, length As Long
    buffer = String$(256, 0)
    length = apiFormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lastDLLError, 0&, buffer, Len(buffer), ByVal 0)
    If length > 0 Then lastDLLErrorDescription = Left$(buffer, length - 2)
End Function

