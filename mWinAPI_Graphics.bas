Attribute VB_Name = "mWinAPI_Graphics"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit
Option Private Module

' Graphics
Public Const WHITE_BRUSH As Integer = 0
Public Const BLACK_BRUSH As Integer = 4
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

' Graphics
Declare Function apiLoadIconA Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function apiLoadCursorA Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Declare Function apiGetStockObject Lib "gdi32" Alias "GetStockObject" (ByVal nIndex As Long) As Long
Declare Function apiLoadCursorFromFileA Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function apiExtractIconA Lib "shell32" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Long) As Long
Declare Function apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function apiOpenIcon Lib "user32" Alias "OpenIcon" (ByVal hwnd As Long) As Long
