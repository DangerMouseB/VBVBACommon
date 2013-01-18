Attribute VB_Name = "mWinAPI_Graphics"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011 David Briant - see https://github.com/DangerMouseB
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'************************************************************************************************************************************************
 
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
