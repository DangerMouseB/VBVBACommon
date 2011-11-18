Attribute VB_Name = "mWinAPI_LoadDLLs"
Option Explicit

Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private myDLLPathsByName As New Dictionary
Private myDLLHandlesByName As New Dictionary

Function DLLHandle(name As String, Optional filename As String) As Long
    Dim hModule As Long
    If myDLLPathsByName.Exists(name) Then DLLHandle = myDLLHandlesByName(name): Exit Function
    hModule = GetModuleHandle(name)                          ' see if it's loaded
    If hModule = 0 And filename <> "" Then hModule = LoadLibrary(filename)    ' if not try to load it
    If hModule = 0 Then Exit Function
    myDLLPathsByName(name) = DLLFilename(hModule)
    myDLLHandlesByName(name) = hModule
    DLLHandle = hModule
End Function

Private Function DLLFilename(hModule As Long) As String
    Dim buffer As String * 1024, length As Long
    buffer = String(1024, Chr$(0))
    length = GetModuleFileName(hModule, buffer, Len(buffer))
    DLLFilename = Left$(buffer, length)
End Function

Function unloadDLL(name As String) As String
    Dim hModule As Long
    hModule = GetModuleHandle(name)
    If hModule <> 0 Then FreeLibrary hModule
    If GetModuleHandle(name) > 0 Then
        unloadDLL = name & " still loaded"
    Else
        unloadDLL = "..."
    End If
End Function

