Attribute VB_Name = "mWinAPI_Common"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit
Option Private Module

Public Type VOID
    VOID As Long
End Type

Public Type HANDLE
    HANDLE As Long
End Type

Public Type HANDLE_OR_ZERO
    HANDLE_OR_ZERO As Long
End Type

Public Type DWORD
    DWORD As Long
End Type

Public Type BOOL
    BOOL As Long
End Type

Public Type ULARGE_INTEGER
    lowWord As Long
    highWord As Long
End Type

Public Type LARGE_INTEGER
    lowWord As Long
    highWord As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength  As Long
    lpSecurityDescriptor As Long
    bInheritHandle  As Boolean
End Type

Public Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public Const CURRENCY_TO_LARGE_INTEGER As Currency = 1 / 10000
Public Const LARGE_INTEGER_TO_CURRENCY As Currency = 10000

Public Const ERROR_SUCCESS As Long = 0&
Public Const ERROR_FILE_NOT_FOUND As Long = 2&
Public Const ERROR_MORE_DATA As Long = 234
Public Const ERROR_NO_MORE_ITEMS As Long = 259&

Public Const INVALID_HANDLE_VALUE = &HFFFFFFFF  '-1
Public Const API_NULL As Long = &H0

Public Const BOOL_FALSE As Long = &H0
Public Const BOOL_TRUE As Long = &H1

Public Const DLL_PROCESS_ATTACH As Long = 1
Public Const DLL_PROCESS_DETACH As Long = 0
Public Const DLL_THREAD_ATTACH As Long = 2
Public Const DLL_THREAD_DETACH As Long = 3

Public Const WM_TIMER As Long = &H113

' Processes, processes, threads
Public Const PROCESSOR_1 As Long = &H1
Public Const PROCESSOR_2 As Long = &H2
Public Const PROCESSOR_3 As Long = &H4
Public Const PROCESSOR_4 As Long = &H8
Public Const PROCESSOR_5 As Long = &H10
Public Const PROCESSOR_6 As Long = &H20
Public Const PROCESSOR_7 As Long = &H40
Public Const PROCESSOR_8 As Long = &H80

' For SHGetFolderPath - only some are supported (the others are commented out see - http://msdn.microsoft.com/en-us/library/bb762181%28VS.85%29.aspx)
Public Enum CSIDL_VALUES
'    CSIDL_DESKTOP = &H0
'    CSIDL_INTERNET = &H1
'    CSIDL_PROGRAMS = &H2
'    CSIDL_CONTROLS = &H3
'    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
'    CSIDL_FAVORITES = &H6
'    CSIDL_STARTUP = &H7
'    CSIDL_RECENT = &H8
'    CSIDL_SENDTO = &H9
'    CSIDL_BITBUCKET = &HA
'    CSIDL_STARTMENU = &HB
'    CSIDL_MYDOCUMENTS = &HC
'    CSIDL_MYMUSIC = &HD
'    CSIDL_MYVIDEO = &HE
'    CSIDL_DESKTOPDIRECTORY = &H10
'    CSIDL_DRIVES = &H11
'    CSIDL_NETWORK = &H12
'    CSIDL_NETHOOD = &H13
'    CSIDL_FONTS = &H14
'    CSIDL_TEMPLATES = &H15
'    CSIDL_COMMON_STARTMENU = &H16
'    CSIDL_COMMON_PROGRAMS = &H17
'    CSIDL_COMMON_STARTUP = &H18
'    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
'    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
'    CSIDL_ALTSTARTUP = &H1D
'    CSIDL_COMMON_ALTSTARTUP = &H1E
'    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
'    CSIDL_PROFILE = &H28
'    CSIDL_SYSTEMX86 = &H29
'    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
'    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
'    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
'    CSIDL_CONNECTIONS = &H31
'    CSIDL_COMMON_MUSIC = &H35
'    CSIDL_COMMON_PICTURES = &H36
'    CSIDL_COMMON_VIDEO = &H37
'    CSIDL_RESOURCES = &H38
'    CSIDL_RESOURCES_LOCALIZED = &H39
'    CSIDL_COMMON_OEM_LINKS = &H3A
'    CSIDL_CDBURN_AREA = &H3B
'    CSIDL_COMPUTERSNEARME = &H3D
'    CSIDL_FLAG_PER_USER_INIT = &H800
'    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
'    CSIDL_FLAG_MASK = &HFF00
End Enum

Public Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Public Const SHGFP_TYPE_DEFAULT = &H1

Private Const TWO_TO_THE_31 As Double = 2# ^ 31
Private Const TWO_TO_THE_32 As Double = 2# ^ 32
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

' GUID
Declare Function apiCoCreateGuid Lib "ole32.dll" Alias "CoCreateGuid" (pguid As GUID) As Long
Declare Function apiStringFromGUID2 Lib "ole32.dll" Alias "StringFromGUID2" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

' DLL
Declare Function apiGetModuleHandleA Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function apiGetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' VB
Declare Function apiVarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (var() As Any) As Long
Declare Function apiVarPtr Lib "msvbvm60.dll" Alias "VarPtr" (var As Any) As Long

' CopyMemory
Declare Function apiCopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As VOID
Declare Function apiZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long) As VOID
Declare Function apiFillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal length As Long, ByVal fill As Byte) As VOID
' Declare Sub CopyPointer Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Long, ByRef source() As Double, ByVal length As Long)
' Declare Sub CopySafeArray2DDesc Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As SAFEARRAY_2D, ByVal source As Long, ByVal length As Long)
' Declare Sub CopyMemoryWrite Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
' Declare Sub CopyMemoryRead Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)

' Stopwatch
Declare Function apiQueryPerformanceCounter Lib "kernel32.dll" Alias "QueryPerformanceCounter" (lpPerformanceCount As LARGE_INTEGER) As BOOL
Declare Function apiQueryPerformanceFrequency Lib "kernel32.dll" Alias "QueryPerformanceFrequency" (lpFrequency As LARGE_INTEGER) As BOOL
Declare Function apiQueryPerformanceFrequency_Currency Lib "kernel32.dll" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As BOOL
Declare Function apiQueryPerformanceCounter_Currency Lib "kernel32.dll" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As BOOL

' Timers
Declare Function apiSetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerProc As Long) As Long
Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long) As BOOL

' Processes, processes, threads
Declare Function apiGetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Declare Function apiSetProcessAffinityMask Lib "kernel32" Alias "SetProcessAffinityMask" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As BOOL

' String functions
Declare Function apiLStrLenA Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Declare Function apiLStrCpyA Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Declare Function apiLStrLenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

' Uncategorised
Declare Function apiMessageBoxA Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Sub apiOutputDebugStringA Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Declare Function apiSHGetFolderPathA Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As HRESULT
' SHGetSpecialFolderPathA is replaced by SHGetFolderPathA
Declare Function apiSHGetSpecialFolderPathA Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As BOOL

' Environment variables
Declare Function apiGetEnvironmentVariableA Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function apiSetEnvironmentVariableA Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long



Function DBCDoubleAsUnsignedLong(aDouble As Double) As Long
    If aDouble < 0 Then Err.Raise 0
    If aDouble >= TWO_TO_THE_31 Then
        DBCDoubleAsUnsignedLong = CLng(aDouble - TWO_TO_THE_32)
    Else
        DBCDoubleAsUnsignedLong = CLng(aDouble)
    End If
End Function

Function DBFunctionPointer(pFunction As Long) As Long          'allows AddressOf operator to set a value in a structure
    DBFunctionPointer = pFunction
End Function

'Returns the hi word from a double word.
Function DBHiWord(DWORD As Long) As Long
    If (DWORD And &H80000000) = &H80000000 Then
        DBHiWord = ((DWORD And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        DBHiWord = (DWORD And &HFFFF0000) \ &H10000
    End If
End Function

'Returns the low word from a double word.
Function DBLoWord(DWORD As Long) As Long
    DBLoWord = (DWORD And &HFFFF&)
End Function

'Receives a string pointer and it turns it into a regular string.
Function DBStringFromPointer(ByVal lPointer As Long) As String
    Dim temp As String, retval As Long
    temp = String$(apiLStrLenA(ByVal lPointer), 0)
    retval = apiLStrCpyA(ByVal temp, ByVal lPointer)
    If retval Then DBStringFromPointer = temp
End Function

'The function takes an unsigned Integer and converts it to a Long for display or arithmetic purposes
Function DBUnsignedToInteger(UINT As Long) As Integer
    If UINT < 0 Or UINT >= OFFSET_2 Then Error 6    ' Overflow
    If UINT <= MAXINT_2 Then
        DBUnsignedToInteger = UINT
    Else
        DBUnsignedToInteger = UINT - OFFSET_2
    End If
End Function

'The function takes a Long containing a value in the range of an unsigned Integer and returns an Integer that you can pass to an API that requires an unsigned Integer
Function DBIntegerToUnsigned(INT16 As Integer) As Long
    If INT16 < 0 Then
        DBIntegerToUnsigned = INT16 + OFFSET_2
    Else
        DBIntegerToUnsigned = INT16
    End If
End Function

Function DBNewGUID() As String
    Dim aGUID As GUID, buffer() As Byte, GUIDStringLength As Long
    Const bufferSize As Long = 40
    apiCoCreateGuid aGUID
    ReDim buffer(0 To (bufferSize * 2) - 1) As Byte
    GUIDStringLength = apiStringFromGUID2(aGUID, VarPtr(buffer(0)), bufferSize)
    DBNewGUID = Left$(buffer, GUIDStringLength - 1)
End Function

