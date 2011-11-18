Attribute VB_Name = "mWinAPI_FileHanding"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit
Option Private Module

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' dwDesiredAccess - CreateFileA
Public Const dwDesiredAccess_query_attributes As Long = &H0
Public Const dwDesiredAccess_GENERIC_READ As Long = &H80000000
Public Const dwDesiredAccess_GENERIC_WRITE As Long = &H40000000
Public Const dwDesiredAccess_DELETE As Long = -1
' and a whole bunch more...

' dwShareMode - CreateFileA
Public Const dwShareMode_none As Long = 0
Public Const dwShareMode_FILE_SHARE_DELETE As Long = -1
Public Const dwShareMode_FILE_SHARE_READ As Long = &H1
Public Const dwShareMode_FILE_SHARE_WRITE As Long = &H2

' dwCreationDisposition - CreateFileA
Public Const dwCreationDisposition_none As Long = 0
Public Const dwCreationDisposition_CREATE_NEW As Long = &H1
Public Const dwCreationDisposition_CREATE_ALWAYS As Long = &H2
Public Const dwCreationDisposition_OPEN_EXISTING As Long = &H3
Public Const dwCreationDisposition_OPEN_ALWAYS As Long = &H4
Public Const dwCreationDisposition_TRUNCATE_EXISTING As Long = &H5

' dwFlagsAndAttributes - CreateFileA
Public Const dwFlagsAndAttributes_none As Long = &H0
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_ENCRYPTED As Long = &H4000
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_NOT_CONTENT_INDEXED As Long = &H2000
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_OFFLINE As Long = &H1000
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_SYSTEM As Long = &H4
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const dwFlagsAndAttributes_FILE_ATTRIBUTE_TEMPORARY As Long = &H100
' and a whole bunch more...

' dwMoveMethod - SetFilePointer
Public Const dwMoveMethod_FILE_BEGIN As Long = &H0
Public Const dwMoveMethod_FILE_CURRENT As Long = &H1
Public Const dwMoveMethod_FILE_END As Long = &H2


Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As BOOL
Declare Function apiCopyFileA Lib "kernel32" Alias "CopyFileA" (ByRef LPCSTR_lpExistingFileName As String, ByRef LPCSTR_lpNewFileName As String, ByVal BOOL_bFailIfExists As Long) As Long
Declare Function apiCopyFileW Lib "kernel32" Alias "CopyFileW" (ByRef LPCWSTR_lpExistingFileName As String, ByRef LPCWSTR_lpNewFileName As String, ByVal BOOL_bFailIfExists As Long) As Long
Declare Function apiCreateFileA Lib "kernel32" Alias "CreateFileA" (ByRef LPCSTR_lpFileName As String, ByVal DWORD_dwDesiredAccess As Long, ByVal DWORD_dwShareMode As Long, ByRef LPSECURITY_ATTRIBUTES_lpSecurityAttributes As Any, ByVal DWORD_dwCreationDisposition As Long, ByVal DWORD_dwFlagsAndAttributes As Long, ByVal HANDLE_hTemplateFile As Long) As HANDLE
Declare Function apiFlushFileBuffers Lib "kernel32" Alias "FlushFileBuffers" (ByVal hFile As Long) As BOOL
Declare Function apiGetDiskFreeSpaceA Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByRef LPCTSTR_lpRootPathName, ByRef LPDWORD_lpSectorsPerCluster As Long, ByRef LPDWORD_lpBytesPerSector As Long, ByRef LPDWORD_lpNumberOfFreeClusters As Long, ByRef LPDWORD_lpTotalNumberOfClusters As Long) As BOOL
Declare Function apiGetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceEx" (ByRef LPCTSTR_lpDirectoryName As Long, ByRef PULARGE_INTEGER_lpFreeBytesAvailable As Long, ByRef PULARGE_INTEGER_lpTotalNumberOfBytes As Long, ByRef PULARGE_INTEGER_lpTotalNumberOfFreeBytes As Long) As BOOL
Declare Function apiGetTempPathA Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function apiSetFilePointer Lib "kernel32" Alias "SetFilePointer" (ByVal HANDLE_hFile As Long, ByVal LONG_lDistanceToMove As Long, PLONG_lpDistanceToMoveHigh As Long, ByVal DWORD_dwMoveMethod As Long) As DWORD
Declare Function apiSetFilePointerEx Lib "kernel32" Alias "SetFilePointerEx" (ByVal HANDLE_hFile As Long, ByVal LARGE_INTEGER_liDistanceToMove As Currency, PLARGE_INTEGER_lpNewFilePointer As Any, ByVal DWORD_dwMoveMethod As Long) As BOOL
Declare Function apiSetEndOfFile Lib "kernel32" Alias "SetEndOfFile" (ByVal HANDLE_hFile As Long) As BOOL
Declare Function apiGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal HANDLE_hFile As Long, LPDWORD_lpFileSizeHigh As Long) As DWORD
Declare Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal HANDLE_hFile As Long, ByVal PLARGE_INTEGER_lpFileSize As Long) As BOOL
Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal HANDLE_hFile As Long, LPVOID_lpBuffer As Any, ByVal DWORD_nNumberOfBytesToRead As Long, LPDWORD_lpNumberOfBytesRead As Long, ByVal LPOVERLAPPED_lpOverlapped As Long) As BOOL


' need cleaning up
Declare Function apiCreateDirectoryA Lib "kernel32" Alias "CreateDirectoryA" (ByRef lpPathName As String, lpSecurityAttributes As Long) As Long
Declare Function apiDeleteFileA Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function apiFileTimeToLocalFileTime Lib "kernel32" Alias "FileTimeToLocalFileTime" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function apiFileTimeToSystemTime Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function apiGetFileInformationByHandle Lib "kernel32" Alias "GetFileInformationByHandle" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Declare Function apiGetFileTime Lib "kernel32" Alias "GetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function apiGetTempFileNameA Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function apiMoveFileA Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function apiMoveFileExA Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Declare Function apiSetFileAttributesA Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function apiSetVolumeLabelA Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Declare Function apiSHFileOperationA Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal LPOVERLAPPED As Long) As BOOL

