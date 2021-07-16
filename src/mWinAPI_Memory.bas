Attribute VB_Name = "mWinAPI_Memory"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit
Option Private Module

' STATUS_NO_MEMORY
' STATUS_ACCESS_VIOLATION

' HEAP_GENERATE_EXCEPTIONS 0x00000004
' HEAP_NO_SERIALIZE 0x00000001
' HEAP_ZERO_MEMORY 0x00000008


Public Const GMEM_FIXED = &H0

Public Const PAGE_READWRITE = &H4

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SECTION_QUERY = &H1
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const SECTION_MAP_EXECUTE = &H8
Public Const SECTION_EXTEND_SIZE = &H10
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Public Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

' start DB - http://msdn.microsoft.com/en-us/library/aa374896%28VS.85%29.aspx

Public Const FILE_MAP_READ As Long = &H4
Public Const FILE_MAP_WRITE As Long = &H2

' Generic Access Rights - defined in Winnt.h
Public Const GENERIC_ALL As Long = &H10000000
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000
Public Const GENERIC_EXECUTE As Long = &H20000000

'' Standard Access Rights - defined in Winnt.h
'STANDARD_DELETE                 ' The right to delete the object.
'STANDARD_READ_CONTOL      ' The right to read the information in the object's security descriptor, not including the information in the system access control list (SACL).
'STANDARD_SYNCHRONIZE      ' The right to use the object for synchronization. This enables a thread to wait until the object is in the signaled state. Some object types do not support this access right.
'STANDARD_WRITE_DAC          ' The right to modify the discretionary access control list (DACL) in the object's security descriptor.
'STANDARD_WRITE_OWNER      ' The right to change the owner in the object's security descriptor.
'
'' Winnt.h also defines the following combinations of the standard access rights constants
'STANDARD_RIGHTS_ALL = Delete Or READ_CONTROL Or WRITE_DAC Or WRITE_OWNER Or SYNCHRONIZE
'STANDARD_RIGHTS_EXECUTE = READ_CONTROL
'STANDARD_RIGHTS_READ = READ_CONTROL
'STANDARD_RIGHTS_REQUIRED = Delete Or READ_CONTROL Or WRITE_DAC Or WRITE_OWNER
'STANDARD_RIGHTS_WRITE = READ_CONTROL

Declare Function apiGetProcessHeap Lib "kernel32" Alias "GetProcessHeap" () As Long
Declare Function apiHeapAlloc Lib "kernel32" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function apiHeapFree Lib "kernel32" Alias "HeapFree" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function apiGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function apiGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long

Declare Function apiGlobalLock Lib "kernel32" Alias "globallock" (ByVal hMem As Long) As Long
Declare Function apiGlobalUnlock Lib "kernel32" Alias "globalunlock" (ByVal hMem As Long) As Long


'CoTaskMemAlloc
'CoTaskMemFree


Declare Function apiCreateFileMappingA Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Declare Function apiOpenFileMappingA Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function apiMapViewOfFile Lib "kernel32" Alias "MapViewOfFile" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Declare Function apiUnmapViewOfFile Lib "kernel32" Alias "UnmapViewOfFile" (ByVal lpBaseAddress As Long) As Long

