Attribute VB_Name = "mWinAPI_RegistryAndIniFiles"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit
Option Private Module

' Constants for Registry top-level keys
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' RegCreateKeyEx options
Public Const REG_OPTION_NON_VOLATILE = 0

' RegCreateKeyEx Disposition
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

' Registry data types
Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

' Registry security attributes
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8

' Registry
Declare Function apiRegEnumValueA Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function apiRegDeleteValueA Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function apiRegDeleteKeyA Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function apiRegOpenKeyExA Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function apiRegCreateKeyExA Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function apiRegQueryValueExA Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function apiRegSetValueExA Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function apiRegEnumKeyA Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function apiRegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long

' INI files
Declare Function apiGetPrivateProfileIntA Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function apiGetPrivateProfileSectionA Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As DWORD
Declare Function apiGetPrivateProfileSectionNamesA Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As DWORD
Declare Function apiGetPrivateProfileStringA Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As DWORD
Declare Function apiGetPrivateProfileStructA Lib "kernel32.dll" Alias "GetPrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Any, ByVal uSizeStruct As Long, ByVal szFile As String) As BOOL
' GetProfileInt             - accesses win.ini
' GetProfileSection      - accesses win.ini
' GetProfileString        - accesses win.ini
Declare Function apiWritePrivateProfileSectionA Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As BOOL
Declare Function apiWritePrivateProfileStringA Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As BOOL
Declare Function apiWritePrivateProfileStructA Lib "kernel32.dll" Alias "WritePrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Any, ByVal uSizeStruct As Long, ByVal szFile As String) As BOOL
' WriteProfileSection    - accesses win.ini
' WriteProfileString      - accesses win.ini

