Attribute VB_Name = "mWinAPI_SxS"
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

Type ACTCTX
    cbSize As Long
    dwFlags As Long
    lpSource As String
    wProcessorArchitecture As Integer
    wLangId As Integer
    lpAssemblyDirectory As String
    lpResourceName As String
    lpApplicationName As String
    hModule As Long
End Type

Const S_OK = 1

Const ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID = &H1
Const ACTCTX_FLAG_LANGID_VALID = &H2
Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID = &H4
Const ACTCTX_FLAG_RESOURCE_NAME_VALID = &H8
Const ACTCTX_FLAG_SET_PROCESS_DEFAULT = &H10
Const ACTCTX_FLAG_APPLICATION_NAME_VALID = &H20
Const ACTCTX_FLAG_SOURCE_IS_ASSEMBLYREF = &H40
Const ACTCTX_FLAG_HMODULE_VALID = &H80

'BOOL ActivateActCtx( __in   HANDLE hActCtx, __out  ULONG_PTR *lpCookie);
Declare Function apiActivateActCtx Lib "kernel32.dll" Alias "ActivateActCtx" (ByVal hActCtx As Long, ByRef lpCookie As Long) As BOOL

'void AddRefActCtx( __in  HANDLE hActCtx);
'Declare Sub apiAddRefActCtx Lib "Kernel32.dll" Alias "AddRefActCtx" (ByVal hActCtx As Long)

'HANDLE CreateActCtx( __inout  PACTCTX pActCtx);
Declare Function apiCreateActCtxW Lib "kernel32.dll" Alias "CreateActCtxW" (ByVal pActCtx As Long) As HANDLE
Declare Function apiCreateActCtxA Lib "kernel32.dll" Alias "CreateActCtxA" (ByVal pActCtx As Long) As HANDLE

'BOOL DeactivateActCtx( __in  DWORD dwFlags, __in  ULONG_PTR ulCookie);
Declare Function apiDeactivateActCtx Lib "kernel32.dll" Alias "DeactivateActCtx" (ByVal dwFlags As Long, ByVal ulCookie As Long) As BOOL

'BOOL FindActCtxSectionGuid( __in   DWORD dwFlags, __in   const GUID *lpExtensionGuid, __in   ULONG ulSectionId, __in   const GUID *lpGuidToFind, __out  PACTCTX_SECTION_KEYED_DATA ReturnedData);
'Declare Function apiFindActCtxSectionGuid Lib "Kernel32.dll" Alias "FindActCtxSectionGuid" () As BOOL

'BOOL FindActCtxSectionString( __in   DWORD dwFlags, __in   const GUID *lpExtensionGuid, __in   ULONG ulSectionId, __in   LPCTSTR lpStringToFind, __out  PACTCTX_SECTION_KEYED_DATA ReturnedData);
'Declare Function apiFindActCtxSectionString Lib "Kernel32.dll" Alias "FindActCtxSectionString" () As BOOL

'BOOL GetCurrentActCtx( __out  HANDLE *lphActCtx);
'Declare Function apiGetCurrentActCtx Lib "Kernel32.dll" Alias "GetCurrentActCtx" () As BOOL

'void IsolationAwareCleanup(void);
'Declare Sub apiIsolationAwareCleanup Lib "Kernel32.dll" Alias "IsolationAwareCleanup" ()

'BOOL QueryActCtxW( __in       DWORD dwFlags, __in       HANDLE hActCtx, __in       PVOID pvSubInstance, __in       ULONG ulInfoClass, __out      PVOID pvBuffer, __in_opt   SIZE_T cbBuffer, __out_opt  SIZE_T *pcbWrittenOrRequired);
'Declare Function apiQueryActCtxW Lib "Kernel32.dll" Alias "QueryActCtxW" () As BOOL

'BOOL QueryActCtxSettingsW( __in_opt   DWORD dwFlags, __in_opt   HANDLE hActCtx, __in_opt   PCWSTR settingsNameSpace, __in       PCWSTR settingName, __out      PWSTR pvBuffer, __in       SIZE_T dwBuffer, __out_opt  SIZE_T *pdwWrittenOrRequired);
'Declare Function apiQueryActCtxSettingsW Lib "Kernel32.dll" Alias "QueryActCtxSettingsW" () As BOOL

'void ReleaseActCtx( __in  HANDLE hActCtx);
Declare Sub apiReleaseActCtx Lib "kernel32.dll" Alias "ReleaseActCtx" (ByVal hActCtx As Long)

'BOOL ZombifyActCtx( __in  HANDLE hActCtx);
'Declare Function apiZombifyActCtx Lib "Kernel32.dll" Alias "ZombifyActCtx" () As BOOL

Function LoadCOMDLL(manifestFilename As String) As String
    Dim struct As ACTCTX, hActCtx As Long, lpCookie As Long, result As Long, obj As Object   ', loader As New QLoader.CLoader
    struct.cbSize = Len(struct)
    struct.lpAssemblyDirectory = ""
    struct.lpSource = manifestFilename
    hActCtx = apiCreateActCtxW(VarPtr(struct)).HANDLE
    If hActCtx < 0 Then LoadCOMDLL = "apiCreateActCtxW failed": Exit Function
    If apiActivateActCtx(hActCtx, lpCookie).BOOL <> S_OK Then LoadCOMDLL = "apiActivateActCtx failed": Exit Function
End Function


