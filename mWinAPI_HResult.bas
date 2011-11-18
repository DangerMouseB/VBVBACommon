Attribute VB_Name = "mWinAPI_HResult"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit
Option Private Module

' HRESULT
Public Type HRESULT
    HRESULT As Long
End Type

' http://msdn.microsoft.com/en-us/library/ms688560%28VS.85%29.aspx
' Because interface methods are virtual, it is not possible for a caller to know the full set of values that may be returned from any one call. One implementation of a method may return five values; another may return eight.

' The documentation lists common values that may be returned for each method; these are the values that you must check for and handle in your code because they have special meanings. Other values may be returned,
' but because they are not meaningful, you do not need to write special code to handle them. A simple check for zero or nonzero is adequate.

' HRESULT Values

' The return value of COM functions and methods is an HRESULT. The values of some HRESULTs have been changed in COM to eliminate all duplication and overlapping with the system error codes. Those that
' duplicate system error codes have been changed to FACILITY_WIN32, and those that overlap remain in FACILITY_NULL. Common HRESULT values and their values are listed in the following table.

' FACILITY_NULL codes
Public Const S_OK As Long = &H0                                     ' The method succeeded. If a boolean return value is expected, the returned value is TRUE.
Public Const S_FALSE As Long = &H1                                 ' The method succeeded and returned the boolean value FALSE

Public Const E_PENDING As Long = &H8000000A                ' The data necessary to complete the operation is not yet available
Public Const E_NOTIMPL As Long = &H80004001                 ' The method is not implemented
Public Const E_NOINTERFACE As Long = &H80004002         ' The QueryInterface method did not recognize the requested interface. The interface is not supported
Public Const E_POINTER As Long = &H80004003                 ' An invalid pointer was used
Public Const E_ABORT As Long = &H80004004                   ' The operation was aborted because of an unspecified error
Public Const E_FAIL As Long = &H80004005                       ' An unspecified failure has occurred
Public Const E_UNEXPECTED As Long = &H8000FFFF           ' A catastrophic failure has occurred

' FACILITY_WIN32 codes
Public Const E_ACCESSDENIED As Long = &H80070005      ' A general access-denied error
Public Const E_HANDLE As Long = &H80070006                  ' An invalid handle was used
Public Const E_OUTOFMEMORY As Long = &H8007000E       ' The method failed to allocate necessary memory
Public Const E_INVALIDARG As Long = &H80070057           ' One or more arguments are not valid

Public Const E_WIN32_042C As Long = &H8007042C            ' The dependency service or group failed to start.

' http://blogs.msdn.com/ericlippert/archive/2003/10/22/53267.aspx
Public Const SCRIPT_E_RECORDED As Long = &H86664004 ' this is how we internally track whether the details of an error have been recorded in the error object or not.  We need a way to say "yes, there was an error, but do not attempt to record information about it again."
Public Const SCRIPT_E_PROPAGATE As Long = &H80020102   'another internal code that we use to track the case where a recorded error is being propagated up the call stack to a waiting catch handler.
Public Const SCRIPT_E_REPORTED  As Long = &H80020101  'the script engines return this to the host when there has been an unhandled error that the host has already been informed about via OnScriptError.

' http://www.codeguru.cn/vc&mfc/APracticalGuideUsingVisualCandATL/39.htm
' http://msdn.microsoft.com/en-us/library/ms690088%28VS.85%29.aspx
' http://vb.mvps.org/hardcore/html/howtoraiseerrors.htm

Public Const H_ERROR As Long = &H80000000

Public Const FACILITY_NULL As Long = &H0
Public Const FACILITY_RPC As Long = &H10000
Public Const FACILITY_DISPATCH As Long = &H20000
Public Const FACILITY_STORAGE As Long = &H30000
Public Const FACILITY_ITF As Long = &H40000                         ' internally used by VB -> VBObjectError
Public Const FACILITY_NOT_DEFINED_5 As Long = &H50000
Public Const FACILITY_NOT_DEFINED_6 As Long = &H60000
Public Const FACILITY_WIN32 As Long = &H70000
Public Const FACILITY_WINDOWS As Long = &H80000
Public Const FACILITY_SSPI As Long = &H90000
Public Const FACILITY_SECURITY As Long = &H90000
Public Const FACILITY_CONTROL As Long = &HA0000                 ' used by VB to raise accross a COM call?
Public Const FACILITY_CERT As Long = &HB0000
Public Const FACILITY_INTERNET As Long = &HC0000
Public Const FACILITY_MEDIASERVER As Long = &HD0000
Public Const FACILITY_MSMQ As Long = &HE0000
Public Const FACILITY_SETUPAPI As Long = &HF0000
Public Const FACILITY_SCARD As Long = &H100000
Public Const FACILITY_COMPLUS As Long = &H110000
Public Const FACILITY_AAF As Long = &H120000
Public Const FACILITY_URT As Long = &H130000
Public Const FACILITY_ACS As Long = &H140000
Public Const FACILITY_DPLAY As Long = &H150000
Public Const FACILITY_UMI As Long = &H160000
Public Const FACILITY_SXS As Long = &H170000
Public Const FACILITY_WINDOWS_CE As Long = &H180000
Public Const FACILITY_HTTP As Long = &H190000
Public Const FACILITY_BACKGROUNDCOPY As Long = &H200000
Public Const FACILITY_CONFIGURATION As Long = &H210000
Public Const FACILITY_STATE_MANAGEMENT As Long = &H220000
Public Const FACILITY_METADIRECTORY As Long = &H230000


Function CHRESULT(aResult As Long) As HRESULT
    CHRESULT.HRESULT = aResult
End Function




