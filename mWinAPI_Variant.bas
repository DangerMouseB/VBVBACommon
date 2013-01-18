Attribute VB_Name = "mWinAPI_Variant"
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

' SAFEARRAY
Public Type SAFEARRAYBOUND
    cElements   As Long      ' # of elements in the array dimension
    lLbound     As Long      ' lower bounds of the array dimension
End Type

Public Type SAFEARRAY
    cDims       As Integer    ' Count of dimensions in this array
    fFeatures   As Integer    ' Flags used by the SAFEARRAY routines
    cbElements  As Long      ' Size of an element of the array.
    cLocks      As Long         ' Number of times the array has been locked without corresponding unlock.
    pvData      As Long        ' Pointer to the data.
    rgSABound() As SAFEARRAYBOUND   ' One bound for each dimension.
    ' An array can have max 60 dimensions, only the first cDims items will be used
    ' note that rgsabound elements are in reverse order,
    '  e.g. for a 2-dimensional array, rgsabound(1) holds info about columns, and rgsabound(2) about rows
End Type

Public Const FADF_AUTO As Integer = &H1              ' An array that is allocated on the stack.
Public Const FADF_STATIC As Integer = &H2           ' An array that is statically allocated.
Public Const FADF_EMBEDDED As Integer = &H4      ' An array that is embedded in a structure.
Public Const FADF_FIXEDSIZE As Integer = &H10     ' An array that may not be resized or reallocated.
Public Const FADF_RECORD As Integer = &H20         ' An array that contains records. When set, there will be a pointer to the IRecordinfo interface at negative offset 4 in the array descriptor.
Public Const FADF_HAVEIID As Integer = &H40        ' An array that has an IID identifying interface. When set, there will be a GUID at negative offset 16 in the safe array descriptor. Flag is set only when FADF_DISPATCH or FADF_UNKNOWN is also set.
Public Const FADF_HAVEVARTYPE As Integer = &H80  ' An array that has a VT type. When set, there will be a VT tag at negative offset 4 in the array descriptor that specifies the element type.
Public Const FADF_BSTR As Integer = &H100           ' An array of BSTRs.
Public Const FADF_UNKNOWN As Integer = &H200    ' An array of IUnknown*.
Public Const FADF_DISPATCH As Integer = &H400    ' An array of IDispatch*.
Public Const FADF_VARIANT As Integer = &H800      ' An array of VARIANTs.
Public Const FADF_RESERVED As Integer = &HF0E8  ' Bits reserved for future use.

' SafeArrayLock
Public Const DISP_E_ARRAYISLOCKED As Long = &H8002000D

Declare Function apiSafeArrayAllocData Lib "oleaut32.dll" Alias "SafeArrayAllocData" (ByVal pSAFEARRAY As Long) As HRESULT
Declare Function apiSafeArrayDestroy Lib "oleaut32.dll" Alias "SafeArrayDestroy" (ByVal pSAFEARRAY As Long) As HRESULT
Declare Function apiSafeArrayDestroyData Lib "oleaut32.dll" Alias "SafeArrayDestroyData" (ByVal pSAFEARRAY As Long) As HRESULT
Declare Function apiSafeArrayGetVartype Lib "oleaut32.dll" Alias "SafeArrayGetVartype" (ByVal pSAFEARRAY As Long, varType As Any) As HRESULT
Declare Function apiSafeArrayLock Lib "oleaut32.dll" Alias "SafeArrayLock" (ByVal pSAFEARRAY As Long) As HRESULT
Declare Function apiSafeArrayUnlock Lib "oleaut32.dll" Alias "SafeArrayUnlock" (ByVal pSAFEARRAY As Long) As HRESULT

'Private Declare Sub SafeArrayAccessData Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef ppvData As Any)
'Private Declare Sub SafeArrayAllocData Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)
'Private Declare Sub SafeArrayAllocDescriptor Lib "oleaut32.dll" (ByVal cDims As Long, ByRef ppsaOut As SAFEARRAY)
'Private Declare Sub SafeArrayAllocDescriptorEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef ppsaOut As SAFEARRAY)
'Private Declare Sub SafeArrayCopy Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef ppsaOut As SAFEARRAY)
'Private Declare Sub SafeArrayCopyData Lib "oleaut32.dll" (ByRef psaSource As SAFEARRAY, ByRef psaTarget As SAFEARRAY)
'Private Declare Function SafeArrayCreate Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND) As Long
'Private Declare Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND, ByRef pvExtra As Any) As Long
'Private Declare Function SafeArrayCreateVector Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long) As Long
'Private Declare Function SafeArrayCreateVectorEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long, ByRef pvExtra As Any) As Long
'Private Declare Sub SafeArrayDestroy Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)
'Private Declare Sub SafeArrayDestroyData Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)
'Private Declare Sub SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)
'Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef psa As SAFEARRAY) As Long
'Private Declare Sub SafeArrayGetElement Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef rgIndices As Long, ByRef pv As Any)
'Private Declare Function SafeArrayGetElemsize Lib "oleaut32.dll" (ByRef psa As SAFEARRAY) As Long
'Private Declare Sub SafeArrayGetIID Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef pguid As GUID)
'Private Declare Sub SafeArrayGetLBound Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByVal nDim As Long, ByRef plLbound As Long)
'Private Declare Sub SafeArrayGetRecordInfo Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef prinfo As Long)
'Private Declare Sub SafeArrayGetUBound Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByVal nDim As Long, ByRef plUbound As Long)
'Private Declare Sub SafeArrayLock Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)
'Private Declare Sub SafeArrayPtrOfIndex Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef rgIndices As Long, ByRef ppvData As Any)
'Private Declare Sub SafeArrayPutElement Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef rgIndices As Long, ByRef pv As Any)
'Private Declare Sub SafeArrayRedim Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef psaboundNew As SAFEARRAYBOUND)
'Private Declare Sub SafeArraySetIID Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByVal GUID As Long)
'Private Declare Sub SafeArraySetRecordInfo Lib "oleaut32.dll" (ByRef psa As SAFEARRAY, ByRef prinfo As Long)
'Private Declare Sub SafeArrayUnaccessData Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)
'Private Declare Sub SafeArrayUnlock Lib "oleaut32.dll" (ByRef psa As SAFEARRAY)

' VARIANT
Public Const VT_EMPTY = 0                    '
Public Const VT_NULL = 1                      '
Public Const VT_I2 = 2                          ' Integer
Public Const VT_I4 = 3                          ' Long
Public Const VT_R4 = 4                          ' Single
Public Const VT_R8 = 5                          ' Double
Public Const VT_CY = 6                          ' Currency
Public Const VT_DATE = 7                      ' Date
Public Const VT_BSTR = 8                      ' String
Public Const VT_DISPATCH = 9               ' vbObject
Public Const VT_ERROR = 10                  ' vbError
Public Const VT_BOOL = 11                    ' Boolean
Public Const VT_VARIANT = 12               ' Variant (used only with arrays of variants)
Public Const VT_UNKNOWN = 13             ' vbDataObject
Public Const VT_DECIMAL = 14               ' Decimal
Public Const VT_I1 = 16                         '
Public Const VT_UI1 = 17                        ' Byte
Public Const VT_UI2 = 18                        '
Public Const VT_UI4 = 19                        '
Public Const VT_I8 = 20                          '
Public Const VT_UI8 = 21                        ' LONGLONG
Public Const VT_INT = 22                        '
Public Const VT_UINT = 23                      '
Public Const VT_VOID = 24                     '
Public Const VT_HRESULT = 25                '
Public Const VT_PTR = 26                        '
Public Const VT_SAFEARRAY = 27             '
Public Const VT_CARRAY = 28                  '
Public Const VT_USERDEFINED = 29          '
Public Const VT_LPSTR = 30                     '
Public Const VT_LPWSTR = 31                  '
Public Const VT_RECORD = 36
Public Const VT_FILETIME = 64                 '
Public Const VT_BLOB = 65                      '
Public Const VT_STREAM = 66                   '
Public Const VT_STORAGE = 67                 '
Public Const VT_STREAMED_OBJECT = 68  '
Public Const VT_STORED_OBJECT = 69      '
Public Const VT_BLOB_OBJECT = 70          '
Public Const VT_CF = 71                           '
Public Const VT_CLSID = 72                      '

Public Const VT_BSTR_BLOB = &HFFF
Public Const VT_VECTOR = &H1000            '
Public Const VT_ARRAY = &H2000               '
Public Const VT_BYREF = &H4000               '
Public Const VT_RESERVED = &H8000         '
Public Const VT_ILLEGAL = &HFFFF              '
Public Const VT_ILLEGALMASKED = &HFFF   '
Public Const VT_TYPEMASK = &HFFF            '



