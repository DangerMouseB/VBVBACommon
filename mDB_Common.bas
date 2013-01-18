Attribute VB_Name = "mDB_Common"
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

' error reporting
Private Const MODULE_NAME As String = "mDB_Common"
Private Const MODULE_VERSION As String = "0.0.1.0"

Public Const GENERAL_ERROR_CODE As Long = 30000
Public Const UNHANDLED_CASE_ERROR_CODE As Long = 30001
Public Const NOT_YET_IMPLIMENTED_ERROR_CODE As Long = 30002

Private Const twoToThe8 As Long = 256
Private Const twoToThe16 As Currency = 65536@
Private Const twoToThe32 As Currency = 4294967296@
Private Const twoToThe48 As Currency = 281474976710656@

' = get.workspace(10000)             - ?
' = get.workspace(44)
' = get.workspace(7124)

' ctrl F11 - brings up the macro editor


'*************************************************************************************************************************************************************************************************************************************************
' array utilities
'*************************************************************************************************************************************************************************************************************************************************

Sub DBClearArray(anArray As Variant)
    Dim SA As SAFEARRAY, length As Long, i As Long, vType As Integer
    DBGetSafeArrayDetails anArray, SA, vType
    length = SA.cbElements
    For i = 1 To SA.cDims
        length = length * SA.rgSABound(i).cElements
    Next
    If length > 0 And SA.pvData > 0 Then apiZeroMemory ByVal SA.pvData, length
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' variant / object assignment
'*************************************************************************************************************************************************************************************************************************************************

Sub slet(oVar As Variant, inVar As Variant)
    If varType(inVar) = vbObject Then
        Set oVar = inVar
    Else
        oVar = inVar
    End If
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' array information
'*************************************************************************************************************************************************************************************************************************************************

Sub DBGetArrayDetails(arrayND As Variant, oDimension As Long, Optional oVType As Integer)
    Dim SA As SAFEARRAY
    DBGetSafeArrayDetails arrayND, SA, oVType
    oDimension = SA.cDims
End Sub

Sub DBGetArrayBounds(arrayND As Variant, dimension As Long, oLowerBound As Long, oUpperBound As Long, Optional oNumberOfElements As Variant)
    On Error GoTo exceptionHandler
    oLowerBound = LBound(arrayND, dimension)
    oUpperBound = UBound(arrayND, dimension)
    If Not IsMissing(oNumberOfElements) Then oNumberOfElements = oUpperBound - oLowerBound + 1
Exit Sub
exceptionHandler:
    oLowerBound = 0
    oUpperBound = -1
    If Not IsMissing(oNumberOfElements) Then oNumberOfElements = 0
End Sub

Function DBGetSafeArrayDetails(anArray As Variant, oSA As SAFEARRAY, Optional vType As Integer) As HRESULT
    Dim ptr As Long, fFeatures As Integer
    ptr = getSafeArrayPointer(anArray)
    If ptr = 0 Then DBGetSafeArrayDetails = CHRESULT(E_INVALIDARG): Exit Function
    apiCopyMemory oSA.cDims, ByVal ptr, 16         ' The fixed part of the SAFEARRAY structure is 16 bytes.
    If oSA.cDims > 0 Then
        ReDim oSA.rgSABound(1 To oSA.cDims)
        apiCopyMemory oSA.rgSABound(1), ByVal ptr + 16&, oSA.cDims * Len(oSA.rgSABound(1))
    End If
    apiSafeArrayGetVartype ptr, vType
End Function

Private Function getSafeArrayPointer(anArray As Variant) As Long
    Dim ptr As Long, vType As Integer
    If Not IsArray(anArray) Then Exit Function
    apiCopyMemory vType, anArray, 2                                                         ' Get the VARTYPE value from the first 2 bytes of the VARIANT structure
    apiCopyMemory ptr, ByVal VarPtr(anArray) + 8, 4                                    ' Get the pointer to the array descriptor (SAFEARRAY structure)   NOTE: A Variant's descriptor, padding & union take up 8 bytes.
    If (vType And VT_BYREF) <> 0 Then apiCopyMemory ptr, ByVal ptr, 4        ' Test if lp is a pointer or a pointer to a pointer and if so get real pointer to the array descriptor (SAFEARRAY structure)
    getSafeArrayPointer = ptr
End Function


'*************************************************************************************************************************************************************************************************************************************************
' array coercion
'*************************************************************************************************************************************************************************************************************************************************

Function DBCByteArray(anArray As Variant) As Byte()
    DBCByteArray = anArray
End Function

Function DBCIntegerArray(anArray As Variant) As Integer()
    DBCIntegerArray = anArray
End Function

Function DBCLongArray(anArray As Variant) As Long()
    DBCLongArray = anArray
End Function

Function DBCSingleArray(anArray As Variant) As Single()
    DBCSingleArray = anArray
End Function

Function DBCDoubleArray(anArray As Variant) As Double()
    DBCDoubleArray = anArray
End Function

Function DBCBooleanArray(anArray As Variant) As Boolean()
    DBCBooleanArray = anArray
End Function

Function DBCDateArray(anArray As Variant) As Date()
    DBCDateArray = anArray
End Function

Function DBCCurrencyArray(anArray As Variant) As Currency()
    DBCCurrencyArray = anArray
End Function

Function DBCStringArray(anArray As Variant) As String()
    DBCStringArray = anArray
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Environment Variables
'*************************************************************************************************************************************************************************************************************************************************

Function VBGetEnv(name As String) As String
    VBGetEnv = Environ$(name)
End Function

Function setEnv(name As String, value As String) As String
    Dim retval As Long
    retval = apiSetEnvironmentVariableA(name, value)
    setEnv = getEnv(name)
End Function

Function getEnv(name As String, Optional maxSize As Long = 2048) As String
    Dim buffer As String, length As Long
    buffer = String(8096, 0)
    length = apiGetEnvironmentVariableA(name, buffer, Len(buffer))
    getEnv = Left$(buffer, length)
End Function

'Function alterPath() As String
'    Dim s As Object, sysVars As Object, userVars As Object
'    Set s = CreateObject("WScript.Shell")
'    Set sysVars = s.Environment("System")
'    Set userVars = s.Environment("User")
'    sysVars("Path") = sysVars("Path") & ";c:\python24Mercury"
'    extendPath = sysVars("Path")
'End Function


'*************************************************************************************************************************************************************************************************************************************************
' array creation
'*************************************************************************************************************************************************************************************************************************************************

Function DBSubArray(anArray As Variant, Optional ByVal i1 As Variant, Optional ByVal i2 As Variant, Optional ByVal j1 As Variant, Optional ByVal j2 As Variant, Optional ByVal k1 As Variant, Optional ByVal k2 As Variant, Optional ByVal newType As Long = -1, Optional ByVal newI1 As Variant, Optional ByVal newJ1 As Variant, Optional ByVal newK1 As Variant) As Variant
    Dim nDimensions As Long, vType As Integer, answerND As Variant, dummy As Long, SA As SAFEARRAY
    Dim x As Long, x1 As Long, x2 As Long, y As Long, y1 As Long, y2 As Long, z As Long, z1 As Long, z2 As Long, newX1 As Long, newY1 As Long, newZ1 As Long, xOffset As Long, yOffset As Long, zOffset As Long
    
    Const METHOD_NAME As String = "DBSubArray"
    
    ' check dimensions
    If DBGetSafeArrayDetails(anArray, SA, vType).HRESULT <> S_OK Then DBErrors_raiseGeneralError ModuleSummary, METHOD_NAME, "Not an array"
    nDimensions = SA.cDims
    If nDimensions = 0 Then DBErrors_raiseGeneralError ModuleSummary, METHOD_NAME, "#Array has no data!"
    If nDimensions > 3 Then DBErrors_raiseGeneralError ModuleSummary, METHOD_NAME, "#Can't handle more than 3 dimensions!"

    ' fill in any missing parameters
    If IsMissing(i1) Then DBGetArrayBounds anArray, 1, x1, dummy Else x1 = i1
    If IsMissing(i2) Then DBGetArrayBounds anArray, 1, dummy, x2 Else x2 = i2
    If IsMissing(newI1) Then newX1 = 1 Else newX1 = newI1
    If newX1 <> x1 Then xOffset = newX1 - x1
    If nDimensions > 1 Then
        If IsMissing(j1) Then DBGetArrayBounds anArray, 2, y1, dummy Else y1 = j1
        If IsMissing(j2) Then DBGetArrayBounds anArray, 2, dummy, y2 Else y2 = j2
        If IsMissing(newJ1) Then newY1 = 1 Else newY1 = newJ1
        If newY1 <> y1 Then yOffset = newY1 - y1
    End If
    If nDimensions > 2 Then
        If IsMissing(k1) Then DBGetArrayBounds anArray, 3, z1, dummy Else z1 = k1
        If IsMissing(k2) Then DBGetArrayBounds anArray, 3, dummy, z2 Else z2 = k2
        If IsMissing(newK1) Then newZ1 = 1 Else newZ1 = newK1
        If newZ1 <> z1 Then zOffset = newZ1 - z1
    End If

    If vType = newType Or newType = -1 Then
        ' no type conversion copy
        Select Case nDimensions
            Case 1
                DBCreateNewArrayOfType answerND, vType, x1 + xOffset, x2 + xOffset
                For x = x1 To x2
                    answerND(x + xOffset) = anArray(x)
                Next
            Case 2
                DBCreateNewArrayOfType answerND, vType, x1 + xOffset, x2 + xOffset, y1 + yOffset, y2 + yOffset
                For x = x1 To x2
                    For y = y1 To y2
                        answerND(x + xOffset, y + yOffset) = anArray(x, y)
                    Next
                Next
            Case 3
                DBCreateNewArrayOfType answerND, vType, x1 + xOffset, x2 + xOffset, y1 + yOffset, y2 + yOffset, z1 + zOffset, z2 + zOffset
                For x = x1 To x2
                    For y = y1 To y2
                        For z = z1 To z2
                            answerND(x + xOffset, y + yOffset, z + zOffset) = anArray(x, y, z)
                        Next
                    Next
                Next
        End Select
    Else
        ' copy with type conversion
        Select Case nDimensions
        
            Case 1
                Select Case newType
                    Case vbByte
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Byte
                        For x = x1 To x2: answerND(x + xOffset) = CByte(anArray(x)): Next
                    Case vbInteger
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Integer
                        For x = x1 To x2: answerND(x + xOffset) = CInt(anArray(x)): Next
                    Case vbLong
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Long
                        For x = x1 To x2: answerND(x + xOffset) = CLng(anArray(x)): Next
                    Case vbSingle
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Single
                        For x = x1 To x2: answerND(x + xOffset) = CSng(anArray(x)): Next
                    Case vbDouble
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Double
                        For x = x1 To x2: answerND(x + xOffset) = CDbl(anArray(x)): Next
                    Case vbBoolean
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Boolean
                        For x = x1 To x2: answerND(x + xOffset) = CBool(anArray(x)): Next
                    Case vbDate
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Date
                        For x = x1 To x2: answerND(x + xOffset) = CDate(anArray(x)): Next
                    Case vbCurrency
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As Currency
                        For x = x1 To x2: answerND(x + xOffset) = CCur(anArray(x)): Next
                    Case vbString
                        ReDim answerND(x1 + xOffset To x2 + xOffset) As String
                        For x = x1 To x2: answerND(x + xOffset) = CStr(anArray(x)): Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary, METHOD_NAME, "#Unsupported Type!"
                End Select
                
            Case 2
                Select Case newType
                    Case vbByte
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Byte
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CByte(anArray(x, y)): Next: Next
                    Case vbInteger
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Integer
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CInt(anArray(x, y)): Next: Next
                    Case vbLong
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Long
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CLng(anArray(x, y)): Next: Next
                    Case vbSingle
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Single
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CSng(anArray(x, y)): Next: Next
                    Case vbDouble
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Double
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CDbl(anArray(x, y)): Next: Next
                    Case vbBoolean
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Boolean
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CBool(anArray(x, y)): Next: Next
                   Case vbDate
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Date
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CDate(anArray(x, y)): Next: Next
                   Case vbCurrency
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As Currency
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CCur(anArray(x, y)): Next: Next
                   Case vbString
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset) As String
                        For x = x1 To x2: For y = y1 To y2: answerND(x + xOffset, y + yOffset) = CStr(anArray(x, y)): Next: Next
                    Case Else
                    DBErrors_raiseGeneralError ModuleSummary, METHOD_NAME, "#Unsupported Type!"
                End Select
                
            Case 3
                Select Case newType
                    Case vbByte
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Byte
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CByte(anArray(x, y, z)): Next: Next: Next
                    Case vbInteger
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Integer
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CInt(anArray(x, y, z)): Next: Next: Next
                    Case vbLong
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Long
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CLng(anArray(x, y, z)): Next: Next: Next
                    Case vbSingle
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Single
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CSng(anArray(x, y, z)): Next: Next: Next
                    Case vbDouble
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Double
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CDbl(anArray(x, y, z)): Next: Next: Next
                    Case vbBoolean
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Boolean
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CBool(anArray(x, y, z)): Next: Next: Next
                   Case vbDate
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Date
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CDate(anArray(x, y, z)): Next: Next: Next
                   Case vbCurrency
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As Currency
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CCur(anArray(x, y, z)): Next: Next: Next
                   Case vbString
                        ReDim answerND(x1 + xOffset To x2 + xOffset, y1 + yOffset To y2 + yOffset, z1 + zOffset To z2 + zOffset) As String
                        For x = x1 To x2: For y = y1 To y2: For z = z1 To z2: answerND(x + xOffset, y + yOffset, z + zOffset) = CStr(anArray(x, y, z)): Next: Next: Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary, METHOD_NAME, "#Unsupported Type!"
                End Select
        End Select
    End If
    DBSubArray = answerND
End Function

Sub DBCreateNewArrayOfType(ByRef oArray As Variant, vType As Integer, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        
        Case y2 < y1
            Select Case vType
                Case vbByte
                    ReDim oArray(x1 To x2) As Byte
                Case vbInteger
                    ReDim oArray(x1 To x2) As Integer
                Case vbLong
                    ReDim oArray(x1 To x2) As Long
                Case vbSingle
                    ReDim oArray(x1 To x2) As Single
                Case vbDouble
                    ReDim oArray(x1 To x2) As Double
                Case vbBoolean
                    ReDim oArray(x1 To x2) As Boolean
                Case vbDate
                    ReDim oArray(x1 To x2) As Date
                Case vbCurrency
                    ReDim oArray(x1 To x2) As Currency
                Case vbString
                    ReDim oArray(x1 To x2) As String
                Case vbVariant
                    ReDim oArray(x1 To x2) As Variant
            End Select
        
        Case z2 < z1
            Select Case vType
                Case vbByte
                    ReDim oArray(x1 To x2, y1 To y2) As Byte
                Case vbInteger
                    ReDim oArray(x1 To x2, y1 To y2) As Integer
                Case vbLong
                    ReDim oArray(x1 To x2, y1 To y2) As Long
                Case vbSingle
                    ReDim oArray(x1 To x2, y1 To y2) As Single
                Case vbDouble
                    ReDim oArray(x1 To x2, y1 To y2) As Double
                Case vbBoolean
                    ReDim oArray(x1 To x2, y1 To y2) As Boolean
                Case vbDate
                    ReDim oArray(x1 To x2, y1 To y2) As Date
                Case vbCurrency
                    ReDim oArray(x1 To x2, y1 To y2) As Currency
                Case vbString
                    ReDim oArray(x1 To x2, y1 To y2) As String
                Case vbVariant
                    ReDim oArray(x1 To x2, y1 To y2) As Variant
            End Select
            
        Case Else
            Select Case vType
                Case vbByte
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Byte
                Case vbInteger
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Integer
                Case vbLong
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Long
                Case vbSingle
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Single
                Case vbDouble
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Double
                Case vbBoolean
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Boolean
                Case vbDate
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Date
                Case vbCurrency
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Currency
                Case vbString
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As String
                Case vbVariant
                    ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Variant
            End Select
    End Select
End Sub

Sub DBCreateNewArrayOfBytes(ByRef oArray() As Byte, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Byte
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Byte
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Byte
    End Select
End Sub

Sub DBCreateNewArrayOfIntegers(ByRef oArray() As Integer, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Integer
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Integer
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Integer
    End Select
End Sub

Sub DBCreateNewArrayOfLongs(ByRef oArray() As Long, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Long
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Long
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Long
    End Select
End Sub

Sub DBCreateNewArrayOfSingles(ByRef oArray() As Single, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Single
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Single
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Single
    End Select
End Sub

Sub DBCreateNewArrayOfDoubles(ByRef oArray() As Double, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Double
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Double
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Double
    End Select
End Sub

Sub DBCreateNewArrayOfBooleans(ByRef oArray() As Boolean, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Boolean
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Boolean
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Boolean
    End Select
End Sub

Sub DBCreateNewArrayOfDates(ByRef oArray() As Date, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Date
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Date
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Date
    End Select
End Sub

Sub DBCreateNewArrayOfCurrencys(ByRef oArray() As Currency, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Currency
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Currency
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Currency
    End Select
End Sub

Sub DBCreateNewArrayOfStrings(ByRef oArray() As String, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As String
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As String
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As String
    End Select
End Sub

Sub DBCreateNewArrayOfVariants(ByRef oArray() As Variant, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2) As Variant
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2) As Variant
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2) As Variant
    End Select
End Sub

Sub DBCreateNewVariantArray(ByRef oArray As Variant, x1 As Long, x2 As Long, Optional y1 As Long = 0, Optional y2 As Long = -1, Optional z1 As Long = 0, Optional z2 As Long = -1)
    Select Case True
        Case y2 < y1
            ReDim oArray(x1 To x2)
        Case z2 < z1
            ReDim oArray(x1 To x2, y1 To y2)
        Case Else
            ReDim oArray(x1 To x2, y1 To y2, z1 To z2)
    End Select
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' stop watch
'*************************************************************************************************************************************************************************************************************************************************
   
Function DBSW_timeNow() As Currency
    If apiQueryPerformanceCounter_Currency(DBSW_timeNow).BOOL <> BOOL_TRUE Then DBErrors_raiseGeneralError ModuleSummary(), "DBSW_timeNow", "kernel32.dll>>QueryPerformanceCounter failed"
End Function

Function DBSW_differenceInMilliseconds(startTime As Currency, stopTime As Currency) As Double
    DBSW_differenceInMilliseconds = (stopTime - startTime) / DBSW_frequency() * 1000#
End Function

Private Function DBSW_frequency() As Currency
    If apiQueryPerformanceFrequency_Currency(DBSW_frequency).BOOL <> BOOL_TRUE Then DBErrors_raiseGeneralError ModuleSummary(), "DBSW_frequency", "kernel32.dll>>QueryPerformanceFrequency failed"
End Function


'*************************************************************************************************************************************************************************************************************************************************
' logging - needs to be linked to vb trace or sysinternals / winapi (apiOutputDebugStringA) + have compile options
'*************************************************************************************************************************************************************************************************************************************************

Sub DBTraceError(ModuleSummary() As Variant, METHOD_NAME As String, errorState() As Variant)
    VBTrace DBTrace_prettySource(ModuleSummary, METHOD_NAME), prettyErrorState(errorState)
End Sub

Sub DBTrace(ModuleSummary() As Variant, METHOD_NAME As String, Description As String)
    VBTrace DBTrace_prettySource(ModuleSummary, METHOD_NAME), Description
End Sub

Function DBTrace_prettySource(ModuleSummary() As Variant, METHOD_NAME As String) As String
    Dim projectName As String, moduleName As String, moduleVersion As String
    Select Case ModuleSummary(0)
        Case 1
            projectName = ModuleSummary(1)
            moduleName = ModuleSummary(2)
            moduleVersion = ModuleSummary(3)
            DBTrace_prettySource = projectName & "." & moduleName & "(v" & moduleVersion & ")" & ">>" & METHOD_NAME
        Case Else
            DBTrace_prettySource = "Unknown module summary version"
    End Select
End Function


'*************************************************************************************************************************************************************************************************************************************************
' error state capture
'*************************************************************************************************************************************************************************************************************************************************

Function DBErrors_errorState() As Variant()
    DBErrors_errorState = Array(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext, Err.lastDLLError, Erl())
End Function

Function DBErrors_errorStateNumber(errorState() As Variant) As Long
    DBErrors_errorStateNumber = errorState(0)
End Function

Function DBErrors_errorStateSource(errorState() As Variant) As String
    DBErrors_errorStateSource = errorState(1)
End Function

Function DBErrors_errorStateDescription(errorState() As Variant) As String
    DBErrors_errorStateDescription = errorState(2)
End Function

Function DBErrors_errorStateHelpFile(errorState() As Variant) As String
    DBErrors_errorStateHelpFile = errorState(3)
End Function

Function DBErrors_errorStateHelpContext(errorState() As Variant) As String
    DBErrors_errorStateHelpContext = errorState(4)
End Function

Function DBErrors_errorStateLastDLLError(errorState() As Variant) As String
    DBErrors_errorStateLastDLLError = errorState(5)
End Function

Function DBErrors_errorStateLineNumber(errorState() As Variant) As String
    DBErrors_errorStateLineNumber = errorState(6)
End Function

Private Function prettyErrorState(errorState() As Variant) As String
    ' don't hex$ the error code as it can be done by the programmer later and is easier to identify in decimal (?)
    prettyErrorState = "<#" & _
        DBErrors_errorStateNumber(errorState) & "(" & _
        DBErrors_errorStateNumber(errorState) - vbObjectError & ")," & _
        DBErrors_errorStateSource(errorState) & "," & _
        DBErrors_errorStateDescription(errorState) & ",@" & _
        DBErrors_errorStateLineNumber(errorState) & ">"
End Function


'*************************************************************************************************************************************************************************************************************************************************
' error raising
'*************************************************************************************************************************************************************************************************************************************************

Sub DBErrors_raiseError(errorNumber As Long, ModuleSummary() As Variant, METHOD_NAME As String, Description As String)
    Err.Raise errorNumber, DBTrace_prettySource(ModuleSummary, METHOD_NAME), Description
End Sub

Sub DBErrors_raiseGeneralError(ModuleSummary() As Variant, METHOD_NAME As String, Description As String)
    DBErrors_raiseError GENERAL_ERROR_CODE, ModuleSummary, METHOD_NAME, Description
End Sub

Sub DBErrors_raiseUnhandledCase(ModuleSummary() As Variant, METHOD_NAME As String, caseID As String)
    DBErrors_raiseError UNHANDLED_CASE_ERROR_CODE, ModuleSummary, METHOD_NAME, "An unhandle case (" & IIf(Len(Trim(caseID)) = 0, "<not specified>", caseID) & ") has occured in " & DBTrace_prettySource(ModuleSummary, METHOD_NAME) & ". Please contact support."
End Sub

Sub DBErrors_raiseNotYetImplemented(ModuleSummary() As Variant, METHOD_NAME As String, feature As String)
    DBErrors_raiseError NOT_YET_IMPLIMENTED_ERROR_CODE, ModuleSummary, METHOD_NAME, feature & " has not been implemented yet in " & DBTrace_prettySource(ModuleSummary, METHOD_NAME) & ". Please contact support."
End Sub

Sub DBErrors_reraiseErrorState(errorState() As Variant)
    Err.Raise _
        DBErrors_errorStateNumber(errorState), _
        DBErrors_errorStateSource(errorState), _
        DBErrors_errorStateDescription(errorState), _
        DBErrors_errorStateHelpFile(errorState), _
        DBErrors_errorStateHelpContext(errorState)
End Sub

Sub DBErrors_reraiseErrorStateFrom(errorState() As Variant, ModuleSummary() As Variant, METHOD_NAME As String)
    DBErrors_raiseError DBErrors_errorStateNumber(errorState), ModuleSummary, METHOD_NAME, "Error " & prettyErrorState(errorState) & " occured"
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' module summary
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function





