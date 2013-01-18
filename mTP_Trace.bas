Attribute VB_Name = "mTP_Trace"
'************************************************************************************************************************************************
'
'    Copyright (c) 2012 David Briant - see https://github.com/DangerMouseB
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

' this module sends traces to the VBTracer available from http://vbaccelerator.com by Steve McMahon

Option Explicit

Private myHWnd As Long
Private myLastSearchTime As Single

#If INCLUDE_TRACES <> 0 Then
Sub VBTrace(ParamArray args() As Variant)
    Dim i As Long, j As Long, textToSend As String, cds As COPYDATASTRUCT, buffer() As Byte, result As Long, retval As Long

    ' if we don't have a handle on the window and it's been at least 1000mS since we last searched then search for the tracer window
    If myHWnd = 0 And (Timer() - myLastSearchTime) * 1000 > 1000 Then
        myLastSearchTime = Timer()
        apiEnumWindows AddressOf EnumWindowsProc, 0
    End If
    If myHWnd = 0 Then Exit Sub
   
    ' contruct textToSend
    On Error Resume Next
    For i = LBound(args) To UBound(args)
        If ((varType(args(i)) And vbArray) = vbArray) Then
            For j = LBound(args(i)) To UBound(args(i))
                textToSend = textToSend & args(i)(j) & vbTab
            Next
        Else
            textToSend = textToSend & args(i) & vbTab
        End If
    Next
    textToSend = App.EXEName & ": " & App.hInstance & ": " & App.ThreadID & ": " & Format$(Now(), "yyyymmdd hhnnss") & ": " & textToSend
    
    ' send it, noting that we need to re-search for tracer on a timeout
    buffer = StrConv(textToSend, vbFromUnicode)
    cds.dwData = 0
    cds.cbData = UBound(buffer) + 1
    cds.lpData = VarPtr(buffer(0))
    retval = apiSendMessageTimeoutA(myHWnd, WM_COPYDATA, 0, cds, SMTO_NORMAL, 200, result)
    If retval <> 1 Then myHWnd = 0

End Sub

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
   If apiGetPropA(hWnd, "vbAcceleratorVBTRACER_TRACEWIN") = 1 Then
      EnumWindowsProc = 0
      myHWnd = hWnd
   Else
      EnumWindowsProc = 1
   End If
End Function

#Else
Sub VBTrace(ParamArray args() As Variant)
End Sub
#End If
