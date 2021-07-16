Attribute VB_Name = "mTP_Trace"
'************************************************************************************************************************************************
'
'    Copyright (c) 2012, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
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
