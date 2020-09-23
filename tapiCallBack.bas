Attribute VB_Name = "tapiCallBack"
Option Explicit
'Call back functions use to check the state of a phone call for
'both monitoring a call and making a call

Global CallInfo As LINECALLINFO
Global CallParams As LINECALLPARAMS
Public Sub LINECALLBACKMON(ByVal hdevice As Long, ByVal dwMessage As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)

    'Handels messages from Tapi32 for Monitoring calls

    Dim hCall As Long
    If dwMessage = LINE_CALLSTATE Then
       hCall = PtrToLong(hdevice)
       Select Case dwParam1
          Case LINECALLSTATE_IDLE 'Call Terminated
            If hCall <> 0 Then
             lineDeallocateCall (hCall)
              AddToLst ("Idle - Monitored call deallocated")
            End If
          Case LINECALLSTATE_DIALING ' Call Dialing
              AddToLst ("Dialing")
          Case LINECALLSTATE_CONNECTED 'Service Connected
            If hCall <> 0 Then
              AddToLst ("Connected")
              
            End If
          Case LINECALLSTATE_PROCEEDING 'Call Proceeding (dialing)
              AddToLst ("Proceeding")
          Case LINECALLSTATE_DISCONNECTED 'Disconnected
            If dwParam2 = LINEDISCONNECTMODE_NORMAL Then
               AddToLst ("Disconnected Normal")
            ElseIf dwParam2 = LINEDISCONNECTMODE_BUSY Then
               AddToLst ("Disconnected Busy")
            End If
          Case LINECALLSTATE_BUSY 'Line Busy
              AddToLst ("Line Busy")
       End Select
   End If
End Sub
Public Sub LINECALLBACKCALL(ByVal hdevice As Long, ByVal dwMessage As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long, ByVal dwParam3 As Long)
    'Handels messages from Tapi32 for Making calls
    Dim hCall As Long
    If dwMessage = LINE_REPLY Then
      If dwParam2 < 0 Then
         AddToLst ("Line Replay Error")
      Else
         AddToLst ("Line Replay Ok")
      End If
   ElseIf dwMessage = LINE_CALLSTATE Then
     hCall = PtrToLong(hdevice)
     Select Case dwParam1
       Case LINECALLSTATE_IDLE
           If hCall <> 0 Then
             lineDeallocateCall (hCall)
             'AddToLst ("Idle - Call")
             frmTapiTest.cmdCall.Enabled = True
             frmTapiTest.cmdHangUp.Enabled = False
           End If
       Case LINECALLSTATE_CONNECTED
           If hCall <> 0 Then
            'AddToLst ("Connected")

           End If
       Case LINECALLSTATE_PROCEEDING
           'Addtolst ("Proceeding")
       Case LINECALLSTATE_DIALING
           'AddToLst ("Dialing")
       Case LINECALLSTATE_DISCONNECTED
           If dwParam2 = LINEDISCONNECTMODE_NORMAL Then
              'AddToLst ("Disconnected Normal")
           ElseIf dwParam2 = LINEDISCONNECTMODE_BUSY Then
              'AddToLst ("Disconnected Busy")
           End If
       Case LINECALLSTATE_BUSY
           'AddToLst ("Line Busy")
     End Select
   End If
End Sub

Public Function PtrToLong(ByVal lngFnPtr As Long) As Long
'Convert Pointer into Long
    PtrToLong = lngFnPtr
End Function

Public Sub AddToLst(strTemp As String)
  'Adds Text to the list box and displays the last entry in
  'the list box
   frmTapiTest.lstStatus.AddItem (strTemp)
   frmTapiTest.lstStatus.TopIndex = _
   frmTapiTest.lstStatus.ListCount - 1
End Sub
