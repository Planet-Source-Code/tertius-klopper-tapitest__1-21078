VERSION 5.00
Begin VB.Form frmTapiTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TAPI Test                                                                       "
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Mediamode"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
      Begin VB.OptionButton optMediaMode 
         Caption         =   "Modem"
         Height          =   375
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optMediaMode 
         Caption         =   "Voice"
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCall 
      Caption         =   "Call"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   360
      Left            =   1815
      TabIndex        =   1
      Top             =   3405
      Width           =   2535
   End
   Begin VB.CommandButton cmdHangUp 
      Caption         =   "Hangup"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdMakeCall 
      Caption         =   "Make a Call"
      Height          =   375
      Left            =   2468
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      ItemData        =   "frmTapiTest.frx":0000
      Left            =   188
      List            =   "frmTapiTest.frx":0002
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Number to Call :"
      Height          =   255
      Left            =   375
      TabIndex        =   9
      Top             =   3458
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tapi Test will monitor outgoing calls from your modem and will make a call of Voice or Data"
      Height          =   495
      Left            =   1155
      TabIndex        =   8
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmTapiTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This app may cause VB to give a exeption error, when running it
'in compiled mode. First save all your work before
'running the app

'TapiTest is an updated version from TAPIMon this app also include
'functions to make a call to either voice or a datamode

'By Tertius
'E-Mail tertiusklopper@hotmail.com for more information ,comments
'or sugestions

'If you downloaded this Source Code From Planet-Source-Code.com
'please vote for me(Thank You)

'Parts of program was translated from Delphi 4 as my first app
'was written in that.

'A Lot of the Declared functions and types in vbTapi.bas is not
'used in this app but it will help you to use them without
'declaring it yourself just include the vbtapi.bas in you app

'vbTapi.bas was found on Planet-Source-code.com but do not
'remember who made it (To the persone who did it THANK YOU)

Dim udtLineCall As LINECALLPARAMS
Dim lines As Long
Dim hInst As Long
Dim lineApp As Long
Dim lphLine As Long
Dim lphCall As Long
Dim adrCallBack As Long

Private Sub cmdCall_Click()
Dim PhoneNum As String
PhoneNum = txtPhoneNum.Text
'Makes call with phone number provide by you
If Len(PhoneNum) > 0 Then
  If lineMakeCall(lphLine, lphCall, PhoneNum, 0, CallParams) < 0 Then
   AddToLst ("Error in Making call")
  Else
   'Call has been connected
   cmdCall.Enabled = False
   cmdHangUp.Enabled = True
  End If
End If

End Sub

Private Sub cmdHangUp_Click()
'Hangs up the call made
 If lineDrop(lphCall, 0, 0) < 0 Then
   AddToLst ("Error in Hanging up call")
   cmdCall.Enabled = True
   cmdHangUp.Enabled = False
 Else
   AddToLst ("Call Hanged up")
 End If
End Sub

Private Sub cmdMakeCall_Click()
'Increas the window size to show the call buttons
'if pressed again it will reduce the size of window again
'and hide the call buttons
If frmTapiTest.Height <> 5265 Then
frmTapiTest.Height = 5265
txtPhoneNum.SetFocus
Else
 frmTapiTest.Height = 3645
End If
 

End Sub

Private Sub Form_Load()
Dim nDevs As Long
Dim tapiVer As Long
Dim extid As LINEEXTENSIONID

frmTapiTest.Height = 3645 'Height of window at startup

'Section for Monitoring calls

If lineInitialize(lineApp, hInst, AddressOf LINECALLBACKMON, 0, nDevs) < 0 Then
    lineApp = 0
ElseIf nDevs = 0 Then  'No Tapi Device
    lineShutdown (lineApp)
    lineApp = 0
ElseIf lineNegotiateAPIVersion(lineApp, 0, 65536, 65540, tapiVer, extid) < 0 Then 'Check for version
    lineShutdown (lineApp)
    lineApp = 0
    lphLine = 0
    'Open a line for monitor (here I use the first device, normally the modem
ElseIf lineOpen(lineApp, 0, lphLine, tapiVer, 0, 0, LINECALLPRIVILEGE_MONITOR, LINEMEDIAMODE_DATAMODEM, 0) < 0 Then
    lineShutdown (lineApp)
    lineApp = 0
    lphLine = 0
End If

If lineApp <> 0 Then
    lstStatus.AddItem ("Monitoring Calls...")
    lstStatus.TopIndex = lstStatus.ListCount - 1
Else
    lstStatus.AddItem ("Error!")
    lstStatus.TopIndex = lstStatus.ListCount - 1
End If

'Section for Making Calls
 With CallParams
            .dwTotalSize = Len(CallParams)
            .dwBearerMode = LINEBEARERMODE_VOICE
            .dwMediaMode = LINEMEDIAMODE_INTERACTIVEVOICE
            'dwMediaMode = LINEMEDIAMODE_DATAMODEM
            'can also place any other call mediamode here
 End With
 
 If lineInitialize(lineApp, hInst, AddressOf LINECALLBACKCALL, 0, nDevs) <> 0 Then
     lineApp = 0
 ElseIf nDevs = 0 Then
     lineShutdown (lineApp)
     lineApp = 0
 ElseIf lineNegotiateAPIVersion(lineApp, 0, 65536, 65540, tapiVer, extid) < 0 Then 'Check for version
    lineShutdown (lineApp)
    lineApp = 0
    lphLine = 0
 ElseIf lineOpen(lineApp, 0, lphLine, tapiVer, 0, 0, LINECALLPRIVILEGE_NONE, 0, CallParams) < 0 Then
    lineShutdown (lineApp)
    lineApp = 0
    lphLine = 0

 End If
 If lphLine = 0 Then
    AddToLst ("Call error")
 End If
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
If lphLine <> 0 Then
   lineClose (lphLine)
End If
If lineApp <> 0 Then
   lineShutdown (lineApp)
End If
End Sub

Private Sub optMediaMode_Click(Index As Integer)
If optMediaMode(0).Value = True Then
 CallParams.dwMediaMode = LINEMEDIAMODE_INTERACTIVEVOICE
Else
 CallParams.dwMediaMode = LINEMEDIAMODE_DATAMODEM
End If
End Sub
