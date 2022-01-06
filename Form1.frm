VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "WAV Volume"
   ClientHeight    =   2955
   ClientLeft      =   2775
   ClientTop       =   2430
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   2175
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   120
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      Left            =   960
      Max             =   1
      Min             =   7
      TabIndex        =   2
      Top             =   720
      Value           =   1
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   2
      Min             =   -2
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play WAV"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2115
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type lVolType
    v As Long
End Type

Private Type VolType
    lv As Integer
    rv As Integer
End Type

Private Declare Function waveOutGetVolume Lib "winmm.dll" _
      (ByVal uDeviceID As Long, lpdwVolume As Long) As Long

Private Declare Function waveOutSetVolume Lib "winmm.dll" _
      (ByVal uDeviceID As Long, ByVal dwVolume As Long) _
       As Long

Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Private Sub Command1_Click()

CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "Sound (*.wav)|*.wav"
CommonDialog1.FLAGS = &H80000 Or &H1000
CommonDialog1.ShowOpen
  i = mciSendString("close voice1", 0&, 0, 0)
i = mciSendString("open " & CommonDialog1.filename & " type waveaudio alias voice1", 0&, 0, 0)


Exit Sub


EH1:
Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
If Err = 59 Then
  MsgBox Err.Description & vbCrLf & vbCrLf & vbTab & "This file appears to be corrupted. Make sure that this file (with the *.as2 extension) is the proper file for this brand of software." & vbCrLf & vbTab & "It is possible that this file is for use with another software package that uses the same As2 (*.as2) file extension.", vbExclamation, "ERR #" & Err
  Exit Sub
End If
Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
Err.Clear

End Sub

Private Sub Command3_Click()



i = mciSendString("play voice1 from 0", 0&, 0, 0)


    VScroll1.SetFocus
End Sub

Private Sub Form_Load()
  HScroll1.Value = 0
  VScroll1.Value = 2
  Form1.Show
  VScroll1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  i = mciSendString("close voice1", 0&, 0, 0)
End Sub


Private Sub Timer1_Timer()
  Dim id As Long, v As Long, i As Long
    id = -1
    
    If VScroll1.Value = 1 And HScroll1.Value = -2 Then _
         i = waveOutSetVolume(id, 0)
    If VScroll1.Value = 1 And HScroll1.Value = -1 Then _
         i = waveOutSetVolume(id, 0)
    If VScroll1.Value = 1 And HScroll1.Value = 0 Then _
         i = waveOutSetVolume(id, 0)
    If VScroll1.Value = 1 And HScroll1.Value = 1 Then _
         i = waveOutSetVolume(id, 0)
    If VScroll1.Value = 1 And HScroll1.Value = 2 Then _
         i = waveOutSetVolume(id, 0)
    
    If VScroll1.Value = 2 And HScroll1.Value = -2 Then _
         i = waveOutSetVolume(id, 10280)
    If VScroll1.Value = 2 And HScroll1.Value = -1 Then _
         i = waveOutSetVolume(id, 379004968)
    If VScroll1.Value = 2 And HScroll1.Value = 0 Then _
         i = waveOutSetVolume(id, 673720360)
    If VScroll1.Value = 2 And HScroll1.Value = 1 Then _
         i = waveOutSetVolume(id, 673714578)
    If VScroll1.Value = 2 And HScroll1.Value = 2 Then _
         i = waveOutSetVolume(id, 673710080)
    
    If VScroll1.Value = 3 And HScroll1.Value = -2 Then _
         i = waveOutSetVolume(id, 20560)
    If VScroll1.Value = 3 And HScroll1.Value = -1 Then _
         i = waveOutSetVolume(id, 757944400)
    If VScroll1.Value = 3 And HScroll1.Value = 0 Then _
         i = waveOutSetVolume(id, 1347440720)
    If VScroll1.Value = 3 And HScroll1.Value = 1 Then _
         i = waveOutSetVolume(id, 1347429155)
    If VScroll1.Value = 3 And HScroll1.Value = 2 Then _
         i = waveOutSetVolume(id, 1347420160)
    
    If VScroll1.Value = 4 And HScroll1.Value = -2 Then _
        i = waveOutSetVolume(id, 31868)
    If VScroll1.Value = 4 And HScroll1.Value = -1 Then _
        i = waveOutSetVolume(id, 1174830204)
    If VScroll1.Value = 4 And HScroll1.Value = 0 Then _
        i = waveOutSetVolume(id, 2088533116)
    If VScroll1.Value = 4 And HScroll1.Value = 1 Then _
        i = waveOutSetVolume(id, 2088515191)
    If VScroll1.Value = 4 And HScroll1.Value = 2 Then _
        i = waveOutSetVolume(id, 2088501248)
    
    If VScroll1.Value = 5 And HScroll1.Value = -2 Then _
        i = waveOutSetVolume(id, 42919)
    If VScroll1.Value = 5 And HScroll1.Value = -1 Then _
        i = waveOutSetVolume(id, 1582213031)
    If VScroll1.Value = 5 And HScroll1.Value = 0 Then _
        i = waveOutSetVolume(id, -1482184793)
    If VScroll1.Value = 5 And HScroll1.Value = 1 Then _
        i = waveOutSetVolume(id, -1482208934)
    If VScroll1.Value = 5 And HScroll1.Value = 2 Then _
        i = waveOutSetVolume(id, -1482227712)
    
    If VScroll1.Value = 6 And HScroll1.Value = -2 Then _
        i = waveOutSetVolume(id, 54227)
    If VScroll1.Value = 6 And HScroll1.Value = -1 Then _
        i = waveOutSetVolume(id, 1554895827)
    If VScroll1.Value = 6 And HScroll1.Value = 0 Then _
        i = waveOutSetVolume(id, -741092397)
    If VScroll1.Value = 6 And HScroll1.Value = 1 Then _
        i = waveOutSetVolume(id, -741122899)
    If VScroll1.Value = 6 And HScroll1.Value = 2 Then _
        i = waveOutSetVolume(id, -741146624)
    
    If VScroll1.Value = 7 And HScroll1.Value = -2 Then _
        i = waveOutSetVolume(id, 65535)
    If VScroll1.Value = 7 And HScroll1.Value = -1 Then _
        i = waveOutSetVolume(id, -1878982657)
    If VScroll1.Value = 7 And HScroll1.Value = 0 Then _
        i = waveOutSetVolume(id, -1)
    If VScroll1.Value = 7 And HScroll1.Value = 1 Then _
        i = waveOutSetVolume(id, -36865)
    If VScroll1.Value = 7 And HScroll1.Value = 2 Then _
        i = waveOutSetVolume(id, -65536)
End Sub

