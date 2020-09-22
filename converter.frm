VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Converter"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "converter.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdConvert 
      Left            =   5490
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save File As...."
      InitDir         =   "C:\"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Convert...."
      Height          =   390
      Left            =   5820
      TabIndex        =   12
      Top             =   2205
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Encoding Options"
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   2550
      TabIndex        =   10
      Top             =   540
      Width           =   5010
      Begin VB.CheckBox Check2 
         Caption         =   "Add Reverb"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1410
         Width           =   1290
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Equalizer"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   1185
         Width           =   1635
      End
      Begin MSComctlLib.Slider bitratesld 
         Height          =   630
         Left            =   75
         TabIndex        =   11
         Top             =   270
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   1111
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   10
         TickStyle       =   2
         Value           =   10
         TextPosition    =   1
      End
      Begin VB.Label curbit 
         Caption         =   "320kbps"
         Height          =   225
         Left            =   4170
         TabIndex        =   18
         Top             =   495
         Width           =   750
      End
      Begin VB.Label hilbl 
         Caption         =   "320"
         Height          =   210
         Left            =   3795
         TabIndex        =   17
         Top             =   915
         Width           =   270
      End
      Begin VB.Label lolbl 
         Caption         =   "32"
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   915
         Width           =   270
      End
   End
   Begin VB.OptionButton convtype 
      Caption         =   "OGG Vorbis VBR"
      Height          =   240
      Index           =   4
      Left            =   15
      TabIndex        =   9
      Top             =   1650
      Width           =   2280
   End
   Begin VB.OptionButton convtype 
      Caption         =   "OGG Vorbis CBR"
      Height          =   240
      Index           =   3
      Left            =   15
      TabIndex        =   8
      Top             =   1380
      Width           =   2280
   End
   Begin VB.OptionButton convtype 
      Caption         =   "Windows Media Audio VBR"
      Height          =   240
      Index           =   5
      Left            =   15
      TabIndex        =   7
      Top             =   1920
      Width           =   2670
   End
   Begin VB.OptionButton convtype 
      Caption         =   "MP3 VBR"
      Height          =   240
      Index           =   2
      Left            =   15
      TabIndex        =   6
      Top             =   1095
      Width           =   2250
   End
   Begin VB.OptionButton convtype 
      Caption         =   "MP3 CBR"
      Height          =   240
      Index           =   1
      Left            =   15
      TabIndex        =   5
      Top             =   825
      Width           =   2265
   End
   Begin VB.OptionButton convtype 
      Caption         =   "WAV Audio (1411kbps)"
      Height          =   240
      Index           =   0
      Left            =   15
      TabIndex        =   4
      Top             =   555
      Value           =   -1  'True
      Width           =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abort"
      Default         =   -1  'True
      Height          =   345
      Left            =   6615
      TabIndex        =   1
      Top             =   3255
      Width           =   930
   End
   Begin MSComctlLib.ProgressBar encProg 
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   2955
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin VB.Label Inputlbl 
      Caption         =   "Input"
      Height          =   195
      Left            =   15
      TabIndex        =   15
      Top             =   0
      Width           =   7455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   75
      X2              =   7440
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0/0 [Position/Length]"
      Height          =   225
      Left            =   45
      TabIndex        =   3
      Top             =   3300
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Converter: Idle"
      Height          =   270
      Left            =   45
      TabIndex        =   2
      Top             =   2700
      Width           =   7515
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public inputx As String
Public inputtype As Boolean
Public outputtype As Integer
 
Private Sub bitratesld_Change()
curbit = bitratesld.value * 32 & "kbps"
End Sub

Private Sub bitratesld_Scroll()
curbit = bitratesld.value * 32 & "kbps"
End Sub

Private Sub Command1_Click()
Call BASS_Encode_Stop(chan)
Call BASS_StreamFree(chan)
BASS_Free
Form1.Enabled = True
Me.Hide
End Sub

Private Sub Command2_Click()
cmdConvert.CancelError = 1
Debug.Print outputtype
Select Case outputtype
Case 0
cmdConvert.Filter = "Windows PCM Wave File (*.WAV)|*.wav"
On Error GoTo nosale
cmdConvert.ShowSave
Encoder.WavWrite inputtype, inputx, cmdConvert.FileName
Case 1
cmdConvert.Filter = "MPEG Layer III (*.MP3)|*.mp3"
On Error GoTo nosale
cmdConvert.ShowSave
Encoder.Converter inputx, "MP3C", Str$(bitratesld.value * 32), cmdConvert.FileName, inputtype
Case 2 '
cmdConvert.Filter = "MPEG Layer III (*.MP3)|*.mp3" '
On Error GoTo nosale
cmdConvert.ShowSave
Encoder.Converter inputx, "MP3V", Str$(bitratesld.value * 32), cmdConvert.FileName, inputtype
Case 3
cmdConvert.Filter = "MPEG Layer II (*.MP2)|*.mp2"
On Error GoTo nosale
cmdConvert.ShowSave
Encoder.Converter inputx, "OGGC", Str$(bitratesld.value * 32), cmdConvert.FileName, inputtype
Case 4
cmdConvert.Filter = "OGG Vorbis Audio (*.OGG)|*.ogg"
On Error GoTo nosale
cmdConvert.ShowSave
Encoder.Converter inputx, "OGGV", Str$(bitratesld.value * 32), cmdConvert.FileName, inputtype
Case 5
cmdConvert.Filter = "Windows Media Audio (*.WMA)|*.wma"
On Error GoTo nosale
cmdConvert.ShowSave
Encoder.CopyToWMA inputx, bitratesld.value * 32000, cmdConvert.FileName, inputtype
End Select
Exit Sub
nosale:
End Sub


Private Sub convtype_Click(Index As Integer)
Select Case Index
Case 0, 1, 5
bitratesld.min = 1
bitratesld.max = 10
lolbl = "32"
hilbl = "320"
curbit = bitratesld.value * 32 & "kbps"
Case 2
bitratesld.min = 1
bitratesld.max = 9
lolbl = "32"
hilbl = "288"
curbit = bitratesld.value * 32 & "kbps"
Case 3, 4
bitratesld.min = 2
bitratesld.max = 10
lolbl = "64"
hilbl = "320"
curbit = bitratesld.value * 32 & "kbps"
End Select
outputtype = Index
End Sub

Private Sub Form_Activate()
If FileExists(Form1.ap & "LAME.exe") = False Then
convtype(1).Enabled = 0
convtype(2).Enabled = 0
Else
convtype(1).Enabled = 1
convtype(2).Enabled = 2
End If
If FileExists(Form1.ap & "OGGENC.exe") = False Then
convtype(3).Enabled = 0
convtype(4).Enabled = 0
Else
convtype(3).Enabled = 1
convtype(4).Enabled = 2
End If
Inputlbl = inputx
End Sub

Private Sub Form_Load()
Inputlbl = inputx
End Sub

