VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "MP3PlayerX2 v5.10"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "X2Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "X2Main.frx":48EA
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox iconbox 
      Height          =   270
      Left            =   2430
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   38
      Top             =   6810
      Width           =   270
   End
   Begin MSComctlLib.ProgressBar Position 
      Height          =   105
      Left            =   75
      TabIndex        =   34
      Top             =   3645
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog AddDialog 
      Left            =   5355
      Top             =   7905
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar Volume 
      Height          =   195
      Left            =   3690
      Max             =   100
      TabIndex        =   32
      Top             =   3870
      Width           =   1500
   End
   Begin VB.PictureBox VisBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   6945
      Picture         =   "X2Main.frx":7CC4
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   31
      Top             =   4995
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   9
      Left            =   6045
      Max             =   299
      TabIndex        =   30
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   8
      Left            =   5790
      Max             =   299
      TabIndex        =   29
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   7
      Left            =   5535
      Max             =   299
      TabIndex        =   28
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   6
      Left            =   5280
      Max             =   299
      TabIndex        =   27
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   5
      Left            =   5025
      Max             =   299
      TabIndex        =   26
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   4
      Left            =   4770
      Max             =   299
      TabIndex        =   25
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   3
      Left            =   4515
      Max             =   299
      TabIndex        =   24
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   2
      Left            =   4260
      Max             =   299
      TabIndex        =   23
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   1
      Left            =   4005
      Max             =   299
      TabIndex        =   22
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.VScrollBar Eq 
      Height          =   1590
      Index           =   0
      Left            =   3750
      Max             =   299
      TabIndex        =   21
      Top             =   4290
      Value           =   149
      Width           =   135
   End
   Begin VB.CheckBox EQOn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3735
      TabIndex        =   20
      ToolTipText     =   "10 Band Equalizer (DirectX8)"
      Top             =   4110
      Width           =   195
   End
   Begin VB.PictureBox Visbox 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   90
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   4140
      Width           =   3600
   End
   Begin VB.CheckBox rvbon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   15
      ToolTipText     =   "Reverb, for that nice 'live' sound"
      Top             =   6000
      Width           =   195
   End
   Begin VB.HScrollBar rvb 
      Height          =   165
      Index           =   0
      Left            =   1320
      Max             =   3000
      Min             =   1
      TabIndex        =   14
      Top             =   6180
      Value           =   2000
      Width           =   2340
   End
   Begin VB.HScrollBar rvb 
      Height          =   165
      Index           =   1
      Left            =   1320
      Max             =   250
      Min             =   1
      TabIndex        =   13
      Top             =   6435
      Value           =   200
      Width           =   2340
   End
   Begin VB.PictureBox Clear 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   846
      Picture         =   "X2Main.frx":8486
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      ToolTipText     =   "Clear the playlist"
      Top             =   3780
      Width           =   240
   End
   Begin VB.PictureBox Randomizer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   594
      Picture         =   "X2Main.frx":8810
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      ToolTipText     =   "Randomize Playlist"
      Top             =   3780
      Width           =   240
   End
   Begin VB.PictureBox MVDWN 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1602
      Picture         =   "X2Main.frx":8B9A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      ToolTipText     =   "Move the Song Down the playlist"
      Top             =   3780
      Width           =   240
   End
   Begin VB.PictureBox MVUP 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1860
      Picture         =   "X2Main.frx":8F24
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      ToolTipText     =   "Move the Song Up the Playlist"
      Top             =   3780
      Width           =   240
   End
   Begin VB.PictureBox RemoveSong 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1350
      Picture         =   "X2Main.frx":92AE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      ToolTipText     =   "Remove the selected song"
      Top             =   3780
      Width           =   240
   End
   Begin VB.PictureBox Addsong 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1098
      Picture         =   "X2Main.frx":9638
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      ToolTipText     =   "Add Audio Files/Playlists"
      Top             =   3780
      Width           =   240
   End
   Begin VB.PictureBox SAvelist 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   342
      Picture         =   "X2Main.frx":99C2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      ToolTipText     =   "Save Playlist (MPX,M3U or PLS)"
      Top             =   3780
      Width           =   240
   End
   Begin MSComDlg.CommonDialog ListDialog 
      Left            =   3180
      Top             =   7755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "MP3Player X2 List (*.mpx)|*.mpx|Winamp 2.91 or Lower Playlist (*.m3u)|*.m3u|Sonique Playlist (*.pls)|*.pls"
   End
   Begin VB.PictureBox OpenList 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   90
      Picture         =   "X2Main.frx":9D4C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      ToolTipText     =   "Open Playlist (MPX,M3U or PLS)"
      Top             =   3780
      Width           =   240
   End
   Begin MSComctlLib.ListView filename1 
      Height          =   960
      Left            =   3840
      TabIndex        =   3
      Top             =   7230
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   1693
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView Playlist1 
      CausesValidation=   0   'False
      Height          =   2970
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   5239
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   7500
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "len"
         Object.Width           =   1500
      EndProperty
   End
   Begin VB.Timer filetimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6405
      Top             =   7725
   End
   Begin VB.Timer vistimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6390
      Top             =   8115
   End
   Begin VB.PictureBox min 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6090
      ScaleHeight     =   240
      ScaleWidth      =   435
      TabIndex        =   1
      ToolTipText     =   "Minimise to Taskbar"
      Top             =   0
      Width           =   435
   End
   Begin VB.PictureBox closee 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6495
      ScaleHeight     =   240
      ScaleWidth      =   255
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   255
   End
   Begin MSComctlLib.ListView cdplaylist 
      Height          =   3015
      Left            =   90
      TabIndex        =   37
      Top             =   645
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   7500
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "len"
         Object.Width           =   1500
      EndProperty
   End
   Begin VB.Image resetimg 
      Height          =   180
      Left            =   6240
      Picture         =   "X2Main.frx":A0D6
      ToolTipText     =   "Reset The EQ"
      Top             =   5670
      Width           =   180
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   3405
      TabIndex        =   40
      Top             =   465
      Width           =   2490
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   1995
      TabIndex        =   39
      Top             =   465
      Width           =   1380
   End
   Begin VB.Image mp3click 
      Height          =   300
      Left            =   5925
      Picture         =   "X2Main.frx":A2C8
      ToolTipText     =   "Show the MP3 Playlist"
      Top             =   495
      Width           =   750
   End
   Begin VB.Image cdclick 
      Height          =   300
      Left            =   5925
      Picture         =   "X2Main.frx":A835
      ToolTipText     =   "Show the CD Player"
      Top             =   810
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reverb"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   300
      TabIndex        =   36
      Top             =   5985
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Equalizer"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   3945
      TabIndex        =   35
      Top             =   4080
      Width           =   885
   End
   Begin VB.Label Vollbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3105
      TabIndex        =   33
      Top             =   3855
      Width           =   615
   End
   Begin VB.Image Play1 
      Height          =   300
      Left            =   9810
      Picture         =   "X2Main.frx":AD6F
      Top             =   6540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Play2 
      Height          =   300
      Left            =   10110
      Picture         =   "X2Main.frx":B261
      Top             =   6540
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Pause1 
      Height          =   300
      Left            =   9810
      Picture         =   "X2Main.frx":B753
      Top             =   6870
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Pause2 
      Height          =   300
      Left            =   10125
      Picture         =   "X2Main.frx":BC45
      Top             =   6870
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Stop1 
      Height          =   300
      Left            =   9795
      Picture         =   "X2Main.frx":C137
      Top             =   7215
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Stop2 
      Height          =   300
      Left            =   10125
      Picture         =   "X2Main.frx":C629
      Top             =   7200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Prev1 
      Height          =   300
      Left            =   9825
      Picture         =   "X2Main.frx":CB1B
      Top             =   7530
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Prev2 
      Height          =   300
      Left            =   10140
      Picture         =   "X2Main.frx":D00D
      Top             =   7515
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Next1 
      Height          =   300
      Left            =   9825
      Picture         =   "X2Main.frx":D4FF
      Top             =   7830
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Next2 
      Height          =   300
      Left            =   10125
      Picture         =   "X2Main.frx":D9F1
      Top             =   7830
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image PlayB 
      Height          =   300
      Left            =   5490
      Picture         =   "X2Main.frx":DEE3
      ToolTipText     =   "Play"
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image PauResB 
      Height          =   300
      Left            =   5790
      Picture         =   "X2Main.frx":E3D5
      ToolTipText     =   "Pause"
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image StopB 
      Height          =   300
      Left            =   6090
      Picture         =   "X2Main.frx":E8C7
      ToolTipText     =   "Stop"
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image NextTrkB 
      Height          =   300
      Left            =   6390
      Picture         =   "X2Main.frx":EDB9
      ToolTipText     =   "Next"
      Top             =   3765
      Width           =   300
   End
   Begin VB.Image PrevTrkB 
      Height          =   300
      Left            =   5190
      Picture         =   "X2Main.frx":F2AB
      ToolTipText     =   "Previous"
      Top             =   3765
      Width           =   300
   End
   Begin VB.Label Tracktime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   5640
      TabIndex        =   18
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reverb Strength"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   1
      Left            =   105
      TabIndex        =   17
      Top             =   6405
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Size"
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Index           =   0
      Left            =   105
      TabIndex        =   16
      Top             =   6195
      Width           =   795
   End
   Begin VB.Label Titlelbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to MP3Player X2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   1860
      TabIndex        =   5
      Top             =   255
      UseMnemonic     =   0   'False
      Width           =   3045
   End
   Begin VB.Image close1 
      Height          =   240
      Left            =   9735
      Picture         =   "X2Main.frx":F79D
      Top             =   5040
      Width           =   255
   End
   Begin VB.Image close2 
      Height          =   240
      Left            =   10035
      Picture         =   "X2Main.frx":FB1F
      Top             =   5040
      Width           =   255
   End
   Begin VB.Image min2 
      Height          =   240
      Left            =   9825
      Picture         =   "X2Main.frx":FEA1
      Top             =   4785
      Width           =   435
   End
   Begin VB.Image min1 
      Height          =   240
      Left            =   9825
      Picture         =   "X2Main.frx":10463
      Top             =   4530
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fxr As Long
Public ap As String 'app path
Public tracknumberX As Long
Public xs As New clsTransForm
Public intNumberOfEntries As Long
Public MediaType As String
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim data(1 To 255) As Single
Dim Data2(1 To 255) As Single
Public s As String
Dim OldCDID As String
Public usebad As Long
Public nodsp As Long
Public novis As Long
Public cdtracknumberx As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public WithEvents ct As CTray
Attribute ct.VB_VarHelpID = -1
Private bh As BITMAPINFO     'bitmap header
Private specbuf() As Byte    'a pointer
Public DXVerX As Long
Const BI_RGB = 0&
Const DIB_RGB_COLORS = 0& '  color table in RGBs
Public vistype As Long
Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(256) As RGBQUAD
End Type
Private fft(1024) As Single



Private Sub Addsong_Click()
    With AddDialog

    ' Set Flags
    .CancelError = 1
    .flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNLongNames
    .Filter = "Audio Files (MP1,MP2,MP3,WAV,OGG,WMA)|*.mp1;*.mp2;*.mp3;*.wav;*.ogg;*.wma|MP3 Playlist (*.mpx, *.m3u, *.pls)|*.mpx;*.m3u;*.pls|"
    ' Max Size
    .MaxFileSize = 32767

    ' Reset FileName
    .FileName = ""

    ' Show the Open Dialog
    On Error GoTo NoFileFound
    .ShowOpen

    ' Check to see if the user selected a file, if not - exit
    If .FileName = "" Then Exit Sub

    ' Counter var
    Dim i As Long

    ' Go through all files selected
    For i = 1 To CountFilesInList(.FileName)

        ' Check the file size
        Select Case FileLen(GetFileFromList(.FileName, i))
            Case Is > 0
            Select Case .FilterIndex
            Case 1
            GoTo addmp3
            Case 2
            
            ReadList GetFileFromList(.FileName, i), 1
            GoTo nextf:
            End Select
addmp3:
                ' Now add the file to the list boxes
                               filename1.ListItems.Add , , GetFileFromList(.FileName, i)
                               filename1.ListItems(filename1.ListItems.Count).SubItems(1) = 0
                                Playlist1.ListItems.Add , , GetFileFromList(.FileName, i)
            Case Else
        End Select
nextf:
    Next
    End With
    Exit Sub
NoFileFound:


End Sub

Private Sub cdclick_Click()
Playlist1.Visible = 0
cdplaylist.Visible = 1
CheckCD
End Sub

Private Sub cdplaylist_DblClick()
If Not cdplaylist.ListItems.Count = 0 Then If Not cdplaylist.SelectedItem.tag = "DATA" Then PlayCDTrack (cdplaylist.SelectedItem.tag)
End Sub

Private Sub cdplaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not cdplaylist.ListItems.Count = 0 Then If Button = 2 Then PopupMenu Form2.mnucd

End Sub

'MP3Player X2 Version 5.00.0001

Private Sub Clear_Click()
Playlist1.ListItems.Clear
filename1.ListItems.Clear
End Sub

Public Sub closee_Click()
If FileExists("C:\TEMP.wav") = True Then
On Error Resume Next 'file may be in use
Kill "C:\TEMP.WAV"
'get rid of the temp wave file used for conversion
End If
BASS_ChannelRemoveFX chan, fxr
For u = 0 To 9
BASS_ChannelRemoveFX chan, fxeq(u)
Next u
BASS_StreamFree chan
BASS_Free
SaveSettings
ct.DeleteIcon
Set xs = Nothing
Set ct = Nothing
End

End Sub

Private Sub closee_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
closee.Picture = close2.Picture
End Sub

Private Sub closee_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
closee.Picture = close1.Picture
End Sub

Private Sub ct_LButtonDblClick()
filetimer.Interval = 1
If Not MediaType = "CONV" Then vistimer.Enabled = 1
Me.Show
End Sub

Private Sub ct_RButtonUp()
PopupMenu Form2.mnutray
End Sub

Sub Dev_Write_Preset()
On Error GoTo nocado:
Open ap & Hour(time) & "_" & Minute(time) & "_" & Second(time) & " " & Day(Date) & "_" & Month(Date) & "_" & "output_debugger.log" For Output As #1
For u = 0 To 9
Print #1, Eq(u).value
Next
Print #1, rvb(0).value
Print #1, rvb(1).value
Close
nocado:
End Sub

Private Sub Eq_Change(Index As Integer)
'change eq
Dim p As BASS_FXPARAMEQ
    Call BASS_FXGetParameters(fxeq(Index), p)
    p.fGain = 15 - (Eq(Index).value / 10)
    Call BASS_FXSetParameters(fxeq(Index), p)
    End Sub

Private Sub Eq_Scroll(Index As Integer)
Call Eq_Change(Index)
End Sub

Public Sub EQOn_Click()
Select Case EQOn.value
Case 1
fxeq(0) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(1) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(2) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(3) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(4) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(5) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(6) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(7) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(8) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
fxeq(9) = BASS_ChannelSetFX(chan, BASS_FX_PARAMEQ)
Dim p As BASS_FXPARAMEQ
p.fBandwidth = 1
p.fCenter = 80
Call BASS_FXSetParameters(fxeq(0), p)
p.fCenter = 120
Call BASS_FXSetParameters(fxeq(1), p)
p.fCenter = 230
Call BASS_FXSetParameters(fxeq(2), p)
p.fCenter = 460
Call BASS_FXSetParameters(fxeq(3), p)
p.fCenter = 690
Call BASS_FXSetParameters(fxeq(4), p)
p.fCenter = 919
Call BASS_FXSetParameters(fxeq(5), p)
p.fCenter = 1838
Call BASS_FXSetParameters(fxeq(6), p)
p.fCenter = 3675
Call BASS_FXSetParameters(fxeq(7), p)
p.fCenter = 7350
Call BASS_FXSetParameters(fxeq(8), p)
p.fCenter = 14700
Call BASS_FXSetParameters(fxeq(9), p)
Dim ix As Integer
For ix = 0 To 9
Call Eq_Change(ix)
Next ix
Case 0
For u = 0 To 9
Call BASS_ChannelRemoveFX(chan, fxeq(u))
Next u
End Select

End Sub

Private Sub filetimer_Timer()
Select Case MediaType
Case "MP3"
If BASS_ChannelGetPosition(chan) = BASS_StreamGetLength(chan) Then GoTo chMP3:
On Error Resume Next
Position.value = BASS_ChannelGetPosition(chan)
Tracktime = FormatTime(modBass.BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan))) & " / " & FormatTime(BASS_ChannelBytes2Seconds(chan, BASS_StreamGetLength(chan)))
Case "CD"
If BASS_ChannelGetPosition(chan) = BASS_StreamGetLength(chan) Then GoTo chCD
On Error Resume Next
Position.value = BASS_ChannelGetPosition(chan)
Tracktime = FormatTime(modBass.BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetPosition(chan))) & " / " & FormatTime(BASS_ChannelBytes2Seconds(chan, BASS_StreamGetLength(chan)))
End Select
Exit Sub
chMP3:
tracknumberX = tracknumberX + 1
If tracknumberX > Playlist1.ListItems.Count Then tracknumberX = 1
If tracknumberX <= 0 Then tracknumberX = 1
PlayMP3 tracknumberX, ""
Exit Sub
chCD:
cdtracknumberx = cdtracknumberx + 1
If cdtracknumberx > cdplaylist.ListItems.Count - 1 Then tracknumberX = 0
If cdtracknumberx <= -1 Then cdtracknumberx = 0
PlayCDTrack cdtracknumberx
Exit Sub
End Sub

Private Sub Form_Load()
vistype = 1
Set ct = New CTray
iconbox.Picture = Me.Icon
ct.PicBox = iconbox
ct.TipText = Me.Caption
ct.ShowIcon
Me.Move Me.Left, Me.Top, 450 * 15, 450 * 15
min.Picture = min1.Picture
closee.Picture = close1.Picture
xs.ShapeMe min, RGB(255, 0, 255)
xs.ShapeMe closee, RGB(255, 0, 255)
xs.ShapeMe OpenList, RGB(255, 0, 255)
xs.ShapeMe SAvelist, RGB(255, 0, 255)
ap = App.Path
If Not Right$(ap, 1) = "\" Then ap = ap & "\"
If FileExists(ap & "settings.ini") = False Then GoTo setup:
LoadSettings
Debug.Print "UseDSP" & nodsp
Debug.Print "UseVIS" & novis
Debug.Print "BadQual" & usebad
MediaType = "MP3"
With bh.bmiHeader
    .biBitCount = 8
    .biPlanes = 1
    .biSize = Len(bh.bmiHeader)
    .biWidth = 240
    .biHeight = 120  'upside down (line 0=bottom)
    .biClrUsed = 256
    .biClrImportant = 256
End With
Dim a As Long
    For a = 1 To 127
        bh.bmiColors(a).rgbGreen = 255 - 2 * a
        bh.bmiColors(a).rgbRed = 2 * a
    Next a
    For a = 0 To 31
        bh.bmiColors(128 + a).rgbBlue = 8 * a
        bh.bmiColors(128 + 32 + a).rgbBlue = 255
        bh.bmiColors(128 + 32 + a).rgbRed = 8 * a
        bh.bmiColors(128 + 64 + a).rgbRed = 255
        bh.bmiColors(128 + 64 + a).rgbBlue = 8 * (31 - a)
        bh.bmiColors(128 + 64 + a).rgbGreen = 8 * a
        bh.bmiColors(128 + 96 + a).rgbRed = 255
        bh.bmiColors(128 + 96 + a).rgbGreen = 255
        bh.bmiColors(128 + 96 + a).rgbBlue = 8 * a
    Next a
If nodsp = 1 Then
        EQOn.value = 0
        EQOn.Enabled = 0
        rvbon.value = 0
        rvbon.Enabled = 0
        rvb(0).Enabled = 0
        rvb(1).Enabled = 0
        Eq(0).Enabled = 0
        Eq(1).Enabled = 0
        Eq(2).Enabled = 0
        Eq(3).Enabled = 0
        Eq(4).Enabled = 0
        Eq(5).Enabled = 0
        Eq(6).Enabled = 0
        Eq(7).Enabled = 0
        Eq(8).Enabled = 0
        Eq(9).Enabled = 0
GoTo skipdx8chk:
End If

BASS_Init 1, 44100, 0, Me.hWnd, 0
    Dim bi As BASS_INFO
    bi.size = LenB(bi)      'LenB(..) returns a byte data
    Call BASS_GetInfo(bi)
    DXVerX = bi.dsver
    If DXVerX < 8 Then
        Call BASS_Free
        Titlelbl = "DirectX 8 not installed, disabling EQ and Reverb"
        EQOn.value = 0
        EQOn.Enabled = 0
        rvbon.value = 0
        rvbon.Enabled = 0
        rvb(0).Enabled = 0
        rvb(1).Enabled = 0
        Eq(0).Enabled = 0
        Eq(1).Enabled = 0
        Eq(2).Enabled = 0
        Eq(3).Enabled = 0
        Eq(4).Enabled = 0
        Eq(5).Enabled = 0
        Eq(6).Enabled = 0
        Eq(7).Enabled = 0
        Eq(8).Enabled = 0
        Eq(9).Enabled = 0
    Else
    End If
If nodsp = 1 Then
        EQOn.value = 0
        EQOn.Enabled = 0
        rvbon.value = 0
        rvbon.Enabled = 0
        rvb(0).Enabled = 0
        rvb(1).Enabled = 0
        Eq(0).Enabled = 0
        Eq(1).Enabled = 0
        Eq(2).Enabled = 0
        Eq(3).Enabled = 0
        Eq(4).Enabled = 0
        Eq(5).Enabled = 0
        Eq(6).Enabled = 0
        Eq(7).Enabled = 0
        Eq(8).Enabled = 0
        Eq(9).Enabled = 0
End If
 BASS_Free
skipdx8chk:
ReadList ap & "playlist.mpx"
If Not tracknumberX <= 0 Then Playlist1.ListItems(tracknumberX).Selected = 1
Exit Sub
setup:
On Error GoTo cannotcreate:
SaveSettings
Exit Sub
'Open ap & "settings.ini" For Output As #1'
'Print #1, "[MPX2]"
'For u = 1 To 10
'Print #1, "EQ" & u & "=150"
'Next u
'Print #1, "reverbdelay=1500"
'Print #1, "reverbstrength=50"
'Print #1, "IsEQOn=0"
'Print #1, "Isreverbon=0"
'Print #1, "Volume=20"
'Print #1, "Cursong=0"
'Close #1
'Exit Sub
cannotcreate:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then xs.DragForm Form1.hWnd, 1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Form2.mnufrm
End Sub


Private Sub min_Click()
vistimer.Enabled = 0
filetimer.Interval = 250
Me.Hide
End Sub

Private Sub min_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = min2.Picture
End Sub

Private Sub min_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = min1.Picture
End Sub

Public Sub PlayMP3(tracknumber As Long, Optional customdspfile As String)
Dim freq As Long, mono As Long
mono = 0
freq = 44100
If usebad = 1 Then
freq = 22050
mono = BASS_DEVICE_MONO
End If
BASS_ChannelRemoveFX chan, fxr
For u = 0 To 9
BASS_ChannelRemoveFX chan, fxeq(u)
Next u
BASS_Free
BASS_Init 1, freq, mono, Form1.hWnd, 0
Call BASS_SetConfig(BASS_CONFIG_FLOATDSP, 1)
On Error GoTo noload:
If FileExists(filename1.ListItems(tracknumber)) = False Then Exit Sub
If Not UCase$(Right$(filename1.ListItems(tracknumber).text, 3)) = "WMA" Then
        If DXVerX <= 8 Then
            chan = BASS_StreamCreateFile(BASSFALSE, filename1.ListItems(tracknumber).text, 0, 0, BASS_STREAM_AUTOFREE)
        Else
            chan = BASS_StreamCreateFile(BASSFALSE, filename1.ListItems(tracknumber).text, 0, 0, BASS_SAMPLE_FLOAT Or BASS_STREAM_AUTOFREE)
        End If
Else
        If DXVerX <= 8 Then
            chan = BASS_WMA_StreamCreateFile(BASSFALSE, filename1.ListItems(tracknumber).text, 0, 0, BASS_STREAM_AUTOFREE)
        Else
            chan = BASS_WMA_StreamCreateFile(BASSFALSE, filename1.ListItems(tracknumber).text, 0, 0, BASS_SAMPLE_FLOAT Or BASS_STREAM_AUTOFREE)
        End If
End If

Position.max = BASS_StreamGetLength(chan)
Playlist1.ListItems(tracknumber).SubItems(1) = FormatTime(BASS_ChannelBytes2Seconds(chan, BASS_StreamGetLength(chan)))
filename1.ListItems(tracknumber).SubItems(1) = BASS_ChannelBytes2Seconds(chan, BASS_StreamGetLength(chan))
s = GetFileName(filename1.ListItems(tracknumber).text)
If Not UCase$(Right$(filename1.ListItems(tracknumber).text, 3)) = "WAV" Then
LoadMp3File filename1.ListItems(tracknumber).text
If TagExists = True Then
Playlist1.ListItems(tracknumber).text = Title
Else
Playlist1.ListItems(tracknumber).text = Left$(s, Len(s) - 4)
End If
Else
Playlist1.ListItems(tracknumber).text = Left$(s, Len(s) - 4)
End If
Call rvbon_Click
Call EQOn_Click
Call Volume_Change
MediaType = "MP3"
Playlist1.ListItems(tracknumberX).Selected = 1
BASS_StreamPlay chan, 0, 0
Titlelbl = Playlist1.ListItems(tracknumber)
Form1.Caption = "MP3PlayerX2 - Now Playing[" & Titlelbl & "]"
ct.TipText = Me.Caption
filetimer.Interval = 1
vistimer.Interval = 1
filetimer.Enabled = 1
vistimer.Enabled = 1
Dim bcc As BASS_CHANNELINFO
Dim us As String
Call BASS_ChannelGetInfo(chan, bcc)
If bcc.flags And BASS_SAMPLE_FLOAT Then
us = "32 Bit"
Else
us = "16 Bit"
End If
Label4 = "Stream - " & Round((FileLen(filename1.ListItems(tracknumber).text) / BASS_ChannelBytes2Seconds(chan, BASS_StreamGetLength(chan))) / 125) & "kbps   Quality:" & us
GetListLen
Exit Sub
noload:
BASS_Free

End Sub
Public Function GetFileName(dd As String) As String
GetFileName = Mid(dd, InStrRev(dd, "\") + 1)
End Function
Public Function FileExists(ByVal FileName As String) As Boolean
On Local Error Resume Next
FileExists = (Dir$(FileName) <> "")
End Function

Private Sub mp3click_Click()
cdplaylist.Visible = 0
Playlist1.Visible = 1
End Sub

Private Sub MVDWN_Click()
On Error Resume Next
Dim lst As Long
Dim lst2
Dim m1, m2, m3, m4
lst = Playlist1.SelectedItem.Index
If lst = Playlist1.ListItems.Count Then
lst2 = 1
Else
lst2 = lst + 1
End If

m1 = Playlist1.ListItems(lst).text
m2 = Playlist1.ListItems(lst).SubItems(1)
m3 = Playlist1.ListItems(lst2).text
m4 = Playlist1.ListItems(lst2).SubItems(1)

Playlist1.ListItems(lst).text = m3
Playlist1.ListItems(lst).SubItems(1) = m4

Playlist1.ListItems(lst2).text = m1
Playlist1.ListItems(lst2).SubItems(1) = m2


m1 = filename1.ListItems(lst).text
m2 = filename1.ListItems(lst2).text
m3 = filename1.ListItems(lst).SubItems(1)
m4 = filename1.ListItems(lst2).SubItems(1)

filename1.ListItems(lst).text = m2
filename1.ListItems(lst).SubItems(1) = m4
filename1.ListItems(lst2).text = m1
filename1.ListItems(lst2).SubItems(1) = m3

Playlist1.ListItems(lst2).Selected = True


End Sub

Private Sub MVUP_Click()
On Error Resume Next
Dim lst As Long
Dim lst2 As Long
Dim m1, m2, m3, m4
lst = Playlist1.SelectedItem.Index
If lst = 1 Then
lst2 = Playlist1.ListItems.Count
Else
lst2 = lst - 1
End If

m1 = Playlist1.ListItems(lst).text
m2 = Playlist1.ListItems(lst).SubItems(1)
m3 = Playlist1.ListItems(lst2).text
m4 = Playlist1.ListItems(lst2).SubItems(1)

Playlist1.ListItems(lst).text = m3
Playlist1.ListItems(lst).SubItems(1) = m4

Playlist1.ListItems(lst2).text = m1
Playlist1.ListItems(lst2).SubItems(1) = m2


m1 = filename1.ListItems(lst).text
m2 = filename1.ListItems(lst2).text
m3 = filename1.ListItems(lst).SubItems(1)
m4 = filename1.ListItems(lst2).SubItems(1)

filename1.ListItems(lst).text = m2
filename1.ListItems(lst).SubItems(1) = m4
filename1.ListItems(lst2).text = m1
filename1.ListItems(lst2).SubItems(1) = m3

Playlist1.ListItems(lst2).Selected = True

End Sub

Private Sub OpenList_Click()
ListDialog.CancelError = 1
On Error GoTo nofilesel
ListDialog.DialogTitle = "Open Playlist..."
ListDialog.ShowOpen
ReadList ListDialog.FileName
nofilesel:
End Sub

Public Sub PlayB_Click()
Select Case MediaType
Case "MP3"
PlayMP3 (tracknumberX)
Case "CD"
PlayCDTrack (cdtracknumberx)
End Select
End Sub

Private Sub PlayB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PlayB.Picture = Play2.Picture
End Sub

Private Sub PlayB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PlayB.Picture = Play1.Picture
End Sub

Private Sub Playlist1_DblClick()
If Playlist1.ListItems.Count = 0 Then Exit Sub
tracknumberX = Playlist1.SelectedItem.Index
PlayMP3 tracknumberX

End Sub
Public Sub ReadList(pathname As String, Optional Noclearentries As Integer)
Dim totaltime As Long
Dim entrnum As Long
Dim m3ustr As String
Dim m3ustr2 As String
Dim m3ustr3 As String
Dim m3ustr4 As String
Dim strin$
If Noclearentries = 1 Then GoTo justadd:
Playlist1.ListItems.Clear
filename1.ListItems.Clear
justadd:
Select Case UCase$(Right$(pathname, 3))
Case "MPX"
Close #1
If FileExists(pathname) = False Then Close: Exit Sub
Open pathname For Input As #1
Line Input #1, m3ustr
For entrnum = 1 To Val(m3ustr)
Line Input #1, m3ustr2 'filename
Line Input #1, m3ustr3 'title
Line Input #1, m3ustr4 'length
If m3ustr4 = "" Then m3ustr4 = 0
filename1.ListItems.Add entrnum, , m3ustr2
Playlist1.ListItems.Add entrnum, , m3ustr3
filename1.ListItems(entrnum).SubItems(1) = Val(m3ustr4)
Playlist1.ListItems(entrnum).SubItems(1) = FormatTime(Val(m3ustr4))
Next
Close
GetListLen
Case "M3U"
Close
If FileExists(pathname) = False Then Exit Sub
Open pathname For Input As #1
u = 1
Do While Not EOF(1)
Line Input #1, m3ustr
If Left$(m3ustr, 8) = "#EXTINF:" Then 'dont do anything - extremely fast loading
Else
If m3ustr = "#EXTM3U" Then GoTo dontdo
filename1.ListItems.Add u, , m3ustr
Playlist1.ListItems.Add u, , (m3ustr)
filename1.ListItems(u).SubItems(1) = 0
Playlist1.ListItems(u).SubItems(1) = (filename1.ListItems(u).SubItems(1))
u = u + 1
dontdo:
End If
GetListLen
Loop

End Select
Exit Sub
founderr:
End Sub


Public Sub WriteList(ListFile As String, ListType As Long)
Select Case ListType
Dim PL As Integer
Case 1
Close #1
If Playlist1.ListItems.Count = 0 Then Exit Sub
Open ListFile For Output As #1
Print #1, Playlist1.ListItems.Count
For PL = 1 To Playlist1.ListItems.Count
Print #1, filename1.ListItems(PL).text
Print #1, Playlist1.ListItems(PL).text
Print #1, filename1.ListItems(PL).SubItems(1)
Next
Close #1
Case 2
Close
Open ListFile For Output As #1
Print #1, "#EXTM3U"
For PL = 1 To Playlist1.ListItems.Count
Print #1, filename1.ListItems(PL).text
Next PL
Close #1
Case 3
IniFileX = ListFile
WriteValue "Playlist", "NumberOfEntries", Playlist1.ListItems.Count
For PL = 1 To Playlist1.ListItems.Count
WriteValue "Playlist", "FILE" & PL, filename1.ListItems(PL).text
Next

End Select
End Sub


Private Sub Playlist1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Playlist1.ListItems.Count = 0 Then If Button = 2 Then PopupMenu Form2.mnump3
End Sub

Private Sub Position_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Position.value = Position.max / Position.Width * X / 15
Call BASS_ChannelSetPosition(chan, Position.value)
End Sub

Private Sub Randomizer_Click()
vistimer.Enabled = 0
RandList
If Not MediaType = "CONV" Or nodsp = 1 Then vistimer.Enabled = 1
End Sub

Private Sub RemoveSong_Click()
On Error Resume Next
Dim REMNUM As Long
REMNUM = Form1.Playlist1.SelectedItem.Index
Form1.Playlist1.ListItems.Remove REMNUM
Form1.filename1.ListItems.Remove REMNUM
If tracknumberX >= REMNUM Then tracknumberX = tracknumberX - 1
End Sub

Private Sub resetimg_Click()
For u = 0 To 9
Eq(u).value = 149
Next u
End Sub

Private Sub rvb_Change(Index As Integer)
Dim p1 As BASS_FXREVERB
    Call BASS_FXGetParameters(fxr, p1)
       p1.fReverbMix = -0.12 * (8000 / rvb(1).value)
       p1.fReverbTime = rvb(0).value
       p1.fInGain = 0
    Call BASS_FXSetParameters(fxr, p1)
Call BASS_FXSetParameters(fxr, p1)
End Sub

Private Sub rvb_Scroll(Index As Integer)
Call rvb_Change(Index)
End Sub

Public Sub rvbon_Click()
Select Case rvbon.value
Case 1
fxr = BASS_ChannelSetFX(chan, BASS_FX_REVERB)
Call rvb_Change(0)
Case 0
Call BASS_ChannelRemoveFX(chan, fxr)
End Select
End Sub
Public Function FormatTime(ByVal sec As Long) As String
Dim s As Long
Dim M As Long
Dim H As Long
s = sec
M = 0
H = 0
If s >= 60 Then
M = Int(s / 60)
s = s - M * 60
End If
If M >= 60 Then
H = Int(M / 60)
M = M - H * 60
End If
If H > 0 Then
FormatTime = Format$(H, "00") & ":"
End If
FormatTime = FormatTime & Format$(M, "00") & ":" & Format$(s, "00")
End Function
Public Sub RandList()
Dim intRandomnumber As Long
        On Error Resume Next
        Randomize
        intNumberOfEntries = Form1.filename1.ListItems.Count
        For Y = 1 To 2
        For X = 1 To intNumberOfEntries
        intRandomnumber = Int(intNumberOfEntries * Rnd) + 1
        tmptext = Form1.filename1.ListItems(X).text
        tmptext2 = Form1.filename1.ListItems(intRandomnumber).text
        tmptext3 = Form1.filename1.ListItems(X).SubItems(1)
        tmptext4 = Form1.filename1.ListItems(intRandomnumber).SubItems(1)
        tmptext5 = Form1.Playlist1.ListItems(X).text
        tmptext6 = Form1.Playlist1.ListItems(intRandomnumber).text
        tmptext7 = Form1.Playlist1.ListItems(X).SubItems(1)
        tmptext8 = Form1.Playlist1.ListItems(intRandomnumber).SubItems(1)
        Form1.filename1.ListItems(X).text = tmptext2
        Form1.filename1.ListItems(X).SubItems(1) = tmptext4
        Form1.Playlist1.ListItems(X).text = tmptext6
        Form1.Playlist1.ListItems(X).SubItems(1) = tmptext8
        
        Form1.filename1.ListItems(intRandomnumber).text = tmptext
        Form1.filename1.ListItems(intRandomnumber).SubItems(1) = tmptext3
        Form1.Playlist1.ListItems(intRandomnumber).text = tmptext5
        Form1.Playlist1.ListItems(intRandomnumber).SubItems(1) = tmptext7
        Next X
        Next Y

End Sub
Public Sub NextTrack()
Select Case MediaType
Case "MP3"
tracknumberX = tracknumberX + 1
If tracknumberX > filename1.ListItems.Count Then tracknumberX = 1
PlayMP3 (tracknumberX)
Case "CD"
cdtracknumberx = cdtracknumberx + 1
If cdtracknumberx > cdplaylist.ListItems.Count - 1 Then cdtracknumberx = 0
PlayCDTrack cdtracknumberx
End Select
End Sub
Public Sub PrevTrack()
Select Case MediaType
Case "MP3"
tracknumberX = tracknumberX - 1
If tracknumberX <= 0 Then tracknumberX = Playlist1.ListItems.Count
PlayMP3 (tracknumberX)
Case "CD"
cdtracknumberx = cdtracknumberx - 1
If cdtracknumberx <= -1 Then cdtracknumberx = cdplaylist.ListItems.Count - 1
PlayCDTrack cdtracknumberx
End Select
End Sub

Private Sub NextTrkB_Click()
NextTrack
End Sub

Private Sub NextTrkB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
NextTrkB.Picture = Next2.Picture
End Sub

Private Sub NextTrkB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
NextTrkB.Picture = Next1.Picture
End Sub
Private Sub PrevTrkB_Click()
PrevTrack
End Sub

Private Sub PrevTrkB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PrevTrkB.Picture = Prev2.Picture
End Sub

Private Sub PrevTrkB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PrevTrkB.Picture = Prev1.Picture
End Sub
Private Sub PauResB_Click()
Pause
End Sub

Private Sub PauResB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PauResB.Picture = Pause2.Picture
End Sub

Private Sub PauResB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PauResB.Picture = Pause1.Picture
End Sub

Private Sub SAvelist_Click()
ListDialog.CancelError = 1
On Error GoTo nofilesel
ListDialog.FileName = ""
ListDialog.DialogTitle = "Save Playlist..."
ListDialog.ShowSave
WriteList ListDialog.FileName, ListDialog.FilterIndex
nofilesel:

End Sub

Private Sub StopB_Click()
StopPlay
End Sub
Private Sub StopB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopB.Picture = Stop2.Picture
End Sub

Private Sub StopB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopB.Picture = Stop1.Picture
End Sub
Public Sub StopPlay()
filetimer.Enabled = False
vistimer.Enabled = False
BASS_ChannelStop chan
For u = 0 To 9
BASS_ChannelRemoveFX chan, fxeq(u)
Next
BASS_ChannelRemoveFX chan, fxr
BASS_StreamFree chan
BASS_Free
End Sub
Public Sub Pause()
If BASS_ChannelIsActive(chan) = BASSTRUE Then
Call BASS_ChannelPause(chan)
Else
Call BASS_ChannelResume(chan)
End If
End Sub
Private Function DrawVis()
Dim i As Long
Dim X As Long
Dim Y As Long
i = 1
X = 1
vistimer.Interval = 1
Visbox.Cls
Select Case vistype
Case 2
While i < 249
Y = (data(i) ^ 0.75 + data(i + 1) ^ 0.75 + data(i + 2) ^ 0.75 + data(i + 3) ^ 0.75 + data(i + 4) ^ 0.75 + data(i + 5) ^ 0.75 + data(i + 6) ^ 0.75 + data(i + 7) ^ 0.75) / 1.83
BitBlt Visbox.hDC, X, 60 - Sin(X * 150) * Y, 1, Y, VisBar.hDC, 1, Round(1), vbSrcCopy
BitBlt Visbox.hDC, X, 60 - (-Sin(X * 150) * Y), 1, Y * 0.83, VisBar.hDC, 1, Round(1), vbSrcCopy
X = X + 1
i = i + 1
Wend
Case 3
While i < 249
Y = (data(i) ^ 0.75 + data(i + 1) ^ 0.75 + data(i + 2) ^ 0.75 + data(i + 3) ^ 0.75 + data(i + 4) ^ 0.75 + data(i + 5) ^ 0.75 + data(i + 6) ^ 0.75 + data(i + 7) ^ 0.75) * 1.3
BitBlt Visbox.hDC, X, 120 - Sin(X * 150) * Y, 1, Y, VisBar.hDC, 1, Round(1), vbSrcCopy
BitBlt Visbox.hDC, X, 120 - (-Sin(X * 150) * Y), 1, Y * 0.83, VisBar.hDC, 1, Round(1), vbSrcCopy
X = X + 1
i = i + 1
Wend
Case 4
While i < 249
Y = (data(i) ^ 0.75 + data(i + 1) ^ 0.75 + data(i + 2) ^ 0.75 + data(i + 3) ^ 0.75 + data(i + 4) ^ 0.75 + data(i + 5) ^ 0.75 + data(i + 6) ^ 0.75 + data(i + 7) ^ 0.75) / 1.333
On Error Resume Next
BitBlt Visbox.hDC, X, 60 - Sin(X * 3) * Y, 1, Y / 8, VisBar.hDC, 1, Round(Y), vbSrcCopy
BitBlt Visbox.hDC, X, 60 - -Sin(X * 3) * Y, 1, Y / 8, VisBar.hDC, 1, Round(Y), vbSrcCopy
X = X + 1
i = i + 1
Wend
Case 5
While i < 249
Y = (data(i) ^ 0.75 + data(i + 1) ^ 0.75 + data(i + 2) ^ 0.75 + data(i + 3) ^ 0.75 + data(i + 4) ^ 0.75 + data(i + 5) ^ 0.75 + data(i + 6) ^ 0.75 + data(i + 7) ^ 0.75)
BitBlt Visbox.hDC, X, 120 - Y, 3, Y, VisBar.hDC, 1, Round(1), vbSrcCopy
X = X + 4
i = i + 4
Wend
End Select
End Function

Private Sub Visbox_Click()
If vistype = 5 Then vistype = 0
vistype = vistype + 1
If vistype = 1 Then Visbox.AutoRedraw = 0
If vistype = 2 Then Visbox.AutoRedraw = 1
If vistype = 3 Then Visbox.AutoRedraw = 1
If vistype = 4 Then Visbox.AutoRedraw = 1
If vistype = 5 Then Visbox.AutoRedraw = 1

End Sub

Private Sub vistimer_Timer()
If novis = 1 Then Exit Sub
Select Case vistype
Case 1
Call BASS_VIS_CODE(0, 0, 0, 0, 0)
Case 2, 3, 4, 5
VisCode
End Select
End Sub

Private Sub Volume_Change()
Call BASS_ChannelSetAttributes(chan, 0, Volume.value, 0)
End Sub

Private Sub Volume_Scroll()
Call Volume_Change
End Sub
Sub VisCode()
If Form1.WindowState = vbMinimized Then Exit Sub
Dim lRslt As Long
Dim fftVals(512) As Single
Dim i As Integer
If chan Then
lRslt = BASS_ChannelGetData(chan, fftVals(0), BASS_DATA_FFT512)
If lRslt <> BASSFALSE And lRslt > BASSFALSE Then
For i = 1 To 255
Data2(i) = (fftVals(i))
If (Data2(i) + (Data2(i) * (i * 0.25))) * 40 > data(i) Then
data(i) = (Data2(i) + (Data2(i) * (i * 0.25))) * 40
If data(i) > 39 Then data(i) = 39
Else
If data(i) > 1 Then
data(i) = data(i) - 1
Else
data(i) = 0
End If
End If
Next

DrawVis
End If
End If

End Sub
Function CountFilesInList(ByVal FileList As String) As Long
    Dim iCount As Integer
    Dim iPos As Integer

    iCount = 0
    For iPos = 1 To Len(FileList)
        If Mid$(FileList, iPos, 1) = Chr$(0) Then iCount = iCount + 1
    Next
    If iCount = 0 Then iCount = 1
    CountFilesInList = iCount
End Function

Function GetFileFromList(ByVal FileList As String, FileNumber As Long) As String
' Get filename of FileNumber from FileList
    Dim iPos                As Long
    Dim iCount              As Long
    Dim iFileNumberStart    As Long
    Dim iFileNumberLen      As Long
    Dim sPath               As String

    If InStr(FileList, Chr(0)) = 0 Then
        GetFileFromList = FileList
    Else
        iCount = 0
        sPath = Left(FileList, InStr(FileList, Chr(0)) - 1)
        If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
        FileList = FileList + Chr(0)
        For iPos = 1 To Len(FileList)
            If Mid$(FileList, iPos, 1) = Chr(0) Then
                iCount = iCount + 1
                Select Case iCount
                    Case FileNumber
                        iFileNumberStart = iPos + 1
                    Case FileNumber + 1
                        iFileNumberLen = iPos - iFileNumberStart
                        Exit For
                End Select
            End If
        Next
        GetFileFromList = sPath + Mid(FileList, iFileNumberStart, iFileNumberLen)
    End If
End Function

Sub BASS_VIS_CODE(ByVal uTimerID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
Dim X As Long, Y As Long, Y1 As Long
Call BASS_ChannelGetData(chan, fft(0), BASS_DATA_FFT2048)
        ReDim specbuf(240 * (119 + 1)) As Byte 'clear display
        For X = 0 To (240 / 2) - 1
          Y = Sqrt(fft(X + 1)) * 2.5 * 119 - 4
            If (Y > 119) Then Y = 119 ' cap it
            If (X) Then  'interpolate from previous to make the display smoother
                Y1 = (Y + Y1) / 2
                
                While (Y1 >= 0)
                    specbuf(Y1 * 240 + X * 2 - 1) = Y1 + 1
                    Y1 = Y1 - 1
                Wend
            End If
            Y1 = Y
            While (Y >= 0)
                specbuf(Y * 240 + X * 2) = Y + 1 ' draw level
                Y = Y - 1
            Wend
        Next X
    Call SetDIBitsToDevice(Visbox.hDC, 0, 1, 240, 119, 0, 0, 0, 119, specbuf(0), bh, 0)
End Sub

Public Function Sqrt(ByVal num As Double) As Double 'its not really sq root
    Sqrt = num ^ 0.5
End Function

Public Sub CheckCD()
If Not BASS_CD_IsReady(0) = 1 Then Exit Sub
cdplaylist.ListItems.Clear
If BASS_CD_GetTracks(0) <= 0 Then Exit Sub
For u = 0 To BASS_CD_GetTracks(0) - 1
If Not BASS_CD_GetTrackLength(0, u) = -1 Then
cdplaylist.ListItems.Add u + 1, , "Audio Track " & u + 1
cdplaylist.ListItems(u + 1).tag = u
cdplaylist.ListItems(u + 1).SubItems(1) = FormatTime(BASS_CD_GetTrackLength(0, u) / 176000)
Else
cdplaylist.ListItems.Add u + 1, , "Data Track " & u + 1
cdplaylist.ListItems(u + 1).tag = "DATA"
cdplaylist.ListItems(u + 1).SubItems(1) = "DATA"
End If
Next u

End Sub
Public Sub PlayCDTrack(cdtracknumber As Long)
Dim freq As Long, mono As Long
mono = 0
freq = 44100
If usebad = 1 Then
freq = 22050
mono = BASS_DEVICE_MONO
End If
BASS_ChannelRemoveFX chan, fxr
For u = 0 To 9
BASS_ChannelRemoveFX chan, fxeq(u)
Next u
BASS_Free
BASS_Init 1, freq, mono, Form1.hWnd, 0
Call BASS_SetConfig(BASS_CONFIG_FLOATDSP, 1)
If DXVerX <= 8 Then
BASS_StreamFree (chan): chan = BASS_CD_StreamCreate(0, cdtracknumber, BASS_STREAM_AUTOFREE)
Else
BASS_StreamFree (chan): chan = BASS_CD_StreamCreate(0, cdtracknumber, BASS_SAMPLE_FLOAT Or BASS_STREAM_AUTOFREE)
End If

If chan = 0 Then BASS_StreamFree (chan): BASS_Free: Exit Sub
'add FX
Call rvbon_Click
Call EQOn_Click
Call Volume_Change
MediaType = "CD"
Call BASS_StreamPlay(chan, 0, 0)
cdplaylist.ListItems(cdtracknumber + 1).Selected = 1
Titlelbl = "Audio Track " & cdtracknumber + 1
Position.max = BASS_StreamGetLength(chan)
Label4 = "CD - 1411kbps"
Position.value = 0
vistimer.Enabled = 1
filetimer.Enabled = 1
cdtracknumberx = cdtracknumber
Form1.Caption = "MP3PlayerX2 - Now Playing[CD Audio Track" & cdtracknumber + 1 & "]"
ct.TipText = Me.Caption
End Sub
Public Sub GetListLen()
If Playlist1.ListItems.Count = 0 Then Exit Sub
totaltime = 0
For u = 1 To Playlist1.ListItems.Count
totaltime = totaltime + filename1.ListItems(u).SubItems(1)
Next u
Label3 = "Total Time: " & FormatTime(totaltime)
End Sub
Public Sub LoadSettings()
On Error Resume Next
IniFileX = ap & "Settings.ini"
tracknumberX = Val(GetValue("MPX2", "Cursong"))
Volume.value = GetValue("MPX2", "Volume")
rvb(0).value = GetValue("MPX2", "ReverbDelay")
rvb(1).value = GetValue("MPX2", "ReverbStrength")
For u = 0 To 9
Eq(u).value = GetValue("MPX2", "EQ" & u + 1)
Next u
EQOn.value = GetValue("MPX2", "IsEqOn")
rvbon.value = GetValue("MPX2", "IsReverbOn")
usebad = GetValue("MPX2", "usebad")
nodsp = GetValue("MPX2", "NODSP")
novis = GetValue("MPX2", "NOVIS")
Form5.usebad.value = usebad
Form5.useDX8.value = nodsp
Form5.useVis.value = novis
End Sub
Public Sub SaveSettings()
On Error Resume Next:
IniFileX = ap & "Settings.ini"
WriteList ap & "Playlist.mpx", 1
WriteValue "MPX2", "Volume", Volume.value
WriteValue "MPX2", "ReverbStrength", rvb(1).value
WriteValue "MPX2", "ReverbDelay", rvb(0).value
WriteValue "MPX2", "EQ1", Eq(0).value
WriteValue "MPX2", "EQ2", Eq(1).value
WriteValue "MPX2", "EQ3", Eq(2).value
WriteValue "MPX2", "EQ4", Eq(3).value
WriteValue "MPX2", "EQ5", Eq(4).value
WriteValue "MPX2", "EQ6", Eq(5).value
WriteValue "MPX2", "EQ7", Eq(6).value
WriteValue "MPX2", "EQ8", Eq(7).value
WriteValue "MPX2", "EQ9", Eq(8).value
WriteValue "MPX2", "EQ10", Eq(9).value
WriteValue "MPX2", "CurSong", Val(Playlist1.SelectedItem.Index)
WriteValue "MPX2", "IsEQOn", EQOn.value
WriteValue "MPX2", "IsReverbOn", rvbon.value
End Sub
