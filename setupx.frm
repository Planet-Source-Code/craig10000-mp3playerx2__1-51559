VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3PlayerX2 Advanced Options"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "setupx.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1260
      TabIndex        =   2
      Top             =   930
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   345
      Left            =   2355
      TabIndex        =   1
      Top             =   930
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Performance/Quality"
      Height          =   915
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   3450
      Begin VB.CheckBox usebad 
         Caption         =   "Use Lower Quality Sound(Less CPU)"
         Height          =   210
         Left            =   75
         TabIndex        =   5
         ToolTipText     =   "Highly Recommended to leave off, unless you have a very old computer - a p-200 or less"
         Top             =   615
         Width           =   3015
      End
      Begin VB.CheckBox useDX8 
         Caption         =   "Disable DX8 FX (Equalizer/Reverb)"
         Height          =   210
         Left            =   75
         TabIndex        =   4
         ToolTipText     =   "Highly Recommended to leave on, Uses DX8 for Equalizer and Reverb, Lower CPU than non-dx8 FX"
         Top             =   405
         Width           =   3015
      End
      Begin VB.CheckBox useVis 
         Caption         =   "Disable Visualisations"
         Height          =   210
         Left            =   75
         TabIndex        =   3
         ToolTipText     =   "Shows the lines/bars/waves with your music"
         Top             =   195
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'read all settings, apply and hide
Close
On Error GoTo cannotcreatefile:
Open Form1.ap & "settings.ini" For Output As #1
Print #1, "[MPX2]"
For u = 1 To 10
Print #1, "EQ" & u & "=" & Form1.Eq(u - 1).value
Next u
Print #1, "ReverbDelay=" & Form1.rvb(0).value
Print #1, "ReverbStrength=" & Form1.rvb(1).value
Print #1, "UseBad=" & usebad.value
Print #1, "NODSP=" & useDX8.value
Print #1, "NOVIS=" & useVis.value
Print #1, "IsEQOn=" & Form1.EQOn.value
Print #1, "IsReverbOn=" & Form1.rvbon.value
Close
cannotcreatefile:
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

