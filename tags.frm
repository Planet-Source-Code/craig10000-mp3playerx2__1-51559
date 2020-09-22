VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Editor"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "tags.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   825
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2340
      Width           =   2220
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   825
      MaxLength       =   30
      TabIndex        =   11
      Top             =   2010
      Width           =   3240
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   825
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1710
      Width           =   840
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   825
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1380
      Width           =   3240
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   825
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1050
      Width           =   3240
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   825
      MaxLength       =   30
      TabIndex        =   3
      Top             =   720
      Width           =   3240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write Tags"
      Height          =   330
      Left            =   3060
      TabIndex        =   0
      Top             =   2325
      Width           =   1050
   End
   Begin VB.Label Label7 
      Caption         =   "Filename"
      Height          =   240
      Left            =   15
      TabIndex        =   12
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label Label6 
      Caption         =   "Comment"
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label Label5 
      Caption         =   "Year"
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   1740
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Album"
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   1410
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Artist"
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Title"
      Height          =   270
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"tags.frx":000C
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   15
      Width           =   4110
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub LoadTags(iFilename As String)
LoadMp3File iFilename
Text1.text = Title
Text2.text = Artist
Text3.text = Album
Text4.text = Year
Text5.text = Comment
Text6 = iFilename
End Sub
Public Sub SaveTags()
LoadMp3File Text6
Title = Text1
Artist = Text2
Album = Text3
Year = Text4
Comment = Text5
CloseMp3File
Form4.Hide
End Sub

Private Sub Command1_Click()
SaveTags
End Sub
