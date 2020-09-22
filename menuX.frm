VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "MenuForm"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "menuX.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnucd 
      Caption         =   "cd"
      Begin VB.Menu mnucdplay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnucdeatt 
         Caption         =   "Exit After This Track"
      End
      Begin VB.Menu mnucdse1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucdconvert 
         Caption         =   "Convert..."
      End
   End
   Begin VB.Menu mnump3 
      Caption         =   "mp3"
      Begin VB.Menu mnump3play 
         Caption         =   "Play"
      End
      Begin VB.Menu mnump3eatt 
         Caption         =   "Exit After This Track"
      End
      Begin VB.Menu mnump3se1 
         Caption         =   "-"
      End
      Begin VB.Menu mnump3tags 
         Caption         =   "Edit Tags"
      End
      Begin VB.Menu mnump3convert 
         Caption         =   "Convert..."
      End
   End
   Begin VB.Menu mnutray 
      Caption         =   "tray"
      Begin VB.Menu mnutrplay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnutrpause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnutrStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnutrse1 
         Caption         =   "-"
      End
      Begin VB.Menu mnutrprev 
         Caption         =   "Previous Track"
      End
      Begin VB.Menu mnutrnext 
         Caption         =   "Next Track"
      End
      Begin VB.Menu mnutrse2 
         Caption         =   "-"
      End
      Begin VB.Menu mnutrsm 
         Caption         =   "Show MP3PlayerX2"
      End
      Begin VB.Menu mnutrexit 
         Caption         =   "Exit MP3PlayerX2"
      End
      Begin VB.Menu mnutrse3 
         Caption         =   "-"
      End
      Begin VB.Menu mnutrcancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnufrm 
      Caption         =   "form"
      Begin VB.Menu mnufrmabout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnufrmsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufrmAdvOp 
         Caption         =   "Advanced Options..."
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnucdconvert_Click()
If Not Form1.cdplaylist.ListItems.Count = 0 Then
Form3.inputx = Form1.cdplaylist.SelectedItem.tag
Form3.bitratesld.value = 160
Form3.Check1.value = 0
Form3.Check2.value = 0
Form3.inputtype = True
Form3.Show
End If
End Sub

Private Sub mnufrmabout_Click()
MsgBox "MP3PlayerX2 v" & App.Major & "." & App.Minor & " ©2004 Craig Anderson, All Rights Reserved." & vbCrLf & vbCrLf & "This software utilises the BASS Audio System, more information at www.un4seen.com" & vbCrLf & "LAME 3.93.1 for MP3 Encoding, more at www.mp3dev.org" & vbCrLf & "OGGTOOLS for Ogg Vorbis encoding, more at www.vorbis.com" & vbCrLf & "Windows Media Audio, information and downloads at microsoft.com" & vbCrLf & vbCrLf & "BASS Audio System is draining: " & Round(BASS_GetCPU(), 2) & "% of your CPU"
End Sub

Private Sub mnufrmAdvOp_Click()
Form5.Show

End Sub

Private Sub mnump3convert_Click()
If Not Form1.Playlist1.ListItems.Count = 0 Then
Form3.inputx = Form1.filename1.ListItems(Form1.Playlist1.SelectedItem.Index).text
Form3.bitratesld.value = 160
Form3.Check1.value = 0
Form3.Check2.value = 0
Form3.inputtype = 0
Form3.Show
End If
End Sub

Private Sub mnump3tags_Click()
If Not Form1.Playlist1.ListItems.Count = 0 Then Form4.LoadTags (Form1.filename1.ListItems(Form1.Playlist1.SelectedItem.Index).text)
Form4.Show
End Sub

Private Sub mnutrexit_Click()
Form1.closee_Click
End Sub

Private Sub mnutrnext_Click()
Form1.NextTrack
End Sub

Private Sub mnutrpause_Click()
Form1.Pause
End Sub

Private Sub mnutrplay_Click()
Select Case Form1.MediaType
Case "CD"
Form1.PlayCDTrack Form1.cdtracknumberx
Case "MP3"
Form1.PlayMP3 Form1.tracknumberX
End Select
End Sub

Private Sub mnutrprev_Click()
Form1.PrevTrack
End Sub

Private Sub mnutrsm_Click()
Form1.filetimer.Interval = 1
If Not Form1.MediaType = "CONV" Then Form1.vistimer.Enabled = 1
Form1.Show
End Sub

Private Sub mnutrStop_Click()
Form1.StopPlay
End Sub
