Attribute VB_Name = "BASSCD"

' BASSCD 2.0 Visual Basic API Header File
' Requires BASS.DLL & BASS.BAS 2.0 - available @ www.un4seen.com
Global Const BASS_ERROR_NOCD = 12      ' no CD in drive
Global Const BASS_ERROR_CDTRACK = 13   ' invalid track number
Global Const BASS_ERROR_NOTAUDIO = 17  ' not an audio track
Global Const BASS_CD_FREEOLD = &H10000
Global Const BASS_CD_SUBCHANNEL = &H20000
Global Const BASS_SYNC_CD_ERROR = 1000
Global Const BASS_CD_DOOR_CLOSE = 0
Global Const BASS_CD_DOOR_OPEN = 1
Global Const BASS_CD_DOOR_LOCK = 2
Global Const BASS_CD_DOOR_UNLOCK = 3
Global Const BASS_CDID_UPC = 1
Global Const BASS_CDID_CDDB = 2
Global Const BASS_CDID_CDDB2 = 3
Global Const BASS_CDID_TEXT = 4
Global Const BASS_CDID_CDPLAYER = 5
Global Const BASS_CHANNEL_STREAM_CD = &H10200
Declare Function BASS_CD_GetDriveDescription Lib "basscd.dll" (ByVal drive As Long) As Long
Declare Function BASS_CD_GetDriveLetter Lib "basscd.dll" (ByVal drive As Long) As Long
Declare Function BASS_CD_Door Lib "basscd.dll" (ByVal drive As Long, ByVal action As Long) As Long
Declare Function BASS_CD_DoorIsLocked Lib "basscd.dll" (ByVal drive As Long) As Long
Declare Function BASS_CD_DoorIsOpen Lib "basscd.dll" (ByVal drive As Long) As Long
Declare Function BASS_CD_IsReady Lib "basscd.dll" (ByVal drive As Long) As Long
Declare Function BASS_CD_GetTracks Lib "basscd.dll" (ByVal drive As Long) As Long
Declare Function BASS_CD_GetTrackLength Lib "basscd.dll" (ByVal drive As Long, ByVal track As Long) As Long
Declare Function BASS_CD_GetID Lib "basscd.dll" (ByVal drive As Long, ByVal id As Long) As Long
Declare Function BASS_CD_StreamCreate Lib "basscd.dll" (ByVal drive As Long, ByVal track As Long, ByVal flags As Long) As Long
Declare Function BASS_CD_StreamCreateFile Lib "basscd.dll" (ByVal f As String, ByVal flags As Long) As Long
Declare Function BASS_CD_StreamGetTrack Lib "basscd.dll" (ByVal handle As Long) As Long

