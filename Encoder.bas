Attribute VB_Name = "Encoder"
Option Explicit
'// Structures for WAV Encoding
Public Type WAVEHEADER_RIFF    '12 bytes
    RIFF            As Long       '"RIFF" = &H46464952
    riffBlockSize   As Long       'pos + 44 - 8
    riffBlockType   As Long       '"WAVE" = &H45564157
End Type

Public Type WAVEHEADER_data    '8 bytes
   dataBlockType    As Long       '"data" = &H61746164
   dataBlockSize    As Long       'pos
End Type

Public Type WAVEFORMAT         '24 bytes
    wfBlockType     As Long       '"fmt " = &H20746D66
    wfBlockSize     As Long
    '--- block size begins from here = 16 bytes
    wFormatTag      As Integer
    nChannels       As Integer
    nSamplesPerSec  As Long
    nAvgBytesPerSec As Long
    nBlockAlign     As Integer
    wBitsPerSample  As Integer
End Type
Global buf() As Byte                           '// Buffer to hold decoding data
Public wr                       As WAVEHEADER_RIFF
Public wf                      As WAVEFORMAT
Public wd                      As WAVEHEADER_data
Public lDriveID                As Long     '// The ID of the Selected Drive
Dim ChanInfo As BASS_CHANNELINFO
Dim channels As Long

'ENCODING Provided by LAME - MP3 make sure the LAME.exe and LAMEENC.dll are in the apppath
'OGG VORBIS ENCODER - Make sure the OGGENC.exe is in hte appath
'Microsoft WMA - Needs the Windows Media 9 Codecs

Public Sub CopyToWMA(lTrack As String, lngBitrate As Long, DestFilename As String, iscd As Boolean)
On Error GoTo Err_Init

Dim flags               As Long
Dim pos                 As Long
Dim lngWMAEncodeHandle  As Long
Dim strOutputFilename   As String
Dim lngReturn           As Long
Dim i                   As Long
BASS_ChannelStop (chan)
BASS_StreamFree (chan)
For u = 0 To 9
Call BASS_ChannelRemoveFX(chan, fxeq(u))
Next u
Call BASS_ChannelRemoveFX(chan, Form1.fxr)

Form1.filetimer.Enabled = 0
Form1.vistimer.Enabled = 0
Form1.MediaType = "CONV"
BASS_Free
BASS_Init 1, 44100, 0, Form1.hWnd, 0

Select Case iscd
Case False
    If UCase$(Right$(lTrack, 3)) = "WMA" Then
    chan = BASS_WMA_StreamCreateFile(BASSFALSE, lTrack, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    Else
    chan = BASS_StreamCreateFile(BASSFALSE, lTrack, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    End If
Case True
    chan = BASS_CD_StreamCreate(0, Val(lTrack), BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
End Select
    
    Form1.filetimer.Enabled = 0
    Form1.vistimer.Enabled = 0
    
    Form1.Titlelbl = "Copying Track " & lTrack & " to WMA. Working..."
    strOutputFilename = DestFilename
    lngWMAEncodeHandle = BASS_WMA_EncodeOpenFile(44100, BASS_WMA_ENCODE_TAGS, lngBitrate, strOutputFilename)
    If lngWMAEncodeHandle = 0 Then
        MsgBox BASS_GetErrorDescription(BASS_ErrorGetCode), BASS_ErrorGetCode
        Exit Sub
    End If
    Form3.encProg.max = BASS_StreamGetLength(chan)
    Call BASS_WMA_EncodeSetTag(lngWMAEncodeHandle, "", "")
    ReDim buf(262144) As Byte
    If Form1.DXVerX >= 8 Then
        If Form3.Check1.value = 1 Then Form1.EQOn_Click
        If Form3.Check2.value = 1 Then Form1.rvbon_Click
    End If
    Do While BASS_ChannelIsActive(chan)
        ReDim Preserve buf(BASS_ChannelGetData(chan, buf(0), 262144) - 1) As Byte
        lngReturn = BASS_WMA_EncodeWrite(lngWMAEncodeHandle, VarPtr(buf(0)), UBound(buf()) + 1)
        If lngReturn = 0 Then
            Exit Do
        End If
        DoEvents
        Form3.encProg.value = BASS_ChannelGetPosition(chan)
        Form3.Label2 = BASS_ChannelGetPosition(chan) & "/" & BASS_StreamGetLength(chan) & " [Position/Length]"

    Loop
    Call BASS_WMA_EncodeClose(lngWMAEncodeHandle)
    Form1.Titlelbl = "WMA Encoding successful."

ExitRoutine:
    Erase buf()
    If chan Then Call BASS_ChannelStop(chan)
    If chan Then Call BASS_StreamFree(chan)
    If lngWMAEncodeHandle Then Call BASS_ChannelStop(lngWMAEncodeHandle)
    If lngWMAEncodeHandle Then Call BASS_StreamFree(lngWMAEncodeHandle)
    Unload Form3
    Exit Sub

Err_Init:
    GoTo ExitRoutine:

End Sub

Public Sub WavWrite(iscd As Boolean, originalfileortrack As String, outputFilename As String)
Form1.vistimer.Enabled = 0
Form1.filetimer.Enabled = 0
BASS_ChannelStop (chan)
BASS_StreamFree (chan)
Form1.MediaType = "CONV"
For u = 0 To 9
Call BASS_ChannelRemoveFX(chan, fxeq(u))
Next u
Call BASS_ChannelRemoveFX(chan, Form1.fxr)
    BASS_StreamFree chan
BASS_Free
BASS_Init 1, 44100, 0, Form1.hWnd, 0
Close
Dim pos             As Long
Dim flags           As Long
Dim ff              As Long
Select Case iscd
Case False
    If UCase$(Right$(originalfileortrack, 3)) = "WMA" Then
    chan = BASS_WMA_StreamCreateFile(BASSFALSE, originalfileortrack, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    Else
    chan = BASS_StreamCreateFile(BASSFALSE, originalfileortrack, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    End If
Case True
    chan = BASS_CD_StreamCreate(0, Val(originalfileortrack), BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
End Select
    If chan = 0 Then
        Form1.Titlelbl = "An error occured, check the " & Form1.ap & "MP3PlayerX2.log file"
        GoTo Err_Init
    End If
    
    Form1.filetimer.Enabled = 0
    Form1.vistimer.Enabled = 0
    Form1.Titlelbl = "Converting. Working..."
    
    Call BASS_ChannelGetInfo(chan, ChanInfo)
    flags = ChanInfo.flags
    channels = ChanInfo.chans

    wf.wFormatTag = 1
    wf.nChannels = channels
    Call BASS_ChannelGetAttributes(chan, wf.nSamplesPerSec, -1, -1)
    wf.wBitsPerSample = IIf(flags And BASS_SAMPLE_8BITS, 8, 16)
    wf.nBlockAlign = wf.nChannels * wf.wBitsPerSample / 8
    wf.wfBlockSize = 16
    wf.nAvgBytesPerSec = wf.nSamplesPerSec * wf.nBlockAlign
    wf.wfBlockType = &H20746D66
        
    wr.RIFF = &H46464952
    wr.riffBlockSize = 0
    wr.riffBlockType = &H45564157
    
    wd.dataBlockType = &H61746164
    wd.dataBlockSize = 0
    
    ff = FreeFile
    On Error GoTo Err_Init
    Open outputFilename For Binary Lock Read Write As #ff
    Form3.encProg.max = BASS_StreamGetLength(chan)
    Form3.Label1 = "Converter: Active, Working..."
    Put #ff, , wr
    Put #ff, , wf
    Put #ff, , wd
    
    pos = 0
    ReDim buf(262144) As Byte
    If Form1.DXVerX >= 8 Then
        If Form3.Check1.value = 1 Then Call Form1.EQOn_Click
        If Form3.Check2.value = 1 Then Call Form1.rvbon_Click
    End If
    While BASS_ChannelIsActive(chan)
        ReDim Preserve buf(BASS_ChannelGetData(chan, buf(0), 262144) - 1) As Byte
        Put #ff, , buf
        pos = BASS_ChannelGetPosition(chan)

        DoEvents
        Form3.encProg.value = BASS_ChannelGetPosition(chan)
        Form3.Label2 = BASS_ChannelGetPosition(chan) & "/" & BASS_StreamGetLength(chan) & " [Position/Length]"
    Wend
    
    Call BASS_ChannelStop(chan)
    Call BASS_StreamFree(chan)
    
    wr.riffBlockSize = pos + 44 - 8
    wd.dataBlockSize = pos
    
    On Error Resume Next
        
    Put #ff, 5, wr.riffBlockSize
    Put #ff, 41, wd.dataBlockSize
    
    Form1.Titlelbl = "Encoding to WAV Successful"
    
No_Err:

    Erase buf()
        
    
    If chan Then Call BASS_ChannelStop(chan)
    If chan Then Call BASS_StreamFree(chan)
For u = 0 To 9
Call BASS_ChannelRemoveFX(chan, fxeq(u))
Next u
Call BASS_ChannelRemoveFX(chan, Form1.fxr)
    Close #ff
    Form1.Enabled = True
    Unload Form3
    

Err_Init:
If BASS_ErrorGetCode > 0 Then
    Open Form1.ap & "MP3PlayerX2.log" For Append As #1
    Print #1, BASS_GetErrorDescription(BASS_ErrorGetCode) & " " & BASS_ErrorGetCode & " occured when MPX2 tried to play/convert: " & originalfileortrack
    Close #1
    Form1.Enabled = 1
    Erase buf()
    Exit Sub
End If
End Sub
Public Sub Converter(origfile As String, toencoder As String, bitrate As String, outfile As String, iscd As Boolean)
Form1.vistimer.Enabled = 0
Form1.filetimer.Enabled = 0
BASS_ChannelStop (chan)
BASS_StreamFree (chan)
Form1.MediaType = "CONV"
Dim enchan As Long
For u = 0 To 9
Call BASS_ChannelRemoveFX(chan, fxeq(u))
Next u
Call BASS_ChannelRemoveFX(chan, Form1.fxr)
BASS_Free
BASS_Init 1, 44100, 0, Form1.hWnd, 0
Call BASS_SetConfig(BASS_CONFIG_FLOATDSP, 0)
BASS_Encode_Stop (chan)
Select Case iscd
Case False
    If UCase$(Right$(origfile, 3)) = "WMA" Then
    chan = BASS_WMA_StreamCreateFile(BASSFALSE, origfile, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    Else
    chan = BASS_StreamCreateFile(BASSFALSE, origfile, 0, 0, BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
    End If
Case True
    chan = BASS_CD_StreamCreate(0, Val(origfile), BASS_STREAM_DECODE Or BASS_STREAM_AUTOFREE)
End Select
        Form3.encProg.max = BASS_StreamGetLength(chan)

    Select Case toencoder
    Case "MP3C"
    enchan = BASS_Encode_Start(chan, Form1.ap & "lame.exe -b " & bitrate & " - " & Chr$(34) & outfile & Chr$(34), 0, 0, 0)
    Case "MP3V"
    enchan = BASS_Encode_Start(chan, Form1.ap & "lame.exe --abr " & bitrate & " - " & Chr$(34) & outfile & Chr$(34), 0, 0, 0)
    Case "OGGV"
    enchan = BASS_Encode_Start(chan, Form1.ap & "oggenc.exe -b " & bitrate & " -  -o " & Chr$(34) & outfile & Chr$(34), 0, 0, 0)
    Case "OGGC"
    enchan = BASS_Encode_Start(chan, Form1.ap & "oggenc.exe --managed -b " & bitrate & " -  -o " & Chr$(34) & outfile & Chr$(34), 0, 0, 0)
        End Select
        If Form3.Check1.value = 1 Then Call Form1.EQOn_Click
        If Form3.Check2.value = 1 Then Call Form1.rvbon_Click
          ReDim buf(262144) As Byte
    While BASS_ChannelIsActive(chan)
        ReDim Preserve buf(BASS_ChannelGetData(chan, buf(0), 262144) - 1) As Byte
        DoEvents
        Form3.encProg.value = BASS_ChannelGetPosition(chan)
        Form3.Label2 = BASS_ChannelGetPosition(chan) & "/" & BASS_StreamGetLength(chan) & " [Position/Length]"
        Form3.Caption = "Converting [" & Round((100 / BASS_StreamGetLength(chan)) * BASS_ChannelGetPosition(chan), 1) & "% ]"
    Wend
    BASS_Encode_Stop (chan)
    BASS_StreamFree (chan)
    BASS_Free
If BASS_ErrorGetCode > 0 Then
    Open Form1.ap & "MP3PlayerX2.log" For Append As #1
    Print #1, time$ & " " & Date & " " & BASS_GetErrorDescription(BASS_ErrorGetCode) & " " & BASS_ErrorGetCode & " occured when MPX2 tried to play/convert: " & origfile
    Close #1
    Form1.Enabled = 1
    Erase buf()
    Exit Sub
End If
    Form3.Hide

End Sub
