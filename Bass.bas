Attribute VB_Name = "modBass"
'BASS 2.0 Multimedia Library
'----------------------------
'(c) 1999-2003 Ian Luck.
'Please report bugs/suggestions/etc... to bass@un4seen.com

Global Const BASSTRUE As Long = 1
Global Const BASSFALSE As Long = 0

Global Const BASS_OK = 0
Global Const BASS_ERROR_MEM = 1
Global Const BASS_ERROR_FILEOPEN = 2
Global Const BASS_ERROR_DRIVER = 3
Global Const BASS_ERROR_BUFLOST = 4
Global Const BASS_ERROR_HANDLE = 5
Global Const BASS_ERROR_FORMAT = 6
Global Const BASS_ERROR_POSITION = 7
Global Const BASS_ERROR_INIT = 8
Global Const BASS_ERROR_START = 9
Global Const BASS_ERROR_ALREADY = 14
Global Const BASS_ERROR_NOPAUSE = 16
Global Const BASS_ERROR_NOTAUDIO = 17
Global Const BASS_ERROR_NOCHAN = 18
Global Const BASS_ERROR_ILLTYPE = 19
Global Const BASS_ERROR_ILLPARAM = 20
Global Const BASS_ERROR_NO3D = 21
Global Const BASS_ERROR_NOEAX = 22
Global Const BASS_ERROR_DEVICE = 23
Global Const BASS_ERROR_NOPLAY = 24
Global Const BASS_ERROR_FREQ = 25
Global Const BASS_ERROR_NOTFILE = 27
Global Const BASS_ERROR_NOHW = 29
Global Const BASS_ERROR_EMPTY = 31
Global Const BASS_ERROR_NONET = 32
Global Const BASS_ERROR_CREATE = 33
Global Const BASS_ERROR_NOFX = 34
Global Const BASS_ERROR_PLAYING = 35
Global Const BASS_ERROR_NOTAVAIL = 37
Global Const BASS_ERROR_DECODE = 38
Global Const BASS_ERROR_DX = 39
Global Const BASS_ERROR_TIMEOUT = 40
Global Const BASS_ERROR_FILEFORM = 41
Global Const BASS_ERROR_SPEAKER = 42
Global Const BASS_ERROR_UNKNOWN = -1

Global Const BASS_DEVICE_8BITS = 1     'use 8 bit resolution, else 16 bit
Global Const BASS_DEVICE_MONO = 2      'use mono, else stereo
Global Const BASS_DEVICE_3D = 4        'enable 3D functionality
Global Const BASS_DEVICE_LATENCY = 256 'calculate device latency (BASS_INFO struct)
Global Const BASS_DEVICE_SPEAKERS = 2048 'force enabling of speaker assignment
Global Const DSCAPS_CONTINUOUSRATE = 16
Global Const DSCAPS_EMULDRIVER = 32
Global Const DSCAPS_CERTIFIED = 64
Global Const DSCAPS_SECONDARYMONO = 256    ' mono
Global Const DSCAPS_SECONDARYSTEREO = 512  ' stereo
Global Const DSCAPS_SECONDARY8BIT = 1024   ' 8 bit
Global Const DSCAPS_SECONDARY16BIT = 2048  ' 16 bit
Global Const DSCCAPS_EMULDRIVER = DSCAPS_EMULDRIVER
Global Const DSCCAPS_CERTIFIED = DSCAPS_CERTIFIED
Global Const WAVE_FORMAT_1M08 = &H1          ' 11.025 kHz, Mono,   8-bit
Global Const WAVE_FORMAT_1S08 = &H2          ' 11.025 kHz, Stereo, 8-bit
Global Const WAVE_FORMAT_1M16 = &H4          ' 11.025 kHz, Mono,   16-bit
Global Const WAVE_FORMAT_1S16 = &H8          ' 11.025 kHz, Stereo, 16-bit
Global Const WAVE_FORMAT_2M08 = &H10         ' 22.05  kHz, Mono,   8-bit
Global Const WAVE_FORMAT_2S08 = &H20         ' 22.05  kHz, Stereo, 8-bit
Global Const WAVE_FORMAT_2M16 = &H40         ' 22.05  kHz, Mono,   16-bit
Global Const WAVE_FORMAT_2S16 = &H80         ' 22.05  kHz, Stereo, 16-bit
Global Const WAVE_FORMAT_4M08 = &H100        ' 44.1   kHz, Mono,   8-bit
Global Const WAVE_FORMAT_4S08 = &H200        ' 44.1   kHz, Stereo, 8-bit
Global Const WAVE_FORMAT_4M16 = &H400        ' 44.1   kHz, Mono,   16-bit
Global Const WAVE_FORMAT_4S16 = &H800        ' 44.1   kHz, Stereo, 16-bit
Global Const BASS_SAMPLE_8BITS = 1          ' 8 bit
Global Const BASS_SAMPLE_FLOAT = 256        ' 32-bit floating-point
Global Const BASS_SAMPLE_MONO = 2           ' mono, else stereo
Global Const BASS_SAMPLE_LOOP = 4           ' looped
Global Const BASS_SAMPLE_3D = 8             ' 3D functionality enabled
Global Const BASS_SAMPLE_SOFTWARE = 16      ' it's NOT using hardware mixing
Global Const BASS_SAMPLE_MUTEMAX = 32       ' muted at max distance (3D only)
Global Const BASS_SAMPLE_VAM = 64           ' uses the DX7 voice allocation & management
Global Const BASS_SAMPLE_FX = 128           ' old implementation of DX8 effects are enabled
Global Const BASS_SAMPLE_OVER_VOL = 65536   ' override lowest volume
Global Const BASS_SAMPLE_OVER_POS = 131072  ' override longest playing
Global Const BASS_SAMPLE_OVER_DIST = 196608 ' override furthest from listener (3D only)
Global Const BASS_MP3_SETPOS = 131072       ' enable pin-point seeking on the MP3/MP2/MP1
Global Const BASS_STREAM_AUTOFREE = 262144  ' automatically free the stream when it stop/ends
Global Const BASS_STREAM_RESTRATE = 524288  ' restrict the download rate of internet file streams
Global Const BASS_STREAM_BLOCK = 1048576    ' download/play internet file stream (MPx/OGG) in small blocks
Global Const BASS_STREAM_DECODE = &H200000  ' don't play the stream, only decode (BASS_ChannelGetData)
Global Const BASS_STREAM_META = &H400000    ' request metadata from a Shoutcast stream
Global Const BASS_STREAM_FILEPROC = &H800000 ' use a STREAMFILEPROC callback

Global Const BASS_UNICODE = &H80000000

Global Const BASS_RECORD_PAUSE = &H8000 ' start recording paused

Global Const BASS_SYNC_POS = 0
Global Const BASS_SYNC_MUSICPOS = 0
Global Const BASS_SYNC_MUSICINST = 1
Global Const BASS_SYNC_END = 2
Global Const BASS_SYNC_MUSICFX = 3
Global Const BASS_SYNC_META = 4
Global Const BASS_SYNC_SLIDE = 5
Global Const BASS_SYNC_STALL = 6
Global Const BASS_SYNC_DOWNLOAD = 7
Global Const BASS_SYNC_MESSAGE = &H20000000
Global Const BASS_SYNC_MIXTIME = &H40000000
Global Const BASS_SYNC_ONETIME = &H80000000

' BASS_ChannelIsActive return values
Global Const BASS_ACTIVE_STOPPED = 0
Global Const BASS_ACTIVE_PLAYING = 1
Global Const BASS_ACTIVE_STALLED = 2
Global Const BASS_ACTIVE_PAUSED = 3

' BASS_ChannelIsSliding return flags
Global Const BASS_SLIDE_FREQ = 1
Global Const BASS_SLIDE_VOL = 2
Global Const BASS_SLIDE_PAN = 4

' BASS_ChannelGetData flags
Global Const BASS_DATA_AVAILABLE = 0         ' query how much data is buffered
Global Const BASS_DATA_FFT512 = &H80000000   ' 512 sample FFT
Global Const BASS_DATA_FFT1024 = &H80000001  ' 1024 FFT
Global Const BASS_DATA_FFT2048 = &H80000002  ' 2048 FFT
Global Const BASS_DATA_FFT4096 = &H80000003  ' 4096 FFT
Global Const BASS_DATA_FFT512S = &H80000010  ' stereo 512 sample FFT
Global Const BASS_DATA_FFT1024S = &H80000011 ' stereo 1024 FFT
Global Const BASS_DATA_FFT2048S = &H80000012 ' stereo 2048 FFT
Global Const BASS_DATA_FFT4096S = &H80000013 ' stereo 4096 FFT
Global Const BASS_DATA_FFT_NOWINDOW = &H20   ' FFT flag: no Hanning window

' BASS_Set/GetConfig options
Global Const BASS_CONFIG_BUFFER = 0
Global Const BASS_CONFIG_UPDATEPERIOD = 1
Global Const BASS_CONFIG_MAXVOL = 3
Global Const BASS_CONFIG_GVOL_SAMPLE = 4
Global Const BASS_CONFIG_GVOL_STREAM = 5
Global Const BASS_CONFIG_GVOL_MUSIC = 6
Global Const BASS_CONFIG_CURVE_VOL = 7
Global Const BASS_CONFIG_CURVE_PAN = 8
Global Const BASS_CONFIG_FLOATDSP = 9
Global Const BASS_CONFIG_3DALGORITHM = 10
Global Const BASS_CONFIG_NET_TIMEOUT = 11
Global Const BASS_CONFIG_NET_BUFFER = 12

' BASS_StreamGetFilePosition modes
Global Const BASS_FILEPOS_DECODE = 0
Global Const BASS_FILEPOS_DOWNLOAD = 1
Global Const BASS_FILEPOS_END = 2

' STREAMFILEPROC actions
Global Const BASS_FILE_CLOSE = 0
Global Const BASS_FILE_READ = 1
Global Const BASS_FILE_QUERY = 2
Global Const BASS_FILE_LEN = 3

Global Const BASS_STREAMPROC_END = &H80000000 ' end of user stream flag

'**************************************************************
'* DirectSound interfaces (for use with BASS_GetDSoundObject) *
'**************************************************************
Global Const BASS_OBJECT_DS = 1                     ' DirectSound
Global Const BASS_OBJECT_DS3DL = 2                  'IDirectSound3DListener

'******************************
'* DX7 voice allocation flags *
'******************************
' Play the sample in hardware. If no hardware voices are available then
' the "play" call will fail
Global Const BASS_VAM_HARDWARE = 1
' Play the sample in software (ie. non-accelerated). No other VAM flags
'may be used together with this flag.
Global Const BASS_VAM_SOFTWARE = 2
Global Const BASS_VAM_TERM_TIME = 4
Global Const BASS_VAM_TERM_DIST = 8
Global Const BASS_VAM_TERM_PRIO = 16

Global Const BASS_3DALG_DEFAULT = 0
Global Const BASS_3DALG_OFF = 1
Global Const BASS_3DALG_FULL = 2
Global Const BASS_3DALG_LIGHT = 3

Type BASS_INFO
    size As Long          ' size of this struct (set this before calling the function)
    flags As Long         ' device capabilities (DSCAPS_xxx flags)
    hwsize As Long        ' size of total device hardware memory
    hwfree As Long        ' size of free device hardware memory
    freesam As Long       ' number of free sample slots in the hardware
    free3d As Long        ' number of free 3D sample slots in the hardware
    minrate As Long       ' min sample rate supported by the hardware
    maxrate As Long       ' max sample rate supported by the hardware
    eax As Long           ' device supports EAX? (always BASSFALSE if BASS_DEVICE_3D was not used)
    minbuf As Long        ' recommended minimum buffer length in ms (requires BASS_DEVICE_LATENCY)
    dsver As Long         ' DirectSound version
    latency As Long       ' delay (in ms) before start of playback (requires BASS_DEVICE_LATENCY)
    initflags As Long     ' "flags" parameter of BASS_Init call
    speakers As Long      ' number of speakers available
    driver As Long        ' driver
End Type

Type BASS_RECORDINFO
    size As Long          ' size of this struct (set this before calling the function)
    flags As Long         ' device capabilities (DSCCAPS_xxx flags)
    formats As Long       ' supported standard formats (WAVE_FORMAT_xxx flags)
    inputs As Long        ' number of inputs
    singlein As Long      ' BASSTRUE = only 1 input can be set at a time
    driver As Long        ' driver
End Type

Type BASS_SAMPLE
    freq As Long          ' default playback rate
    Volume As Long        ' default volume (0-100)
    pan As Long           ' default pan (-100=left, 0=middle, 100=right)
    flags As Long         ' BASS_SAMPLE_xxx flags
    length As Long        ' length (in samples, not bytes)
    max As Long           ' maximum simultaneous playbacks
    ' The following are the sample's default 3D attributes (if the sample
    ' is 3D, BASS_SAMPLE_3D is in flags) see BASS_ChannelSet3DAttributes
    mode3d As Long        ' BASS_3DMODE_xxx mode
    mindist As Single     ' minimum distance
    MAXDIST As Single     ' maximum distance
    iangle As Long        ' angle of inside projection cone
    oangle As Long        ' angle of outside projection cone
    outvol As Long        ' delta-volume outside the projection cone
    ' The following are the defaults used if the sample uses the DirectX 7
    ' voice allocation/management features.
    vam As Long           ' voice allocation/management flags (BASS_VAM_xxx)
    priority As Long      ' priority (0=lowest, &Hffffffff=highest)
End Type

Type BASS_CHANNELINFO
        freq As Long          ' default playback rate
        chans As Long         ' channels
        flags As Long         ' BASS_SAMPLE/STREAM/MUSIC/SPEAKER flags
        ctype As Long         ' type of channel
End Type

' BASS_CHANNELINFO types
Global Const BASS_CTYPE_SAMPLE = 1
Global Const BASS_CTYPE_RECORD = 2
Global Const BASS_CTYPE_STREAM = &H10000
Global Const BASS_CTYPE_STREAM_WAV = &H10001
Global Const BASS_CTYPE_STREAM_OGG = &H10002
Global Const BASS_CTYPE_STREAM_MP1 = &H10003
Global Const BASS_CTYPE_STREAM_MP2 = &H10004
Global Const BASS_CTYPE_STREAM_MP3 = &H10005
Global Const BASS_CTYPE_MUSIC_MOD = &H20000
Global Const BASS_CTYPE_MUSIC_MTM = &H20001
Global Const BASS_CTYPE_MUSIC_S3M = &H20002
Global Const BASS_CTYPE_MUSIC_XM = &H20003
Global Const BASS_CTYPE_MUSIC_IT = &H20004
Global Const BASS_CTYPE_MUSIC_MO3 = &H100    ' mo3 flag

Global Const BASS_FX_PARAMEQ = 7        ' GUID_DSFX_STANDARD_PARAMEQ
Global Const BASS_FX_REVERB = 8         ' GUID_DSFX_WAVES_REVERB
Type BASS_FXPARAMEQ             ' DSFXParamEq
    fCenter As Single
    fBandwidth As Single
    fGain As Single
End Type

Type BASS_FXREVERB              ' DSFXWavesReverb
    fInGain As Single                ' [-96.0,0.0]            default: 0.0 dB
    fReverbMix As Single             ' [-96.0,0.0]            default: 0.0 db
    fReverbTime As Single            ' [0.001,3000.0]         default: 1000.0 ms
    fHighFreqRTRatio As Single       ' [0.001,0.999]          default: 0.001
End Type

Global Const BASS_FX_PHASE_NEG_180 = 0
Global Const BASS_FX_PHASE_NEG_90 = 1
Global Const BASS_FX_PHASE_ZERO = 2
Global Const BASS_FX_PHASE_90 = 3
Global Const BASS_FX_PHASE_180 = 4

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Declare Function BASS_SetConfig Lib "bass.dll" (ByVal opt As Long, ByVal value As Long) As Long
Declare Function BASS_GetConfig Lib "bass.dll" (ByVal opt As Long) As Long
Declare Function BASS_GetVersion Lib "bass.dll" () As Long
Declare Function BASS_GetDeviceDescription Lib "bass.dll" (ByVal device As Long) As Long
Declare Function BASS_ErrorGetCode Lib "bass.dll" () As Long
Declare Function BASS_Init Lib "bass.dll" (ByVal device As Long, ByVal freq As Long, ByVal flags As Long, ByVal win As Long, ByVal clsid As Long) As Long
Declare Function BASS_SetDevice Lib "bass.dll" (ByVal device As Long) As Long
Declare Function BASS_GetDevice Lib "bass.dll" () As Long
Declare Function BASS_Free Lib "bass.dll" () As Long
Declare Function BASS_GetDSoundObject Lib "bass.dll" (ByVal object As Long) As Long
Declare Function BASS_GetInfo Lib "bass.dll" (ByRef info As BASS_INFO) As Long
Declare Function BASS_Update Lib "bass.dll" () As Long
Declare Function BASS_GetCPU Lib "bass.dll" () As Single
Declare Function BASS_Start Lib "bass.dll" () As Long
Declare Function BASS_Stop Lib "bass.dll" () As Long
Declare Function BASS_Pause Lib "bass.dll" () As Long
Declare Function BASS_SetVolume Lib "bass.dll" (ByVal Volume As Long) As Long
Declare Function BASS_GetVolume Lib "bass.dll" () As Long

Declare Function BASS_StreamCreate Lib "bass.dll" (ByVal freq As Long, ByVal chans As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_StreamCreateFile Lib "bass.dll" (ByVal mem As Long, ByVal f As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
Declare Function BASS_StreamCreateURL Lib "bass.dll" (ByVal url As String, ByVal offset As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_StreamCreateFileUser Lib "bass.dll" (ByVal buffered As Long, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Sub BASS_StreamFree Lib "bass.dll" (ByVal handle As Long)
Declare Function BASS_StreamGetLength Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_StreamGetTags Lib "bass.dll" (ByVal handle As Long, ByVal tags As Long) As Long
Declare Function BASS_StreamPreBuf Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_StreamPlay Lib "bass.dll" (ByVal handle As Long, ByVal flush As Long, ByVal flags As Long) As Long
Declare Function BASS_StreamGetFilePosition Lib "bass.dll" (ByVal handle As Long, ByVal mode As Long) As Long

Private Declare Function BASS_ChannelBytes2Seconds64 Lib "bass.dll" Alias "BASS_ChannelBytes2Seconds" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long) As Single
Declare Function BASS_ChannelSeconds2Bytes Lib "bass.dll" (ByVal handle As Long, ByVal pos As Single) As Long
Declare Function BASS_ChannelGetDevice Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelIsActive Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelGetInfo Lib "bass.dll" (ByVal handle As Long, ByRef info As BASS_CHANNELINFO) As Long
Declare Function BASS_ChannelStop Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelPause Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelResume Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelSetAttributes Lib "bass.dll" (ByVal handle As Long, ByVal freq As Long, ByVal Volume As Long, ByVal pan As Long) As Long
Declare Function BASS_ChannelGetAttributes Lib "bass.dll" (ByVal handle As Long, ByRef freq As Long, ByRef Volume As Long, ByRef pan As Long) As Long
Declare Function BASS_ChannelSlideAttributes Lib "bass.dll" (ByVal handle As Long, ByVal freq As Long, ByVal Volume As Long, ByVal pan As Long, ByVal time As Long) As Long
Declare Function BASS_ChannelIsSliding Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelSet3DAttributes Lib "bass.dll" (ByVal handle As Long, ByVal mode As Long, ByVal min As Single, ByVal max As Single, ByVal iangle As Long, ByVal oangle As Long, ByVal outvol As Long) As Long
Declare Function BASS_ChannelGet3DAttributes Lib "bass.dll" (ByVal handle As Long, ByRef mode As Long, ByRef min As Single, ByRef max As Single, ByRef iangle As Long, ByRef oangle As Long, ByRef outvol As Long) As Long
Declare Function BASS_ChannelSet3DPosition Lib "bass.dll" (ByVal handle As Long, ByRef pos As Any, ByRef orient As Any, ByRef vel As Any) As Long
Declare Function BASS_ChannelGet3DPosition Lib "bass.dll" (ByVal handle As Long, ByRef pos As Any, ByRef orient As Any, ByRef vel As Any) As Long
Private Declare Function BASS_ChannelSetPosition64 Lib "bass.dll" Alias "BASS_ChannelSetPosition" (ByVal handle As Long, ByVal pos As Long, ByVal poshigh As Long) As Long
Declare Function BASS_ChannelGetPosition Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelGetLevel Lib "bass.dll" (ByVal handle As Long) As Long
Declare Function BASS_ChannelGetData Lib "bass.dll" (ByVal handle As Long, ByRef buffer As Any, ByVal length As Long) As Long
Private Declare Function BASS_ChannelSetSync64 Lib "bass.dll" Alias "BASS_ChannelSetSync" (ByVal handle As Long, ByVal atype As Long, ByVal param As Long, ByVal paramhigh As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_ChannelRemoveSync Lib "bass.dll" (ByVal handle As Long, ByVal sync As Long) As Long
Declare Function BASS_ChannelSetDSP Lib "bass.dll" (ByVal handle As Long, ByVal proc As Long, ByVal user As Long, ByVal priority As Long) As Long
Declare Function BASS_ChannelRemoveDSP Lib "bass.dll" (ByVal handle As Long, ByVal dsp As Long) As Long
Declare Function BASS_ChannelSetEAXMix Lib "bass.dll" (ByVal handle As Long, ByVal mix As Single) As Long
Declare Function BASS_ChannelGetEAXMix Lib "bass.dll" (ByVal handle As Long, ByRef mix As Single) As Long
Declare Function BASS_ChannelSetLink Lib "bass.dll" (ByVal handle As Long, ByVal chan As Long) As Long
Declare Function BASS_ChannelRemoveLink Lib "bass.dll" (ByVal handle As Long, ByVal chan As Long) As Long
Declare Function BASS_ChannelSetFX Lib "bass.dll" (ByVal handle As Long, ByVal atype As Long) As Long
Declare Function BASS_ChannelRemoveFX Lib "bass.dll" (ByVal handle As Long, ByVal fx As Long) As Long
Declare Function BASS_FXSetParameters Lib "bass.dll" (ByVal handle As Long, ByRef par As Any) As Long
Declare Function BASS_FXGetParameters Lib "bass.dll" (ByVal handle As Long, ByRef par As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Function BASS_ChannelBytes2Seconds(ByVal handle As Long, ByVal pos As Long) As Single
BASS_ChannelBytes2Seconds = BASS_ChannelBytes2Seconds64(handle, pos, 0)
End Function
Function BASS_ChannelSetPosition(ByVal handle As Long, ByVal pos As Long) As Long
BASS_ChannelSetPosition = BASS_ChannelSetPosition64(handle, pos, 0)
End Function
Function BASS_ChannelSetSync(ByVal handle As Long, ByVal atype As Long, ByVal param As Long, ByVal proc As Long, ByVal user As Long) As Long
BASS_ChannelSetSync = BASS_ChannelSetSync64(handle, atype, param, 0, proc, user)
End Function
Function STREAMPROC(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
End Function
Function STREAMFILEPROC(ByVal action As Long, ByVal param1 As Long, ByVal param2 As Long, ByVal user As Long) As Long
End Function
Sub DOWNLOADPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
End Sub
Sub SYNCPROC(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
End Sub
Sub DSPPROC(ByVal handle As Long, ByVal channel As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
End Sub
Function RECORDPROC(ByVal handle As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long) As Long
End Function
Function BASS_GetDeviceDescriptionString(ByVal device As Long) As String
Dim pstring As Long
Dim sstring As String
On Error Resume Next
pstring = BASS_GetDeviceDescription(device)
If pstring Then
    sstring = VBStrFromAnsiPtr(pstring)
End If
BASS_GetDeviceDescriptionString = sstring
End Function
Function BASS_GetStringVersion() As String
BASS_GetStringVersion = Trim(Str(LoWord(BASS_GetVersion))) & "." & Trim(Str(HiWord(BASS_GetVersion)))
End Function
Public Function HiWord(lParam As Long) As Long
HiWord = lParam \ &H10000 And &HFFFF&
End Function
Public Function LoWord(lParam As Long) As Long
LoWord = lParam And &HFFFF&
End Function
Function MakeLong(LoWord As Long, HiWord As Long) As Long
MakeLong = (LoWord And &HFFFF&) Or (HiWord * &H10000)
End Function
Public Function VBStrFromAnsiPtr(ByVal lpStr As Long) As String
Dim bStr() As Byte
Dim cChars As Long
On Error Resume Next
cChars = lstrlen(lpStr)
ReDim bStr(0 To cChars - 1) As Byte
Call CopyMemory(bStr(0), ByVal lpStr, cChars)
VBStrFromAnsiPtr = StrConv(bStr, vbUnicode)
End Function
Public Function BASS_GetErrorDescription(ErrorCode As Long) As String
Select Case ErrorCode
    Case BASS_OK
        BASS_GetErrorDescription = "All is OK"
    Case BASS_ERROR_MEM
        BASS_GetErrorDescription = "Memory Error"
    Case BASS_ERROR_FILEOPEN
        BASS_GetErrorDescription = "Can't Open the File"
    Case BASS_ERROR_DRIVER
        BASS_GetErrorDescription = "Can't Find a Free Sound Driver"
    Case BASS_ERROR_BUFLOST
        BASS_GetErrorDescription = "The Sample Buffer Was Lost - Please Report This!"
    Case BASS_ERROR_HANDLE
        BASS_GetErrorDescription = "Invalid Handle"
    Case BASS_ERROR_FORMAT
        BASS_GetErrorDescription = "Unsupported Format"
    Case BASS_ERROR_POSITION
        BASS_GetErrorDescription = "Invalid Playback Position"
    Case BASS_ERROR_INIT
        BASS_GetErrorDescription = "BASS_Init Has Not Been Successfully Called"
    Case BASS_ERROR_START
        BASS_GetErrorDescription = "BASS_Start Has Not Been Successfully Called"
    Case BASS_ERROR_INITCD
        BASS_GetErrorDescription = "Can't Initialize CD"
    Case BASS_ERROR_CDINIT
        BASS_GetErrorDescription = "BASS_CDInit Has Not Been Successfully Called"
    Case BASS_ERROR_NOCD
        BASS_GetErrorDescription = "No CD in drive"
    Case BASS_ERROR_CDTRACK
        BASS_GetErrorDescription = "Can't Play the Selected CD Track"
    Case BASS_ERROR_ALREADY
        BASS_GetErrorDescription = "Already Initialized"
    Case BASS_ERROR_CDVOL
        BASS_GetErrorDescription = "CD Has No Volume Control"
    Case BASS_ERROR_NOPAUSE
        BASS_GetErrorDescription = "Not Paused"
    Case BASS_ERROR_NOTAUDIO
        BASS_GetErrorDescription = "Not An Audio Track"
    Case BASS_ERROR_NOCHAN
        BASS_GetErrorDescription = "Can't Get a Free Channel"
    Case BASS_ERROR_ILLTYPE
        BASS_GetErrorDescription = "An Illegal Type Was Specified"
    Case BASS_ERROR_ILLPARAM
        BASS_GetErrorDescription = "An Illegal Parameter Was Specified"
    Case BASS_ERROR_NO3D
        BASS_GetErrorDescription = "No 3D Support"
    Case BASS_ERROR_NOEAX
        BASS_GetErrorDescription = "No EAX Support"
    Case BASS_ERROR_DEVICE
        BASS_GetErrorDescription = "Illegal Device Number"
    Case BASS_ERROR_NOPLAY
        BASS_GetErrorDescription = "Not Playing"
    Case BASS_ERROR_FREQ
        BASS_GetErrorDescription = "Illegal Sample Rate"
    Case BASS_ERROR_NOTFILE
        BASS_GetErrorDescription = "The Stream is Not a File Stream (WAV/MP3)"
    Case BASS_ERROR_NOHW
        BASS_GetErrorDescription = "No Hardware Voices Available"
    Case BASS_ERROR_EMPTY
        BASS_GetErrorDescription = "The MOD music has no sequence data"
    Case BASS_ERROR_NONET
        BASS_GetErrorDescription = "No Internet connection could be opened"
    Case BASS_ERROR_CREATE
        BASS_GetErrorDescription = "Couldn't create the file"
    Case BASS_ERROR_NOFX
        BASS_GetErrorDescription = "Effects are not enabled"
    Case BASS_ERROR_PLAYING
        BASS_GetErrorDescription = "The channel is playing"
    Case BASS_ERROR_NOTAVAIL
        BASS_GetErrorDescription = "The requested data is not available"
    Case BASS_ERROR_DECODE
        BASS_GetErrorDescription = "The channel is a 'decoding channel'"
    Case BASS_ERROR_UNKNOWN
        BASS_GetErrorDescription = "Some Other Mystery Error"
End Select
End Function
