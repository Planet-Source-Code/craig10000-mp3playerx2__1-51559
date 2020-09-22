Attribute VB_Name = "BASSenc"

' BASSenc 2.0 Visual Basic API Header File
' Requires BASS.DLL & BASS.BAS - available @ www.un4seen.com

' See the BASSENC.CHM file for more complete documentation


' BASS_Encode_Start flags
Global Const BASS_ENCODE_NOHEAD = 1        'do NOT send a WAV header to the encoder
Global Const BASS_ENCODE_FP_8BIT = 2   'convert floating-point sample data to 8-bit integer
Global Const BASS_ENCODE_FP_16BIT = 4  'convert floating-point sample data to 16-bit integer
Global Const BASS_ENCODE_FP_24BIT = 6  'convert floating-point sample data to 24-bit integer


Declare Function BASS_Encode_Start Lib "bassenc.dll" (ByVal chan As Long, ByVal cmdline As String, ByVal flags As Long, ByVal proc As Long, ByVal user As Long) As Long
Declare Function BASS_Encode_IsActive Lib "bassenc.dll" (ByVal chan As Long) As Long
Declare Function BASS_Encode_Stop Lib "bassenc.dll" (ByVal chan As Long) As Long
Declare Function BASS_Encode_SetPaused Lib "bassenc.dll" (ByVal chan As Long, ByVal paused As Long) As Long


Sub ENCODEPROC(ByVal channel As Long, ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
    
    'CALLBACK FUNCTION !!!

    ' Encoding callback function.
    ' channel: The channel handle
    ' buffer : Buffer containing the encoded data
    ' length : Number of bytes
    ' user   : The 'user' parameter value given when calling BASS_EncodeStart
    
End Sub
