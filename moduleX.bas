Attribute VB_Name = "xx"
'inireader
'audio file id3 tag reader/writer
'randomizer

Option Explicit
Global u As Long
Global chan As Long

Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)

Global rtn As String
Global success As String
Global IniFileX As String
Public fxeq(9) As Long '10 dx8 eq bands

Private strData As String * 127
Private strPath As String
Private strArtist As String
Private strTitle As String
Private strAlbum As String
Private strYear As String
Private strComment As String
Private TagCreated As Boolean

Function GetValue(section$, key$) As String
If IniFileX = "" Then
Exit Function
End If
Dim KeyValue$
Dim characters As Long
    
KeyValue$ = String$(128, 0)
    
characters = GetPrivateProfileStringByKeyName(section$, key$, "", KeyValue$, 127, IniFileX)

If characters > 1 Then
   KeyValue$ = Left$(KeyValue$, characters)
End If
    
GetValue = KeyValue$

End Function
Function WriteValue(section$, key$, value$)
If IniFileX = "" Then Exit Function
WritePrivateProfileStringByKeyName section$, key$, value$, IniFileX

End Function

' Reads and Writes ID3v1 TAGS
Public Function FileExists(ByVal FileName As String) As Boolean
On Local Error Resume Next
FileExists = (Dir$(FileName) <> "")
End Function

Public Sub LoadMp3File(valPath As String)
    strPath = valPath
    If FileExists(valPath) = False Then Exit Sub
    Close
    Open strPath For Binary As #1

    Get #1, FileLen(valPath) - 127, strData
    Close #1
    
    TagCreated = False
    
    If TagExists = True Then
        strArtist = Mid(strData, 34, 30)
        strTitle = Mid(strData, 4, 30)
        strAlbum = Mid(strData, 64, 30)
        strComment = Mid(strData, 98, 30)
        strYear = Mid(strData, 94, 4)
    Else
        strArtist = ""
        strTitle = ""
        strAlbum = ""
        strYear = ""
        strComment = ""
    End If
    End Sub
                
Property Get Artist() As String
    
    Artist = RTrim(strArtist)
        
End Property

Property Get Title() As String

    Title = RTrim(strTitle)
        
End Property

Property Get Album() As String

    Album = RTrim(strAlbum)
            
End Property

Property Get Year() As String

    Year = RTrim(strYear)
    
End Property

Property Get Comment() As String
        
    Comment = RTrim(strComment)
          
End Property

Public Sub CloseMp3File()
        Dim ToBeWritten As String
    
    SetAttr strPath, vbNormal
    On Error Resume Next
    Open strPath For Binary As #1
    
    FileLen (strPath)
    ToBeWritten = "TAG"
    Put #1, FileLen(strPath) - 127, ToBeWritten
   '
    ToBeWritten = strTitle & String(30 - Len(strTitle), " ")
    Put #1, FileLen(strPath) - 124, ToBeWritten
    
    ToBeWritten = strArtist & String(30 - Len(strArtist), " ")
    Put #1, FileLen(strPath) - 94, ToBeWritten
    
    ToBeWritten = strAlbum & String(30 - Len(strAlbum), " ")
    Put #1, FileLen(strPath) - 64, ToBeWritten
    
    ToBeWritten = strYear & String(4 - Len(strYear), " ")
    Put #1, FileLen(strPath) - 34, ToBeWritten
    
    ToBeWritten = strComment & String(30 - Len(strComment), " ")
    Put #1, FileLen(strPath) - 30, ToBeWritten
    
    Close #1
    
    TagCreated = True
errfOUNDx:
End Sub

Public Function TagExists() As Boolean
    If InStr(strData, "TAG") >= 1 Or TagCreated = True Then
        If Right(strData, Len(strData) - 3) <> String(Len(strData) - 3, " ") Then
            TagExists = True
            Exit Function
        End If
    End If
    
    TagExists = False
    
End Function

Property Let Artist(valArtist As String)
     
    strArtist = valArtist
    
End Property

Property Let Title(valTitle As String)
     
    strTitle = valTitle
        
End Property

Property Let Album(valAlbum As String)
     
    strAlbum = valAlbum
    
End Property

Property Let Year(valYear As String)
     
    strYear = valYear
    
End Property

Property Let Comment(valComment As String)
     
    strComment = valComment
    
End Property

Public Sub RandList()
'randomises one playlist
Dim intRandomnumber As Long
Dim tmptext As String
Dim tmptext2 As String
Dim tmptext3 As String
Dim tmptext4 As String
Dim tmptext5 As String
Dim tmptext6 As String
Dim tmptext7 As String
Dim tmptext8 As String
Dim intNumberOfEntries As Long
Dim Y As Long
Dim X As Long
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

