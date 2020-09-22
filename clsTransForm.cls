VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long

Private Const RGN_DIFF = 4

Dim CurRgn As Long, TempRgn As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private objName As Object
Private lngHeight As Long, lngWidth As Long

Public Sub ShapeMe(Name As Object, Color As Long, Optional Load As Boolean = True, Optional FileName As String = vbNullString)

Dim X As Long, Y As Long
Dim booMiddleOfSet As Boolean
Dim colPoints As Collection
Set colPoints = New Collection
Dim z As Variant
Dim lngTransY As Long
Dim lngTransStartX As Long
Dim lngTransEndX As Long
Dim intStoreScaleMode As Integer
Dim lngHDC As Long
Set objName = Name
If TypeOf objName Is Form Then
    InitializeForFormOrPictureBox
    
    If FileName <> vbNullString And Load = True Then
        LoadRegionDataFromFile FileName
        
    Else
        lngHDC = objName.hDC
        booMiddleOfSet = False
        For Y = 0 To lngHeight
        
            lngTransY = Y
            For X = 0 To lngWidth
        
                If GetPixel(lngHDC, X, Y) = Color Then
                    If booMiddleOfSet = False Then
                        lngTransStartX = X
                        lngTransEndX = X
                        booMiddleOfSet = True
                    Else
                        lngTransEndX = X
                    End If
                Else
                    If booMiddleOfSet Then
                        colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
                        booMiddleOfSet = False
                    End If
                End If
            Next X
        Next Y
        CurRgn = CreateRectRgn(1, 1, lngWidth - 1, lngHeight - 1)
    End If
    
ElseIf TypeOf objName Is PictureBox Then
    InitializeForFormOrPictureBox
    If FileName <> vbNullString And Load = True Then
        LoadRegionDataFromFile FileName
    Else
        booMiddleOfSet = False
        For Y = 0 To lngHeight
            lngTransY = Y
            For X = 0 To lngWidth
                If objName.Point(X, Y) = Color Then
                    If booMiddleOfSet = False Then
                        lngTransStartX = X
                        lngTransEndX = X
                        booMiddleOfSet = True
                    Else
                        lngTransEndX = X
                    End If
                Else
                    If booMiddleOfSet Then
                        colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
                        booMiddleOfSet = False
                    End If
                End If
            Next X
        Next Y
        CurRgn = CreateRectRgn(0, 0, lngWidth, lngHeight)
    End If
    
ElseIf TypeOf objName Is CommandButton Or TypeOf objName Is OptionButton Or TypeOf objName Is CheckBox Then 'check to see if this is a button
    With objName
        intStoreScaleMode = .Parent.ScaleMode
        .Parent.ScaleMode = 3
        .Caption = ""
        .BackColor = Color
        .Refresh
        lngHDC = GetWindowDC(.hWnd)
        lngHeight = .Height
        lngWidth = .Width
    End With
    
    If FileName <> vbNullString And Load = True Then 'load region data from file
        LoadRegionDataFromFile FileName
        
    Else
        booMiddleOfSet = False
        For Y = 0 To lngHeight
            lngTransY = Y
            For X = 0 To lngWidth
                If GetPixel(lngHDC, X, Y) = Color Then
                    If booMiddleOfSet = False Then
                        lngTransStartX = X
                        lngTransEndX = X
                        booMiddleOfSet = True
                    Else
                        lngTransEndX = X
                    End If
                Else
                    If booMiddleOfSet Then
                        colPoints.Add Array(lngTransY, lngTransStartX, lngTransEndX)
                        booMiddleOfSet = False
                    End If
                End If
            Next X
        Next Y
        CurRgn = CreateRectRgn(2, 2, lngWidth - 2, lngHeight - 2)
    End If
    
Else
    Err.Raise vbObjectError + 512 + 2000, "TransForm", "Must pass in the name of a Form, PictureBox, CommandButton, CheckBox or OptionButton.  TransForm ShapeMe method failed."
    Exit Sub
End If

If FileName <> vbNullString And Load = True Then
Else
    For Each z In colPoints
        TempRgn = CreateRectRgn(z(1), z(0), z(2) + 1, z(0) + 1)
        CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF
        DeleteObject (TempRgn)
    Next
End If

If FileName <> vbNullString And Load = False Then
    SaveRegionDataToFile FileName
End If

SetWindowRgn objName.hWnd, CurRgn, True

ReleaseDC objName.hWnd, lngHDC


If TypeOf objName Is CommandButton Then
    objName.Parent.ScaleMode = intStoreScaleMode
End If

Set colPoints = Nothing

End Sub
Public Sub DragForm(hWnd As Long, intButton As Integer)

On Error Resume Next

If intButton = vbLeftButton Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End If

End Sub

Private Sub InitializeForFormOrPictureBox()
With objName
    .AutoRedraw = True
    .ScaleMode = 3
    .ClipControls = False
    lngHeight = .ScaleHeight
    lngWidth = .ScaleWidth
End With
End Sub

Private Sub SaveRegionDataToFile(ByVal sPath As String)
Dim iFile As Long
Dim nBytes As Long
Dim b() As Byte

On Error GoTo ErrorHandler
      
      nBytes = GetRegionData(CurRgn, 0, ByVal 0&)
      If nBytes > 0 Then
         ReDim b(0 To nBytes - 1) As Byte
         If nBytes = GetRegionData(CurRgn, nBytes, b(0)) Then
            On Error Resume Next
            Kill sPath
            On Error GoTo ErrorHandler
            iFile = FreeFile
            Open sPath For Binary Access Write Lock Read As #iFile
            Put #iFile, , b
            Close #iFile
         Else
            Err.Raise vbObjectError + 512 + 2001, "TransForm", "Unable to get region data in SaveRegionDataToFile"
         End If
      Else
         Err.Raise vbObjectError + 512 + 2002, "TransForm", "Unable to determine size of region in SaveRegionDataToFile"
      End If
   
   Exit Sub
   
ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If iFile > 0 Then
      Close #iFile
   End If
   Err.Raise lErr, "TransForm", sErr
   Exit Sub
   
End Sub

Private Sub LoadRegionDataFromFile(ByVal sFileName As String)
Dim iFile As Long
Dim b() As Byte
Dim dwCount As Long
On Error GoTo ErrorHandler

   iFile = FreeFile
   Open sFileName For Binary Access Read Lock Write As #iFile
   ReDim b(0 To LOF(iFile) - 1) As Byte
   Get #iFile, , b
   Close #iFile
   
   dwCount = UBound(b) - LBound(b) + 1
   CurRgn = ExtCreateRegion(ByVal 0&, dwCount, b(0))

   Exit Sub

ErrorHandler:
Dim lErr As Long, sErr As String
   lErr = Err.Number: sErr = Err.Description
   If iFile > 0 Then
      Close #iFile
   End If
   Err.Raise lErr, App.EXEName & ".cDIBSectionRegion", sErr
   Exit Sub
End Sub


