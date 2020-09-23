VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' DIB Helper Class
' (c) Damian, 2000
' Modified by Paul Berlin

Private MyhDC As Long, lpRGB As Long
Private bmH As BITMAPINFOHEADER
Private hBMO As Long, hDIB As Long
Private TxtRect As RECT

Sub Create(ByVal W As Long, ByVal H As Long)
'Creates or resizes the current DIB
    Class_Terminate
    With bmH
        .biSize = Len(bmH)
        .biWidth = W
        .biHeight = H
        .biPlanes = 1
        .biBitCount = 24
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
    End With
    MyhDC = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(MyhDC, bmH, 0, lpRGB, 0, 0)
    If hDIB Then
        hBMO = SelectObject(MyhDC, hDIB)
        Cls
    Else
        Err.Raise 825, , "DIB failed to create" & vbCrLf & "Please restart your system and try again"
    End If
End Sub

Sub Clone(ByVal pic As StdPicture, Optional ByVal FitSize As Boolean = True)
'Clones a Picture object, you may use LoadPicture to
'directly create the DIB without calling Create()

    Dim bmp As BITMAP
    
    GetObjectA pic.handle, Len(bmp), bmp
    If FitSize Then Create bmp.bmWidth, bmp.bmHeight
    If hDIB Then
        Dim hDCt As Long, hBMOt As Long
        hDCt = CreateCompatibleDC(MyhDC)
        hBMOt = SelectObject(hDCt, pic.handle)
        StretchBlt MyhDC, 0, 0, Width, Height, hDCt, 0, 0, bmp.bmWidth, bmp.bmHeight, vbSrcCopy
        SelectObject hDCt, hBMOt
        DeleteDC hDCt
    Else
        Err.Raise -1, , "DIB has to be created first"
    End If
End Sub

Sub CloneHDC(ByVal theHDC As Long, ByVal dwRop As Long, Optional x As Long = 0, Optional y As Long = 0)
'clones an entire hDC or a selected area
'useful for image cropping/copying
    BitBlt MyhDC, 0, 0, bmH.biWidth, bmH.biHeight, theHDC, x, y, dwRop
End Sub

Public Sub PaintTo(ByVal toDC As Long, ByVal atX As Long, ByVal atY As Long, ByVal dwRop As Long)
'paints the entire picture to selected hDC
    BitBlt toDC, atX, atY, bmH.biWidth, bmH.biHeight, MyhDC, 0, 0, dwRop
End Sub

Public Sub PaintToTrans(ByVal toDC As Long, ByVal atX As Long, ByVal atY As Long, Optional ByVal Ratio As Long = 100)
'paints the entire picture to selected hDC transparent, with ability to resize
    TransparentBlt toDC, atX, atY, bmH.biWidth * (Ratio * 0.01), bmH.biHeight * (Ratio * 0.01), MyhDC, 0, 0, bmH.biWidth, bmH.biHeight, &HFF00FF
End Sub

'nice addition that paints a resized image without changing the original picture
Sub PaintZoom(ByVal toDC As Long, ByVal atX As Long, ByVal atY As Long, ByVal Ratio As Long, ByVal dwRop As Long)
    StretchBlt toDC, atX, atY, bmH.biWidth * (Ratio * 0.01), bmH.biHeight * (Ratio * 0.01), MyhDC, 0, 0, bmH.biWidth, bmH.biHeight, dwRop
End Sub

Sub Cls()
    If hDIB Then ZeroMemory ByVal lpRGB, bmH.biSizeImage
End Sub

Private Sub Class_Terminate()
    If MyhDC Then
        If hBMO Then DeleteObject SelectObject(MyhDC, hBMO): hBMO = 0
        DeleteObject MyhDC: MyhDC = 0
    End If
End Sub

Property Get RGBSize() As Long
    RGBSize = bmH.biSizeImage
End Property

Property Get Width() As Long
    Width = bmH.biWidth
End Property

Property Get Height() As Long
    Height = bmH.biHeight
End Property

Public Property Get hdc() As Long
  hdc = MyhDC
End Property

Public Sub PrintTxt(ByVal theText As String, ByVal x As Long, ByVal y As Long)
  'It prints text at the desired location

  TxtRect.Left = x
  TxtRect.Right = Len(theText) * 16 + x
  TxtRect.Top = y
  TxtRect.Bottom = y + 50

  DrawText MyhDC, theText, Len(theText), TxtRect, DT_LEFT
End Sub

Public Sub ChangeColors(ByVal TextColor As Long, ByVal BgColor As Long, Optional ByVal BgTransparent As Boolean)
  'This changes the text color of the dib
  If BgTransparent Then
    SetBkMode MyhDC, 1
  Else
    SetBkMode MyhDC, 0
    SetBkColor MyhDC, BgColor
  End If
  SetTextColor MyhDC, TextColor
End Sub
