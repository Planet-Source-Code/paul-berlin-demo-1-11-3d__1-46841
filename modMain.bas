Attribute VB_Name = "ModMain"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Type tCamera
  XRot As Single 'rotation on X axis
  YRot As Single 'rotation on Y axis
  Zoom As Single 'Zoom level
  Perspective As Single 'Perspective
End Type

Type tVextex
  x As Single 'Location on X axis in 3d space
  y As Single 'Location on Y axis in 3d space
  z As Single 'Location on Z axis in 3d space
  iX As Single  'Location on X axis in 2d space
  iY As Single  'Location on Y axis in 2d space
  Camdist As Single 'The distance to the camera this object is
  tSel As Boolean 'Used when drawing
  bFlag As Byte 'Flag used when animating
End Type

Public Const MaxShapes As Byte = 10
Public Const TextCol As Long = vbWhite

'The camera (eye) is based with reference to point 0,0,0 in the 3d-space.
'The camera is always pointed towards 0,0,0.
Public Camera As tCamera

'This is the backbuffer in memory
'All objects are drawn onto this virtual image, then the backbuffer
'is shown on frmmain. This to remove flicker.
Public Backbuffer As New CDIB

'gfx holders
Public Sphere As New CDIB
Public Nebula As New CDIB

Public IsDebug As Boolean 'show debug info onto form, key D
Public DebugMode As Byte
Public Rot As Boolean 'auto rotate shape, key X
Public Dotted As Boolean 'if true, show dots instead of spheres
Public Fps As Boolean 'if true, show frame time

Public Stars(300) As tVextex 'just resize this array to remove/add stars, 1-based
Public Vertex() As tVextex 'Using an 1 based array for simpleness
Public Sound As New clsFmod 'The sound engine

Public Sub ProcessVertex(ByRef Vert() As tVextex, ByVal lWidth As Long, ByVal lHeight As Long, ByVal Theta As Single, ByVal Alt As Single, ByVal Size As Single, ByVal Perspective As Single)
  Dim x As Long
  
  'Define center coordinates of plotting area picture box
  Dim cX As Single
  Dim cY As Single

  'Define 3D viewpoint (eye) coordinates
  Dim vX As Single
  Dim vY As Single
  Dim vZ As Single

  'Define zenith distance angle.
  Dim Phi As Single
  Phi = 90 - Alt

  'Define sines and cosines of Theta and Phi
  Dim Sin_Theta As Single
  Dim Cos_Theta As Single
  Dim Sin_Phi   As Single
  Dim Cos_Phi   As Single
  
  'Set the center coordinate values of the plotting area
  cX = lWidth / 2
  cY = lHeight / 2

  'Compute the sines and cosines of the Theta and Phi angles.
  'This way they don't have to be computed more than once and
  'it speeds things up a tiny bit.
  Sin_Theta = Sine(Theta)
  Cos_Theta = Cosine(Theta)
  Sin_Phi = Sine(Phi)
  Cos_Phi = Cosine(Phi)

  'Compute viewpoint (eye) coordinates of point (X,Y,Z)
  For x = 1 To UBound(Vert)
    With Vert(x)
      vX = -.x * Sin_Theta + .y * Cos_Theta
      vY = -.x * Cos_Theta * Cos_Phi - .y * Sin_Theta * Cos_Phi + .z * Sin_Phi
      vZ = -.x * Cos_Theta * Sin_Phi - .y * Sin_Theta * Sin_Phi - .z * Cos_Phi + Perspective
      .Camdist = vZ
      
      'Convert to 2d coordinates
      .iX = cX + Size * vX / vZ
      .iY = cY - Size * vY / vZ
    End With
  Next x

End Sub

Public Function Sine(ByVal Degrees_Arg As Single) As Single
  Sine = Sin(Degrees_Arg * Atn(1) / 45)
End Function

Public Function Cosine(ByVal Degrees_Arg As Single) As Single
  Cosine = Cos(Degrees_Arg * Atn(1) / 45)
End Function

Public Sub DrawVertex(ByRef Vert() As tVextex, ByVal hdc As Long)
  'This sub will check each vertex distance to camera and draw the ones
  'further away first
  Dim x As Long, Drawn As Long
  Dim CurDepth As Single, oldSel As Single
  Dim oldDepth As Single
  
  
  For x = 1 To UBound(Vert)
    Vert(x).tSel = False
  Next x
  
  Do
    CurDepth = 32000
    oldDepth = 32000
    
    For x = 1 To UBound(Vert)
      If oldDepth - Vert(x).Camdist <= CurDepth And Not Vert(x).tSel Then
        oldSel = x
        CurDepth = oldDepth - Vert(x).Camdist
      End If
    Next x
    
    oldDepth = Vert(oldSel).Camdist
    Drawn = Drawn + 1
    Vert(oldSel).tSel = True
    
    'Draw
    If Dotted Then
      SetPixel Backbuffer.hdc, Vert(oldSel).iX, Vert(oldSel).iY, vbGreen
    Else
      Sphere.PaintToTrans hdc, Vert(oldSel).iX - (Sphere.Width / 2), Vert(oldSel).iY - (Sphere.Height / 2)
    End If
    
    If IsDebug Then 'Show debug info
      With Backbuffer
        Select Case DebugMode
          Case 1
            .PrintTxt "Draw order. VertexNumber:DrawOrder", 0, 0
            .PrintTxt oldSel & ":" & Drawn, Vert(oldSel).iX, Vert(oldSel).iY
          Case 2
            .PrintTxt "Distance to camera. VertexNumber:DistanceToCamera", 0, 0
            .PrintTxt oldSel & ":" & Vert(oldSel).Camdist, Vert(oldSel).iX, Vert(oldSel).iY
          Case 3
            .PrintTxt "Vertex number. VertexNumber (Total:" & UBound(Vert) & ")", 0, 0
            .PrintTxt oldSel, Vert(oldSel).iX, Vert(oldSel).iY
          Case 4
            .PrintTxt "X Rotation: " & Camera.XRot, 0, .Height - 48
            .PrintTxt "Y Rotation: " & Camera.YRot, 0, .Height - 32
            .PrintTxt "Zoom:" & Camera.Zoom, 0, .Height - 16
          Case Else
            .PrintTxt "Select debugmode with buttons 0-4.", 0, 0
        End Select
      End With
    End If
    
  Loop Until Drawn = UBound(Vert)

End Sub

Public Sub DoStars()
  Dim x As Long, y As Single
  
  'process 2d coordinates
  ProcessVertex Stars(), frmMain.ScaleWidth, frmMain.ScaleHeight, 0, -90, 2500, 5000
  
  For x = 1 To UBound(Stars)
    'calculate color
    y = Int((15000 - Stars(x).z) * (255 / 15000)) + 20
    If y > 255 Then y = 255
    'draw pixel
    SetPixel Backbuffer.hdc, Stars(x).iX, Stars(x).iY, "&H00" & CStr(Hex(y)) & CStr(Hex(y)) & CStr(Hex(y))
    'change location
    Stars(x).z = Stars(x).z - 50
    If Stars(x).z < -2500 Then
      Stars(x).z = 15000
    End If
  Next x
End Sub
