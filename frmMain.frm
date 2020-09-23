VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Demo"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      Top             =   0
      Width           =   135
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShpTmr As New clsStopwatch

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then Camera.YRot = Camera.YRot + 2
  If KeyCode = vbKeyUp Then Camera.YRot = Camera.YRot - 2
  If KeyCode = vbKeyLeft Then Camera.XRot = Camera.XRot + 2
  If KeyCode = vbKeyRight Then Camera.XRot = Camera.XRot - 2
  If KeyCode = vbKeyA Then Camera.Zoom = Camera.Zoom + 100
  If KeyCode = vbKeyZ Then Camera.Zoom = Camera.Zoom - 100
  If KeyCode = vbKeyD Then IsDebug = Not IsDebug: Fps = IsDebug
  If KeyCode = vbKeyV Then Dotted = Not Dotted
  If KeyCode = vbKeyN Then ShpTmr.SetStartTime ShpTmr.GetTime - 21000
  If KeyCode = vbKeyF Then Fps = Not Fps
  If KeyCode = vbKey1 And IsDebug Then DebugMode = 1
  If KeyCode = vbKey2 And IsDebug Then DebugMode = 2
  If KeyCode = vbKey3 And IsDebug Then DebugMode = 3
  If KeyCode = vbKey4 And IsDebug Then DebugMode = 4
  If KeyCode = vbKey0 And IsDebug Then DebugMode = 0
  If KeyCode = vbKeyEscape Then Unload Me
  If KeyCode = vbKeyX Then Rot = Not Rot
End Sub

Private Sub Form_Load()
  'Dim Sw As New clsStopwatch
  Dim ns As Long, vs As Long 'fps calc
  Dim mStart As Long, mTime(9) As Long 'fps calc
  Dim CurShape As Byte 'holds current shape
  Dim x As Long
  
  Me.Caption = Me.Caption + " v1." & App.Revision & " [By Paul Berlin 2002]"
  Me.Show
  
  'Init graphics, sound, varaibles, 3d-shapes
  Backbuffer.Create Me.ScaleWidth, Me.ScaleHeight
  Backbuffer.ChangeColors TextCol, vbRed, False
  Sphere.Clone LoadPicture(App.Path & "\sphere.bmp")
  Nebula.Clone LoadPicture(App.Path & "\nebula.bmp")
  ReDim Vertex(0) As tVextex
  
  For x = 1 To UBound(Stars)
    With Stars(x)
      .x = (Rnd * 2000) - 1000
      .y = (Rnd * 2000) - 1000
      .z = (Rnd * 15000)
    End With
  Next x
  
  CurShape = 0
  Rot = True
  Camera.Perspective = 10000
  
  LoadShape CurShape
  
  With Sound
    .InitBuffer 100
    .Init 44100, 16
    .MusicPlay App.Path & "\music.mod"
  End With
  
  'main loop
  ShpTmr.StartWatch 'Timer for changing shapes
  Do
    'Sw.StartWatch 'timer for measuring drawtime
    
    DoEvents
    Nebula.PaintTo Backbuffer.hdc, (Backbuffer.Width / 2) - (Nebula.Width / 2), (Backbuffer.Height / 2) - (Nebula.Height / 2), vbSrcCopy
    DoStars 'Draw stars
    AnimateShape CurShape 'Animate 3d-shapes
    ProcessVertex Vertex(), Me.ScaleWidth, Me.ScaleHeight, Camera.XRot, Camera.YRot, Camera.Zoom, Camera.Perspective 'Convert XYZ coordinates to XY coordinates
    DrawVertex Vertex(), Backbuffer.hdc 'Draw 3d-shape
    
    Backbuffer.PaintTo Me.hdc, 0, 0, vbSrcCopy 'copy backbuffer to form
    Backbuffer.Cls 'clear backbuffer
    
       
    If ShpTmr.GetTime >= 20000 Then 'change shape after 20 sec
      If Camera.Zoom > 0 Then
        Camera.Zoom = Camera.Zoom - 50 'but first zoom out
      Else
        CurShape = CurShape + 1
        If CurShape > MaxShapes Then CurShape = 0
        LoadShape CurShape
        ShpTmr.StartWatch
      End If
    Else
      With Camera
        If .Zoom < 2500 Then 'make sure to zoom out after new loaded shape
          .Zoom = .Zoom + 50
        End If
        If Rot Then 'Auto rotate
          .XRot = .XRot + 1.25
          .YRot = .YRot + 0.75
        End If
        If .XRot > 360 Then .XRot = .XRot - 360
        If .XRot < -360 Then .XRot = .XRot + 360
        If .YRot > 360 Then .YRot = .YRot - 360
        If .YRot < -360 Then .YRot = .YRot + 360
      End With
    End If
    
    
    'calculate fps
    If Fps Then
      vs = 0
      For ns = 0 To 8
        mTime(ns) = mTime(ns + 1)
        vs = vs + mTime(ns)
      Next ns
      mTime(ns) = timeGetTime - mStart
      Backbuffer.PrintTxt CStr(10000 \ (vs + mTime(ns))), Backbuffer.Width - (Len(CStr(10000 \ (vs + mTime(ns)))) * 8), 0
    End If
    mStart = timeGetTime
  Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set Sound = Nothing
  Set Backbuffer = Nothing
  Set Sphere = Nothing
  End
End Sub

Public Sub AddVertex(ByVal x As Single, ByVal y As Single, ByVal z As Single)
  ReDim Preserve Vertex(UBound(Vertex) + 1) As tVextex
  With Vertex(UBound(Vertex))
    .x = x
    .y = y
    .z = z
  End With
End Sub

Public Sub LoadShape(Shp As Byte)
  'This creates an 3d shape in memory
  Dim x As Long, y As Long
  ReDim Vertex(0) As tVextex
  
  Select Case Shp
  Case 0 'cube
    For x = -250 To 250 Step 50
      AddVertex 250, x, -250
      AddVertex -250, x, 250
      AddVertex -250, x, -250
      AddVertex 250, x, 250
    Next x
    For x = -200 To 200 Step 50
      AddVertex x, -250, -250
      AddVertex x, -250, 250
      AddVertex x, 250, -250
      AddVertex x, 250, 250
    Next x
    For x = -200 To 200 Step 50
      AddVertex 250, -250, x
      AddVertex -250, -250, x
      AddVertex 250, 250, x
      AddVertex -250, 250, x
    Next x
  Case 1
    For x = -250 To 250 Step 50
      AddVertex x, -250, -250
      AddVertex x, -250, 250
      AddVertex x, 250, -250
      AddVertex x, 250, 250
    Next x
    For x = -200 To 200 Step 50
      AddVertex 250, -250, x
      AddVertex -250, -250, x
      AddVertex 250, 250, x
      AddVertex -250, 250, x
    Next x
    For x = 0 To 360 Step 10
      AddVertex 250 * Cosine(x), 0, 250 * Sine(x)
    Next x
      AddVertex 0, 0, 0
      AddVertex 0, -50, 0
      AddVertex 0, 50, 0
      AddVertex 0, 100, 0
      AddVertex 0, -100, 0
  Case 2
    For x = -200 To 200 Step 50
      AddVertex x, -200, 0
      AddVertex x, 200, 0
    Next x
    For x = -150 To 150 Step 50
      AddVertex -200, x, 0
      AddVertex 200, x, 0
    Next x
    For x = 0 To 400 Step 25
      AddVertex 200 - (x / 2), 200 - (x / 2), x
    Next x
    For x = 0 To 350 Step 50
      AddVertex -200 + (x / 2), 200 - (x / 2), x
      AddVertex 200 - (x / 2), -200 + (x / 2), x
      AddVertex -200 + (x / 2), -200 + (x / 2), x
    Next x
  Case 3
    For x = 0 To 360 Step 5
      AddVertex 0, 500 * Cosine(x), 500 * Sine(x)
    Next x
    For x = -100 To 100 Step 50
      AddVertex 0, x, 0
      AddVertex x, 0, 0
      AddVertex 0, 0, x
    Next x
  Case 4
    For x = 0 To 360 Step 20
      AddVertex Cosine(x), 250 * Cosine(x), 200 * Sine(x)
      AddVertex 100 * Cosine(x), 250 * Cosine(x), 200 * Sine(x)
      AddVertex 200 * Cosine(x), 250 * Cosine(x), 200 * Sine(x)
      AddVertex 300 * Cosine(x), 250 * Cosine(x), 200 * Sine(x)
      AddVertex 400 * Cosine(x), 250 * Cosine(x), 200 * Sine(x)
    Next x
    AddVertex 0, 0, 0
    AddVertex 50, 0, 0
    AddVertex -50, 0, 0
    AddVertex 100, 0, 0
    AddVertex -100, 0, 0
    AddVertex -150, 0, 0
    AddVertex 150, 0, 0
  Case 5
    For x = -200 To 200 Step 100
      For y = -200 To 200 Step 100
        AddVertex x, y, 0
        AddVertex x, y, -200
        AddVertex x, y, 200
        AddVertex x, y, 100
        AddVertex x, y, -100
      Next y
    Next x
    For x = 0 To 360 Step 10
      AddVertex 0, 500 * Cosine(x), 500 * Sine(x)
    Next x
  Case 6
    For y = 50 To 500 Step 50
      For x = 0 To 360 Step 30
        AddVertex 50 * Sine(y), y * Cosine(x), y * Sine(x)
      Next x
    Next y
  Case 7
    For y = 0 To 200 Step 25
      For x = 0 To 360 Step 20
        AddVertex 500 * Sine(y), y * Cosine(x), y * Sine(x)
      Next x
    Next y
  Case 8
    For x = -600 To 600 Step 40
      AddVertex 50 * Sine(x), x, 0
    Next x
    For x = 20 To 100 Step 20
      AddVertex x - 40, 600 + x, 0
      AddVertex -x - 40, 600 + x, 0
    Next x
    For x = -80 To 80 Step 20
      AddVertex x - 40, 700, 0
    Next x
    For x = 20 To 100 Step 20
      AddVertex -x + 40, -600 - x, 0
      AddVertex x + 40, -600 - x, 0
    Next x
    For x = -80 To 80 Step 20
      AddVertex x + 40, -700, 0
    Next x
  Case 9
    For x = 0 To 360 Step 40
      AddVertex -500, 100 * Cosine(x), 100 * Sine(x)
      AddVertex -450, 100 * Cosine(x), 100 * Sine(x)
      AddVertex -550, 100 * Cosine(x), 100 * Sine(x)
      AddVertex -600, 100 * Cosine(x), 100 * Sine(x)
      AddVertex -400, 100 * Cosine(x), 100 * Sine(x)
      AddVertex -650, 100 * Cosine(x), 100 * Sine(x)
      AddVertex -700, 100 * Cosine(x), 100 * Sine(x)
    Next x
    For x = 0 To 360 Step 40
      AddVertex 700, 100 * Cosine(x), 100 * Sine(x)
    Next x
    For x = 0 To 360 Step 10
      AddVertex 700, 400 * Cosine(x), 400 * Sine(x)
    Next x
    AddVertex -700, -25, 25
    AddVertex -700, 25, 25
    AddVertex -700, -25, -25
    AddVertex -700, 25, -25
    AddVertex -650, 0, 0
    Case 10
    For x = 0 To 300 Step 50
      AddVertex 150 - (x / 2), x + 200, -150 + (x / 2)
      AddVertex -150 + (x / 2), x + 200, -150 + (x / 2)
      AddVertex -150 + (x / 2), x + 200, 150 - (x / 2)
      AddVertex 150 - (x / 2), x + 200, 150 - (x / 2)
    Next x
    For x = 0 To 300 Step 50
      AddVertex 150 - (x / 2), -150 + (x / 2), x + 200
      AddVertex -150 + (x / 2), -150 + (x / 2), x + 200
      AddVertex -150 + (x / 2), 150 - (x / 2), x + 200
      AddVertex 150 - (x / 2), 150 - (x / 2), x + 200
    Next x
    For x = 0 To 300 Step 50
      AddVertex x + 200, 150 - (x / 2), -150 + (x / 2)
      AddVertex x + 200, -150 + (x / 2), -150 + (x / 2)
      AddVertex x + 200, -150 + (x / 2), 150 - (x / 2)
      AddVertex x + 200, 150 - (x / 2), 150 - (x / 2)
    Next x
    For x = 0 To -300 Step -50
      AddVertex -150 - (x / 2), x - 200, 150 + (x / 2)
      AddVertex 150 + (x / 2), x - 200, -150 - (x / 2)
      AddVertex -150 - (x / 2), x - 200, -150 - (x / 2)
      AddVertex 150 + (x / 2), x - 200, 150 + (x / 2)
    Next x
    For x = 0 To -300 Step -50
      AddVertex -150 - (x / 2), 150 + (x / 2), x - 200
      AddVertex 150 + (x / 2), -150 - (x / 2), x - 200
      AddVertex -150 - (x / 2), -150 - (x / 2), x - 200
      AddVertex 150 + (x / 2), 150 + (x / 2), x - 200
    Next x
    For x = 0 To -300 Step -50
      AddVertex x - 200, -150 - (x / 2), 150 + (x / 2)
      AddVertex x - 200, 150 + (x / 2), -150 - (x / 2)
      AddVertex x - 200, -150 - (x / 2), -150 - (x / 2)
      AddVertex x - 200, 150 + (x / 2), 150 + (x / 2)
    Next x
  End Select
End Sub

Public Sub AnimateShape(Shp As Byte)
  'This sub animates the 3d shape by moving its vertices
  Dim x As Long, y As Long
   
  Select Case Shp
    Case 1
      For x = UBound(Vertex) - 4 To UBound(Vertex)
        With Vertex(x)
          If .bFlag = 0 Then
            .y = .y + 5
            If .y = 400 Then .bFlag = 1
          ElseIf .bFlag = 1 Then
            .z = .z - 5
            If .z = -400 Then .bFlag = 2
          ElseIf .bFlag = 2 Then
            .y = .y - 5
            If .y = -400 Then .bFlag = 3
          ElseIf .bFlag = 3 Then
            .z = .z + 5
            If .z = 0 Then .bFlag = 0
          End If
        End With
      Next x
    Case 4
      For x = UBound(Vertex) - 6 To UBound(Vertex)
        With Vertex(x)
          If .bFlag = 0 Then
            .x = .x + 5 + (.x / 100)
            If .x >= 400 Then .bFlag = 1
          ElseIf .bFlag = 1 Then
            .x = .x - 5 - -(.x / 100)
            If .x <= -400 Then .bFlag = 0
          End If
        End With
      Next x
    Case 6
      For x = 1 To UBound(Vertex)
        With Vertex(x)
          If .bFlag = 0 Then
            .x = .x + 2
            If .x >= 50 Then
              .bFlag = 1
              .x = 50
            End If
          ElseIf .bFlag = 1 Then
            .x = .x - 2
            If .x <= -50 Then
              .bFlag = 0
              .x = -50
            End If
          End If
        End With
      Next x
    Case 7
      For x = 1 To UBound(Vertex)
        With Vertex(x)
          If .bFlag = 0 Then
            .x = .x + 20
            If .x >= 500 Then .bFlag = 1
          ElseIf .bFlag = 1 Then
            .x = .x - 20
            If .x <= -500 Then .bFlag = 0
          End If
        End With
      Next x
    Case 8
      For x = 1 To 30
        With Vertex(x)
          If .bFlag = 0 Then
            .x = .x + 2
            If .x >= 50 Then
              .x = 50
              .bFlag = 1
            End If
          ElseIf .bFlag = 1 Then
            .x = .x - 2
            If .x <= -50 Then
              .x = -50
              .bFlag = 0
            End If
          End If
        End With
      Next x
      For x = 31 To 50
        With Vertex(x)
          If Vertex(30).bFlag = 0 Then
            .x = .x + 2
          ElseIf Vertex(30).bFlag = 1 Then
            .x = .x - 2
          End If
        End With
      Next x
      For x = 51 To UBound(Vertex)
        With Vertex(x)
          If Vertex(1).bFlag = 0 Then
            .x = .x + 2
          ElseIf Vertex(1).bFlag = 1 Then
            .x = .x - 2
          End If
        End With
      Next x
    Case 9
      With Vertex(UBound(Vertex))
        .x = .x + 30
        If .x > 2500 Then
          .x = -650
        End If
      End With
  End Select
End Sub
