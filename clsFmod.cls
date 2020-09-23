VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Written by Paul berlin
Dim stream As Long 'Pointers to loaded streams
Dim channel As Long 'Pointers to used channels
Dim module As Long 'pointer to loaded module

Private Sub Class_Terminate()
  FSOUND_Close
  FSOUND_StopSound FSOUND_ALL
  If stream <> 0 Then FSOUND_Stream_Close stream
  If module <> 0 Then
    FMUSIC_StopSong module
    FMUSIC_FreeSong module
  End If
End Sub

Public Function Init(mixrate As Long, Channels As Long)
  'This inits FMOD with selected settings
  On Error GoTo ErrH
  
  If FSOUND_Init(mixrate, Channels, FSOUND_INIT_GLOBALFOCUS) = False Then GoTo ErrH
  
  
  Init = True
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "Init - FMOD Error"
  Init = False
End Function

Public Function ReturnDrivers(DrvArray() As String) As Boolean
  'This returns the names of availabe drives in the array
  'Output must be set before using this or there will be 0 drivers
  Dim x As Integer
  
  On Error GoTo ErrH
  
  If FSOUND_GetNumDrivers > 0 Then
    ReDim DrvArray(FSOUND_GetNumDrivers - 1)
  
    For x = 0 To FSOUND_GetNumDrivers - 1
      DrvArray(x) = GetStringFromPointer(FSOUND_GetDriverName(x))
    Next x
    ReturnDrivers = True
  Else
    ReturnDrivers = False
  End If
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "ReturnDrivers - FMOD Error"
  ReturnDrivers = False
End Function

Public Function InitOutput(OutType As Byte) As Boolean
  'Sets output type
  'MUST BE CALLED BEFORE Init
  'use ReturnDrivers to see which drives support this output type
  On Error GoTo ErrH
  InitOutput = True
  
  If OutType = 0 Then
    If FSOUND_SetOutput(-1) = False Then GoTo ErrH
  ElseIf OutType = 1 Then
    If FSOUND_SetOutput(FSOUND_OUTPUT_DSOUND) = False Then GoTo ErrH
  ElseIf OutType = 2 Then
    If FSOUND_SetOutput(FSOUND_OUTPUT_WINMM) = False Then GoTo ErrH
  End If
  
  If DebugMode Then
    Debug.Print "-InitOutput-"
    Debug.Print "Wanting to use: " & OutType
    If FSOUND_GetOutput And FSOUND_OUTPUT_DSOUND Then
      Debug.Print "Using: DirectSound"
    ElseIf FSOUND_GetOutput And FSOUND_OUTPUT_WINMM Then
      Debug.Print "Using: WinMM"
    End If
  End If
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitOutput - FMOD Error"
  InitOutput = False
End Function

Public Function InitMixer(mixer As Byte) As Boolean
  'Sets mixer to use
  'MUST BE CALLED BEFORE Init
  '0 = autodetect, 1 = CPU/FPU 32-bit, MMX P5 16-bit, MMX P6 16-bit
  On Error GoTo ErrH
  InitMixer = True
  
  Select Case mixer
    Case 0
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_AUTODETECT) = False Then GoTo ErrH
    Case 1
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_FPU) = False Then GoTo ErrH
    Case 2
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_MMXP5) = False Then GoTo ErrH
    Case 3
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_MMXP6) = False Then GoTo ErrH
  End Select
  
  If DebugMode Then
    Debug.Print "-InitMixer-"
    Debug.Print "Using Mixer: " & FSOUND_GetMixer
  End If
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitMixer - FMOD Error"
  InitMixer = False
End Function

Public Function InitBuffer(ms As Long) As Boolean
  'Sets buffer size in ms
  'MUST BE CALLED BEFORE Init
  InitBuffer = True
  If FSOUND_SetBufferSize(ms) = False Then GoTo ErrH

  If DebugMode Then
    Debug.Print "-InitBuffer-"
    Debug.Print "Buffer size: " & ms
  End If

  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitBuffer - FMOD Error"
  InitBuffer = False
End Function

Public Function InitDriver(DriveNum As Byte) As Boolean
  'Must be called before init
  'drive num 0 will set to default driver
  On Error GoTo ErrH
  InitDriver = True
  If FSOUND_SetDriver(DriveNum) = False Then GoTo ErrH
  
  If DebugMode Then
    Debug.Print "-InitDriver-"
    Debug.Print "Num Drivers: " & FSOUND_GetNumDrivers
    Debug.Print "Wanting to use Driver: " & DriveNum, GetStringFromPointer(FSOUND_GetDriverName(DriveNum))
    Debug.Print "Using Driver: " & FSOUND_GetDriver, GetStringFromPointer(FSOUND_GetDriverName(FSOUND_GetDriver))
  End If

  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitDriver - FMOD Error"
  InitDriver = False
End Function

Public Function StreamPlay(Filename As String) As Boolean
  'Will play the filename stream
  On Error GoTo ErrH
  StreamPlay = True

  FSOUND_StopSound FSOUND_ALL
  If stream <> 0 Then FSOUND_Stream_Close stream
  stream = FSOUND_Stream_OpenFile(Filename, FSOUND_NORMAL, 0)
  channel = FSOUND_Stream_Play(FSOUND_FREE, stream)

  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamPlay = False
End Function

Public Function StreamPausePlay() As Boolean
  'this will pause or play the current stream
  On Error GoTo ErrH
  
  If stream <> 0 Then
    StreamPausePlay = True

    If FSOUND_GetPaused(channel) = 0 Then
      FSOUND_SetPaused channel, True
    Else
      FSOUND_SetPaused channel, False
    End If
  End If
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamPausePlay = False
End Function

Public Function StreamStop() As Boolean
  'this will stop and unload stream
  On Error GoTo ErrH
  StreamStop = True

  FSOUND_StopSound FSOUND_ALL
  If stream <> 0 Then FSOUND_Stream_Close stream
  stream = 0
  channel = 0
    

  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamStop = False
End Function

Public Property Get StreamFrequency() As Long
  StreamFrequency = FSOUND_GetFrequency(channel)
End Property

Public Property Let StreamFrequency(ByVal lFreq As Long)
  FSOUND_SetFrequency channel, lFreq
End Property

Public Function StreamMute() As Boolean
  'this will mute/unmute stream
  On Error GoTo ErrH
  
  If stream <> 0 Then
    StreamMute = True

    If FSOUND_GetMute(channel) = 0 Then
      FSOUND_SetMute channel, True
    Else
      FSOUND_SetMute channel, False
    End If
  End If
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
  StreamMute = False
End Function

Public Property Get Volume() As Byte
  Volume = FSOUND_GetVolume(channel)
End Property

Public Property Let Volume(ByVal bVolume As Byte)
  FSOUND_SetVolume channel, bVolume
End Property

Public Property Get IsLoaded() As Boolean
  If stream <> 0 Then IsLoaded = True
End Property

Public Property Get IsPlaying() As Boolean
  If FSOUND_GetPaused(channel) = 0 Then
    IsPlaying = True
  End If
End Property

Public Property Get SongLen() As Long
  If stream <> 0 Then
    SongLen = FSOUND_Stream_GetLengthMs(stream)
  End If
End Property

Public Property Get SongPos() As Long
  If stream <> 0 And channel <> 0 Then
    SongPos = FSOUND_Stream_GetTime(stream)
  End If
End Property

Public Property Let SongPos(ByVal lPos As Long)
  If stream <> 0 And channel <> 0 Then
    FSOUND_Stream_SetTime stream, lPos
  End If
End Property

Public Function GetStreamLength(Filename As String) As Long
  Dim tmpStream As Long
  On Error GoTo ErrH
  
  tmpStream = FSOUND_Stream_OpenFile(Filename, FSOUND_NORMAL, 0)
  GetStreamLength = FSOUND_Stream_GetLengthMs(tmpStream) / 1000
  FSOUND_Stream_Close tmpStream
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
End Function

Public Function GetMusicName(Filename As String) As String
  Dim tmpMusic As Long
  On Error GoTo ErrH
  
  tmpMusic = FMUSIC_LoadSong(Filename)
  GetMusicName = GetStringFromPointer(FMUSIC_GetName(tmpMusic))
  FMUSIC_FreeSong tmpMusic
  
  Exit Function
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
End Function

Public Sub MusicPlay(sFilename As String)
  'Will play the filename music
  On Error GoTo ErrH

  If module <> 0 Then
    FMUSIC_StopSong module
    FMUSIC_FreeSong module
  End If
  module = FMUSIC_LoadSong(sFilename)
  Call FMUSIC_PlaySong(module)

  Exit Sub
ErrH:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "FMOD Error"
End Sub
