Attribute VB_Name = "FMod"
Option Explicit

'
'FMOD VB6 Module
'
'Created by: Adion <adion@quakecity.net>
'Updated by: KarLKoX <XxKarLKoXxX@aol.com>
'Last Update: 14/02/02 (dd/mm/yy)
'Revision # : 24
'For FMod version : 3.5
'Please mail me for any errors you find here

Public Const FMOD_VERSION = 3.5

'************
'* Enums
'************
'FSOUND_GetError
Public Enum FMOD_ERRORS
    FMOD_ERR_NONE             ' No errors
    FMOD_ERR_BUSY             ' Cannot call this command after FSOUND_Init.  Call FSOUND_Close first.
    FMOD_ERR_UNINITIALIZED    ' This command failed because FSOUND_Init was not called
    FMOD_ERR_INIT             ' Error initializing output device.
    FMOD_ERR_ALLOCATED        ' Error initializing output device, but more specifically, the output device is already in use and cannot be reused.
    FMOD_ERR_PLAY             ' Playing the sound failed.
    FMOD_ERR_OUTPUT_FORMAT    ' Soundcard does not support the features needed for this soundsystem (16bit stereo output)
    FMOD_ERR_COOPERATIVELEVEL ' Error setting cooperative level for hardware.
    FMOD_ERR_CREATEBUFFER     ' Error creating hardware sound buffer.
    FMOD_ERR_FILE_NOTFOUND    ' File not found
    FMOD_ERR_FILE_FORMAT      ' Unknown file format
    FMOD_ERR_FILE_BAD         ' Error loading file
    FMOD_ERR_MEMORY           ' Not enough memory
    FMOD_ERR_VERSION          ' The version number of this file format is not supported
    FMOD_ERR_INVALID_PARAM    ' An invalid parameter was passed to this function
    FMOD_ERR_NO_EAX           ' Tried to use an EAX command on a non EAX enabled channel or output.
    FMOD_ERR_CHANNEL_ALLOC    ' Failed to allocate a new channel
    FMOD_ERR_RECORD           ' Recording is not supported on this machine
    FMOD_ERR_MEDIAPLAYER      ' Required Mediaplayer codec is not installed
End Enum

'FSOUND_SetOutput, FSOUND_GetOutput
Public Enum FSOUND_OUTPUTTYPES
    FSOUND_OUTPUT_NOSOUND    'NoSound driver, all calls to this succeed but do nothing.
    FSOUND_OUTPUT_WINMM      'Windows Multimedia driver.
    FSOUND_OUTPUT_DSOUND     'DirectSound driver.  You need this to get EAX2 or EAX3 support, or FX api support.
    
    FSOUND_OUTPUT_A3D        'A3D driver.  You need this to get geometry and EAX reverb support.
    FSOUND_OUTPUT_OSS        'Linux/Unix OSS (Open Sound System) driver, i.e. the kernel sound drivers.
    FSOUND_OUTPUT_ESD        'Linux/Unix ESD (Enlightment Sound Daemon) driver.
    
    FSOUND_OUTPUT_ALSA       'Linux Alsa driver.
    
     FSOUND_OUTPUT_XBOX      'Xbox driver
End Enum

'FSOUND_SetMixer, FSOUND_GetMixer
Public Enum FSOUND_MIXERTYPES
    FSOUND_MIXER_AUTODETECT         'Windows CE only - Non interpolating/low quality mixer. */
    FSOUND_MIXER_BLENDMODE          'removed / obsolete.
    FSOUND_MIXER_MMXP5              'removed / obsolete.
    FSOUND_MIXER_MMXP6              'removed / obsolete.

    FSOUND_MIXER_QUALITY_AUTODETECT 'All platforms - Autodetect the fastest quality mixer based on your cpu.
    FSOUND_MIXER_QUALITY_FPU        'Win32/Linux only - Interpolating/volume ramping FPU mixer.
    FSOUND_MIXER_QUALITY_MMXP5      'Win32/Linux only - Interpolating/volume ramping FPU mixer.
    FSOUND_MIXER_QUALITY_MMXP6      'Win32/Linux only - Interpolating/volume ramping ppro+ MMX mixer.
    
    FSOUND_MIXER_MONO               ' Windows CE only - MONO non interpolating/low quality mixer. For speed
    FSOUND_MIXER_QUALITY_MONO       ' Windows CE only - MONO Interpolating mixer.  For speed
End Enum

'FMUSIC_GetType
Public Enum FMUSIC_TYPES
    FMUSIC_TYPE_NONE
    FMUSIC_TYPE_MOD        'Protracker / Fasttracker
    FMUSIC_TYPE_S3M        'ScreamTracker 3
    FMUSIC_TYPE_XM         'FastTracker 2
    FMUSIC_TYPE_IT         'Impulse Tracker.
    FMUSIC_TYPE_MIDI       'MIDI file
End Enum

Public Enum FSOUND_DSP_PRIORITIES
    FSOUND_DSP_DEFAULTPRIORITY_CLEARUNIT = 0           'DSP CLEAR unit - done first
    FSOUND_DSP_DEFAULTPRIORITY_SFXUNIT = 100           'DSP SFX unit - done second
    FSOUND_DSP_DEFAULTPRIORITY_MUSICUNIT = 200         'DSP MUSIC unit - done third
    FSOUND_DSP_DEFAULTPRIORITY_USER = 300              'User priority, use this as reference
    FSOUND_DSP_DEFAULTPRIORITY_FFTUNIT = 900           'This reads data for FSOUND_DSP_GetSpectrum, so it comes after user units
    FSOUND_DSP_DEFAULTPRIORITY_CLIPANDCOPYUNIT = 1000  'DSP CLIP AND COPY unit - last
End Enum

'FSOUND_DSP_Create, FSOUND_DSP_SetPriority
Public Enum FSOUND_CAPS
    FSOUND_CAPS_HARDWARE = &H1                   ' This driver supports hardware accelerated 3d sound.
    FSOUND_CAPS_EAX2 = &H2                       ' This driver supports EAX 2 reverb
    FSOUND_CAPS_GEOMETRY_OCCLUSIONS = &H4        ' This driver supports (A3D) geometry occlusions
    FSOUND_CAPS_GEOMETRY_REFLECTIONS = &H8       ' This driver supports (A3D) geometry reflections
    FSOUND_CAPS_EAX3 = &H10                      ' This driver supports EAX 3 reverb
End Enum

'FSOUND_MODES
Public Enum FSOUND_MODES
    FSOUND_LOOP_OFF = &H1            ' For non looping samples.
    FSOUND_LOOP_NORMAL = &H2         ' For forward looping samples.
    FSOUND_LOOP_BIDI = &H4           ' For bidirectional looping samples.  (no effect if in hardware).
    FSOUND_8BITS = &H8               ' For 8 bit samples.
    FSOUND_16BITS = &H10             ' For 16 bit samples.
    FSOUND_MONO = &H20               ' For mono samples.
    FSOUND_STEREO = &H40             ' For stereo samples.
    FSOUND_UNSIGNED = &H80           ' For source data containing unsigned samples.
    FSOUND_SIGNED = &H100            ' For source data containing signed data.
    FSOUND_DELTA = &H200             ' For source data stored as delta values.
    FSOUND_IT214 = &H400             ' For source data stored using IT214 compression.
    FSOUND_IT215 = &H800             ' For source data stored using IT215 compression.
    FSOUND_HW3D = &H1000             ' Attempts to make samples use 3d hardware acceleration. (if the card supports it)
    FSOUND_2D = &H2000               ' Ignores any 3d processing.  overrides FSOUND_HW3D.  Located in software.
    FSOUND_STREAMABLE = &H4000       ' For realtime streamable samples.  If you dont supply this sound may come out corrupted.
    FSOUND_LOADMEMORY = &H8000       ' For FSOUND_Sample_Load - 'name' will be interpreted as a pointer to data
    FSOUND_LOADRAW = &H10000         ' For FSOUND_Sample_Load/FSOUND_Stream_Open - will ignore file format and treat as raw pcm.
    FSOUND_MPEGACCURATE = &H20000    ' For FSOUND_Stream_Open - scans MP2/MP3 (VBR also) for accurate FSOUND_Stream_GetLengthMs/FSOUND_Stream_SetTime.
    FSOUND_FORCEMONO = &H40000       ' For forcing stereo streams and samples to be mono - needed with FSOUND_HW3D - incurs speed hit
    FSOUND_HW2D = &H80000            ' 2d hardware sounds.  allows hardware specific effects
    FSOUND_ENABLEFX = &H100000       ' Allows DX8 FX to be played back on a sound.  Requires DirectX 8 - Note these sounds cant be played more than once, or have a changing frequency  
    
    FSOUND_MPEGHALFRATE = &H200000   ' For FMODCE only - decodes mpeg streams using a lower quality decode, but faster execution
    FSOUND_XADPCM = &H400000         ' For XBOX only - Describes a user sample that its contents are compressed as XADPCM
    
    FSOUND_NORMAL = FSOUND_LOOP_OFF Or FSOUND_8BITS Or FSOUND_MONO
End Enum

'FSOUND_CD_SetPlayMode
Public Enum FSOUND_CDPLAYMODES
    FSOUND_CD_PLAYCONTINUOUS        'Starts from the current track and plays to end of CD.
    FSOUND_CD_PLAYONCE              'Plays the specified track then stops.
    FSOUND_CD_PLAYLOOPED            'Plays the specified track looped, forever until stopped manually.
    FSOUND_CD_PLAYRANDOM            'Plays tracks in random order
End Enum

'Miscellaneous values for FMOD functions.
'FSOUND_PlaySound, FSOUND_PlaySoundEx, FSOUND_Sample_Alloc, FSOUND_Sample_Load, FSOUND_SetPan
Public Enum FSOUND_CHANNELSAMPLEMODE
    FSOUND_FREE = -1                 ' definition for dynamically allocated channel or sample
    FSOUND_UNMANAGED = -2            ' definition for allocating a sample that is NOT managed by fsound
    FSOUND_ALL = -3                  ' for a channel index or sample index, this flag affects ALL channels or samples available!  Not supported by all functions.
    FSOUND_STEREOPAN = -1            ' definition for full middle stereo volume on both channels
    FSOUND_SYSTEMCHANNEL = -1000     ' special channel ID for channel based functions that want to alter the global FSOUND software mixing output channel
End Enum

' FSOUND_Reverb_SetProperties, FSOUND_Reverb_GetProperties, FSOUND_REVERB_PROPERTYFLAGS
Public Type FSOUND_REVERB_PROPERTIES
                                    ' MIN     MAX    DEFAULT DESCRIPTION
    Environment         As Long     ' 0       25     0       sets all listener properties
    EnvSize             As Single   ' 1.0     100.0  7.5     environment size in meters
    EnvDiffusion        As Single   ' 0.0     1.0    1.0     environment diffusion
    Room                As Long     ' -10000  0      -1000   room effect level (at mid frequencies)
    RoomHF              As Long     ' -10000  0      -100    relative room effect level at high frequencies
    RoomLF              As Long     ' -10000  0      0       relative room effect level at low frequencies
    DecayTime           As Single   ' 0.1     20.0   1.49    reverberation decay time at mid frequencies
    DecayHFRatio        As Single   ' 0.1     2.0    0.83    high-frequency to mid-frequency decay time ratio
    DecayLFRatio        As Single   ' 0.1     2.0    1.0     low-frequency to mid-frequency decay time ratio
    Reflections         As Long     ' -10000  1000   -2602   early reflections level relative to room effect
    ReflectionsDelay    As Single   ' 0.0     0.3    0.007   initial reflection delay time
    ReflectionsPan(3)   As Single   '                0,0,0   early reflections panning vector
    Reverb              As Long     ' -1000   2000   200     late reverberation level relative to room effect
    ReverbDelay         As Single   ' 0.0     0.1    0.011   late reverberation delay time relative to initial reflection
    ReverbPan(3)        As Single   '                0,0,0   late reverberation panning vector
    EchoTime            As Single   ' .075    0.25   0.25    echo time
    EchoDepth           As Single   ' 0.0     1.0    0.0     echo depth
    ModulationTime      As Single   ' 0.04    4.0    0.25    modulation time
    ModulationDepth     As Single   ' 0.0     1.0    0.0     modulation depth
    AirAbsorptionHF     As Single   ' -100    0.0    -5.0    change in level per meter at high frequencies
    HFReference         As Single   ' 1000.0  20000  5000.0  reference high frequency (hz)
    LFReference         As Single   ' 20.0    1000.0 250.0   reference low frequency (hz)
    RoomRolloffFactor   As Single   ' 0.0     10.0   0.0     like FSOUND_3D_Listener_SetRolloffFactor but for room effect
    Diffusion           As Single   ' 0.0     100.0  100.0   Value that controls the echo density in the late reverberation decay. (xbox only)
    Density             As Single   ' 0.0     100.0  100.0   Value that controls the modal density in the late reverberation decay (xbox only)
    Flags               As Long     '                        modifies the behavior of above properties
End Type


'Values for the Flags member of the FSOUND_REVERB_PROPERTIES structure.
Public Enum FSOUND_REVERB_PROPERTYFLAGS
    FSOUND_REVERBFLAGS_DECAYTIMESCALE = &H1        ' EnvironmentSize affects reverberation decay time
    FSOUND_REVERBFLAGS_REFLECTIONSSCALE = &H2      ' EnvironmentSize affects reflection level
    FSOUND_REVERBFLAGS_REFLECTIONSDELAYSCALE = &H4 ' EnvironmentSize affects initial reflection delay time
    FSOUND_REVERBFLAGS_REVERBSCALE = &H8           ' EnvironmentSize affects reflections level
    FSOUND_REVERBFLAGS_REVERBDELAYSCALE = &H10     ' EnvironmentSize affects late reverberation delay time
    FSOUND_REVERBFLAGS_DECAYHFLIMIT = &H20         ' AirAbsorptionHF affects DecayHFRatio
    FSOUND_REVERBFLAGS_ECHOTIMESCALE = &H40        ' EnvironmentSize affects echo time
    FSOUND_REVERBFLAGS_MODULATIONTIMESCALE = &H80  ' EnvironmentSize affects modulation time
	FSOUND_REVERBFLAGS_DEFAULT = FSOUND_REVERBFLAGS_DECAYTIMESCALE Or FSOUND_REVERBFLAGS_REFLECTIONSSCALE Or FSOUND_REVERBFLAGS_REFLECTIONSDELAYSCALE Or FSOUND_REVERBFLAGS_REVERBSCALE Or FSOUND_REVERBFLAGS_REVERBDELAYSCALE Or FSOUND_REVERBFLAGS_DECAYHFLIMIT
End Enum

' FSOUND_REVERB_PRESETS
' don't know how to set that in VB ?
' #define FSOUND_PRESET_GENERIC          {0,  7.5f,   1.000f, -1000,  -100,   0,  1.49f,  0.83f,  1.00f,  -2602,  0.007f, 0.00f,0.00f,0.00f,  200,    0.011f,     0.00f,0.00f,0.00f,  0.250f, 0.000f, 0.250f, 0.000f, -5.0f,  5000.0f, 250.0f, 0.00f, 0x3f }

'FSOUND_Reverb_SetEnvironment, FSOUND_Reverb_SetEnvironmentAdvanced
Public Type FSOUND_REVERB_CHANNELPROPERTIES
    Direct              As Long     ' direct path level (at low and mid frequencies)
    DirectHF            As Long     ' relative direct path level at high frequencies
    Room                As Long     ' room effect level (at low and mid frequencies)
    RoomHF              As Long     ' relative room effect level at high frequencies
    Obstruction         As Long     ' main obstruction control (attenuation at high frequencies)
    ObstructionLFRatio  As Single   ' obstruction low-frequency level re. main control
    Occlusion           As Long     ' main occlusion control (attenuation at high frequencies)
    OcclustionLFRatio   As Single   ' occlusion low-frequency level re. main control
    OcclusionRoomRatio  As Single   ' relative occlusion control for room effect
    OcclusionDirectRatio As Single  ' relative occlusion control for direct path
    Exclusion           As Long     ' main exlusion control (attenuation at high frequencies)
    ExclusionLFRatio    As Single   ' exclusion low-frequency level re. main control
    OutsideVolumeHF     As Long     ' outside sound cone level at high frequencies
    DopplerFactor       As Single   ' like DS3D flDopplerFactor but per source
    RolloffFactor       As Single   ' like DS3D flRolloffFactor but per source
    RoomRolloffFactor   As Single   ' like DS3D flRolloffFactor but for room effect
    AirAbsorptionFactor As Single   ' multiplies AirAbsorptionHF member of FSOUND_REVERB_PROPERTIES
    Flags               As Long     ' modifies the behavior of properties
End Type

'FSOUND_GEOMETRY_MODES
'FSOUND_Geometry_AddPolygon
Public Enum FSOUND_GEOMETRY_MODES
    FSOUND_GEOMETRY_NORMAL = &H0                 ' Default geometry type.  Occluding polygon
    FSOUND_GEOMETRY_REFLECTIVE = &H1             ' This polygon is reflective
    FSOUND_GEOMETRY_OPENING = &H2                ' Overlays a transparency over the previous polygon.  The 'openingfactor' value supplied is copied internally.
    FSOUND_GEOMETRY_OPENING_REFERENCE = &H4      ' Overlays a transparency over the previous polygon.  The 'openingfactor' supplied is pointed to (for access when building a list)
End Enum


'FSOUND_FX_MODES
Public Enum FSOUND_FX_MODES
    FSOUND_FX_CHORUS
    FSOUND_FX_COMPRESSOR
    FSOUND_FX_DISTORTION
    FSOUND_FX_ECHO
    FSOUND_FX_FLANGER
    FSOUND_FX_GARGLE
    FSOUND_FX_I3DL2REVERB
    FSOUND_FX_PARAMEQ
    FSOUND_FX_WAVES_REVERB
End Enum


'FSOUND_SPEAKERMODES
'FSOUND_SetSpeakerMode
'These are speaker types defined for use with the FSOUND_SetSpeakerMode command.
'Only works with FSOUND_OUTPUT_DSOUND output mode.
Public Enum FSOUND_SPEAKERMODES
    FSOUND_SPEAKERMODE_5POINT1       ' The audio is played through a speaker arrangement of surround speakers with a subwoofer.
    FSOUND_SPEAKERMODE_HEADPHONE     ' The speakers are headphones.
    FSOUND_SPEAKERMODE_MONO          ' The speakers are monaural.
    FSOUND_SPEAKERMODE_QUAD          ' The speakers are quadraphonic.
    FSOUND_SPEAKERMODE_STEREO        ' The speakers are stereo (default value).
    FSOUND_SPEAKERMODE_SURROUND      ' The speakers are surround sound.
End Enum

'************
'* DEFINES (Constants)
'************
Public Const FSOUND_INIT_USEDEFAULTMIDISYNTH = &H1  'Causes MIDI playback to force software decoding.
Public Const FSOUND_INIT_GLOBALFOCUS = &H2          'For DirectSound output - sound is not muted when window is out of focus.
Public Const FSOUND_INIT_ENABLEOUTPUTFX = &H4       'For DirectSound output - Allows FSOUND_FX api to be used on global software mixer output!
Public Const FSOUND_INIT_ACCURATEGETCURRENTVU = &H8 'This latency adjusts FSOUND_GetCurrentVU, but incurs a cpu and very small memory hit

'************
'* FSOUND Init/Global
'************
'Initialize (Before FSOUND_Init)
Public Declare Function FSOUND_SetOutput Lib "fmod.dll" Alias "_FSOUND_SetOutput@4" (ByVal outputtype As FSOUND_OUTPUTTYPES) As Byte
Public Declare Function FSOUND_SetDriver Lib "fmod.dll" Alias "_FSOUND_SetDriver@4" (ByVal driver As Long) As Byte
Public Declare Function FSOUND_SetMixer Lib "fmod.dll" Alias "_FSOUND_SetMixer@4" (ByVal mixer As FSOUND_MIXERTYPES) As Byte
Public Declare Function FSOUND_SetBufferSize Lib "fmod.dll" Alias "_FSOUND_SetBufferSize@4" (ByVal lenms As Long) As Byte
Public Declare Function FSOUND_SetHWND Lib "fmod.dll" Alias "_FSOUND_SetHWND@4" (ByVal hwnd As Long) As Byte
Public Declare Function FSOUND_SetMinHardwareChannels Lib "fmod.dll" Alias "_FSOUND_SetMinHardwareChannels@4" (ByVal min As Integer) As Byte
Public Declare Function FSOUND_SetMaxHardwareChannels Lib "fmod.dll" Alias "_FSOUND_SetMaxHardwareChannels@4" (ByVal min As Integer) As Byte
Public Declare Function FSOUND_SetMemorySystem Lib "fmod.dll" Alias "_FSOUND_SetMemorySystem@20" (ByVal pool As Long, ByVal poollen As Long, ByVal useralloc As Long, ByVal userrealloc As Long, ByVal userfree As Long) As Byte
'Main
Public Declare Function FSOUND_Init Lib "fmod.dll" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal Flags As Long) As Byte
Public Declare Function FSOUND_Close Lib "fmod.dll" Alias "_FSOUND_Close@0" () As Long
'Runtime
Public Declare Function FSOUND_SetSpeakerMode Lib "fmod.dll" Alias "_FSOUND_SetSpeakerMode@4" (ByVal speakermode As Long) As Long
Public Declare Function FSOUND_SetSFXMasterVolume Lib "fmod.dll" Alias "_FSOUND_SetSFXMasterVolume@4" (ByVal volume As Long) As Long
Public Declare Function FSOUND_SetPanSeperation Lib "fmod.dll" Alias "_FSOUND_SetPanSeperation@4" (ByVal pansep As Single) As Long
Public Declare Function FSOUND_File_SetCallbacks Lib "fmod.dll" Alias "_FSOUND_File_SetCallbacks@20" (ByVal OpenCallback As Long, ByVal CloseCallback As Long, ByVal ReadCallback As Long, ByVal SeekCallback As Long, ByVal TellCallback As Long) As Long
'System Information
Public Declare Function FSOUND_GetError Lib "fmod.dll" Alias "_FSOUND_GetError@0" () As FMOD_ERRORS
Public Declare Function FSOUND_GetVersion Lib "fmod.dll" Alias "_FSOUND_GetVersion@0" () As Single
Public Declare Function FSOUND_GetOutput Lib "fmod.dll" Alias "_FSOUND_GetOutput@0" () As FSOUND_OUTPUTTYPES
Public Declare Function FSOUND_GetOutputHandle Lib "fmod.dll" Alias "_FSOUND_GetOutputHandle@0" () As Long
Public Declare Function FSOUND_GetDriver Lib "fmod.dll" Alias "_FSOUND_GetDriver@0" () As Long
Public Declare Function FSOUND_GetMixer Lib "fmod.dll" Alias "_FSOUND_GetMixer@0" () As FSOUND_MIXERTYPES
Public Declare Function FSOUND_GetNumDrivers Lib "fmod.dll" Alias "_FSOUND_GetNumDrivers@0" () As Long
Public Declare Function FSOUND_GetDriverName Lib "fmod.dll" Alias "_FSOUND_GetDriverName@4" (ByVal id As Long) As Long
Public Declare Function FSOUND_GetDriverCaps Lib "fmod.dll" Alias "_FSOUND_GetDriverCaps@8" (ByVal id As Long, ByRef caps As Long) As Byte

Public Declare Function FSOUND_GetOutputRate Lib "fmod.dll" Alias "_FSOUND_GetOutputRate@0" () As Long
Public Declare Function FSOUND_GetMaxChannels Lib "fmod.dll" Alias "_FSOUND_GetMaxChannels@0" () As Long
Public Declare Function FSOUND_GetMaxSamples Lib "fmod.dll" Alias "_FSOUND_GetMaxSamples@0" () As Long
Public Declare Function FSOUND_GetSFXMasterVolume Lib "fmod.dll" Alias "_FSOUND_GetSFXMasterVolume@0" () As Long
Public Declare Function FSOUND_GetNumHardwareChannels Lib "fmod.dll" Alias "_FSOUND_GetNumHardwareChannels@0" () As Long
Public Declare Function FSOUND_GetChannelsPlaying Lib "fmod.dll" Alias "_FSOUND_GetChannelsPlaying@0" () As Long
Public Declare Function FSOUND_GetCPUUsage Lib "fmod.dll" Alias "_FSOUND_GetCPUUsage@0" () As Single

'************
'* FSOUND Samples
'************
'Sample creation and management functions
Public Declare Function FSOUND_Sample_Load Lib "fmod.dll" Alias "_FSOUND_Sample_Load@16" (ByVal Index As Long, ByVal name As String, ByVal mode As FSOUND_MODES, ByVal memlength As Long) As Long
Public Declare Function FSOUND_Sample_Alloc Lib "fmod.dll" Alias "_FSOUND_Sample_Alloc@28" (ByVal Index As Long, ByVal length As Long, ByVal mode As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long) As Long
Public Declare Function FSOUND_Sample_Free Lib "fmod.dll" Alias "_FSOUND_Sample_Free@4" (ByVal sptr As Long) As Long
Public Declare Function FSOUND_Sample_Upload Lib "fmod.dll" Alias "_FSOUND_Sample_Upload@12" (ByVal sptr As Long, ByVal srcdata As Long, ByVal mode As Long) As Byte
Public Declare Function FSOUND_Sample_Lock Lib "fmod.dll" Alias "_FSOUND_Sample_Lock@28" (ByVal sptr As Long, ByVal offset As Long, ByVal length As Long, ByVal ptr1 As Long, ByVal ptr2 As Long, ByVal len1 As Long, ByVal len2 As Long) As Byte
Public Declare Function FSOUND_Sample_Unlock Lib "fmod.dll" Alias "_FSOUND_Sample_Unlock@20" (ByVal sptr As Long, ByVal sptr1 As Long, ByVal sptr2 As Long, ByVal len1 As Long, ByVal len2 As Long) As Byte

'Sample control functions
Public Declare Function FSOUND_Sample_SetLoopMode Lib "fmod.dll" Alias "_FSOUND_Sample_SetLoopMode@8" (ByVal sptr As Long, ByVal loopmode As FSOUND_MODES) As Byte
Public Declare Function FSOUND_Sample_SetLoopPoints Lib "fmod.dll" Alias "_FSOUND_Sample_SetLoopPoints@12" (ByVal sptr As Long, ByVal loopstart As Long, ByVal loopend As Long) As Byte
Public Declare Function FSOUND_Sample_SetDefaults Lib "fmod.dll" Alias "_FSOUND_Sample_SetDefaults@20" (ByVal sptr As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long) As Byte
Public Declare Function FSOUND_Sample_SetMinMaxDistance Lib "fmod.dll" Alias "_FSOUND_Sample_SetMinMaxDistance@12" (ByVal sptr As Long, ByVal min As Single, ByVal max As Single) As Byte

'Sample information
Public Declare Function FSOUND_Sample_Get Lib "fmod.dll" Alias "_FSOUND_Sample_Get@4" (ByVal sampno As Long) As Long
Public Declare Function FSOUND_Sample_GetName Lib "fmod.dll" Alias "_FSOUND_Sample_GetName@4" (ByVal sptr As Long) As Long
Public Declare Function FSOUND_Sample_GetLength Lib "fmod.dll" Alias "_FSOUND_Sample_GetLength@4" (ByVal sptr As Long) As Long
Public Declare Function FSOUND_Sample_GetLoopPoints Lib "fmod.dll" Alias "_FSOUND_Sample_GetLoopPoints@12" (ByVal sptr As Long, ByRef loopstart As Long, ByRef loopend As Long) As Byte
Public Declare Function FSOUND_Sample_GetDefaults Lib "fmod.dll" Alias "_FSOUND_Sample_GetDefaults@20" (ByVal sptr As Long, ByRef deffreq As Long, ByRef defvol As Long, ByRef defpan As Long, ByRef defpri As Long) As Byte
Public Declare Function FSOUND_Sample_GetMode Lib "fmod.dll" Alias "_FSOUND_Sample_GetMode@4" (ByVal sptr As Long) As Long

'************
'* Channel control functions
'************
'Playing and stopping sounds
Public Declare Function FSOUND_PlaySound Lib "fmod.dll" Alias "_FSOUND_PlaySound@8" (ByVal channel As Long, ByVal sptr As Long) As Long
Public Declare Function FSOUND_PlaySoundEx Lib "fmod.dll" Alias "_FSOUND_PlaySoundEx@16" (ByVal channel As Long, ByVal sptr As Long, ByVal dsp As Long, ByVal startpaused As Byte) As Long
Public Declare Function FSOUND_StopSound Lib "fmod.dll" Alias "_FSOUND_StopSound@4" (ByVal channel As Long) As Byte

'Functions to control playback of a channel
Public Declare Function FSOUND_SetFrequency Lib "fmod.dll" Alias "_FSOUND_SetFrequency@8" (ByVal channel As Long, ByVal freq As Long) As Byte
Public Declare Function FSOUND_SetVolume Lib "fmod.dll" Alias "_FSOUND_SetVolume@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Public Declare Function FSOUND_SetVolumeAbsolute Lib "fmod.dll" Alias "_FSOUND_SetVolumeAbsolute@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Public Declare Function FSOUND_SetPan Lib "fmod.dll" Alias "_FSOUND_SetPan@8" (ByVal channel As Long, ByVal pan As Long) As Byte
Public Declare Function FSOUND_SetSurround Lib "fmod.dll" Alias "_FSOUND_SetSurround@8" (ByVal channel As Long, ByVal surround As Long) As Byte
Public Declare Function FSOUND_SetMute Lib "fmod.dll" Alias "_FSOUND_SetMute@8" (ByVal channel As Long, ByVal mute As Byte) As Byte
Public Declare Function FSOUND_SetPriority Lib "fmod.dll" Alias "_FSOUND_SetPriority@8" (ByVal channel As Long, ByVal Priority As Long) As Byte
Public Declare Function FSOUND_SetReserved Lib "fmod.dll" Alias "_FSOUND_SetReserved@8" (ByVal channel As Long, ByVal reserved As Long) As Byte
Public Declare Function FSOUND_SetPaused Lib "fmod.dll" Alias "_FSOUND_SetPaused@8" (ByVal channel As Long, ByVal Paused As Byte) As Byte
Public Declare Function FSOUND_SetLoopMode Lib "fmod.dll" Alias "_FSOUND_StopSound@4" (ByVal channel As Long, ByVal loopmode As Long) As Byte
Public Declare Function FSOUND_SetCurrentPosition Lib "fmod.dll" Alias "_FSOUND_SetCurrentPosition@8" (ByVal channel As Long, ByVal offset As Long) As Byte

'Functions to control DX8 only effects processing.
'Note that FX enabled samples can only be played once at a time.
Public Declare Function FSOUND_FX_Enable Lib "fmod.dll" Alias "_FSOUND_FX_Enable@8" (ByVal channel As Long, ByVal fx As FSOUND_FX_MODES) As Long
Public Declare Function FSOUND_FX_SetChorus Lib "fmod.dll" Alias "_FSOUND_FX_SetChorus@32" (ByVal fxid As Long, ByVal WetDryMix As Single, ByVal Depth As Single, ByVal Feedback As Single, ByVal Frequency As Single, ByVal Waveform As Long, ByVal Delay As Single, ByVal Phase As Long) As Byte
Public Declare Function FSOUND_FX_SetCompressor Lib "fmod.dll" Alias "_FSOUND_FX_SetCompressor@28" (ByVal fxid As Long, ByVal Gain As Single, ByVal Attack As Single, ByVal Release As Single, ByVal Threshold As Single, ByVal Ratio As Single, ByVal Predelay As Single) As Byte
Public Declare Function FSOUND_FX_SetDistortion Lib "fmod.dll" Alias "_FSOUND_FX_SetDistortion@24" (ByVal fxid As Long, ByVal Gain As Single, ByVal Edge As Single, ByVal PostEQCenterFrequency As Single, ByVal PostEQBandwidth As Single, ByVal PreLowpassCutoff As Single) As Byte
Public Declare Function FSOUND_FX_SetEcho Lib "fmod.dll" Alias "_FSOUND_FX_SetEcho@24" (ByVal fxid As Long, ByVal WetDryMix As Single, ByVal Feedback As Single, ByVal LeftDelay As Single, ByVal RightDelay As Single, ByVal PanDelay As Long) As Byte
Public Declare Function FSOUND_FX_SetFlanger Lib "fmod.dll" Alias "_FSOUND_FX_SetFlanger@32" (ByVal fxid As Long, ByVal WetDryMix As Single, ByVal Depth As Single, ByVal Feedback As Single, ByVal Frequency As Single, ByVal Waveform As Long, ByVal Delay As Single, ByVal Phase As Long) As Byte
Public Declare Function FSOUND_FX_SetGargle Lib "fmod.dll" Alias "_FSOUND_FX_SetGargle@12" (ByVal fxid As Long, ByVal RateHz As Long, ByVal WaveShape As Long) As Byte
Public Declare Function FSOUND_FX_SetI3DL2Reverb Lib "fmod.dll" Alias "_FSOUND_FX_SetI3DL2Reverb@52" (ByVal fxid As Long, ByVal Room As Long, ByVal RoomHF As Long, ByVal RoomRolloffFactor As Single, ByVal DecayTime As Single, ByVal DecayHFRatio As Single, ByVal Reflections As Long, ByVal ReflectionsDelay As Single, ByVal Reverb As Long, ByVal ReverbDelay As Single, ByVal Diffusion As Single, ByVal Density As Single, ByVal HFReference As Single) As Byte
Public Declare Function FSOUND_FX_SetParamEQ Lib "fmod.dll" Alias "_FSOUND_FX_SetParamEQ@16" (ByVal fxid As Long, ByVal Center As Single, ByVal Bandwidth As Single, ByVal Gain As Single) As Byte
Public Declare Function FSOUND_FX_SetWavesReverb Lib "fmod.dll" Alias "_FSOUND_FX_SetWavesReverb@20" (ByVal fxid As Long, ByVal InGain As Single, ByVal ReverbMix As Single, ByVal ReverbTime As Single, ByVal HighFreqRTRatio As Single) As Byte

'Channel information
Public Declare Function FSOUND_IsPlaying Lib "fmod.dll" Alias "_FSOUND_IsPlaying@4" (ByVal channel As Long) As Byte
Public Declare Function FSOUND_GetFrequency Lib "fmod.dll" Alias "_FSOUND_GetFrequency@4" (ByVal channel As Long) As Long
Public Declare Function FSOUND_GetVolume Lib "fmod.dll" Alias "_FSOUND_GetVolume@4" (ByVal channel As Long) As Long
Public Declare Function FSOUND_GetPan Lib "fmod.dll" Alias "_FSOUND_GetPan@4" (ByVal channel As Long) As Long
Public Declare Function FSOUND_GetSurround Lib "fmod.dll" Alias "_FSOUND_GetSurround@4" (ByVal channel As Long) As Byte
Public Declare Function FSOUND_GetMute Lib "fmod.dll" Alias "_FSOUND_GetMute@4" (ByVal channel As Long) As Byte
Public Declare Function FSOUND_GetPriority Lib "fmod.dll" Alias "_FSOUND_GetPriority@4" (ByVal channel As Long) As Long
Public Declare Function FSOUND_GetReserved Lib "fmod.dll" Alias "_FSOUND_GetReserved@4" (ByVal channel As Long) As Byte
Public Declare Function FSOUND_GetPaused Lib "fmod.dll" Alias "_FSOUND_GetPaused@4" (ByVal channel As Long) As Byte
Public Declare Function FSOUND_GetCurrentPosition Lib "fmod.dll" Alias "_FSOUND_GetCurrentPosition@4" (ByVal channel As Long) As Long
Public Declare Function FSOUND_GetCurrentSample Lib "fmod.dll" Alias "_FSOUND_GetCurrentSample@4" (ByVal channel As Long) As Long
Public Declare Function FSOUND_GetCurrentVU Lib "fmod.dll" Alias "_FSOUND_GetCurrentVU@4" (ByVal channel As Long) As Single

'************
'* FSOUND 3D
'************
'see also FSOUND_Sample_SetMinMaxDistance (above)
Public Declare Function FSOUND_3D_Update Lib "fmod.dll" Alias "_FSOUND_3D_Update@0" () As Long
Public Declare Function FSOUND_3D_SetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_SetAttributes@12" (ByVal channel As Long, ByRef Pos As Single, ByRef vel As Single) As Byte
Public Declare Function FSOUND_3D_GetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_GetAttributes@12" (ByVal channel As Long, ByRef Pos As Single, ByRef vel As Single) As Byte
Public Declare Function FSOUND_3D_Listener_SetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_Listener_SetAttributes@32" (ByVal Pos As Single, ByVal vel As Single, ByVal fx As Single, ByVal fy As Single, ByVal fz As Single, ByVal tx As Single, ByVal ty As Single, ByVal tz As Single) As Long
Public Declare Function FSOUND_3D_Listener_GetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_Listener_GetAttributes@32" (ByRef Pos As Single, ByRef vel As Single, ByRef fx As Single, ByRef fy As Single, ByRef fz As Single, ByRef tx As Single, ByRef ty As Single, ByRef tz As Single) As Long
Public Declare Function FSOUND_3D_Listener_SetDopplerFactor Lib "fmod.dll" Alias "_FSOUND_3D_Listener_SetDopplerFactor@4" (ByVal fscale As Single) As Long
Public Declare Function FSOUND_3D_Listener_SetDistanceFactor Lib "fmod.dll" Alias "_FSOUND_3D_Listener_SetDistanceFactor@4" (ByVal fscale As Single) As Long
Public Declare Function FSOUND_3D_Listener_SetRolloffFactor Lib "fmod.dll" Alias "_FSOUND_3D_Listener_SetRolloffFactor@4" (ByVal fscale As Single) As Long

'************
'* FSOUND Streams
'************
Public Declare Function FSOUND_Stream_OpenFile Lib "fmod.dll" Alias "_FSOUND_Stream_OpenFile@12" (ByVal filename As String, ByVal mode As FSOUND_MODES, ByVal memlength As Long) As Long
Public Declare Function FSOUND_Stream_Create Lib "fmod.dll" Alias "_FSOUND_Stream_Create@20" (ByVal callback As Long, ByVal length As Long, ByVal mode As Long, ByVal samplerate As Long, ByVal userdata As Long) As Long
Public Declare Function FSOUND_Stream_Play Lib "fmod.dll" Alias "_FSOUND_Stream_Play@8" (ByVal channel As Long, ByVal stream As Long) As Long
Public Declare Function FSOUND_Stream_PlayEx Lib "fmod.dll" Alias "_FSOUND_Stream_PlayEx@16" (ByVal channel As Long, ByVal stream As Long, ByVal dsp As Long, ByVal startpaused As Byte) As Long
Public Declare Function FSOUND_Stream_Stop Lib "fmod.dll" Alias "_FSOUND_Stream_Stop@4" (ByVal stream As Long) As Byte
Public Declare Function FSOUND_Stream_Close Lib "fmod.dll" Alias "_FSOUND_Stream_Close@4" (ByVal stream As Long) As Byte
Public Declare Function FSOUND_Stream_SetEndCallback Lib "fmod.dll" Alias "_FSOUND_Stream_SetEndCallback@12" (ByVal stream As Long, ByVal callback As Long, ByVal userdata As Long) As Byte
Public Declare Function FSOUND_Stream_SetSynchCallback Lib "fmod.dll" Alias "_FSOUND_Stream_SetSynchCallback@12" (ByVal stream As Long, ByVal callback As Long, ByVal userdata As Long) As Byte
Public Declare Function FSOUND_Stream_GetSample Lib "fmod.dll" Alias "_FSOUND_Stream_GetSample@4" (ByVal stream As Long) As Long
Public Declare Function FSOUND_Stream_CreateDSP Lib "fmod.dll" Alias "_FSOUND_Stream_CreateDSP@16" (ByVal stream As Long, ByVal callback As Long, ByVal Priority As Long, ByVal userdata As Long) As Long
Public Declare Function FSOUND_Stream_SetBufferSize Lib "fmod.dll" Alias "_FSOUND_Stream_SetBufferSize@4" (ByVal ms As Long) As Byte

Public Declare Function FSOUND_Stream_SetPosition Lib "fmod.dll" Alias "_FSOUND_Stream_SetPosition@8" (ByVal stream As Long, ByVal positition As Long) As Byte
Public Declare Function FSOUND_Stream_GetPosition Lib "fmod.dll" Alias "_FSOUND_Stream_GetPosition@4" (ByVal stream As Long) As Long
Public Declare Function FSOUND_Stream_SetTime Lib "fmod.dll" Alias "_FSOUND_Stream_SetTime@8" (ByVal stream As Long, ByVal ms As Long) As Byte
Public Declare Function FSOUND_Stream_GetTime Lib "fmod.dll" Alias "_FSOUND_Stream_GetTime@4" (ByVal stream As Long) As Long
Public Declare Function FSOUND_Stream_GetLength Lib "fmod.dll" Alias "_FSOUND_Stream_GetLength@4" (ByVal stream As Long) As Long
Public Declare Function FSOUND_Stream_GetLengthMs Lib "fmod.dll" Alias "_FSOUND_Stream_GetLengthMs@4" (ByVal stream As Long) As Long

'************
'* FSOUND CD
'************
Public Declare Function FSOUND_CD_Play Lib "fmod.dll" Alias "_FSOUND_CD_Play@4" (ByVal Track As Long) As Byte
Public Declare Function FSOUND_CD_SetPlayMode Lib "fmod.dll" Alias "_FSOUND_CD_SetPlayMode@4" (ByVal mode As FSOUND_CDPLAYMODES) As Long
Public Declare Function FSOUND_CD_Stop Lib "fmod.dll" Alias "_FSOUND_CD_Stop@0" () As Byte
Public Declare Function FSOUND_CD_SetPaused Lib "fmod.dll" Alias "_FSOUND_CD_SetPaused@4" (ByVal Paused As Byte) As Byte
Public Declare Function FSOUND_CD_SetVolume Lib "fmod.dll" Alias "_FSOUND_CD_SetVolume@4" (ByVal volume As Long) As Byte
Public Declare Function FSOUND_CD_Eject Lib "fmod.dll" Alias "_FSOUND_CD_Eject@0" () As Byte

Public Declare Function FSOUND_CD_GetPaused Lib "fmod.dll" Alias "_FSOUND_CD_GetPaused@0" () As Byte
Public Declare Function FSOUND_CD_GetTrack Lib "fmod.dll" Alias "_FSOUND_CD_GetTrack@0" () As Long
Public Declare Function FSOUND_CD_GetNumTracks Lib "fmod.dll" Alias "_FSOUND_CD_GetNumTracks@0" () As Long
Public Declare Function FSOUND_CD_GetVolume Lib "fmod.dll" Alias "_FSOUND_CD_GetVolume@0" () As Long
Public Declare Function FSOUND_CD_GetTrackLength Lib "fmod.dll" Alias "_FSOUND_CD_GetTrackLength@4" (ByVal Track As Long) As Long
Public Declare Function FSOUND_CD_GetTrackTime Lib "fmod.dll" Alias "_FSOUND_CD_GetTrackTime@0" () As Long

'************
'* FSOUND DSP
'************
'DSP Unit control and information functions
Public Declare Function FSOUND_DSP_Create Lib "fmod.dll" Alias "_FSOUND_DSP_Create@12" (ByVal callback As Long, ByVal Priority As Long, ByVal param As Long) As Long
Public Declare Function FSOUND_DSP_Free Lib "fmod.dll" Alias "_FSOUND_DSP_Free@4" (ByVal unit As Long) As Long
Public Declare Function FSOUND_DSP_SetPriority Lib "fmod.dll" Alias "_FSOUND_DSP_SetPriority@8" (ByVal unit As Long, ByVal Priority As Long) As Long
Public Declare Function FSOUND_DSP_GetPriority Lib "fmod.dll" Alias "_FSOUND_DSP_GetPriority@4" (ByVal unit As Long) As Long
Public Declare Function FSOUND_DSP_SetActive Lib "fmod.dll" Alias "_FSOUND_DSP_SetActive@8" (ByVal unit As Long, ByVal active As Integer) As Long
Public Declare Function FSOUND_DSP_GetActive Lib "fmod.dll" Alias "_FSOUND_DSP_GetActive@4" (ByVal unit As Long) As Byte

'Functions to get hold of FSOUND 'system DSP unit' handles
Public Declare Function FSOUND_DSP_GetClearUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetClearUnit@0" () As Long
Public Declare Function FSOUND_DSP_GetSFXUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetSFXUnit@0" () As Long
Public Declare Function FSOUND_DSP_GetMusicUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetMusicUnit@0" () As Long
Public Declare Function FSOUND_DSP_GetClipAndCopyUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetClipAndCopyUnit@0" () As Long
Public Declare Function FSOUND_DSP_GetFFTUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetFFTUnit@0" () As Long

'misc DSP functions
Public Declare Function FSOUND_DSP_MixBuffers Lib "fmod.dll" Alias "_SOUND_DSP_MixBuffers@28" (ByVal destbuffer As Long, ByVal srcbuffer As Long, ByVal length As Long, ByVal freq As Long, ByVal Vol As Long, ByVal pan As Long, ByVal mode As Long) As Byte
Public Declare Function FSOUND_DSP_ClearMixBuffer Lib "fmod.dll" Alias "_FSOUND_DSP_ClearMixBuffer@0" () As Long
Public Declare Function FSOUND_DSP_GetBufferLength Lib "fmod.dll" Alias "_FSOUND_DSP_GetBufferLength@0" () As Long
' Total buffer length due to FSOUND_SetBufferSize
Public Declare Function FSOUND_DSP_GetBufferLengthTotal Lib "fmod.dll" Alias "_FSOUND_DSP_GetBufferLengthTotal@0" () As Long
Public Declare Function FSOUND_DSP_GetSpectrum Lib "fmod.dll" Alias "_FSOUND_DSP_GetSpectrum@0" () As Long

'************
'* FSOUND Geometry
'************
'scene/polygon functions
Public Declare Function FSOUND_Geometry_AddPolygon Lib "fmod.dll" Alias "_FSOUND_Geometry_AddPolygon@28" (ByRef p1 As Single, ByRef p2 As Single, ByRef p3 As Single, ByRef p4 As Single, ByRef normal As Single, ByVal mode As Long, ByRef openingfactor As Single) As Byte
Public Declare Function FSOUND_Geometry_AddList Lib "fmod.dll" Alias "_FSOUND_Geometry_AddList@4" (ByVal geomlist As Long) As Long

'polygon list functions
Public Declare Function FSOUND_Geometry_List_Create Lib "fmod.dll" Alias "_FSOUND_Geometry_List_Create@4" (ByVal boundingvolume As Long) As Long
Public Declare Function FSOUND_Geometry_List_Free Lib "fmod.dll" Alias "_FSOUND_Geometry_List_Free@4" (ByVal geomlist As Long) As Byte
Public Declare Function FSOUND_Geometry_List_Begin Lib "fmod.dll" Alias "_FSOUND_Geometry_List_Begin@4" (ByVal geomlist As Long) As Byte
Public Declare Function FSOUND_Geometry_List_End Lib "fmod.dll" Alias "_FSOUND_Geometry_List_End@4" (ByVal geomlist As Long) As Byte

'material functions
Public Declare Function FSOUND_Geometry_Material_Create Lib "fmod.dll" Alias "_FSOUND_Geometry_Material_Create@0" () As Long
Public Declare Function FSOUND_Geometry_Material_Free Lib "fmod.dll" Alias "_FSOUND_Geometry_Material_Free@4" (ByVal Material As Long) As Byte
Public Declare Function FSOUND_Geometry_Material_SetAttributes Lib "fmod.dll" Alias "_FSOUND_Geometry_Material_SetAttributes@20" (ByVal Material As Long, ByRef reflectancegain As Single, ByRef reflectancefreq As Single, ByRef transmittancegain As Single, ByRef transmittancefreq As Single) As Byte
Public Declare Function FSOUND_Geometry_Material_GetAttributes Lib "fmod.dll" Alias "_FSOUND_Geometry_Material_GetAttributes@20" (ByVal Material As Long, ByRef reflectancegain As Single, ByRef reflectancefreq As Single, ByRef transmittancegain As Single, ByRef transmittancefreq As Single) As Byte
Public Declare Function FSOUND_Geometry_Material_Set Lib "fmod.dll" Alias "_FSOUND_Geometry_Material_Set@4" (ByVal Material As Long) As Byte

'************
'* FSOUND Reverb functions. (eax2/3 reverb) (NOT SUPPORTED IN LINUX/CE)
'************
'The FSOUND_REVERB_PRESETS have not been included in VB yet so they cannot yet be used here...
Public Declare Function FSOUND_Reverb_SetProperties Lib "fmod.dll" Alias "_FSOUND_Reverb_SetProperties@4" (ByRef prop As FSOUND_REVERB_PROPERTIES) As Byte
Public Declare Function FSOUND_Reverb_GetProperties Lib "fmod.dll" Alias "_FSOUND_Reverb_GetProperties@4" (ByRef prop As FSOUND_REVERB_PROPERTIES) As Byte
Public Declare Function FSOUND_Reverb_SetChannelProperties Lib "fmod.dll" (ByVal channel As Long, ByRef prop As FSOUND_REVERB_CHANNELPROPERTIES) As Byte
Public Declare Function FSOUND_Reverb_GetChannelProperties Lib "fmod.dll" (ByVal channel As Long, ByRef prop As FSOUND_REVERB_CHANNELPROPERTIES) As Byte


'************
'* FSOUND RECORD
'************
'initialization functions
Public Declare Function FSOUND_Record_SetDriver Lib "fmod.dll" Alias "_FSOUND_Record_SetDriver@4" (ByVal outputtype As Long) As Byte
Public Declare Function FSOUND_Record_GetNumDrivers Lib "fmod.dll" Alias "_FSOUND_Record_GetNumDrivers@0" () As Long
Public Declare Function FSOUND_Record_GetDriverName Lib "fmod.dll" Alias "_FSOUND_Record_GetDriverName@4" (ByVal id As Long) As Long
Public Declare Function FSOUND_Record_GetDriver Lib "fmod.dll" Alias "_FSOUND_Record_GetDriver@0" () As Long

'recording functionality.  Only one recording session will work at a time
Public Declare Function FSOUND_Record_StartSample Lib "fmod.dll" Alias "_FSOUND_Record_StartSample@8" (ByVal sample As Long, ByVal loopit As Boolean) As Byte
Public Declare Function FSOUND_Record_Stop Lib "fmod.dll" Alias "_FSOUND_Record_Stop@0" () As Byte
Public Declare Function FSOUND_Record_GetPosition Lib "fmod.dll" Alias "_FSOUND_Record_GetPosition@0" () As Long

'************
'* FMUSIC Modules
'************
'Song management / playback functions
Public Declare Function FMUSIC_LoadSong Lib "fmod.dll" Alias "_FMUSIC_LoadSong@4" (ByVal name As String) As Long
Public Declare Function FMUSIC_LoadSongMemory Lib "fmod.dll" Alias "_FMUSIC_LoadSongMemory@8" (ByRef Data As Long, ByVal length As Long) As Long
Public Declare Function FMUSIC_FreeSong Lib "fmod.dll" Alias "_FMUSIC_FreeSong@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_PlaySong Lib "fmod.dll" Alias "_FMUSIC_PlaySong@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_StopSong Lib "fmod.dll" Alias "_FMUSIC_StopSong@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_StopAllSongs Lib "fmod.dll" Alias "_FMUSIC_StopAllSongs@0" () As Long
Public Declare Function FMUSIC_SetZxxCallback Lib "fmod.dll" Alias "_FMUSIC_SetZxxCallback@8" (ByVal module As Long, ByVal callback As Long) As Byte
Public Declare Function FMUSIC_SetRowCallback Lib "fmod.dll" Alias "_FMUSIC_SetRowCallback@12" (ByVal module As Long, ByVal callback As Long, ByVal rowstep As Long) As Byte
Public Declare Function FMUSIC_SetOrderCallback Lib "fmod.dll" Alias "_FMUSIC_SetOrderCallback@12" (ByVal module As Long, ByVal callback As Long, ByVal rowstep As Long) As Byte
Public Declare Function FMUSIC_SetInstCallback Lib "fmod.dll" Alias "_FMUSIC_SetInstCallback@12" (ByVal module As Long, ByVal callback As Long, ByVal instrument As Long) As Byte
Public Declare Function FMUSIC_SetSample Lib "fmod.dll" Alias "_FMUSIC_SetSample@12" (ByVal module As Long, ByVal sampno As Long, ByRef sptr As Long) As Byte
Public Declare Function FMUSIC_OptimizeChannels Lib "fmod.dll" Alias "_FMUSIC_OptimizeChannels@12" (ByVal module As Long, ByVal maxchannels As Long, ByVal minvolume As Long) As Byte

'Runtime song functions
Public Declare Function FMUSIC_SetReverb Lib "fmod.dll" Alias "_FMUSIC_SetReverb@4" (ByVal Reverb As Byte) As Byte
Public Declare Function FMUSIC_SetOrder Lib "fmod.dll" Alias "_FMUSIC_SetOrder@8" (ByVal module As Long, ByVal order As Long) As Byte
Public Declare Function FMUSIC_SetPaused Lib "fmod.dll" Alias "_FMUSIC_SetPaused@8" (ByVal module As Long, ByVal Pause As Byte) As Byte
Public Declare Function FMUSIC_SetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_SetMasterVolume@8" (ByVal module As Long, ByVal volume As Long) As Byte
Public Declare Function FMUSIC_SetPanSeperation Lib "fmod.dll" Alias "_FMUSIC_SetPanSeperation@8" (ByVal module As Long, ByVal pansep As Single) As Byte

'Static song information functions
Public Declare Function FMUSIC_GetName Lib "fmod.dll" Alias "_FMUSIC_GetName@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetType Lib "fmod.dll" Alias "_FMUSIC_GetType@4" (ByVal module As Long) As FMUSIC_TYPES
Public Declare Function FMUSIC_GetNumOrders Lib "fmod.dll" Alias "_FMUSIC_GetNumOrders@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetNumPatterns Lib "fmod.dll" Alias "_FMUSIC_GetNumPatterns@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetNumInstruments Lib "fmod.dll" Alias "_FMUSIC_GetNumInstruments@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetNumSamples Lib "fmod.dll" Alias "_FMUSIC_GetNumSamples@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetNumChannels Lib "fmod.dll" Alias "_FMUSIC_GetNumChannels@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetSample Lib "fmod.dll" Alias "_FMUSIC_GetSample@8" (ByVal module As Long, ByVal sampno As Long) As Long
Public Declare Function FMUSIC_GetPatternLength Lib "fmod.dll" Alias "_FMUSIC_GetPatternLength@8" (ByVal module As Long, ByVal orderno As Long) As Long

'Runtime song information
Public Declare Function FMUSIC_IsFinished Lib "fmod.dll" Alias "_FMUSIC_IsFinished@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_IsPlaying Lib "fmod.dll" Alias "_FMUSIC_IsPlaying@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_GetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_GetMasterVolume@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetGlobalVolume Lib "fmod.dll" Alias "_FMUSIC_GetGlobalVolume@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetOrder Lib "fmod.dll" Alias "_FMUSIC_GetOrder@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetPattern Lib "fmod.dll" Alias "_FMUSIC_GetPattern@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetSpeed Lib "fmod.dll" Alias "_FMUSIC_GetSpeed@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetBPM Lib "fmod.dll" Alias "_FMUSIC_GetBPM@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetRow Lib "fmod.dll" Alias "_FMUSIC_GetRow@4" (ByVal module As Long) As Long
Public Declare Function FMUSIC_GetPaused Lib "fmod.dll" Alias "_FMUSIC_GetPaused@4" (ByVal module As Long) As Byte
Public Declare Function FMUSIC_GetTime Lib "fmod.dll" Alias "_FMUSIC_GetTime@4" (ByVal module As Long) As Long

'************
'* Windows Declarations (Added by Adion)
'************
'Required for GetStringFromPointer
Private Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long ' Notice the As Long return value replacing the As String given by the API Viewer.
'Required for the FFT/Spectral functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

'************
'* FUNCTIONS (Added by Adion)
'************
'Usage: myerrorstring = FSOUND_GetErrorString(FSOUND_GetError)
Public Function FSOUND_GetErrorString(ByVal errorcode As Long) As String
    Select Case errorcode
        Case FMOD_ERR_NONE:             FSOUND_GetErrorString = "No errors"
        Case FMOD_ERR_BUSY:             FSOUND_GetErrorString = "Cannot call this command after FSOUND_Init.  Call FSOUND_Close first."
        Case FMOD_ERR_UNINITIALIZED:    FSOUND_GetErrorString = "This command failed because FSOUND_Init was not called"
        Case FMOD_ERR_PLAY:             FSOUND_GetErrorString = "Playing the sound failed."
        Case FMOD_ERR_INIT:             FSOUND_GetErrorString = "Error initializing output device."
        Case FMOD_ERR_ALLOCATED:        FSOUND_GetErrorString = "The output device is already in use and cannot be reused."
        Case FMOD_ERR_OUTPUT_FORMAT:    FSOUND_GetErrorString = "Soundcard does not support the features needed for this soundsystem (16bit stereo output)"
        Case FMOD_ERR_COOPERATIVELEVEL: FSOUND_GetErrorString = "Error setting cooperative level for hardware."
        Case FMOD_ERR_CREATEBUFFER:     FSOUND_GetErrorString = "Error creating hardware sound buffer."
        Case FMOD_ERR_FILE_NOTFOUND:    FSOUND_GetErrorString = "File not found"
        Case FMOD_ERR_FILE_FORMAT:      FSOUND_GetErrorString = "Unknown file format"
        Case FMOD_ERR_FILE_BAD:         FSOUND_GetErrorString = "Error loading file"
        Case FMOD_ERR_MEMORY:           FSOUND_GetErrorString = "Not enough memory "
        Case FMOD_ERR_VERSION:          FSOUND_GetErrorString = "The version number of this file format is not supported"
        Case FMOD_ERR_INVALID_PARAM:    FSOUND_GetErrorString = "An invalid parameter was passed to this function"
        Case FMOD_ERR_NO_EAX:           FSOUND_GetErrorString = "Tried to use an EAX command on a non EAX enabled channel or output."
        Case FMOD_ERR_CHANNEL_ALLOC:    FSOUND_GetErrorString = "Failed to allocate a new channel"
        Case FMOD_ERR_RECORD:           FSOUND_GetErrorString = "Recording is not supported on this machine"
        Case FMOD_ERR_MEDIAPLAYER:      FSOUND_GetErrorString = "Required Mediaplayer codec is not installed"
        Case Else:                      FSOUND_GetErrorString = "Unknown error"
    End Select
End Function

'Thanks to KarLKoX for the following function
'Example: MyDriverName = GetStringFromPointer(FSOUND_GetDriverName(count))
Public Function GetStringFromPointer(ByVal lpString As Long) As String
Dim zpos As Long
Dim s As String

s = String(255, 0)
ConvCStringToVBString s, lpString
' Look for the null char ending the C string
zpos = InStr(s, vbNullChar)
s = Left(s, zpos - 1)
GetStringFromPointer = s
End Function

'These functions are added by Adion
Public Function GetSingleFromPointer(ByVal lpSingle As Long) As Single
'A Single is 4 bytes, so we copy 4 bytes
CopyMemory GetSingleFromPointer, ByVal lpSingle, 4
End Function
'Warning: You should set the fft dsp to active before retreiving the spectrum
'Also make sure the array you pass is dimensioned and ready to use
'FSOUND_DSP_SetActive FSOUND_DSP_GetFFTUnit, 1
Public Function GetSpectrum(ByRef Spectrum() As Single)
Dim nrOfVals As Long, lpSpectrum As Long
Dim a As Long
If UBound(Spectrum) > 511 Then nrOfVals = 512 Else nrOfVals = UBound(Spectrum) + 1
lpSpectrum = FSOUND_DSP_GetSpectrum
CopyMemory Spectrum(0), ByVal lpSpectrum, nrOfVals * 4
End Function
