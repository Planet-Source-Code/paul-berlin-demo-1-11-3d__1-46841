VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Used to measure time

Private lStartTime As Long

Public Sub StartWatch()
  lStartTime = timeGetTime
End Sub

Public Function GetTime() As Long
  GetTime = timeGetTime - lStartTime
End Function

Public Function SetStartTime(ByVal lTime As Long)
  'ex: setstarttime gettime-20000
  lStartTime = lTime
End Function
