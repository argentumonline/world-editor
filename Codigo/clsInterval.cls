VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetTickCount Lib "kernel32" () As Long

Private interval As Long
Private time As Long
Private can As Boolean
Public Sub Init(Inter As Long)
    interval = Inter
    time = GetTickCount + interval
End Sub
Public Function ICan() As Boolean
    If GetTickCount <= interval Then
        time = GetTickCount + interval
    End If
    If GetTickCount >= time Then
        can = True
        time = GetTickCount + interval
    End If
    
    ICan = can
    can = False

End Function

