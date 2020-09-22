VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PointStartX As Long
Public PointStartY As Long

Public PointEndX As Long
Public PointEndY As Long

Public LineWidth As Long
Public LineColour As Long

Public IsCircle As Boolean

Public Connects As Collection

Public Sub Connect(Start As Boolean, LineIndex As Long)
    Dim ConnectData As New ClsConnect

    ConnectData.Start = Start
    ConnectData.LineIndex = LineIndex

    Connects.Add ConnectData
    
    Set ConnectData = Nothing
End Sub

Private Sub Class_Initialize()
    Set Connects = New Collection
End Sub

Private Sub Class_Terminate()
    Set Connects = Nothing
End Sub
