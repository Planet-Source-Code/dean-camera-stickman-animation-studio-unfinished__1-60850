VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Lines As Collection
Private Name As String

Private Sub Class_Initialize()
    Set Lines = New Collection
End Sub

Private Sub Class_Terminate()
    Set Lines = Nothing

    FrmMain.RefreshObjectsList
End Sub

Property Let ObjectName(NewName As String)
    Name = NewName
    FrmMain.RefreshObjectsList
End Property

Property Get ObjectName() As String
    ObjectName = Name
End Property
