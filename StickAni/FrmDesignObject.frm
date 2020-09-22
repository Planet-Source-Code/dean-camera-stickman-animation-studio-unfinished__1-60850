VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmDesignObject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Design Object"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "FrmDesignObject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   1455
   End
   Begin VB.PictureBox picDesObject 
      Height          =   6255
      Left            =   1920
      ScaleHeight     =   6195
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
   Begin VB.Frame fmeDesOptions 
      Caption         =   "Edit Options"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.OptionButton optCircle 
         Enabled         =   0   'False
         Height          =   495
         Left            =   840
         Picture         =   "FrmDesignObject.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton optLine 
         Enabled         =   0   'False
         Height          =   495
         Left            =   280
         Picture         =   "FrmDesignObject.frx":04F0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtLineWidth 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "1"
         Top             =   820
         Width           =   495
      End
      Begin MSComDlg.CommonDialog cdgDialogs 
         Left            =   840
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdSetLineColour 
         Caption         =   "Set Line Colour"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblLineWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Line Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDesignObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ================================================================================

Public DesObject As ClsObject

Private TempFrame As ClsFrame
Private DesActiveLine As ClsLine
Private SelActiveLine As ClsLine
Private DesActivePoint As Integer

Private IsMoving As Boolean
Private IsResizing As Boolean
Private MouseOffset As tMouseOffset

Private DesDC As Long
Private DesBitmap As Long
Private OldBitmap As Long

' ================================================================================

Private Sub cmdClose_Click()
    Unload Me
End Sub

' ================================================================================

Private Sub Form_Load()
    Set TempFrame = New ClsFrame
    TempFrame.Objects.Add DesObject

    CreateDesDC
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TempFrame = Nothing
    
    DeleteDesDC
End Sub

' ================================================================================

Public Sub CreateDesDC()
    Dim ScreenDC As Long
    
    ScreenDC = GetDC(0)
    
    DesDC = CreateCompatibleDC(ScreenDC)
    DesBitmap = CreateCompatibleBitmap(ScreenDC, 400, 400)
    OldBitmap = SelectObject(DesDC, DesBitmap)

    ReleaseDC 0&, ScreenDC
End Sub

Public Sub DeleteDesDC()
    DeleteDC DesDC
    DeleteObject DesBitmap
End Sub

' ================================================================================

Private Sub picDesObject_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static cObject As ClsObject, cLine As ClsLine
    Static Px As Long, Py As Long

    IsMoving = Shift

    Px = x / Screen.TwipsPerPixelX
    Py = y / Screen.TwipsPerPixelY

    With MouseOffset
        .MovingOffsetX = Px
        .MovingOffsetX = Py
        .MovingPrevOffsetX = 0
        .MovingPrevOffsetY = 0
    End With
    
    Set SelActiveLine = Nothing

    For Each cObject In ActiveFrame.Objects
        For Each cLine In cObject.Lines
            If (Px > cLine.PointStartX - 3) And (Px < cLine.PointStartX + 3) And _
                (Py > cLine.PointStartY - 3) And (Py < cLine.PointStartY + 3) Then
                Set DesActiveLine = cLine
                Set SelActiveLine = cLine
                DesActivePoint = POINTSTART
                
                Exit For
            ElseIf (Px > cLine.PointEndX - 3) And (Px < cLine.PointEndX + 3) And _
                    (Py > cLine.PointEndY - 3) And (Py < cLine.PointEndY + 3) Then
                Set DesActiveLine = cLine
                Set SelActiveLine = cLine
                DesActivePoint = POINTEND
                
                Exit For
            End If
        Next

        Set cLine = Nothing
    Next

    If Not DesActiveLine Is Nothing Then
        cmdSetLineColour.Enabled = True
        txtLineWidth.Enabled = True
        txtLineWidth.Text = DesActiveLine.LineWidth
        optLine.Enabled = True
        optCircle.Enabled = True
                
        If DesActiveLine.IsCircle Then
            optCircle.Value = True
        Else
            optLine.Value = True
        End If
                
        If Not IsMoving Then IsResizing = True
    End If

    Set cObject = Nothing
End Sub

Private Sub picDesObject_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static Xloc As Long, Yloc As Long
    Static ConnectData As ClsConnect, cLine As ClsLine

    Xloc = x / Screen.TwipsPerPixelX
    Yloc = y / Screen.TwipsPerPixelY

    If SelActiveLine Is Nothing Then Exit Sub

    If Not IsMoving Then
        If IsResizing Then
            If DesActivePoint = POINTSTART Then
                DesActiveLine.PointStartX = Xloc
                DesActiveLine.PointStartY = Yloc

                For Each ConnectData In DesActiveLine.Connects
                    If ConnectData.Start = POINTEND Then
                        Set cLine = DesObject.Lines(ConnectData.LineIndex)
                        cLine.PointEndX = Xloc
                        cLine.PointEndY = Yloc
                        Set cLine = Nothing
                    End If
                Next
            
                Set ConnectData = Nothing
            Else
                DesActiveLine.PointEndX = Xloc
                DesActiveLine.PointEndY = Yloc

                For Each ConnectData In DesActiveLine.Connects
                    If ConnectData.Start = POINTSTART Then
                        Set cLine = DesObject.Lines(ConnectData.LineIndex)
                        cLine.PointStartX = Xloc
                        cLine.PointStartY = Yloc
                        Set cLine = Nothing
                    End If
                Next
            End If
        End If
    Else
        If MouseOffset.MovingPrevOffsetX Then
                For Each cLine In DesObject.Lines
                    cLine.PointStartX = cLine.PointStartX + (MouseOffset.MovingPrevOffsetX - (MouseOffset.MovingOffsetX - Xloc))
                    cLine.PointStartY = cLine.PointStartY + (MouseOffset.MovingPrevOffsetY - (MouseOffset.MovingOffsetY - Yloc))

                    cLine.PointEndX = cLine.PointEndX + (MouseOffset.MovingPrevOffsetX - (MouseOffset.MovingOffsetX - Xloc))
                    cLine.PointEndY = cLine.PointEndY + (MouseOffset.MovingPrevOffsetY - (MouseOffset.MovingOffsetY - Yloc))
                Next

                Set cLine = Nothing
        End If

        MouseOffset.MovingPrevOffsetX = MouseOffset.MovingOffsetX - Xloc
        MouseOffset.MovingPrevOffsetY = MouseOffset.MovingOffsetY - Yloc
    End If

    picDesObject_Paint
End Sub

Private Sub picDesObject_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsMoving = False
    IsResizing = False
    
    picDesObject_Paint
End Sub

Private Sub picDesObject_Paint()
    RenderFrame TempFrame, DesDC, True, False, True, False, True, DesActiveLine
    BitBlt picDesObject.hdc, 0, 0, 400, 400, DesDC, 0, 0, vbSrcCopy
    
    FrmMain.RefreshFrame
End Sub

' ================================================================================

Private Sub optCircle_Click()
    If Not DesActiveLine Is Nothing Then
        DesActiveLine.IsCircle = True
        
        picDesObject_Paint
    End If
End Sub

Private Sub optLine_Click()
    If Not DesActiveLine Is Nothing Then
        DesActiveLine.IsCircle = False
        
        picDesObject_Paint
    End If
End Sub

Private Sub cmdSetLineColour_Click()
    If Not DesActiveLine Is Nothing Then
        cdgDialogs.ShowColor
        DesActiveLine.LineColour = cdgDialogs.Color
        
        picDesObject_Paint
    End If
End Sub

Private Sub txtLineWidth_Change()
    If txtLineWidth.Text = vbNullString Then txtLineWidth.Text = "1"
    If Int(txtLineWidth.Text) = 0 Then txtLineWidth.Text = "1"
    
    DesActiveLine.LineWidth = Int(txtLineWidth.Text)
    picDesObject_Paint
End Sub
