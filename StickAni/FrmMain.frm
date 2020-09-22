VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Stick-Figure Animation Studio"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9795
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeAnimation 
      Caption         =   "Animation"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   6360
      Width           =   2175
      Begin VB.HScrollBar sclAniSpeed 
         Height          =   255
         Left            =   120
         Max             =   300
         Min             =   15
         TabIndex        =   20
         Top             =   640
         Value           =   100
         Width           =   1935
      End
      Begin VB.CommandButton cmdAnimation 
         Caption         =   "Play"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblCurrFrame 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Frame 1 of 1"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   925
         Width           =   1935
      End
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   7560
   End
   Begin MSComDlg.CommonDialog cdgDialogs 
      Left            =   1080
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar sclVertPos 
      Height          =   6495
      Left            =   9480
      TabIndex        =   16
      Top             =   1440
      Width           =   255
   End
   Begin VB.HScrollBar sclHozPos 
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   7920
      Width           =   7095
   End
   Begin VB.CommandButton cmdDeleteFrame 
      Caption         =   "Delete Frame"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdNewFrame 
      Caption         =   "New Frame"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5160
      Width           =   1935
   End
   Begin StickmanAniStudio.CtlThumbs ctlThumbs 
      Height          =   1270
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2143
   End
   Begin VB.Frame fmeObjects 
      Caption         =   "Objects"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
      Begin VB.CommandButton cmdEditObject 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton cmdDeleteObject 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdCenterObject 
         Caption         =   "Center"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   975
      End
      Begin VB.PictureBox picExtraAttributes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1065
         ScaleWidth      =   1905
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
         Begin VB.Label lblMoving 
            BackStyle       =   0  'Transparent
            Caption         =   "Moving"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   750
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblObjectVisible 
            BackStyle       =   0  'Transparent
            Caption         =   "N/A"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   750
            Width           =   735
         End
         Begin VB.Label lblObjectName 
            BackStyle       =   0  'Transparent
            Caption         =   "No Object Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   30
            TabIndex        =   7
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblSelectedObjectCpnts 
            BackStyle       =   0  'Transparent
            Caption         =   "Components: N/A"
            Height          =   255
            Left            =   30
            TabIndex        =   6
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdNewObjects 
         Caption         =   "Add New Object"
         Height          =   375
         Left            =   120
         Picture         =   "FrmMain.frx":0442
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbObjects 
         Height          =   315
         ItemData        =   "FrmMain.frx":0884
         Left            =   120
         List            =   "FrmMain.frx":0886
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "cmbObjects"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblChoseObject 
         Caption         =   "Selected Object:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.PictureBox picCurrFrame 
      BackColor       =   &H00808080&
      Height          =   6495
      Left            =   2400
      ScaleHeight     =   6435
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuSetWrkSpceSze 
         Caption         =   "Set &Workspace Size"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'                   STICKMAN ANIMATION STUDIO
'                     (C) Dean Camera, 2005

' ================================================================================

Private TempDC As Long
Private TempBitmap As Long

Private AniFrame As Long
    
Private IsMoving As Boolean
Private MouseOffset As tMouseOffset

' ================================================================================

Private Sub cmbObjects_Change()
    Dim cObject As ClsObject

    For Each cObject In ActiveFrame.Objects
        If cObject.ObjectName = cmbObjects.Text Then
            Set ActiveObject = cObject
            Exit For
        End If
    Next

    Set cObject = Nothing

    UpdateSelectedObjectStats
End Sub

Private Sub cmbObjects_Click()
    Dim cObject As ClsObject

    For Each cObject In ActiveFrame.Objects
        If cObject.ObjectName = cmbObjects.Text Then
            Set ActiveObject = cObject
            Exit For
        End If
    Next

    Set cObject = Nothing

    UpdateSelectedObjectStats
End Sub

Private Sub cmdCenterObject_Click()
    Dim cLine As ClsLine
    Dim MaxBounds As RECT

    If ActiveObject Is Nothing Then Exit Sub

    For Each cLine In ActiveObject.Lines
        If MaxBounds.Left > cLine.PointStartX Then MaxBounds.Left = cLine.PointStartX
        If MaxBounds.Left > cLine.PointEndX Then MaxBounds.Left = cLine.PointEndX

        If MaxBounds.Right < cLine.PointStartX Then MaxBounds.Right = cLine.PointStartX
        If MaxBounds.Right < cLine.PointEndX Then MaxBounds.Right = cLine.PointEndX

        If MaxBounds.Bottom < cLine.PointStartY Then MaxBounds.Bottom = cLine.PointStartY
        If MaxBounds.Bottom < cLine.PointEndY Then MaxBounds.Bottom = cLine.PointEndY
    Next
        
    MaxBounds.Left = (WorkArea(0) / 2) - (MaxBounds.Right - MaxBounds.Left)
    MaxBounds.Top = (WorkArea(1) / 2) - MaxBounds.Bottom + 100
    
    For Each cLine In ActiveObject.Lines
        cLine.PointStartX = cLine.PointStartX + MaxBounds.Left
        cLine.PointEndX = cLine.PointEndX + MaxBounds.Left

        cLine.PointStartY = cLine.PointStartY + MaxBounds.Top
        cLine.PointEndY = cLine.PointEndY + MaxBounds.Top
    Next

    Set cLine = Nothing

    ctlThumbs.RefreshThumbs
    UpdateSelectedObjectStats
    RefreshFrame
End Sub

Private Sub cmdDeleteFrame_Click()
    Dim cFrame As ClsFrame
    Dim fIndex As Integer
    Dim PrevFrameIndex As Integer

    If ActiveFrame Is Nothing Then Exit Sub

    For fIndex = ActiveFrame.Index + 1 To Frames.Count
        Set cFrame = Frames(fIndex)
        cFrame.Index = cFrame.Index - 1
        Set cFrame = Nothing
    Next

    PrevFrameIndex = ActiveFrame.Index - 1
    Frames.Remove ActiveFrame.Index

    If PrevFrameIndex Then
        Set ActiveFrame = Frames(PrevFrameIndex)
    Else
        If Frames.Count Then
            Set ActiveFrame = Frames(Frames.Count)
        Else
            Set ActiveFrame = Nothing
            cmdDeleteFrame.Enabled = False
            fmeObjects.Enabled = False
            lblCurrFrame.Caption = "Frame 0 of 0"
        End If
    End If

    RefreshObjectsList
    ctlThumbs.RefreshThumbs
    RefreshFrame
End Sub

Private Sub cmdDeleteObject_Click()
    Dim lCount As Long
    Dim cObject As ClsObject

    For lCount = 1 To ActiveFrame.Objects.Count
        Set cObject = ActiveFrame.Objects(lCount)
        If cObject Is ActiveObject Then
            If ActiveObject Is cObject Then Set ActiveObject = Nothing
            ActiveFrame.Objects.Remove lCount
            Exit For
        End If
    Next

    Set cObject = Nothing

    ctlThumbs.RefreshThumbs
    RefreshObjectsList
    UpdateSelectedObjectStats
    RefreshFrame
End Sub

Private Sub cmdEditObject_Click()
    Set frmDesignObject.DesObject = ActiveObject
    
    frmDesignObject.Caption = "Design Object - " & ActiveObject.ObjectName
    frmDesignObject.Show 1, Me
End Sub

Private Sub cmdNewFrame_Click()
    Dim cFrame As ClsFrame

    If ActiveFrame Is Nothing Then
        Set cFrame = New ClsFrame

        cFrame.Index = 1
        Frames.Add cFrame

        fmeObjects.Enabled = True
        cmdDeleteFrame.Enabled = True
        
        RefreshObjectsList
    Else
        For Each cFrame In Frames
            If cFrame.Index > ActiveFrame.Index Then cFrame.Index = cFrame.Index + 1
        Next
        
        CopyFrame ActiveFrame, cFrame
        Frames.Add cFrame
    End If

    Set ActiveFrame = cFrame
        
    Me.RefreshFrame
    Me.ctlThumbs.RefreshThumbs
        
    Set cFrame = Nothing
End Sub

Private Sub cmdNewObjects_Click()
    With cdgDialogs
        .Filter = "Stickman Ani Studio Object (*.sao)|*.sao"
        .FileName = vbNullString
        .ShowOpen

        If .FileName <> vbNullString Then LoadObject .FileName
    End With
End Sub

Private Sub cmdAnimation_Click()
    If tmrAnimation.Enabled = False Then
        fmeObjects.Enabled = False
        cmdNewFrame.Enabled = False
        cmdDeleteFrame.Enabled = False
        picCurrFrame.Enabled = False
        cmbObjects.Enabled = False

        cmdAnimation.Caption = "Stop"

        tmrAnimation.Enabled = True
    Else
        AniFrame = Frames.Count + 1
        tmrAnimation_Timer
    End If
End Sub

' ================================================================================

Private Sub Form_Load()
    Dim cFrame As New ClsFrame

    ScaleMode = vbPixels
    AutoRedraw = True

    Me.lblObjectVisible.Caption = vbNullString

    WorkArea(0) = 400
    WorkArea(1) = 400

    CreateWorkAreaDC
    ctlThumbs.CreateThumbDC

    Set Frames = New Collection

    cFrame.Index = 1
    Frames.Add cFrame
    Set ActiveFrame = cFrame

    Set cFrame = Nothing

    ctlThumbs.RefreshThumbs
    RefreshObjectsList
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If Me.Width <= 8400 Then Me.Width = 8400
    If Me.Height <= 8000 Then Me.Height = 8000

    Me.ScaleMode = vbTwips

    picCurrFrame.Width = Me.Width - picCurrFrame.Left - 400
    picCurrFrame.Height = Me.Height - picCurrFrame.Top - 990

    sclHozPos.Width = Me.Width - sclHozPos.Left - 360
    sclHozPos.Top = Me.Height - 990

    sclVertPos.Height = Me.Height - sclVertPos.Top - 1010
    sclVertPos.Left = Me.Width - 400

    ctlThumbs.Width = Me.Width - 300

    Me.ScaleMode = vbPixels

    DeleteWorkAreaDC
    CreateWorkAreaDC
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteWorkAreaDC

    Set Frames = Nothing

    Set ActiveFrame = Nothing
    Set ActiveObject = Nothing
    Set ActiveLine = Nothing
End Sub

' ================================================================================

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSetWrkSpceSze_Click()
    FrmSetWorkspaceSize.Show 1, Me

    UpdateSelectedObjectStats
End Sub

' ================================================================================

Private Sub picCurrFrame_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveObject Is Nothing Then Exit Sub

    lblMoving.Visible = Shift
End Sub

Private Sub picCurrFrame_Keyup(KeyCode As Integer, Shift As Integer)
    lblMoving.Visible = False
End Sub

Private Sub PicCurrFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static Xloc As Long, Yloc As Long
    Static ConnectData As ClsConnect, cLine As ClsLine

    Xloc = (x / Screen.TwipsPerPixelX) + sclHozPos.Value
    Yloc = (y / Screen.TwipsPerPixelY) + sclVertPos.Value

    If Not IsMoving Then
        If ActiveLine Is Nothing Then Exit Sub

        If ActivePoint = POINTSTART Then
            ActiveLine.PointStartX = Xloc
            ActiveLine.PointStartY = Yloc

            For Each ConnectData In ActiveLine.Connects
                If ConnectData.Start = POINTEND Then
                    Set cLine = ActiveObject.Lines(ConnectData.LineIndex)
                    cLine.PointEndX = Xloc
                    cLine.PointEndY = Yloc
                    Set cLine = Nothing
                End If
            Next

            Set ConnectData = Nothing
        Else
            ActiveLine.PointEndX = Xloc
            ActiveLine.PointEndY = Yloc

            For Each ConnectData In ActiveLine.Connects
                If ConnectData.Start = POINTSTART Then
                    Set cLine = ActiveObject.Lines(ConnectData.LineIndex)
                    cLine.PointStartX = Xloc
                    cLine.PointStartY = Yloc
                    Set cLine = Nothing
                End If
            Next
        End If
    Else
        If MouseOffset.MovingPrevOffsetX Then
            If Not ActiveObject Is Nothing Then
                For Each cLine In ActiveObject.Lines
                    cLine.PointStartX = cLine.PointStartX + (MouseOffset.MovingPrevOffsetX - (MouseOffset.MovingOffsetX - Xloc))
                    cLine.PointStartY = cLine.PointStartY + (MouseOffset.MovingPrevOffsetY - (MouseOffset.MovingOffsetY - Yloc))

                    cLine.PointEndX = cLine.PointEndX + (MouseOffset.MovingPrevOffsetX - (MouseOffset.MovingOffsetX - Xloc))
                    cLine.PointEndY = cLine.PointEndY + (MouseOffset.MovingPrevOffsetY - (MouseOffset.MovingOffsetY - Yloc))
                Next

                Set cLine = Nothing
            End If
        End If

        MouseOffset.MovingPrevOffsetX = MouseOffset.MovingOffsetX - Xloc
        MouseOffset.MovingPrevOffsetY = MouseOffset.MovingOffsetY - Yloc
    End If

    RefreshFrame
End Sub

Private Sub PicCurrFrame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static cObject As ClsObject, cLine As ClsLine
    Static Px As Long, Py As Long

    If ActiveFrame Is Nothing Then Exit Sub

    If x / Screen.TwipsPerPixelX > WorkArea(0) Or y / Screen.TwipsPerPixelY > WorkArea(1) Then Exit Sub

    Set ActiveLine = Nothing

    IsMoving = Shift

    Px = (x / Screen.TwipsPerPixelX) + sclHozPos.Value
    Py = (y / Screen.TwipsPerPixelY) + sclVertPos.Value

    With MouseOffset
        .MovingOffsetX = Px
        .MovingOffsetX = Py
        .MovingPrevOffsetX = 0
        .MovingPrevOffsetY = 0
    End With

    For Each cObject In ActiveFrame.Objects
        For Each cLine In cObject.Lines
            If (Px > cLine.PointStartX - 3) And (Px < cLine.PointStartX + 3) And _
                (Py > cLine.PointStartY - 3) And (Py < cLine.PointStartY + 3) Then
                Set ActiveObject = cObject
                Set ActiveLine = cLine
                ActivePoint = POINTSTART
                Exit For
            ElseIf (Px > cLine.PointEndX - 3) And (Px < cLine.PointEndX + 3) And _
                    (Py > cLine.PointEndY - 3) And (Py < cLine.PointEndY + 3) Then
                Set ActiveObject = cObject
                Set ActiveLine = cLine
                ActivePoint = POINTEND
                Exit For
            End If
        Next

        Set cLine = Nothing
    Next

    Set cObject = Nothing

    UpdateSelectedObjectStats
End Sub

Private Sub PicCurrFrame_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsMoving = False

    UpdateSelectedObjectStats
    RefreshFrame

    ctlThumbs.RefreshThumbs

    Set ActiveLine = Nothing
End Sub

Private Sub PicCurrFrame_Paint()
    Dim NoClearBG As Boolean
    Static cFrame As ClsFrame

    If Not ActiveFrame Is Nothing Then
        If ActiveFrame.Index > 1 Then
            For Each cFrame In Frames
                If cFrame.Index = ActiveFrame.Index - 1 Then
                    RenderFrame cFrame, TempDC, False, False, True, True, False
                    NoClearBG = True
                    Exit For
                End If
            Next
        End If
    End If

    RenderFrame ActiveFrame, TempDC, True, IsMoving, (Not NoClearBG), False, False

    BitBlt picCurrFrame.hdc, 0 - sclHozPos.Value, 0 - sclVertPos.Value, WorkArea(0), WorkArea(1), TempDC, 0, 0, vbSrcCopy
End Sub

' ================================================================================

Public Sub CreateWorkAreaDC()
    Dim ScreenDC As Long
    
    ScreenDC = GetDC(0)

    TempDC = CreateCompatibleDC(ScreenDC)
    TempBitmap = CreateCompatibleBitmap(ScreenDC, WorkArea(0), WorkArea(1))
    
    ReleaseDC 0&, ScreenDC
    
    SelectObject TempDC, TempBitmap
    SetBkMode TempDC, 1

    If (picCurrFrame.Height < 35) Or (picCurrFrame.Width < 35) Then
        sclVertPos.Visible = False
        sclHozPos.Visible = False

        picCurrFrame.Width = Me.Width - picCurrFrame.Left - 100
        picCurrFrame.Height = Me.Height - picCurrFrame.Top - 600
    Else
        sclVertPos.Visible = True
        sclHozPos.Visible = True
    End If

    If (picCurrFrame.Height < WorkArea(1)) Then
        sclVertPos.Enabled = True
        sclVertPos.Max = WorkArea(1) - picCurrFrame.Height
    Else
        sclVertPos.Enabled = False
    End If

    If (picCurrFrame.Width < WorkArea(0)) Then
        sclHozPos.Enabled = True
        sclHozPos.Max = WorkArea(0) - picCurrFrame.Width
    Else
        sclHozPos.Enabled = False
    End If

    RefreshFrame
End Sub

Public Sub DeleteWorkAreaDC()
    DeleteDC TempDC
    DeleteObject TempBitmap
End Sub

' ================================================================================

Public Sub UpdateSelectedObjectStats()
    Dim cLine As ClsLine
    Dim Onscreen As RECT
    Dim OffscreenPicIndex As Integer

    If ActiveObject Is Nothing Then
        cmbObjects.Text = vbNullString
        lblObjectName.Caption = "No Object Selected"
        lblSelectedObjectCpnts.Caption = "Components: N/A"
        lblObjectVisible.Caption = vbNullString
        cmdDeleteObject.Enabled = False
        cmdCenterObject.Enabled = False
        cmdEditObject.Enabled = False
        
        picExtraAttributes.Cls
    Else
        cmbObjects.Text = ActiveObject.ObjectName
        lblObjectName.Caption = ActiveObject.ObjectName
        lblSelectedObjectCpnts.Caption = "Components: " & ActiveObject.Lines.Count

        cmdDeleteObject.Enabled = True
        cmdCenterObject.Enabled = True
        cmdEditObject.Enabled = True
        
        For Each cLine In ActiveObject.Lines
            If cLine.PointStartX > 0 Then Onscreen.Left = True
            If cLine.PointStartX < WorkArea(0) Then Onscreen.Right = True
            If cLine.PointStartY > 0 Then Onscreen.Top = True
            If cLine.PointStartY < WorkArea(1) Then Onscreen.Bottom = True

            If cLine.PointEndX > 0 Then Onscreen.Left = True
            If cLine.PointEndX < WorkArea(0) Then Onscreen.Right = True
            If cLine.PointEndY > 0 Then Onscreen.Top = True
            If cLine.PointEndY < WorkArea(1) Then Onscreen.Bottom = True
        Next

        Set cLine = Nothing

        lblObjectVisible.Caption = "Offscreen"

        If Onscreen.Left = False And Onscreen.Top = False Then
            OffscreenPicIndex = 5
        ElseIf Onscreen.Top = False And Onscreen.Right = False Then
            OffscreenPicIndex = 6
        ElseIf Onscreen.Right = False And Onscreen.Bottom = False Then
            OffscreenPicIndex = 7
        ElseIf Onscreen.Bottom = False And Onscreen.Left = False Then
            OffscreenPicIndex = 8
        ElseIf Onscreen.Top = False Then
            OffscreenPicIndex = 1
        ElseIf Onscreen.Bottom = False Then
            OffscreenPicIndex = 2
        ElseIf Onscreen.Left = False Then
            OffscreenPicIndex = 3
        ElseIf Onscreen.Right = False Then
            OffscreenPicIndex = 4
        Else
            lblObjectVisible.Caption = "Visible"
            OffscreenPicIndex = 9
        End If

        picExtraAttributes.PaintPicture LoadPicture(App.Path & "\Res\OffScreen.bmp"), 35, 770, 160, 160, 150 * (OffscreenPicIndex - 1), 0, 160, 160
    End If
End Sub

Public Sub RefreshObjectsList()
    Dim cObj As ClsObject

    cmbObjects.Clear

    If ActiveFrame Is Nothing Then Exit Sub

    For Each cObj In ActiveFrame.Objects
        cmbObjects.AddItem cObj.ObjectName
    Next

    Set cObj = Nothing
End Sub

Public Sub RefreshFrame()
    If Not ActiveFrame Is Nothing Then lblCurrFrame.Caption = "Frame " & ActiveFrame.Index & " of " & Frames.Count
    PicCurrFrame_Paint
End Sub

Private Sub sclAniSpeed_Change()
    tmrAnimation.Interval = sclAniSpeed.Value
End Sub

' ================================================================================

Private Sub sclHozPos_Change()
    PicCurrFrame_Paint
End Sub

Private Sub sclVertPos_Change()
    PicCurrFrame_Paint
End Sub

' ================================================================================

Private Sub tmrAnimation_Timer()
    lblCurrFrame.Caption = "Frame " & AniFrame + 1 & " of " & Frames.Count

    AniFrame = AniFrame + 1

    If AniFrame <= Frames.Count Then
        RenderFrame Frames(AniFrame), TempDC, False, False, True, False, False
        BitBlt picCurrFrame.hdc, 0 - sclHozPos.Value, 0 - sclVertPos.Value, WorkArea(0), WorkArea(1), TempDC, 0, 0, vbSrcCopy
    Else
        fmeObjects.Enabled = True
        cmdNewFrame.Enabled = True
        cmdDeleteFrame.Enabled = True
        picCurrFrame.Enabled = True
        cmbObjects.Enabled = True
        
        cmdAnimation.Caption = "Play"

        tmrAnimation.Enabled = False
        AniFrame = 0
        RefreshFrame
    End If
End Sub
