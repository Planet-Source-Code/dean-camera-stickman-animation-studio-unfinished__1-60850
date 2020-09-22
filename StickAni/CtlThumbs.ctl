VERSION 5.00
Begin VB.UserControl CtlThumbs 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1215
   ScaleWidth      =   4800
   Begin VB.HScrollBar sclThumbScroll 
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "CtlThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TempDC As Long
Private TempBitmap As Long
Private TempOldBitmap As Long

Private ThumbDC As Long
Private ThumbBitmap As Long
Private OldBitmap As Long

Private Sub sclThumbScroll_Change()
    RefreshThumbs
End Sub

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels

    CreateTempDC
End Sub

Public Sub CreateThumbDC()
    Dim ScreenDC As Long
    
    ScreenDC = GetDC(0)
    
    ThumbDC = CreateCompatibleDC(ScreenDC)
    ThumbBitmap = CreateCompatibleBitmap(ScreenDC, WorkArea(0), WorkArea(1))
    OldBitmap = SelectObject(ThumbDC, ThumbBitmap)

    ReleaseDC 0&, ScreenDC
End Sub

Private Sub CreateTempDC()
    Dim ScreenDC As Long
    
    ScreenDC = GetDC(0)
    
    TempDC = CreateCompatibleDC(ScreenDC)

    If Frames Is Nothing Then
        TempBitmap = CreateCompatibleBitmap(ScreenDC, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY)
    Else
        TempBitmap = CreateCompatibleBitmap(ScreenDC, Frames.Count * 60 + 100, UserControl.Height)
    End If
    
    ReleaseDC 0&, ScreenDC

    TempOldBitmap = SelectObject(TempDC, TempBitmap)
    SetBkMode TempDC, 1
End Sub

Public Sub DeleteThumbDC()
    SelectObject ThumbDC, OldBitmap
    DeleteObject ThumbBitmap
    DeleteDC ThumbDC
End Sub

Private Sub DeleteTempDC()
    SelectObject TempDC, TempOldBitmap
    DeleteObject TempBitmap
    DeleteDC TempDC
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static cFrame As ClsFrame

    If Button <> 1 Or Y < 10 Or Y > 60 Then Exit Sub

    For Each cFrame In Frames
        If (X >= (cFrame.Index - 1) * 60 + 10) - IIf(sclThumbScroll.Enabled, 0 - sclThumbScroll.Value, 0) And X < (((cFrame.Index - 1) * 60) + 60) - IIf(sclThumbScroll.Enabled, 0 - sclThumbScroll.Value, 0) Then
            Set ActiveFrame = cFrame
            Set ActiveObject = Nothing
            Set ActiveLine = Nothing

            FrmMain.cmbObjects.Text = vbNullString

            Exit For
        End If
    Next

    Set cFrame = Nothing

    RefreshThumbs
    FrmMain.RefreshObjectsList
    FrmMain.UpdateSelectedObjectStats
    FrmMain.RefreshFrame
End Sub

Private Sub UserControl_Paint()
    RefreshThumbs
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = vbTwips
    sclThumbScroll.Width = UserControl.Width - 60
    UserControl.ScaleMode = vbPixels

    RefreshThumbs
End Sub

Private Sub UserControl_Terminate()
    DeleteThumbDC
    DeleteTempDC
End Sub

Public Sub RefreshThumbs()
    Static cFrame As ClsFrame, cObject As ClsObject, cLine As ClsLine
    Static nBrush As Long, sObjHdl As Long
    Static rPRect As RECT
    Static TextFont As IFont
    Dim ThumbIndex As Long

    If Frames Is Nothing Then Exit Sub

    DeleteTempDC
    CreateTempDC

    Set TextFont = New StdFont
    TextFont.Name = "Small Fonts"
    TextFont.Size = 5

    nBrush = CreateSolidBrush(&HE0E0E0)
    SetRect rPRect, 0, 0, Frames.Count * 60 + 100, UserControl.Height / Screen.TwipsPerPixelY
    FillRect TempDC, rPRect, nBrush
    DeleteObject nBrush

    For Each cFrame In Frames
        ThumbIndex = cFrame.Index - 1

        RenderFrame cFrame, ThumbDC, False, False, True, False, False
        StretchBlt TempDC, 60 * ThumbIndex + 10, 10, 50, 50, ThumbDC, 0, 0, WorkArea(0), WorkArea(1), vbSrcCopy

        If Not ActiveFrame Is Nothing Then
            If ActiveFrame Is cFrame Then
                nBrush = CreateSolidBrush(vbRed)
            Else
                nBrush = CreateSolidBrush(vbBlack)
            End If
        Else
            nBrush = CreateSolidBrush(vbBlack)
        End If

        SetRect rPRect, 60 * ThumbIndex + 10, 10, 60 * ThumbIndex + 60, 60
        FrameRect TempDC, rPRect, nBrush
        DeleteObject nBrush

        sObjHdl = SelectObject(TempDC, TextFont.hFont)
        SetRect rPRect, 60 * ThumbIndex + 10, 0, 60 * ThumbIndex + 60, 10
        DrawText TempDC, ThumbIndex + 1, Len(Str(ThumbIndex)) - 1, rPRect, DT_CENTER
        SelectObject TempDC, sObjHdl
    Next

    If ((ThumbIndex + 1) * 60) - 10 > UserControl.Width / Screen.TwipsPerPixelX Then
        sclThumbScroll.Enabled = True
        sclThumbScroll.Max = (UserControl.Width / Screen.TwipsPerPixelX) - (ThumbIndex * 60) - 10
    Else
        sclThumbScroll.Enabled = False
    End If

    BitBlt hdc, 0, 0, UserControl.Width, UserControl.Height, TempDC, IIf(sclThumbScroll.Enabled, 0 - sclThumbScroll.Value, 0), 0, vbSrcCopy

    Set TextFont = Nothing
    Set cFrame = Nothing
    Set cObject = Nothing
    Set cLine = Nothing
End Sub
