Attribute VB_Name = "ModMainDeclares"
Option Explicit

Public WorkArea(1) As Long
Public Frames As Collection

Public ActiveFrame As ClsFrame
Public ActiveObject As ClsObject
Public ActiveLine As ClsLine
Public ActivePoint As Boolean

Public TempPoint As POINTAPI

Public Const POINTSELECTED = vbRed
Public Const POINTMOVE = &H800080
Public Const POINTCOLOUR = vbBlue
Public Const POINTSTART = True
Public Const POINTEND = False

Public Type tScrollOffset
    OffsetX As Long
    OffsetY As Long
End Type

Public Type tMouseOffset
    MovingOffsetX As Long
    MovingOffsetY As Long
    MovingPrevOffsetX As Long
    MovingPrevOffsetY As Long
End Type

Public Sub RenderFrame(Frame As ClsFrame, DC As Long, ShowHandles As Boolean, IsMoving As Boolean, ClearBG As Boolean, Ghost As Boolean, IsDesigning As Boolean, Optional HLActiveLine As ClsLine)
    Static cObject As ClsObject, cLine As ClsLine
    Static nBrush As Long, nPen As Long, sObjHdl As Long
    Static rPRect As RECT

    If Frame Is Nothing Then
        nBrush = CreateSolidBrush(&H808080)
        SetRect rPRect, 0, 0, WorkArea(0), WorkArea(1)
        FillRect DC, rPRect, nBrush
        DeleteObject nBrush
        Exit Sub
    End If

    If ClearBG Then
        nBrush = CreateSolidBrush(vbWhite)
        SetRect rPRect, 0, 0, WorkArea(0), WorkArea(1)
        FillRect DC, rPRect, nBrush
        DeleteObject nBrush
    End If

    For Each cObject In Frame.Objects
        If IsDesigning Then
                nBrush = CreateSolidBrush(POINTCOLOUR)
        ElseIf ShowHandles Then
            If Not ActiveObject Is Nothing Then
                If ActiveObject Is cObject Then
                    nBrush = CreateSolidBrush(IIf(IsMoving, POINTMOVE, POINTSELECTED))
                Else
                    nBrush = CreateSolidBrush(POINTCOLOUR)
                End If
            Else
                nBrush = CreateSolidBrush(POINTCOLOUR)
            End If
        End If

        For Each cLine In cObject.Lines
            MoveToEx DC, cLine.PointStartX, cLine.PointStartY, TempPoint
            nPen = CreatePen(PS_SOLID, cLine.LineWidth, IIf(Ghost, &HAAAAAA, cLine.LineColour))
            sObjHdl = SelectObject(DC, nPen)
            
            If cLine.IsCircle Then
                Ellipse DC, cLine.PointStartX + 2 * (cLine.PointEndX - cLine.PointStartX), cLine.PointStartY, cLine.PointEndX - (cLine.PointEndX - cLine.PointStartX), cLine.PointEndY
            Else
                LineTo DC, cLine.PointEndX, cLine.PointEndY
            End If

            If ShowHandles Then
                SetRect rPRect, cLine.PointStartX - 3, cLine.PointStartY - 3, cLine.PointStartX + 3, cLine.PointStartY + 3
                FillRect DC, rPRect, nBrush
                SetRect rPRect, cLine.PointEndX - 3, cLine.PointEndY - 3, cLine.PointEndX + 3, cLine.PointEndY + 3
                FillRect DC, rPRect, nBrush
            End If

            SelectObject DC, sObjHdl
            DeleteObject nPen
        Next

        If IsDesigning And Not HLActiveLine Is Nothing Then
            DeleteObject nBrush
            
            nBrush = CreateSolidBrush(POINTSELECTED)
            
            SetRect rPRect, HLActiveLine.PointStartX - 3, HLActiveLine.PointStartY - 3, HLActiveLine.PointStartX + 3, HLActiveLine.PointStartY + 3
            FillRect DC, rPRect, nBrush
            SetRect rPRect, HLActiveLine.PointEndX - 3, HLActiveLine.PointEndY - 3, HLActiveLine.PointEndX + 3, HLActiveLine.PointEndY + 3
            FillRect DC, rPRect, nBrush
        End If
        
        DeleteObject nBrush
    Next

    Set cObject = Nothing
    Set cLine = Nothing
End Sub

Public Sub CopyFrame(SourceFrame As ClsFrame, DestFrame As ClsFrame)
    Dim cObj As ClsObject
    Dim sObj As ClsObject
    Dim cLine As ClsLine
    Dim sLine As ClsLine
    Dim cIndex As Long

    Set DestFrame = New ClsFrame
    DestFrame.Index = SourceFrame.Index + 1

    For Each sObj In SourceFrame.Objects
        Set cObj = New ClsObject
        cObj.ObjectName = sObj.ObjectName

        For Each sLine In sObj.Lines
            Set cLine = New ClsLine

            cLine.IsCircle = sLine.IsCircle
            cLine.LineColour = sLine.LineColour
            cLine.LineWidth = sLine.LineWidth

            cLine.PointEndX = sLine.PointEndX
            cLine.PointEndY = sLine.PointEndY
            cLine.PointStartX = sLine.PointStartX
            cLine.PointStartY = sLine.PointStartY

            For cIndex = 1 To sLine.Connects.Count
                cLine.Connects.Add sLine.Connects.Item(cIndex)
            Next

            cObj.Lines.Add cLine
        Next

        DestFrame.Objects.Add cObj
    Next
End Sub

Public Sub LoadObject(FileName As String)
    Dim FileNum As Integer
    Dim LineData As String
    Dim SptLineData() As String
    Dim SubSplit() As String
    Dim cIndex As Long
    Dim NewObjIndex As Long
    Dim cObject As ClsObject
    Dim cLine As ClsLine

    If Dir(FileName) = "" Then
        MsgBox "The specified object file """ & FileName & """ does not exist.", vbCritical, "Stickman Animation Studio"
    Else
        FileNum = FreeFile
        Open FileName For Input As FileNum

        Line Input #FileNum, LineData

        For Each cObject In ActiveFrame.Objects
            If InStr(1, cObject.ObjectName, "(") Then
                NewObjIndex = NewObjIndex + 1
            End If
        Next

        Set cObject = New ClsObject

        cObject.ObjectName = LineData & " (" & NewObjIndex + 1 & ")"

        Do While Not EOF(FileNum)
            Set cLine = New ClsLine

            Line Input #FileNum, LineData
            SptLineData = Split(LineData, ",")

            SubSplit = Split(SptLineData(0), ":")
            cLine.PointStartX = Int(SubSplit(0))
            cLine.PointStartY = Int(SubSplit(1))

            SubSplit = Split(SptLineData(1), ":")
            cLine.PointEndX = Int(SubSplit(0))
            cLine.PointEndY = Int(SubSplit(1))

            cLine.LineWidth = Int(SptLineData(2))
            cLine.LineColour = Int(SptLineData(4))
            cLine.IsCircle = Int(SptLineData(5))

            SptLineData = Split(SptLineData(3), "|")
            For cIndex = 0 To UBound(SptLineData) - 1
                SubSplit = Split(SptLineData(cIndex), ":")

                cLine.Connect Int(SubSplit(0)), Int(SubSplit(1))
            Next

            cObject.Lines.Add cLine
            Set cLine = Nothing
        Loop

        Close #FileNum

        ActiveFrame.Objects.Add cObject
        Set ActiveObject = cObject
        Set cObject = Nothing

        With FrmMain
            .RefreshFrame
            .RefreshObjectsList
            .UpdateSelectedObjectStats
            .ctlThumbs.RefreshThumbs
        End With
    End If
End Sub
