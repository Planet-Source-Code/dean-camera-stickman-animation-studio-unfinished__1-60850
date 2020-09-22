VERSION 5.00
Begin VB.Form FrmSetWorkspaceSize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Workspace Size"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmSetWorkspaceSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "100"
      Top             =   1280
      Width           =   855
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "100"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblWidthHeight 
      Caption         =   "Width (pixels):                   Height (pixels):"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   900
      Width           =   1095
   End
   Begin VB.Image imgWorkspace 
      Height          =   480
      Left            =   600
      Picture         =   "FrmSetWorkspaceSize.frx":5C12
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lblAffectNote 
      Caption         =   "Note: These dimentions will affect your entire project."
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblEnterDimentions 
      Caption         =   "Please enter the new dimentions of the workspace:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmSetWorkspaceSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ================================================================================

Private Sub Form_Load()
    txtWidth = WorkArea(0)
    txtHeight = WorkArea(1)
End Sub

' ================================================================================

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If Val(txtHeight.Text) < 200 Then txtHeight.Text = WorkArea(1)
    If Val(txtWidth.Text) < 200 Then txtWidth.Text = WorkArea(0)

    If Val(txtHeight.Text) > 800 Then txtHeight.Text = "800"
    If Val(txtWidth.Text) > 800 Then txtWidth.Text = "800"

    WorkArea(0) = Val(txtWidth.Text)
    WorkArea(1) = Val(txtHeight.Text)
        
    With FrmMain
        .picCurrFrame.Cls
        .DeleteWorkAreaDC
        .CreateWorkAreaDC
    
        .ctlThumbs.DeleteThumbDC
        .ctlThumbs.CreateThumbDC
        .ctlThumbs.RefreshThumbs
    End With
    
    Unload Me
End Sub
