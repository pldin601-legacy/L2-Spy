VERSION 5.00
Begin VB.Form frmMask 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E2E2&
   BorderStyle     =   0  'None
   Caption         =   "Lineage II Wnet Spy by Roman Gemini"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMask.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   480
      Picture         =   "frmMask.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   840
      Picture         =   "frmMask.frx":09C6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   1200
      Picture         =   "frmMask.frx":0F4A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   1200
      Picture         =   "frmMask.frx":14CE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   3060
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   840
      Picture         =   "frmMask.frx":1A52
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   3060
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   480
      Picture         =   "frmMask.frx":1FD6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   3060
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   1200
      Picture         =   "frmMask.frx":255A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   840
      Picture         =   "frmMask.frx":2ADE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   480
      Picture         =   "frmMask.frx":3062
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   2700
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1200
      Picture         =   "frmMask.frx":35E6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   2340
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   840
      Picture         =   "frmMask.frx":3B6A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cbi 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   480
      Picture         =   "frmMask.frx":40EE
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   2340
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox cb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2190
      Picture         =   "frmMask.frx":4672
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      ToolTipText     =   "Свернуть"
      Top             =   90
      Width           =   315
   End
   Begin VB.PictureBox cb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2540
      Picture         =   "frmMask.frx":4BF6
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      ToolTipText     =   "Развернуть"
      Top             =   90
      Width           =   315
   End
   Begin VB.PictureBox cb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2900
      Picture         =   "frmMask.frx":517A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      ToolTipText     =   "Закрыть"
      Top             =   90
      Width           =   315
   End
   Begin VB.PictureBox p8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   1800
      Picture         =   "frmMask.frx":56FE
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox p7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   1440
      Picture         =   "frmMask.frx":5772
      ScaleHeight     =   60
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox p6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   1380
      Picture         =   "frmMask.frx":58B6
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox p5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      Picture         =   "frmMask.frx":593A
      ScaleHeight     =   300
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox p4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      Picture         =   "frmMask.frx":5A6E
      ScaleHeight     =   285
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   1860
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox p3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1740
      Picture         =   "frmMask.frx":5B96
      ScaleHeight     =   450
      ScaleWidth      =   90
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox p2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1560
      Picture         =   "frmMask.frx":5E32
      ScaleHeight     =   450
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1380
      Picture         =   "frmMask.frx":6146
      ScaleHeight     =   450
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lineage II Wnet Spy by Roman Gemini"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   420
      TabIndex        =   20
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   100
      Picture         =   "frmMask.frx":64D2
      Stretch         =   -1  'True
      Top             =   100
      Width           =   240
   End
End
Attribute VB_Name = "frmMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Capturing As Boolean
Dim OldX As Single, OldY As Single
Dim ChildHwnd As Long


Sub DrawSkin()

Me.Cls

Call StretchBlt(Me.hdc, 0, 0, 9, 30, p1.hdc, 0, 0, 9, 30, vbSrcCopy)
Call StretchBlt(Me.hdc, 9, 0, (Me.Width / tppX) - 15, 30, p2.hdc, 0, 0, 8, 30, vbSrcCopy)
Call StretchBlt(Me.hdc, (Me.Width / tppX) - 6, 0, 7, 30, p3.hdc, 0, 0, 7, 30, vbSrcCopy)

Call StretchBlt(Me.hdc, 0, 30, 4, (Me.Height / tppY) - 34, p4.hdc, 0, 0, 4, 19, vbSrcCopy)
Call StretchBlt(Me.hdc, (Me.Width / tppX) - 4, 30, 4, (Me.Height / tppY) - 34, p5.hdc, 0, 0, 4, 20, vbSrcCopy)

Call StretchBlt(Me.hdc, 0, (Me.Height / tppY) - 4, 5, 4, p6.hdc, 0, 0, 5, 4, vbSrcCopy)
Call StretchBlt(Me.hdc, (Me.Width / tppX) - 4, (Me.Height / tppY) - 4, 5, 4, p8.hdc, 0, 0, 4, 4, vbSrcCopy)
Call StretchBlt(Me.hdc, 4, (Me.Height / tppY) - 4, (Me.Width / tppX) - 8, 4, p7.hdc, 0, 0, 21, 4, vbSrcCopy)

cb(0).Top = 5 * tppX
cb(1).Top = 5 * tppX
cb(2).Top = 5 * tppX

cb(0).Left = Me.Width - (26 * tppX)
cb(1).Left = cb(0).Left - (24 * tppX)
cb(2).Left = cb(1).Left - (24 * tppX)


End Sub

Private Sub cb_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button = 1 Then cb(index).Picture = cbi((index * 3) + 2).Picture
End Sub

Private Sub cb_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If IsMouseInWindow(cb(index).hWnd) Then
        If Not Capturing Then
            SetCapture cb(index).hWnd
            Capturing = True
            cb(index).Picture = cbi((index * 3) + 1).Picture
        End If
    Else
        If Capturing Then
            ReleaseCapture
            Capturing = False
            cb(index).Picture = cbi((index * 3)).Picture
        End If
    End If

End Sub

Private Sub cb_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error Resume Next
    
    If Button = 1 Then
        cb(index).Picture = cbi((index * 3) + 1).Picture
        Select Case index
        Case 0
            Unload frmSpy
            Unload Me
        Case 1
            If Me.WindowState = vbMaximized Then Me.WindowState = vbNormal Else Me.WindowState = vbMaximized
        Case 2
            Me.WindowState = vbMinimized
        End Select
    End If
    SetCapture cb(index).hWnd
End Sub

Private Sub Form_Load()
    SetFormTColorXP Me, vbMagenta, 240

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    OldX = x
    OldY = y
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
 Me.Left = Me.Left + (x - OldX)
 Me.Top = Me.Top + (y - OldY)
End If

End Sub


Private Sub Form_Resize()
    DrawSkin
End Sub



Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseMove Button, Shift, x, y
End Sub





