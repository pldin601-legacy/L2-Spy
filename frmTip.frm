VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Timer tmt 
         Interval        =   10
         Left            =   60
         Top             =   960
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00EEF2F2&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   10
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00EEF2F2&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   9
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00EEF2F2&
         Caption         =   "0с. назад"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   8
         Top             =   660
         Width           =   1740
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00EEF2F2&
         Caption         =   "0с."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   7
         Top             =   420
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   0
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Убийства:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Дуэли:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Посл. раз заходил:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Возраст:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   420
         Width           =   660
      End
      Begin VB.Label lNick 
         Alignment       =   2  'Center
         BackColor       =   &H00EEF2F2&
         Caption         =   "Marusha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   0
         Width           =   2835
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lX As POINTAPI
Dim y As POINTAPI

Private Sub tmt_Timer()
    Call GetCursorPos(y)
    If y.x <> lX.x And y.y <> lX.y Then Me.Hide: tmt.Enabled = False
End Sub

Sub RaiseMe()
    If IsMouseInWindow(frmSpy.lv1.hWnd) Then
        Call GetCursorPos(lX)
        tmt.Enabled = True
        Me.Move tppX * lX.x, tppY * lX.y
        SetFormTColorXP Me, vbMagenta, 200
        Me.Visible = True
    End If
End Sub
