VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "О программе"
   ClientHeight    =   4800
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7335
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0442
   ScaleHeight     =   3313.045
   ScaleMode       =   0  'User
   ScaleWidth      =   6887.944
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   4260
      Width           =   1395
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Версия"
      Height          =   195
      Left            =   3720
      TabIndex        =   10
      Top             =   900
      Width           =   3225
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Программа посвящается клану Phoenix."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   4560
      Width           =   3045
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Написать письмо автору"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3720
      MouseIcon       =   "frmAbout.frx":34C86
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "roman.gemini@gmail.com"
      Top             =   3300
      Width           =   1875
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":: Обратная связь ::"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   6
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":: Автор ::"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   5
      Top             =   2340
      Width           =   1035
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Задача программы - отображать онлайн-статус избранных персонажей и кланов на сервере Wnet Lineage II."
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   3315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":: О программе ::"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   3
      Top             =   1260
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Роман Лахтадыр (aka Roman Gemini)"
      Height          =   195
      Left            =   3720
      TabIndex        =   2
      Top             =   2640
      Width           =   2745
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lineage II Wnet Spy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3660
      TabIndex        =   1
      Top             =   480
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Woobind Software © 2007"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   2205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   3155.214
      X2              =   3155.214
      Y1              =   0
      Y2              =   3313.045
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label10.Caption = "Версия " & GetVersion
        
        
        ' On Error Resume Next
    Dim y As Single, x As Long
    Me.ScaleMode = vbPixels
    x = GetSysColor(&HF)
    
    For y = 0 To Me.ScaleHeight
        Me.Line (225, y)-(Me.ScaleWidth, y), RGBBright(x, 255 - (100 / Me.ScaleHeight * y))
    Next y

End Sub

Private Sub Label8_Click()
    Call RunWEB("mailto:roman.gemini@gmail.com")
End Sub
