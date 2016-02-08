VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form dlgClan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Обзор клана"
   ClientHeight    =   6780
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgClan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   ScaleHeight     =   6780
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   3420
      Top             =   6180
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Обновить"
      Height          =   375
      Left            =   4860
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.PictureBox pScale 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4500
      ScaleHeight     =   315
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   120
      Width           =   2715
      Begin VB.Label lRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   2535
      End
      Begin VB.Image imgFore 
         Height          =   190
         Left            =   60
         Picture         =   "dlgClan.frx":000C
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
      Begin VB.Image imgBack 
         Height          =   190
         Left            =   540
         Picture         =   "dlgClan.frx":00DE
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2085
      End
   End
   Begin ComctlLib.ListView lst 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9763
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Профессия"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Звание"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Статус"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Закрыть"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lUpd 
      BackStyle       =   0  'Transparent
      Caption         =   "Обновление..."
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   6300
      Width           =   1815
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "dlgClan.frx":01B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "dlgClan.frx":0502
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Онлайн:"
      Height          =   195
      Left            =   3780
      TabIndex        =   4
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lClan 
      BackStyle       =   0  'Transparent
      Caption         =   "Клан"
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
      Left            =   960
      TabIndex        =   3
      Top             =   180
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Клан:"
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   180
      Width           =   435
   End
   Begin VB.Menu gmnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Добавить """" в общий список"
      End
   End
End
Attribute VB_Name = "dlgClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kach As clsKachalka
Option Explicit




Private Sub Command1_Click()

    sRefresh
    
End Sub

Private Sub Form_Load()
    
    DrawOnline 0, 0
    Me.Visible = True
    Me.BackColor = GetSysColor(&HF)
    '    Parse LoadFile("d:\block.htm"), CurrClan, CurrOnline
    sRefresh
    
End Sub


Sub sRefresh()
On Error Resume Next
Me.MousePointer = 13
UpdateOn
Set kach = New clsKachalka
Dim tmpstr As String
tmpstr = kach.DownloadToString("http://l2w2new.wnet.ua/UserPanel/clans.asp?action=show&name=" & CurrClan)
Set kach = Nothing
Parse tmpstr, CurrClan, CurrOnline
Me.MousePointer = 0
UpdateOff
End Sub


Sub Parse(inText As String, zClan As String, Optional onlyOnline As Boolean = False)

    lst.ListItems.Clear
    
    Dim tmpstr As String, u, i, j, k As Integer, L As Integer
    Dim fin As Boolean
    Dim prsNick As String
    Dim prsProf As String
    Dim prsCaption As String
    Dim prsOnline As String
    Dim itm As ListItem
    On Error Resume Next
    u = 1
    
    Do
        
        
        i = InStr(u, inText, "<tr class=""CharRow")
        If i = 0 Then Exit Do
        i = InStr(i, inText, "<td><span class=")
        i = InStr(i, inText, """") + 1
        j = InStr(i + 1, inText, """")
        prsOnline = Mid(inText, i, j - i)
        
        
        i = InStr(i, inText, "show&name")
        i = InStr(i, inText, "=") + 1
        j = InStr(i, inText, """")
        prsNick = Mid(inText, i, j - i)
        
        i = InStr(i, inText, "<br>") + 4
        j = InStr(i, inText, "</td>")
        prsCaption = Mid(inText, i, j - i)
        
        i = InStr(i, inText, "<br>") + 4
        j = InStr(i, inText, "</td>")
        prsProf = Mid(inText, i, j - i)
        
        If prsOnline = "on" Then L = L + 1
        k = k + 1
        u = i
        
        If onlyOnline Then
        
          If prsOnline = "on" Then
            Set itm = lst.ListItems.Add(, prsNick, prsNick, 0, 0)
            itm.SubItems(1) = prsProf
            itm.SubItems(2) = prsCaption
            itm.SubItems(3) = prsOnline
            If prsOnline = "on" Then
                itm.SmallIcon = 2
            Else
                itm.SmallIcon = 1
            End If
          End If
          
        Else
    
            Set itm = lst.ListItems.Add(, "itm" & prsNick, prsNick, 0, 0)
            itm.SubItems(1) = prsProf
            itm.SubItems(2) = prsCaption
            itm.SubItems(3) = prsOnline
            If prsOnline = "on" Then
                itm.SmallIcon = 2
            Else
                itm.SmallIcon = 1
            End If
        
        End If
        
    Loop Until Not True
    
    If k = 0 Then
        MsgBox ("Либо Wnet висит, либо нет такого клана!")
        DrawOnline 0, 0
    Else
        DrawOnline L, k
    End If
    
    lClan.Caption = zClan
    
End Sub

Sub DrawOnline(inOnline As Integer, inTotal As Integer)

    Dim i As Integer, j  As Integer, k As Integer
    
    If inTotal > 0 Then
        i = pScale.Width - 120
        j = Fix(i / inTotal * inOnline)
        k = i - j
        
        imgFore.Width = j
        imgFore.Left = 60
    
        imgBack.Left = j + 60
        imgBack.Width = k
        lRate.Caption = Format(inOnline, "0") & "/" & Format(inTotal, "0") & " (" & Format(Fix(100 / inTotal * inOnline), "0") & "%)"
    Else
        i = pScale.Width - 120
        j = Fix(0)
        k = i - j
        
        imgFore.Width = j
        imgFore.Left = 60
    
        imgBack.Left = j + 60
        imgBack.Width = k
        lRate.Caption = Format(inOnline, "0") & "/" & Format(inTotal, "0") & " (0%)"
    End If
    
    
End Sub

Private Sub lst_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    lst.SortKey = ColumnHeader.index - 1
    
End Sub

Private Sub lst_DblClick()
  
  If Not IsNothing(lst.SelectedItem) Then
    RunWEB "http://l2w2new.wnet.ua/UserPanel/ShowChar.asp?action=show&name=" & lst.SelectedItem.Text
  End If

End Sub

Private Sub lst_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim u As ListItem
    
    On Error Resume Next
    
    If Button = 2 Then
        Set u = lst.SelectedItem
        If Not IsNothing(u) Then
            mnuAdd.Caption = "Добавить """ & u.Text & """ в общий список..."
            mnuAdd.Tag = u.Text
            PopupMenu Me.gmnu
        End If
    End If
    
End Sub

Private Sub mnuAdd_Click()
    frmSpy.AddNick mnuAdd.Tag
End Sub

Private Sub OKButton_Click()
    
    Unload Me
    
End Sub

Private Sub Timer1_Timer()
    sRefresh
End Sub


Sub UpdateOn()

    Dim o As Integer
    lUpd.Visible = True
    
    For o = 255 To 0 Step -16
        lUpd.ForeColor = RGBBright(Me.BackColor, o)
        Dream 25
    Next o
    
End Sub

Sub UpdateOff()

    Dim o As Integer
    
    For o = 0 To 255 Step 16
        lUpd.ForeColor = RGBBright(Me.BackColor, o)
        Dream 25
    Next o
    
    lUpd.Visible = False
    
End Sub

