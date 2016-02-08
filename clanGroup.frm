VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form clanGroup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Онлайн-статус группы кланов"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "clanGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pScale 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   4290
      Width           =   2235
      Begin VB.Label lRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   60
         Width           =   2115
      End
      Begin VB.Image imgBack 
         Height          =   195
         Left            =   540
         Picture         =   "clanGroup.frx":0442
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1620
      End
      Begin VB.Image imgFore 
         Height          =   190
         Left            =   60
         Picture         =   "clanGroup.frx":0514
         Stretch         =   -1  'True
         Top             =   60
         Width           =   525
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Обновить"
      Height          =   375
      Left            =   4500
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Сохранить..."
      Height          =   375
      Left            =   4500
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Загрузить..."
      Height          =   375
      Left            =   4500
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Удалить..."
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   900
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Добавить..."
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   420
      Width           =   1695
   End
   Begin ComctlLib.ListView lv1 
      Height          =   3795
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   6694
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Клан"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Онлайн"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Всего"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Суммарный онлайн:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Группа кланов:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
   Begin VB.Menu mnuCLN 
      Caption         =   "show"
      Visible         =   0   'False
      Begin VB.Menu mnuObz 
         Caption         =   "Обзор клана """""
      End
   End
End
Attribute VB_Name = "clanGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClanList() As String
Dim kach As clsKachalka
Dim Working As Boolean
Dim xoO As Integer, xoT As Integer

Private Sub Command1_Click()
    
    If Working Then Exit Sub

    On Error Resume Next
    
    Dim N As ListItem
    Dim u As String, V As String
    
    If lv1.SelectedItem = "" Then Exit Sub
    
    Set N = lv1.SelectedItem
    If MsgBox("Удалить клан ''" & N.Text & "'' из списка?", vbQuestion + vbYesNo) = vbYes Then
        V = Join(ClanList, vbCrLf)
        V = Replace(V, vbCrLf & N, "", , , vbBinaryCompare)
        ClanList = Split(V, vbCrLf)
        SaveClanList
        xoO = xoO - N.SubItems(1)
        xoT = xoT - N.SubItems(2)
        lv1.ListItems.Remove N.index
        DrawOnline xoO, xoT
    End If
    
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


Private Sub Command2_Click()

    If Working Then Exit Sub
    
    Dim cdlg As New CommonDlg
    
    cdlg.DefaultExt = "gf"
    cdlg.hWndOwner = Me.hWnd
    cdlg.DialogTitle = "Открытие списка кланов"
    cdlg.Filter = "Файлы списков кланов (*.gf)|*.fg"
    cdlg.FileName = ""
    cdlg.ShowOpen
    
    If cdlg.FileName > "" Then
        Call LoadClanList(cdlg.FileName)
        Call RefreshClans
        Call SaveClanList
    End If

End Sub

Sub RefreshClans()
    
    If Working Then Exit Sub
    
    On Error Resume Next
    
    Dim lstitm As ListItem, sClan As Variant, oO As Integer, oT As Integer
    xoO = 0
    xoT = 0
    
    lv1.ListItems.Clear
    Label2.Caption = "Обновление..."
    
    pScale.Visible = False
    Working = True
    Me.MousePointer = 13
    
    For Each sClan In ClanList
        If sClan > "" Then
            Set lstitm = lv1.ListItems.Add()
            lstitm.Text = sClan
            lstitm.SubItems(1) = "..."
            lstitm.SubItems(2) = "..."
        End If
    Next sClan

    For Each lstitm In lv1.ListItems
        DoEvents
        If GetCInfo(lstitm.Text, oO, oT) Then
            lstitm.SubItems(1) = oO
            lstitm.SubItems(2) = oT
            xoO = xoO + oO
            xoT = xoT + oT
        Else
            If MsgBox("При получении информации о клане """ & lstitm.Text & """ возникла ошибка! Продолжить обновление?", vbQuestion + vbYesNo, "L2 Wnet Spy") = vbNo Then Exit For
        End If
    Next lstitm

    StatusX
        
    Working = False
    Me.MousePointer = 0
    
End Sub


Sub StatusX()
    
    
    If xoT > 0 Then
        pScale.Visible = True
        Label2.Caption = "Суммарный онлайн:"
        DrawOnline xoO, xoT
    Else
        Label2.Caption = "Пусто."
    End If

End Sub

Sub AddClan()

    Dim u As String, V As String
    Dim lstitm As ListItem, oO As Integer, oT As Integer
    
    u = InputBox("Введите название клана:", "L2 Wnet Spy")
    If Len(u) = 0 Then Exit Sub
    
    V = Join(ClanList, vbCrLf)
    If InStr(V, u & vbCrLf) Then _
            MsgBox "Этот клан уже есть в списке!", vbExclamation, "Ахтунг! 8)": _
            Exit Sub
            
    V = V & vbCrLf & u
    
    ClanList = Split(V, vbCrLf)
    
    Call SaveClanList
    
    Set lstitm = lv1.ListItems.Add()
    lstitm.Text = u
    lstitm.SubItems(1) = "..."
    lstitm.SubItems(2) = "..."
    Label2.Caption = "Обновление..."
    pScale.Visible = False
    DoEvents
    
    If GetCInfo(u, oO, oT) Then
            lstitm.SubItems(1) = oO
            lstitm.SubItems(2) = oT
            xoO = xoO + oO
            xoT = xoT + oT
    Else
            Call MsgBox("При получении информации о клане """ & u & """ возникла ошибка!", vbExclamation, "L2 Wnet Spy")
    End If
    
    StatusX
    
End Sub

Sub LoadClanList(Optional inFilename As String = "")
    
    On Error Resume Next
    If inFilename = "" Then inFilename = LowPath(App.Path) & "default.gf"
    
    Dim tmpstr As String
    
    ClanList = Split(LoadFile(inFilename), vbCrLf)
    
End Sub

Sub SaveClanList(Optional inFilename As String = "")
    
    On Error Resume Next
    If inFilename = "" Then inFilename = LowPath(App.Path) & "default.gf"
    
    Dim tmpstr As String
    
    Call SaveFile(inFilename, Join(ClanList, vbCrLf))
    
End Sub

Private Sub Command3_Click()

    Dim cdlg As New CommonDlg
    
    cdlg.DefaultExt = "gf"
    cdlg.hWndOwner = Me.hWnd
    cdlg.DialogTitle = "Сохранение списка кланов"
    cdlg.Filter = "Файлы списков кланов (*.gf)|*.fg"
    cdlg.FileName = ""
    cdlg.ShowSave
    
    If cdlg.FileName > "" Then
        If Not FileExists(cdlg.FileName) Then
            Call SaveClanList(cdlg.FileName)
        Else
            If MsgBox(cdlg.FileName & " уже существует" & vbCrLf & "Заменить?", vbExclamation + vbYesNo) = vbYes Then
                Call SaveClanList(cdlg.FileName)
            End If
        End If
    End If
    
End Sub

Private Sub Command4_Click()

    RefreshClans
    
End Sub

Private Sub Command5_Click()
    
    If Working Then Exit Sub
    AddClan
    
End Sub

Private Sub Form_Load()
    
    Me.Show
    LoadClanList
    RefreshClans
    
End Sub

Function GetCInfo(inClan As String, outOnline As Integer, outTotal As Integer) As Boolean
    
    On Error Resume Next
    Dim tmpstr As String
    
    Set kach = New clsKachalka
    tmpstr = kach.DownloadToString("http://l2w2new.wnet.ua/UserPanel/clans.asp?action=show&name=" & inClan)
    Set kach = Nothing
    
    GetCInfo = Parse(tmpstr, outOnline, outTotal)

End Function


Function Parse(inText As String, outOnline As Integer, outTotal As Integer) As Boolean

    
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
        
      
          
        If prsOnline = "on" Then L = L + 1
        
        k = k + 1
        u = i
        
    Loop Until Not True
    
    If k = 0 Then
        Parse = False
        outOnline = 0
        outTotal = 0
    Else
        Parse = True
        outOnline = L
        outTotal = k
    End If
    
    
End Function

Private Sub Form_Unload(Cancel As Integer)
If Working Then Cancel = 1
End Sub

Private Sub lv1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cln As ListItem
    Set cln = lv1.SelectedItem
    
    If Button = 2 And cln.Text > "" Then
        mnuObz.Caption = "Обзор клана """ & cln & """"
        mnuObz.Tag = cln
        PopupMenu mnuCLN
    End If
End Sub

Private Sub mnuObz_Click()
    CurrClan = lv1.SelectedItem.Text
    dlgClan.sRefresh
End Sub
