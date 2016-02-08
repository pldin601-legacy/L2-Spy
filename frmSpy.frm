VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSpy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lineage II Wnet Spy by Roman Gemini"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7440
      Top             =   5520
   End
   Begin VB.PictureBox isUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E2E2&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6900
      Picture         =   "frmSpy.frx":0442
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox isDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E2E2&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Picture         =   "frmSpy.frx":0875
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox tray 
      Height          =   255
      Left            =   5880
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   17
      Top             =   5100
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Сервер"
      Height          =   1395
      Left            =   5940
      TabIndex        =   12
      Top             =   2580
      Width           =   1575
      Begin VB.CommandButton Command7 
         Caption         =   "Обновить"
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "В мире:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lIngame 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Image imSS 
         Height          =   285
         Left            =   1080
         Top             =   240
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Статус:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fast!"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   9
      ToolTipText     =   "Быстрый просмотр онлайн-статуса персонажа"
      Top             =   5340
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Группа кланов..."
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      ToolTipText     =   "Просмотреть онлайн-статус группы кланов"
      Top             =   1980
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Обзор клана..."
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Просмотреть состав клана..."
      Top             =   1500
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   5880
      ScaleHeight     =   675
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   4320
      Width           =   1755
      Begin VB.CheckBox Check2 
         Caption         =   "Уведомления"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         ToolTipText     =   "Посылать уведомления в область уведомлений"
         Top             =   360
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Автообновление"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         ToolTipText     =   "Обновлять онлайн персонажей периодически"
         Top             =   60
         Width           =   1635
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7320
      Top             =   5040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Обновить"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      ToolTipText     =   "Обновить статус персонажей в списке"
      Top             =   5340
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Удалить..."
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      ToolTipText     =   "Удалить персонажа из списка"
      Top             =   5340
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Добавить..."
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Добавить нового персонажа в список"
      Top             =   5340
      Width           =   1395
   End
   Begin ComctlLib.ListView lv1 
      Height          =   3735
      Left            =   180
      TabIndex        =   1
      Top             =   1500
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "IL"
      SmallIcons      =   "IL"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":: О программе ::"
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
      Left            =   6120
      MouseIcon       =   "frmSpy.frx":1098
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "О программе"
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   7800
      Y1              =   1060
      Y2              =   1060
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   0
      MouseIcon       =   "frmSpy.frx":13A2
      MousePointer    =   99  'Custom
      Picture         =   "frmSpy.frx":16AC
      ToolTipText     =   "Кликните чтобы посетить сайт клана"
      Top             =   0
      Width           =   7860
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":: Другие программы by Roman Gemini ::"
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
      Left            =   4500
      MouseIcon       =   "frmSpy.frx":1CAEC
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Если есть желание ознакомиться с другими моими творениями, то вам сюда..."
      Top             =   1170
      Width           =   3030
   End
   Begin ComctlLib.ImageList IL 
      Left            =   1980
      Top             =   1080
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
            Picture         =   "frmSpy.frx":1CDF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSpy.frx":1D148
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Статус персонажей:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1170
      Width           =   1560
   End
   Begin VB.Menu prMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu clp1 
         Caption         =   "Lineage II Wnet Spy"
         Enabled         =   0   'False
      End
      Begin VB.Menu clp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowHide 
         Caption         =   "Показать/скрыть основное окно"
      End
      Begin VB.Menu mnuClnShow 
         Caption         =   "Обзор клана..."
      End
      Begin VB.Menu mnuGRP 
         Caption         =   "Группа кланов..."
      End
      Begin VB.Menu clp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefrsh 
         Caption         =   "Освежить онлайн персонажей"
      End
      Begin VB.Menu clp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе..."
      End
      Begin VB.Menu mnuAutorun 
         Caption         =   "Запускать вместе с Windows"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Закрыть"
      End
   End
   Begin VB.Menu mnuCols 
      Caption         =   "Columns"
      Visible         =   0   'False
      Begin VB.Menu mnuCSel 
         Caption         =   "Раса"
         Index           =   0
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Пол"
         Index           =   1
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Дуэлей"
         Index           =   2
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Убийств"
         Index           =   3
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Карма"
         Index           =   4
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Возраст"
         Index           =   5
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Посл. заход"
         Index           =   6
      End
      Begin VB.Menu mnuCSel 
         Caption         =   "Наиграно"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kach As clsKachalka
Dim Working As Boolean
Dim CancelBoul As Boolean
Dim Capturing As Boolean
Dim CharBuffer() As CharConstructive
Dim WasOnline As Boolean

Sub SaveColumnSizes()
    Dim b As Integer
    For b = 1 To lv1.ColumnHeaders.Count
        SaveSetting "l2wnet spy", "ColumnSizes", lv1.ColumnHeaders(b).Text, CStr(lv1.ColumnHeaders(b).Width)
    Next b
End Sub

Function IsServerOnline(Optional pCount As Integer = 0) As Boolean

On Error Resume Next

Set kach = New clsKachalka
Dim tmpstr As String
Dim iProof As String
Dim L As Integer
Dim M As Integer
Dim N As Integer

tmpstr = kach.DownloadToString("http://www.l2.wnet.ua/")

If tmpstr = "" Then IsServerOnline = False: Exit Function

L = InStr(tmpstr, "Игровой сервер:")

If L > 0 Then
    M = InStr(L, tmpstr, "<img src=""")
Else
    Exit Function
End If

N = InStr(M + 11, tmpstr, """")
iProof = Mid(tmpstr, M + 10, N - M - 10)

If iProof = "/images/up.gif" Then
    IsServerOnline = True
Else
    IsServerOnline = False
End If

L = InStr(tmpstr, "Сейчас в мире")

If L > 0 Then
    M = InStr(L, tmpstr, "<td align=""right"">")
Else
    Exit Function
End If

N = InStr(M + 18, tmpstr, "</td>")
iProof = Mid(tmpstr, M + 18, N - M - 18)

pCount = Val(iProof)

Set kach = Nothing

End Function

Sub AddNick(Optional uNick As String = "")

    Dim u As String, V As String, z As ListItem
    
    If uNick = "" Then
        u = InputBox("Введите ник персонажа:", "L2 Wnet Spy")
    Else
        u = uNick
    End If
    
    If Len(u) = 0 Then Exit Sub
    
    V = Join(Nicks, vbCrLf)
    If InStr(V, u & vbCrLf) Then MsgBox "Этот персонаж уже есть в списке!", vbExclamation, "Ахтунг! 8)": Exit Sub
    V = V & u & vbCrLf
    Nicks = Split(V, vbCrLf)
    SaveNicks
    Set z = lv1.ListItems.Add()
    z.Text = u
    RefreshNicks
    
End Sub

Sub PingServer()
    Dim iG As Integer
    Dim iso As Boolean
    Me.MousePointer = 13
    iso = IsServerOnline(iG)
    imSS.Picture = IIf(iso, isUp.Picture, isDown.Picture)
    lIngame.Caption = Format(iG, "### ##0")
    Me.MousePointer = 0
    If iso Then
        TrayModify tray, "L2 Wnet Spy " + GetVersion & vbCrLf & "Server: online", PicToIco(isUp.hdc, 18, 19)
    Else
        TrayModify tray, "L2 Wnet Spy " + GetVersion & vbCrLf & "Server: offline", PicToIco(isDown.hdc, 18, 19)
    End If
    
    If WasOnline <> iso And Check2.Value > 0 Then
      If iso Then
        TrayBalloon tray, "L2 Wnet Spy " + GetVersion, 1, "Сервер запущен!"
      Else
        TrayBalloon tray, "L2 Wnet Spy " + GetVersion, 1, "Сервер отключен!"
      End If
    End If

    WasOnline = iso

End Sub

Sub ReloadNicks()

    On Error Resume Next
    Dim Nick As Variant
    Dim lAdd As ListItem
    
    If Working Then Exit Sub
    lv1.ListItems.Clear
    
    For Each Nick In Nicks
        If Nick > "" Then
            Set lAdd = lv1.ListItems.Add()
            lAdd.Text = Nick
        End If
    Next Nick
    
End Sub

Sub RefreshNicks()
    If Working Then Exit Sub
    Working = True
    On Error Resume Next
    Dim Nick As Variant
    Dim i As Integer
    Dim itm As ListItem
    Dim T As Boolean
    Dim u As String
    Dim V As Boolean
    Dim e As String
    Dim offline As Boolean
    Dim splt As String
    Dim o As Integer
    Dim csChar As CharConstructive
    Me.MousePointer = 13
    
    ' Проверка смены статуса
    If lv1.ListItems.Count > 0 Then
        Dim xCheck() As String
        Dim xMessage As String
        ReDim xCheck(1 To lv1.ListItems.Count) As String
        ReDim CharBuffer(1 To lv1.ListItems.Count) As CharConstructive
    End If
    
    
    For o = 1 To lv1.ListItems.Count
        Set itm = lv1.ListItems(o)
            
            i = i + 1
            xCheck(o) = itm.SubItems(2)
            itm.SubItems(2) = "Смотрим..."
            DoEvents
            Call IsCharOnline(CStr(itm.Text), csChar)
            CharBuffer(o) = csChar
            
            Select Case csChar.ccOnline
            Case 0
                itm.SubItems(2) = "Не шпилит"
                itm.SmallIcon = 1
                If xCheck(o) > "" And xCheck(o) <> itm.SubItems(2) Then xMessage = xMessage & vbCrLf & itm.Text & ": " & itm.SubItems(2)
            Case 1
                itm.SubItems(2) = "Шпилит!"
                itm.SmallIcon = 2
                If xCheck(o) > "" And xCheck(o) <> itm.SubItems(2) Then xMessage = xMessage & vbCrLf & itm.Text & ": " & itm.SubItems(2)
            Case 2
                itm.SubItems(2) = "Нет такого"
                itm.SmallIcon = 1
            Case 3
                itm.SubItems(2) = "Ошибко!"
                Select Case MsgBox("Панель игрока не отвечает на запросы." + vbCrLf + "Продолжать долбить панель игрока (Да) или отменить операцию (Нет)?", vbQuestion + vbYesNo)
                Case vbNo
                    Exit For
                End Select
            End Select
            
            Select Case csChar.ccOnline
            Case 0, 1
                Call UpdateItemD(itm, csChar, lv1)
            End Select
            
        DoEvents
    
    Next o
    
    Me.MousePointer = 0
    Working = False
    If xMessage > "" And Check2.Value Then TrayBalloon tray, "L2 Wnet Spy " + GetVersion, 1, xMessage & vbCrLf
    
    
End Sub


Private Sub Check1_Click()
If Check1.Value = 1 Then Timer1.Enabled = True Else Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
    If Working Then MsgBox "BUSY!": Exit Sub
    AddNick
End Sub

Private Sub Command2_Click()
            RefreshNicks
End Sub

Sub SaveSettings()
    On Error Resume Next
    
    Dim ColSelects As Integer
    Dim i As Integer
    
    For i = mnuCSel.LBound To mnuCSel.UBound
        If mnuCSel(i).Checked Then
            ColSelects = ColSelects Or UnBit(i + 1)
        End If
    Next i
    
    If Check1.Value Then
            ColSelects = ColSelects Or UnBit(9)
    End If
    
    If Check2.Value Then
            ColSelects = ColSelects Or UnBit(10)
    End If
    
    SaveSetting "l2wnet spy", "Options", "Start Visible", xStr(Me.Visible)
    SaveSetting "l2wnet spy", "Options", "Settings", Str(ColSelects)
    
    SaveColumnSizes
    
End Sub

Sub LoadSettings()

    On Error Resume Next
    Dim ColSelects As Integer
    Dim i As Integer
    
    ColSelects = Val(GetSetting("l2wnet spy", "options", "Settings", "0"))
    
    If TrueBit(ColSelects, 9) Then Check1.Value = 1 Else Check1.Value = 0
    If TrueBit(ColSelects, 10) Then Check2.Value = 1 Else Check2.Value = 0

    For i = mnuCSel.LBound To mnuCSel.UBound
        If TrueBit(ColSelects, i + 1) Then mnuCSel(i).Checked = True Else mnuCSel(i).Checked = False
    Next i

    HotClan = GetSetting("l2wnet spy", "options", "Hot Clan", "")
    
End Sub


Function TrueBit(Value As Integer, BitNumber As Integer) As Boolean
    Dim z
    z = UnBit(BitNumber)
    If (Value And z) / z = 1 Then TrueBit = True Else TrueBit = False
End Function

Function UnBit(inBit As Integer)
    Dim k, L
    L = 1
    For k = 1 To inBit
        L = L * 2
    Next k
    UnBit = L
End Function

Private Sub Command3_Click()
    If Working Then MsgBox "BUSY!": Exit Sub

    On Error Resume Next
    Dim N As ListItem
    Dim u As String, V As String
    If lv1.SelectedItem = "" Then Exit Sub
    
    Set N = lv1.SelectedItem
    If MsgBox("Удалить ник ''" & N.Text & "''?", vbQuestion + vbYesNo) = vbYes Then
        V = Join(Nicks, vbCrLf)
        V = Replace(V, N & vbCrLf, "", , , vbBinaryCompare)
        Nicks = Split(V, vbCrLf)
        SaveNicks
        RemoveItemBuffer CharBuffer, N.index
        lv1.ListItems.Remove N.index
    End If
    

End Sub

Sub RemoveFromList(inNick As String)
    lv1.ListItems.Remove "key" & inNick
End Sub


Private Sub Command4_Click()
    
        On Error Resume Next
        Dim Nick As Variant
        Dim i As Integer
        Dim itm As ListItem
        Dim T As Boolean
        Dim u As String
        Dim V As Boolean
        Dim outStatus As String
        Dim outProof As String
        Dim offline As Boolean
        Dim ccS As CharConstructive
    
        Nick = InputBox("Введите ник персонажа:")
        If Len(Trim(Nick)) > 0 Then
                Me.MousePointer = 13
                Command4.Caption = "Смотрим..."
                Command4.Enabled = False
                DoEvents
                V = IsCharOnline(CStr(Nick), ccS)
                Command4.Caption = "Fast!"
                Command4.Enabled = True
                
                If ccS.ccOnline = coIsOnline Then
                    outStatus = "Шпилит!"
                    outProof = ccS.ccProf
                ElseIf ccS.ccOnline = coIsOffline Then
                    outStatus = "Не шпилит"
                    outProof = ccS.ccProf
                ElseIf ccS.ccOnline = coIsNone Then
                    outStatus = "Нет такого"
                    outProof = ""
                End If
                Me.MousePointer = 0
                If ccS.ccOnline = coServerDown Then
                    MsgBox "Панель игрока недоступна!", vbInformation
                Else
                    MsgBox "Никнейм: " & Nick & vbCrLf & "Профессия: " & outProof & vbCrLf & vbCrLf & "Статус: " & outStatus
                End If
        End If

End Sub

Private Sub Command5_Click()

    clanPrompt.Show
    
End Sub

Private Sub Command6_Click()
    clanGroup.Show
End Sub


Private Sub Command7_Click()
    PingServer
End Sub

Private Sub Form_Load()
    
    InitCommonControls
    ' tray.Picture =
    TrayAdd tray, "L2 Wnet Spy " + GetVersion, PicToIco(isDown.hdc, 18, 19)
    Me.Visible = CBol(GetSetting("l2wnet spy", "Options", "Start Visible", "1"))
    Me.MousePointer = 13
    DoEvents
    
    LoadNicks
    LoadSettings
    ReloadColumns
    ReloadNicks
    RefreshNicks
    
    PingServer
    
    Me.MousePointer = 0

    
End Sub



Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Visible = False: Me.WindowState = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
    TrayRemove
    End
End Sub

Private Sub Image1_Click()
    RunWEB "http://www.phoenixclan.at.ua"
End Sub

Private Sub Label2_Click()
    frmAbout.Show
End Sub

Private Sub Label4_Click()
    RunWEB "http://networkmeter.pp.net.ua"
End Sub



Private Sub lv1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    lv1.SortKey = ColumnHeader.index - 1
    
End Sub

Private Sub lv1_DblClick()
  If Not IsNothing(lv1.SelectedItem) Then
    RunWEB "http://l2w2new.wnet.ua/UserPanel/ShowChar.asp?action=show&name=" & lv1.SelectedItem.Text
  End If
End Sub


Private Sub lv1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuCols
    
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuAutorun_Click()
    mnuAutorun.Checked = Not mnuAutorun.Checked
    If mnuAutorun.Checked Then
        SetAutorunUser
    Else
        KillAutorunUser
    End If
End Sub

Private Sub mnuClnShow_Click()
 clanPrompt.Show
End Sub

Private Sub mnuCSel_Click(index As Integer)
    SaveColumnSizes
    mnuCSel(index).Checked = Not mnuCSel(index).Checked
    ReloadColumns
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuGRP_Click()
 clanGroup.Show
End Sub

Private Sub mnuRefrsh_Click()
RefreshNicks
End Sub

Private Sub mnuShowHide_Click()
    tray_MouseMove 1, 0, WM_LBUTTONDOWN, 0
End Sub

Private Sub Timer1_Timer()
 
    RefreshNicks
End Sub


Private Sub Timer2_Timer()
   PingServer
End Sub

Private Sub tray_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim lMsg As Single, hFG As Long
lMsg = x


Select Case lMsg
    Case WM_LBUTTONUP
    Case WM_RBUTTONUP
        mnuAutorun.Checked = IfAutorunUser
        PopupMenu prMenu
    Case WM_BALLOONL
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDOWN
        Me.Visible = Not IsWindowOnScreen(Me.hWnd)
        If Me.Visible Then Me.SetFocus
    Case WM_RBUTTONDOWN
    Case Else
End Select

End Sub

Sub ReloadColumns()
    
    Dim a As ColumnHeader
    Dim b() As Variant, c As Variant
    Dim d As Integer
    Dim e As ListItem
    
    lv1.ColumnHeaders.Clear
    b = Array("Раса", "Пол", "Дуэлей", "Убийств", "Карма", "Возраст", "Посл. вход", "Наиграно")
    
    Set a = lv1.ColumnHeaders.Add()
        a.Text = "Ник"
        a.Width = CCur(GetSetting("l2wnet spy", "ColumnSizes", a.Text, 1440))
    
    Set a = lv1.ColumnHeaders.Add()
        a.Text = "Профессия"
        a.Width = CCur(GetSetting("l2wnet spy", "ColumnSizes", a.Text, 1400))
    
  
    Set a = lv1.ColumnHeaders.Add()
        a.Text = "Статус"
        a.Width = CCur(GetSetting("l2wnet spy", "ColumnSizes", a.Text, 1000))
    
    For Each c In b
      If mnuCSel(d).Checked Then
        Set a = lv1.ColumnHeaders.Add()
            a.Text = c
            a.Width = CCur(GetSetting("l2wnet spy", "ColumnSizes", c, 1440))
      End If
      d = d + 1
    Next c
    
    For Each e In lv1.ListItems
        Call UpdateItemD(e, CharBuffer(e.index), lv1)
    Next e
    
End Sub

