VERSION 5.00
Begin VB.Form clanPrompt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Проверить онлайн клана"
   ClientHeight    =   1590
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "clanPrompt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Онлайн"
      Height          =   195
      Left            =   900
      TabIndex        =   4
      ToolTipText     =   "Показывать только тех, кто в игре"
      Top             =   780
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Го"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Клан:"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   420
      Width           =   435
   End
End
Attribute VB_Name = "clanPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lClans() As String
Dim CCode As Byte

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub Combo1_Change()
    
    If CCode >= 32 Then Predictate

    If Combo1.Text = "" Then OKButton.Enabled = False Else OKButton.Enabled = True
    
End Sub

Sub Predictate()

    Dim itm As String
    Dim itmr As Integer
    Dim lgt As String
    
    lgt = Len(Combo1.Text)
    
    For itmr = 0 To Combo1.ListCount - 1
        itm = Combo1.List(itmr)
        If LCase(Combo1.Text) = LCase(Mid(itm, 1, lgt)) Then
            Combo1.Text = itm
            Combo1.SelStart = lgt
            Combo1.SelLength = Len(itm) - (lgt)
            Exit For
        End If
    Next itmr
    
End Sub

Private Sub Combo1_Click()

    If Combo1.Text = "" Then OKButton.Enabled = False Else OKButton.Enabled = True

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    CCode = KeyAscii

End Sub

Private Sub Form_Load()

    LoadRecent
    
End Sub


Sub SaveRecent()

    Dim sf As String, sbuff As String
    sf = LowPath(App.Path) & "clans"
    
    sbuff = Join(lClans, vbCrLf)
    
    SaveFile sf, sbuff
    
End Sub

Sub LoadRecent()

    Dim sf As String, sbuff As String
    sf = LowPath(App.Path) & "clans"
    sbuff = LoadFile(sf)
    lClans = Split(sbuff, vbCrLf)
    
    Call RefreshCombo
    
End Sub

Sub RefreshCombo()

    Dim lClan As Variant
    Combo1.Clear
    For Each lClan In lClans
        Combo1.AddItem lClan
    Next lClan
    Combo1.Text = HotClan
    
End Sub

Private Sub Form_Paint()

    Combo1.SetFocus
    
End Sub

Private Sub OKButton_Click()

    If ClanIsNew(Combo1.Text) Then
        AddToClans Combo1.Text
        SaveRecent
    End If
    
    Me.MousePointer = 13
    HotClan = Combo1.Text
    SaveSetting "l2wnet spy", "options", "Hot Clan", HotClan
    
    ' Show Clan
    If Check1.Value = 1 Then CurrOnline = True Else CurrOnline = False
    CurrClan = Combo1.Text
    Unload Me
    dlgClan.Show
    
    
End Sub

Function ClanIsNew(inClan As String) As Boolean

    Dim lClan As Variant
    For Each lClan In lClans
        If LCase(lClan) = LCase(inClan) Then ClanIsNew = False: Exit Function
    Next lClan

    ClanIsNew = True

End Function

Sub AddToClans(inClan As String)

    Dim i As Integer
    i = UBound(lClans)
    ReDim Preserve lClans(i + 1)
    lClans(i + 1) = inClan
    
End Sub
