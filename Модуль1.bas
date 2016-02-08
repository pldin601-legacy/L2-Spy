Attribute VB_Name = "Модуль1"
Public Nicks() As String
Public HotClan As String

Public CurrClan As String
Public CurrOnline As Boolean

Public Enum CharOnline
    coIsOffline = 0
    coIsOnline = 1
    coIsNone = 2
    coServerDown = 3
End Enum

Public Type CharConstructive
    ccNick As String
    ccOnline As CharOnline
    ccProf As String
    ccRas As String
    ccGender As String
    ccPVP As String
    ccPK As String
    ccKarma As String
    ccBorn As String
    ccLastIn As String
    ccPlayed As String
End Type

Dim BlnNotify As Boolean

Sub SaveNicks()
    
    Dim N As String
    N = Join(Nicks, vbCrLf)
    Call SaveFile(LowPath(App.Path) & "nicks", N)
    
End Sub

Sub LoadNicks()

    Dim N As String
    N = LoadFile(LowPath(App.Path) & "nicks")
    Nicks = Split(N, vbCrLf)

End Sub

Public Function IsCharOnline(inChar As String, ccS As CharConstructive) As Boolean

On Error Resume Next


Dim tmpstr As String, iProof As String, L As Integer, M As Integer, N As Integer
Dim zA As Integer, zB As Integer, zC As Integer, zTime As String

Set kach = New clsKachalka
tmpstr = kach.DownloadToString("http://l2w2new.wnet.ua/UserPanel/ShowChar.asp?action=show&name=" & inChar)
Set kach = Nothing

' tmpstr = LoadFile("D:\muse.txt")

If tmpstr = "" Then ccS.ccOnline = coServerDown: Exit Function

L = InStr(tmpstr, "&nbsp;Профессия"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccProf = Mid(tmpstr, M + 4, N - M - 4)

L = InStr(tmpstr, "&nbsp;Раса"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccRas = Mid(tmpstr, M + 4, N - M - 4)

L = InStr(tmpstr, "&nbsp;Пол"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccGender = Mid(tmpstr, M + 4, N - M - 4)

L = InStr(tmpstr, "&nbsp;Дуэлей"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccPVP = Mid(tmpstr, M + 4, N - M - 4)

L = InStr(tmpstr, "&nbsp;Убийств"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccPK = Mid(tmpstr, M + 4, N - M - 4)

L = InStr(tmpstr, "&nbsp;Карма"): M = InStr(InStr(L, tmpstr, "<td>") + 5, tmpstr, ">"): N = InStr(M, tmpstr, "<")
ccS.ccKarma = Mid(tmpstr, M + 1, N - M - 1)

L = InStr(tmpstr, "&nbsp;Дата создания персонажа"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccBorn = DateBorn(Mid(tmpstr, M + 4, N - M - 4))

L = InStr(tmpstr, "&nbsp;Дата последнего захода в игру"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccLastIn = DateX(DateDiff("s", Mid(tmpstr, M + 4, N - M - 4), Now))

L = InStr(tmpstr, "&nbsp;Всего наиграно времени"): M = InStr(L, tmpstr, "<td>"): N = InStr(M, tmpstr, "</td>")
ccS.ccPlayed = Mid(tmpstr, M + 4, N - M - 4)

If InStr(tmpstr, "Персонаж в игре") Then
    ccS.ccOnline = coIsOnline
ElseIf InStr(tmpstr, "Персонаж не в игре") Then
    ccS.ccOnline = coIsOffline
ElseIf InStr(tmpstr, "Персонажа с таким именем не существует") Then
    ccS.ccOnline = coIsNone
End If

End Function


Sub UpdateItemD(inListItem As ListItem, inMass As CharConstructive, inListBox As ListView)
    Dim b() As Variant
    b = Array("Профессия", "Раса", "Пол", "Дуэлей", "Убийств", "Карма", "Возраст", "Посл. вход", "Наиграно")
    
    UpdateItm inListItem, b(0), inMass.ccProf, inListBox
    UpdateItm inListItem, b(1), inMass.ccRas, inListBox
    UpdateItm inListItem, b(2), inMass.ccGender, inListBox
    UpdateItm inListItem, b(3), inMass.ccPVP, inListBox
    UpdateItm inListItem, b(4), inMass.ccPK, inListBox
    UpdateItm inListItem, b(5), inMass.ccKarma, inListBox
    UpdateItm inListItem, b(6), inMass.ccBorn, inListBox
    UpdateItm inListItem, b(7), inMass.ccLastIn, inListBox
    UpdateItm inListItem, b(8), inMass.ccPlayed, inListBox
    
End Sub

Sub UpdateItm(inListItem As ListItem, inItem As Variant, inValue As String, inListBox As ListView)
    Dim u As Integer
    For u = 1 To inListBox.ColumnHeaders.Count
        If inListBox.ColumnHeaders(u).Text = inItem Then
            inListItem.SubItems(inListBox.ColumnHeaders(u).index - 1) = inValue
            Exit For
        End If
    Next u
End Sub

Sub RemoveItemBuffer(Expression() As CharConstructive, RemIndex As Integer)
    Dim z() As CharConstructive, i As Integer, V As Integer
    ReDim z(1 To UBound(Expression) - 1) As CharConstructive
    
    For V = 1 To UBound(Expression)
        If V <> RemIndex Then
            i = i + 1
            z(i) = Expression(V)
        End If
    Next V
    
    ReDim Expression(1 To UBound(z)) As CharConstructive
    Expression() = z()
End Sub
