Attribute VB_Name = "modMask"
Public Const WS_CHILD = &H40000000

Sub AttachMe(inForm As Form, Optional noMin As Boolean = False)

    Dim df As Form
    
    Set df = New frmMask
    
    Load df
    
    df.Width = inForm.Width + (8 * tppX)
    df.Height = inForm.Height + (34 * tppY)
    
    inForm.WindowState = vbNormal
    inForm.BorderStyle = 0
    inForm.Top = 30 * tppY
    inForm.Left = 4 * tppX
    
    df.Label1.Caption = inForm.Caption
    df.Caption = df.Label1.Caption
    df.Icon = inForm.Icon
    df.Image1.Picture = inForm.Icon
    df.Tag = CStr(inForm.hwnd)

    SetParent inForm.hwnd, df.hwnd
 
    If noMin Then
        df.CB(2).Visible = False
        df.CB(1).Visible = False
    End If
    
    df.Show
    DoEvents
    
End Sub
