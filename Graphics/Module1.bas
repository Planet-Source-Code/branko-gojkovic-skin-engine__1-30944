Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Enum eDugme
    Da = 1
    Ne = 2
    Ok = 0
End Enum

Public Enum eDugmad
    DaNe = 1
    Ok = 0
End Enum

Public Enum eZnak
    Info = 0
    Pitanje = 1
    Upozorenje = 2
    Kritièno = 3
End Enum

Public fMainForm As frmMain
Public gOdgovor As eDugme
Public gstrBaza As String
Public Const gstrPwd  As String = ";pwd=doga"
Public ws As Workspace
Public db As Database





Sub Main()
    LoadSkinSettings
    frmSplash.Show
    frmSplash.Refresh
    frmScreen.Show
    If CanPlaySound > 0 Then gSound = True
    gstrBaza = App.Path & "\bgsoft1.mdb"
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash
    fMainForm.Show 1
End Sub

Public Sub Pomeri(Who As Form)
    Call ReleaseCapture
    Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Function msg(ByVal Tekst As String, ByVal myZnak As eZnak, Optional ByVal mVrsta As eDugmad, Optional ByVal myTitle As String) As eDugme
    '
    Dim uvp As Integer
    Dim uvI As Integer
    
    If Len(myTitle) > 0 Then
        frmPoruka.Caption = myTitle
    Else
        frmPoruka.Caption = App.Title
    End If
    
    frmPoruka.pnlTitle.Caption = Tekst
    frmPoruka.pnlTitle.ForeColor = Title.ColorC
    frmPoruka.pnlTitle.Refresh
    '
    With frmPoruka
    .Image1.Picture = .IL.ListImages(myZnak + 1).Picture
    
    uvp = .pnlTitle.Top + .pnlTitle.Height
    uvI = .Image1.Top + .Image1.Height
    
    If mVrsta > 0 Then
        .cmd(0).Visible = False
        .cmd(1).Visible = True
        .cmd(2).Visible = True
        If uvp >= uvI Then
            .cmd(1).Top = uvp + 100
            .cmd(2).Top = uvp + 100
        Else
            .cmd(1).Top = uvI + 100
            .cmd(2).Top = uvI + 100
        End If
        
        .cmd(1).Left = 300
        .cmd(2).Left = .Width - (.cmd(2).Width + 300)
        .Height = .cmd(1).Top + .cmd(1).Height + 200
    Else
        .cmd(0).Visible = True
        .cmd(1).Visible = False
        .cmd(2).Visible = False
        If uvp >= uvI Then
            .cmd(0).Top = uvp + 100
        Else
            .cmd(0).Top = uvI + 100
        End If
        .cmd(0).Left = (.Width - .cmd(0).Width) / 2
        .Height = .cmd(0).Top + .cmd(0).Height + 200
    End If
    
    End With
    frmPoruka.Show 1
    msg = gOdgovor
    
End Function
