Attribute VB_Name = "modSkin"
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Enum menumPP
    Dejavu = 0
    First = 1
End Enum

Public Enum msPozadina
    Sa = 0
    Bez = 1
End Enum

'Public Enum menumAktivnost
'    Aktivan = 0
'    Inaktivan = 1
'End Enum

Type mTaèka
    X As Integer
    Y As Integer
    Boja As Long
End Type
    

Private Type mSkinTip
    Width As Integer
    Height As Integer
    X As Integer
    Y As Integer
End Type

Private Type mskinTitle
    Top As Byte
    ColorA As Long
    ColorB As Long
    CloseColor As Long
    CloseTitle As Byte
    FontSize As Byte
    Back As Byte
    Tile As Byte
End Type


Private mLeftCUP As mSkinTip
Private mLeft As mSkinTip
Private mLeftCD As mSkinTip

Private mTitleLeft As mSkinTip
Private mTitle As mSkinTip
Private mTitleBack As mSkinTip
Private mTitleBackEnd As mSkinTip

Private mClose As mSkinTip
Private mBottom As mSkinTip

Private mRightCUP As mSkinTip
Private mRight As mSkinTip
Private mRightCD As mSkinTip

Public mBack As mSkinTip
Public mButton As mSkinTip
Public mRadio As mSkinTip

Private mskinTitle As mskinTitle




Public Sub Pomeri(Who As Form)
    Call ReleaseCapture
    Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Sub GetSkinSettings5()
    '
    Dim strINI As String
    Dim slika As String
    
    strINI = GetINI
    slika = Mid(strINI, 1, Len(strINI) - 4) & ".bmp"
    frmScreen.p.Picture = LoadPicture(slika)
    
    'Uèitavanje skin kljuèeva iz [skin].ini
    
    mLeftCUP.Height = GetPrivateProfileInt("LeftCornerUP", "Height", "0", strINI)
    mLeftCUP.Width = GetPrivateProfileInt("LeftCornerUP", "Width", "0", strINI)
    mLeftCUP.X = GetPrivateProfileInt("LeftCornerUP", "X", "0", strINI)
    mLeftCUP.Y = GetPrivateProfileInt("LeftCornerUP", "Y", "0", strINI)
    
    mLeft.Height = GetPrivateProfileInt("Left", "Height", "0", strINI)
    mLeft.Width = GetPrivateProfileInt("Left", "Width", "0", strINI)
    mLeft.X = GetPrivateProfileInt("Left", "X", "0", strINI)
    mLeft.Y = GetPrivateProfileInt("Left", "Y", "0", strINI)
    
    mLeftCD.Height = GetPrivateProfileInt("LeftCornerDOWN", "Height", "0", strINI)
    mLeftCD.Width = GetPrivateProfileInt("LeftCornerDOWN", "Width", "0", strINI)
    mLeftCD.X = GetPrivateProfileInt("LeftCornerDOWN", "X", "0", strINI)
    mLeftCD.Y = GetPrivateProfileInt("LeftCornerDOWN", "Y", "0", strINI)
    
    mTitleLeft.Height = GetPrivateProfileInt("TitleLeft", "Height", "0", strINI)
    mTitleLeft.Width = GetPrivateProfileInt("TitleLeft", "Width", "0", strINI)
    mTitleLeft.X = GetPrivateProfileInt("TitleLeft", "X", "0", strINI)
    mTitleLeft.Y = GetPrivateProfileInt("TitleLeft", "Y", "0", strINI)
    
    mTitle.Height = GetPrivateProfileInt("Title", "Height", "0", strINI)
    mTitle.Width = GetPrivateProfileInt("Title", "Width", "0", strINI)
    mTitle.X = GetPrivateProfileInt("Title", "X", "0", strINI)
    mTitle.Y = GetPrivateProfileInt("Title", "Y", "0", strINI)
    
    mTitleBack.Height = GetPrivateProfileInt("TitleBack", "Height", "0", strINI)
    mTitleBack.Width = GetPrivateProfileInt("TitleBack", "Width", "0", strINI)
    mTitleBack.X = GetPrivateProfileInt("TitleBack", "X", "0", strINI)
    mTitleBack.Y = GetPrivateProfileInt("TitleBack", "Y", "0", strINI)
    
    mTitleBackEnd.Height = GetPrivateProfileInt("TitleBackEnd", "Height", "0", strINI)
    mTitleBackEnd.Width = GetPrivateProfileInt("TitleBackEnd", "Width", "0", strINI)
    mTitleBackEnd.X = GetPrivateProfileInt("TitleBackEnd", "X", "0", strINI)
    mTitleBackEnd.Y = GetPrivateProfileInt("TitleBackEnd", "Y", "0", strINI)
    
    mRightCUP.Height = GetPrivateProfileInt("RightCornerUP", "Height", "0", strINI)
    mRightCUP.Width = GetPrivateProfileInt("RightCornerUP", "Width", "0", strINI)
    mRightCUP.X = GetPrivateProfileInt("RightCornerUP", "X", "0", strINI)
    mRightCUP.Y = GetPrivateProfileInt("RightCornerUP", "Y", "0", strINI)
    
    mRight.Height = GetPrivateProfileInt("Right", "Height", "0", strINI)
    mRight.Width = GetPrivateProfileInt("Right", "Width", "0", strINI)
    mRight.X = GetPrivateProfileInt("Right", "X", "0", strINI)
    mRight.Y = GetPrivateProfileInt("Right", "Y", "0", strINI)
    
    mRightCD.Height = GetPrivateProfileInt("RightCornerDOWN", "Height", "0", strINI)
    mRightCD.Width = GetPrivateProfileInt("RightCornerDOWN", "Width", "0", strINI)
    mRightCD.X = GetPrivateProfileInt("RightCornerDOWN", "X", "0", strINI)
    mRightCD.Y = GetPrivateProfileInt("RightCornerDOWN", "Y", "0", strINI)
    
    mClose.Height = GetPrivateProfileInt("Close", "Height", "0", strINI)
    mClose.Width = GetPrivateProfileInt("Close", "Width", "0", strINI)
    mClose.X = GetPrivateProfileInt("Close", "X", "0", strINI)
    mClose.Y = GetPrivateProfileInt("Close", "Y", "0", strINI)
    
    mBottom.Height = GetPrivateProfileInt("Bottom", "Height", "0", strINI)
    mBottom.Width = GetPrivateProfileInt("Bottom", "Width", "0", strINI)
    mBottom.X = GetPrivateProfileInt("Bottom", "X", "0", strINI)
    mBottom.Y = GetPrivateProfileInt("Bottom", "Y", "0", strINI)
        
    mBack.Height = GetPrivateProfileInt("Back", "Height", "0", strINI)
    mBack.Width = GetPrivateProfileInt("Back", "Width", "0", strINI)
    mBack.X = GetPrivateProfileInt("Back", "X", "0", strINI)
    mBack.Y = GetPrivateProfileInt("Back", "Y", "0", strINI)
    
    mskinTitle.Top = CByte(GetPrivateProfileInt("FormTitle", "Top", "0", strINI))
    mskinTitle.ColorA = GetPrivateProfileInt("FormTitle", "ColorA", "0", strINI)
    mskinTitle.ColorB = GetPrivateProfileInt("FormTitle", "ColorB", "0", strINI)
    mskinTitle.CloseColor = GetPrivateProfileInt("FormTitle", "CloseColor", "0", strINI)
    mskinTitle.CloseTitle = CByte(GetPrivateProfileInt("FormTitle", "CloseTitle", "0", strINI))
    mskinTitle.FontSize = CByte(GetPrivateProfileInt("FormTitle", "FontSize", "0", strINI))
    mskinTitle.Back = CByte(GetPrivateProfileInt("FormTitle", "Back", "0", strINI))
    mskinTitle.Tile = CByte(GetPrivateProfileInt("FormTitle", "Tile", "0", strINI))
    
    mButton.Height = GetPrivateProfileInt("Button", "Height", "0", strINI)
    mButton.Width = GetPrivateProfileInt("Button", "Width", "0", strINI)
    mButton.X = GetPrivateProfileInt("Button", "X", "0", strINI)
    mButton.Y = GetPrivateProfileInt("Button", "Y", "0", strINI)
    
    mRadio.Height = GetPrivateProfileInt("Radio", "Height", "0", strINI)
    mRadio.Width = GetPrivateProfileInt("Radio", "Width", "0", strINI)
    mRadio.X = GetPrivateProfileInt("Radio", "X", "0", strINI)
    mRadio.Y = GetPrivateProfileInt("Radio", "Y", "0", strINI)
    
End Sub

Public Sub SetSkin(ByVal SkinName As String)
    SaveSetting App.Title, "Skin", "SkinINI", App.Path & "\Skins\" & SkinName
    GetSkinSettings
End Sub

Private Function GetINI() As String
    '
    GetINI = GetSetting(App.Title, "Skin", "SkinINI", App.Path & "\Skins\sigma.ini")
    '
End Function



    
    Dim ret As Long
    Dim DeskHdc As Long
    Dim i As Long
    Dim Širina As Integer
    Dim Visina As Integer
    Dim ŠirinaNaslova As Integer
    Dim uŠirina As Integer
    Dim uVisina As Integer
    Dim lTitleLeft As Integer
    Dim lCloseLeft As Integer
    Dim Pic As Picture
    Dim ctl As Control

    
    Širina = F.Width / Screen.TwipsPerPixelX
    Visina = F.Height / Screen.TwipsPerPixelY
    DeskHdc = frmScreen.p.hDC
    
    F.lblTitle.Top = mskinTitle.Top
    F.lblTitle.FontSize = mskinTitle.FontSize
    F.lblClose.Top = mskinTitle.Top
    F.lblClose.ForeColor = mskinTitle.CloseColor
    If mskinTitle.CloseTitle = 0 Then F.lblClose.Caption = ""
    If Stanje = Aktivan Then
        F.lblTitle.ForeColor = mskinTitle.ColorA
    Else
        F.lblTitle.ForeColor = mskinTitle.ColorB
    End If
    
    F.lblTitle.Caption = F.Caption
    ŠirinaNaslova = F.lblTitle.Width / Screen.TwipsPerPixelX
    uŠirina = 0
    uVisina = 0
    
    'Pozadina forme
    If mPozadina = Sa Then
        If mskinTitle.Back = 1 And mUslov = First Then
            If mskinTitle.Tile = 0 Then
                StretchBlt F.hDC, 0, 0, Širina, Visina, DeskHdc, mBack.X, mBack.Y, mBack.Width, mBack.Height, vbSrcCopy
            Else
                
            End If
        End If
    End If
    
    Select Case Stanje
        Case 0
            'Left Corner Up
            BitBlt F.hDC, uŠirina, uVisina, mLeftCUP.Width, mLeftCUP.Height, DeskHdc, mLeftCUP.X, mLeftCUP.Y, vbSrcCopy
            uŠirina = uŠirina + mLeftCUP.Width
            uVisina = uVisina + mLeftCUP.Height
            'Left
            StretchBlt F.hDC, 0, mLeftCUP.Height, mLeft.Width, Visina, DeskHdc, mLeft.X, mLeft.Y, mLeft.Width, mLeft.Height, vbSrcCopy
            'Donji Korner
            BitBlt F.hDC, 0, Visina - mLeftCD.Height, mLeftCD.Width, mLeftCD.Height, DeskHdc, mLeftCD.X, mLeftCD.Y, vbSrcCopy
            'Title Left
            BitBlt F.hDC, uŠirina, 0, mTitleLeft.Width, mTitleLeft.Height, DeskHdc, mTitleLeft.X, mTitleLeft.Y, vbSrcCopy
            uŠirina = uŠirina + mTitleLeft.Width
            lTitleLeft = uŠirina * Screen.TwipsPerPixelX
            'PozadinaNaslova
            StretchBlt F.hDC, uŠirina, 0, ŠirinaNaslova, mTitleBack.Height, DeskHdc, mTitleBack.X, mTitleBack.Y, mTitleBack.Width, mTitleBack.Height, vbSrcCopy
            uŠirina = uŠirina + ŠirinaNaslova
            'Kraj naslova
            BitBlt F.hDC, uŠirina, 0, mTitleBackEnd.Width, mTitleBackEnd.Height, DeskHdc, mTitleBackEnd.X, mTitleBackEnd.Y, vbSrcCopy
            uŠirina = uŠirina + mTitleBackEnd.Width
            'TitleBar
            StretchBlt F.hDC, uŠirina, 0, Širina - uŠirina, mTitle.Height, DeskHdc, mTitle.X, mTitle.Y, mTitle.Width, mTitle.Height, vbSrcCopy
            'Close
            BitBlt F.hDC, Širina - (mClose.Width + mRightCUP.Width), 0, mClose.Width, mClose.Height, DeskHdc, mClose.X, mClose.Y, vbSrcCopy
            lCloseLeft = ((Širina - mClose.Width) * Screen.TwipsPerPixelX) - 40
            'Desni gornji ugao
            BitBlt F.hDC, Širina - mRightCUP.Width, 0, mRightCUP.Width, mRightCUP.Height, DeskHdc, mRightCUP.X, mRightCUP.Y, vbSrcCopy
            'Desna ivica
            StretchBlt F.hDC, Širina - mRight.Width, uVisina, mRight.Width, Visina - uVisina, DeskHdc, mRight.X, mRight.Y, mRight.Width, mRight.Height, vbSrcCopy
            'Donja linija
            StretchBlt F.hDC, mLeftCD.Width, Visina - mBottom.Height, Širina - (mLeftCD.Width + mRightCD.Width), mBottom.Height, DeskHdc, mBottom.X, mBottom.Y, mBottom.Width, mBottom.Height, vbSrcCopy
            'Donji Korner
            BitBlt F.hDC, Širina - mRightCD.Width, Visina - mRightCD.Height, mRightCD.Width, mRightCD.Height, DeskHdc, mRightCD.X, mRightCD.Y, vbSrcCopy
        Case 1
            'Left Corner Up
            BitBlt F.hDC, uŠirina, uVisina, mLeftCUP.Width, mLeftCUP.Height, DeskHdc, mLeftCUP.X + mLeftCUP.Width, mLeftCUP.Y, vbSrcCopy
            uŠirina = uŠirina + mLeftCUP.Width
            uVisina = uVisina + mLeftCUP.Height
            'Left
            StretchBlt F.hDC, 0, mLeftCUP.Height, mLeft.Width, Visina, DeskHdc, mLeft.X + mLeft.Width, mLeft.Y, mLeft.Width, mLeft.Height, vbSrcCopy
            'Donji Korner
            BitBlt F.hDC, 0, Visina - mLeftCD.Height, mLeftCD.Width, mLeftCD.Height, DeskHdc, mLeftCD.X + mLeftCD.Width, mLeftCD.Y, vbSrcCopy
            'Title Left
            BitBlt F.hDC, uŠirina, 0, mTitleLeft.Width, mTitleLeft.Height, DeskHdc, mTitleLeft.X, mTitleLeft.Y + mTitleLeft.Height, vbSrcCopy
            uŠirina = uŠirina + mTitleLeft.Width
            lTitleLeft = uŠirina * Screen.TwipsPerPixelX
            'PozadinaNaslova
            StretchBlt F.hDC, uŠirina, 0, ŠirinaNaslova, mTitleBack.Height, DeskHdc, mTitleBack.X, mTitleBack.Y + mTitleBack.Height, mTitleBack.Width, mTitleBack.Height, vbSrcCopy
            uŠirina = uŠirina + ŠirinaNaslova
            'Kraj naslova
            BitBlt F.hDC, uŠirina, 0, mTitleBackEnd.Width, mTitleBackEnd.Height, DeskHdc, mTitleBackEnd.X, mTitleBackEnd.Y + mTitleBackEnd.Height, vbSrcCopy
            uŠirina = uŠirina + mTitleBackEnd.Width
            'TitleBar
            StretchBlt F.hDC, uŠirina, 0, Širina - uŠirina, mTitle.Height, DeskHdc, mTitle.X, mTitle.Y + mTitle.Height, mTitle.Width, mTitle.Height, vbSrcCopy
            'Close
            BitBlt F.hDC, Širina - (mClose.Width + mRightCUP.Width), 0, mClose.Width, mClose.Height, DeskHdc, mClose.X, mClose.Y + mClose.Height, vbSrcCopy
            lCloseLeft = ((Širina - mClose.Width) * Screen.TwipsPerPixelX) - 40
            'Desni gornji ugao
            BitBlt F.hDC, Širina - mRightCUP.Width, 0, mRightCUP.Width, mRightCUP.Height, DeskHdc, mRightCUP.X + mRightCUP.Width, mRightCUP.Y, vbSrcCopy
            'Desna ivica
            StretchBlt F.hDC, Širina - mRight.Width, uVisina, mRight.Width, Visina - uVisina, DeskHdc, mRight.X + mRight.Width, mRight.Y, mRight.Width, mRight.Height, vbSrcCopy
            'Donja linija
            StretchBlt F.hDC, mLeftCD.Width, Visina - mBottom.Height, Širina - (mLeftCD.Width + mRightCD.Width), mBottom.Height, DeskHdc, mBottom.X, mBottom.Y + mBottom.Height, mBottom.Width, mBottom.Height, vbSrcCopy
            'Donji Korner
            BitBlt F.hDC, Širina - mRightCD.Width, Visina - mRightCD.Height, mRightCD.Width, mRightCD.Height, DeskHdc, mRightCD.X + mRightCD.Width, mRightCD.Y, vbSrcCopy
        End Select
    '
    
    'StretchBlt F.pict.hDC, 0, 0, mButton.Width, mButton.Height, DeskHdc, mButton.X + (3 * mButton.Width), mButton.Y, mButton.Width, mButton.Height, vbSrcCopy
    'SavePicture F.pict, "Dugme.bmp"
    'F.Image1.Picture = LoadPicture(F.pict.Picture)
    'F.Image1.Picture = F.pict.Picture
    F.lblClose.Left = lCloseLeft
    F.lblClose.Top = 0
    F.lblClose.Width = (mClose.Width * Screen.TwipsPerPixelX) - 20
    F.lblClose.Height = mTitle.Height * Screen.TwipsPerPixelY
    F.lblTitle.Left = lTitleLeft
    F.Refresh
    
    For Each ctl In F.Controls
        If Left$(ctl.Tag, 1) = "1" Then ctl.Refresh
    Next
End Sub


Sub Specijalka(ByVal F As Form)
    '
    
    Dim xP As Long
    Dim i As Integer
    Dim c As Integer
    Dim t() As mTaèka
    Dim k As Integer
    Dim Redovi As Integer
    Dim vReda As Integer
    Dim uT As Integer
    Dim O As Boolean
    Dim Petica As Integer
    Dim D As Integer
    
    Redovi = 1000
    vReda = mTitle.Height / 2
    
    uT = Redovi * vReda
    ReDim t(1 To uT)
    D = 80
    Petica = D * vReda
    
    k = 1
    
    F.AutoRedraw = False
    For i = 1 To Redovi
        If i = D + 1 Then O = True
        For c = 1 To vReda
            t(k).X = i
            t(k).Y = c
            t(k).Boja = GetPixel(F.hDC, i, c)
            Call SetPixel(F.hDC, i, c, &HC0FFFF + i)
            If O = True Then Call SetPixel(F.hDC, t(k - Petica).X, c, t(k - Petica).Boja)
            k = k + 1
        Next c
    Next i
    F.AutoRedraw = True
    Erase t
End Sub
