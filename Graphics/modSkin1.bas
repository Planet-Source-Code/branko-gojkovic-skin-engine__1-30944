Attribute VB_Name = "modSkin1"
Option Explicit

Public gSkinPicture As Picture
Public SourceHdc As Long
Public gV As Integer
Public gŠ As Integer
Public mSkin As String
Public iniFajl As String


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Enum eStanje
    Aktivan = 0
    Neaktivan = 1
End Enum

Private Type eSkinTitle
    Top As Byte
    ColorA As Long
    ColorB As Long
    ColorC As Long
    ColorD As Long
    ColorEnabled As Long
    ColorDisabled As Long
    CloseColor As Long
    CloseTitle As Byte
    FontSize As Byte
    Back As Byte
    BackColor As Long
    AutoSize As Byte
End Type


Private Type mSkinPart
    Width As Integer
    Height As Integer
    X As Integer
    Y As Integer
End Type

Type mTaèka
    X As Integer
    Y As Integer
    Boja As Long
End Type

Private LCUP As mSkinPart
Private L As mSkinPart
Private LCD As mSkinPart
Private TL As mSkinPart
Public Title As eSkinTitle
Private TitleBack As mSkinPart
Private TitleBackEnd As mSkinPart
Private RightCUP As mSkinPart
Public mClose As mSkinPart
Private mTitle As mSkinPart
Private mRight As mSkinPart
Private mRightCD As mSkinPart
Private mBottom As mSkinPart
Public mBack As mSkinPart
Public mDugme As mSkinPart
Public mOption As mSkinPart
Public mCheck As mSkinPart




Public Sub GetSkin()
    mSkin = GetSetting(App.Title, "Skin", "SkinName", "SigmaPro")
    iniFajl = App.Path & "\Skins\" & mSkin & ".ini"
End Sub

Public Sub SetSkin(ByVal mNaziv As String)
    SaveSetting App.Title, "Skin", "SkinNAme", mNaziv
    mSkin = mNaziv
    iniFajl = App.Path & "\Skins\" & mSkin & ".ini"
End Sub

Sub LoadSkinSettings()
    '
    GetSkin
    'mSkin = GetSetting(App.Title, "Skin", "SkinName", "SigmaPro")
    'iniFajl = App.Path & "\Skins\" & mSkin & ".ini"
    
    Set gSkinPicture = LoadPicture(App.Path & "\Skins\" & mSkin & ".bmp")
    
    Set frmScreen.picSource.Picture = gSkinPicture
    SourceHdc = frmScreen.picSource.hdc
    
    
    'LeviKorner (Gornji)
    LCUP.Width = GetPrivateProfileInt("LeftCornerUP", "Width", "0", iniFajl)
    LCUP.Height = GetPrivateProfileInt("LeftCornerUP", "Height", "0", iniFajl)
    LCUP.X = GetPrivateProfileInt("LeftCornerUP", "X", "0", iniFajl)
    LCUP.Y = GetPrivateProfileInt("LeftCornerUP", "Y", "0", iniFajl)
    'Levaivica
    L.Width = GetPrivateProfileInt("Left", "Width", "0", iniFajl)
    L.Height = GetPrivateProfileInt("Left", "Height", "0", iniFajl)
    L.X = GetPrivateProfileInt("Left", "X", "0", iniFajl)
    L.Y = GetPrivateProfileInt("Left", "Y", "0", iniFajl)
    'LeviKorner (Donji)
    LCD.Height = GetPrivateProfileInt("LeftCornerDOWN", "Height", "0", iniFajl)
    LCD.Width = GetPrivateProfileInt("LeftCornerDOWN", "Width", "0", iniFajl)
    LCD.X = GetPrivateProfileInt("LeftCornerDOWN", "X", "0", iniFajl)
    LCD.Y = GetPrivateProfileInt("LeftCornerDOWN", "Y", "0", iniFajl)
    'TitleLeft
    TL.Height = GetPrivateProfileInt("TitleLeft", "Height", "0", iniFajl)
    TL.Width = GetPrivateProfileInt("TitleLeft", "Width", "0", iniFajl)
    TL.X = GetPrivateProfileInt("TitleLeft", "X", "0", iniFajl)
    TL.Y = GetPrivateProfileInt("TitleLeft", "Y", "0", iniFajl)
    'Title
    Title.Top = CByte(GetPrivateProfileInt("FormTitle", "Top", "0", iniFajl))
    Title.ColorA = GetPrivateProfileInt("FormTitle", "ColorA", "0", iniFajl)
    Title.ColorB = GetPrivateProfileInt("FormTitle", "ColorB", "0", iniFajl)
    Title.ColorC = GetPrivateProfileInt("FormTitle", "ColorC", "0", iniFajl)
    Title.ColorD = GetPrivateProfileInt("FormTitle", "ColorD", "0", iniFajl)
    Title.ColorEnabled = GetPrivateProfileInt("FormTitle", "ColorEnabled", "0", iniFajl)
    Title.ColorDisabled = GetPrivateProfileInt("FormTitle", "ColorDisabled", "0", iniFajl)
    Title.FontSize = CByte(GetPrivateProfileInt("FormTitle", "FontSize", "0", iniFajl))
    Title.Back = CByte(GetPrivateProfileInt("FormTitle", "Back", "0", iniFajl))
    Title.BackColor = GetPrivateProfileInt("FormTitle", "BackColor", "0", iniFajl)
    Title.AutoSize = CByte(GetPrivateProfileInt("FormTitle", "AutoSize", "1", iniFajl))
    
    'TitleBack
    If Title.Back = 1 Then
        TitleBack.Height = GetPrivateProfileInt("TitleBack", "Height", "0", iniFajl)
        TitleBack.Width = GetPrivateProfileInt("TitleBack", "Width", "0", iniFajl)
        TitleBack.X = GetPrivateProfileInt("TitleBack", "X", "0", iniFajl)
        TitleBack.Y = GetPrivateProfileInt("TitleBack", "Y", "0", iniFajl)
    End If
    
    'TitleBackEnd
    TitleBackEnd.Height = GetPrivateProfileInt("TitleBackEnd", "Height", "0", iniFajl)
    TitleBackEnd.Width = GetPrivateProfileInt("TitleBackEnd", "Width", "0", iniFajl)
    TitleBackEnd.X = GetPrivateProfileInt("TitleBackEnd", "X", "0", iniFajl)
    TitleBackEnd.Y = GetPrivateProfileInt("TitleBackEnd", "Y", "0", iniFajl)
    
    'TitleBar
    mTitle.Height = GetPrivateProfileInt("Title", "Height", "0", iniFajl)
    mTitle.Width = GetPrivateProfileInt("Title", "Width", "0", iniFajl)
    mTitle.X = GetPrivateProfileInt("Title", "X", "0", iniFajl)
    mTitle.Y = GetPrivateProfileInt("Title", "Y", "0", iniFajl)
    
    'Close
    mClose.Height = GetPrivateProfileInt("Close", "Height", "0", iniFajl)
    mClose.Width = GetPrivateProfileInt("Close", "Width", "0", iniFajl)
    mClose.X = GetPrivateProfileInt("Close", "X", "0", iniFajl)
    mClose.Y = GetPrivateProfileInt("Close", "Y", "0", iniFajl)
    gV = mClose.Height * Screen.TwipsPerPixelY
    
    'RightCornerUp
    RightCUP.Height = GetPrivateProfileInt("RightCornerUP", "Height", "0", iniFajl)
    RightCUP.Width = GetPrivateProfileInt("RightCornerUP", "Width", "0", iniFajl)
    RightCUP.X = GetPrivateProfileInt("RightCornerUP", "X", "0", iniFajl)
    RightCUP.Y = GetPrivateProfileInt("RightCornerUP", "Y", "0", iniFajl)
    
    'Right
    mRight.Height = GetPrivateProfileInt("Right", "Height", "0", iniFajl)
    mRight.Width = GetPrivateProfileInt("Right", "Width", "0", iniFajl)
    mRight.X = GetPrivateProfileInt("Right", "X", "0", iniFajl)
    mRight.Y = GetPrivateProfileInt("Right", "Y", "0", iniFajl)
    
    'RightCD
    mRightCD.Height = GetPrivateProfileInt("RightCornerDOWN", "Height", "0", iniFajl)
    mRightCD.Width = GetPrivateProfileInt("RightCornerDOWN", "Width", "0", iniFajl)
    mRightCD.X = GetPrivateProfileInt("RightCornerDOWN", "X", "0", iniFajl)
    mRightCD.Y = GetPrivateProfileInt("RightCornerDOWN", "Y", "0", iniFajl)
    
    'Bootom
    mBottom.Height = GetPrivateProfileInt("Bottom", "Height", "0", iniFajl)
    mBottom.Width = GetPrivateProfileInt("Bottom", "Width", "0", iniFajl)
    mBottom.X = GetPrivateProfileInt("Bottom", "X", "0", iniFajl)
    mBottom.Y = GetPrivateProfileInt("Bottom", "Y", "0", iniFajl)
    
    'Pozadina
    mBack.Height = GetPrivateProfileInt("Back", "Height", "0", iniFajl)
    mBack.Width = GetPrivateProfileInt("Back", "Width", "0", iniFajl)
    mBack.X = GetPrivateProfileInt("Back", "X", "0", iniFajl)
    mBack.Y = GetPrivateProfileInt("Back", "Y", "0", iniFajl)
    
    'Dugme
    mDugme.Height = GetPrivateProfileInt("Button", "Height", "0", iniFajl)
    mDugme.Width = GetPrivateProfileInt("Button", "Width", "0", iniFajl)
    mDugme.X = GetPrivateProfileInt("Button", "X", "0", iniFajl)
    mDugme.Y = GetPrivateProfileInt("Button", "Y", "0", iniFajl)
    
    'Option
    mOption.Height = GetPrivateProfileInt("Radio", "Height", "0", iniFajl)
    mOption.Width = GetPrivateProfileInt("Radio", "Width", "0", iniFajl)
    mOption.X = GetPrivateProfileInt("Radio", "X", "0", iniFajl)
    mOption.Y = GetPrivateProfileInt("Radio", "Y", "0", iniFajl)
    
    'Check
    mCheck.Height = GetPrivateProfileInt("Check", "Height", "0", iniFajl)
    mCheck.Width = GetPrivateProfileInt("Check", "Width", "0", iniFajl)
    mCheck.X = GetPrivateProfileInt("Check", "X", "0", iniFajl)
    mCheck.Y = GetPrivateProfileInt("Check", "Y", "0", iniFajl)
    
    '
    Set gSkinPicture = Nothing
    '
End Sub

Public Sub LoadSkin(F As Form, Stanje As eStanje, PrviPut As Boolean, Optional ByVal Pozadina As Integer)
    '
    Dim uŠirina As Integer
    Dim uVisina As Integer
    Dim Širina As Integer
    Dim Visina As Integer
    Dim TitleLeft As Integer
    Dim TitleWidth As Integer
    Dim cmdClose As Object
    
    uŠirina = 0
    uVisina = 0
    Širina = F.Width / Screen.TwipsPerPixelX
    Visina = F.Height / Screen.TwipsPerPixelY
     
    
    If Pozadina = 1 And PrviPut = True Then StretchBlt F.hdc, L.Width, TL.Height, Širina, Visina, SourceHdc, mBack.X, mBack.Y, mBack.Width, mBack.Height, vbSrcCopy
     
    If Stanje = Aktivan Then
        
        'LeftCornerUp
        BitBlt F.hdc, uŠirina, uVisina, LCUP.Width, LCUP.Height, SourceHdc, LCUP.X, LCUP.Y, vbSrcCopy
        uŠirina = uŠirina + LCUP.Width
        uVisina = uVisina + LCUP.Height
        
        'Left
        StretchBlt F.hdc, 0, uVisina, L.Width, Visina, SourceHdc, L.X, L.Y, L.Width, L.Height, vbSrcCopy
    
        'LeftCornerDown
        BitBlt F.hdc, 0, Visina - LCD.Height, LCD.Width, LCD.Height, SourceHdc, LCD.X, LCD.Y, vbSrcCopy
    
        'TitleLeft
        BitBlt F.hdc, uŠirina, 0, TL.Width, TL.Height, SourceHdc, TL.X, TL.Y, vbSrcCopy
            uŠirina = uŠirina + TL.Width
            TitleLeft = uŠirina * Screen.TwipsPerPixelX
    
        If PrviPut = True Then
     
            F.Controls.Add "vb.label", "lblTitle"
            With F!lbltitle
                .Caption = F.Caption
                .FontSize = Title.FontSize
                .Top = Title.Top
                .BackStyle = 0
                If Title.AutoSize = 1 Then
                    .AutoSize = True
                    .Left = TitleLeft
                Else
                    .Left = TitleLeft
                    .Width = F.Width - (TitleLeft + (TitleBackEnd.Width * Screen.TwipsPerPixelX) + (RightCUP.Width * Screen.TwipsPerPixelX) + (mClose.Width * Screen.TwipsPerPixelX))
                    .Alignment = 2
                End If
                .Visible = True
            End With
        End If
            F!lbltitle.ForeColor = Title.ColorA
            TitleWidth = F!lbltitle.Width / Screen.TwipsPerPixelX
        'PozadinaNaslova
        If Title.Back = 1 Then StretchBlt F.hdc, uŠirina, 0, TitleWidth, TitleBack.Height, SourceHdc, TitleBack.X, TitleBack.Y, TitleBack.Width, TitleBack.Height, vbSrcCopy
        uŠirina = uŠirina + TitleWidth
        'Kraj naslova
        BitBlt F.hdc, uŠirina, 0, TitleBackEnd.Width, TitleBackEnd.Height, SourceHdc, TitleBackEnd.X, TitleBackEnd.Y, vbSrcCopy
        uŠirina = uŠirina + TitleBackEnd.Width
        'TitleBar
        StretchBlt F.hdc, uŠirina, 0, Širina - uŠirina, mTitle.Height, SourceHdc, mTitle.X, mTitle.Y, mTitle.Width, mTitle.Height, vbSrcCopy
        'Close
        BitBlt F.hdc, Širina - (mClose.Width + RightCUP.Width), 0, mClose.Width, mClose.Height, SourceHdc, mClose.X, mClose.Y, vbSrcCopy
        If PrviPut = True Then
            'Set cmdClose = F.Controls.Add("graphics.bgClose", "cmdClose")
            'F!cmdClose.Visible = True
            F!cmdClose.Width = mClose.Width * Screen.TwipsPerPixelX
            F!cmdClose.Height = mClose.Height * Screen.TwipsPerPixelY
            F!cmdClose.Top = 0
            F!cmdClose.Left = (Širina - (mClose.Width + RightCUP.Width)) * Screen.TwipsPerPixelX
        Else
            F!cmdClose.Ena
        End If
        
        'Desni gornji ugao
        BitBlt F.hdc, Širina - RightCUP.Width, 0, RightCUP.Width, RightCUP.Height, SourceHdc, RightCUP.X, RightCUP.Y, vbSrcCopy
        'Desna ivica
        StretchBlt F.hdc, Širina - mRight.Width, uVisina, mRight.Width, Visina - uVisina, SourceHdc, mRight.X, mRight.Y, mRight.Width, mRight.Height, vbSrcCopy
        'Donja linija
        StretchBlt F.hdc, LCD.Width, Visina - mBottom.Height, Širina - (LCD.Width + mRightCD.Width), mBottom.Height, SourceHdc, mBottom.X, mBottom.Y, mBottom.Width, mBottom.Height, vbSrcCopy
        'Donji Korner
        BitBlt F.hdc, Širina - mRightCD.Width, Visina - mRightCD.Height, mRightCD.Width, mRightCD.Height, SourceHdc, mRightCD.X, mRightCD.Y, vbSrcCopy
    Else
        'LeftCornerUp
        BitBlt F.hdc, uŠirina, uVisina, LCUP.Width, LCUP.Height, SourceHdc, LCUP.X + LCUP.Width, LCUP.Y, vbSrcCopy
        uŠirina = uŠirina + LCUP.Width
        uVisina = uVisina + LCUP.Height
        
        'Left
        StretchBlt F.hdc, 0, uVisina, L.Width, Visina, SourceHdc, L.X + L.Width, L.Y, L.Width, L.Height, vbSrcCopy
    
        'LeftCornerDown
        BitBlt F.hdc, 0, Visina - LCD.Height, LCD.Width, LCD.Height, SourceHdc, LCD.X + LCD.Width, LCD.Y, vbSrcCopy
    
        'TitleLeft
        BitBlt F.hdc, uŠirina, 0, TL.Width, TL.Height, SourceHdc, TL.X, TL.Y + TL.Height, vbSrcCopy
            uŠirina = uŠirina + TL.Width
            TitleLeft = uŠirina * Screen.TwipsPerPixelX
            F!lbltitle.ForeColor = Title.ColorB
            TitleWidth = F!lbltitle.Width / Screen.TwipsPerPixelX
        'PozadinaNaslova
        If Title.Back = 1 Then StretchBlt F.hdc, uŠirina, 0, TitleWidth, TitleBack.Height, SourceHdc, TitleBack.X, TitleBack.Y + TitleBack.Height, TitleBack.Width, TitleBack.Height, vbSrcCopy
        uŠirina = uŠirina + TitleWidth
        'Kraj naslova
        BitBlt F.hdc, uŠirina, 0, TitleBackEnd.Width, TitleBackEnd.Height, SourceHdc, TitleBackEnd.X, TitleBackEnd.Y + TitleBackEnd.Height, vbSrcCopy
        uŠirina = uŠirina + TitleBackEnd.Width
        'TitleBar
        StretchBlt F.hdc, uŠirina, 0, Širina - uŠirina, mTitle.Height, SourceHdc, mTitle.X, mTitle.Y + mTitle.Height, mTitle.Width, mTitle.Height, vbSrcCopy
        'Close
        BitBlt F.hdc, Širina - (mClose.Width + RightCUP.Width), 0, mClose.Width, mClose.Height, SourceHdc, mClose.X, mClose.Y + mClose.Height, vbSrcCopy
        F!cmdClose.Dis
        'Desni gornji ugao
        BitBlt F.hdc, Širina - RightCUP.Width, 0, RightCUP.Width, RightCUP.Height, SourceHdc, RightCUP.X + RightCUP.Width, RightCUP.Y, vbSrcCopy
        'Desna ivica
        StretchBlt F.hdc, Širina - mRight.Width, uVisina, mRight.Width, Visina - uVisina, SourceHdc, mRight.X + mRight.Width, mRight.Y, mRight.Width, mRight.Height, vbSrcCopy
        'Donja linija
        StretchBlt F.hdc, LCD.Width, Visina - mBottom.Height, Širina - (LCD.Width + mRightCD.Width), mBottom.Height, SourceHdc, mBottom.X, mBottom.Y + mBottom.Height, mBottom.Width, mBottom.Height, vbSrcCopy
        'Donji Korner
        BitBlt F.hdc, Širina - mRightCD.Width, Visina - mRightCD.Height, mRightCD.Width, mRightCD.Height, SourceHdc, mRightCD.X + mRightCD.Width, mRightCD.Y, vbSrcCopy
    
    End If
    gŠ = F!cmdClose.Left
    '
    F.Refresh
    '
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
    vReda = mTitle.Height
    
    uT = Redovi * vReda
    ReDim t(1 To uT)
    D = 20
    Petica = D * vReda
    
    k = 1
    
    F.AutoRedraw = False
    For i = 1 To Redovi
        If i = D + 1 Then O = True
        For c = 1 To vReda
            t(k).X = i
            t(k).Y = c
            t(k).Boja = GetPixel(F.hdc, i, c)
            'Call SetPixel(F.hdc, i, c, &HC0FFFF + i)
            Call SetPixel(F.hdc, i, c, &HFFFF&)
            If O = True Then Call SetPixel(F.hdc, t(k - Petica).X, c, t(k - Petica).Boja)
            k = k + 1
        Next c
    Next i
    F.AutoRedraw = True
    Erase t
End Sub

