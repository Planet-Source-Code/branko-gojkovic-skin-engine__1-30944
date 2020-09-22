VERSION 5.00
Begin VB.UserControl bgDugme 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "bgDugme.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   420
   ScaleWidth      =   945
   ToolboxBitmap   =   "bgDugme.ctx":0152
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   660
   End
End
Attribute VB_Name = "bgDugme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ucW As Integer
Dim ucH As Integer
Dim mMouse As Boolean
Dim LL As Integer
Dim LT As Integer

Private mvarCaption As String 'local copy
Private mvarEnabled As Boolean
Private mvarFont As Font

Private Enum eSt
    Normalno = 0
    Pritisnuto = 1
    Fokus = 2
    Disabled = 3
End Enum

Public Event Click()




Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call UserControl_MouseDown(1, 1, 1, 1)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(1, 1, 1, 1)
End Sub

Private Sub UserControl_GotFocus()
    If mMouse = False Then Iscrtaj Fokus
End Sub

Private Sub UserControl_Initialize()
    '
    'Caption = Extender.Name
    'lblCaption.ForeColor = Title.ColorC
    mMouse = False
    '
End Sub

Private Sub UserControl_InitProperties()
    '
    Enabled = True
    Caption = Extender.Name
    Set Font = Ambient.Font
    '
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp: SendKeys "+{tab}"
        Case vbKeyDown: SendKeys "{tab}"
        Case vbKeyReturn
            Call UserControl_MouseDown(1, 1, 1, 1)
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        Case vbKeyReturn
            Call UserControl_MouseUp(1, 1, 1, 1)
    End Select
    '
End Sub

Private Sub UserControl_LostFocus()
    '
    Iscrtaj Normalno
    '
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Exit Sub
    mMouse = True
    Cap 1
    Iscrtaj Pritisnuto
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
        Iscrtaj Fokus
        Screen.MousePointer = 0
        mMouse = False
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If Button = 2 Then Exit Sub
    Iscrtaj Fokus
    Cap 0
    If mMouse = True Then RaiseEvent Click
    mMouse = False
    '
End Sub

Private Sub UserControl_Resize()
    
    ucW = UserControl.Width / Screen.TwipsPerPixelX
    ucH = UserControl.Height / Screen.TwipsPerPixelY
    'pnl1.Width = UserControl.Width
    'pnl1.Height = UserControl.Height
    Iscrtaj Normalno
    ResizeCaption
End Sub



Private Sub ResizeCaption()
    '
    LT = (UserControl.Height - lblCaption.Height) / 2
    LL = (UserControl.Width - lblCaption.Width) / 2
    lblCaption.Top = LT
    lblCaption.Left = LL
    '
End Sub

Private Sub Cap(ByVal ind As Integer)
    Select Case ind
        Case 0
            lblCaption.Left = LL
            lblCaption.Top = LT
        Case 1
            lblCaption.Left = LL + 20
            lblCaption.Top = LT + 20
    End Select
End Sub

Private Sub Iscrtaj(ByVal mStanje As eSt)
    '
    
    
    
    Select Case mStanje
        '
        Case 0
            'LCUP
            BitBlt UserControl.hdc, 0, 0, 4, 4, SourceHdc, mDugme.X, mDugme.Y, vbSrcCopy
            'L
            StretchBlt UserControl.hdc, 0, 4, 4, ucH - 8, SourceHdc, mDugme.X, mDugme.Y + 4, 4, 4, vbSrcCopy
            'LCD
            BitBlt UserControl.hdc, 0, ucH - 4, 6, 4, SourceHdc, mDugme.X, mDugme.Y + (mDugme.Height - 4), vbSrcCopy
            'Top
            StretchBlt UserControl.hdc, 4, 0, ucW - 8, 4, SourceHdc, mDugme.X + 4, mDugme.Y, 4, 4, vbSrcCopy
            'RCUP
            BitBlt UserControl.hdc, ucW - 4, 0, 4, 6, SourceHdc, mDugme.X + (mDugme.Width - 4), mDugme.Y, vbSrcCopy
            'R
            StretchBlt UserControl.hdc, ucW - 4, 6, 4, ucH - 10, SourceHdc, mDugme.X + (mDugme.Width - 4), mDugme.Y + 6, 4, 4, vbSrcCopy
            'RCD
            BitBlt UserControl.hdc, ucW - 4, ucH - 4, 4, 4, SourceHdc, mDugme.X + (mDugme.Width - 4), mDugme.Y + (mDugme.Height - 4), vbSrcCopy
            'Bottom
            StretchBlt UserControl.hdc, 6, ucH - 4, ucW - 10, 4, SourceHdc, mDugme.X + 6, mDugme.Y + (mDugme.Height - 4), 4, 4, vbSrcCopy
            'Popuna
            StretchBlt UserControl.hdc, 4, 4, ucW - 8, ucH - 8, SourceHdc, mDugme.X + 4, mDugme.Y + 4, mDugme.Width - 8, mDugme.Height - 8, vbSrcCopy
            Cap 0
        Case 1
            'LCUP
            BitBlt UserControl.hdc, 0, 0, 4, 4, SourceHdc, mDugme.X + mDugme.Width, mDugme.Y, vbSrcCopy
            'L
            StretchBlt UserControl.hdc, 0, 4, 4, ucH - 8, SourceHdc, mDugme.X + mDugme.Width, mDugme.Y + 4, 4, 4, vbSrcCopy
            'LCD
            BitBlt UserControl.hdc, 0, ucH - 4, 6, 4, SourceHdc, mDugme.X + mDugme.Width, mDugme.Y + (mDugme.Height - 4), vbSrcCopy
            'Top
            StretchBlt UserControl.hdc, 4, 0, ucW - 8, 4, SourceHdc, mDugme.X + mDugme.Width + 4, mDugme.Y, 4, 4, vbSrcCopy
            'RCUP
            BitBlt UserControl.hdc, ucW - 4, 0, 4, 6, SourceHdc, mDugme.X + mDugme.Width + (mDugme.Width - 4), mDugme.Y, vbSrcCopy
            'R
            StretchBlt UserControl.hdc, ucW - 4, 6, 4, ucH - 10, SourceHdc, mDugme.X + mDugme.Width + (mDugme.Width - 4), mDugme.Y + 6, 4, 4, vbSrcCopy
            'RCD
            BitBlt UserControl.hdc, ucW - 4, ucH - 4, 4, 4, SourceHdc, mDugme.X + mDugme.Width + (mDugme.Width - 4), mDugme.Y + (mDugme.Height - 4), vbSrcCopy
            'Bottom
            StretchBlt UserControl.hdc, 6, ucH - 4, ucW - 10, 4, SourceHdc, mDugme.X + mDugme.Width + 6, mDugme.Y + (mDugme.Height - 4), 4, 4, vbSrcCopy
            'Popuna
            StretchBlt UserControl.hdc, 4, 4, ucW - 8, ucH - 8, SourceHdc, mDugme.X + mDugme.Width + 4, mDugme.Y + 4, mDugme.Width - 8, mDugme.Height - 8, vbSrcCopy
            Cap 1
        Case 2
            'LCUP
            BitBlt UserControl.hdc, 0, 0, 4, 4, SourceHdc, mDugme.X + (3 * mDugme.Width), mDugme.Y, vbSrcCopy
            'L
            StretchBlt UserControl.hdc, 0, 4, 4, ucH - 8, SourceHdc, mDugme.X + (3 * mDugme.Width), mDugme.Y + 4, 4, 4, vbSrcCopy
            'LCD
            BitBlt UserControl.hdc, 0, ucH - 4, 6, 4, SourceHdc, mDugme.X + (3 * mDugme.Width), mDugme.Y + (mDugme.Height - 4), vbSrcCopy
            'Top
            StretchBlt UserControl.hdc, 4, 0, ucW - 8, 4, SourceHdc, mDugme.X + (3 * mDugme.Width) + 4, mDugme.Y, 4, 4, vbSrcCopy
            'RCUP
            BitBlt UserControl.hdc, ucW - 4, 0, 4, 6, SourceHdc, mDugme.X + (3 * mDugme.Width) + (mDugme.Width - 4), mDugme.Y, vbSrcCopy
            'R
            StretchBlt UserControl.hdc, ucW - 4, 6, 4, ucH - 10, SourceHdc, mDugme.X + (3 * mDugme.Width) + (mDugme.Width - 4), mDugme.Y + 6, 4, 4, vbSrcCopy
            'RCD
            BitBlt UserControl.hdc, ucW - 4, ucH - 4, 4, 4, SourceHdc, mDugme.X + (3 * mDugme.Width) + (mDugme.Width - 4), mDugme.Y + (mDugme.Height - 4), vbSrcCopy
            'Bottom
            StretchBlt UserControl.hdc, 6, ucH - 4, ucW - 10, 4, SourceHdc, mDugme.X + (3 * mDugme.Width) + 6, mDugme.Y + (mDugme.Height - 4), 4, 4, vbSrcCopy
            'Popuna
            StretchBlt UserControl.hdc, 4, 4, ucW - 8, ucH - 8, SourceHdc, mDugme.X + (3 * mDugme.Width) + 4, mDugme.Y + 4, mDugme.Width - 8, mDugme.Height - 8, vbSrcCopy
            Cap 0
    End Select
    UserControl.Refresh
End Sub

'Properties

Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property


Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    
    mvarCaption = vData
    lblCaption.Caption = mvarCaption
    lblCaption.ForeColor = Title.ColorC
    ResizeCaption
    lblCaption.Refresh
    PropertyChanged ("Caption")


End Property


Public Property Get Font() As Font
 Set Font = mvarFont
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set mvarFont = NewFont
    Set UserControl.Font = mvarFont
    With lblCaption
        .Font.Name = mvarFont.Name
        .Font.Size = mvarFont.Size
        .Font.Bold = mvarFont.Bold
        .Font.Italic = mvarFont.Italic
        .Font.Strikethrough = mvarFont.Strikethrough
        .Font.Underline = mvarFont.Underline
        .Font.Charset = mvarFont.Charset
        .Refresh
    End With
    PropertyChanged "Font"
    Call ResizeCaption
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Property Let Enabled(ByVal NewBool As Boolean)

    mvarEnabled = NewBool

    If mvarEnabled = False Then
        lblCaption.ForeColor = Title.ColorD
        lblCaption.Refresh
    Else
        lblCaption.ForeColor = Title.ColorC
        lblCaption.Refresh
    End If
    PropertyChanged "Enabled"
    UserControl.Enabled = mvarEnabled
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    mvarCaption = PropBag.ReadProperty("Caption", Extender.Name)
    Set mvarFont = PropBag.ReadProperty("Font", Ambient.Font)
    mvarEnabled = PropBag.ReadProperty("Enabled", True)
    
    
    Caption = mvarCaption
    Set Font = mvarFont
    Enabled = mvarEnabled
    '
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    Call PropBag.WriteProperty("Caption", mvarCaption, Extender.Name)
    Call PropBag.WriteProperty("Font", mvarFont, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", mvarEnabled, True)
    '
End Sub

