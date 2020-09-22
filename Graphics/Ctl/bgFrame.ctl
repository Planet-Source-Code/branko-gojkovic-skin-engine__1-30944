VERSION 5.00
Begin VB.UserControl bgFrame 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1710
   ScaleWidth      =   2055
   ToolboxBitmap   =   "bgFrame.ctx":0000
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   435
   End
End
Attribute VB_Name = "bgFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ucW As Integer
Dim ucH As Integer
Dim mvarCaption As String
Dim mvarFont As Font
Dim mvarEnabled As Boolean


Private Sub Iscrtaj()
    '
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
    StretchBlt UserControl.hdc, 4, 4, ucW - 8, ucH - 8, SourceHdc, mBack.X, mBack.Y, mBack.Width, mBack.Height, vbSrcCopy
    
End Sub

Private Sub UserControl_InitProperties()
    '
    Enabled = True
    Set Font = Ambient.Font
    Caption = Extender.Name
    '
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp: SendKeys "+{tab}"
        Case vbKeyDown: SendKeys "{tab}"
    End Select
End Sub

Private Sub UserControl_Resize()
    ucW = UserControl.Width / Screen.TwipsPerPixelX
    ucH = UserControl.Height / Screen.TwipsPerPixelY
    Iscrtaj
End Sub


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
