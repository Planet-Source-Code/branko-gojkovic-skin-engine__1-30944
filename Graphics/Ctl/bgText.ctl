VERSION 5.00
Begin VB.UserControl bgText 
   BackColor       =   &H00E7EBEF&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   375
   ScaleWidth      =   1230
   Begin VB.TextBox txtText 
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   0
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "BG soft"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "bgText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




'Local variables
Private mvarText As String
Private mvarEnabled As Boolean
Private mvarFont As Font
Private mvarAlign As eAlign
Private mvarPassWordChar As String
Private mvarMaxLenght As Integer
Private mvarAppearance As eAppearance
Private mvarBorderStyle As eBorderStyle

'User Control Events
Public Event KeyPress(KeyAscii As Integer)







'********************************************
'********************************************
'********************************************
'User Control Custom Properties


Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "122c"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Text = mvarText
    '
End Property


Public Property Let Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    
    mvarText = vData
    txtText.Text = mvarText
    txtText.Refresh
    
    PropertyChanged ("Text")
    '
End Property

Public Property Get Font() As Font
    '
    Set Font = mvarFont
    '
End Property

Public Property Set Font(ByVal NewFont As Font)
    '
    Set mvarFont = NewFont
    Set UserControl.Font = mvarFont
    With txtText
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
    UserControl.Width = txtText.Width
    '
End Property

Public Property Get Enabled() As Boolean
    '
    Enabled = mvarEnabled
    '
End Property

Public Property Let Enabled(ByVal NewBool As Boolean)
    '
    mvarEnabled = NewBool

    If mvarEnabled = False Then
        txtText.ForeColor = Title.ColorDisabled
        txtText.Refresh
    Else
        txtText.ForeColor = Title.ColorEnabled
        txtText.Refresh
    End If
    '
    PropertyChanged "Enabled"
    UserControl.Enabled = mvarEnabled
    '
End Property

Public Property Let Align(ByVal vData As eAlign)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarAlign = vData
    txtText.Alignment = mvarAlign
    PropertyChanged "Align"
End Property


Public Property Get Align() As eAlign
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Align = mvarAlign
End Property



Public Property Get PasswordChar() As String
    PasswordChar = mvarPassWordChar
End Property

Public Property Let PasswordChar(ByVal vData As String)
    mvarPassWordChar = vData
    txtText.PasswordChar = mvarPassWordChar
    PropertyChanged "PasswordChar"
End Property



Public Property Get MaxLenght() As Integer
    MaxLenght = mvarMaxLenght
End Property

Public Property Let MaxLenght(ByVal vData As Integer)
    mvarMaxLenght = vData
    txtText.MaxLength = mvarMaxLenght
    PropertyChanged "MaxLenght"
End Property



Public Property Get Appearance() As eAppearance
    Appearance = mvarAppearance
End Property

Public Property Let Appearance(ByVal vData As eAppearance)
    mvarAppearance = vData
    txtText.Appearance = vData
    PropertyChanged "Appearance"
End Property

Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = mvarBorderStyle
End Property

Public Property Let BorderStyle(ByVal vData As eBorderStyle)
    mvarBorderStyle = vData
    txtText.BorderStyle = vData
    PropertyChanged "BorderStyle"
End Property




'***************************************
'***************************************
'***************************************
'UserControl Subs





Private Sub txtText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub UserControl_Initialize()
    '
    txtText.ForeColor = Title.ColorEnabled
    txtText.BackColor = Title.BackColor
    '
End Sub

Private Sub UserControl_InitProperties()
    '
    Enabled = True
    Appearance = 1
    BorderStyle = FixedSingle
    Set Font = Ambient.Font
    Text = Extender.Name
    '
End Sub

Private Sub UserControl_Resize()
    '
    txtText.Height = UserControl.Height
    txtText.Width = UserControl.Width
    '
End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        Case vbKeySpace
            'Call UserControl_MouseDown(1, 1, 1, 1)
    End Select
    '
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        'Case vbKeyUp: SendKeys "+{tab}"
        'Case vbKeyDown: SendKeys "{tab}"
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub



Private Sub UserControl_GotFocus()
    'lblCaption.Font.Underline = True
End Sub

Private Sub UserControl_LostFocus()
    'lblCaption.Font.Underline = False
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    mvarAppearance = PropBag.ReadProperty("Appearance", 1)
    mvarBorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    mvarAlign = PropBag.ReadProperty("Align", 0)
    mvarMaxLenght = PropBag.ReadProperty("MaxLenght", 0)
    mvarPassWordChar = PropBag.ReadProperty("PasswordChar", "")
    mvarText = PropBag.ReadProperty("Text", Extender.Name)
    Set mvarFont = PropBag.ReadProperty("Font", Ambient.Font)
    mvarEnabled = PropBag.ReadProperty("Enabled", True)
    
    Align = mvarAlign
    MaxLenght = mvarMaxLenght
    PasswordChar = mvarPassWordChar
    Text = mvarText
    Set Font = mvarFont
    Enabled = mvarEnabled
    Appearance = mvarAppearance
    BorderStyle = mvarBorderStyle
    '
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    Call PropBag.WriteProperty("Align", mvarAlign, 0)
    Call PropBag.WriteProperty("MaxLenght", mvarMaxLenght, 0)
    Call PropBag.WriteProperty("PasswordChar", mvarPassWordChar, "")
    Call PropBag.WriteProperty("Text", mvarText, Extender.Name)
    Call PropBag.WriteProperty("Font", mvarFont, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", mvarEnabled, True)
    Call PropBag.WriteProperty("Appearance", mvarAppearance, 1)
    Call PropBag.WriteProperty("BorderStyle", mvarBorderStyle, 1)
    '
End Sub


