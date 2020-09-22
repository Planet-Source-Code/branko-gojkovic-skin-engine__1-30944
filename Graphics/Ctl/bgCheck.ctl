VERSION 5.00
Begin VB.UserControl bgCheck 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   DataBindingBehavior=   1  'vbSimpleBound
   MouseIcon       =   "bgCheck.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   255
   ScaleWidth      =   960
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check 1"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "bgCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Local variables
Private mvarCaption As String
Private mvarEnabled As Boolean
Private mvarFont As Font
Private vy As Integer
Private mStanje As Boolean
Private mvarValue As Boolean 'local copy
Private mvarAlign As eAlign

'User Control Events
Public Event KeyPress(KeyAscii As Integer)







'********************************************
'********************************************
'********************************************
'User Control Custom Properties


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
    '
End Property


Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    
    mvarCaption = vData
    lblCaption.Caption = mvarCaption
    lblCaption.ForeColor = Title.ColorEnabled
    lblCaption.Refresh
    
    PropertyChanged ("Caption")
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
        lblCaption.ForeColor = Title.ColorDisabled
        lblCaption.Refresh
    Else
        lblCaption.ForeColor = Title.ColorEnabled
        lblCaption.Refresh
    End If
    '
    PropertyChanged "Enabled"
    UserControl.Enabled = mvarEnabled
    '
End Property


Public Property Let Value(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
    Iscrtaj mvarValue
    PropertyChanged "Value"
End Property


Public Property Get Value() As Boolean
Attribute Value.VB_MemberFlags = "24"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property



Public Property Let Align(ByVal vData As eAlign)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarAlign = vData
    Iscrtaj mvarValue
    PropertyChanged "Align"
End Property


Public Property Get Align() As eAlign
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Align = mvarAlign
End Property




'********************************************
'********************************************
'********************************************
'User control Custom Subs & Functions


Private Sub Iscrtaj(ByVal myStanje As Boolean)
    '
    If mvarAlign = Left Then
        If myStanje = False Then
            BitBlt UserControl.hdc, 0, vy, mCheck.Width, mCheck.Height, SourceHdc, mCheck.X, mCheck.Y, vbSrcCopy
        Else
            BitBlt UserControl.hdc, 0, vy, mCheck.Width, mCheck.Height, SourceHdc, mCheck.X + mCheck.Width, mCheck.Y, vbSrcCopy
        End If
        If mCheck.Width > 0 Then
            lblCaption.Left = (mCheck.Width * Screen.TwipsPerPixelX) + 70
        Else
            lblCaption.Left = 270
        End If
    Else
        If myStanje = False Then
            BitBlt UserControl.hdc, (UserControl.Width / Screen.TwipsPerPixelX) - mCheck.Width, vy, mCheck.Width, mCheck.Height, SourceHdc, mCheck.X, mCheck.Y, vbSrcCopy
        Else
            BitBlt UserControl.hdc, (UserControl.Width / Screen.TwipsPerPixelX) - mCheck.Width, vy, mCheck.Width, mCheck.Height, SourceHdc, mCheck.X + mCheck.Width, mCheck.Y, vbSrcCopy
        End If
        
        If mCheck.Width > 0 Then
            lblCaption.Left = (UserControl.Width - (mCheck.Width * Screen.TwipsPerPixelX)) - (lblCaption.Width + 70)
        Else
            lblCaption.Left = (UserControl.Width - lblCaption.Width) - 70
        End If
    End If
        
    UserControl.Refresh
    '
End Sub













'***************************************
'***************************************
'***************************************
'UserControl Subs

Private Sub lblCaption_Click()
    '
    Call UserControl_MouseDown(1, 1, 1, 1)
    '
End Sub

Private Sub UserControl_InitProperties()
    '
    Enabled = True
    Set Font = Ambient.Font
    Caption = Extender.Name
    '
End Sub

Private Sub UserControl_Resize()
    '
    Value = False
    lblCaption.Font.Underline = False
    lblCaption.Top = ((UserControl.Height - lblCaption.Height) / 2) - 10
    vy = ((UserControl.Height - (mOption.Height * Screen.TwipsPerPixelY)) / 2) / Screen.TwipsPerPixelY
    StretchBlt UserControl.hdc, 0, 0, UserControl.Width, UserControl.Height, SourceHdc, mBack.X, mBack.Y, mBack.Width, mBack.Height, vbSrcCopy
    Iscrtaj False
End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        Case vbKeySpace
            Call UserControl_MouseDown(1, 1, 1, 1)
    End Select
    '
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp: SendKeys "+{tab}"
        Case vbKeyDown: SendKeys "{tab}"
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub



Private Sub UserControl_GotFocus()
    lblCaption.Font.Underline = True
End Sub


Private Sub UserControl_LostFocus()
    lblCaption.Font.Underline = False
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If Value = True Then
        Value = False
    Else
        Value = True
    End If
End Sub




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    mvarAlign = PropBag.ReadProperty("Align", 0)
    mvarCaption = PropBag.ReadProperty("Caption", Extender.Name)
    Set mvarFont = PropBag.ReadProperty("Font", Ambient.Font)
    mvarEnabled = PropBag.ReadProperty("Enabled", True)
    mvarValue = PropBag.ReadProperty("Value", False)
    
    Align = mvarAlign
    Caption = mvarCaption
    Set Font = mvarFont
    Enabled = mvarEnabled
    Value = mvarValue
    
    '
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    Call PropBag.WriteProperty("Align", mvarAlign, 0)
    Call PropBag.WriteProperty("Caption", mvarCaption, Extender.Name)
    Call PropBag.WriteProperty("Font", mvarFont, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", mvarEnabled, True)
    Call PropBag.WriteProperty("Value", mvarValue, False)
    '
End Sub

