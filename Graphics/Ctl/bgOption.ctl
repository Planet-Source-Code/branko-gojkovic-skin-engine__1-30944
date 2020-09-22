VERSION 5.00
Begin VB.UserControl bgOption 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   DataBindingBehavior=   1  'vbSimpleBound
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "bgOption.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   360
   ScaleWidth      =   1560
   ToolboxBitmap   =   "bgOption.ctx":0152
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   450
      X2              =   1425
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Option"
      Height          =   240
      Left            =   525
      TabIndex        =   0
      Top             =   75
      Width           =   795
   End
End
Attribute VB_Name = "bgoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarCaption As String
Private mvarEnabled As Boolean
Private mvarFont As Font
Private vy As Integer
Private mStanje As Boolean
Private mvarValue As Boolean 'local copy


Sub ResizeC()
    Line1.Y1 = (lblCaption.Top + lblCaption.Height) + 10
    Line1.Y2 = (lblCaption.Top + lblCaption.Height) + 10
    Line1.X1 = lblCaption.Left
    Line1.X2 = lblCaption.Left + lblCaption.Width
End Sub

Private Sub Iscrtaj(ByVal myStanje As Boolean)
    '
    If myStanje = False Then
        StretchBlt UserControl.hdc, 0, 0, UserControl.Width, UserControl.Height, SourceHdc, mBack.X, mBack.Y, mBack.Width, mBack.Height, vbSrcCopy
        BitBlt UserControl.hdc, 0, vy, mOption.Width, mOption.Height, SourceHdc, mOption.X, mOption.Y, vbSrcCopy
    Else
        StretchBlt UserControl.hdc, 0, 0, UserControl.Width, UserControl.Height, SourceHdc, mBack.X, mBack.Y, mBack.Width, mBack.Height, vbSrcCopy
        BitBlt UserControl.hdc, 0, vy, mOption.Width, mOption.Height, SourceHdc, mOption.X + mOption.Width, mOption.Y, vbSrcCopy
    End If
    'picOption.Width = mOption.Width * Screen.TwipsPerPixelX
    'picOption.Height = mOption.Height * Screen.TwipsPerPixelY
    lblCaption.Left = (mOption.Width * Screen.TwipsPerPixelX) + 50
    UserControl.Refresh
    '
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
    'ResizeC
    lblCaption.Refresh
    ResizeC
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
    'Call ResizeC
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


Private Sub UserControl_GotFocus()
    Value = True
    Line1.Visible = True
End Sub

Private Sub UserControl_Initialize()
    '
    Value = False
    Line1.Visible = False
    '
End Sub

Private Sub UserControl_LostFocus()

   Value = False
   Line1.Visible = False
   
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    mvarCaption = PropBag.ReadProperty("Caption", Extender.Name)
    Set mvarFont = PropBag.ReadProperty("Font", Ambient.Font)
    mvarEnabled = PropBag.ReadProperty("Enabled", True)
    mvarValue = PropBag.ReadProperty("Value", False)
    
    Caption = mvarCaption
    Set Font = mvarFont
    Enabled = mvarEnabled
    Value = mvarValue
    '
End Sub

Public Property Let Value(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
    Iscrtaj mvarValue
End Property


Public Property Get Value() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property

Private Sub UserControl_Resize()
    '
    'picOption.Top = (UserControl.Height - picOption.Height) / 2
    lblCaption.Top = ((UserControl.Height - lblCaption.Height) / 2) - 10
    vy = ((UserControl.Height - (mOption.Height * Screen.TwipsPerPixelY)) / 2) / Screen.TwipsPerPixelY
    
    Iscrtaj False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    Call PropBag.WriteProperty("Caption", mvarCaption, Extender.Name)
    Call PropBag.WriteProperty("Font", mvarFont, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", mvarEnabled, True)
    Call PropBag.WriteProperty("Value", mvarValue, False)
    '
End Sub

