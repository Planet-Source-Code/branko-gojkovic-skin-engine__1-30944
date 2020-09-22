VERSION 5.00
Begin VB.Form frmLogIn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Identifikacija korisnika"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Graphics.bgDugme cmd 
      Height          =   465
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   2100
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "&Acept"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Graphics.bgFrame bgFrame1 
      Height          =   1140
      Left            =   225
      Top             =   600
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   2011
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Graphics.bgText bgText2 
         Height          =   390
         Left            =   1500
         TabIndex        =   3
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         MaxLenght       =   15
         PasswordChar    =   "*"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Graphics.bgText bgText1 
         Height          =   390
         Left            =   1500
         TabIndex        =   2
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         MaxLenght       =   15
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   1
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User name:"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
   End
   Begin Graphics.bgDugme cmd 
      Height          =   465
      Index           =   1
      Left            =   2625
      TabIndex        =   5
      Top             =   2100
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "&Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Aktiv As Boolean
Private WithEvents cmdClose As VBControlExtender
Attribute cmdClose.VB_VarHelpID = -1



Private Sub cmdClose_ObjectEvent(Info As EventInfo)
    If Info = "Klik" Then Unload Me
End Sub

Private Sub Form_Load()
    Aktiv = False
    Set cmdClose = Me.Controls.Add("graphics.bgclose", "cmdClose")
    cmdClose.Visible = True
End Sub

Private Sub Form_Activate()
    '
    If Aktiv = False Then
        Aktiv = True
        LoadSkin Me, Aktivan, True, 1
    Else
        LoadSkin Me, Aktivan, False
    End If
End Sub

Private Sub Form_Deactivate()
    '
    LoadSkin Me, Neaktivan, False
    '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
    '
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If Y <= gV And X < gŠ Then
            Screen.MouseIcon = LoadResPicture(101, vbResCursor)
            Screen.MousePointer = 99
            Pomeri Me
        Else
            
        End If
    End If
    If Screen.MousePointer <> 0 Then Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Aktiv = False
End Sub
