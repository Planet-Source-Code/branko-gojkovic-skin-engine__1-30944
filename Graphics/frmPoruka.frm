VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPoruka 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
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
   ScaleHeight     =   2685
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   225
      Top             =   1575
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   375
      Top             =   3075
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPoruka.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPoruka.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPoruka.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPoruka.frx":0CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Graphics.bgDugme cmd 
      Height          =   465
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   3900
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "&OK"
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
   Begin Graphics.bgDugme cmd 
      Height          =   465
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   4425
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "&Yes"
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
   Begin Graphics.bgDugme cmd 
      Height          =   465
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Top             =   4950
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   820
      Caption         =   "&No"
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
   Begin VB.Label pnlTitle 
      BackStyle       =   0  'Transparent
      Height          =   1440
      Left            =   1200
      TabIndex        =   3
      Top             =   450
      Width           =   4665
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   375
      Top             =   600
      Width           =   465
   End
End
Attribute VB_Name = "frmPoruka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Aktiv As Boolean

Private WithEvents cmdClose As VBControlExtender
Attribute cmdClose.VB_VarHelpID = -1

Private Sub cmdClose_ObjectEvent(Info As EventInfo)
    If Info = "Klik" Then
        gOdgovor = 2
        Unload Me
    End If
End Sub

Sub Izlaz()
    Unload Me
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0: gOdgovor = 0
        Case 1: gOdgovor = 1
        Case 2: gOdgovor = 2
    End Select
    Unload Me
End Sub

Private Sub Form_Activate()
    '
    If Aktiv = False Then
        Aktiv = True
        LoadSkin Me, Aktivan, True, 1
        Play PopUp
    Else
        LoadSkin Me, Aktivan, False
    End If
End Sub

Private Sub Form_Deactivate()
    LoadSkin Me, Neaktivan, False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        '
        Case vbKeyEscape
            gOdgovor = Ne
            Unload Me
        '
    End Select
    '
End Sub

Private Sub Form_Load()
    Aktiv = False
    Set cmdClose = Me.Controls.Add("graphics.bgclose", "cmdClose")
    cmdClose.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    Aktiv = False
    '
End Sub

Private Sub Timer1_Timer()
    If Me!lbltitle.Visible = True Then
        Me!lbltitle.Visible = False
    Else
        Me!lbltitle.Visible = True
    End If
End Sub
