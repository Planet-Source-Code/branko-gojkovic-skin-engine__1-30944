VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "BG Soft graphics engine X"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Graphics.bgFrame bgFrame1 
      Height          =   1815
      Left            =   450
      Top             =   2775
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3201
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
      Begin Graphics.bgCheck chk 
         Height          =   240
         Index           =   0
         Left            =   975
         TabIndex        =   3
         Top             =   225
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Check1"
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
      Begin Graphics.bgCheck chk 
         Height          =   240
         Index           =   1
         Left            =   975
         TabIndex        =   4
         Top             =   675
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Check2"
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
   Begin Graphics.bgDugme cmd 
      Height          =   540
      Index           =   0
      Left            =   4575
      TabIndex        =   0
      ToolTipText     =   "Konfiguracija programa"
      Top             =   600
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   953
      Caption         =   "&Set skin"
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
      Height          =   540
      Index           =   1
      Left            =   4575
      TabIndex        =   1
      ToolTipText     =   "Informacije o programu"
      Top             =   1200
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   953
      Caption         =   "About"
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
      Height          =   540
      Index           =   2
      Left            =   4575
      TabIndex        =   2
      ToolTipText     =   "Izlaz iz programa"
      Top             =   1800
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   953
      Caption         =   "Exit"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Aktiv As Boolean

Private WithEvents cmdClose As VBControlExtender
Attribute cmdClose.VB_VarHelpID = -1

Private Sub bgDugme1_Click()
    frmLogIn.Show 1
End Sub

Private Sub cmdClose_ObjectEvent(Info As EventInfo)
    If Info = "Klik" Then Unload Me
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0: frmOptions.Show 1
        Case 1: frmAbout.Show 1
        Case 2: Unload Me
    End Select
End Sub



Private Sub Form_Activate()
    If Aktiv = False Then
        Aktiv = True
        LoadSkin Me, Aktivan, True, 1
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
        'Case vbKeyEscape: Unload Me
    End Select
    '
End Sub

Private Sub Form_Load()
    '
    Set cmdClose = Me.Controls.Add("graphics.bgclose", "cmdClose")
    cmdClose.Visible = True
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
    Dim i As Integer
    
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    End
End Sub

