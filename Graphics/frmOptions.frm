VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Konfiguracija programa"
   ClientHeight    =   8025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
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
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1021"
   Begin Graphics.bgDugme cmd 
      Height          =   465
      Index           =   0
      Left            =   6975
      TabIndex        =   12
      ToolTipText     =   "Izlaz na glavni meni"
      Top             =   7350
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      Caption         =   "Izlaz"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   300
      TabIndex        =   6
      Top             =   600
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Grafika"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblAktivni"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmd(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Picture1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "picDemo"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "List1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.ListBox List1 
         BackColor       =   &H8000000F&
         Height          =   4860
         Left            =   300
         TabIndex        =   11
         Top             =   975
         Width           =   1815
      End
      Begin VB.PictureBox picDemo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   2250
         ScaleHeight     =   4830
         ScaleWidth      =   5655
         TabIndex        =   9
         Top             =   975
         Width           =   5715
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   13
         Top             =   0
         Width           =   15
      End
      Begin Graphics.bgDugme cmd 
         Height          =   465
         Index           =   1
         Left            =   300
         TabIndex        =   14
         ToolTipText     =   "Postavljanje selektovanog skina za aktivni"
         Top             =   6000
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   820
         Caption         =   "Set as active skin"
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
      Begin VB.Label lblAktivni 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1425
         TabIndex        =   10
         Top             =   675
         Width           =   3390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Aktivni skin:"
         Height          =   240
         Left            =   300
         TabIndex        =   8
         Top             =   675
         Width           =   1140
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2022
         Left            =   505
         TabIndex        =   5
         Tag             =   "1025"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2022
         Left            =   406
         TabIndex        =   4
         Tag             =   "1024"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2022
         Left            =   307
         TabIndex        =   2
         Tag             =   "1023"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   1125
      TabIndex        =   7
      Top             =   3300
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
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

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
        Case 0: Unload Me
        Case 1: PostaviSkin
    End Select
    
End Sub

Private Sub PostaviSkin()
    '
    SetSkin List1.Text
    lblAktivni.Caption = List1.Text
    If msg("To teke efects you mast restart program!" & vbCr & vbCr & "Do you want restart program now?", Pitanje, DaNe, "Skin is changed") = Da Then End
    '
End Sub

Private Sub cmdClose_Klik()
    Unload Me
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
    End Select
    '
End Sub

Private Sub Form_Load()
    '
    Dim i As Integer
    Dim BrojK As Integer
    Dim LI As Integer
    
    Set cmdClose = Me.Controls.Add("graphics.bgclose", "cmdClose")
    cmdClose.Visible = True
    
    File1.Path = App.Path & "\Skins\"
    File1.Pattern = "*.bmp"
    For i = 0 To File1.ListCount - 1
        BrojK = Len(File1.List(i)) - 4
        List1.AddItem Mid$(File1.List(i), 1, BrojK)
    Next i
    GetSkin
    For i = 0 To List1.ListCount - 1
        If mSkin = List1.List(i) Then
            LI = i
            Exit For
        End If
    Next i
    List1.Text = List1.List(LI)
    lblAktivni.Caption = mSkin
    cmd(1).Caption = "Aktivni skin"
    cmd(1).Enabled = False
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Aktiv = False
End Sub

Private Sub List1_Click()
    '
    On Error Resume Next
    picDemo.Picture = LoadPicture(App.Path & "\Skins\" & List1.Text & ".demo")
    If List1.Text <> lblAktivni.Caption Then
        cmd(1).Caption = "Postavi '" & List1.Text & "' za aktivni skin"
        cmd(1).Enabled = True
    Else
        cmd(1).Caption = "Aktivni skin"
        cmd(1).Enabled = False
    End If
    '
End Sub
