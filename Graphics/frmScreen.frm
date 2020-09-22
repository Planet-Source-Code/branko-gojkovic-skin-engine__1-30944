VERSION 5.00
Begin VB.Form frmScreen 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Graphics.bgClose bgClose1 
      Height          =   315
      Left            =   -2000
      TabIndex        =   1
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   150
      ScaleHeight     =   990
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   50000
      Width           =   1815
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

