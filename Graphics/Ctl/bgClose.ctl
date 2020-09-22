VERSION 5.00
Begin VB.UserControl bgClose 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   MouseIcon       =   "bgClose.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   435
   ScaleWidth      =   480
   ToolboxBitmap   =   "bgClose.ctx":0152
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "bgClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mMouse As Boolean
Private mDown As Boolean
Private Pt As POINTAPI
Private mWnd As Long
Public Event Klik()


Private Sub Timer1_Timer()
    '
    If Ambient.UserMode = False Then Exit Sub
    GetCursorPos Pt
    mWnd = WindowFromPoint(Pt.X, Pt.Y)
    
    If mWnd = UserControl.hWnd Then
        If mDown = True Then Exit Sub
        Iscrtaj 1
        mMouse = True
    Else
        Iscrtaj 0
        mMouse = False
        Timer1.Enabled = False
    End If
    
End Sub

Private Sub UserControl_Initialize()
    '
    Iscrtaj 0
    UserControl.Width = mClose.Width * Screen.TwipsPerPixelX
    UserControl.Height = mClose.Height * Screen.TwipsPerPixelY
    '
End Sub

Private Sub Iscrtaj(mst As Integer)
    '
    Select Case mst
        Case 0 'Normalno
            BitBlt UserControl.hdc, 0, 0, mClose.Width, mClose.Height, SourceHdc, mClose.X, mClose.Y, vbSrcCopy
        Case 1 'Mouse move
            BitBlt UserControl.hdc, 0, 0, mClose.Width, mClose.Height, SourceHdc, mClose.X + mClose.Width, mClose.Y, vbSrcCopy
        Case 2 'Click
            BitBlt UserControl.hdc, 0, 0, mClose.Width, mClose.Height, SourceHdc, mClose.X + (mClose.Width * 2), mClose.Y, vbSrcCopy
    End Select
    UserControl.Refresh
    '
End Sub



Public Sub Dis()
    '
    BitBlt UserControl.hdc, 0, 0, mClose.Width, mClose.Height, SourceHdc, mClose.X, mClose.Y + mClose.Height, vbSrcCopy
    UserControl.Refresh
    '
End Sub

Public Sub Ena()
    '
    BitBlt UserControl.hdc, 0, 0, mClose.Width, mClose.Height, SourceHdc, mClose.X, mClose.Y, vbSrcCopy
    UserControl.Refresh
    '
End Sub

Private Sub UserControl_InitProperties()
    '
    Enabled = True
    '
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Iscrtaj 2
    mDown = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < 0 Or X > ScaleWidth Or Y < 0 Or Y > ScaleHeight Then
    
    Else
        Timer1.Enabled = True
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMouse = True Then
        Iscrtaj 1
        RaiseEvent Klik
    Else
        Iscrtaj 0
    End If
    mDown = False
End Sub
