Attribute VB_Name = "modUserCtls"
Option Explicit

Public Enum eAlign
    Left = 0
    Right = 1
    Center = 2
End Enum

Public Enum eAppearance
    Flat = 0
    ThreeD = 1
End Enum

Public Enum eBorderStyle
    None = 0
    FixedSingle = 1
End Enum

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

