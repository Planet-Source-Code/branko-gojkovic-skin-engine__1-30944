Attribute VB_Name = "modSound"
Option Explicit

Private Const SND_ALIAS = &H10000
Public Const SND_ASYNC = &H1
Private Const SND_FILENAME = &H20000
Private Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Private Const SND_NOSTOP = &H10
Private Const SND_NOWAIT = &H2000
Private Const SND_SYNC = &H0
Public Const SND_MEMORY = &H4
Public gSound As Boolean

Private Declare Function PlaySound Lib "WINMM.DLL" Alias _
   "PlaySoundA" (ByVal lpszName As String, _
   ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function mciSendString Lib "WINMM.DLL" Alias _
   "mciSendStringA" (ByVal lpstrCommand As String, _
   ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
   ByVal hwndCallback As Long) As Long
   
Private Const MMSYSERR_NOERROR = 0

Public Const AUDIO_NONE = 0
Public Const AUDIO_WAVE = 1
Public Const AUDIO_MIDI = 2
Public SoundBuffer() As Byte

Public Enum eIzbor
    Klik = 0
    PopUp = 1
End Enum

    
Public Function CanPlaySound() As Integer
    ' Returns 1 if wave output
    ' Returns 2 if midi output
    ' Returns 3 if both
    '
    Dim i As Integer
    '
    i = AUDIO_NONE
    '
    If waveOutGetNumDevs > 0 Then
        i = AUDIO_WAVE
    End If
    '
    If midiOutGetNumDevs > 0 Then
        i = i + AUDIO_MIDI
    End If
    '
    CanPlaySound = i
    '
End Function

Public Function Play(ByVal izbor As eIzbor, Optional async As Variant, Optional sLoop As Variant) As Boolean

    Dim i As Integer
    Dim F As String
    Dim j As Long
    Dim Filename As String
    If gSound = False Then Exit Function
    '
    Select Case izbor
        Case Klik: Filename = App.Path & "\Skins\" & "Klik.Wav"
        Case PopUp: Filename = App.Path & "\Skins\" & "PopUp.Wav"
    End Select
    '
    i = Len(Filename)
    F = UCase(Filename)

    If IsMissing(async) Then
        j = SND_ASYNC
    Else
        If async Then
            j = SND_ASYNC
        Else
            j = SND_SYNC
        End If
    End If

    If Not IsMissing(sLoop) Then
        If sLoop And (j = SND_ASYNC) Then
        j = j + SND_LOOP
        End If
    End If

    j = j + SND_NOSTOP + SND_NOWAIT

    If InStr(F, ".WAV") = i - 3 Then
      If AUDIO_WAVE = AUDIO_WAVE Then
         j = j + SND_FILENAME + SND_NODEFAULT
         i = PlaySound(Filename, 0, j)
         Play = IIf(i = 0, False, True)
      Else
         Beep
         Play = True
      End If
    End If
End Function
