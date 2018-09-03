Attribute VB_Name = "auxrout2"
Option Explicit


' Names of shell windows sought
Const g_cstrShellViewWnd As String = "Progman"
Const g_cstrShellTaskBarWnd As String = "Shell_TrayWnd"

' For ShowWindow()
Const SW_HIDE = 0
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10


' Used by GetWindowWord to find next window
Const GW_HWNDNEXT = 2


'Taskbar
Private TaskbarVisible As Boolean

' Sound

Const SND_SYNC = &H0            ' play synchronously
Const SND_ASYNC = &H1           ' play asynchronously
Const SND_ALIAS = &H10000       ' name is a WIN.INI [sounds] entry
Const SND_NODEFAULT = &H2       ' No default message if error
Const SND_FILENAME = &H20000    ' name is a file name
Const SND_PURGE = &H40          ' Stop Sound
Const MMSYSERR_BASE = 0
Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)           ' Unspecified
Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)     ' device ID out of range
Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)       ' Specified resource already allocated.
Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)
Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)        ' no device driver present
Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)
Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)

'Midi sound
Const MAXPNAMELEN = 32 ' Max product name length
Const MOD_MAPPER = 5   ' Midi Mapper
Const MM = "MIDIMapper.Configurator.exe"
Const VM = "VirtualMIDISynth.dll"
Const CALLBACK_NULL = 0

' For bitmapdb
Const Msgdup = "There is a duplicate of this quote already in database. Quotes just appended are located under ""New"". Please change the text"

' System sound definitions

Public Type SystemSoundDefinitions
  GroupName As String
  SoundName As String
  RegKey As String
  Current As String
  Default As String
End Type

'User-defined variable storing information about the MIDI output device.
Type MIDIOUTCAPS
 wMid As Integer                   ' Mfg id of device driver for MIDI output dev.

 wPid As Integer                   ' Product Id of MIDI output dev.

 vDriverVersion As Long            ' Version no. of device driver for MIDI output device. High-order byte: major version no. & low-order byte: minor version no.

 szPname As String * MAXPNAMELEN   ' Product name in null-terminated string.

 wTechnology As Integer
 ' One of following describing MIDI output device:
 '  MOD_FMSYNTH-Device is an FM synthesizer.
 '  MOD_MAPPER-Device is the Microsoft MIDI mapper.
 '  MOD_MIDIPORT-Device is a MIDI hardware port.
 '  MOD_SQSYNTH-Device is a square wave synthesizer.
 '  MOD_SYNTH-Device is a synthesizer.

  wVoices As Integer          ' No. of voices supported by an internal synthesizer device. If device is port, this member is not meaningful and set to 0.

  wNotes As Integer           ' Maximum number of simultaneous notes that can be played by an internal synthesizer device. If device is port, this member is not meaningful and set to 0.

  wChannelMask As Integer     ' Channels that an internal synthesizer device responds to, where the least significant bit refers
                              ' to channel 0 and most significant bit to channel 15. Port devices that transmit on all channels set this member to 0xFFFF.

  dwSupport As Long
  ' One of the following describes optional functionality supported by device:
  ' MIDICAPS_CACHE-Supports patch caching.
  ' MIDICAPS_LRVOLUME-Supports separate left & right volume control.
  ' MIDICAPS_STREAM-Provides direct support for midiStreamOut function.
  ' MIDICAPS_VOLUME-Supports volume control.
  ' If a device supports volume changes, the MIDICAPS_VOLUME flag will be set for dwSupport member. If device supports separate
  ' volume changes on the left & right channels, both MIDICAPS_VOLUME & MIDICAPS_LRVOLUME flags will be set for this member.
End Type


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
' Retrieves the number of MIDI output devices present in the system. The function returns
' the number of MIDI output devices. A zero return value means no MIDI devices in the system.
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
' Queries a specified MIDI output device to determine its capabilities. Requires the following parms;
'   uDeviceID- Uint variable identifying of the MIDI output device. The device id specified by this parm
'   varies from zero to one less than the no. of devices present. It can also be a properly cast device handle.
'   lpMidiOutCaps- Address of MIDIOUTCAPS structure, which is filled with info about capabilities of device.
'   cbMidiOutCaps- Size, in bytes, of the MIDIOUTCAPS structure. Use Len function with MIDIOUTCAPS var as argument to get this value.
'
'
' Sound APIs
'
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
' Function closes the specified MIDI output device. Requires a handle to MIDI output device. If function is successful, handle is no longer valid after call to this function. A successful function call returns MMSYSERR_NOERROR or 0

Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
'                      Function opens a MIDI output device for playback, requiring the following parms
'  lphmo-              Address of HMIDIOUT handle. This location is filled with a handle identifying the opened
'                      MIDI output device. The handle is used to identify device in calls to other MIDI output functions.
'  uDeviceID-          Identifier of the MIDI output device that is to be opened.
'  dwCallback-         Address of a callback function, event handle, a thread identifier, or handle of a window
'                      or thread called during MIDI playback to process messages related to the progress of
'                      the playback. If no callback is desired, set this value to 0.
'  dwCallbackInstance- User instance data passed to the callback. Set this value to 0.
'  dwFlags-            Callback flag for opening the device. Set this value to 0.
'

Private Declare Function midiStreamOpen Lib "winmm.dll" (ByVal hms As LONG_PTR, ByVal puDeviceID As LONG_PTR, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function midiStreamStop Lib "winmm.dll" (ByVal hms As Long) As Long
Private Declare Function midiStreamClose Lib "winmm.dll" (ByVal hms As Long) As Long

' Required API declarations for Find Window
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long



' Show Window
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd As Long, ByVal hWndChild As Long, ByVal lpszClassName As String, ByVal lpszWindow As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private workfle(8) As String
Private response As Long, ct As Long, ct1 As Long, hWnd As Long, nRet As Long

' Midi vars
Private midirc As Long, oldsndfname As String, CanPlayWaves As Boolean, hmidi(5) As Long, hmsmidi(5) As Long, curmididev As Long, strPlaying As String, VMSynth As Boolean

' Picnames
Private picnames(450) As String, picnamesort(450) As String
Dim c As New cRegistry
Private Function FindWindowPartial(Optional InstallUpdate As Boolean = False) As Long
Dim TitleTmp As String

ct = 0

  ' Find first window and loop through all subsequent windows in master window list.


  hWnd = FindWindow(vbNullString, vbNullString)
  Do Until hWnd = 0

    ' Make sure this window has no parent.

    If GetParent(hWnd) = 0 Then

      ' Retrieve caption text from current window.

      TitleTmp = Space(256)
      nRet = GetWindowText(hWnd, TitleTmp, Len(TitleTmp))
      If nRet Then

        ' Clean up return string, prepare for case-insensitive comparison.

        TitleTmp = UCase(Left(TitleTmp, nRet))
        If InstallUpdate = True Then
          'Don't worry about partials
                  If InStr(TitleTmp, UCase("SetupMyReels")) Or InStr(TitleTmp, UCase("UpdateMyReels")) Then
                  FindWindowPartial = 1
                  Exit Function
                  End If

        ' Use appropriate method to determine if current window's caption either starts with or contains passed string.

        ElseIf InStr(TitleTmp, UCase("MyReels")) Then
          If InStr(TitleTmp, UCase("microsoft visual basic")) > 0 Then
          ct = ct - 1
          ElseIf Not (InStr(TitleTmp, UCase("SetupMyReels")) > 0 Or InStr(TitleTmp, UCase("UpdateMyReels")) > 0 Or InStr(TitleTmp, UCase("WinDbg")) > 0 Or InStr(TitleTmp, UCase("explor")) > 0 Or InStr(TitleTmp, UCase("\")) > 0 Or InStr(TitleTmp, UCase("/")) > 0) Then
          ct = ct + 2
          End If
        FindWindowPartial = ct
        End If
      End If
    End If

    ' Get next window in master window list & continue.
    hWnd = GetWindow(hWnd, GW_HWNDNEXT)
  Loop
End Function
Public Sub setformpos(objform As Object, Optional ByVal noVert As Boolean = False)
c.hDC = objform.hDC

With objform
.Left = (c.DCWidthTwips - .Width) / 2

If (c.DCHeightTwips - .Height > 225 * resY) Then
 If noVert = False Then .Top = (c.DCHeightTwips - .Height) / 2
ct = .Top
End If

End With

End Sub
Public Sub PokeResolution(objform As Object)
Dim intreel As Long, resTmp As Integer


c.hWnd = objform.hWnd
c.hDC = objform.hDC


' Is taskbar showing?
TaskbarVisible = FindShellTaskBar()

With objform



If (c.VistaorLater() = True) Then
resX = c.WWidth * Screen.TwipsPerPixelX / .Width
resY = c.WHeight * Screen.TwipsPerPixelY / .Height
Else
Select Case c.DCWidth
Case Is < 1023

resX = 1

Case 1023 To 1151

resX = 32 / 25

Case 1152 To 1279

resX = 35 / 25

Case 1279 To 1281

resX = 38 / 25

Case Else
resX = 40 / 25
End Select
resY = resX
End If


' fix TaskbarVisible logic
If TaskbarVisible Then
  If (c.IsWin2000NT = True) Or (c.XP = False) Or (c.VistaorLater() = False) Then
 .Width = resX * .Width
 .Height = resY * .Height
  Dotaskwindow objform, True
  Else
   If (c.Width - c.WWidth) < (c.Height - c.WHeight) Then

     If c.WTop = c.Top Then
     .Top = c.Top * Screen.TwipsPerPixelY
     Else
     .Top = c.WTop * Screen.TwipsPerPixelY
     End If
    .Left = c.Left * Screen.TwipsPerPixelX

   Else ' Vertical Taskbar


     If c.WLeft = c.Left Then
     .Left = c.Left * Screen.TwipsPerPixelX
     Else
     .Left = c.WLeft * Screen.TwipsPerPixelX
     End If
   .Top = c.Top * Screen.TwipsPerPixelY

   End If
   .Height = c.WHeight * Screen.TwipsPerPixelY
   .Width = c.WWidth * Screen.TwipsPerPixelX
  End If
  Else
 .Width = resX * .Width
 .Height = resY * .Height
End If


.lblmisc(1).Move resX * .lblmisc(1).Left, resY * .lblmisc(1).Top
If (c.XP Or c.VistaorLater) Then
resTmp = 0
Else
resTmp = 30
End If

.lblprizemeter(1).Move resX * (resTmp + .lblprizemeter(1).Left - 200 * gt(159)), resY * .lblprizemeter(1).Top
.lblmisc(8).Left = resX * (.lblmisc(8).Left - 300 * gt(159))
.lblmisc(9).Left = resX * (.lblmisc(9).Left - 300 * gt(159))

If gt(200) = 0 And (Stringvars(3) = "" Or gt(191) = 0) Then
.Quotez.Visible = False
Else
.Quotez.Move resX * .Quotez.Left, resY * .Quotez.Top, resX * .Quotez.Width, resY * .Quotez.Height
End If

For ct = 0 To 1
  If resX > 35 / 25 Then
  .Candy(ct).Width = (35 / 25) * .Candy(ct).Width
  Else
  .Candy(ct).Width = resX * .Candy(ct).Width
  End If
.Candy(ct).Height = resY * .Candy(ct).Height

If resX > 1 Then .lblrandom(ct).Move resX * .lblrandom(ct).Left, resY * .lblrandom(ct).Top + 90
Next

 
 If gt(159) = 1 Then

.frapicarea.Move resX * 3360, resY * 2525, resX * 7640, resY * 4600

.Quotez.Visible = False
.Lnptr(0).Move 3210, 4750
.Lnptr(1).Move 11040, 4750
.Lnptr(2).Move 3210, 3150
.Lnptr(3).Move 11040, 3150
.Lnptr(4).Move 3210, 6150
.Lnptr(5).Move 11040, 6150

' MB
.lblmisc(1).Caption = " Money Back POT ....                  Money Back Status   ==>"
.lblmisc(1).Width = resX * 7995
.lblmisc(7).Move resX * 3480, resY * 840

.lblprizemeter(1).Top = resY * 7920 'old 7560
.lblmisc(8).Top = resY * 7580 ' old 7200
.lblmisc(9).Top = resY * 8260 ' old 8040

.lblmisc(6).Move resX * 4900, resY * 8160
.lblprizemeter(0).Move resX * 5300, resY * 8160
.lblmisc(0).Move resX * 10740, resY * 1170
.lblmisc(4).Move resX * 6000, resY * 7760


For ct = 0 To 1
.Candy(ct).Left = .frapicarea.Left + (1 + 2 * ct) / 4 * .frapicarea.Width - .Candy(ct).Width / 2
.Candy(ct).Top = resY * 7080
Next

If resX > 1 Then

For ct = 0 To 5
Select Case ct
Case 0, 2, 4
.Lnptr(ct).Move (30 + resX * .Lnptr(ct).Left), resY * .Lnptr(ct).Top
Case Else
.Lnptr(ct).Move resX * .Lnptr(ct).Left, resY * .Lnptr(ct).Top
End Select
Next

End If

Else    ' gt(159)=0

If (c.VistaorLater() = False) Then
Select Case resX
Case Is <= 32 / 25
resTmp = 0
Case 35 / 25
resTmp = 10
Case Else
resTmp = 20
End Select
Else
resTmp = 0
End If

.frapicarea.Move resX * .frapicarea.Left - resTmp, resY * .frapicarea.Top - 10, resX * .frapicarea.Width, resY * .frapicarea.Height
.lblmisc(5).Move resX * .lblmisc(5).Left, resY * .lblmisc(5).Top

For ct = 0 To 13
For intreel = 0 To 4
.lblprize(14 * intreel + ct).Left = resX * .lblprize(14 * intreel + ct).Left
.lblprize(14 * intreel + ct).Top = resY * .lblprize(14 * intreel + ct).Top
.lblprizeamt(14 * intreel + ct).Move resX * .lblprizeamt(14 * intreel + ct).Left, resY * .lblprizeamt(14 * intreel + ct).Top
Next
Next

If resX > 1 Then

For ct = 0 To 5
Select Case ct
Case 0, 2, 4
.Lnptr(ct).Move .frapicarea.Left - .Lnptr(ct).Width, resY * .Lnptr(ct).Top
Case Else
.Lnptr(ct).Move .frapicarea.Left + .frapicarea.Width, resY * .Lnptr(ct).Top
End Select
Next


.lblmisc(1).Width = resX * .lblmisc(1).Width

.lblmisc(7).Move resX * .lblmisc(7).Left, resY * .lblmisc(7).Top
.lblmisc(8).Top = resY * .lblmisc(8).Top
.lblmisc(9).Top = resY * .lblmisc(9).Top

.lblprizemeter(0).Move resX * .lblprizemeter(0).Left, resY * .lblprizemeter(0).Top
.lblmisc(6).Move resX * .lblmisc(6).Left, resY * .lblmisc(6).Top
.lblmisc(0).Move resX * .lblmisc(0).Left, resY * .lblmisc(0).Top
.lblmisc(4).Move resX * .lblmisc(4).Left, resY * .lblmisc(4).Top

For ct = 0 To 1
' dup statement for gt(159) = 1 but frapicarea changed!
.Candy(ct).Left = .frapicarea.Left + (1 + 2 * ct) / 4 * .frapicarea.Width - .Candy(ct).Width / 2
.Candy(ct).Top = resY * .Candy(ct).Top
Next

End If

End If
End With

setformpos objform, c.VistaorLater

End Sub
Public Function fixpw(intct As Long, pw As Long) As Long
Dim res As Single
If (c.VistaorLater() = True) Then
res = resY
Else
res = resX
End If

If (Pokemach.frapicarea.Height / 3) > fixpw Then
' NVIDIA
fixpw = (Pokemach.frapicarea.Height / 3) - (7 * res)
Else
Select Case res
Case Is < 32 / 25
fixpw = pw
Case Is >= 32 / 25, Is < 35 / 25
If gt(159) = 0 Then
fixpw = 23 * pw / 18
Else
fixpw = 99 * pw / 78
End If
Case Is >= 35 / 25, Is < 38 / 25
If gt(159) = 0 Then
fixpw = 25 * pw / 18
Else
fixpw = 111 * pw / 78 '12
End If
Case Is >= 38 / 25, Is < 40 / 25
If gt(159) = 0 Then
fixpw = 28 * pw / 18
Else
fixpw = 21 * pw / 13 '15
End If
Case Is >= 40 / 25
If gt(159) = 0 Then
fixpw = 31 * pw / 18
Else
fixpw = 144 * pw / 78 '18
End If
End Select
End If

fixpw = intct * fixpw
End Function
Private Sub Main()
Dim dt As Date

VOGchg = -1
gt(0) = -4
LoadFrmSplsh 3000
genoptsgen = False  ' can't obviously do this in PokeLoad
isMainActive = True


' Check for previous instance app.previnstance not working
If c.MutexChk() Then
  MsgBox "Sorry, only 1 copy of MyReels allowed"
  Unload frmSplsh
  c.CloseMutexhandle
  Exit Sub
Else
  If c.mch2000nt = True And FindWindowPartial > 2 Then
  response = MsgBox("Only 1 copy of MyReels should run." & vbNewLine & "If there is an explorer folder or a task containing the words ""MyReels"" currently open," & vbNewLine & "click ""OK"", close it and retry the game from a shortcut." & vbNewLine & "Else click ""Cancel"" for possible errors.", vbOKCancel)
  If response = vbOKCancel Then
  Unload frmSplsh
  c.CloseMutexhandle
  Exit Sub
  End If
  End If
End If

' Init quotes db pics
For ct = 1 To 450
Select Case ct
Case 1 To 28, 30 To 40, 107, 109 To 157, 170, 184, 185
picnames(ct) = "lear"
Case 29
picnames(ct) = "anger"
Case 41
picnames(ct) = "cocktail"
Case 42
picnames(ct) = "collins"
Case 43
picnames(ct) = "old-fashioned"
Case 44
picnames(ct) = "highball"
Case 45
picnames(ct) = "margarita"
Case 46
picnames(ct) = "sherry"
Case 47
picnames(ct) = "sour"
Case 48
picnames(ct) = "punch cup"
Case 49
picnames(ct) = "irish coffee"
Case 50
picnames(ct) = "parfait"
Case 51
picnames(ct) = "champagne"
Case 52
picnames(ct) = "beer mug"
Case 53
picnames(ct) = "brandy snifter"
Case 54
picnames(ct) = "baseball"
Case 55
picnames(ct) = "laugh"
Case 56
picnames(ct) = "ability"
Case 57
picnames(ct) = "dead"
Case 58
picnames(ct) = "saw"
Case 59
picnames(ct) = "pageturn"
Case 60
picnames(ct) = "diff"
Case 61
picnames(ct) = "sad"
Case 62
picnames(ct) = "boy"
Case 63
picnames(ct) = "paradise"
Case 64
picnames(ct) = "travel"
Case 65
picnames(ct) = "write"
Case 66
picnames(ct) = "clown"
Case 67
picnames(ct) = "god"
Case 68
picnames(ct) = "world"
Case 69
picnames(ct) = "light"
Case 70
picnames(ct) = "home"
Case 71
picnames(ct) = "rose"
Case 72
picnames(ct) = "attend"
Case 73
picnames(ct) = "actor"
Case 74
picnames(ct) = "mouse"
Case 75
picnames(ct) = "coins"
Case 76
picnames(ct) = "gotit"
Case 77
picnames(ct) = "not"
Case 78
picnames(ct) = "rabbit"
Case 79
picnames(ct) = "sweet"
Case 80
picnames(ct) = "computer"
Case 81
picnames(ct) = "justice"
Case 82
picnames(ct) = "music"
Case 83
picnames(ct) = "moon"
Case 84
picnames(ct) = "me"
Case 85
picnames(ct) = "baby"
Case 86
picnames(ct) = "babe"
Case 87
picnames(ct) = "heart"
Case 88
picnames(ct) = "tree"
Case 89
picnames(ct) = "ear"
Case 90
picnames(ct) = "fruit"
Case 91
picnames(ct) = "clipbrd"
Case 92
picnames(ct) = "yawn"
Case 93
picnames(ct) = "bird"
Case 94
picnames(ct) = "mug"
Case 95
picnames(ct) = "bee"
Case 96
picnames(ct) = "time"
Case 97
picnames(ct) = "car"
Case 98
picnames(ct) = "plane"
Case 99
picnames(ct) = "phone"
Case 100
picnames(ct) = "paste"
Case 101
picnames(ct) = "tv"
Case 102
picnames(ct) = "city"
Case 103
picnames(ct) = "search"
Case 104
picnames(ct) = "star"
Case 105
picnames(ct) = "think"
Case 106
picnames(ct) = "bottle"
Case 108
picnames(ct) = "teacher"
Case 158
picnames(ct) = "news"
Case 159
picnames(ct) = "sign"
Case 160
picnames(ct) = "aahh"
Case 161
picnames(ct) = "ship"
Case 162
picnames(ct) = "bark"
Case 163
picnames(ct) = "fish"
Case 164
picnames(ct) = "welcome"
Case 165
picnames(ct) = "nurse"
Case 166
picnames(ct) = "cat"
Case 167
picnames(ct) = "xmas"
Case 168
picnames(ct) = "compass"
Case 169
picnames(ct) = "warning"
Case 171
picnames(ct) = "gallop"
Case 172
picnames(ct) = "flush"
Case 173 To 175, 177 To 181
picnames(ct) = "vispun"
Case 176
picnames(ct) = "dragon"
Case 182
picnames(ct) = "drome"
Case 183
picnames(ct) = "spoon"
Case 186
picnames(ct) = "chicken"
Case 187
picnames(ct) = "bicycle"
Case 188
picnames(ct) = "pig"
Case 189
picnames(ct) = "castle"
Case 190
picnames(ct) = "frog"
Case 191
picnames(ct) = "pyramid"
Case 192
picnames(ct) = "elephant"
Case 193
picnames(ct) = "sheep"
Case 194
picnames(ct) = "cow"
Case 195
picnames(ct) = "teeth"
Case 196
picnames(ct) = "heaven"
Case 197
picnames(ct) = "pants"
Case 198
picnames(ct) = "camel"
Case 199
picnames(ct) = "feet"
Case 200
picnames(ct) = "maid"
Case 201
picnames(ct) = "crown"
Case 202
picnames(ct) = "bell"
Case 203
picnames(ct) = "lion"
Case 204
picnames(ct) = "kiss"
Case 205
picnames(ct) = "mountain"
Case 206
picnames(ct) = "fire"
Case 207
picnames(ct) = "rainbow"
Case 208
picnames(ct) = "egg"
Case 209
picnames(ct) = "finger"
Case 210
picnames(ct) = "cheese"
Case 211
picnames(ct) = "monkey"
Case 212
picnames(ct) = "drum"
Case 213
picnames(ct) = "winter"
Case 214
picnames(ct) = "turtle"
Case 215
picnames(ct) = "rain"
Case 216
picnames(ct) = "dice"
Case 217
picnames(ct) = "path"
Case 218
picnames(ct) = "wave"
Case 219
picnames(ct) = "sun"
Case 220
picnames(ct) = "donkey"
Case 221
picnames(ct) = "butfly"
Case 222
picnames(ct) = "yowsah"
Case 223
picnames(ct) = "talk"
Case 224
picnames(ct) = "magic"
Case 225
picnames(ct) = "unlock"
Case 226
picnames(ct) = "snail"
Case 227
picnames(ct) = "snake"
Case 228
picnames(ct) = "town"
Case 229
picnames(ct) = "window"
Case 230
picnames(ct) = "hair"
Case 231
picnames(ct) = "angel"
Case 232
picnames(ct) = "hammer"
Case 233
picnames(ct) = "wind"
Case 234
picnames(ct) = "shoe"
Case 235
picnames(ct) = "sneeze"
Case 236
picnames(ct) = "bear"
Case 237
picnames(ct) = "overhill"
Case 238
picnames(ct) = "crops"
Case 239
picnames(ct) = "water"
Case 240
picnames(ct) = "polly"
Case 241
picnames(ct) = "thunder"
Case 242
picnames(ct) = "tiger"
Case 243
picnames(ct) = "sword"
Case 244
picnames(ct) = "violet"
Case 245
picnames(ct) = "church"
Case 246
picnames(ct) = "beepbeep"
Case 247
picnames(ct) = "willow"
Case 248
picnames(ct) = "jewels"
Case 249
picnames(ct) = "swan"
Case 250
picnames(ct) = "worm"
Case 251
picnames(ct) = "chain"
Case 252
picnames(ct) = "shell"
Case 253
picnames(ct) = "congrats"
Case 254
picnames(ct) = "eagle"
Case 255
picnames(ct) = "creaky"
Case 256
picnames(ct) = "poppy"
Case 257
picnames(ct) = "splash"
Case 258
picnames(ct) = "stag"
Case 259
picnames(ct) = "spider"
Case 260
picnames(ct) = "claw"
Case 261
picnames(ct) = "wolf"
Case 262
picnames(ct) = "violin"
Case 263
picnames(ct) = "weight"
Case 264
picnames(ct) = "mole"
Case 265
picnames(ct) = "arrow"
Case 266
picnames(ct) = "fishing"
Case 267
picnames(ct) = "giraffe"
Case 268
picnames(ct) = "duck"
Case 269
picnames(ct) = "icecream"
Case 270
picnames(ct) = "dolly"
Case 271
picnames(ct) = "goat"
Case 272
picnames(ct) = "rhino"
Case 273
picnames(ct) = "squirrel"
Case 274
picnames(ct) = "diner"
Case 275
picnames(ct) = "bridge"
Case 276
picnames(ct) = "faceoff"
Case 277
picnames(ct) = "kids"
Case 278
picnames(ct) = "prison"
Case 279
picnames(ct) = "bus"
Case 280
picnames(ct) = "grapes"
Case 281
picnames(ct) = "fox"
Case 282
picnames(ct) = "porcupine"
Case 283
picnames(ct) = "painter"
Case 284
picnames(ct) = "autumn"
Case 285
picnames(ct) = "gun"
Case 286
picnames(ct) = "whip"
Case 287
picnames(ct) = "sloth"
Case 288
picnames(ct) = "bubbles"
Case 289
picnames(ct) = "anvil"
Case 290
picnames(ct) = "rooster"
Case 291
picnames(ct) = "tear"
Case 292
picnames(ct) = "cloud9"
Case 293
picnames(ct) = "dove"
Case 294
picnames(ct) = "anchor"
Case 295
picnames(ct) = "snore"
Case 296
picnames(ct) = "station"
Case 297
picnames(ct) = "burglar"
Case 298
picnames(ct) = "hippo"
Case 299
picnames(ct) = "roo"
Case 300
picnames(ct) = "witch"
Case 301
picnames(ct) = "octowreck"
Case 302
picnames(ct) = "owl"
Case 303
picnames(ct) = "tricky"
Case 304
picnames(ct) = "butter"
Case 305
picnames(ct) = "crow"
Case 306
picnames(ct) = "alien"
Case 307
picnames(ct) = "hell"
Case 308
picnames(ct) = "stairs"
Case 309
picnames(ct) = "lightning"
Case 310
picnames(ct) = "seal"
Case 311
picnames(ct) = "dragonfly"
Case 312
picnames(ct) = "dressmaker"
Case 313
picnames(ct) = "pie"
Case 314
picnames(ct) = "cactus"
Case 315
picnames(ct) = "dino"
Case 316
picnames(ct) = "liar"
Case 317
picnames(ct) = "punch"
Case 318
picnames(ct) = "blood"
Case 319
picnames(ct) = "monk"
Case 320
picnames(ct) = "belfry"
Case 321
picnames(ct) = "sniff"
Case 322
picnames(ct) = "truth"
Case 323
picnames(ct) = "space"
Case 324
picnames(ct) = "honour"
Case 325
picnames(ct) = "nectar"
Case 326
picnames(ct) = "well"
Case 327
picnames(ct) = "harp"
Case 328
picnames(ct) = "locust"
Case 329
picnames(ct) = "chess"
Case 330
picnames(ct) = "running"
Case 331
picnames(ct) = "bones"
Case 332
picnames(ct) = "meds"
Case 333
picnames(ct) = "beaver"
Case 334
picnames(ct) = "harbour"
Case 335
picnames(ct) = "shake"
Case 336
picnames(ct) = "plum"
Case 337
picnames(ct) = "fountain"
Case 338
picnames(ct) = "ask"
Case 339
picnames(ct) = "jungle"
Case 340
picnames(ct) = "copter"
Case 341
picnames(ct) = "secret"
Case 342
picnames(ct) = "cake"
Case 343
picnames(ct) = "daisy"
Case 344
picnames(ct) = "badback"
Case 345
picnames(ct) = "demolition"
Case 346
picnames(ct) = "diamonds"
Case 347
picnames(ct) = "trojanhorse"
Case 348
picnames(ct) = "vulture"
Case 349
picnames(ct) = "thankyou"
Case 350
picnames(ct) = "care"
Case 351
picnames(ct) = "cogs"
Case 352
picnames(ct) = "barber"
Case 353
picnames(ct) = "poison"
Case 354
picnames(ct) = "solid"
Case 355
picnames(ct) = "stork"
Case 356
picnames(ct) = "weed"
Case 357
picnames(ct) = "demo"
Case 358
picnames(ct) = "scorpio"
Case 359
picnames(ct) = "stuck"
Case 360
picnames(ct) = "goose"
Case 361
picnames(ct) = "chameleon"
Case 362
picnames(ct) = "bingo"
Case 363
picnames(ct) = "map"
Case 364
picnames(ct) = "oyster"
Case 365
picnames(ct) = "cell"
Case 366
picnames(ct) = "mermaid"
Case 367
picnames(ct) = "forgive"
Case 368
picnames(ct) = "power"
Case 369
picnames(ct) = "leopard"
Case 370
picnames(ct) = "twins"
Case 371
picnames(ct) = "pelican"
Case 372
picnames(ct) = "adameve"
Case 373
picnames(ct) = "penguins"
Case 374
picnames(ct) = "greed"
Case 375
picnames(ct) = "ruined"
Case 376
picnames(ct) = "windmill"
Case 377
picnames(ct) = "ram"
Case 378
picnames(ct) = "lungs"
Case 379
picnames(ct) = "fly"
Case 380
picnames(ct) = "cupid"
Case 381
picnames(ct) = "spring"
Case 382
picnames(ct) = "dandelion"
Case 383
picnames(ct) = "crystalball"
Case 384
picnames(ct) = "shhh"
Case 385
picnames(ct) = "phoenix"
Case 386
picnames(ct) = "please"
Case 387
picnames(ct) = "apollo"
Case 388
picnames(ct) = "salt"
Case 389
picnames(ct) = "island"
Case 390
picnames(ct) = "command"
Case 391
picnames(ct) = "dew"
Case 392
picnames(ct) = "hay"
Case 393
picnames(ct) = "falcon"
Case 394
picnames(ct) = "jealous"
Case 395
picnames(ct) = "clean"
Case 396
picnames(ct) = "oil"
Case 397
picnames(ct) = "blush"
Case 398
picnames(ct) = "entertain"
Case 399
picnames(ct) = "character"
Case 400
picnames(ct) = "noble"
Case 401
picnames(ct) = "singers"
Case 402
picnames(ct) = "scream"
Case 403
picnames(ct) = "purity"
Case 404
picnames(ct) = "conscience"
Case 405
picnames(ct) = "wisdom"
Case 406
picnames(ct) = "immortal"
Case 407
picnames(ct) = "daffodils"
Case 408
picnames(ct) = "piano"
Case 409
picnames(ct) = "goal"
Case 410
picnames(ct) = "target"
Case 411
picnames(ct) = "different"
Case 412
picnames(ct) = "miser"
Case 413
picnames(ct) = "math"
Case 414
picnames(ct) = "nothing"
Case 415
picnames(ct) = "dawn"
Case 416
picnames(ct) = "sunset"
Case 417
picnames(ct) = "dusk"
Case 418
picnames(ct) = "morality"
Case 419
picnames(ct) = "magnet"
Case 420
picnames(ct) = "gravity"
Case 421
picnames(ct) = "grain"
Case 422
picnames(ct) = "history"
Case 423
picnames(ct) = "philosophy"
Case 424
picnames(ct) = "unicorn"
Case 425
picnames(ct) = "hope"
Case 426
picnames(ct) = "amigos"
Case 427
picnames(ct) = "habits"
Case 428
picnames(ct) = "echo"
Case 429
picnames(ct) = "nightingale"
Case 430
picnames(ct) = "market"
Case 431
picnames(ct) = "meditate"
Case 432
picnames(ct) = "faith"
Case 433
picnames(ct) = "ivy"
Case 434
picnames(ct) = "dna"
Case 435
picnames(ct) = "law"
Case 436
picnames(ct) = "culture"
Case 437
picnames(ct) = "fate"
Case 438
picnames(ct) = "influence"
Case 439
picnames(ct) = "stupor"
Case 440
picnames(ct) = "colour"
Case 441
picnames(ct) = "armour"
Case 442
picnames(ct) = "attitude"
Case 443
picnames(ct) = "suspicion"
Case 444
picnames(ct) = "farewell"
Case 445
picnames(ct) = "audience"
Case 446
picnames(ct) = "spiral"
Case 447
picnames(ct) = "hotpotato"
Case 448
picnames(ct) = "begin"
Case 449
picnames(ct) = "passion"
Case 450
picnames(ct) = "sunflower"
Case 451
picnames(ct) = "wheel"
End Select
Next

VMSynth = CSFT
SndMidInit

App.HelpFile = App.Path & "\MyReels.chm"
prepspin = False

' Shell ("regsvr32.exe actcndy4.ocx /s"), vbMinimizedFocus

Load Pokemach
If procend = True Then
procend = False
 If resCheck = False Then
 QuitNow
 Exit Sub
 ElseIf gt(200) <> 1 Then 'Avoids warning on setup/update launch- but excludes rare case of new config while installer running!
  If FindWindowPartial(True) <> 0 Then
  Pokemach.waitimer.Enabled = False
  response = MsgBox("MyReels cannot run with the Installer or Updater open." & vbNewLine & "If the Updater or Installer is running, click ""OK"", close it and retry the game from a shortcut." & vbNewLine & "Else click ""Cancel"" for possible errors.", vbSystemModal Or vbOKCancel)
   If response = vbOKCancel Then
   QuitNow
   Exit Sub
   Else
   Pokemach.waitimer.Enabled = True
   End If
  End If
 End If

Pokemach.Show

On Error Resume Next

If FindShellTaskBar Then Dotaskwindow Pokemach

' Must close handles if midi off!
If gt(185) = 0 Then ZapMidihWnd

If gt(200) = 1 Then
' Shell ("winhlp32.exe  -N 17 " & App.Path & "\MyReels.hlp"), vbNormalFocus
 On Error GoTo Filedeleteerror
 If FileExists(App.Path & "\building.wav") Then Kill (App.Path & "\building.wav")
 If FileExists(App.Path & "\peace.wav") Then Kill (App.Path & "\peace.wav")
 If FileExists(App.Path & "\fabulous.wav") Then Kill (App.Path & "\fabulous.wav")
 If FileExists(App.Path & "\yeah.wav") Then Kill (App.Path & "\yeah.wav")
For ct = 1 To 50000
   DoEvents
   Sleep 50
   If justcached = True Then
   DoEvents
   c.CallHelp
   Exit For
   End If
   Next ct
 DoEvents
 gt(200) = 0 ' Trigger turned off for good
 End If
' gt(200) = 1 'can set trigger here


Else
 QuitNow
 If runAdmin = True Then
  DoEvents
  c.RunElevated App.Path & "\" & App.EXEName
 End If
procend = False
End If

isMainActive = False

Exit Sub
Filedeleteerror:
ShowError
MsgBox "Please check file permissions in the selected directory and restart program!", vbExclamation
QuitNow
End Sub
Private Sub QuitNow()
Stopnoise 2
TaskbarVisible = FindShellTaskBar()
Dotaskwindow , , True
Unload Pokemach
Set Pokemach = Nothing
Unload frmSplsh
Set frmSplsh = Nothing
c.CloseMutexhandle
End Sub
Public Sub SndMidInit()
CanPlayWaves = waveOutGetNumDevs()
curmididev = 0
strPlaying = ""
For ct = 0 To midiDevTot - 1
hmidi(ct) = 0
hmsmidi(ct) = 0
Next
If gt(185) = 0 Then Midichg 'not on chgdir
If CanPlayWaves = False Then MsgBox "Couldn't open wave device"
End Sub
Private Function ZapMidihWnd(Optional ByVal hWndIndex As Long = -1) As Boolean
Dim mhWnd As Long, mshWnd As Long, noErr As Boolean, tmp As String
DoEvents
' may need midiOutReset but this does not send EOX byte. midiStreamStop may be buggy clearing handle prematurely

 noErr = Not c.VistaorLater And Not c.XP
 tmp = ""

If hWndIndex = -1 Then
 For ct = 0 To midiDevTot - 1
 mhWnd = hmidi(ct)
 mshWnd = hmsmidi(ct)
 On Error Resume Next
 If mshWnd > 0 And VMSynth = True Then
  If midiStreamStop(mshWnd) = 0 Then  ' And midiStreamClose(mshWnd) = 0 ' crashes
  hmsmidi(ct) = 0
  mshWnd = 0
  Else
  Exit For
  End If
 End If
 If mhWnd > 0 And VMSynth = False Then
  If noErr Then
  mhWnd = 0
  Else
   If midiOutClose(mhWnd) = 0 Then
   hmidi(ct) = 0
   mhWnd = 0
   Else
   Exit For
   End If
  End If
 End If
 Next
Else
 mshWnd = hmsmidi(hWndIndex)
'MsgBox "before mshWnd: " & mshWnd & " gt(185): " & gt(185)
 If mshWnd > 0 And VMSynth = True Then
  If midiStreamStop(mshWnd) = 0 Then ' And midiStreamClose(mshWnd) = 0 Then 'crashes
'MsgBox "after mshWnd: " & mshWnd & " gt(185): " & gt(185)
  hmsmidi(hWndIndex) = 0
  mshWnd = 0
  End If
 End If
 mhWnd = hmidi(hWndIndex)
 If mhWnd > 0 And VMSynth = False Then
  If noErr Then
  mhWnd = 0
  Else
   If midiOutClose(mhWnd) = 0 Then
   hmidi(hWndIndex) = 0
   mhWnd = 0
   End If
  End If
 End If
End If

If mshWnd > 0 Then tmp = "Midi Stream"
If mhWnd > 0 Then
 If mshWnd = 0 Then
 If VMSynth = False Then tmp = "MidiOut"
 Else
 If VMSynth = False Then tmp = tmp & " and MidiOut handles"
 End If
Else
If mshWnd > 0 Then tmp = tmp & " handle"
End If
 
 
If tmp = "" Then
ZapMidihWnd = False
Else
MsgBox "There was a problem closing " & tmp & ". Restart MyReels recommended."
ZapMidihWnd = True
End If
DoEvents
End Function
Public Sub REORGSLOT()
Dim Llong(666) As Long, Striing(100) As String, DefStriing(50) As Long, DoReorg As Boolean
DoReorg = False

Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")
With rectemp
On Error GoTo ErrHndl
.MoveFirst

If IsNull(![Lorng]) = True Then DoReorg = True
End With

If DoReorg = True Then

Set rectemp = dbsCurrent.OpenRecordset("Inpoot")

With rectemp
.MoveFirst


For ct = 1 To 282
.MoveNext
Next


' 345 moves before gt, 116 after
For ct = 1 To 661
Llong(ct) = ![Lorng]
.MoveNext
Next


For ct = 1 To 100
Striing(ct) = ![Streeng]
.MoveNext
Next


End With

Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")
With rectemp
.MoveFirst
For ct = 1 To 661
.Edit
![Lorng] = Llong(ct)
.Update
.MoveNext
Next
For ct = 1 To 382
.DELETE
.MoveNext
Next


End With


Set rectemp = dbsCurrent.OpenRecordset("SELECT Streeng FROM Inpoot")
With rectemp
.MoveFirst
For ct = 1 To 100
.Edit
![Streeng] = Striing(ct)
.Update
.MoveNext
Next


End With
End If


' get gt(192)
Set rectemp = dbsCurrent.OpenRecordset("Inpoot", dbOpenForwardOnly)

' 345 moves before gt, 116 after
For ct = 1 To 536
rectemp.MoveNext
Next
gt(192) = rectemp![Lorng]




If gt(192) < 3 Then
Set rectemp = dbsCurrent.OpenRecordset("Inpoot")

With rectemp
.MoveFirst

For ct = 1 To 5
.AddNew     ' Defaultgt now 65
' Initialise more defaults later
![Lorng] = 0
.Update
Next



.MoveFirst
For ct = 1 To 100
Striing(ct) = ![Streeng]
.MoveNext
Next

End With


Set rectemp = dbsCurrent.OpenRecordset("SELECT Streeng FROM Inpoot")
With rectemp
.MoveFirst
For ct = 1 To 200
.Edit
Select Case ct
Case Is > 150
![Streeng] = ""
Case Is < 51
![Streeng] = Striing(ct)
Case Is > 100
![Streeng] = Striing(ct - 50)
Case Else
![Streeng] = ""
End Select
.Update
.MoveNext
Next
End With
End If

If gt(192) < 4 Then     ' Prior to 2.2.3
' Swap 185 with 195
Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")
With rectemp
.MoveFirst
.Move 529
gt(195) = rectemp![Lorng]
.Move 10
.Edit
rectemp![Lorng] = gt(195)
.Update


End With

End If



If gt(192) < 5 Then     ' Prior to 2.2.20
' Fix 184, 187
Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")
With rectemp
.MoveFirst
.Move 528 ' gt (184) = 0 FSFG bonus default
.Edit
rectemp![Lorng] = 0
.Move 3
.Edit
rectemp![Lorng] = -1 ' init for midipos

' NOTE: ADD THIS AT END OF FUTURE VERSION CHANGES!
.Move 5 'Or whatever to get gt(192)
.Edit
rectemp![Lorng] = 5


.Update


End With


End If



Exit Sub
ErrHndl:
ShowError

End Sub
Public Function PlayMidiFile(ByVal FileName As String, interupt As Boolean) As Boolean
Dim reply As String * 255, f As String
PlayMidiFile = False

If midirc <> 0 Or FileName = "" Then Exit Function  ' Not valid midi device

' Parse filename (or use GetShortPathName)

  If c.XP Or c.VistaorLater Then
  f = c.GetShortFileName(FileName)
  Else
  f = c.GetShortName(FileName)
  End If
'MCI_DEVTYPE_SEQUENCER = 523 ;
nRet = mciSendString("status " & strPlaying & " mode", reply, 255, 0)
DoEvents

  If interupt = False Then
  If Left$(reply, 7) = "playing" Then Exit Function
  strPlaying = f
  End If

  If f = "" Then
  nRet = 1
  Else
  StopMidiFile
  strPlaying = f 'when interupt = True
  If VMSynth = False Then
     'just Roland GSW
    If ZapMidihWnd(curmididev) Then Exit Function
    End If
  DoEvents
  nRet = mciSendString("Open " & strPlaying, vbNullString, 0, 0)
  nRet = mciSendString("Play " & strPlaying & " from 0", vbNullString, 0, 0)
  End If

PlayMidiFile = (nRet = 0)

End Function
Private Function CSFT()
Dim MMPW As String, MMP As String
MMPW = c.SysDir & "SysWOW64\" & "MIDIMapper\"
MMP = c.SysDir & "MIDIMapper\"
CSFT = True
If (findafile(MMPW, MM)) Then Exit Function
If (findafile(c.SysDir(True), VM)) Then Exit Function
If (findafile(MMP, MM)) Then Exit Function
CSFT = False
End Function
Public Sub StopMidiFile()
nRet = mciSendString("Stop " & strPlaying, vbNullString, 0, 0)
nRet = mciSendString("Seek " & strPlaying & " to 0", vbNullString, 0, 0)
nRet = mciSendString("Close " & strPlaying, vbNullString, 0, 0)
DoEvents
End Sub
Public Function Midichg(Optional ByVal inputDev As Long = -1)
Dim caps As MIDIOUTCAPS

Select Case inputDev
Case -1
' open
 If gt(185) > 0 Then
 If strPlaying <> "" Then StopMidiFile
  If ZapMidihWnd(curmididev) Then
    If OpenmidiDev <> 0 Then
    gt(185) = 0
    curmididev = 0
    End If
  Else
  Midichg = gt(185)
  gt(185) = 0
  Exit Function
  End If
 Else ' gt(185) = 0
  midiDevTot = midiOutGetNumDevs()
  If midiDevTot > 4 Then midiDevTot = 4
  ' Set first device as midi mapper then get rest of midi devices
  nRet = midiOutGetDevCaps(-1, caps, Len(caps))
  If nRet = 0 And caps.wTechnology = MOD_MAPPER Then
   For ct = 0 To (midiDevTot - 1)
    If midiOutGetDevCaps(ct, caps, Len(caps)) = 0 Then
    curmididev = ct
    Exit For
    End If
   Next
    If OpenmidiDev = 0 Then
    gt(185) = curmididev + 1
    Else
    gt(185) = 0
    End If
  Else
   MsgBox "Problem with the Midi Mapper, code: " & nRet & "."
  End If
  End If
Case 0
' close
    If strPlaying <> "" Then
    StopMidiFile
    strPlaying = ""
    End If
  gt(185) = 0
    If ZapMidihWnd(curmididev) Then
    Midichg = 1
    Exit Function
    End If
Case Else
  If (inputDev > midiDevTot) Then
  Midichg = 0
  Exit Function
  Else
    curmididev = inputDev - 1
    If OpenmidiDev = 0 Then
    gt(185) = curmididev + 1
    Else
    curmididev = gt(185) - 1
    End If
  End If
End Select
Midichg = gt(185)
DoEvents
End Function
Private Function OpenmidiDev()
  Dim tmp As String, errStr As String, badStream As Boolean, hWndAddr As Long, hwndMidi As Long, midiAddr As Long, oldmididev As Long
  badStream = False
  If hmidi(curmididev) = 0 Then midirc = midiOutOpen(hmidi(curmididev), curmididev, 0, 0, 0)

  If VMSynth = True And midirc = 0 And hmsmidi(curmididev) = 0 Then
  oldmididev = curmididev
  hwndMidi = 0
  hWndAddr = VarPtr(hwndMidi)
  midiAddr = VarPtr(curmididev)
  midirc = midiStreamOpen(hWndAddr, midiAddr, 1, 0, 0, CALLBACK_NULL)
  curmididev = oldmididev ' Function should not change device to -2 on fail!
  hmsmidi(curmididev) = hwndMidi
  If midirc = 4 Then midirc = 0 ' MMSYSERR_ALLOCATED: GS Wavetable is currently allocated. We know
  If midirc <> 0 Then badStream = True
  End If

  Select Case midirc
  Case MMSYSERR_ERROR
  tmp = "MMSYSERR_UNSPECIFIED_ERROR"
  Case MMSYSERR_BADDEVICEID
  tmp = "MMSYSERR_BADDEVICEID"
  Case MMSYSERR_ALLOCATED
  tmp = "MMSYSERR_ALLOCATED"
  Case MMSYSERR_INVALHANDLE
  tmp = "MMSYSERR_INVALHANDLE"
  Case MMSYSERR_NOMEM
  tmp = "string problem"
  Case MMSYSERR_INVALPARAM
  tmp = "Invalid handle or flags issue"
  Case 64
  tmp = "MIDIERR_UNPREPARED"
  Case 65
  tmp = "MIDIERR_STILLPLAYING"
  Case 67
  tmp = "MIDIERR_NOTREADY"
  Case 69
  tmp = "MIDIERR_INVALIDSETUP"
  Case 257
  tmp = "MCIERR_INVALID_DEVICE_ID"
  Case 262
  tmp = "MCIERR_HARDWARE"
  Case 263
  tmp = "MCIERR_INVALID_DEVICE_NAME"
  Case 265
  tmp = "MCIERR_DEVICE_OPEN"
  Case 266
  tmp = "MCIERR_CANNOT_LOAD_DRIVER"
  Case 337
  tmp = "MCIERR_SEQ_PORT_INUSE"
  Case 338
  tmp = "MCIERR_SEQ_PORT_NONEXISTENT"
  Case 339
  tmp = "MCIERR_SEQ_PORT_MAPNODEVICE"
  Case 340
  tmp = "MCIERR_SEQ_PORT_MISCERROR"
  Case 341
  tmp = "MCIERR_SEQ_TIMER"
  Case 342
  tmp = "MCIERR_SEQ_PORTUNSPECIFIED"
  Case 343
  tmp = "MCIERR_SEQ_NOMIDIPRESENT"
  Case Else
  tmp = ""
  End Select

 If midirc > 0 Then
  If badStream = True Then
  errStr = "MidiStreamOpen failed with"
  Else
  errStr = "midiOutOpen failed with"
  End If

  If tmp = "" Then
  MsgBox errStr & "code: " & midirc & "!"
  Else
  MsgBox errStr & ": " & tmp & "!"
  End If
 End If

OpenmidiDev = midirc
End Function
Public Function Stopnoise(Optional LoadNext As Integer = 0) As Boolean
If strPlaying <> "" Then StopMidiFile
PlaySndF "", True
DoEvents
If gt(185) > 0 Then
Select Case LoadNext 'ChgDir is 2
Case 0
gt(185) = 0
Stopnoise = ZapMidihWnd
Case 1
Stopnoise = ZapMidihWnd
Case 2
Stopnoise = ZapMidihWnd(gt(185) - 1)
End Select
End If
End Function
Public Sub PlaySndF(ByVal FileName As String, Optional stopnow As Boolean = False)

If CanPlayWaves = False Or gt(186) < 2 Then Exit Sub


If oldsndfname <> "" Then Call PlaySound("", 0&, SND_PURGE Or SND_NODEFAULT)

If FileName = "Thumbb" Then

If gt(186) = 3 Then ' Quotesound turned on

FileName = picnames(BmpIX)


Select Case BmpIX
Case 41 To 45, 47 To 50, 53
FileName = "cocktail"
Case 46, 51, 52, 106
FileName = "pour"
End Select





Call PlaySound(App.Path & "\" & FileName & ".wav", 0&, SND_ASYNC)

End If

Else
If stopnow = False Then Call PlaySound(FileName, 0&, SND_ASYNC)
End If

oldsndfname = FileName
End Sub
Public Function SystemSoundNames(Snds() As SystemSoundDefinitions) As Long
Dim i As Long, j As Long, Ndx As Long
Dim subsectkeys() As String, sectkeys() As String, SKeys() As String, sectCount As Long, subsectct As Long

ct = 0
With c
.ClassKey = HKEY_CURRENT_USER
.SectionKey = "AppEvents\Schemes\Apps"
If .EnumerateSections(SKeys(), sectCount) = True Then
ReDim sectkeys(1 To sectCount) As String
For i = 1 To sectCount
sectkeys(i) = SKeys(i)
Next
ReDim SKeys(1 To sectCount) As String

For i = 1 To sectCount

' subsection
.SectionKey = "AppEvents\Schemes\Apps\" & sectkeys(i)
If .EnumerateSections(SKeys(), subsectct) = True Then


ReDim subsectkeys(0 To subsectct) As String
For j = 1 To subsectct
subsectkeys(j) = SKeys(j)
Next
ReDim SKeys(1 To subsectct) As String


For j = 1 To subsectct
.ValueKey = subsectkeys(j)
If .EnumerateValues(SKeys(), Ndx) = True Then


ReDim Preserve Snds(0 To ct) As SystemSoundDefinitions

  With Snds(ct)
  .GroupName = c.SectionKey
  .SoundName = c.ValueKey
  .RegKey = c.SectionKey & "\" & subsectkeys(j)
  .Current = c.CurrentValue
  .Default = c.DefaultValue
  End With

ct = 1 + ct

End If


Next

End If

Next



End If
End With
SystemSoundNames = ct
End Function
Private Function FindShellTaskBar() As Boolean

On Error Resume Next

hWnd = FindWindowEx(0&, 0&, g_cstrShellTaskBarWnd, vbNullString)

If hWnd = 0 Then
FindShellTaskBar = False
Else
FindShellTaskBar = True
End If
    
End Function
Private Function FindShellWindow() As Long
'not used

On Error Resume Next

hWnd = FindWindowEx(0&, 0&, g_cstrShellViewWnd, vbNullString)

FindShellWindow = hWnd
    
End Function
Private Sub HideShowWindow(ByVal hWnd As Long, Optional ByVal Hide As Boolean = False)

Dim lngShowCmd As Long

On Error Resume Next

If Hide = True Then
lngShowCmd = SW_HIDE
    
Else
lngShowCmd = SW_SHOW
    
End If

Call ShowWindow(hWnd, lngShowCmd)


End Sub
Public Sub Dotaskwindow(Optional objform As Object, Optional ByVal Hidenow As Boolean = False, Optional noobjekt As Boolean)
On Error Resume Next
If noobjekt = False Then
If resX > 1 Then objform.Top = 0
End If

If TaskbarVisible And Not (c.VistaorLater()) Then
 If Hidenow = False Then
 Call HideShowWindow(hWnd)
 Else
 Call HideShowWindow(hWnd, True)
 End If
End If

End Sub
Public Function Docaptions(numb As Long, Optional sortpicnamesnow As Boolean, Optional sortpicnames As Boolean) As String
If sortpicnamesnow = True Then
    For ct = 1 To 450
    picnamesort(ct) = picnames(ct)
    Next
    c.InsertSortStringsStart picnamesort
Else
    If sortpicnames - False Then
    Docaptions = CStr(numb) & "     " & picnamesort(numb)
    Else
    Docaptions = CStr(numb) & "     " & picnames(numb)
    End If
End If
End Function
Public Sub Getzhiddnstat()
    For ct = 1 To 450
    If picnamesort(-zhiddnstatus) = picnames(ct) Then
    zhiddnstatus = -ct
    Exit Sub
    End If
    Next
End Sub
Public Sub BitmapDb(picsloadstatus As Long, RHSoffset As Long, RHSpoz As Long, LHSpoz As Long, imgselprevindex As Long, wipeimgsel As Boolean, Optional imgseljustchanged As Boolean = False)
Dim i As Long


On Error GoTo Quotezerr

sDatabaseName = Stringvars(3) & "Quotes.s$t"


If OpenDb(sDatabaseName, 2) = False Then Exit Sub


Set rectemp = gdbCurrentDB.OpenRecordset("Quotez")


With rectemp

.Index = gdbCurrentDB.TableDefs(.Name).Indexes(0).Name

Select Case picsloadstatus

Case -3 ' get imgselprev


.MoveFirst
.Move RHSpoz
imgselprevindex = ![Bmpindex]
If imgselprevindex > 0 Then
wipeimgsel = False
Else
wipeimgsel = True
End If


Case -2 ' Fill quotetext

gt(196) = 0 ' zhidden picnames: init no of different DBpics

.MoveFirst
Quotebrs.quotetext.Clear
Quotebrs.quotelist.Clear
gt(196) = 0 ' init no of different DBpics
For ct = 0 To gt(191) - 1
If ![Bmpindex] > gt(196) Then gt(196) = ![Bmpindex]
' If ![Bmpindex] = 81 Then
Quotebrs.quotetext.AddItem ![Quotestr]
Quotebrs.quotelist.AddItem CStr(ct + 1)
' End If

.MoveNext
Next


Quotebrs.quotetext.ListIndex = LHSpoz
Quotebrs.quotetext.Text = Quotebrs.quotetext.List(LHSpoz)




Case -1 ' Delete quote

If gt(191) = 1 Then
.Close
GoTo Quotezerr
End If

.MoveFirst
.Move LHSpoz


ct1 = ![Bmpindex]
i = ct1
' Is bmpindex(LHSpoz) last of its type?


.MoveFirst
For ct = 0 To gt(191) - 1

If ![Bmpindex] = ct1 And ct <> LHSpoz Then ct1 = 0
.MoveNext
Next



  If ct1 > 0 Then ' Last of its type
  .MoveFirst

  For ct = 0 To gt(191) - 1
    If ![Bmpindex] > ct1 Then
    .Edit
    ![Bmpindex] = ![Bmpindex] - 1
    .Update
    End If
   .MoveNext
  Next
  Else



  ' Is there a bmp really here?
  .MoveFirst
  .Move LHSpoz


    If IZempty = False Then
    response = MsgBox("Other assignments to this thumbnail will be deleted. OK?", vbYesNo)
      If response = vbNo Then
      .Close
      GoTo Quotezerr
      Else
      .MoveFirst
      For ct = 0 To gt(191) - 1
      Select Case ![Bmpindex]
      Case i
      .Edit
      ![Bmpindex] = 0
      .Update
      Case Is > i ' reduce uppers
      .Edit
      ![Bmpindex] = ![Bmpindex] - 1
      .Update
      End Select
      .MoveNext

      Next
      End If
    End If

  End If

  ' Now decrement apt bmpindex
  gt(191) = gt(191) - 1

  ' And current position
  If gt(195) = 0 Then
  If gt(188) >= LHSpoz Then gt(188) = gt(188) - 1
End If

' Now delete
.MoveFirst
.Move LHSpoz
.DELETE


If LHSpoz > 0 Then LHSpoz = LHSpoz - 1



Case 0  ' Fill images on RH screen

  If zhiddnstatus < 0 Then ' find real RHSoffset to match passed RHSoffset
  zhiddnstatus = -zhiddnstatus
  .MoveFirst
  For ct = 0 To gt(191) - 1
  If ![Bmpindex] = 0 Then DoEvents
    If ![Bmpindex] = zhiddnstatus Then
    RHSoffset = ct
    Exit For
  Else
    .MoveNext
  End If
  Next
  End If

  For ct = 0 To 8

  If ct + RHSoffset < gt(191) Then
  Quotebrs.RHSquote(ct).Visible = True
  Quotebrs.Spinimage(ct).Visible = True
  .MoveFirst
  .Move RHSoffset + ct

    If Len(![Quotestr]) > 90 * resX Then
    Quotebrs.RHSquote(ct) = Left(![Quotestr], 90 * resX) & " ..."
    Else
    Quotebrs.RHSquote(ct) = Left(![Quotestr], 90 * resX)
    End If

    i = ![Bmpindex]
    If i > 0 Then
       Quotebrs.lblindex(ct).Caption = i
       If IZempty = True Then
         Quotebrs.lblindex(ct).ForeColor = &H80FF&
         .MoveFirst
         For ct1 = 0 To gt(191) - 1
         If ![Bmpindex] = i And IZempty = False Then Exit For
         .MoveNext
         Next
         Quotebrs.lblindex(ct).ToolTipText = "Thumbnail located at: " & CStr(ct1 + 1)
       Else
         Quotebrs.lblindex(ct).ForeColor = &HFFFF00
         Quotebrs.lblindex(ct).ToolTipText = ""
       End If

    If Chunker(rectemp, False, ct) = False Then GoTo Quotezerr
       Set Quotebrs.Spinimage(ct).Picture = LoadPicture(loaddirectory & "q" & ct & ".bmp")
    Else
    ' blank pic
      Set Quotebrs.Spinimage(ct).Picture = LoadPicture(App.Path & "\blank.bmp")
      Quotebrs.lblindex(ct).Caption = ""
    End If
Else
Quotebrs.RHSquote(ct).Visible = False
Quotebrs.Spinimage(ct).Visible = False
Quotebrs.lblindex(ct).Caption = ""
End If


Next








Case 1  ' Put bmp in imgsel on quotetext_click


.MoveFirst
.Move LHSpoz
i = ![Bmpindex]

 If ![Bmpindex] > 0 Then

    If IZempty = False Then ' bmp here



    If Chunker(rectemp, False) = False Then GoTo Quotezerr

    Else    ' No bmp here
    ' Search for picture
    .MoveFirst
    For ct = 0 To gt(191) - 1
    If ![Bmpindex] = i Then
    If IZempty = False Then Exit For ' found bmp here
    End If
    .MoveNext
   ' No pic
    Next
    
    If Chunker(rectemp, False) = False Then GoTo Quotezerr
    End If







Quotebrs.imgsel.Picture = LoadPicture(loaddirectory & "q0.bmp")
imgselprevindex = i
wipeimgsel = False
Else
' Clear imgsel
wipeimgsel = True
Quotebrs.imgsel.Picture = LoadPicture(App.Path & "\blank.bmp")
End If






Case 2 ' Assigns text & picture to position LHSpoz of db


' Update text first
If Quotebrs.quotetext.Text = "" Then
MsgBox "Please type some text"
.Close
gt(191) = -gt(191)
GoTo Quotezerr
End If

Quotebrs.quotetext.Text = Left(Quotebrs.quotetext.Text, 172)
If Namevalid(2, Quotebrs.quotetext.Text) = False Then
.Close
gt(191) = -gt(191)
GoTo Quotezerr
End If

' Now check for dups
.MoveFirst


For ct = 0 To gt(191) - 1
    If Quotebrs.quotetext.Text = ![Quotestr] Then
        If ct <> LHSpoz Then
        .Close
        MsgBox Msgdup, vbOKOnly
        gt(191) = -gt(191)
        GoTo Quotezerr
        End If
    End If
.MoveNext
Next

.MoveFirst
.Move LHSpoz
ct1 = ![Bmpindex]

.Edit
![Quotestr] = Quotebrs.quotetext.Text
.Update


' LHSpoz changes if textchanges
.MoveFirst

For ct = 0 To gt(191) - 1
If ![Quotestr] = Quotebrs.quotetext.Text Then Exit For
.MoveNext
Next
LHSpoz = ct



.MoveFirst
' .Move LHSpoz - 1


Quotebrs.quotetext.Clear

For ct = 0 To gt(191) - 1
Quotebrs.quotetext.AddItem ![Quotestr]
.MoveNext
Next

' Much slower
' For ct = LHSpoz - 1 To 0 Step -1
' Quotebrs.quotetext.List(ct) = ![Quotestr]
' If Not .BOF Then .MovePrevious
' Next


.MoveFirst
.Move LHSpoz
Quotebrs.quotetext.List(LHSpoz) = ![Quotestr]
Quotebrs.quotetext.ListIndex = LHSpoz



' Update picture next
If imgseljustchanged = False Then
.Close
GoTo Quotezerr
End If

.MoveFirst
.Move LHSpoz





If wipeimgsel = True Then

i = ct1
If ct1 = 0 Then
' Nothing to replace
.Close
GoTo Quotezerr
End If

' Is this the last assignment?

.MoveFirst
For ct = 0 To gt(191) - 1
If ![Bmpindex] = ct1 And ct <> LHSpoz Then
ct1 = 0
Exit For
End If
.MoveNext
Next


.MoveFirst
.Move LHSpoz



    ' Is a bmp really here?
    If IZempty = False Then

    If ct1 = 0 Then
    response = MsgBox("Other assignments to the thumbnail just replaced will be erased. OK?", vbYesNo)
        If response = vbNo Then
        gt(191) = -gt(191)
        .Close
        GoTo Quotezerr
        End If
    End If
    ' Clear picture assignment from DB
    .Edit
    ![Bmpfile] = Null
    ![Bmpindex] = 0
    .Update


    If ct1 = 0 Then ' zero others if > 1 entry
    .MoveFirst
    For ct = 0 To gt(191) - 1
        If ![Bmpindex] = i Then
        .Edit
        ![Bmpindex] = 0
        .Update
        End If
    .MoveNext
    Next
    End If

    .MoveFirst
    For ct = 0 To gt(191) - 1
        If ![Bmpindex] > i Then
        .Edit
        ![Bmpindex] = ![Bmpindex] - 1
        .Update
        End If
    .MoveNext
    Next

    Else ' IZempty true- Clear picture assignment from DB
    .Edit
    ![Bmpindex] = 0
    .Update
    End If



Else    ' wipeimgsel false


    If IZempty = False Then

    If ct1 > 0 Then ' ct1 should never be 0 anyway!

        ' Don't replace root bmp with same
        If imgselprevindex = ct1 Then
        ' Still need to update RHS
        .Close
        GoTo Quotezerr
        End If


    'check for other assignments
    .MoveFirst
    For ct = 0 To gt(191) - 1
        If ![Bmpindex] = ct1 And ct <> LHSpoz Then
        response = MsgBox("Other assignments to the thumbnail that was here will point to the new picture. OK?", vbYesNo)
        If response = vbNo Then
        gt(191) = -gt(191)
        .Close
        GoTo Quotezerr
        Else
        Exit For
        End If
        End If
    .MoveNext
    Next

    End If ' ct1 > 0


   .MoveFirst
   .Move LHSpoz

        If imgselprevindex = 0 Then
        If FileExists(loaddirectory & "q0.bmp") Then Kill (loaddirectory & "q0.bmp")
        SavePicture Quotebrs.picTruesize.Image, loaddirectory & "q0.bmp"


        If Chunker(rectemp, True) = False Then GoTo Quotezerr

        .Edit
        ![Bmpindex] = ct1
        .Update
        Else
        ' Point assignments to new picture, condense

        If imgselprevindex > ct1 Then imgselprevindex = imgselprevindex - 1

        .MoveFirst
        For ct = 0 To gt(191) - 1
        If ![Bmpindex] > ct1 Then
        .Edit
        ![Bmpindex] = ![Bmpindex] - 1
        .Update
        ElseIf ![Bmpindex] = ct1 Then
        .Edit
        ![Bmpindex] = imgselprevindex
        .Update
        End If
        .MoveNext
        Next



        ' Must clear
        .MoveFirst
        .Move LHSpoz
        .Edit
        ![Bmpfile] = Null
        .Update

        End If  ' imgselprev


    Else    ' izempty true
        If imgselprevindex = 0 Then
        ' Need next available bmpindex number for new bmp from scratch
        .MoveFirst
        For ct = 0 To gt(191) - 1
        If ![Bmpindex] > ct1 Then ct1 = ![Bmpindex]
        .MoveNext
        Next

        .MoveFirst
        .Move LHSpoz
        .Edit
        ![Bmpindex] = ct1 + 1
        .Update


        If FileExists(loaddirectory & "q0.bmp") Then Kill (loaddirectory & "q0.bmp")
        SavePicture Quotebrs.picTruesize.Image, loaddirectory & "q0.bmp"


        If Chunker(rectemp, True) = False Then GoTo Quotezerr
        Else
        .Edit
        ![Bmpindex] = imgselprevindex
        .Update
        End If
    End If  ' izempty


End If





Case 3  ' add new record, picture always cleared

' Now check for dups
.MoveFirst

For ct = 0 To gt(191) - 1
    If ![Quotestr] = "New" Then
    .Close
    MsgBox Msgdup, vbOKOnly
    gt(191) = -gt(191)
    GoTo Quotezerr
    End If
.MoveNext
Next



wipeimgsel = True

.AddNew
![Quotestr] = "New"
![Bmpindex] = 0
![Bmpfile] = Null
.Update


.MoveFirst
For ct = 0 To gt(191)
' trap null to get LHSpoz
If .Fields(0).Value = "New" Then Exit For
.MoveNext
Next

imgselprevindex = 0
LHSpoz = ct
gt(191) = gt(191) + 1

' And current position
If gt(195) = 0 Then
If gt(188) >= LHSpoz Then gt(188) = gt(188) + 1
End If

Quotebrs.imgsel.Picture = LoadPicture(App.Path & "\blank.bmp")



Case 4  ' exit screen


' Get rid of null quotes
.MoveFirst
For ct = 0 To gt(191) - 1
If .Fields(0).Value = "" Then .DELETE
.MoveNext
Next



If gt(189) = 1 Then
' Remove any existing dest. file.
Genopts.Cdlg.FileName = ""
Genopts.Cdlg.FilterIndex = 1
Genopts.Cdlg.Filter = "text (*.,txt)|*.txt|All Files (*.*)|*.*"
' Specify default filter
Genopts.Cdlg.DialogTitle = " Write to Text File"

Genopts.Cdlg.InitDir = Stringvars(3)
Genopts.Cdlg.CancelError = False

Genopts.Cdlg.ShowOpen

If Not FileExists(Genopts.Cdlg.FileName) Or Genopts.Cdlg.FileName = "" Then
.Close
GoTo Quotezerr
End If

hfile = FreeFile()

On Error GoTo Quotezerr
' Not good as we miss the .close

Open Genopts.Cdlg.FileName For Output As #hfile
Close #hfile

hfile = FreeFile()

Open Genopts.Cdlg.FileName For Output As #hfile
.MoveFirst

For ct = 0 To gt(191) - 2
' If !bmpindex > 40 And !bmpindex < 54 Then
' If ![Bmpindex] = 81 Then
' If Asc(Left$(![Quotestr], 1)) > 96 Then
Print #hfile, ![Quotestr]
' End If
.MoveNext
Next
Print #hfile, ![Quotestr];

Close #hfile
End If

End Select


.Close
End With


Set rectemp = Nothing
Set dbsCurrent = Nothing
killdb sDatabaseName


Exit Sub
Quotezerr:
rectemp.Close
Set rectemp = Nothing
Set dbsCurrent = Nothing
ShowError
killdb sDatabaseName
End Sub
Public Function IZempty()

On Error GoTo Nullerror

If rectemp.Fields("Bmpfile").FieldSize > 0 Then
IZempty = False
Else
IZempty = True
End If
Exit Function

Nullerror:
Err.Clear
IZempty = True

End Function
Public Sub stripMBstats(STRIPRH As Boolean)
' Use this method to clear old/new MB strategies
For ct = 89 To 103
If ct - 88 <= gt(10) Then
If STRIPRH = True Then  ' clear new
        If gt(ct) > 30 Then
        If Mid(CStr(gt(ct)), 3, 1) = "0" Then
        gt(ct) = CLng(Left(CStr(gt(ct)), 2))
        ElseIf Mid(CStr(gt(ct)), 2, 1) = "0" Then
        gt(ct) = CLng(Left(CStr(gt(ct)), 1))
        End If
        End If
Else
        ct1 = CLng(Right(CStr(gt(ct)), 2))
        gt(ct) = ct1
        gt(88) = gt(88) + ct1 ' stats
        gt(129) = gt(129) + 1
End If
ElseIf ct - 88 > gt(151) Then
gt(ct) = 0
End If
Next
End Sub
Public Function Namevalid(testtype As Long, chartxt As String)
Namevalid = True
' testtype 0 Dirnamenopath,1 Dirnamewithpath, 2 valid Ascii
If testtype = 0 Then
For ct = 1 To Len(chartxt)
Select Case Asc(Mid(chartxt, ct, 1))
Case 32 To 46, 48 To 57, 59, 61, 64 To 91, 93 To 122, 125, 126
Case Else   ' 47 = /,58 = :,60 = <,62 = >,63 = ?, 92 = \,123 = {,124 = |,125= },126 = ~
ct1 = Asc(Mid(chartxt, ct, 1))
Namevalid = False
Exit Function
End Select
Next
ElseIf testtype = 1 Then
For ct = 1 To Len(chartxt)
Select Case Asc(Mid(chartxt, ct, 1))
Case 32 To 46, 48 To 58, 59, 61, 64 To 122, 125, 126
Case Else
Namevalid = False
chartxt = Asc(Mid(chartxt, ct, 1))
Exit Function
End Select
Next
Else    ' testype = 2
For ct = 1 To Len(chartxt)
Select Case Asc(Mid(chartxt, ct, 1))
' 133 =…, 146 (not supported)? ,156 =œ,161 = ¡,162 = ¢, 163 = £, 169 = ©, 174 = ® ,175 = ¯, 176 = °, 177 = ±, 178 =², 179 = ³, 180 = ´, 181 = µ, 185 =¹, 188 - 255 various
Case 32 To 123, 125, 126, 133, 146, 156, 161 To 163, 169, 174 To 181, 185, 188 To 255
Case Else
Namevalid = False
Exit Function
End Select
Next
End If
End Function
Public Function Chunker(Rekordset As Recordset, Torek As Boolean, Optional kt As Long = 0)
Dim NumBlocks As Integer, FileLength As Long, LeftOver As Long, filedata As String
Const blocksize = 32768
hfile = FreeFile()
On Error GoTo Chunkerr
With Rekordset
If Torek = True Then
    ' Open source file.
    Open loaddirectory & "q" & kt & ".bmp" For Binary Access Read As hfile

    ' Get file length.
    FileLength = LOF(hfile)


    If FileLength = 0 Then GoTo Chunkerr


    ' Calculate no. of blocks to read, & leftover bytes.
    NumBlocks = FileLength \ blocksize
    LeftOver = FileLength Mod blocksize
    ' Put first record in edit mode.


    ' Read leftover data, writing it to table.
    filedata = String$(LeftOver, 32)
    Get #hfile, , filedata
    .Edit
    ![Bmpfile].AppendChunk (filedata)


    ' Read remaining blocks of data, writing them to table.
    filedata = String$(blocksize, 32)
    For ct = 1 To NumBlocks
    Get #hfile, , filedata
    ![Bmpfile].AppendChunk (filedata)
    Next
    .Update
    Close #hfile
    ' Update record and terminate function.



    Else
        ' Get size of the field.




        FileLength = ![Bmpfile].FieldSize()


        If FileLength = 0 Then GoTo Chunkerr


        ' Calculate no. of blocks to write, & leftover bytes.
        NumBlocks = FileLength \ blocksize
        LeftOver = FileLength Mod blocksize

        hfile = FreeFile()
        ' Remove any existing destination file.
        Open loaddirectory & "q" & kt & ".bmp" For Binary As #hfile
        Close #hfile

        hfile = FreeFile()
        Open loaddirectory & "q" & kt & ".bmp" For Binary As #hfile


        ' Write leftover data to output file.
        filedata = ![Bmpfile].GetChunk(0, LeftOver)
        Put #hfile, , filedata


        ' Write remaining blocks of data to output file.
        For ct1 = 1 To NumBlocks ' Reads a chunk and writes it to output file.
        filedata = ![Bmpfile].GetChunk((ct1 - 1) * blocksize + LeftOver, blocksize)
        Put #hfile, , filedata
        Next
        Close #hfile

End If

End With
Chunker = True
Exit Function
Chunkerr:
Chunker = False
ShowError
Rekordset.Close
End Function
