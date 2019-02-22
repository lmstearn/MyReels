VERSION 5.00
Begin VB.Form cfgthumb 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MyReels: Configure new Pictures"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   7
   Icon            =   "cfgthumb.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   Begin VB.CheckBox Shortcutter 
      Caption         =   "Create Shortcut"
      Height          =   210
      Left            =   7680
      TabIndex        =   15
      ToolTipText     =   "Create a shortcut on Programs\MyReels"
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CheckBox chkalwayssquare 
      Caption         =   "Square crops"
      Height          =   210
      Left            =   7680
      TabIndex        =   14
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Timer Drawlines 
      Enabled         =   0   'False
      Interval        =   18
      Left            =   1560
      Top             =   7080
   End
   Begin VB.PictureBox rub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "cfgthumb.frx":000C
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picTruesize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "cfgthumb.frx":0316
      ForeColor       =   &H80000008&
      Height          =   1400
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6375
      Left            =   9120
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   180
      Left            =   2520
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   6600
   End
   Begin VB.CommandButton Quit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Savenew 
      Caption         =   "&Open/Create a new Save Directory"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CommandButton Savepictures 
      Caption         =   " "
      Default         =   -1  'True
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   7920
      Width           =   5415
   End
   Begin VB.ComboBox PatternCombo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.DriveListBox drvsource 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.DirListBox dirsource 
      Appearance      =   0  'Flat
      Height          =   1290
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.FileListBox filesource 
      Height          =   3870
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Nextscreen 
      Caption         =   "&Next"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   7320
      Width           =   1455
   End
   Begin VB.PictureBox Prevue 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DragIcon        =   "cfgthumb.frx":0620
      DragMode        =   1  'Automatic
      Height          =   6720
      Left            =   2160
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   7
      Top             =   240
      Width           =   6720
      Begin VB.PictureBox Prevuechild 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         DragIcon        =   "cfgthumb.frx":076A
         DragMode        =   1  'Automatic
         Height          =   6720
         Left            =   0
         ScaleHeight     =   448
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   448
         TabIndex        =   11
         ToolTipText     =   "Right - click and drag mouse to crop region"
         Top             =   0
         Width           =   6720
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   72
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   1
            Visible         =   0   'False
            X1              =   96
            X2              =   168
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   2
            Visible         =   0   'False
            X1              =   96
            X2              =   168
            Y1              =   8
            Y2              =   8
         End
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   3
            Visible         =   0   'False
            X1              =   0
            X2              =   72
            Y1              =   8
            Y2              =   8
         End
      End
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   13
      Left            =   10800
      Top             =   7320
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   12
      Left            =   10800
      Top             =   6120
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   11
      Left            =   10800
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   10
      Left            =   10800
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   9
      Left            =   10800
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   8
      Left            =   10800
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   7
      Left            =   10800
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   6
      Left            =   9360
      Top             =   7320
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   5
      Left            =   9360
      Top             =   6120
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   4
      Left            =   9360
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   3
      Left            =   9360
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   2
      Left            =   9360
      Top             =   2520
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   1
      Left            =   9360
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image imgsel 
      Height          =   975
      Index           =   0
      Left            =   9360
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "cfgthumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_PATH = 260


'The following BFFM_ constants for reference. Others are actually used in auxrouts1
'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

'specific to the STRING method
Const LMEM_FIXED = &H0
Const LMEM_ZEROINIT = &H40
Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)



'Shortcutting
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const NOERROR = 0
Const FO_COPY = &H2
Const FOF_SILENT = &H4
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10
Const FOF_NOCONFIRMMKDIR = &H200

'Recent declares
Const SHARD_PATH = &H2&

Private Type SHFILEOPSTRUCT
hWnd      As Long
wFunc      As Long
pFrom      As String
pTo        As String
fFlags     As Integer
fAborted   As Boolean
hNameMaps  As Long
sProgress  As String
End Type

Enum STGM
    STGM_DIRECT = &H0&
    STGM_TRANSACTED = &H10000
    STGM_SIMPLE = &H8000000
    STGM_READ = &H0&
    STGM_WRITE = &H1&
    STGM_READWRITE = &H2&
    STGM_SHARE_DENY_NONE = &H40&
    STGM_SHARE_DENY_READ = &H30&
    STGM_SHARE_DENY_WRITE = &H20&
    STGM_SHARE_EXCLUSIVE = &H10&
    STGM_PRIORITY = &H40000
    STGM_DELETEONRELEASE = &H4000000
    STGM_CREATE = &H1000&
    STGM_CONVERT = &H20000
    STGM_FAILIFTHERE = &H0&
    STGM_NOSCRATCH = &H100000
End Enum
'
' Shell Folder Path Constants...
'
' on NT:
'   ..\WinNT\profiles\username
'
' on Windows 9x:
'   ..\Windows
Enum SHELLFOLDERS
    CSIDL_DESKTOP = &H0&            ' \Desktop
    CSIDL_PROGRAMS = &H2&           ' \Start Menu\Programs
    CSIDL_CONTROLS = &H3&           ' No Path
    CSIDL_PRINTERS = &H4&           ' No Path
    CSIDL_PERSONAL = &H5&           ' \Personal
    CSIDL_FAVORITES = &H6&          ' \Favorites
    CSIDL_STARTUP = &H7&            ' \Start Menu\Programs\Startup
    CSIDL_RECENT = &H8&             ' \Recent
    CSIDL_SENDTO = &H9&             ' \SendTo
    CSIDL_BITBUCKET = &HA&          ' No Path
    CSIDL_STARTMENU = &HB&          ' \Start Menu
    CSIDL_DESKTOPDIRECTORY = &H10&  ' \Desktop
    CSIDL_DRIVES = &H11&            ' No Path
    CSIDL_NETWORK = &H12&           ' No Path
    CSIDL_NETHOOD = &H13&           ' \NetHood
    CSIDL_FONTS = &H14&             ' \fonts
    CSIDL_TEMPLATES = &H15&         ' \ShellNew
    CSIDL_COMMON_STARTMENU = &H16&  ' ..\WinNT\profiles\All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17&   ' ..\WinNT\profiles\All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18&    ' ..\WinNT\profiles\All Users\Start Menu\Programs\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19& '..\WinNT\profiles\All Users\Desktop
    CSIDL_APPDATA = &H1A&           ' ..\WinNT\profiles\username\Application Data
    CSIDL_PRINTHOOD = &H1B&         ' ..\WinNT\profiles\username\PrintHood
End Enum

Enum SHOWCMDFLAGS
    SHOWNORMAL = 5
    SHOWMAXIMIZE = 3
    SHOWMINIMIZE = 7
End Enum

Private Type BrowseInfo    'bi
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As String   'return display name of item selected
  lpszTitle As String        'text to go in the banner over the tree
  ulFlags As Long            'flags that control the return stuff
  lpfnCallback As Long
  lParam As Long             'extra info passed back in callbacks
  iImage As Long             'output var: where to return the Image index
End Type

Const BIF_RETURNONLYFSDIRS = &H1&, BIF_NEWDIALOGSTYLE = &H40&, BIF_NONEWFOLDERBUTTON = &H200&


'Following for SHChangeNotify
Enum SHCN_Flags
  SHCNF_IDLIST = &H0      ' LPITEMIDLIST
  SHCNF_PATHA = &H1       ' path name
  SHCNF_PRINTERA = &H2    ' printer friendly name
  SHCNF_DWORD = &H3       ' DWORD
  SHCNF_PATHW = &H5       ' path name
  SHCNF_PRINTERW = &H6    ' printer friendly name
  SHCNF_TYPE = &HFF
  ' Flushes the system event buffer. The function does not return until the system is
  ' finished processing the given event.
  SHCNF_FLUSH = &H1000
  ' Flushes the system event buffer. The function returns immediately regardless of
  ' whether the system is finished processing the given event.
  SHCNF_FLUSHNOWAIT = &H2000
  SHCNE_UPDATEDIR = &H1000
  SHCNF_PATH = SHCNF_PATHW
  SHCNF_PRINTER = SHCNF_PRINTERW

End Enum

Const SHCNE_CREATE = &H2

Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal pwszPath As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)


'To notify Start menu of new shortcut
Private Declare Function SHChangeNotify Lib "shell32.dll" (ByVal wEventID As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long) As Long

'Reference only: specific to the STRING method
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long


Dim currentdirectory As String
Dim thumbs(14) As StdPicture, trueThumbs(14) As StdPicture
Dim firstchangeofdir As Boolean, blnsaveyet As Boolean, warning As Boolean, warning1 As Boolean, justclicked(14) As Boolean, writehallfame As Boolean
Dim inttracker(14) As Long, ct As Long, ct1 As Long, oldpicno As Long, response As Long, oldX As Long, oldY As Long, cropwidth As Long, cropheight As Long
Dim newX As Long, newY As Long, entryX As Long, entryY As Long
Dim Hscrollvis As Boolean, Vscrollvis As Boolean, scrolltolH As Long, scrolltolV As Long, rightbuttdown As Boolean, boosloterror As Boolean
Dim bordN As Boolean, bordE As Boolean, bordS As Boolean, bordW As Boolean, DoingDragDrop As Boolean, dirused As Boolean
Dim testtan As Single, tan1 As Single, tan3 As Single, tan5 As Single, tan7 As Single
Private Sub Form_Load()

procend = False
Dim smRes As Single
smRes = (resY - 1) / 2 + 1

If resX = 1 Then
setformpos Me
Else
Dotaskwindow Me, True
With Me

.Width = resX * .Width
.Height = resY * .Height

End With

setformpos Me, True

Prevue.Width = resX * Prevue.Width
Prevue.Height = resY * Prevue.Height
Prevuechild.Width = resX * Prevuechild.Width
Prevuechild.Height = resY * Prevuechild.Height
With Savepictures
.Left = resX * .Left
.Top = resY * .Top
.Width = smRes * .Width
.Height = smRes * .Height
.Fontsize = resY * Int(10 * textwidthratio)
End With
With Savenew
.Left = resX * .Left
.Top = resY * .Top
.Width = smRes * .Width
.Height = smRes * .Height
.Fontsize = resY * Int(10 * textwidthratio)
End With
With Nextscreen
.Left = resX * .Left
.Top = resY * .Top
.Width = smRes * .Width
.Height = smRes * .Height
.Fontsize = resY * Int(10 * textwidthratio)
End With
With Quit
.Left = resX * .Left
.Top = resY * .Top
.Width = smRes * .Width
.Height = smRes * .Height
.Fontsize = resY * Int(10 * textwidthratio)
End With
With chkalwayssquare
.Left = resX * .Left
.Top = resY * .Top
.Fontsize = resY * Int(8 * textwidthratio)
End With
With Shortcutter
.Left = resX * .Left
.Top = resY * .Top
.Fontsize = resY * Int(8 * textwidthratio)
End With
For ct = 0 To 13
With imgsel(ct)
.Top = resY * .Top
.Left = resX * .Left
.Height = resY * .Height
.Width = resX * .Width
End With
Next
End If


With PatternCombo
.AddItem "(*.bmp;*.ico;*.wmf;*.gif;*.jpg)"
.AddItem "Bitmap (*.bmp)"
.AddItem "Icon (*.ico)"
.AddItem "Metafile (*.wmf)"
.AddItem "Graphics Interchange (*.gif)"
.AddItem "Lossy format (*.jpg)"
.AddItem "All Files (*.*)"
.ListIndex = 0
End With

tan1 = Tan(3.1416 / 8)
tan3 = Tan(3 * 3.1416 / 8)
tan5 = Tan(5 * 3.1416 / 8)
tan7 = Tan(7 * 3.1416 / 8)

chkalwayssquare.Value = gt(44)
Shortcutter.Value = gt(190)

 
DoingDragDrop = False
rightbuttdown = False
currentdirectory = CurDir$
olddirectory = ""
oldpicno = 0
dirused = False
boosloterror = False

Savepictures.Caption = "&Allocate to " & CStr(currentdirectory)
Savepictures.ToolTipText = infofortt(0, drvsource.Drive)
Savenew.ToolTipText = infofortt(0, drvsource.Drive)

warning = True
warning1 = True
blnsaveyet = False
firstchangeofdir = True

For ct = 0 To 3
Clipborder(ct).Visible = False
Next

For ct = 1 To 14
inttracker(ct) = 0
justclicked(ct) = True
imgsel(ct - 1).BorderStyle = 1
Next
ct = 0
cfgthumb.Show

procend = True
End Sub
Private Sub Form_Resize()
Dim wid As Single, hgt As Single
cfgthumb.ScaleMode = 1
Const GAP = 60
cfgthumb.Caption = "Configure new Symbols"  'Space(109)


    'If WindowState = 1 Then Exit Sub

    wid = drvsource.Width
    drvsource.Move GAP, GAP, wid
    PatternCombo.Move GAP, (ScaleHeight - 100) - PatternCombo.Height - GAP, wid
    
    hgt = (PatternCombo.Top - drvsource.Top - drvsource.Height - GAP) / 2
    If hgt < 100 Then hgt = 100
    dirsource.Move GAP, drvsource.Top + drvsource.Height + GAP, wid, hgt
    filesource.Move GAP, dirsource.Top + dirsource.Height + GAP, wid, hgt
    
cfgthumb.ScaleMode = 3
Dotaskwindow Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Drawlines.Enabled = False
End Sub
Private Sub Prevuechild_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
rightbuttdown = True
entryX = 0
entryY = 0
Set picTruesize = Nothing
oldX = CLng(X)
oldY = CLng(Y)
newX = oldX
newY = oldY
Clipborder(0).X1 = oldX
Clipborder(0).X2 = oldX
Clipborder(0).Y1 = oldY
Clipborder(0).Y2 = newY
Clipborder(1).X1 = oldX
Clipborder(1).X2 = newX
Clipborder(1).Y1 = newY
Clipborder(1).Y2 = newY
Clipborder(2).X1 = newX
Clipborder(2).X2 = newX
Clipborder(2).Y1 = newY
Clipborder(2).Y2 = oldY
Clipborder(3).X1 = oldX
Clipborder(3).X2 = newX
Clipborder(3).Y1 = oldY
Clipborder(3).Y2 = oldY

For ct1 = 0 To 3
With Clipborder(ct1)
.BorderStyle = 2
.Visible = True
End With
Drawlines.Enabled = True
Next
End If
End Sub
Private Sub Drawlines_Timer()
Dim tolX As Long, tolY As Long
Clipborder(0).X1 = oldX
Clipborder(0).X2 = oldX
Clipborder(0).Y1 = oldY
Clipborder(0).Y2 = newY
Clipborder(1).X1 = oldX
Clipborder(1).X2 = newX
Clipborder(1).Y1 = newY
Clipborder(1).Y2 = newY
Clipborder(2).X1 = newX
Clipborder(2).X2 = newX
Clipborder(2).Y1 = newY
Clipborder(2).Y2 = oldY
Clipborder(3).X1 = oldX
Clipborder(3).X2 = newX
Clipborder(3).Y1 = oldY
Clipborder(3).Y2 = oldY
If Hscrollvis = True And Vscrollvis = True Then
If entryX <= 0 Or entryY <= 0 Or entryX > Prevue.Width + HScroll1.Max Or entryY > Prevue.Height + VScroll1.Max Then
entryX = 0
entryY = 0
Exit Sub
End If
ElseIf Hscrollvis = True Then
If entryX <= 0 Or entryY <= 0 Or entryX > Prevue.Width + HScroll1.Max Then
entryX = 0
entryY = 0
Exit Sub
End If
ElseIf Vscrollvis = True Then
If entryX <= 0 Or entryY <= 0 Or entryX > Prevue.Width + VScroll1.Max Then
entryX = 0
entryY = 0
Exit Sub
End If
Else
Exit Sub
End If


tolX = HScroll1.Value + scrolltolH / 2
tolY = VScroll1.Value + scrolltolV / 2


If newX = entryX Or newY = entryY Then
Exit Sub
Else
testtan = (newX - entryX) / (entryY - newY)
End If

bordN = False
bordE = False
bordS = False
bordW = False

'At top,left edge, note tan(A + PI) = Tan (A)
If newX < tolX Or newY < tolY Then
    If newX < tolX Then bordW = True
    If newY < tolY Then bordN = True
    'set angleofentry: tan decreases as angle increases, invert Y angles
    If testtan > tan3 Then  'big: going right right up or left left down
        If newX > tolX Then
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
        Else
        If Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
        End If
        If newY < tolY Then
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
        Else
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        End If
        End If
    ElseIf testtan <= tan5 Then 'big - : going right right down or left left up
        If Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        ElseIf Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
    ElseIf testtan <= tan3 And testtan > tan1 Then  'up & right OR left & down
        If newX > tolX Then
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
        Else
        If Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
        End If
        If newY < tolY Then
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
        Else
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        End If
        End If
    ElseIf testtan <= 1 And testtan > tan7 Then 'going up (left)
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        ElseIf Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    ElseIf testtan <= tan7 And testtan > tan5 Then  'left & up
        If Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
    Else    'Going down (left)
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        ElseIf Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
    End If
Exit Sub
End If


tolX = HScroll1.Value + VScroll1.Left - Prevue.Left - scrolltolH / 2
tolY = VScroll1.Value + HScroll1.Top - Prevue.Top - scrolltolV / 2



'at lower, right edge: scroll bars visible?
If newX > tolX Or newY > tolY Then
    If newX > tolX Then bordE = True
    If newY > tolY Then bordS = True
    'set angleofentry: tan decreases as angle increases, invert Y angles
    If testtan > tan3 Then  'big: going right right up or left left down
        If newX > tolX Then
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
        Else
        If Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
        End If
        If newY < tolY Then
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
        Else
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        End If
        End If
    ElseIf testtan <= tan5 Then 'big - : going right right down or left left up
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        ElseIf Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        End If
    ElseIf testtan <= tan3 And testtan > tan1 Then  'up & right OR down & left
        If newX > tolX Then
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
        Else
        If Hscrollvis = True Then
        If newX <= entryX And HScroll1.Value - HScroll1.SmallChange >= 0 Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
        End If
        If newY < tolY Then
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
        Else
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        End If
        End If
    ElseIf testtan <= tan1 And testtan >= 0 Then 'going up (right)
        If Vscrollvis = True Then
        If newY <= entryY And VScroll1.Value - VScroll1.SmallChange >= 0 Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
        End If
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    ElseIf testtan <= tan7 And testtan > tan5 Then  'left & up
        If Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        End If
    Else    'Going down (right)
        If Vscrollvis = True Then
        If newY >= entryY And VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
        ElseIf Hscrollvis = True Then
        If newX >= entryX And HScroll1.Value + HScroll1.SmallChange <= HScroll1.Max Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    End If
Exit Sub
End If

End Sub
Private Sub Prevuechild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then

'Points of entry to set angles
If X >= Prevue.Width + HScroll1.Value - scrolltolH Then
    If entryX = 0 Then
        If X >= newX Then
        entryX = X - 1
        Else
        entryX = X + 1
        End If
    End If
    If entryY = 0 Then
        If Y >= newY Then
        entryY = Y - 1
        Else
        entryY = Y + 1
        End If
    End If
newX = CLng(X)
newY = CLng(Y)
ElseIf Y >= Prevue.Height + VScroll1.Value - scrolltolV Then
    If entryX = 0 Then
        If X >= newX Then
        entryX = X - 1
        Else
        entryX = X + 1
        End If
    End If
    If entryY = 0 Then
        If Y >= newY Then
        entryY = Y - 1
        Else
        entryY = Y + 1
        End If
    End If
newX = CLng(X)
newY = CLng(Y)
ElseIf X <= scrolltolH + HScroll1.Value Then
    If entryX = 0 Then
        If X >= newX Then
            If X > 1 Then
            entryX = X - 1
            Else
            entryX = 1
            End If
        Else
        entryX = X + 1
        End If
    End If
    If entryY = 0 Then
        If Y >= newY Then
        entryY = Y - 1
        Else
        entryY = Y + 1
        End If
    End If
newX = CLng(X)
newY = CLng(Y)
ElseIf Y <= scrolltolV + VScroll1.Value Then
    If entryX = 0 Then
        If X >= newX Then
        entryX = X - 1
        Else
        entryX = X + 1
        End If
    End If
    If entryY = 0 Then
        If Y >= newY Then
            If Y > 1 Then
            entryY = Y - 1
            Else
            entryY = 1
            End If
        Else
        entryY = Y + 1
        End If
    End If
newX = CLng(X)
newY = CLng(Y)
Else
newX = CLng(X)
newY = CLng(Y)
entryX = 0
entryY = 0
End If

End If
End Sub
Private Sub Prevuechild_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
rightbuttdown = False
Drawlines.Enabled = False

newX = CLng(X)
newY = CLng(Y)

cropwidth = newX - oldX
cropheight = newY - oldY

With Prevuechild
If gt(44) = 1 Then
    If cropwidth > 0 Then
        If cropwidth > .Width - oldX Then
        cropwidth = .Width - oldX
        newX = .Width
        End If
    Else
        If cropwidth < -oldX Then
        cropwidth = -oldX
        newX = 0
        End If
    End If


    If cropheight > 0 Then
        If cropheight > .Height - oldY Then
        cropheight = .Height - oldY
        newY = .Height
        End If
    Else
        If cropheight < -oldY Then
        cropheight = -oldY
        newY = 0
        End If
    End If

If cropwidth > 0 And cropheight > 0 Then
    If cropwidth > cropheight Then
        newX = oldX + cropheight
        If newX > .Width Then
        newX = .Width
        cropheight = newX - oldX
        End If
        cropwidth = cropheight
    ElseIf cropwidth < cropheight Then
        newY = oldY + cropwidth
        If newY > .Height Then
        newY = .Height
        cropwidth = newY - oldY
        End If
        cropheight = cropwidth
    End If
ElseIf cropwidth > 0 And cropheight < 0 Then
    If cropwidth > -cropheight Then
        newX = oldX - cropheight
        If newX > .Width Then
        newX = .Width
        cropheight = oldX - newX
        End If
        cropwidth = -cropheight
    ElseIf cropwidth < -cropheight Then
        newY = oldY - cropwidth
        If newY < 0 Then
        newY = 0
        cropwidth = oldY
        End If
        cropheight = -cropwidth
    End If
ElseIf cropwidth < 0 And cropheight > 0 Then
    If -cropwidth > cropheight Then
        newX = oldX - cropheight
        If newX < 0 Then
        newX = 0
        cropheight = -oldX
        End If
        cropwidth = -cropheight
    ElseIf -cropwidth < cropheight Then
        newY = oldY - cropwidth
        If newY > .Height Then
        newY = .Height
        cropwidth = newY - oldY
        End If
        cropheight = -cropwidth
    End If
Else
    If -cropwidth > -cropheight Then
        newX = oldX + cropheight
        If newX < 0 Then
        newX = 0
        cropheight = -oldX
        End If
        cropwidth = cropheight
    ElseIf -cropwidth < -cropheight Then
        newY = oldY + cropwidth
        If newY < 0 Then
        newY = 0
        cropwidth = -oldY
        End If
        cropheight = cropwidth
    End If
End If
Else    'gt(44)
        If newX > .Width Then
        newX = .Width
        cropwidth = newX - oldX
        ElseIf newX < 0 Then
        newX = 0
        cropwidth = -oldX
        End If
        If newY > .Height Then
        newY = .Height
        cropheight = newY - oldY
        ElseIf newY < 0 Then
        newY = 0
        cropheight = -oldY
        End If
End If
End With

If cropwidth = 0 Or cropheight = 0 Then
oldX = 0
oldY = 0
Exit Sub
End If

Clipborder(0).X1 = oldX
Clipborder(0).X2 = oldX
Clipborder(0).Y1 = oldY
Clipborder(0).Y2 = newY
Clipborder(1).X1 = oldX
Clipborder(1).X2 = newX
Clipborder(1).Y1 = newY
Clipborder(1).Y2 = newY
Clipborder(2).X1 = newX
Clipborder(2).X2 = newX
Clipborder(2).Y1 = newY
Clipborder(2).Y2 = oldY
Clipborder(3).X1 = oldX
Clipborder(3).X2 = newX
Clipborder(3).Y1 = oldY
Clipborder(3).Y2 = oldY


For ct1 = 0 To 3
With Clipborder(ct1)
.BorderStyle = 1
.Visible = True
End With
Next

Set rub = Nothing
On Error GoTo badpic
    If cropwidth > 0 And cropheight > 0 Then
    With rub
    .Width = cropwidth
    .Height = cropheight
    .PaintPicture Prevuechild.Picture, 0, 0, .Width, .Height, oldX, oldY, cropwidth, cropheight
    End With
    ElseIf cropwidth > 0 And cropheight < 0 Then
    With rub
    .Width = cropwidth
    .Height = -cropheight
    .PaintPicture Prevuechild.Picture, 0, 0, cropwidth, -cropheight, oldX, newY, cropwidth, -cropheight
    End With
    ElseIf cropwidth < 0 And cropheight > 0 Then
    With rub
    .Width = -cropwidth
    .Height = cropheight
    .PaintPicture Prevuechild.Picture, 0, 0, -cropwidth, cropheight, newX, oldY, -cropwidth, cropheight
    End With
    Else
    With rub
    .Width = -cropwidth
    .Height = -cropheight
    .PaintPicture Prevuechild.Picture, 0, 0, -cropwidth, -cropheight, newX, newY, -cropwidth, -cropheight
    End With
    End If
    
Set rub.Picture = rub.Image
End If
Exit Sub
badpic:
ShowError
End Sub
Private Sub chkalwayssquare_Click()
If procend = False Then Exit Sub
If chkalwayssquare.Value = 0 Then
gt(44) = 0
Else
gt(44) = 1
End If
End Sub
Private Sub Quit_Click()
If blnsaveyet = False And ct > 0 Then
response = MsgBox("Discard the selected images?", vbYesNo)
If response = vbYes Then
Unload Me
Set cfgthumb = Nothing
Load gametype
gametype.Show
End If
Else
Unload Me
Set cfgthumb = Nothing
Load gametype
gametype.Show
End If
End Sub
Private Function BrowseForFolderByPath(sSelPath As String) As String

Dim BI As BrowseInfo
Dim pidl1 As Long, pidl2 As Long
Dim sPath As String * MAX_PATH 'SysReAllocStringLen didn't work

With BI

    .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_NONEWFOLDERBUTTON

   'owner of the dialog. Pass 0 for the desktop.
    If Not Screen.ActiveForm Is Nothing Then
    .hWndOwner = Screen.ActiveForm.hWnd
    Else
    .hWndOwner = 0
    End If
   
   'The desktop folder will be the dialog's root folder.
   'SHSimpleIDListFromPath can also be used to set this value.
    .pIDLRoot = 0
    
   'Set the dialog's prompt string
    .lpszTitle = "Pre-selecting the folder using the folder's name. No backout after 'ok'"
    
    .pszDisplayName = ""
    
   'Obtain and set the address of the callback function
    .lpfnCallback = FARPROC(AddressOf BrowseCallbackProcStr)

     pidl1 = ILCreateFromPathW(StrPtr(sSelPath))
     .lParam = pidl1
    
End With

  'Shows the browse dialog and doesn't return until the
  'dialog is closed. The BrowseCallbackProcStr will
  'receive all browse dialog specific messages while
  'the dialog is open. pidl will contain the pidl of the
  'selected folder if the dialog is not canceled: (re-using pidl)
   pidl2 = SHBrowseForFolder(BI)



If pidl2 Then
   
     'Get the path from the selected folder's pidl returned
     'from the SHBrowseForFolder call (rtns True on success,
     'sPath must be pre-allocated!)

      If SHGetPathFromIDList(pidl2, sPath) Then
        'SysReAllocStringLen VarPtr(sPath), , MAX_PATH
        'Return the path
         BrowseForFolderByPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
         
      End If
      
     'Free the memory the shell allocated for the pidl.
      Call CoTaskMemFree(pidl2)
   
End If
'If lpSelPath <> 0 Then Call LocalFree(lpSelPath)
Call CoTaskMemFree(pidl1)
End Function
Private Sub Savenew_Click()
Dim snewdir As String, sBuffer As String

cfgthumb.Enabled = False
sBuffer = Left(loaddirectory, Len(loaddirectory) - 1)
sBuffer = BrowseForFolderByPath(sBuffer)
cfgthumb.Enabled = True
If sBuffer = "" Then Exit Sub

dirused = False
oldpicno = 0
snewdir = InputBox("Please type in the name of New Directory under: " & sBuffer & " & click OK, or click Cancel to choose: " & sBuffer, "New Directory Name", "Newdir")

'snewdir = "" if user cancels, so sbuffer is currentdirectory
If Namevalid(0, snewdir) = True Then
    If snewdir = "" Then
    currentdirectory = sBuffer
    ChDrive (Left$(currentdirectory, 1))
    ChDir currentdirectory
        If findafile(currentdirectory, "Slotdata.s$t") > 0 Then
        dirused = True
        For ct1 = 1 To 9
        If findafile(currentdirectory, ct1 & ".bmp") > 0 Then oldpicno = oldpicno + 1
        Next
        For ct1 = 0 To 4
        If findafile(currentdirectory, "1" & ct1 & ".bmp") > 0 Then oldpicno = oldpicno + 1
        Next
        End If
    
    Else
    currentdirectory = sBuffer & "\" & snewdir
    On Error GoTo cantcreatedir
    ChDrive (Left$(currentdirectory, 1))
    MkDir currentdirectory
    ChDir currentdirectory
    warning = False
    End If
    
    Nextscreen.Caption = "&Next"
    Savepictures.Caption = "&Allocate to " & CStr(currentdirectory)
    Savepictures.ToolTipText = infofortt(ct - oldpicno, drvsource.Drive)
    Savenew.ToolTipText = infofortt(ct, drvsource.Drive)
    'Cleanup old directory
    If Dir(loaddirectory & "q0" & ".bmp") <> "" Then Kill (loaddirectory & "q0" & ".bmp")
        If firstchangeofdir = True Then 'write data to out once only
        
            If Anychanges = True Then
            response = MsgBox("As there were changes to your old configuration, your VOG has been changed to 999.99. Yes to recalculate next time, or No to discard changes", vbYesNo)
                If response = vbYes Then
                gt(29) = 999
                gt(30) = 99
                gt(31) = 999
                gt(32) = 99
                gt(33) = 0
                gt(34) = 0
                ChDrive (Left$(loaddirectory, 1))
                ChDir loaddirectory
                If outputvars = False Then MsgBox "Unable to write to old slotdata.s$t.", vbOKOnly
                ChDrive (Left$(currentdirectory, 1))
                ChDir currentdirectory
                End If
            End If
    
        firstchangeofdir = False
        End If
    
    gt(0) = 5 'No going back
    Quit.Enabled = False
    writehallfame = False
    blnsaveyet = False
   
Else
MsgBox "Invalid Characters, please retry", vbOKOnly
End If  'namevalid
Exit Sub
cantcreatedir:
MsgBox "Directory already exists", vbOKOnly
End Sub
Private Sub Savepictures_Click()
Dim oldgt(192) As Long, oldvars(8) As String, oldvog  As Single

boosloterror = False 'In case of retry file delete
On Error GoTo Criterror 'Takes care of copy errors (in errorhandlers as well)

If ct < 3 Then
response = MsgBox("Need three images or more!", vbOKOnly)
Exit Sub
End If

'Bug in 2000 NT?, curdir truncates path names ~1 etc

If warning = True Then
    
    If InStr(currentdirectory, "~") > 0 Then
    If Left(Right(currentdirectory, 8), 6) = Left(Right(App.Path, 8), 6) Then
    response = MsgBox("Your Install path name contains a ""~"" which usually MSDOS truncation. Unable to tell from the new path name whether or not you are changing the generic default Configuration. If that is the case then it is not recommended that this be changed. Are you sure you want to continue?", vbYesNo)
        If response = vbNo Then
        warning1 = False
        Exit Sub
        End If
    End If
    ElseIf currentdirectory = App.Path Then
    response = MsgBox("It is not recommended that the generic default configuration is changed. Are you sure you want to continue?", vbYesNo)
        If response = vbNo Then
        warning1 = False
        Exit Sub
        End If
    End If
End If

If findafile(currentdirectory, "Slotdata.s$t") > 0 Then dirused = True

    If dirused = True Then 'First time here
        If writehallfame = True Then
        response = MsgBox("Any bitmap files previously used by the game in the directory name on the Allocate Button will be killed. Clicking Yes will also commit to new reel distributions", vbYesNo)
        Else
        response = MsgBox("Any bitmap files previously used by the game in the directory name on the Allocate Button will be killed. Clicking Yes will also commit to new reel distributions and current game data in that directory to be written to the Hall_Of_Fame", vbYesNo)
        End If
        If response = vbNo Then
        warning1 = False
        Exit Sub
        End If
    End If

    KillExistingPics 14

    CopypicstoDir

    gt(0) = 5 'No going back
    Quit.Enabled = False
    oldpicno = 0
    warning1 = True
    filesource.Refresh
    blnsaveyet = True

        If writehallfame = True Then  'don't need to write to hallfame twice!
        If currentdirectory <> loaddirectory Then FileCopy loaddirectory & "Slotdata.s$t", currentdirectory & "Slotdata.s$t"
        GoTo Writesuccess
        End If


    On Error GoTo Sloterror
        If findafile(currentdirectory, "Slotdata.s$t") > 0 Then

        'Write to Hallfame name found in new Slotdata


        sDatabaseName = currentdirectory & "Slotdata.s$t"


        If OpenDb(sDatabaseName, 0) = False Then
        MsgBox "Slotdata.s$t in new directory has been corrupted, data cannot be written to Hall_Of_Fame.", vbOKOnly
        Else
        Set dbsCurrent = gdbCurrentDB


        'REORG  Slotdata
        REORGSLOT
        Set rectemp = dbsCurrent.OpenRecordset("SELECT Lorng FROM Inpoot")

            With rectemp
            On Error GoTo ErrHandlerr

            .MoveFirst

            '345 moves before gt, 116 after
            For ct1 = 1 To 345
            .MoveNext
            Next

            For ct1 = 1 To 192
            Select Case ct1
            Case 29
            oldgt(29) = ![Lorng]
                If oldgt(29) < 0 Then
                VOGchg = 1
                oldgt(29) = -oldgt(29)
                Else
                VOGchg = 0
                End If
            Case 2, 10, 26 To 28, 31, 49 To 87, 133, 150 To 152, 155, 156, 192
            oldgt(ct1) = ![Lorng]
                If oldgt(ct1) < 0 Then
                oldgt(ct1) = 0
                ElseIf oldgt(ct1) > 2147483647 Then oldgt(ct1) = 2147483647 'necessary?
                End If
            Case 30, 32
            oldgt(ct1) = ![Lorng]
                If oldgt(ct1) < 0 Then
                oldgt(ct1) = 0
                ElseIf oldgt(ct1) > 9999 Then oldgt(ct1) = 9999
                End If

            End Select
            .MoveNext
            Next

                If oldgt(10) > 0 Then
                gt(31) = oldgt(31)
                gt(32) = oldgt(32)
                oldvog = CSng(inputformat(31))
                Else
                gt(29) = oldgt(29)
                gt(30) = oldgt(30)
                oldvog = CSng(inputformat(29))
                End If
            End With



        Set rectemp = dbsCurrent.OpenRecordset("SELECT Streeng FROM Inpoot")
            With rectemp
            On Error GoTo ErrHandlerr

            .MoveFirst

            For ct1 = 1 To 8
            Select Case ct1
            Case 4
            'Hallfame Directory
            oldvars(4) = ![Streeng]
            If Len(oldvars(4)) > 100 Or Namevalid(1, oldvars(4)) = False Then
            response = MsgBox("The Hall_Of_Fame name in Slotdata.s$t in the new directory is invalid. Would you like to try the current name?", vbYesNo)
            If response = True Then
            oldvars(4) = Stringvars(4)
            Else    'Exit Hallfame write
            GoTo Writesuccess
            End If
            End If
            Case 5, 8
            oldvars(ct1) = ![Streeng]
            If Len(oldvars(ct1)) > 20 Then
            oldvars(ct1) = Left(oldvars(ct1), 20)
            ElseIf Namevalid(2, oldvars(ct1)) = False Then
            oldvars(ct1) = "Invalid_Name"
            End If

            End Select
            .MoveNext
            Next



            End With
        Set gdbCurrentDB = dbsCurrent
        killdb sDatabaseName
        Set rectemp = Nothing
        Set dbsCurrent = Nothing




        'Write to Hallfame

        sDatabaseName = oldvars(4) & "Hallfame.s$t"

        If OpenDb(sDatabaseName, 1) = False Then
        MsgBox "Hallfame.s$t missing or corrupted, data cannot be written there ; continuing ...", vbOKOnly
        GoTo Writesuccess
        End If

        Set dbsCurrent = gdbCurrentDB

        Set rectemp = dbsCurrent.OpenRecordset("Hall")

        With rectemp

        .MoveLast

        .AddNew

        On Error GoTo Writesuccess

        ![Player] = oldvars(8)

        ![Game] = oldvars(5)

        ct1 = CLng(Left(CStr(oldgt(155)), 3))
        If ct1 <= 500 Then
        ![Startcash] = ct1
        Else
        ![Startcash] = 50
        End If

        ![Endcash] = oldgt(2)

        If oldgt(156) = 1 And oldgt(192) > 1 Then
        ![Reason] = 2
        ElseIf oldgt(2) = 0 Then
        ![Reason] = 1
        Else
        ![Reason] = 3
        End If

        ![Date] = CDate(Date)

        ![Spintotal] = oldgt(49) + oldgt(50) + oldgt(51)

        ![ReelChange] = oldgt(26)

        ![SymbolChange] = oldgt(27)

        ![GeneralChange] = oldgt(28)

        ![StatsResets] = oldgt(150)


        If oldgt(55) < 0 Then oldgt(55) = 0
        ![GameReplays] = oldgt(55)


        If oldgt(133) > 0 Then
        ![SVOG] = ((oldgt(59) + oldgt(63) + oldgt(67) + oldgt(71) + oldgt(75) + oldgt(79) + oldgt(83) + oldgt(87)) / oldgt(133)) * 100
        Else
        ![SVOG] = 0
        End If

        If VOGchg = 1 Then
        ![VOG] = -oldvog
        Else
        ![VOG] = oldvog
        End If

        If oldgt(152) < 0 Then oldgt(152) = -oldgt(152)
        ![MonteCarlo] = oldgt(152)

        .Update


        .Close
        End With

        Set rectemp = Nothing
        Set dbsCurrent = Nothing
        killdb sDatabaseName
        writehallfame = True
        If compactdb(1) = False Then GoTo Writesuccess

        End If  'Slotdata open error



    Else
    'No Slotdata here yet
    FileCopy loaddirectory & "Slotdata.s$t", currentdirectory & "Slotdata.s$t"
    End If

If loaddirectory <> currentdirectory And olddirectory = "" Then olddirectory = loaddirectory
loaddirectory = currentdirectory


Writesuccess:
Savepictures.ToolTipText = infofortt(ct, drvsource.Drive)
Savenew.ToolTipText = "The config in the old directory might be a problem if you continue here. See help"

Exit Sub


Criterror:
boosloterror = True
ShowError
MsgBox "Please check file permissions in the selected directory and retry!", vbOKOnly
Exit Sub
Sloterror:
boosloterror = True
ShowError
MsgBox "Slotdata.s$t has been deleted, program halted!!!", vbOKOnly
Exit Sub
ErrHandlerr:
Set rectemp = Nothing
Set dbsCurrent = Nothing
ShowError
killdb sDatabaseName
MsgBox "Slotdata.s$t in new directory has been corrupted, data cannot be written to Hall_Of_Fame ; continuing ...", vbOKOnly
'Delete old Slotdata
Kill currentdirectory & "Slotdata.s$t"
FileCopy loaddirectory & "Slotdata.s$t", currentdirectory & "Slotdata.s$t"
If loaddirectory <> currentdirectory And olddirectory = "" Then olddirectory = loaddirectory
loaddirectory = currentdirectory
End Sub
Private Sub Nextscreen_Click()

Select Case ct
Case 0 To 2
response = MsgBox("Need three images or more!", vbOKOnly)
Exit Sub
Case 3 To 14
    If blnsaveyet = False Then
    
        If Nextscreen.Caption = "&Sure?" Then
        Savepictures_Click
        Nextscreen.Caption = "&Next"
        If warning1 = False Then Exit Sub
            If boosloterror = True Then
            Unload cfgthumb
            Set cfgthumb = Nothing
            Exit Sub
            End If
        Else
        Nextscreen.Caption = "&Sure?"
        Exit Sub
        End If
    
    End If
End Select
CopypicstoDir True
If gt(190) = 1 Then Shortcutting
FinishPictureConfig trueThumbs, ct
Unload Me
Set cfgthumb = Nothing
Load cornfig
End Sub
Private Sub CopypicstoDir(Optional ReassignThumbs = False)
Dim ct2 As Long, thumbTmp(14) As StdPicture
ct2 = 0

For ct1 = 1 To 14
If inttracker(ct1) <> 0 Then
ct2 = ct2 + 1
Set thumbTmp(ct2) = thumbs(ct1)
End If
Next

KillExistingPics ct2

For ct1 = 1 To ct2 'ct is equivalent
If ReassignThumbs Then Set trueThumbs(ct1) = thumbTmp(ct1)
SavePicture thumbTmp(ct1), CStr(ct1) & ".bmp"
Next
End Sub
Private Sub KillExistingPics(ByVal picMax As Integer)
If picMax < 10 Then
    'Kill existing files
    For ct1 = 1 To picMax
    If findafile(currentdirectory, ct1 & ".bmp") > 0 Then Kill ct1 & ".bmp"
    Next
    Else
    For ct1 = 1 To 9
    If findafile(currentdirectory, ct1 & ".bmp") > 0 Then Kill ct1 & ".bmp"
    Next
    For ct1 = 0 To picMax - 10
    If findafile(currentdirectory, "1" & ct1 & ".bmp") > 0 Then Kill "1" & ct1 & ".bmp"
    Next
    End If
End Sub
Private Sub imgsel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
If DoingDragDrop = True Then Exit Sub
Nextscreen.Caption = "&Next"
DoDragdrop CLng(Index + 1)
End Sub
Private Sub imgsel_Click(Index As Integer)
If justclicked(Index + 1) = True Then Exit Sub
Nextscreen.Caption = "&Next"
If inttracker(Index + 1) > 0 Then
Reduceint Index + 1, inttracker
inttracker(Index + 1) = 0
Set imgsel(Index).Picture = LoadPicture()
Set thumbs(Index + 1) = Nothing
If ct > 0 Then ct = ct - 1
End If
justclicked(Index + 1) = True
Savepictures.ToolTipText = infofortt(ct - oldpicno, drvsource.Drive)
Savenew.ToolTipText = infofortt(ct, drvsource.Drive)
End Sub
Private Sub dirsource_Change()
    filesource.Path = dirsource.Path
End Sub
Private Sub drvsource_Change()
    On Error GoTo DriveError
    dirsource.Path = drvsource.Drive
    filesource.Path = ""
    dirsource.Refresh
    Exit Sub

DriveError:
    drvsource.Drive = dirsource.Path
End Sub
Private Sub filesource_Click()
Dim fName As String
cropwidth = 0
cropheight = 0
scrolltolH = 0
scrolltolV = 0
HScroll1.Value = 0
VScroll1.Value = 0
fName = ""
For ct1 = 0 To 3
With Clipborder(ct1)
.BorderStyle = 2
.Visible = False
End With
Next

    
    On Error GoTo LoadPictureError
    'fname = Left(CStr(drvsource.Drive), 2) not needed
    
    If Mid(filesource.Path, Len(filesource.Path), 1) = "\" Then
    fName = fName + filesource.Path + filesource.FileName
    Else
    fName = fName + filesource.Path + "\" + filesource.FileName
    End If
    
    Prevuechild.Picture = LoadPicture(fName)
    
    
    Prevue.Move resX * 144, resY * 16, resX * 448, resY * 448
    Prevuechild.Move 0, 0
    
    
    HScroll1.Left = resX * 144
    If Prevuechild.Height < resY * 448 Then
    HScroll1.Top = resY * 16 + Prevuechild.Height
    VScroll1.Height = Prevuechild.Height
    Prevue.Height = Prevuechild.Height
    Else
    HScroll1.Top = resY * (16 + 448)
    VScroll1.Height = resY * 448
    End If
    
    VScroll1.Top = resY * 16
    If Prevuechild.Width < resX * 448 Then
    VScroll1.Left = resX * 144 + Prevuechild.Width
    HScroll1.Width = Prevuechild.Width
    Prevue.Width = Prevuechild.Width
    Else
    VScroll1.Left = resX * (144 + 448)
    HScroll1.Width = resX * 448
    End If
    
    
    'make .max multiple of 16
    HScroll1.Max = findhcmplusone(CInt(Prevuechild.Width - resX * 448), CInt(resX * 16))
    VScroll1.Max = findhcmplusone(CInt(Prevuechild.Height - resY * 448), CInt(resY * 16))
    
    Hscrollvis = Prevue.Width < Prevuechild.Width
    Vscrollvis = Prevue.Height < Prevuechild.Height
    HScroll1.Visible = Hscrollvis
    VScroll1.Visible = Vscrollvis
    'set arbitrary tolerance
    If Hscrollvis = True Then
    HScroll1.LargeChange = HScroll1.Max / 8
    HScroll1.SmallChange = HScroll1.Max / 16
    scrolltolH = 1 + Int(HScroll1.SmallChange)
    scrolltolV = scrolltolH
    End If
    If Vscrollvis = True Then
    VScroll1.LargeChange = VScroll1.Max / 8
    VScroll1.SmallChange = VScroll1.Max / 16
    scrolltolV = 1 + Int(VScroll1.SmallChange)
    If scrolltolH = 0 Then scrolltolH = scrolltolV
    End If
    
    Caption = "Viewer [" & fName & "]"
    Prevuechild.Visible = True
    Exit Sub

LoadPictureError:
    Beep
    Caption = "Viewer [Invalid picture]"
    Exit Sub
End Sub
Private Sub PatternCombo_Click()
Dim pat As String
Dim p1 As Long
Dim p2 As Long

    pat = PatternCombo.List(PatternCombo.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    filesource.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub
Private Sub HScroll1_Change()
Prevuechild.Left = -HScroll1.Value
If rightbuttdown = True Then
If entryX < Prevue.Width Then
If bordW = True Then entryX = HScroll1.Value + Int(scrolltolH / 2)
Else
If bordE = True Then entryX = Prevue.Width - Int(scrolltolH / 2) + HScroll1.Value
End If
End If
End Sub
Private Sub Shortcutter_Click()
If procend = False Then Exit Sub
If Shortcutter.Value = 0 Then
gt(190) = 0
Else
gt(190) = 1
End If
End Sub
Private Sub VScroll1_Change()
Prevuechild.Top = -VScroll1.Value
If rightbuttdown = True Then
If entryY < Prevue.Height Then
If bordN = True Then entryY = VScroll1.Value + Int(scrolltolV / 2)
Else
If bordS = True Then entryY = Prevue.Height - Int(scrolltolV / 2) + VScroll1.Value
End If
End If
End Sub
Private Sub DoDragdrop(whichindex As Long)
On Error GoTo Quit
Prevuechild.ToolTipText = ""
DoingDragDrop = True
Screen.MousePointer = vbHourglass
For ct1 = 0 To 3
Clipborder(ct1).Visible = True
Next
If inttracker(whichindex) = 0 Then
ct = ct + 1
inttracker(whichindex) = ct
End If
Set picTruesize = Nothing
ScaleMode = 1

With picTruesize
If cropwidth = 0 Then
.Width = 1820
.Height = 1820
.PaintPicture Prevuechild.Picture, 0, 0, 1820, 1820
Set thumbs(whichindex) = .Image
Set picTruesize = Nothing
.Width = resX * 915
.Height = resY * 915
.PaintPicture Prevuechild.Picture, 0, 0, resX * 915, resY * 915
Else
'scalemode twips
.Width = 1820
.Height = 1820
.PaintPicture rub.Picture, 0, 0, 1820, 1820
Set thumbs(whichindex) = .Image
Set picTruesize = Nothing
.Width = resX * 915
.Height = resY * 915
.PaintPicture rub.Picture, 0, 0, resX * 915, resY * 915
End If
imgsel(whichindex - 1).Picture = .Image
End With

oldX = 0
oldY = 0
cropwidth = 0
cropheight = 0

ScaleMode = 3
For ct1 = 0 To 3
With Clipborder(ct1)
.BorderStyle = 2
.Visible = False
End With
Next
justclicked(whichindex) = False
Screen.MousePointer = vbDefault
DoingDragDrop = False
Savepictures.ToolTipText = infofortt(ct - oldpicno, drvsource.Drive)
Savenew.ToolTipText = infofortt(ct, drvsource.Drive)
oldX = 0
oldY = 0
cropwidth = 0
cropheight = 0
Exit Sub

Quit:
oldX = 0
oldY = 0
cropwidth = 0
cropheight = 0
ScaleMode = 3
For ct1 = 0 To 3
With Clipborder(ct1)
.BorderStyle = 2
.Visible = False
End With
Next
justclicked(whichindex) = False
Screen.MousePointer = vbDefault
DoingDragDrop = False
ct = ct - 1
inttracker(whichindex) = 0
ShowError
End Sub
Private Sub Shortcutting()
Dim indexoficon As Long, startupmode As Long, sBuff As String, startMenuPath As String, StartMenuName As String, dirlen As Long

indexoficon = 0
startupmode = 0

'- We also need the user's Start Menu folder
startMenuPath = GetSpecialFolder(CSIDL_STARTMENU)

If startMenuPath = "" Then
MsgBox "Sorry, no StartMenuPath, no shortcut."
Exit Sub
End If
  
'set up the StartMenuPath now to reflect the folder
'we'll be installing the shortcuts into a bit later
startMenuPath = startMenuPath & "Programs\MyReels\"

'trim right string off currentdirectory
dirlen = Len(currentdirectory) - 1

For ct1 = dirlen To 1 Step -1
If Mid(currentdirectory, ct1, 1) = "\" Then Exit For
Next

StartMenuName = Mid(currentdirectory, ct1 + 1, dirlen - ct1) & ".lnk"

If findafile(startMenuPath, StartMenuName) > 0 Then
GetsBuff:
sBuff = InputBox("A shortcut with the name " & StartMenuName & " already exists. Please type in another name for it: " & " click OK, or click Cancel to overwrite the existing shortcut.", "New Shortcut Name", "New Shortcut Name")
  If sBuff = StartMenuName Then
  GoTo GetsBuff
  ElseIf sBuff <> "" Then
  StartMenuName = sBuff & ".lnk"
  End If
End If
StartMenuName = startMenuPath & StartMenuName
'fCreateShellLink StartMenuName, App.Path & "\MyReels.exe", currentdirectory, App.Path & "\MyReels.exe", indexoficon, startupmode
fCreateShellLink StartMenuName, currentdirectory & "\Slotdata.s$t", currentdirectory, "", indexoficon, startupmode


'SHChangeNotify SHCNE_CREATE, SHCNF_PATH, StrPtr(StartMenuName), 0
SHChangeNotify SHCNE_UPDATEDIR, SHCNF_PATH, StrPtr(StartMenuName), 0

End Sub
Private Sub ShellRenameFile(sOldName As String, sNewName As String)
  
  'set some working variables
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim r As Long
  
  'add a pair of terminating nulls to each string
   sOldName = sOldName & Chr$(0) & Chr$(0)
   sNewName = sNewName & Chr$(0) & Chr$(0)
  
  'for debugging - print the resulting strings
   Print sOldName
   Print sNewName
  
  'set up the options
   With SHFileOp
      .wFunc = FO_RENAME
      .pFrom = sOldName
      .pTo = sNewName
      .fFlags = FOF_SILENT Or FOF_NOCONFIRMATION
   End With
  
  'and rename the file
   r = SHFileOperation(SHFileOp)

End Sub
Private Function GetSpecialFolder(CSIDL As Long) As String
  
'a few local variables needed
Dim r As Long, sPath As String, pidl As Long
   
Const NOERROR = 0
Const MAX_LENGTH = 260
  
'fill pidl with the specified folder item
r = SHGetSpecialFolderLocation(Me.hWnd, CSIDL, pidl)
   
If r = NOERROR Then
     
'Of the structure is filled, initialize and
'retrieve the path from the id list, and return
'the folder with a trailing slash appended.
sPath = Space$(MAX_LENGTH)
r = SHGetPathFromIDList(ByVal pidl, ByVal sPath)
       
      If r Then
      GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1) & "\"
      Call CoTaskMemFree(pidl)
      End If
      
End If

End Function
Private Function fCreateShellLink(sLnkFile As String, sExeFile As String, sWorkDir As String, sIconFile As String, lIconIdx As Long, ShowCmd As SHOWCMDFLAGS) As Long

Dim cShellLink   As ShellLinkA   ' An explorer IShellLinkA(Win 9x/Win NT) instance
Dim cPersistFile As IPersistFile ' An explorer IPersistFile instance
    
If (sLnkFile = "") Or (sExeFile = "") Then Exit Function

On Error GoTo fCreateShellLinkError
Set cShellLink = New ShellLinkA   'Create new IShellLink interface
Set cPersistFile = cShellLink     'Implement cShellLink's IPersistFile interface
    
With cShellLink
    'Set command line exe name & path to new ShortCut.
    .SetPath sExeFile
    
    'Set working directory in shortcut
    If sWorkDir <> "" Then .SetWorkingDirectory sWorkDir
   
    
    'Set shortcut description
    .SetDescription "MyReels Profile" & vbNullChar
'   If (LnkDesc <> "") Then .SetDescription pszName
    
    'Set shortcut icon location & index
    If sIconFile <> "" Then .SetIconLocation sIconFile, lIconIdx
    
    'Set shortcut's startup mode (min,max,normal)
    .SetShowCmd ShowCmd
End With

cShellLink.Resolve 0, SLR_UPDATE
cPersistFile.Save StrConv(sLnkFile, vbUnicode), 0 'Unicode conversion that must be done!
fCreateShellLink = True                           'Return Success

fCreateShellLinkError:
Set cPersistFile = Nothing
Set cShellLink = Nothing
End Function
