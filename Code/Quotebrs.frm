VERSION 5.00
Begin VB.Form Quotebrs 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MyReels: Configure Quote Pictures"
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
   HelpContextID   =   16
   Icon            =   "Quotebrs.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   Begin VB.CommandButton cmmd 
      Caption         =   "&Compact"
      Height          =   375
      Index           =   4
      Left            =   4200
      TabIndex        =   39
      Top             =   8205
      Width           =   855
   End
   Begin VB.Timer Scrolltim 
      Interval        =   504
      Left            =   6480
      Top             =   6720
   End
   Begin VB.CheckBox Writetofile 
      Caption         =   "Write to text"
      Height          =   210
      Left            =   1560
      TabIndex        =   27
      Top             =   8280
      Width           =   1335
   End
   Begin VB.ComboBox quotetext 
      CausesValidation=   0   'False
      Height          =   2280
      Left            =   0
      Style           =   1  'Simple Combo
      TabIndex        =   17
      Text            =   " "
      Top             =   5880
      Width           =   5055
   End
   Begin VB.VScrollBar PicScroll 
      Height          =   2715
      Left            =   7560
      TabIndex        =   16
      Top             =   5880
      Width           =   180
   End
   Begin VB.CommandButton cmmd 
      Caption         =   "&Delete Quote"
      Height          =   375
      Index           =   3
      Left            =   6120
      TabIndex        =   15
      Top             =   8200
      Width           =   1215
   End
   Begin VB.CommandButton cmmd 
      Caption         =   "&Assign"
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   14
      Top             =   7270
      Width           =   1215
   End
   Begin VB.CommandButton cmmd 
      Caption         =   "A&ppend Quote"
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   13
      Top             =   7740
      Width           =   1215
   End
   Begin VB.ListBox quotelist 
      Height          =   2370
      Left            =   5160
      TabIndex        =   12
      Top             =   5880
      Width           =   855
   End
   Begin VB.CheckBox chkalwayssquare 
      Caption         =   "Square crops"
      Height          =   210
      Left            =   3000
      TabIndex        =   11
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Timer Drawlines 
      Enabled         =   0   'False
      Interval        =   18
      Left            =   7440
      Top             =   5520
   End
   Begin VB.PictureBox rub 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "Quotebrs.frx":000C
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5040
      Left            =   7440
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   180
      Left            =   2280
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.ComboBox PatternCombo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5400
      Width           =   1935
   End
   Begin VB.DriveListBox drvsource 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1700
   End
   Begin VB.DirListBox dirsource 
      Appearance      =   0  'Flat
      Height          =   1290
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.FileListBox filesource 
      Height          =   2610
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton cmmd 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   8200
      Width           =   1335
   End
   Begin VB.PictureBox Prevue 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DragIcon        =   "Quotebrs.frx":0316
      DragMode        =   1  'Automatic
      Height          =   5040
      Left            =   2160
      ScaleHeight     =   336
      ScaleMode       =   0  'User
      ScaleWidth      =   336
      TabIndex        =   5
      Top             =   240
      Width           =   5040
      Begin VB.PictureBox Prevuechild 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         DragIcon        =   "Quotebrs.frx":0460
         DragMode        =   1  'Automatic
         Height          =   5040
         Left            =   0
         ScaleHeight     =   336
         ScaleMode       =   0  'User
         ScaleWidth      =   336
         TabIndex        =   8
         ToolTipText     =   "Right - click and drag mouse to crop region"
         Top             =   0
         Width           =   5145
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   70.531
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   1
            Visible         =   0   'False
            X1              =   94.041
            X2              =   164.571
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   2
            Visible         =   0   'False
            X1              =   94.041
            X2              =   164.571
            Y1              =   8
            Y2              =   8
         End
         Begin VB.Line Clipborder 
            BorderStyle     =   2  'Dash
            Index           =   3
            Visible         =   0   'False
            X1              =   0
            X2              =   70.531
            Y1              =   8
            Y2              =   8
         End
      End
   End
   Begin VB.PictureBox picTruesize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "Quotebrs.frx":05AA
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   9
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Picspinindex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6060
      TabIndex        =   38
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lblProcess 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Processing ..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   1650
      TabIndex        =   37
      Top             =   6840
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   8
      Left            =   7800
      TabIndex        =   36
      Top             =   8040
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   7
      Left            =   7800
      TabIndex        =   35
      Top             =   7080
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   6
      Left            =   7800
      TabIndex        =   34
      Top             =   6120
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   5
      Left            =   7800
      TabIndex        =   33
      Top             =   5160
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   4
      Left            =   7800
      TabIndex        =   32
      Top             =   4200
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   3
      Left            =   7800
      TabIndex        =   31
      Top             =   3240
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   2
      Left            =   7800
      TabIndex        =   30
      Top             =   2280
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   1
      Left            =   7800
      TabIndex        =   29
      Top             =   1320
      Width           =   165
   End
   Begin VB.Label lblindex 
      Height          =   210
      Index           =   0
      Left            =   7800
      TabIndex        =   28
      Top             =   360
      Width           =   165
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   0
      Left            =   8250
      Top             =   20
      Width           =   900
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   8
      Left            =   9300
      TabIndex        =   26
      Top             =   7755
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   7
      Left            =   9300
      TabIndex        =   25
      Top             =   6795
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   6
      Left            =   9300
      TabIndex        =   24
      Top             =   5835
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   5
      Left            =   9300
      TabIndex        =   23
      Top             =   4875
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   4
      Left            =   9300
      TabIndex        =   22
      Top             =   3915
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   3
      Left            =   9300
      TabIndex        =   21
      Top             =   2955
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   2
      Left            =   9300
      TabIndex        =   20
      Top             =   1995
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   1
      Left            =   9300
      TabIndex        =   19
      Top             =   1035
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label RHSquote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   780
      Index           =   0
      Left            =   9300
      TabIndex        =   18
      Top             =   75
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   8
      Left            =   8250
      Top             =   7700
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   7
      Left            =   8250
      Top             =   6740
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   6
      Left            =   8250
      Top             =   5780
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   5
      Left            =   8250
      Top             =   4820
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   4
      Left            =   8250
      Top             =   3860
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   3
      Left            =   8250
      Top             =   2900
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   2
      Left            =   8250
      Top             =   1940
      Width           =   900
   End
   Begin VB.Image Spinimage 
      Height          =   900
      Index           =   1
      Left            =   8250
      Top             =   980
      Width           =   900
   End
   Begin VB.Image imgsel 
      BorderStyle     =   1  'Fixed Single
      Height          =   905
      Left            =   6215
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   905
   End
End
Attribute VB_Name = "Quotebrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thumbs As StdPicture
Dim LHSpoz As Long, imgselprevindex As Long, RHSpoz As Long, RHSoffset As Long, wipeimgsel As Boolean, imgseljustchanged As Boolean, ScrollKeydown As Boolean
Dim ct As Long, ct1 As Long, oldpicno As Long, oldX As Long, oldY As Long, cropwidth As Long, cropheight As Long
Dim newX As Long, newY As Long, entryX As Long, entryY As Long
Dim Hscrollvis As Boolean, Vscrollvis As Boolean, scrolltolH As Long, scrolltolV As Long, rightbuttdown As Boolean, boosloterror As Boolean
Dim bordN As Boolean, bordE As Boolean, bordS As Boolean, bordW As Boolean, DoingDragDrop As Boolean
Dim testtan As Single, tan1 As Single, tan3 As Single, tan5 As Single, tan7 As Single
Private Sub filesource_DblClick()
filesource.Refresh
End Sub
Private Sub Form_Load()
procend = False

On Error GoTo quoterror

If resX = 1 Then
setformpos Me
Else
Dotaskwindow Me, True
With Me

.Width = resX * .Width
.Height = resY * .Height

End With

setformpos Me

Prevue.Width = resX * Prevue.Width
Prevue.Height = resY * Prevue.Height
Prevuechild.Width = resX * Prevuechild.Width
Prevuechild.Height = resY * Prevuechild.Height
With imgsel
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With

With quotetext
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With

With chkalwayssquare
.Left = resX * .Left
.Top = resY * .Top
.Value = gt(44)
End With

With lblProcess
.Left = resX * .Left
.Top = resY * .Top
End With

With Writetofile
.Left = resX * .Left
.Top = resY * .Top
.Value = gt(189)
End With


For ct = 0 To 4
With cmmd(ct)
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
End With
Next

With quotelist
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With
End If



With Picspinindex
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
.Fontsize = 14 * resY
.Caption = padzeros(1)
End With


For ct = 0 To 8
With lblindex(ct)
.BackStyle = 1
.Left = resX * .Left
.Top = resY * .Top
.AutoSize = True
.Fontsize = CInt(8 * resY)
End With
With RHSquote(ct)
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
.Fontsize = 8 * resY
End With
With Spinimage(ct)
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
.Stretch = True
End With
Next
RHSquote(0).BorderStyle = 1
Spinimage(0).BorderStyle = 1

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


DoingDragDrop = False
rightbuttdown = False
wipeimgsel = False
oldpicno = 0
RHSoffset = 0
LHSpoz = 0
RHSpoz = 0
boosloterror = False
imgseljustchanged = False


For ct = 0 To 3
Clipborder(ct).Visible = False
Next


With PicScroll
If gt(191) > 0 Then
If gt(191) < 10 Then .Visible = False
.Value = 0
.Max = findhcmplusone(CInt(gt(191)), resY * 16)
.SmallChange = 1
.LargeChange = 9
.Height = resY * .Height
.Left = resX * .Left
.Top = resY * .Top
Else
.Height = resY * .Height
.Left = resX * .Left
.Top = resY * .Top
GoTo quoterror
End If
End With



'Need to clear Quotes.s$t + Filesize warnings!
If Stringvars(3) = "" Or Dir(Stringvars(3) & "Quotes.s$t") = "" Then GoTo quoterror 'delete contents of quotes


BitmapDb -2, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

BitmapDb 1, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel 'load first picture in imagsel

DoEvents

BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel


'FILL quotelist
quotelist.Selected(0) = True


Quotebrs.Show
procend = True
Exit Sub


quoterror:
ShowError
MsgBox "Quotes.s$t could not be opened", vbOKOnly
Prevue.Enabled = False
Prevuechild.Enabled = False
chkalwayssquare.Enabled = False
Writetofile.Enabled = False
PicScroll.Enabled = False

Picspinindex.Caption = ""

For ct = 1 To 4
cmmd(ct).Enabled = False
Next

For ct = 0 To 8
lblindex(ct).Visible = False
Spinimage(ct).Visible = False
RHSquote(ct).Visible = False
Next

quotetext.Visible = False
quotelist.Visible = False


imgsel.Enabled = False
With lblProcess
.Width = 2 * .Width
.Caption = "Inaccessible: Click ""Use Base Dir"" && retry."
End With
drvsource.Visible = False
dirsource.Visible = False
filesource.Visible = False
PatternCombo.Visible = False

Quotebrs.Show
procend = True
End Sub
Private Sub Form_Activate()
Dim f As Form 'For DB reference
For Each f In Forms
If f.Name = "Zhidden" Then
Zhidden.Show
Exit Sub
End If
Next
If zhiddnstatus < 0 Then
BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
zhiddnstatus = 0
End If
End Sub
Private Sub Form_Resize()
Dim wid As Single, hgt As Single
Quotebrs.ScaleMode = 1
Const GAP = 60
Quotebrs.Caption = "Configure Quote Thumbnails"  'Space(109)


filesource.Path = CurDir

    'If WindowState = 1 Then Exit Sub

    wid = drvsource.Width
    drvsource.Move GAP, GAP, wid
    PatternCombo.Move GAP, ScaleHeight - quotetext.Height - 2 * cmmd(0).Height - 100 - PatternCombo.Height - GAP, wid
    hgt = (PatternCombo.Top - drvsource.Top - drvsource.Height - GAP) / 2
    If hgt < 100 Then hgt = 100
    dirsource.Move GAP, drvsource.Top + drvsource.Height + GAP, wid, hgt
    filesource.Move GAP, dirsource.Top + dirsource.Height + GAP, wid, hgt
    
Quotebrs.ScaleMode = 3
Dotaskwindow Me
End Sub
Private Sub cmmd_Click(Index As Integer)
Dim tempstr As String
procend = False

On Error GoTo Sloterror

Select Case Index
Case 0

If drvsource.Enabled = False Then
gt(195) = 0
Stringvars(3) = ""
With Genopts
.chkgenopt(6).Value = 0
.chkgenopt(7).Enabled = False
.chkgenopt(9).Enabled = False
.cmdgenopts(5).Enabled = False
.cmdgenopts(35).Enabled = False
End With

'And this dialog closed
End If

'Clean up
For ct = 0 To 8
If Dir(loaddirectory & "q" & ct & ".bmp") <> "" Then Kill (loaddirectory & "q" & ct & ".bmp")
Next

If Stringvars(3) <> "" Then BitmapDb 4, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel


Unload Quotebrs
Set Quotebrs = Nothing
Genopts.Enabled = True
Genopts.Show



Case 1  'Assign
ProcessQuotez True
BitmapDb 2, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel, imgseljustchanged
    If gt(191) > 0 Then 'Assign succeeded
    BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
    imgseljustchanged = False
    Else
    gt(191) = -gt(191)
    End If
quotelist.ListIndex = LHSpoz
quotetext.ListIndex = LHSpoz
ProcessQuotez
Case 2  'Insert
ProcessQuotez True
BitmapDb 3, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

If gt(191) > 0 Then  'Append succeeded
'LHSpoz is absolute position of last modified
quotetext.AddItem "New"
quotelist.AddItem LHSpoz
quotelist.List(LHSpoz) = CStr(LHSpoz + 1)
BitmapDb -2, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

If gt(191) < 10 Or RHSoffset > LHSpoz - 10 Then BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

If RHSpoz >= LHSpoz Then Picspinindex.Caption = padzeros(RHSpoz + 1)

quotetext.ListIndex = LHSpoz
quotelist.ListIndex = LHSpoz
'refresh
With PicScroll
If gt(191) > 9 Then .Visible = True
.Max = findhcmplusone(CInt(gt(191)), resY * 16)
.SmallChange = 1
.LargeChange = 9
.Value = RHSoffset
End With
Else
gt(191) = -gt(191)
End If

ProcessQuotez
Case 3  'delete

If cmmd(3).Caption = "&Delete Quote" Then
cmmd(3).Caption = "Sure?"
procend = True
Exit Sub
End If

ProcessQuotez True
BitmapDb -1, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
'refresh
BitmapDb -2, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
BitmapDb 1, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

If ((gt(191) - LHSpoz < 10 And LHSpoz - RHSoffset < 10)) And RHSoffset > 0 And gt(191) > 9 Then
RHSoffset = RHSoffset - 1
RHSpoz = RHSpoz - 1
End If
If (RHSoffset >= LHSpoz - 10 And RHSoffset < gt(191) - 8) Or gt(191) <= 10 Then
BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
Picspinindex.Caption = padzeros(RHSpoz + 1)
End If

If gt(191) < 10 Then PicScroll.Visible = False
ProcessQuotez
Case 4
ProcessQuotez True
sDatabaseName = Stringvars(3) & "Quotes.s$t"
If compactdb(2) = False Then GoTo Sloterror
sDatabaseName = ""
ProcessQuotez
End Select

procend = True
Exit Sub

Sloterror:
procend = True
ShowError
ProcessQuotez
End Sub
Private Sub imgsel_DragDrop(Source As Control, X As Single, Y As Single)
If DoingDragDrop = True Then Exit Sub
DoDragdrop
End Sub
Private Sub imgsel_Click()
On Error GoTo Noblank
imgseljustchanged = True
wipeimgsel = True
imgselprevindex = 0
Set imgsel.Picture = LoadPicture(App.Path & "\blank.bmp")
Set thumbs = Nothing
Exit Sub

Noblank:
Set thumbs = Nothing
ShowError
End Sub
Private Sub quotetext_Click()
If procend = False Then Exit Sub
procend = False

cmmd(3).Caption = "&Delete Quote"
LHSpoz = quotetext.ListIndex
quotelist.ListIndex = LHSpoz
BitmapDb 1, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
imgseljustchanged = True

'Refresh RHS if apt
If LHSpoz > RHSoffset And LHSpoz < RHSoffset + 9 Then BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

procend = True
End Sub
Private Sub quotelist_Click()
If procend = False Then Exit Sub
'select quotext

cmmd(3).Caption = "&Delete Quote"
LHSpoz = quotelist.ListIndex
quotetext.ListIndex = LHSpoz
BitmapDb 1, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
imgseljustchanged = True
'Refresh RHS if apt

If LHSpoz > RHSoffset And LHSpoz < RHSoffset + 9 Then BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel

End Sub
Private Sub RHSquote_Click(Index As Integer)
Spinimage_Click (Index)
End Sub
Private Sub Spinimage_Click(Index As Integer)
RHSpoz = Index + RHSoffset
BitmapDb -3, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
Set imgsel.Picture = Spinimage(Index).Picture
imgseljustchanged = True
For ct = 0 To 8
If ct = Index Then
Spinimage(ct).BorderStyle = 1
RHSquote(ct).BorderStyle = 1
Else
Spinimage(ct).BorderStyle = 0
RHSquote(ct).BorderStyle = 0
End If
Next
Picspinindex.Caption = padzeros(RHSpoz + 1)
End Sub
Private Sub lblindex_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If procend = False Or (gt(191) < 20000 Or gt(196) < 400) Then Exit Sub  'Only generic DBs
procend = False
    If Button = 1 Then
    BmpIX = CLng(lblindex(Index).Caption)
    PlaySndF ("Thumbb")
    Else
    Unload Zhidden
    Set Zhidden = Nothing
    zhiddnstatus = -1
    Load Zhidden
    End If
procend = True
End Sub
Private Sub PicScroll_Scroll()
Picspinindex.Caption = padzeros(PicScroll.Value + 1)
End Sub
Private Sub PicScroll_Change()
If procend = False Or ScrollKeydown = True Then Exit Sub
procend = False

cmmd(3).Caption = "&Delete Quote"
RHSoffset = PicScroll.Value
If RHSoffset + 9 > gt(191) Then RHSoffset = gt(191) - 9
RHSpoz = RHSoffset
PicScroll.Value = RHSoffset
Picspinindex.Caption = padzeros(RHSoffset + 1)


BitmapDb 0, RHSoffset, RHSpoz, LHSpoz, imgselprevindex, wipeimgsel
For ct = 0 To 8
Spinimage(ct).BorderStyle = 0
RHSquote(ct).BorderStyle = 0
Next
Quotebrs.Enabled = False
Scrolltim.Enabled = True
procend = True
End Sub
Private Sub PicScroll_KeyUp(KeyCode As Integer, Shift As Integer)
ScrollKeydown = False
End Sub
Private Sub PicScroll_KeyDown(KeyCode As Integer, Shift As Integer)
ScrollKeydown = True
End Sub
Private Sub Scrolltim_Timer()
Quotebrs.Enabled = True
Scrolltim.Enabled = False
End Sub
Private Sub dirsource_Change()
filesource.Path = dirsource.Path
End Sub
Private Sub chkalwayssquare_Click()
If procend = False Then Exit Sub
If chkalwayssquare.Value = 0 Then
gt(44) = 0
Else
gt(44) = 1
End If
End Sub
Private Sub Writetofile_Click()
If procend = False Then Exit Sub
If Writetofile.Value = 0 Then
gt(189) = 0
Else
gt(189) = 1
End If
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
    
    
    Prevue.Move resX * 118, resY * 12, resX * 336, resY * 336
    Prevuechild.Move 0, 0
    
    
    HScroll1.Left = resX * 118
    If Prevuechild.Height < resY * 336 Then
    HScroll1.Top = resY * 12 + Prevuechild.Height
    VScroll1.Height = Prevuechild.Height
    Prevue.Height = Prevuechild.Height
    Else
    HScroll1.Top = resY * (12 + 336)
    VScroll1.Height = resY * 336
    End If
    
    VScroll1.Top = resY * 12
    If Prevuechild.Width < resX * 336 Then
    VScroll1.Left = resX * 118 + Prevuechild.Width
    HScroll1.Width = Prevuechild.Width
    Prevue.Width = Prevuechild.Width
    Else
    VScroll1.Left = resX * (118 + 336)
    HScroll1.Width = resX * 336
    End If
    
    
    'make .max multiple of 12
    HScroll1.Max = findhcmplusone(Prevuechild.Width - resX * 336, resX * 12)
    VScroll1.Max = findhcmplusone(Prevuechild.Height - resY * 336, resY * 12)
    
    Hscrollvis = Prevue.Width < Prevuechild.Width
    Vscrollvis = Prevue.Height < Prevuechild.Height
    HScroll1.Visible = Hscrollvis
    VScroll1.Visible = Vscrollvis
    'set arbitrary tolerance
    If Hscrollvis = True Then
    HScroll1.LargeChange = (HScroll1.Max / 6)
    HScroll1.SmallChange = HScroll1.Max / 12
    scrolltolH = 1 + Int(HScroll1.SmallChange)
    scrolltolV = scrolltolH
    End If
    If Vscrollvis = True Then
    VScroll1.LargeChange = VScroll1.Max / 6
    VScroll1.SmallChange = VScroll1.Max / 12
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
Private Sub DoDragdrop()
On Error GoTo Quit
Prevuechild.ToolTipText = ""
DoingDragDrop = True
Screen.MousePointer = vbHourglass
For ct1 = 0 To 3
Clipborder(ct1).Visible = True
Next
Set picTruesize = Nothing
ScaleMode = 1

With picTruesize
If cropwidth = 0 Then
.Width = 1160
.Height = 1160
.PaintPicture Prevuechild.Picture, 0, 0, 1160, 1160
Set thumbs = .Image
Set picTruesize = Nothing
.Width = 1160
.Height = 1160
.PaintPicture Prevuechild.Picture, 0, 0, 1160, 1160
Else
'scalemode twips
.Width = 1160
.Height = 1160
.PaintPicture rub.Picture, 0, 0, 1160, 1160
Set thumbs = .Image
Set picTruesize = Nothing
.Width = 1160
.Height = 1160
.PaintPicture rub.Picture, 0, 0, 1160, 1160
End If
imgsel.Picture = .Image
imgselprevindex = 0
imgseljustchanged = True
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
Screen.MousePointer = vbDefault
DoingDragDrop = False
wipeimgsel = False
'.ToolTipText = infofortt(ct - oldpicno, drvsource.Drive)
'.ToolTipText = infofortt(ct, drvsource.Drive)
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
Screen.MousePointer = vbDefault
DoingDragDrop = False
ShowError
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
Private Function padzeros(argz As Long) As String
If argz < gt(191) Then
ct = argz
Else
ct = gt(191)
End If

Select Case ct
Case Is < 10
padzeros = "0000000" & ct
Case Is < 100
padzeros = "000000" & ct
Case Is < 1000
padzeros = "00000" & ct
Case Is < 10000
padzeros = "0000" & ct
Case Is < 100000
padzeros = "000" & ct
Case Is < 1000000
padzeros = "00" & ct
Case Is < 10000000
padzeros = "0" & ct
Case Else
padzeros = ct
End Select
End Function
Private Sub ProcessQuotez(Optional Dizable As Boolean = False)
cmmd(3).Caption = "&Delete Quote"
If Dizable = True Then
Genopts.enabled = false
Gametype.enabled = false
quotetext.Visible = False
quotelist.Visible = False
Quotebrs.Refresh
Quotebrs.Enabled = False
Quotebrs.MousePointer = vbHourglass
Else
Genopts.MidiPlay.Interval = 18
Genopts.MidiPlay.Enabled = True
Quotebrs.MousePointer = vbDefault
Quotebrs.Enabled = True
quotetext.Visible = True
quotelist.Visible = True
End If
End Sub
