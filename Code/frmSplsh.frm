VERSION 5.00
Begin VB.Form frmSplsh 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2850
   ClientLeft      =   3495
   ClientTop       =   2940
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplsh.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5240
      Begin VB.Image imgLogo 
         Height          =   1335
         Left            =   120
         Picture         =   "frmSplsh.frx":000C
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright 1999-2017"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblWarning 
         Caption         =   "Usage and distribution- See Disclaimer in Help for details."
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3750
         TabIndex        =   4
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Win9x ==> Win10"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1760
         TabIndex        =   5
         Top             =   1440
         Width           =   3150
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   7
         Top             =   600
         Width           =   165
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "Licensed for personal use only"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Stearn && DisAssociates Presents"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   105
         TabIndex        =   6
         Top             =   120
         Width           =   4965
      End
   End
   Begin VB.Label lblconfigload 
      AutoSize        =   -1  'True
      Caption         =   " Loading Profile Configuration Mgr ...."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5115
   End
End
Attribute VB_Name = "frmSplsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
If gt(0) = -4 Then gt(0) = 0
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblProductName.Caption = App.Title
DoMe
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoMe
End Sub
Private Sub Form_Click()
DoMe
End Sub
Private Sub Form_Resize()
Dim response As Integer
'resCheck
Dim Mon As New cRegistry
Mon.hDC = Me.hDC
resCheck = True
Select Case Mon.DCWidth
Case Is <= 640
response = MsgBox("Sorry, you need a resolution of 800 X 600 or greater.", vbOKOnly)
resCheck = False
End Select
setformpos Me
End Sub
Private Sub DoMe()
Sleep 50
DoEvents
lblPlatform.ToolTipText = "Win95; Win98; WinNT; Win2000; Win XP; Win7; Win8.x; Win 10; Me."
End Sub
