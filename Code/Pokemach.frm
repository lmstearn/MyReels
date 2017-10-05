VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Object = "{C3CBD80D-C8D1-11D2-9F8E-0080C7CE5CDC}#4.1#0"; "ActCndy2.ocx"
Begin VB.Form Pokemach 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "MyReels"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   ClipControls    =   0   'False
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
   HelpContextID   =   1
   Icon            =   "Pokemach.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12060
   Begin ActiveCandy.CandyCommand Candy 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   176
      Top             =   6360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ForeColor       =   0
      FontName        =   "Times New Roman"
      FontSize        =   8.25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ActiveCandy.CandyCommand Candy 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   177
      Top             =   6360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ForeColor       =   0
      FontName        =   "Times New Roman"
      FontSize        =   8.25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox TA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1695
      Left            =   3240
      ScaleHeight     =   1695
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Frame frapicarea 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   " "
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   3350
      TabIndex        =   3
      Top             =   3050
      Width           =   5325
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   1
         Left            =   0
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   2
         Left            =   0
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   3
         Left            =   0
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   4
         Left            =   1080
         Top             =   0
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   5
         Left            =   1080
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   6
         Left            =   1080
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   7
         Left            =   1080
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   8
         Left            =   2160
         Top             =   0
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   9
         Left            =   2160
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   10
         Left            =   2160
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   11
         Left            =   2160
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   12
         Left            =   3240
         Top             =   0
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   13
         Left            =   3240
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   14
         Left            =   3240
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   15
         Left            =   3240
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   16
         Left            =   4320
         Top             =   0
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   17
         Left            =   4320
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   18
         Left            =   4320
         Top             =   2160
         Width           =   1000
      End
      Begin VB.Image M 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1000
         Index           =   19
         Left            =   4320
         Top             =   3240
         Width           =   1000
      End
   End
   Begin VB.Timer Vanishlines 
      Enabled         =   0   'False
      Interval        =   18
      Left            =   6800
      Top             =   1560
   End
   Begin VB.Timer Prizemeter 
      Enabled         =   0   'False
      Left            =   6840
      Top             =   1560
   End
   Begin VB.Timer gamespinwait 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   7800
      Top             =   1560
   End
   Begin VB.PictureBox Spare 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   400
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer jackpot 
      Enabled         =   0   'False
      Interval        =   288
      Left            =   7560
      Top             =   1560
   End
   Begin VB.Timer waitimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7920
      Top             =   1560
   End
   Begin VB.Timer timoneyback 
      Enabled         =   0   'False
      Interval        =   288
      Left            =   6960
      Top             =   1560
   End
   Begin VB.Timer Prizeflash 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   7320
      Top             =   1560
   End
   Begin VB.Timer midiplay 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   8280
      Top             =   1560
   End
   Begin VB.Image Quotez 
      Appearance      =   0  'Flat
      Height          =   905
      Left            =   8070
      Stretch         =   -1  'True
      Top             =   7595
      Width           =   905
   End
   Begin VB.Shape Lnptr 
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   5
      Left            =   8680
      Shape           =   2  'Oval
      Top             =   5620
      Width           =   105
   End
   Begin VB.Shape Lnptr 
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   4
      Left            =   3230
      Shape           =   2  'Oval
      Top             =   5620
      Width           =   105
   End
   Begin VB.Shape Lnptr 
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   3
      Left            =   8680
      Shape           =   2  'Oval
      Top             =   3480
      Width           =   105
   End
   Begin VB.Shape Lnptr 
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   3230
      Shape           =   2  'Oval
      Top             =   3480
      Width           =   105
   End
   Begin VB.Shape Lnptr 
      FillStyle       =   0  'Solid
      Height          =   200
      Index           =   1
      Left            =   8680
      Shape           =   2  'Oval
      Top             =   4520
      Width           =   100
   End
   Begin VB.Shape Lnptr 
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   3230
      Shape           =   2  'Oval
      Top             =   4520
      Width           =   105
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   13
      Left            =   5500
      TabIndex        =   175
      Top             =   8270
      Width           =   60
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   5500
      TabIndex        =   174
      Top             =   8030
      Width           =   60
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   11
      Left            =   5500
      TabIndex        =   173
      Top             =   7790
      Width           =   60
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   10
      Left            =   5500
      TabIndex        =   172
      Top             =   7550
      Width           =   60
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   13
      Left            =   11520
      Top             =   6120
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   12
      Left            =   120
      Top             =   6120
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   11
      Left            =   11520
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   10
      Left            =   120
      Top             =   5160
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   9
      Left            =   11520
      Top             =   4200
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   8
      Left            =   120
      Top             =   4200
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   7
      Left            =   11520
      Top             =   3240
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   6
      Left            =   120
      Top             =   3240
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   5
      Left            =   11520
      Top             =   2280
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   4
      Left            =   120
      Top             =   2280
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   3
      Left            =   11520
      Top             =   1320
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   2
      Left            =   120
      Top             =   1320
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   1
      Left            =   11520
      Top             =   360
      Width           =   405
   End
   Begin VB.Image imgprizethumb 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   0
      Left            =   120
      Top             =   360
      Width           =   405
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   4
      Left            =   5955
      TabIndex        =   1
      Top             =   6840
      Width           =   60
   End
   Begin VB.Label lblmisc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   9
      Left            =   9960
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblmisc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   7780
      TabIndex        =   171
      Top             =   810
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblmisc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   8
      Left            =   9960
      TabIndex        =   170
      Top             =   6720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   7320
      TabIndex        =   169
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   6960
      TabIndex        =   168
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scatters"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   960
      TabIndex        =   167
      Top             =   6720
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   9240
      TabIndex        =   166
      Top             =   7920
      Width           =   315
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   8400
      TabIndex        =   165
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblrandom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   8595
      TabIndex        =   164
      Top             =   130
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Random Jackpot now at ..."
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   3240
      TabIndex        =   163
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblrandom 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0000000000"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   6720
      TabIndex        =   162
      Top             =   130
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   3240
      TabIndex        =   160
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   3480
      TabIndex        =   159
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   3720
      TabIndex        =   158
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   3960
      TabIndex        =   157
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   4200
      TabIndex        =   156
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   5
      Left            =   4440
      TabIndex        =   155
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   6
      Left            =   4680
      TabIndex        =   154
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   7
      Left            =   4920
      TabIndex        =   153
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   5160
      TabIndex        =   152
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   9
      Left            =   5520
      TabIndex        =   151
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   5880
      TabIndex        =   150
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   11
      Left            =   6240
      TabIndex        =   149
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblmoneyback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   6600
      TabIndex        =   148
      Top             =   960
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   147
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   1
      Left            =   10920
      TabIndex        =   146
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   145
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   3
      Left            =   10920
      TabIndex        =   144
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   4
      Left            =   600
      TabIndex        =   143
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   5
      Left            =   10920
      TabIndex        =   142
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   6
      Left            =   600
      TabIndex        =   141
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   7
      Left            =   10920
      TabIndex        =   140
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   8
      Left            =   600
      TabIndex        =   139
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   9
      Left            =   10920
      TabIndex        =   138
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   10
      Left            =   600
      TabIndex        =   137
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   11
      Left            =   10920
      TabIndex        =   136
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   12
      Left            =   600
      TabIndex        =   135
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   13
      Left            =   10920
      TabIndex        =   134
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   14
      Left            =   1200
      TabIndex        =   133
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   15
      Left            =   10320
      TabIndex        =   132
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   16
      Left            =   1200
      TabIndex        =   131
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   17
      Left            =   10320
      TabIndex        =   130
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   18
      Left            =   1200
      TabIndex        =   129
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   19
      Left            =   10320
      TabIndex        =   128
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   20
      Left            =   1200
      TabIndex        =   127
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   21
      Left            =   10320
      TabIndex        =   126
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   22
      Left            =   1200
      TabIndex        =   125
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   23
      Left            =   10320
      TabIndex        =   124
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   24
      Left            =   1200
      TabIndex        =   123
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   25
      Left            =   10320
      TabIndex        =   122
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   26
      Left            =   1200
      TabIndex        =   121
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   27
      Left            =   10320
      TabIndex        =   120
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   28
      Left            =   1800
      TabIndex        =   119
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   29
      Left            =   9840
      TabIndex        =   118
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   30
      Left            =   1800
      TabIndex        =   117
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   31
      Left            =   9840
      TabIndex        =   116
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   32
      Left            =   1800
      TabIndex        =   115
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   33
      Left            =   9840
      TabIndex        =   114
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   34
      Left            =   1800
      TabIndex        =   113
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   35
      Left            =   9840
      TabIndex        =   112
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   36
      Left            =   1800
      TabIndex        =   111
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   37
      Left            =   9840
      TabIndex        =   110
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   38
      Left            =   1800
      TabIndex        =   109
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   39
      Left            =   9840
      TabIndex        =   108
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   40
      Left            =   1800
      TabIndex        =   107
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   41
      Left            =   9840
      TabIndex        =   106
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   42
      Left            =   2280
      TabIndex        =   105
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   43
      Left            =   9360
      TabIndex        =   104
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   44
      Left            =   2280
      TabIndex        =   103
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   45
      Left            =   9360
      TabIndex        =   102
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   46
      Left            =   2280
      TabIndex        =   101
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   47
      Left            =   9360
      TabIndex        =   100
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   48
      Left            =   2280
      TabIndex        =   99
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   49
      Left            =   9360
      TabIndex        =   98
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   50
      Left            =   2280
      TabIndex        =   97
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   51
      Left            =   9360
      TabIndex        =   96
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   52
      Left            =   2280
      TabIndex        =   95
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   53
      Left            =   9360
      TabIndex        =   94
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   54
      Left            =   2280
      TabIndex        =   93
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   55
      Left            =   9360
      TabIndex        =   92
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   56
      Left            =   2760
      TabIndex        =   91
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   57
      Left            =   8880
      TabIndex        =   90
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   58
      Left            =   2760
      TabIndex        =   89
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   59
      Left            =   8880
      TabIndex        =   88
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   60
      Left            =   2760
      TabIndex        =   87
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   61
      Left            =   8880
      TabIndex        =   86
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   62
      Left            =   2760
      TabIndex        =   85
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   63
      Left            =   8880
      TabIndex        =   84
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   64
      Left            =   2760
      TabIndex        =   83
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   65
      Left            =   8880
      TabIndex        =   82
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   66
      Left            =   2760
      TabIndex        =   81
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   67
      Left            =   8880
      TabIndex        =   80
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   68
      Left            =   2760
      TabIndex        =   79
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprize 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   69
      Left            =   8880
      TabIndex        =   78
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   77
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   1
      Left            =   10920
      TabIndex        =   76
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   2
      Left            =   600
      TabIndex        =   75
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   3
      Left            =   10920
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   4
      Left            =   600
      TabIndex        =   73
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   5
      Left            =   10920
      TabIndex        =   72
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   6
      Left            =   600
      TabIndex        =   71
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   7
      Left            =   10920
      TabIndex        =   70
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   8
      Left            =   600
      TabIndex        =   69
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   9
      Left            =   10920
      TabIndex        =   68
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   10
      Left            =   600
      TabIndex        =   67
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   11
      Left            =   10920
      TabIndex        =   66
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   12
      Left            =   600
      TabIndex        =   65
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   13
      Left            =   10920
      TabIndex        =   64
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   14
      Left            =   1200
      TabIndex        =   63
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   15
      Left            =   10320
      TabIndex        =   62
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   16
      Left            =   1200
      TabIndex        =   61
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   17
      Left            =   10320
      TabIndex        =   60
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   18
      Left            =   1200
      TabIndex        =   59
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   19
      Left            =   10320
      TabIndex        =   58
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   20
      Left            =   1200
      TabIndex        =   57
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   21
      Left            =   10320
      TabIndex        =   56
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   22
      Left            =   1200
      TabIndex        =   55
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   23
      Left            =   10320
      TabIndex        =   54
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   24
      Left            =   1200
      TabIndex        =   53
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   25
      Left            =   10320
      TabIndex        =   52
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   26
      Left            =   1200
      TabIndex        =   51
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   27
      Left            =   10320
      TabIndex        =   50
      Top             =   6480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   28
      Left            =   1800
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   29
      Left            =   9840
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   30
      Left            =   1800
      TabIndex        =   47
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   31
      Left            =   9840
      TabIndex        =   46
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   32
      Left            =   1800
      TabIndex        =   45
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   33
      Left            =   9840
      TabIndex        =   44
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   34
      Left            =   1800
      TabIndex        =   43
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   35
      Left            =   9840
      TabIndex        =   42
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   36
      Left            =   1800
      TabIndex        =   41
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   37
      Left            =   9840
      TabIndex        =   40
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   38
      Left            =   1800
      TabIndex        =   39
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   39
      Left            =   9840
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   40
      Left            =   1800
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   41
      Left            =   9840
      TabIndex        =   36
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   42
      Left            =   2280
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   43
      Left            =   9360
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   44
      Left            =   2280
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   45
      Left            =   9360
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   46
      Left            =   2280
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   47
      Left            =   9360
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   48
      Left            =   2280
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   49
      Left            =   9360
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   50
      Left            =   2280
      TabIndex        =   27
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   51
      Left            =   9360
      TabIndex        =   26
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   52
      Left            =   2280
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   53
      Left            =   9360
      TabIndex        =   24
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   54
      Left            =   2280
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   55
      Left            =   9360
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   56
      Left            =   2760
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   57
      Left            =   8880
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   58
      Left            =   2760
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   59
      Left            =   8880
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   60
      Left            =   2760
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   61
      Left            =   8880
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   62
      Left            =   2760
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   63
      Left            =   8880
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   64
      Left            =   2760
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   65
      Left            =   8880
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   66
      Left            =   2760
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   67
      Left            =   8880
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   68
      Left            =   2760
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizeamt 
      Height          =   240
      Index           =   69
      Left            =   8880
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprizemeter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   9600
      TabIndex        =   7
      Top             =   7920
      Width           =   105
   End
   Begin ComctlLib.ImageList Thumbslist 
      Index           =   0
      Left            =   3240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList Thumbslist 
      Index           =   1
      Left            =   3840
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList Thumbslist 
      Index           =   2
      Left            =   4440
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList Thumbslist 
      Index           =   3
      Left            =   5040
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList Thumbslist 
      Index           =   4
      Left            =   5640
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList Thumbslist 
      Index           =   5
      Left            =   6240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   3335
      X2              =   3335
      Y1              =   6215
      Y2              =   3020
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   3350
      X2              =   8660
      Y1              =   6215
      Y2              =   6215
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   8670
      X2              =   8670
      Y1              =   6215
      Y2              =   3020
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   3350
      X2              =   8680
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label lblprizemeter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   10040
      TabIndex        =   6
      Top             =   7162
      Width           =   90
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   9000
      X2              =   8880
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   9840
      X2              =   9720
      Y1              =   6960
      Y2              =   7200
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   8880
      X2              =   9000
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   9840
      X2              =   9720
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   9240
      X2              =   9360
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   5
      X1              =   9480
      X2              =   9360
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   9360
      X2              =   9240
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   7
      X1              =   9360
      X2              =   9480
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   8
      X1              =   9000
      X2              =   9120
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   9
      X1              =   9240
      X2              =   9120
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   10
      X1              =   9480
      X2              =   9600
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   11
      X1              =   9720
      X2              =   9600
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   12
      X1              =   9120
      X2              =   9000
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   13
      X1              =   9120
      X2              =   9240
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   14
      X1              =   9480
      X2              =   9600
      Y1              =   7440
      Y2              =   7680
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   15
      X1              =   9600
      X2              =   9720
      Y1              =   7680
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   16
      X1              =   8760
      X2              =   9000
      Y1              =   7200
      Y2              =   7320
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   17
      X1              =   9000
      X2              =   8760
      Y1              =   7320
      Y2              =   7440
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   18
      X1              =   9960
      X2              =   9720
      Y1              =   7200
      Y2              =   7320
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   19
      X1              =   9960
      X2              =   9720
      Y1              =   7440
      Y2              =   7320
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   20
      X1              =   8760
      X2              =   8880
      Y1              =   7200
      Y2              =   6960
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   21
      X1              =   9840
      X2              =   9960
      Y1              =   6960
      Y2              =   7200
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   22
      X1              =   9960
      X2              =   9840
      Y1              =   7440
      Y2              =   7680
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   23
      X1              =   8760
      X2              =   8880
      Y1              =   7440
      Y2              =   7680
   End
   Begin VB.Label lblmisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   9285
      TabIndex        =   5
      Top             =   7155
      Width           =   105
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   24
      X1              =   9000
      X2              =   8880
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   25
      X1              =   9840
      X2              =   9720
      Y1              =   6840
      Y2              =   7080
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   26
      X1              =   9840
      X2              =   9720
      Y1              =   7800
      Y2              =   7560
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   27
      X1              =   8880
      X2              =   9000
      Y1              =   7800
      Y2              =   7560
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   28
      X1              =   9240
      X2              =   9360
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   29
      X1              =   9480
      X2              =   9360
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   30
      X1              =   9360
      X2              =   9240
      Y1              =   7800
      Y2              =   7560
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   31
      X1              =   9480
      X2              =   9360
      Y1              =   7560
      Y2              =   7800
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   32
      X1              =   9000
      X2              =   9120
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   33
      X1              =   9240
      X2              =   9120
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   34
      X1              =   9480
      X2              =   9600
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   35
      X1              =   9720
      X2              =   9600
      Y1              =   7080
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   36
      X1              =   9120
      X2              =   9000
      Y1              =   7800
      Y2              =   7560
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   37
      X1              =   9240
      X2              =   9120
      Y1              =   7560
      Y2              =   7800
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   38
      X1              =   9600
      X2              =   9480
      Y1              =   7800
      Y2              =   7560
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   39
      X1              =   9600
      X2              =   9720
      Y1              =   7800
      Y2              =   7560
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   40
      X1              =   8880
      X2              =   8640
      Y1              =   7320
      Y2              =   7200
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   41
      X1              =   8640
      X2              =   8880
      Y1              =   7440
      Y2              =   7320
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   42
      X1              =   10080
      X2              =   9840
      Y1              =   7200
      Y2              =   7320
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   43
      X1              =   10080
      X2              =   9840
      Y1              =   7440
      Y2              =   7320
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   44
      X1              =   8640
      X2              =   8880
      Y1              =   7200
      Y2              =   6840
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   45
      X1              =   9840
      X2              =   10080
      Y1              =   6840
      Y2              =   7200
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   46
      X1              =   10080
      X2              =   9840
      Y1              =   7440
      Y2              =   7800
   End
   Begin VB.Line liwnf 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   47
      X1              =   8640
      X2              =   8880
      Y1              =   7440
      Y2              =   7800
   End
   Begin VB.Label lblmisc 
      BackStyle       =   0  'Transparent
      Caption         =   "Money Back Status ....                POT ...."
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   161
      Top             =   480
      Visible         =   0   'False
      Width           =   5595
   End
End
Attribute VB_Name = "Pokemach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim storit(4) As Long, trakker(4) As Long, movesize(4) As Long
Dim ydisp(4, 4) As Long, medfastslowmove(4, 24) As Long, ydisp0(4) As Long, spinztemp(4) As Long, multiprize(2, 9, 3) As Long
Dim wheelvec(5, 14) As Long, thumbs(14) As StdPicture
Dim wheelorder(4, 24) As Long, oldlblprize(15) As Long
Dim pw As Long, pwScale As Long, TOL As Long, prizetotal As Long, prizeaccum As Long, betqut As Long, aptot As Long, jackpotprize As Long, pzctcum As Long
Dim intreel As Long, ct As Long, ct1 As Long, ct2 As Long, ct3 As Long
'ct3: Jack & Monback flash
Dim reelmin As Long, reelmax As Long, oldspecial As Long, special As Long, oldMidiNo As Long
Dim spinstop As Boolean, spinzstart As Long, piccount As Long, currold As Long, currline As Long, lnbet As Long
Dim gt80 As Boolean, gt76 As Boolean, gt72 As Boolean, gt132 As Boolean, keyPrsd As Boolean
Dim ydispmin As Long, ydispmax As Long, cumpzzeros As String, pzzeros As String
Dim mt As Long, degreeoftitle As Long, prizecount(2) As Long, prizeaccold As Long
Dim spincount As Long, freegamecount As Long, kept1or2 As Long, gamespintot As Long, lbloffset As Long, wid As Long, hgt As Long
Dim gamenotspin As Boolean, response As Long, justrestored As Boolean, flashtoggle As Boolean, prizeflashon As Boolean, activatctrls As Boolean
Dim gamesaved As Boolean, spinsaved As Boolean, featurereset As Boolean, EOGEF As Boolean, playsuccess As Boolean

Const EOGTT = "End_of_Game. You will have to start a new one in the Configuration Panel."
Const Betlim = "Bet Total limit"
Const Sublim = "Substitute prize limit"
Const Scatlim = "Scatter prize limit"
Const Pzlim = "Naturals prize limit"
Const Jacklim = "Jackpot limit"
Const Mbaklim = "Money Back limit"
Const Fgamlim = "Free games limit"
Const Fspnlim = "Free spins limit"
Const Spnlim = "Spins limit"
Const Cashlim = "Cash limit"
Const Multln = "Multiline"
Private Sub Form_Load()
Dim errno As Long, loadnow As Long
Dim gamestartup As Long, seedchanged As Boolean, c As New cRegistry
Const vbBackslash = "\"
loadnow = False
justrestored = False
flashtoggle = False
spincount = 0
freegamecount = 0
gamestartup = 0
EOGEF = False
prizeaccold = 0
waitimmarker = 0
spinzstart = 0
activatctrls = True
changecalculated = True

Select Case gt(0)
Case Is < 1
gamestartup = 1
Case 1
gamestartup = 2
Case 2
gamestartup = 3
End Select


    If gamestartup = 1 Then

       'Slotdata.s$t exists in curdir?
       If gt(0) = 0 Then loadnow = findafile(CurDir, "Slotdata.s$t")

       If loadnow = 0 Then
       CommonDialog1.InitDir = CurDir
         CommonDialog1.Filter = "MyReels (Slotdata.s$t)|Slotdata.s$t|All Files (*.*)|*.*"
         ' Specify default filter
         CommonDialog1.FilterIndex = 1
         ' Set CancelError is True
         CommonDialog1.CancelError = True
         On Error GoTo Errcancel
         CommonDialog1.ShowOpen
           If CommonDialog1.FileTitle <> "Slotdata.s$t" Then
           response = MsgBox("Please respecify Slotdata.s$t!", vbOKOnly)
           Unload frmSplsh
           Stopnoise
           Exit Sub
           End If
        If gt(0) = -1 Then
        LoadFrmSplsh 440
        Else
        frmSplsh.Refresh
        End If
      End If

    End If  'gamestartup condition

loaddirectory = CurDir & "\"



'firstload, chdir or gametype cancel

        If gamestartup < 3 Then

        firstgametypeload

                If Inputvars = False Then
                If runAdmin = False Then response = MsgBox("Input file corrupted! Cannot continue.", vbOKOnly)
                Stopnoise
                Unload frmSplsh
                Exit Sub
                End If
        End If

PokeResolution Me


If gt(0) = -1 And gt(185) > 0 Then SndMidInit


gt(0) = 0
Pokemach.KeyPreview = False


'New start or restart game after accepting config
If gamestartup = 1 Or (gamestartup > 1 And gt(45) = 1) Or genoptsgen = True Then Randomisethem seedchanged


zerogamspnvars

getthumbspiccount thumbs, piccount, wheelvec, wheelorder

setinauxrouts piccount

gamspintot gamespintot

zeroscatter

degreeoftitle = gt(3)

If gt(200) = 0 And (gamestartup = 1 Or (gamestartup = 3 And gt(45) = 1)) Then
    If seedchanged = False And genoptsgen = False Then
    gt(55) = gt(55) + 1
    Else
    gt(56) = gt(56) + 1
    End If
    ElseIf gamestartup = 3 And genoptsgen = True Then
    gt(56) = gt(56) + 1
End If

If gt(55) = -1 Then gt(55) = 0   'fix for stats
If gt(56) = -1 Then gt(56) = 0   'fix for stats

'Get scatter vars, copy prizelists
For ct = 0 To 13
With imgprizethumb(ct)
.Left = resX * .Left
.Top = resY * .Top
.BorderStyle = 0
.Visible = False
End With
Next


If gt(159) = 1 Then
pw = 1560  'set sparepic width
For intreel = 0 To 4
For ct = 0 To 13
imgprizethumb(ct).Move 70 * resX, resY * (80 + ct * 600)
    If intreel < 3 Then
    wid = 600 * (intreel + 1)
    Else
    wid = 1800 + 480 * (intreel - 2)
    End If
lblprize(14 * intreel + ct).Move resX * wid, resY * (60 + ct * 600)
lblprizeamt(14 * intreel + ct).Move resX * wid, resY * (340 + ct * 600)
Next
Next
Else
pw = 1080
End If

randomspinvec pw, medfastslowmove, special, spinzstart

If gt(185) > 0 Then
playsuccess = PlayMidiFile(Stringvars(39), True)
MidiPlay.Enabled = playsuccess
If gt(187) > 0 Then medfastslowmove(0, 5) = 44
Else
playsuccess = False
End If


For pct = 1 To piccount
Scatterinit wheelvec(1, pct), pct, False
'order of intscattervec change?
Next

lbloffset = 0
errno = 0

If piccount = 14 Then   'ct temp here
ct = 13
Else
ct = 12
End If

If intscatternumber > 0 Then
With lblmisc(5)
.Visible = True
If gt(159) = 1 Then
.Left = resX * 960
If intscatternumber = 2 Then
.Top = resY * 7120
Else
.Top = resY * (7080 + 620 * -CLng(CBool(ct = 13)))
End If
End If
End With
End If


If intscatternumber = 2 Then

If gt(159) = 0 Then
imgprizethumb(12).Move resX * 120, resY * 7180
imgprizethumb(13).Move resX * 120, resY * 7980
Else
imgprizethumb(12).Top = resY * 7520
imgprizethumb(13).Top = resY * 8100
End If

For intreel = 0 To 4
  If intreel < 3 Then
  wid = resX * 600 * (intreel + 1)
  Else
  wid = resX * (1800 + 480 * (intreel - 2))
  End If
    
  If gt(159) = 0 Then
  lblprize(14 * intreel + 12).Move wid, resY * 7020
  lblprizeamt(14 * intreel + 12).Move wid, resY * 7500
  lblprize(14 * intreel + 13).Move wid, resY * 7820
  lblprizeamt(14 * intreel + 13).Move wid, resY * 8300
  Else
  lblprize(14 * intreel + 12).Move wid, resY * 7500
  lblprizeamt(14 * intreel + 12).Move wid, resY * 7780
  lblprize(14 * intreel + 13).Move wid, resY * 8080
  lblprizeamt(14 * intreel + 13).Move wid, resY * 8360
  End If

Next
ElseIf intscatternumber = 1 Then
If gt(159) = 0 Then
imgprizethumb(ct).Move resX * 120, resY * 7180
Else
imgprizethumb(ct).Top = resY * (7520 + 550 * -CLng(CBool(ct = 13)))
End If

For intreel = 0 To 4
    If intreel < 3 Then
    wid = resX * 600 * (intreel + 1)
    Else
    wid = resX * (1800 + 480 * (intreel - 2))
    End If
    If gt(159) = 0 Then

    lblprize(14 * intreel + ct).Move wid, resY * 7040
    lblprizeamt(14 * intreel + ct).Move wid, resY * 7520
    Else
    lblprize(14 * intreel + ct).Move wid, resY * (7500 + 550 * -CLng(CBool(ct = 13)))
    lblprizeamt(14 * intreel + ct).Move wid, resY * (7780 + 550 * -CLng(CBool(ct = 13)))
    End If
Next
End If


'Get pics, fix prize labels

For pct = 1 To piccount
sortlabels False, errno, pct
If errno = 1 Then GoTo ErrZymbols
Next

If gt(159) = 1 Then
With Line2(0)
.X1 = resX * 3350
.X2 = resX * 3350
.Y1 = resY * 2510
.Y2 = resY * 7043
.BorderWidth = 2
End With
With Line2(1)
.X1 = resX * 3350
.X2 = resX * 11005
.Y1 = resY * 7043
.Y2 = resY * 7043
.BorderWidth = 2
End With
With Line2(2)
.X1 = resX * 11005
.X2 = resX * 11005
.Y1 = resY * 2510
.Y2 = resY * 7043
.BorderWidth = 2
End With
With Line2(3)
.X1 = resX * 3350
.X2 = resX * 11005
.Y1 = resY * 2510
.Y2 = resY * 2510
.BorderWidth = 2
End With

For ct = 0 To 4
For ct1 = 0 To 3
With M(4 * ct + ct1)
.Width = resX * 1400
.Height = resY * 1400
.Top = resY * (ct1 - 1) * 1560
.Left = resX * ct * 1560
End With
Next
Next
Else
For ct = 0 To 3
With Line2(ct)
.X1 = resX * .X1
.X2 = resX * .X2
.Y1 = resY * .Y1
.Y2 = resY * .Y2
If resX > 1 Then
If (ct = 0 Or ct = 2) Then
.Y1 = .Y1 - 5
.Y2 = .Y2 + 10
End If
End If

End With
Next
For ct = 0 To 4
For ct1 = 0 To 3
With M(4 * ct + ct1)
.Width = resX * 1000
.Height = resY * 1000
.Top = resY * (ct1 - 1) * 1080
.Left = resX * ct * 1080
End With
Next
Next
End If


With Spare
If gt(159) = 0 Then
.Width = resX * 1000
.Height = resY * 1000
Else
.Width = resX * 1400
.Height = resY * 1400
End If
End With


'On gameload spinz is hidden picture (ct = 0), so bring it back 1 to kick of init
If gamestartup = 1 Or gt(45) = 1 Or genoptsgen = True Then
For intreel = 0 To 4
spinz(intreel) = Int(24 * Rnd + 1)
Next
    If Rnd < CSng(gt(6) / 4) Then
    dirofspin = -1
    Else
    dirofspin = 1
    End If

Else
    If prepspin = False Then
    For intreel = 0 To 4
    spinz(intreel) = Advanz(spinz(intreel), -dirspin(intreel))
    Next
    End If
End If

prepspin = False
genoptsgen = False


For intreel = 0 To 4

hreel(intreel) = False
dirspin(intreel) = dirofspin


Thumbslist(5).ListImages.Clear
For ct = 1 To 24
Thumbslist(5).ListImages.Add (ct), , thumbs(wheelorder(intreel, ct))
Spare.PaintPicture Thumbslist(5).ListImages(ct).Picture, 0, 0, resX * (pw - (gt(159) + 1) * 80), (resY * (pw - (gt(159) + 1) * 80))
Thumbslist(intreel).ListImages.Add (ct), , Spare.Image
Set Spare = Nothing
Next

spinztemp(intreel) = spinz(intreel)
spinztemp(intreel) = Advanz(spinztemp(intreel), 2 * dirofspin)
storit(intreel) = 0
trakker(intreel) = 3
For ct = 0 To 3
    picnum(intreel, 3 - ct) = 4 * intreel + ct
    With M(picnum(intreel, 3 - ct))
    If dirofspin = -1 Then
    pwScale = fixpw(3 - ct, pw)
    .Top = pwScale
    Set .Picture = Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), -dirofspin)).Picture
    If intreel = 0 Then ydisp0(ct) = pwScale
    Else
    pwScale = fixpw(ct - 1, pw)
    .Top = pwScale
    Set .Picture = Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), -dirofspin)).Picture
    If intreel = 0 Then ydisp0(ct) = pwScale
    End If
    End With
Next
Next


For ct = 0 To 47
With liwnf(ct)
.Visible = False
.BorderColor = gt(178)
If gt(159) = 1 Then
.X1 = resX * (.X1 + 1430)
.X2 = resX * (.X2 + 1430)
.Y1 = resY * (.Y1 - 6000)
.Y2 = resY * (.Y2 - 6000)
Else
.X1 = resX * .X1
.X2 = resX * .X2
.Y1 = resY * .Y1
.Y2 = resY * .Y2
End If
End With
Next

'Init object/font props
Tryagain:
On Error GoTo Fontprob
textwidthratio = Spare.TextWidth(Stringvars(10))
Spare.fontname = Stringvars(10)
textwidthratio = textwidthratio / Spare.TextWidth(Stringvars(10))
If textwidthratio < 0.72 Then textwidthratio = 0.72
For ct = 0 To 1
With lblprizemeter(ct)
.fontname = Stringvars(10)
.Fontsize = resY * Int(14 * textwidthratio)
.Font.Charset = 0
.Font.Weight = 600
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0
.ForeColor = gt(175)
End With
With lblrandom(ct)
.ForeColor = gt(175)
.fontname = Stringvars(10)
.Fontsize = Int(14 * textwidthratio)
.Font.Charset = 0
.Font.Weight = 600
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0
End With
Next


lblmisc(2).Width = resX * lblmisc(2).Width
For ct = 0 To 9
With lblmisc(ct)
If ct = 7 Then
.ForeColor = gt(175)
Else
.ForeColor = gt(163)
End If
.fontname = Stringvars(10)
If ct < 8 Then
.Fontsize = resY * Int(14 * textwidthratio)
Else
.Fontsize = resY * Int(8 * textwidthratio)
End If
.Font.Charset = 0
.Font.Weight = 700
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0
.Height = resY * .Height
End With

Next


For ct = 0 To 4
For ct1 = 0 To 13

With lblprize(ct * 14 + ct1)
.fontname = Stringvars(10)
.Fontsize = CLng(12 * textwidthratio)
.Font.Charset = 0
.Font.Weight = 600
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0
.BackColor = gt(165 + ct)
.ForeColor = gt(170 + ct)
If ct < 2 Then
.Caption = "X " & 5 - ct
Else
.Caption = "X" & 5 - ct
End If
End With

'lblprizeamt forecolour as above
With lblprizeamt(ct * 14 + ct1)
.fontname = Stringvars(10)
.Fontsize = resY * Int(10 * textwidthratio)
.Font.Charset = 0
.Font.Weight = 600
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0
.BackStyle = 0
.ForeColor = gt(170 + ct)
If ct < 2 And resX > 1 Then .Width = 220 + .Width
End With
Next
Next

special = 0
freegamecount = 0
spincount = 0
prizeaccum = gt(2)


cumpzzeros = ""
For ct1 = 1 To Len(CStr(gt(47))) - Len(CStr(prizeaccum))
cumpzzeros = "0" & cumpzzeros
Next
lblprizemeter(0).Caption = cumpzzeros & CStr(prizeaccum)

pzzeros = ""

For ct1 = 1 To Len(CStr(gt(47)))
pzzeros = "0" & pzzeros
Next
lblprizemeter(1).Caption = pzzeros

betqut = gt(20 + gt(20))
With lblmisc(0)
If .Height < 330 * resY Then
.Top = .Top + (330 * resY - .Height) / 2
lblprizemeter(1).Top = lblprizemeter(1).Top + (330 * resY - .Height) / 2
End If
.Caption = "BET " & betqut & "X"
End With
indicln


For ct = 0 To 1
With Candy(ct)
.BackPicture = gt(180 + 2 * ct)
.Fontsize = textwidthratio * 12
.fontname = Stringvars(11 + ct)
.ForeColor = gt(179 + 2 * ct)
.Caption = Stringvars(ct + 6)
.Enabled = False
End With
Next


'money back
If gt(10) > 0 Then

lblmisc(1).Visible = True
lblmisc(7).Visible = True
lblmisc(7).Caption = gt(13)
For ct = 1 To gt(10)

With lblmoneyback(ct - 1)
If gt(159) = 1 Then
.Left = resX * 11540
.Top = resY * (ct - 1) * 480
Else
.Left = resX * .Left
.Top = resY * .Top
End If
.AutoSize = True
.fontname = Stringvars(10)
.Fontsize = Int(14 * textwidthratio)
.Font.Charset = 0
.Font.Weight = 700
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0

.Caption = CStr(ct)
.Alignment = 2  'Center
.BackStyle = 0  'Transparent
.ForeColor = gt(163)
.Visible = True
If gt(11) = ct Then .ForeColor = gt(177)

End With

Next
End If


If gt(14) = 1 Then     'Activate Random jackpot

For ct = 2 To 3
With lblmisc(ct)
.Top = resY * .Top
.Left = resX * .Left
.Visible = True
End With
Next

lblmisc(3).Caption = decsep

With lblrandom(0)
.Caption = ""
.Visible = True
For ct1 = Len(CStr(gt(16))) To Len(CStr(gt(47))) - 1
.Caption = "0" & .Caption
Next
.Caption = .Caption & CStr(gt(16))
End With
With lblrandom(1)
.Visible = True
.Caption = gt(17)
End With
End If


'Quotes
For ct = 0 To 3
With lblmisc(10 + ct)
.Top = resY * .Top
.fontname = Stringvars(10)
.Fontsize = Int(resY * 11.25 * textwidthratio)
.Font.Charset = 0
.Font.Weight = 300
.FontUnderline = 0
.FontItalic = 0
.FontStrikethru = 0
.ForeColor = gt(163)
If quotestring(ct) <> "" Then
.Caption = quotestring(ct)
Else
.Visible = False
End If
End With
Next


'Now set the tiled background

If gt(161) = -1 Then

If loadbackbitmap("Background", 1) = True Then


wid = Spare.Picture.Width
hgt = Spare.Picture.Height

For ct = 0 To Pokemach.ScaleWidth Step 51 * wid / 90
For ct1 = 0 To Pokemach.ScaleHeight Step 51 * hgt / 90
Pokemach.PaintPicture Spare.Picture, ct, ct1
Next
Next
End If


Else

Pokemach.BackColor = gt(161)

End If  'background bitmap condition


'merge background

With TA
.Width = resX * .Width
.Height = resY * .Height
If gt(159) = 1 Then
.Left = resX * 4540
.Top = resY * 800
Else
.Left = resX * .Left
.Top = resY * .Top
If gt(14) = 0 And gt(10) = 0 Then .Top = resY * 1050
End If
'OPTION OF ANOTHER BITMAP HERE
If gt(162) = -1 Then
  If Stringvars(1) = Stringvars(2) Then
  .Visible = False
  procend = True
  Unload frmSplsh
  waitimmarker = 5
  waitimer.Enabled = True
  Exit Sub
  Else
  If loadbackbitmap("Title", 2) = True Then paintitlearea
  procend = True
  Unload frmSplsh
  End If
Else
  procend = True
  Unload frmSplsh
  .BackColor = gt(162)
End If
TextCircle TA, Stringvars(5), .ScaleWidth / 2, .ScaleHeight, .ScaleHeight, degreeoftitle, .Fontsize
End With

If waitimmarker = 0 Then
waitimmarker = 6
waitimer.Enabled = True
End If

Exit Sub

Fontprob:
response = MsgBox("Problem with Fonts ; load game defaults?", vbOKCancel)
If response = 2 Then
Stopnoise
Unload frmSplsh
Exit Sub
Else
For ct = 9 To 12
Stringvars(ct) = "Times New Roman"
Next
Err.Clear
GoTo Tryagain
End If
Errcancel:
'User clicked Cancel
response = MsgBox("Request to open Slotdata.s$t cancelled!", vbOKOnly)
Stopnoise
Unload frmSplsh
Exit Sub
ErrZymbols:
response = MsgBox("Cannot locate nn.bmp; Aborting load!", vbOKOnly)
cleanup
Unload frmSplsh
Stopnoise
Exit Sub
End Sub
Private Function loadbackbitmap(zTitle As String, stringvarindex As Long)
loadbackbitmap = False
On Error GoTo Errbitmap
Set Spare.Picture = LoadPicture(Stringvars(stringvarindex))
loadbackbitmap = True
Exit Function
Errbitmap:
response = MsgBox("There is a problem loading the " & zTitle & " Bitmap. Continuing to load with default colour ...", vbOKOnly)
If stringvarindex = 1 Then
gt(161) = &H80FF&
Pokemach.BackColor = gt(161)
End If
Err.Clear
End Function
Public Sub indicln()
'Line indicator
For ct = 1 To 6
If ct > 2 * (gt(153) + 1) Then
Lnptr(ct - 1).FillColor = &H80FF&
Else
Lnptr(ct - 1).FillColor = &HFFFF00
End If
Next
lnbet = betqut * (gt(153) + 1)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
keyPrsd = True
If zhiddnstatus > 9 Then
Zhidden.Show
Else
If Prizemeter.Enabled = True Then
  If KeyCode = vbKeyEscape Then
  lblprizemeter(0).Caption = cumpzzeros & CStr(prizeaccum - mt)
  DoPrz
  LinesBegone
  End If
Exit Sub
ElseIf timoneyback.Enabled = True Or jackpot.Enabled = True Then
  If KeyCode = vbKeyEscape Then
  ct3 = 14
  End If
Exit Sub
End If

If spincount > 0 Or freegamecount > 0 Or waitimer.Enabled = True Or waitimmarker > 0 Then Exit Sub

Select Case KeyCode
Case vbKey1
PlaySndF Stringvars(27)
gt(153) = 0
indicln
Case vbKey2
PlaySndF Stringvars(27)
gt(153) = 1
indicln
Case vbKey3
PlaySndF Stringvars(27)
gt(153) = 2
indicln
Case vbKeyA
Frmaboutt
Case vbKeyC
chgchtr True
Preparetospin
Case vbKeyD
Changedir
Case vbKeyM
Zhidden.Musicc_Click
Case vbKeyQ
Zhidden.Thumbb_Click
Case vbKeyS
Zhidden.Soundd_Click
Case vbKeyT
TA_Click
Case vbKeyControl
Cyclemultibets
activatectrls
Case vbKeyReturn
Configurationn
Case vbKeySpace
chgchtr
Preparetospin
Case vbKeyShift
PopupMenu Zhidden!OptionZ, 2
Case vbKeyEscape
Quitt
Case vbKeyF1
PlaySndF App.Path & "\help.wav"
End Select
End If
End Sub
Private Sub candy_Click(Index As Integer)
Select Case Index
Case 0
chgchtr
Preparetospin
Case 1
Cyclemultibets
activatectrls
End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
keyPrsd = False
If zhiddnstatus > 9 Then
If Button = 2 Then Exit Sub
Zhidden.Show
Else
If Button = 2 Then
gt(0) = 0
 If Prizemeter.Enabled = True Then
 lblprizemeter(0).Caption = cumpzzeros & CStr(prizeaccum - mt)
 DoPrz
 LinesBegone
 End If
Prizeflash.Enabled = False
PopupMenu Zhidden!OptionZ, 2
End If
End If
End Sub
Private Sub Form_Activate()
If zhiddnstatus > 9 Then Zhidden.Show
End Sub
Private Sub imgprizethumb_Click(Index As Integer)
Unload Zhidden
Set Zhidden = Nothing
zhiddnstatus = 10 + Index
Load Zhidden
End Sub
Private Sub MidiPlay_Timer()

If medfastslowmove(0, 5) > 0 And waitimmarker <> 1 Then playsuccess = PlayMidiFile(Stringvars(medfastslowmove(0, 5)), False)
If playsuccess = True Then

  If gt(187) > 0 Or (gt(187) < 0 And (waitimer.Enabled = gamespinwait.Enabled = Prizemeter.Enabled = timoneyback.Enabled = jackpot.Enabled = False)) Then  'set next track

    If gt(187) > 0 Then
     gt(187) = gt(187) + 1
     If gt(187) > midisNum Then gt(187) = 1
     medfastslowmove(0, 5) = 43 + gt(187)
    Else

     gt(187) = gt(187) - 1
     If Abs(gt(187)) > midisNum Then gt(187) = -1

      If (oldMidiNo = medfastslowmove(0, 5)) Then
       If medfastslowmove(0, 5) < 50 Then
       medfastslowmove(0, 5) = oldMidiNo - gt(187)
       Else
       medfastslowmove(0, 5) = 44
       End If
      Else
      oldMidiNo = medfastslowmove(0, 5)
      End If
    
    End If

  playsuccess = False
  End If
End If
End Sub
Public Sub TA_Click()
With TA
If waitimer.Enabled = True Then Exit Sub
If gt(162) = -1 Then
        If Stringvars(1) = Stringvars(2) Then
        .Visible = False
        Pokemach.KeyPreview = False
        For ct2 = 0 To 1
        Candy(ct2).Enabled = False
        Next
        Pokemach.Show
        activatctrls = True
        waitimmarker = 5
        waitimer.Interval = 2 * gt(194)
        waitimer.Enabled = True
        Else
        paintitlearea
        End If
Else
.Cls
End If
degreeoftitle = degreeoftitle + 1
If degreeoftitle = 6 Then degreeoftitle = 0
TextCircle TA, Stringvars(5), .ScaleWidth / 2, .ScaleHeight, .ScaleHeight, degreeoftitle / 2, .Fontsize
End With
gt(3) = degreeoftitle
End Sub
Private Sub Form_DblClick()
activatctrls = True
PopupMenu Zhidden!OptionZ, 2
End Sub
Private Sub Form_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
activatctrls = True
PopupMenu Zhidden!OptionZ, 2
End If
End Sub
Private Sub Quotez_Click()
PlaySndF ("Thumbb")
End Sub
Public Sub Preparetospin()
Dim RH As Long, LH As Long, gendirchange As Boolean
Prizeflash.Enabled = False
TA.Enabled = False
Pokemach.KeyPreview = False
Candy(0).Enabled = False
Candy(1).Enabled = False
activatctrls = True
gendirchange = False
jackpotprize = 0

indicln
lblprizemeter(1).ToolTipText = ""

If lnbet > prizeaccum Then  'hardly any money left
If prizeaccum > 0 Then
lblmisc(4).Caption = "Bet smaller please"
lblmisc(4).ToolTipText = "Bet smaller or start a new game"
activatectrls
Else
lblmisc(4).Caption = "No Money Left!"
showendofgame
End If
Exit Sub

End If

lblmisc(4).Caption = ""
lblmisc(4).ToolTipText = ""

If justcached = False Then
'generate number for spin up or down
If Rnd < CSng(gt(6) / 4) Then
If dirofspin = 1 Then gendirchange = True
dirofspin = -1
Else
If dirofspin = -1 Then gendirchange = True
dirofspin = 1
End If
Else
'Don't change spindir
justcached = False
End If


On Error GoTo ErrOverflowspins
If dirofspin = -1 Then gt(52) = gt(52) + 1

If gt(52) > gt(47) Then GoTo ErrOverflowspins



'Holdq is free spin hold counter
'Spinz is top line symbols (win1)

If gamesaved = True And spinsaved = True Then
lblmisc(4).Caption = "Game(s); Spin(s) saved! "
ElseIf gamesaved = True Then lblmisc(4).Caption = "Game(s) saved! "
ElseIf spinsaved = True Then lblmisc(4).Caption = "Spin(s) saved! "
End If

If freegamecount > 0 Then
       
        If freegamecount = 1 Then

        arrangespins gendirchange


        If justrestored = True Then
        lblmisc(4).Caption = lblmisc(4).Caption & "Games restored!"
        ElseIf featurereset = True Then
        lblmisc(4).Caption = lblmisc(4).Caption & "Feature reset!"
        End If
            
        lblmisc(8).Caption = "Feature multiplier : " & freegamesettings(kept1or2, 6) & "X"
        lblmisc(8).ForeColor = gt(175)
        
        waitimmarker = 3
        waitimer.Interval = gt(194) * 240
        waitimer.Enabled = True
        End If

lblmisc(4).Caption = lblmisc(4).Caption & " Game " & freegamecount & " of " & freegamesettings(kept1or2, 8)
freegamecount = freegamecount + 1

On Error GoTo ErrOverflowfrgam
gt(49) = gt(49) + 1       'stats
If gt(49) > gt(47) Then GoTo ErrOverflowfrgam

ElseIf spincount > 0 Then

    If spincount = 1 Then
    
        arrangespins gendirchange

        If justrestored = True Then
        lblmisc(4).Caption = lblmisc(4).Caption & "Spins restored!"
        ElseIf featurereset = True Then
        lblmisc(4).Caption = lblmisc(4).Caption & "Feature reset!"
        End If
        
        lblmisc(8).Caption = "Feature multiplier : " & spinsettings(kept1or2, 12) & "X"
        lblmisc(8).ForeColor = gt(178)
        
        waitimmarker = 3
        waitimer.Interval = gt(194) * 240
        waitimer.Enabled = True
        End If


lblmisc(4).Caption = lblmisc(4).Caption & " Spin " & spincount & " of " & spinsettings(kept1or2, 14)
spincount = spincount + 1


On Error GoTo ErrOverflowfrspin
If gt(50) > gt(47) Then GoTo ErrOverflowfrspin
gt(50) = gt(50) + 1       'stats

Else    'No FSFG

gt80 = False
gt76 = False
gt72 = False
gt132 = False

For intreel = 0 To 4
hreel(intreel) = False
dirspin(intreel) = dirofspin
Next

lblmisc(8).Visible = False

If chtr <= 0 Then
On Error GoTo ErrOverflowbetmult    'stats
gt(133) = gt(133) + betqut * (gt(153) + 1)
If gt(133) > gt(47) Then GoTo ErrOverflowbetmult

On Error GoTo ErrOverflowspins
If gt(51) > gt(47) Then GoTo ErrOverflowspins
gt(51) = gt(51) + 1       'stats

prizeaccum = prizeaccum - lnbet
lblmisc(4).Caption = Stringvars(14)
Else
lblmisc(4).Caption = "Go Cheater!"
End If

End If



If dirofspin = -1 Then  'spinning up
    If gendirchange = False Then GoTo Clearforspin
    For intreel = 0 To 4
    If hreel(intreel) = False Then
    dirspin(intreel) = dirofspin
    trakker(intreel) = 3
    
    For ct = 0 To 3
    picnum(intreel, 3 - ct) = 4 * intreel + ct
    Next
    For ct = 0 To 3
    spinztemp(intreel) = spinz(intreel) 'advanz
    pwScale = fixpw(3 - ct, pw)
    M(picnum(intreel, ct)).Top = pwScale
    Set M(picnum(intreel, ct)).Picture = Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), ct - 3)).Picture
    ydisp0(ct) = pwScale
    Next
    End If
    Next
   
Else
    If gendirchange = False Then GoTo Clearforspin
    For intreel = 0 To 4
    
    If hreel(intreel) = False Then
    dirspin(intreel) = dirofspin
    trakker(intreel) = 3
    
    For ct = 0 To 3
    picnum(intreel, ct) = 4 * intreel + ct
    Next
    
    For ct = 0 To 3
    spinztemp(intreel) = spinz(intreel) 'advanz
    pwScale = fixpw(ct - 1, pw)
    M(picnum(intreel, 3 - ct)).Top = pwScale
    Set M(picnum(intreel, 3 - ct)).Picture = Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), 3 - ct)).Picture
    ydisp0(ct) = pwScale
    Next
    End If
    Next
    
End If


Clearforspin:

'release held reels
If freegamecount > 0 Then
For intreel = 0 To 4
hreel(intreel) = False
Next
End If

pwScale = fixpw(1, pw)

If dirofspin = 1 Then
ydispmin = -pwScale
ydispmax = 0
Else
ydispmin = 3 * pwScale
ydispmax = 2 * pwScale
End If




'Total spins <= gt(47)
On Error GoTo ErrOverflowspins
If gt(49) + gt(50) + gt(51) > gt(47) Then GoTo ErrOverflowspins


With lblprizemeter(1)
If .Caption > 0 Then  'Don't need to redisplay same amount

.Caption = ""
For ct1 = 1 To Len(CStr(gt(47)))
.Caption = "0" & lblprizemeter(1).Caption
Next
End If
End With

With lblprizemeter(0)
.Caption = ""
For ct1 = Len(CStr(prizeaccum)) + 1 To Len(CStr(gt(47)))
.Caption = "0" & .Caption
Next
.Caption = .Caption & CStr(prizeaccum)
End With



Pokemach.AutoRedraw = True


For ct1 = 0 To 47
liwnf(ct1).Visible = False
Next
lblmisc(0).ForeColor = gt(163)
lblmisc(0).Caption = "BET " & betqut & "X"
lblmisc(9).Visible = False
aptot = 0

'Reset prize lights

For ct = 1 To pzctcum
For ct1 = 0 To 70 Step 14
If oldlblprize(ct) < ct1 Then
lblprize(oldlblprize(ct)).ForeColor = gt(170 + Int(oldlblprize(ct) / 14))
lblprizeamt(oldlblprize(ct)).BackStyle = 0
lblprizeamt(oldlblprize(ct)).ForeColor = gt(170 + Int(oldlblprize(ct) / 14))
Exit For
End If
Next
Next


If gt(10) > 0 And spincount = 0 And freegamecount = 0 And chtr <= 0 Then
'Append new Stats to old
If gt(11) > 0 Then gt(88 + gt(11)) = CLng(gt(88 + gt(11)) & "0" & lnbet) 'Money back light off if reset
If (gt(12) = 0 And gt(11) = 1) Or (gt(12) > 0 And gt(11) = 0) Then lblmoneyback(gt(10) - 1).ForeColor = gt(163)
End If


'Random jackpot with no overflow

If gt(14) = 1 Then
LH = gt(16)
RH = gt(17)

If chtr <= 0 Then
gt(48) = gt(48) + lnbet    'stats
gt(128) = gt(128) + 1
'Adjust RH increment
RH = RH + gt(15) * lnbet
End If  'chtr
lblrandom(0).Caption = ""
        If RH > 9 Then
        On Error GoTo Erroverflowjackinc
        LH = LH + Int(RH / 10)
        If LH > gt(47) Then GoTo Erroverflowjackinc
        RH = RH - 10 * Int(RH / 10)
        End If
For ct1 = Len(CStr(LH)) To Len(CStr(gt(47))) - 1
lblrandom(0).Caption = "0" & lblrandom(0).Caption
Next
lblrandom(0).Caption = lblrandom(0).Caption & CStr(LH)
lblrandom(1).Caption = RH
For ct1 = 0 To 1
lblmisc(2).Caption = "Random Jackpot now at ..."
lblmisc(ct1 + 2).ForeColor = gt(163)
lblrandom(ct1).ForeColor = gt(175)
Next
gt(16) = LH
gt(17) = RH
End If


prepspin = True
For ct = 10 To 13
lblmisc(ct).Caption = ""
Next
randomspinvec pw, medfastslowmove, special, spinzstart
For ct = 0 To 3
With lblmisc(10 + ct)
If quotestring(ct) <> "" Then
.Caption = quotestring(ct)
.Visible = True
Else
.Visible = False
End If
End With
Next



Pokemach.AutoRedraw = False
PlaySndF Stringvars(26)

'if waitimmarker = 3, wait while old spin/games are restored
If waitimmarker = 0 Then
waitimer.Interval = 2 * gt(194)
waitimmarker = 1
waitimer.Enabled = True
End If

Exit Sub
Erroverflowjackinc:
reachedlimit Jacklim
Exit Sub
ErrOverflowbetmult:
reachedlimit Betlim
Exit Sub
ErrOverflowfrgam:
gt(49) = gt(47)
reachedlimit Fgamlim
Exit Sub
ErrOverflowfrspin:
gt(50) = gt(47)
reachedlimit Fspnlim
Exit Sub
ErrOverflowspins:
If gt(49) > 0 Then
gt(49) = gt(49) - 1
ElseIf gt(50) > 0 Then
gt(50) = gt(50) - 1
Else
gt(51) = gt(47)
End If
reachedlimit Spnlim
End Sub
Private Sub Vanishlines_Timer()
LinesBegone
End Sub
Private Sub LinesBegone()
If Quotez.Picture.Handle <> 0 Then
For ct1 = 0 To 47
liwnf(ct1).Visible = False
Next
Quotez.Visible = True
End If
Vanishlines.Enabled = False
End Sub
Private Sub waitimer_Timer()
Select Case waitimmarker

Case 0  'Put waitimer to sleep

If freegamecount = 0 And spincount = 0 Then
'Time out buffered keystrokes
If Pokemach.KeyPreview = True Then waitimer.Enabled = False
    If timoneyback.Enabled = False Then
    activatectrls
    Vanishlines.Enabled = True
    End If

Else
waitimer.Enabled = False
End If

Case 1
waitimer.Enabled = False

spindereels  'Called from here or Preparetospin ... set up pay & show prize

'Snapshot reels for holdqfreegame purpose
For currline = 0 To gt(153)
For intreel = 0 To 4    'Top line
If hreel(intreel) = False Then hq(currline, intreel) = spinz(intreel)
Next
Next

sortsubstitutes prizecount, multiprize

displayprize


waitimmarker = 2

Case 2  'called from prizemeter timer

eligiblespingame freegamecount, spincount, kept1or2, prizetotal, gamenotspin, gamesaved, spinsaved, featurereset

If gt(10) > 0 And EOGEF = False And chtr <= 0 Then
If freegamecount = 1 Or spincount = 1 And justrestored = False And featurereset = False Then prizeaccold = prizeaccum - prizetotal - jackpotprize
'prizeaccold now set

moneybackdrama jackpotprize + prizetotal   'Money back
    If aptot > 0 Then
    prizeonmeter aptot
    Else
    If jackpotprize = 0 And Quotez.Picture.Handle <> 0 Then PlaySndF ("Thumbb")
    End If
Else
If jackpotprize = 0 And Quotez.Picture.Handle <> 0 Then PlaySndF ("Thumbb")
End If

If freegamecount > 0 Or spincount > 0 Then

    lblmisc(8).Visible = True
    justrestored = False
    
        If freegamecount > 0 Then
        lblmisc(8).Caption = "Feature multiplier : " & freegamesettings(kept1or2, 6) & "X"
        lblmisc(8).ForeColor = gt(175)
        ElseIf spincount > 0 Then
        lblmisc(8).Caption = "Feature multiplier : " & spinsettings(kept1or2, 12) & "X"
        lblmisc(8).ForeColor = gt(178)
        End If

    Candy(0).Enabled = False
    Candy(1).Enabled = False

    waitimer.Interval = (28 * gt(194) * gt(157)) + 360
    gamespinwait.Enabled = True

End If
waitimmarker = 0
Case 3  'Called from Preparetospin when game/spins active
waitimmarker = 1
waitimer.Interval = 2 * gt(194)
Case 4  'MB if FSFG nowin
If Prizemeter.Enabled = False Then waitimmarker = 0
Case 5  'Merge BG bitmaps on LOAD
With TA
Set .Picture = CaptureWindow(.hWnd, False, 0, 0, .ScaleX(.Width, vbTwips, vbPixels), .ScaleY(.Height, vbTwips, vbPixels))
.Visible = True
If gt(156) = 1 Or prizeaccum = 0 Then
TextCircle TA, "End_of_Game", .ScaleWidth / 2, .ScaleHeight, .ScaleHeight, degreeoftitle, .Fontsize
lblmisc(4).ToolTipText = EOGTT
For ct = 0 To 1
Candy(ct).ToolTipText = EOGTT
Next
prizeaccum = 0
Else
TextCircle TA, Stringvars(5), .ScaleWidth / 2, .ScaleHeight, .ScaleHeight, degreeoftitle, .Fontsize
End If
End With
waitimmarker = 6
Case 6
If TA.ToolTipText = "" Then
If gt(156) <> 1 And prizeaccum <> 0 Then Cachereels
lblmisc(4) = Stringvars(13)     'Welcome
TA.ToolTipText = "Click to Adjust"
Pokemach.Show
activatectrls
Else
    If gt(156) = 1 Then
    Pokemach.Refresh
    activatectrls
    End If
End If
waitimmarker = 0
End Select
End Sub
Private Sub gamespinwait_Timer()
Dim test1or2 As Long
If waitimer.Enabled = True Or jackpot.Enabled = True Or timoneyback.Enabled = True Then Exit Sub
gamespinwait.Enabled = False
waitimmarker = 4


For test1or2 = 0 To 1
If test1or2 = kept1or2 Then
        If freegamecount > freegamesettings(test1or2, 8) Then
                If clearq(kept1or2, freegamecount, spincount, gamenotspin) = True Then
                fixmonback
                Exit Sub
                Else    'do other queues
                justrestored = True
                End If
        ElseIf spincount > spinsettings(test1or2, 14) Then
        currold = currheld
                If clearq(kept1or2, freegamecount, spincount, gamenotspin) = True Then
                fixmonback
                Exit Sub
                Else
                justrestored = True
                End If
        End If
waitimmarker = 0
Preparetospin
Exit Sub        'In case of unwanted second loop pass
End If
Next
End Sub
Private Sub spindereels()
Dim spd1 As Long, spd2 As Long, spd3 As Long
Dim dur1 As Long, dur2 As Long, dur3 As Long


If Rnd < CSng(gt(7) / 4) Then
gt(53) = gt(53) + 1       'stats
dur1 = gt(194) * 2
dur2 = gt(194) * 4
dur3 = gt(194) * 10
Else
dur1 = gt(194) * 6
dur2 = gt(194) * 10
dur3 = gt(194) * 14
End If


If Rnd < CSng(gt(8) / 4) Then
gt(54) = gt(54) + 1       'stats
spd1 = 1
spd2 = 2
spd3 = 3
Else
spd1 = 2
spd2 = 3
spd3 = 4
End If

'Now the spinning
reelmin = 0
reelmax = 0
For reelmax = 0 To 4
movesize(reelmax) = medfastslowmove(3, reelmax)
For ct = 1 To medfastslowmove(0, reelmax)
spinning spd3
Next
Next
reelmax = 4
For ct = 1 To dur1
spinning spd1
Next
For ct = 1 To dur2
spinning spd2
Next
For ct = 1 To dur1
spinning spd1
Next
For reelmin = 0 To 4


If chtr > 0 Then
For ct = 1 To 999
spinztemp(reelmin) = spinz(reelmin)
If wheelorder(reelmin, Advanz(spinztemp(reelmin), -dirofspin)) = chtr Then Exit For
spinning spd1
Next

End If

For ct = 1 To dur3
spinstop = False
endspin spd3
If spinstop = True Then Exit For
Next
Next

End Sub
Private Sub spinning(movechoice)
For intreel = reelmin To reelmax

If hreel(intreel) = False Then

storit(intreel) = storit(intreel) + 1

For ct1 = 0 To 3
ydisp(intreel, ct1) = ydisp0(ct1) + dirofspin * storit(intreel) * movesize(intreel)
M(picnum(intreel, ct1)).Top = ydisp(intreel, ct1)
Next

'cleanup
If Abs(ydisp(intreel, 0) - ydispmax) < TOL Then
spinz(intreel) = Advanz(spinz(intreel), dirofspin)
M(picnum(intreel, 3)).Top = ydispmin
Set M(picnum(intreel, 3)).Picture = Thumbslist(intreel).ListImages(spinz(intreel)).Picture
trakker(intreel) = trakker(intreel) + 1
If trakker(intreel) = 4 Then trakker(intreel) = 0
For ct1 = 0 To 3
picnum(intreel, ct1) = 4 * intreel + Cycle(trakker(intreel), ct1)
Next
storit(intreel) = 0
spinzstart = Advanz(spinzstart, 1)
spinzstart = 1
movesize(intreel) = medfastslowmove(movechoice, spinzstart)
End If

End If  'holdq
Next
End Sub
Private Sub endspin(movechoice)
For intreel = reelmin To reelmax

If hreel(intreel) = False Then

storit(intreel) = storit(intreel) + 1

For ct1 = 0 To 3
ydisp(intreel, ct1) = ydisp0(ct1) + dirofspin * storit(intreel) * movesize(intreel)
M(picnum(intreel, ct1)).Top = ydisp(intreel, ct1)
Next
 
'cleanup
If Abs(ydisp(intreel, 0) - ydispmax) < TOL Then
If intreel = reelmin Then
        If Int((special + 1) / 2) = 1 Then
        special = oldspecial    'restore input value
        spinstop = True
        spinzstart = Advanz(spinzstart, 1)
        trakker(intreel) = trakker(intreel) + 1
        If trakker(intreel) = 4 Then trakker(intreel) = 0
        For ct1 = 0 To 3
        picnum(intreel, ct1) = 4 * intreel + Cycle(trakker(intreel), ct1)
        Next
        Else    'do another rotation, happens when special is 0 as well!
        oldspecial = special
        special = Int(2 * Rnd + 1)
        End If
Else
spinz(intreel) = Advanz(spinz(intreel), dirofspin)
M(picnum(intreel, 3)).Top = ydispmin
Set M(picnum(intreel, 3)).Picture = Thumbslist(intreel).ListImages(spinz(intreel)).Picture
trakker(intreel) = trakker(intreel) + 1
If trakker(intreel) = 4 Then trakker(intreel) = 0
For ct1 = 0 To 3
picnum(intreel, ct1) = 4 * intreel + Cycle(trakker(intreel), ct1)
Next
spinzstart = Advanz(spinzstart, 1)
movesize(intreel) = medfastslowmove(movechoice, spinzstart)
End If
storit(intreel) = 0
End If


End If  'holdq condition
Next    'intreel loop
End Sub
Private Sub displayprize()
Dim currprize As Long, picno As Long, multtemp As Long
Dim capnats As Boolean, capsubs As Boolean, pzcum As Long, pztool(2) As String, lineoferror As Long
pzctcum = 0
lbloffset = 0
prizetotal = 0
capnats = False
capsubs = False

If chtr > 0 Then
prizecount(1) = 0
prizecount(2) = 0
End If

For currline = 0 To gt(153)
pzcum = 0
pztool(currline) = ""
lineoferror = currline
ct1 = 0
ct2 = 0

'Pay
    For ct = 1 To prizecount(currline)
        picno = multiprize(currline, ct, 2)
        ct2 = sst(picno, multiprize(currline, ct, 1))
            If multiprize(currline, ct, 3) = 1 Then
            lblmisc(9).Visible = True
                If gt(4) > 0 Then
                capsubs = True
                multtemp = gt(4)
                Else    'gt(5) > 0
                capnats = True
                multtemp = gt(5)
                End If
        pztool(currline) = pztool(currline) & " Pic " & picno & ": " & multtemp & " X (" & betqut * ct2 & ") "
        currprize = multtemp * ct2
        Else    'not in a substitute family
        multtemp = 1
        pztool(currline) = pztool(currline) & " Pic " & picno & ": (" & betqut * ct2 & ") "
        currprize = ct2
        End If
        sortlabels True, ct, picno
        pzcum = currprize + pzcum
        multiprize(currline, ct, 3) = multtemp
    Next


'FSFG bonus
ct2 = 1
If gt(184) = 0 Or (currline = currheld And gt(184) = 1) Then
If pzcum > 0 Then
gt(134) = gt(134) + betqut 'stats
    If spincount > 0 Then
        If hq(currheld, 5) = gamespinsymbol(2) Then
        ct2 = spinsettings(0, 12)
        Else
        ct2 = spinsettings(1, 12)
        End If
    ElseIf freegamecount > 0 Then
        If hq(currheld, 5) = gamespinsymbol(0) Then
        ct2 = freegamesettings(0, 6)
        Else
        ct2 = freegamesettings(1, 6)
        End If
    End If
pzcum = ct2 * pzcum
Else
'FSFG nopay
    If spincount > 0 Then
        If hq(currheld, 5) = gamespinsymbol(2) Then
        ct1 = spinsettings(0, 13)
        Else
        ct1 = spinsettings(1, 13)
        End If
    ElseIf freegamecount > 0 Then
        If hq(currheld, 5) = gamespinsymbol(0) Then
        ct1 = freegamesettings(0, 7)
        Else
        ct1 = freegamesettings(1, 7)
        End If
    End If
    If ct1 > 0 Then
    pztool(currheld) = "No-Prize feature bonus : (" & ct1 * betqut & ")"
    gt(130) = gt(130) + 1   'stats
    gt(131) = gt(131) + ct1
    If gt132 = False Then gt(132) = gt(132) + betqut
    gt132 = True
    Else
    pztool(currline) = "0 "
    End If
End If
Else
If pzcum = 0 Then pztool(currline) = "0 "
End If


'STATS
If chtr <= 0 Then
For ct = 1 To prizecount(currline)
currprize = betqut * ct2 * multiprize(currline, ct, 3) * sst(multiprize(currline, ct, 2), multiprize(currline, ct, 1))
'Substitutes & scatters mut. exclusive
If multiprize(currline, ct, 0) = 1 Then  'subs stats

        If spincount > 0 Or freegamecount > 0 Then
        If currprize >= gt(73) Then
        gt(73) = currprize
        gt(74) = gt(49) + gt(50) + gt(51)
        End If
        On Error GoTo ErrOverflowsubs
        gt(75) = gt(75) + currprize
        If gt76 = False Then gt(76) = gt(76) + betqut
        gt76 = True
        gt(126) = gt(126) + 1
        Else
        If currprize >= gt(61) Then
        gt(61) = currprize
        gt(62) = gt(49) + gt(50) + gt(51)
        End If
        On Error GoTo ErrOverflowsubs
        gt(63) = gt(63) + currprize
        gt(64) = gt(64) + betqut
        gt(123) = gt(123) + 1
        End If

If gt(63) > gt(47) Or gt(75) > gt(47) Then GoTo ErrOverflowsubs

ElseIf multiprize(currline, ct, 2) = intscattervec(1, 2) Or multiprize(currline, ct, 2) = intscattervec(2, 2) Then   'stats
        If spincount > 0 Or freegamecount > 0 Then
        If currprize >= gt(69) Then
        gt(69) = currprize
        gt(70) = gt(49) + gt(50) + gt(51)
        End If
        On Error GoTo ErrOverflowscat
        gt(71) = gt(71) + currprize
        If gt72 = False Then gt(72) = gt(72) + betqut
        gt72 = True
        gt(125) = gt(125) + 1
        Else
        If currprize >= gt(57) Then
        gt(57) = currprize
        gt(58) = gt(49) + gt(50) + gt(51)
        End If
        On Error GoTo ErrOverflowscat
        gt(59) = gt(59) + currprize
        gt(60) = gt(60) + betqut
        gt(122) = gt(122) + 1
        End If

If gt(71) > gt(47) Or gt(59) > gt(47) Then GoTo ErrOverflowscat

Else    '"Natural" prizes
If spincount > 0 Or freegamecount > 0 Then
If currprize >= gt(77) Then
gt(77) = currprize
gt(78) = gt(49) + gt(50) + gt(51)
End If
On Error GoTo ErrOverflownats
gt(79) = gt(79) + currprize
If gt80 = False Then gt(80) = gt(80) + betqut
gt80 = True
gt(127) = gt(127) + 1
Else
If currprize >= gt(65) Then
gt(65) = currprize
gt(66) = gt(49) + gt(50) + gt(51)
End If
On Error GoTo ErrOverflownats
gt(67) = gt(67) + currprize
gt(68) = gt(68) + betqut
gt(124) = gt(124) + 1
End If

If gt(67) > gt(47) Or gt(79) > gt(47) Then GoTo ErrOverflownats

End If
Next
End If 'chtr


If gt(47) - pzctcum < prizecount(currline) Then
pzctcum = gt(47)
Else
pzctcum = prizecount(currline) + pzctcum
End If
If ct1 > 0 Then prizecount(currline) = prizecount(currline) + 1
prizetotal = ct1 + pzcum + prizetotal

If pzcum > 0 Then
Lnptr(2 * currline).FillColor = gt(176)
Lnptr(2 * currline + 1).FillColor = gt(176)
End If

Select Case currline
Case 0
pztool(currline) = "*Middle*: " & pztool(currline)
Case 1
pztool(currline) = "*Top*: " & pztool(currline)
Case 2
pztool(currline) = "*Bottom*: " & pztool(currline)
End Select
lblprizemeter(1).ToolTipText = lblprizemeter(1).ToolTipText & pztool(currline)

Next    'currline

With Prizeflash
If prizetotal > 0 Then
If gt(160) > 0 Then
.Interval = (4 - gt(160)) * 400
Else
.Interval = 800
End If
.Enabled = True
prizeflashon = True
End If
End With

If chtr <= 0 Then
On Error GoTo ErrOverflowln
gt(149) = gt(149) + gt(153) + 1 'multln

If gt(14) = 1 And lblrandom(0).Caption <> "RESET" Then  'RJ


'RJ game value 5%: jackpot of 1 get random between 1 & 20
'1.1 calculate random between 1 and 22 etc (VOG = 1/22*22/20 = 1/20)

If gt(18) = CLng(2 * Rnd * (10 * gt(16) + gt(17)) + 1) Then
aptot = gt(16)
 
'"Fair" with middle digit of RH
If aptot < gt(47) Then
If gt(17) > 5 Then
aptot = aptot + 1
ElseIf gt(17) = 5 Then
If Right(gt(49) + gt(50) + gt(51), 1) > 5 Then aptot = aptot + 1
End If
End If

jackpotprize = aptot

For ct = 0 To 1
lblmisc(ct + 2).ForeColor = gt(177)
lblrandom(ct).ForeColor = gt(177)
Next

lblmisc(4).Caption = Stringvars(16)
ct3 = 0
PlaySndF Stringvars(29)
jackpot.Enabled = True
Candy(0).Enabled = False
Candy(1).Enabled = False

'reset jackpot
gt(16) = gt(19)
gt(17) = 0

If aptot >= gt(81) Then 'stats
gt(81) = aptot
gt(82) = gt(49) + gt(50) + gt(51)
End If
gt(84) = gt(84) + gt(48)    'running total
gt(48) = 0
On Error GoTo Erroverflowjack
gt(83) = gt(83) + aptot
If gt(83) > gt(47) Then GoTo Erroverflowjack

End If
End If
Else
If gt(47) - betqut * ct2 * prizetotal < prizeaccum Then prizetotal = gt(47) - prizeaccum
End If  'chtr

If capsubs = True Then
lblmisc(9).Caption = "Substitute multiplier: " & gt(4) & "X"
ElseIf capnats = True Then
lblmisc(9).Caption = "Naturals multiplier : " & gt(5) & "X"
End If


If aptot > 0 Then
prizeonmeter aptot
Else
prizeonmeter betqut * prizetotal
End If

Exit Sub

ErrOverflowln:
gt(149) = gt(47)
DPzerr currline, pzcum, pztool
Exit Sub
ErrOverflowsubs:
If spincount > 0 Or freegamecount > 0 Then
gt(75) = gt(47)
Else
gt(63) = gt(47)
End If
DPzerr currline, pzcum, pztool
Exit Sub
ErrOverflowscat:
If spincount > 0 Or freegamecount > 0 Then
gt(71) = gt(47)
Else
gt(59) = gt(47)
End If
DPzerr currline, pzcum, pztool
Exit Sub
ErrOverflownats:
If spincount > 0 Or freegamecount > 0 Then
gt(79) = gt(47)
Else
gt(67) = gt(47)
End If
DPzerr currline, pzcum, pztool
Exit Sub
Erroverflowjack:
EOGEF = True
gt(83) = gt(47)
prizeonmeter aptot
Exit Sub
End Sub
Private Sub DPzerr(lineoferror As Long, pzcum As Long, pztool() As String)

If lineoferror < 3 Then 'not ln error
prizetotal = pzcum + prizetotal
pzctcum = prizecount(currline) + pzctcum
End If

If prizetotal > 0 And gt(160) > 0 Then
Prizeflash.Interval = (4 - gt(160)) * 400
Prizeflash.Enabled = True
prizeflashon = True
End If

For currline = 0 To gt(153)
If prizecount(currline) > 0 Then
Select Case currline
Case 0
pztool(currline) = "*Middle*: " & pztool(currline)
Case 1
pztool(currline) = "*Top*: " & pztool(currline)
Case 2
pztool(currline) = "*Bottom*: " & pztool(currline)
End Select
Else
pztool(currline) = "(0) "
End If
lblprizemeter(1).ToolTipText = lblprizemeter(1).ToolTipText & pztool(currline)
Next
EOGEF = True
prizeonmeter betqut * prizetotal
End Sub
Private Sub prizeonmeter(prize As Long)
Dim midipz As Long

If prize > 0 Then Quotez.Visible = False

Select Case prize / betqut
Case 25 To 99
midipz = 1
Case 100 To 249
midipz = 2
Case 249 To 999
midipz = 3
Case Is >= 1000
midipz = 4
End Select
If midipz > 0 And gt(185) > 0 Then playsuccess = PlayMidiFile(Stringvars(39 + midipz), True)

With Prizemeter

If aptot > 0 Then

mt = 1

Select Case prize
Case 1 To 4
.Interval = 288
For ct = 0 To 3
liwnf(ct).Visible = True
Next
Case 5 To 9
.Interval = 216
For ct = 0 To 7
liwnf(ct).Visible = True
Next
Case 10 To 24
.Interval = 162
For ct = 0 To 15
liwnf(ct).Visible = True
Next
Case 25 To 49
.Interval = 108
For ct = 0 To 23
liwnf(ct).Visible = True
Next
Case 50 To 99
.Interval = 72
For ct = 0 To 27
liwnf(ct).Visible = True
Next
Case 100 To 249
.Interval = 54
For ct = 0 To 31
liwnf(ct).Visible = True
Next
Case 250 To 999
.Interval = 36
For ct = 0 To 39
liwnf(ct).Visible = True
Next
Case Else
.Interval = 18    '18 clicks/sec
For ct = 0 To 47
liwnf(ct).Visible = True
Next
End Select

Else    'spun prizes
mt = CLng(betqut)

Select Case prize
Case 0
waitimer.Enabled = True
.Enabled = False
Exit Sub
Case mt To 4 * mt
.Interval = 288
For ct = 0 To 3
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(30)
lblmisc(4).Caption = Stringvars(17)
Case 5 * mt To 9 * mt
.Interval = 216
For ct = 0 To 7
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(31)
lblmisc(4).Caption = Stringvars(18)
Case 10 * mt To 24 * mt
.Interval = 162
For ct = 0 To 15
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(32)
lblmisc(4).Caption = Stringvars(19)
Case 25 * mt To 49 * mt
.Interval = 108
For ct = 0 To 23
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(33)
lblmisc(4).Caption = Stringvars(20)
Case 50 * mt To 99 * mt
.Interval = 72
For ct = 0 To 27
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(34)
lblmisc(4).Caption = Stringvars(21)
Case 100 * mt To 249 * mt
.Interval = 54
For ct = 0 To 31
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(35)
lblmisc(4).Caption = Stringvars(22)
Case 250 * mt To 999 * mt
.Interval = 36
For ct = 0 To 39
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(36)
lblmisc(4).Caption = Stringvars(23)
Case 1000 To 4999
.Interval = 18    '18 clicks/sec
For ct = 0 To 47
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(37)
lblmisc(4).Caption = Stringvars(24)
Case Else
lblmisc(4).Caption = Stringvars(25)
.Interval = 10
For ct = 0 To 47
liwnf(ct).Visible = True
Next
PlaySndF Stringvars(38)

End Select

If chtr > 0 Then lblmisc(4).Caption = "Cheater's Gold!"

End If

'Put prize on meter
If Len(pzzeros) <> Len(CStr(gt(47))) - Len(CStr(prize)) Then
pzzeros = ""
For ct1 = Len(CStr(prize)) + 1 To Len(CStr(gt(47)))
pzzeros = "0" & pzzeros
Next
End If
lblprizemeter(1).Caption = pzzeros & CStr(prize)

If aptot = 0 Then
 If gt(159) = 0 Then
 lblmisc(0).Caption = betqut & "X ="
 Else
 lblmisc(0).Caption = "-" & betqut & "X-"
 End If
Else
lblmisc(0).Caption = "WIN"
End If
lblmisc(0).ForeColor = gt(177)


On Error GoTo ErrOverflowCash
prizeaccum = prize + prizeaccum

'If no overflow, do things gracefully
If prizeaccum >= gt(47) Then EOGEF = True


If jackpot.Enabled = False And timoneyback.Enabled = False Then .Enabled = True

End With
Exit Sub

ErrOverflowCash:
gt(2) = gt(47)
reachedlimit Cashlim
End Sub
Private Sub Prizemeter_Timer()
DoPrz
End Sub
Private Sub DoPrz()
Dim temp As String

If Prizemeter.Interval = 1000 Then
prizeonmeter betqut * prizetotal
Exit Sub
End If

With lblprizemeter(0)

If EOGEF = True Then

If gt(47) = 2147483647 Then 'Happens a lot?
If CLng(.Caption) >= 2147483647 - lnbet Then
.Caption = 2147483647
GoTo EOGerror
End If
ElseIf .Caption >= gt(47) Then
.Caption = gt(47)
GoTo EOGerror
End If
End If


temp = CStr(CLng(.Caption) + mt)


If Len(cumpzzeros) <> Len(CStr(gt(47))) - Len(temp) Then
cumpzzeros = ""
For ct1 = 1 To Len(CStr(gt(47))) - Len(temp)
cumpzzeros = "0" & cumpzzeros
Next
.AutoSize = True
.AutoSize = False
End If

.Caption = cumpzzeros & temp


If .Caption = prizeaccum Then
If EOGEF = True Then GoTo EOGerror
If aptot > 0 Then
aptot = 0
If prizetotal > 0 Then
Prizemeter.Interval = 1000
Exit Sub
End If
If Quotez.Picture.Handle <> 0 Then PlaySndF ("Thumbb")
Else
If chtr > 0 Then prizeaccum = prizeaccum - betqut * prizetotal
End If
'Eligiblespingame if waitimmarker > 0
waitimer.Enabled = True
Prizemeter.Enabled = False
End If


End With

Exit Sub
EOGerror:
gt(2) = gt(47)
reachedlimit Cashlim
End Sub
Private Sub timoneyback_Timer()
If flashtoggle = False Then
lblmisc(1).Caption = "Money Paid Back"
ct3 = ct3 + 1
flashtoggle = True
Else
lblmisc(1).Caption = ""
ct3 = ct3 + 1
flashtoggle = False
End If
If ct3 = 15 Then
flashtoggle = False
timoneyback.Enabled = False
lblmisc(1).ForeColor = gt(163)
lblmisc(1).Caption = "Money Back Status ....                POT ...."
Prizemeter.Enabled = True
End If
End Sub
Private Sub jackpot_Timer()
If flashtoggle = False Then
lblmisc(2).Caption = "Jackpot Activated"
ct3 = ct3 + 1
flashtoggle = True
Else
lblmisc(2).Caption = "Amount Won = " & aptot
ct3 = ct3 + 1
flashtoggle = False
End If
If ct3 = 15 Then
flashtoggle = False
jackpot.Enabled = False
Prizemeter.Enabled = True
End If
End Sub
Private Sub prizeflash_Timer()
Dim c1 As Integer, c2 As Integer

If prizeflashon = True Then
indicln
For c1 = 1 To pzctcum
For c2 = 0 To 70 Step 14
If oldlblprize(c1) < c2 Then
lblprize(oldlblprize(c1)).ForeColor = gt(170 + Int(oldlblprize(c1) / 14))
lblprizeamt(oldlblprize(c1)).BackStyle = 0
lblprizeamt(oldlblprize(c1)).ForeColor = gt(170 + Int(oldlblprize(c1) / 14))
Exit For
End If
Next
Next
Else
If gt(160) = 0 Then Prizeflash.Enabled = False

For c1 = 0 To 2
If prizecount(c1) > 0 Then
Lnptr(2 * c1).FillColor = gt(176)
Lnptr(2 * c1 + 1).FillColor = gt(176)
End If
Next

For c1 = 1 To pzctcum
For c2 = 0 To 70 Step 14
If oldlblprize(c1) < c2 Then
lblprize(oldlblprize(c1)).ForeColor = gt(176)
lblprizeamt(oldlblprize(c1)).ForeColor = gt(176)
Exit For
End If
Next
Next

End If


prizeflashon = Not (prizeflashon)

End Sub
Private Sub paintitlearea()
wid = Spare.Picture.Width
hgt = Spare.Picture.Height

For ct = 0 To TA.ScaleWidth Step 51 * wid / 90
For ct1 = 0 To TA.ScaleHeight Step 51 * hgt / 90
TA.PaintPicture Spare.Picture, ct, ct1
Next
Next
End Sub
Private Sub prizesonform(errno As Long, picno As Long, picoffset As Long, scatno As Long)
Dim picplace As Long, testval As Long

testval = -1
'Scatno 12 or 13

If scatno > 0 Then
picplace = scatno
Else
picplace = picno - picoffset - 1
End If

On Error GoTo Errsymbols


Set thumbs(picno) = LoadPicture(CStr(picno) & ".bmp")
Thumbslist(5).ListImages.Add (picno), , thumbs(picno)
Spare.PaintPicture Thumbslist(5).ListImages(picno).Picture, 0, 0, 400, 400

imgprizethumb(picplace).Visible = True
imgprizethumb(picplace).Picture = Spare.Image

Set Spare = Nothing


For intreel = 0 To 4
'Don't show some prize values if lt 5
        If sst(picno, 2) = 1 Or sortprizes(intreel + 6, picno, testval) = True Then
        lblprize(14 * intreel + picplace).Visible = True
        lblprizeamt(14 * intreel + picplace).Visible = True
        lblprizeamt(14 * intreel + picplace) = sst(picno, 6 + intreel)
        End If
Next
Exit Sub


Errsymbols:
errno = 1
End Sub
Private Sub sortlabels(lightlbl As Boolean, multinum As Long, loopcount As Long)

If intscatternumber = 2 Then
    If loopcount < intscattervec(2, 2) Then
        If loopcount < intscattervec(1, 2) Then
        lbloffset = 0
        ElseIf loopcount = intscattervec(1, 2) Then
            If lightlbl = False Then
            prizesonform multinum, loopcount, 0, 12
            Else
            lightupprize multinum, loopcount, 0, 12
            End If
        Else
        lbloffset = 1
        End If
    ElseIf loopcount = intscattervec(2, 2) Then
        If lightlbl = False Then
        prizesonform multinum, loopcount, 0, 13
        Else
        lightupprize multinum, loopcount, 0, 13
        End If
    Else
    lbloffset = 2
    End If
ElseIf intscatternumber = 1 Then
    If loopcount < intscattervec(1, 2) Then
    lbloffset = 0
    ElseIf loopcount = intscattervec(1, 2) Then
        If piccount = 14 Then
            If lightlbl = False Then
            prizesonform multinum, loopcount, 0, 13
            Else
            lightupprize multinum, loopcount, 0, 13
            End If
        Else
            If lightlbl = False Then
            prizesonform multinum, loopcount, 0, 12
            Else
            lightupprize multinum, loopcount, 0, 12
            End If
        End If
    Else
    lbloffset = 1
    End If
End If

If loopcount <> intscattervec(1, 2) And loopcount <> intscattervec(2, 2) Then
If lightlbl = False Then
prizesonform multinum, loopcount, lbloffset, 0
Else
lightupprize multinum, loopcount, lbloffset, 0
End If
End If

End Sub
Private Sub lightupprize(pzect As Long, picno As Long, picoffset As Long, scatno As Long)
Dim parttot As Long, picplace As Long

If scatno > 0 Then
picplace = scatno
Else
picplace = picno - picoffset - 1
End If


'oldlblprize(loopcount) = (prizelevel - 6) + scaled picno to identify prize
parttot = (multiprize(currline, pzect, 1) - 6) * 14 + picplace


oldlblprize(pzect + pzctcum) = parttot
With lblprizeamt(parttot)
.BackStyle = 1
.BackColor = gt(176)
.ForeColor = lblprize(parttot).BackColor
.ForeColor = gt(176)
End With


End Sub
Public Sub Cyclemultibets()
Dim cb As Long
cb = gt(20)


For ct1 = 0 To 47
liwnf(ct1).Visible = False
Next
lblmisc(0).ForeColor = gt(163)

cb = cb + 1
If cb = 6 Or gt(20 + cb) = 0 Then cb = 1
betqut = gt(20 + cb)

For ct1 = 1 To 10
If betqut = ct1 Then
lblmisc(0).Caption = "BET " & betqut & "X"
gt(20) = cb
PlaySndF Stringvars(27)
Exit For
End If
Next

If gt(37) = 1 Then special = Int(4 * Rnd + 1)
End Sub
Private Sub reachedlimit(errReason As String)

If gt(75) = gt(47) Or gt(63) = gt(47) Then
errReason = Sublim
ElseIf gt(71) = gt(47) Or gt(59) = gt(47) Then
errReason = Scatlim
ElseIf gt(79) = gt(47) Or gt(67) = gt(47) Then
errReason = Pzlim
ElseIf gt(83) = gt(47) Then
errReason = Jacklim
ElseIf gt(87) = gt(47) Or gt(67) = gt(47) Then
errReason = Mbaklim
ElseIf prizeaccum = gt(47) Then
errReason = Cashlim
ElseIf gt(149) = gt(47) Then
errReason = Multln
Else
response = MsgBox("Unknown Error, Quitting.", vbOKOnly)
Quitt
End If

If prizeaccum > 0 Then gt(156) = 1 'Stats

lblmisc(4).Caption = errReason & "!"

waitimer.Enabled = False
Prizemeter.Enabled = False
gamespinwait.Enabled = False
jackpot.Enabled = False
timoneyback.Enabled = False
spincount = 0
freegamecount = 0

showendofgame

response = MsgBox("Congratulations, you have successfully completed your game!. The specified " & errReason & " of " & gt(47) & " in your Game Stats has been reached. You will need to start a new game by clicking the start new Game button on the main Configuration Panel. Good luck and thank you for playing!", vbOKOnly)
End Sub
Private Sub showendofgame()
With TA
.Enabled = False
If gt(162) = -1 Then
        If Stringvars(1) = Stringvars(2) Then
        .Visible = False
        .Enabled = False
        waitimmarker = 5
        waitimer.Interval = 2 * gt(194)
        waitimer.Enabled = True
        Exit Sub
        Else
        paintitlearea
        End If
Else
.BackColor = gt(162)
End If
TextCircle TA, "End_of_Game", .ScaleWidth / 2, .ScaleHeight, .ScaleHeight, degreeoftitle, .Fontsize
lblmisc(4).ToolTipText = EOGTT
For ct = 0 To 1
Candy(ct).ToolTipText = EOGTT
Next
activatectrls
End With
End Sub
Public Sub activatectrls()
If activatctrls = False Then Exit Sub
'Here so VBreturn works

TA.Enabled = True
TA.Visible = True

Pokemach.KeyPreview = True

Candy(0).Enabled = True
Candy(1).Enabled = True

activatctrls = False
End Sub
Private Sub moneybackdrama(currprizeamt As Long)
Dim countmove As Long, moneyback As Long
moneyback = gt(13)
countmove = gt(11)
'prizeaccold = prizeaccum beginning of FSFG


If freegamecount = 0 And spincount = 0 Then
    If currprizeamt = 0 Then
        If countmove > 0 Then
        countmove = countmove + 1
            If countmove > gt(10) Then
            aptot = moneyback + lnbet
            
            lblmisc(1).ForeColor = gt(177)
            lblmisc(4).Caption = Stringvars(15)
            ct3 = 0
            stripMBstats False
            gt(151) = gt(10)
            gt(85) = moneyback + lnbet    'stats
            gt(86) = gt(49) + gt(50) + gt(51)
            On Error GoTo ErroverflowMB
            gt(87) = gt(87) + gt(85)
            If gt(87) > gt(47) Then GoTo ErroverflowMB
            timoneyback.Enabled = True
            PlaySndF Stringvars(28)
            waitimer.Enabled = False
            If gt(12) = 0 Then 'MB kicks off at start of game!
            countmove = 1
            lblmoneyback(0).ForeColor = gt(177)
            Else
            countmove = 0
            End If
            moneyback = 0
            lblmisc(7).Caption = ""
            Else
            moneyback = moneyback + lnbet
            lblmisc(7).Caption = moneyback      'Pot
            lblmoneyback(countmove - 2).ForeColor = gt(163)
            lblmoneyback(countmove - 1).ForeColor = gt(177)
            End If
        ElseIf gt(12) = 0 Then 'money back kicks off at start of game!
        countmove = 1
        lblmoneyback(0).ForeColor = gt(177)
        End If
    ElseIf currprizeamt >= gt(12) Then
    moneyback = 0
    lblmisc(7).Caption = moneyback
    If countmove > 1 Then lblmoneyback(countmove - 1).ForeColor = gt(163)
    countmove = 1
    lblmoneyback(0).ForeColor = gt(177)
    stripMBstats True   'Stats
    Else
        If countmove > 0 Then
        lblmoneyback(countmove - 1).ForeColor = gt(163)
        Else
        lblmoneyback(0).ForeColor = gt(163)
        End If
    stripMBstats True   'Stats
    lblmisc(7).Caption = ""
    countmove = 0
    End If
End If
gt(11) = countmove
gt(13) = moneyback
Exit Sub

ErroverflowMB:
EOGEF = True
gt(87) = gt(47)
prizeonmeter aptot
End Sub
Private Sub fixmonback()

Select Case currold
Case 1
currold = -1
Case 2
currold = 1
End Select

For intreel = 0 To 4    'Just finished FSFG so align held reels
'adjust restore pos for dirofspin
If hreel(intreel) = True Then

    If dirspin(intreel) <> dirofspin Then
        
        
    If dirofspin = -1 Then  'spinning up
    trakker(intreel) = 3
    
    For ct = 0 To 3
    picnum(intreel, 3 - ct) = 4 * intreel + ct
    Next
    For ct = 0 To 3
    spinztemp(intreel) = spinz(intreel) 'advanz
    pwScale = fixpw(3 - ct, pw)
    M(picnum(intreel, ct)).Top = pwScale
    Set M(picnum(intreel, ct)).Picture = Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), ct - 3 - currold)).Picture
    ydisp0(ct) = pwScale
    Next
    
    
    Else
    trakker(intreel) = 3
    

    
    For ct = 0 To 3
    picnum(intreel, ct) = 4 * intreel + ct
    Next
    For ct = 0 To 3
    spinztemp(intreel) = spinz(intreel) 'advanz
    pwScale = fixpw(ct - 1, pw)
    M(picnum(intreel, 3 - ct)).Top = pwScale
    Set M(picnum(intreel, 3 - ct)).Picture = Thumbslist(intreel).ListImages(Advanz(spinztemp(intreel), 3 - ct + currold)).Picture
    ydisp0(ct) = pwScale
    Next
    End If
    
    'compensate for position change
    spinz(intreel) = Advanz(spinz(intreel), 2 * dirofspin - dirspin(intreel) * currold)
    Else
    spinz(intreel) = Advanz(spinz(intreel), -dirspin(intreel) * currold)
    End If

End If
Next

'Reset MB if won prize
waitimer.Interval = 2 * gt(194)
If prizeaccold > 0 Then
moneybackdrama prizeaccum - prizeaccold
prizeaccold = 0
    If EOGEF = False And aptot > 0 Then
    prizeonmeter aptot
    Exit Sub
    End If
End If

waitimmarker = 0
waitimer.Enabled = True
End Sub
Private Sub Cachereels()

pwScale = fixpw(1, pw)

TOL = pwScale / 6 'just below min movesize

TA.Enabled = False
lblmisc(4).Caption = "Caching Reels : Please Wait"
For ct = 1 To 24
medfastslowmove(3, ct) = pwScale / 3
Next

For intreel = 0 To 4
movesize(intreel) = pwScale / 3
Next

If dirofspin = 1 Then
ydispmin = -pwScale
ydispmax = 0
Else
ydispmin = 3 * pwScale
ydispmax = 2 * pwScale
End If

With frapicarea
.BackColor = vbWhite
.Caption = "Caching"
.Fontsize = 24
.Refresh
End With
Pokemach.AutoRedraw = False
For ct = 1 To 75
reelmin = 0
reelmax = 4
spinning 3
Next
frapicarea.BackColor = vbBlack
frapicarea.Caption = ""
justcached = True
Pokemach.Enabled = True
activatectrls
Pokemach.AutoRedraw = True
End Sub
Public Sub Frmaboutt()
If zhiddnstatus > 9 Then Exit Sub
Prizeflash.Enabled = False
Load frmAbout
Pokemach.Hide
End Sub
Public Sub Configurationn()
If zhiddnstatus > 9 Then Exit Sub
LoadFrmSplsh 460
setthumbspiccount thumbs, piccount, wheelvec, wheelorder
cleanup
UnloadForms
If outputvars = True Then
Load gametype
gametype.Show
Stopnoise 3
Else
Quitt
End If
Unload frmSplsh
End Sub
Public Sub Quitt()
Dim c As New cRegistry

'gt(55) = 0   'trigger
'gt(56) = 0

cleanup
UnloadForms
If outputvars = False Then Exit Sub
If gt(46) = 1 Then
gt(0) = -2
Load frmAbout
frmAbout.Show
End If
Stopnoise
c.CloseMutexhandle
End Sub
Public Sub Changedir()
If zhiddnstatus > 9 Then Exit Sub
gt(0) = -1
cleanup
Stopnoise 1
UnloadForms
procend = False
If outputvars = True Then
Load Pokemach
If procend = True Then
Pokemach.Show
Dotaskwindow Pokemach
Else
Unload Pokemach
Set Pokemach = Nothing
End If
End If
End Sub
Private Sub UnloadForms()
If keyPrsd = True Then
Unload Zhidden
Set Zhidden = Nothing
End If
Unload Pokemach
Set Pokemach = Nothing
End Sub
Private Sub cleanup()
Dim ctl As VB.Control
 For Each ctl In Me.Controls
  If TypeOf ctl Is VB.Timer Then ctl.Interval = 0
 Next ctl
If gt(156) = 0 Then gt(2) = prizeaccum
If Dir(loaddirectory & "q0.bmp") <> "" Then Kill (loaddirectory & "q0.bmp")
End Sub
