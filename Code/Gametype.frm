VERSION 5.00
Object = "{10AB04E3-3AEA-101C-96E6-0020AF38F4BB}#1.0#0"; "TEGSPIN3.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form gametype 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7620
   ClientLeft      =   3450
   ClientTop       =   1155
   ClientWidth     =   8010
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
   HelpContextID   =   5
   Icon            =   "Gametype.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8010
   Begin VB.CheckBox chkgametype 
      Caption         =   "Feature Bonus restricted to Payline from whence the Feature originated."
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   52
      Top             =   3360
      Width           =   5295
   End
   Begin VB.CommandButton cmdgametype 
      Caption         =   " "
      Height          =   435
      Index           =   6
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Frame fragametype 
      Caption         =   "Multiple Prizes on a Payline (Substitutes maximised but used once only)"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   38
      Top             =   2520
      Width           =   5535
      Begin VB.OptionButton optgametype 
         Caption         =   " Highest win pays"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Scatter wins are compared against centre (or middle) line wins"
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optgametype 
         Caption         =   "All wins pay"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fragametype 
      Caption         =   "Bet Multiplier Spin Boxes"
      Height          =   625
      Index           =   4
      Left            =   2400
      TabIndex        =   37
      Top             =   1820
      Width           =   5535
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   7
         Left            =   240
         TabIndex        =   41
         Top             =   210
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   8
         Left            =   1320
         TabIndex        =   43
         Top             =   210
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   9
         Left            =   2400
         TabIndex        =   45
         Top             =   210
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   10
         Left            =   3480
         TabIndex        =   47
         Top             =   210
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   11
         Left            =   4560
         TabIndex        =   49
         Top             =   210
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   4920
         TabIndex        =   50
         Top             =   210
         Width           =   375
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3840
         TabIndex        =   48
         Top             =   210
         Width           =   375
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2760
         TabIndex        =   46
         Top             =   210
         Width           =   375
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1680
         TabIndex        =   44
         Top             =   210
         Width           =   375
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   42
         Top             =   210
         Width           =   375
      End
   End
   Begin VB.Frame fragametype 
      Caption         =   "A Picture with a Substituter"
      Height          =   1215
      Index           =   3
      Left            =   2400
      TabIndex        =   28
      Top             =   240
      Width           =   5080
      Begin VB.OptionButton optgametype 
         Caption         =   "Pays"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Doesn't apply to single - picture prizes"
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optgametype 
         Caption         =   "Pays"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   5
         Left            =   840
         TabIndex        =   32
         Top             =   240
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   6
         Left            =   840
         TabIndex        =   34
         Top             =   720
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin VB.Label labgametype 
         Caption         =   "times if Naturals only in winning combination"
         Height          =   255
         Index           =   6
         Left            =   1545
         TabIndex        =   36
         Top             =   840
         Width           =   3270
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   " "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   35
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   " "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.Label labgametype 
         Caption         =   "times if Substituter is used in winning combination"
         Height          =   255
         Index           =   0
         Left            =   1540
         TabIndex        =   31
         Top             =   360
         Width           =   3445
      End
   End
   Begin VB.Frame fragametype 
      Height          =   975
      Index           =   2
      Left            =   2400
      TabIndex        =   20
      Top             =   4680
      Width           =   5535
      Begin VB.CheckBox chkgametype 
         Caption         =   "Money Back feature commences only after a prize paying at least"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4935
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   4
         Left            =   3720
         TabIndex        =   23
         Top             =   480
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   27
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.Label labgametype 
         Caption         =   "times original bet. All bets repaid after"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   25
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label labgametype 
         Caption         =   "nowin spins"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdgametype 
      Caption         =   "&General Options ..."
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   14
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdgametype 
      Caption         =   "&Recalculate Return Now"
      Height          =   375
      Index           =   4
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "This can take a little time "
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton cmdgametype 
      Caption         =   "&New Reel Pictures ..."
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   12
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdgametype 
      Caption         =   "&Edit Reel Combinations ..."
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   11
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdgametype 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Frame fragametype 
      Height          =   975
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   3600
      Width           =   5535
      Begin VB.CheckBox chkgametype 
         Caption         =   "Activate Random cumulative Jackpot, starting amount : "
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   0
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin TegspinLibCtl.TegoSpin spngametype 
         Height          =   360
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         BevelWidth      =   3
         Interval        =   125
      End
      Begin VB.Label labgametype 
         Caption         =   "increments of 0.1 at a time"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label labgametype 
         Caption         =   "with"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblgametype 
         BackColor       =   &H8000000E&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.PictureBox picspare 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdgametype 
      Caption         =   "Save && &Quit"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin TegspinLibCtl.TegoSpin spngametype 
      Height          =   360
      Index           =   2
      Left            =   6960
      TabIndex        =   16
      Top             =   5760
      Width           =   345
      _Version        =   65536
      _ExtentX        =   609
      _ExtentY        =   635
      _StockProps     =   64
      BevelWidth      =   3
      Interval        =   125
   End
   Begin VB.Label lblgametype 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   405
      Index           =   12
      Left            =   2520
      TabIndex        =   19
      Top             =   6240
      Width           =   90
   End
   Begin VB.Label lblgametype 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   17
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label labgametype 
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   9
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Label labgametype 
      Caption         =   "Free Game  - Free Spin Contentions - click Picture to promote"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Image imggamespin 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   3
      Left            =   7560
      Top             =   0
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imggamespin 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   2
      Left            =   7560
      Top             =   480
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imggamespin 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   1
      Left            =   7560
      Top             =   960
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imggamespin 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   0
      Left            =   7560
      Top             =   1440
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   13
      Left            =   1320
      Top             =   6600
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   12
      Left            =   1320
      Top             =   5520
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   11
      Left            =   1320
      Top             =   4440
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   10
      Left            =   1320
      Top             =   3360
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   9
      Left            =   1320
      Top             =   2280
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   8
      Left            =   1320
      Top             =   1200
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   7
      Left            =   1320
      Top             =   120
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   6
      Left            =   120
      Top             =   6600
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   5
      Left            =   120
      Top             =   5520
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   4
      Left            =   120
      Top             =   4440
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   3
      Left            =   120
      Top             =   3360
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   2
      Left            =   120
      Top             =   2280
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   1
      Left            =   120
      Top             =   1200
      Width           =   915
   End
   Begin VB.Image imgsymbolsel 
      Appearance      =   0  'Flat
      Height          =   915
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   915
   End
   Begin VB.Line lnegamespin 
      BorderWidth     =   2
      Index           =   0
      Tag             =   " "
      Visible         =   0   'False
      X1              =   6960
      X2              =   7320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line lnegamespin 
      BorderWidth     =   2
      Index           =   1
      Tag             =   " "
      Visible         =   0   'False
      X1              =   7200
      X2              =   7320
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line lnegamespin 
      BorderWidth     =   2
      Index           =   2
      Tag             =   " "
      Visible         =   0   'False
      X1              =   7200
      X2              =   7320
      Y1              =   1800
      Y2              =   1680
   End
   Begin ComctlLib.ImageList Thumblist 
      Left            =   6720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "gametype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thumbs(14) As StdPicture, wheelorder(4, 24) As Long, wheelvec(5, 14) As Long
Dim piccount As Long, ct As Long, temp As Long, response As Long, gamespintot As Long
Dim remembchkmonback As Boolean, remembchkjackpot As Boolean, newpiccalcreturn As Long, confirmnewgame As Boolean
Dim captval As String, zcolour As Long, estp0 As Single, leftgt155 As Long, rightgt155 As Long
Dim tgs(2, 9) As Long, tss(2, 15) As Long, tsub(14, 14) As Long, tgamespinkeep(3) As Long, tgamespinsymbol(3) As Long
Private Sub Form_Load()
'gt(0)=-2 when g(46) = 1, we have quit pokemach and wish to view the about page
'gt(0)=-1 on right click - change directory
'gt(0)=0 on first load
'gt(0)=1 on gametype cancel
'gt(0)=2 second load of pokemach (and gametype save)
'gt(0)=3 when cornfig is invoked
'gt(0)=4 when cornfig is cancelled
'gt(0)=5 on cfgthumb or load of SSTAB without saving new config, or game limit reached in Pokemach, or just stated new game
'gt(0)=6 on second load of cornfig or cfgthumb after cfgthumb without saving new config
'gt(1)=1 In cornfig requests ordered reel combinations
'gt(2)= money on hand
'gt(3)= degree of title in pokemach
'gt(4)= prize multiplier with substitute
'gt(5)= prize multiplier with naturals only
'gt(6)= 0 - 4 spin direction percentage
'gt(7)= 0 - 4 spin duration percentage
'gt(8)= 0 - 4 spin speed percentage
'gt(9)= 1 when All wins pay
'gt(10)= number of spins before money back
'gt(11)= current money back status (0 - 15)
'gt(12)= money back feature only begins after prize of this or more
'gt(13)= money back pot
'gt(14)= 1 when random jackpot is selected
'gt(15)= Random jackpot increments of 0.1
'gt(16)= random jackpot meter value left
'gt(17)= random jackpot meter value right
'gt(18)= secret jackpot value
'gt(19)= starting jackpot value
'gt(20)= current multiple bet ratio - one of gt(21-25)
'gt(21-25)=  spinbox bet multipliers
'gt(26)= gt(26) + 1 if, on cornfig "continue" and wheelvec change
'gt(27)= gt(27) + 1 if a change in SStab
'gt(28)= gt(28) + 1 if a change in gametype (see pokrouts anychanges for details)
'gt(29)= LH value of the game
'gt(30)= RH value of the game
'gt(31)= LH moneybackvalue of the game
'gt(32)= RH moneybackvalue of the game
'gt(33)= LH value of P0
'gt(34)= RH value of P0
'gt(35)= old random seed (see genopts)
'gt(36)= current seed-change interval index sec/min/hour etc
'gt(37)= allow special to change random process
'gt(38)= time of randomization in units of sec
'gt(39)= time of randomization in units of min
'gt(40)= time of randomization in units of hour
'gt(41)= week of randomization in units of day of month
'gt(42)= month of randomization in units of month
'gt(43)= year of randomization in units of year
'gt(44)= 1 crops square in cfgthumb, Thumbnails
'gt(45)= 1 if game is restarted with same random seed after clicking main configuration OK, 0 to continue game
'gt(46)= View about page when quitting game
'gt(47)= Game tolerance
'gt(48)= bet total of RJ wins SINCE last RJ win
'gt(49)= Number of free games taken
'gt(50)= Total number of free spins taken
'gt(51)= Total number of spins - free games - free spins
'gt(52)= No of spins up
'gt(53)= No of spins of short duration
'gt(54)= No of spins at fast speed
'gt(55)= Number of times game replayed with same random seed
'gt(56)= Number of times game replayed with new random seed
'gt(57)= Largest scatter prize wsf (Not in FS FG)
'gt(58)= How many turns ago
'gt(59)= Total scatter prizemoney
'gt(60)= Bet total of scatter wins
'gt(61)= Largest substitute prize wsf (Not in FS FG)
'gt(62)= How many turns ago
'gt(63)= Total substitute prizemoney
'gt(64)= Bet total of substitute wins
'gt(65)= Largest naturals prize wsf (Not in FS FG)
'gt(66)= How many turns ago
'gt(67)= Total naturals prizemoney
'gt(68)= Bet total of naturals wins
'gt(69)= Largest scatter prize wsf FS FG
'gt(70)= How many turns ago
'gt(71)= Total scatter prizemoney
'gt(72)= Bet total of scatter wins FS FG
'gt(73)= Largest substitute prize wsf FS FG
'gt(74)= How many turns ago
'gt(75)= Total substitute prizemoney
'gt(76)= Bet total of substitute wins FS FG
'gt(77)= Largest naturals prize wsf FS FG
'gt(78)= How many turns ago
'gt(79)= Total naturals prizemoney
'gt(80)= Bet total of naturals wins FS FG
'gt(81)= Largest RJ prize wsf
'gt(82)= How many turns ago
'gt(83)= Total RJ prizemoney
'gt(84)= Bet total UP TO last RJ win - can be = gt(48)
'gt(85)= Last MB prize wsf
'gt(86)= How many turns ago
'gt(87)= Total MB prizemoney
'gt(88)= Bet total of MB wins
'gt(89 - 103)= MB strategy used in gt(85) - Concatenated with new strategies see stripMBstats
'gt(104)= LH scatter VOG when Frmabout was last invoked
'gt(105)= RH gt(104)
'gt(106)= LH Substitute VOG when Frmabout was last invoked
'gt(107)= RH gt(106)
'gt(108)= LH Naturals VOG when Frmabout was last invoked
'gt(109)= RH gt(108)
'gt(110-115)= gt(104-109) in FG FS
'gt(116)= Random Jackpot VOG when Frmabout was last invoked
'gt(117)= RH gt(116)
'gt(118)= Money Back VOG when Frmabout was last invoked
'gt(119)= RH gt(118)
'gt(120)= FS FG no prize bonus VOG when Frmabout was last invoked
'gt(121)= RH gt(118)
'gt(122)= number of turns in prize wins scatters not FSFG
'gt(123)= number of turns in prize wins substitutes not FSFG
'gt(124)= number of turns in prize wins Naturals not FSFG
'gt(125)= number of turns in prize wins scatters FSFG
'gt(126)= number of turns in prize wins substitutes FSFG
'gt(127)= number of turns in prize wins Naturals FSFG
'gt(128)= number of turns while RJ active
'gt(129)= number of turns in prize wins MB
'gt(130)= number of turns in FS FG noprize bonus
'gt(131)= No prize bonus total
'gt(132)= bet total of no prize bonus
'gt(133)= total Bet quantity
'gt(134)= Total bet quantity of prizes won
'gt(135)= total Bet quantity when Frmabout was last invoked
'gt(136)= total prize money when Frmabout was last invoked
'gt(137)= leading zeros for RH gt(30)
'gt(138)= above for gt(32)
'gt(139)= above for gt(34)
'gt(140)= above for gt(105)
'gt(141)= above for gt(107)
'gt(142)= above for gt(109)
'gt(143)= above for gt(111)
'gt(144)= above for gt(113)
'gt(145)= above for gt(115)
'gt(146)= above for gt(117)
'gt(147)= above for gt(119)
'gt(148)= above for gt(121)
'gt(149)= Total number of lines played
'gt(150)= Number of times STATS reset
'gt(151)= last MB gt(10) in case it was changed
'gt(152)= MonteCarlo iterations, 0 for no Monte Carlo choice, neg for recalcnow
'gt(153)= Current number of lines played :0 for 1 line , 1 for 2 lines, 2 for 3 lines
'gt(154)= probgtequalthan(gt(12))
'gt(155)= LHS starting money, RHS starting money subject to change
'gt(156)= 1 if game ends with overflow
'gt(157)= Time between FSFGs 0, short, 1, med, 2, long (28 * gt(194) * gt(157)) + 360
'gt(158)= Default status 0 if saved, 1 current, 2 generic
'gt(159)= 1 if Big pictures 0 otherwise
'gt(160)= Flashing prizes speed 0 - none 1 - slow 2 - med 3 - fast
'gt(161)= wallpaper colour
'gt(162)= title colour
'gt(163)= wallpaper forecolour
'gt(164)= title forecolour
'gt(165)= Pokemach prize 5s colour default &H00C000C0&
'gt(166)= prize 4s colour &H00FF8080&
'gt(167)= prize 3s colour &H0000C000&
'gt(168)= prize 2s colour &H00008080&
'gt(169)= prize 1s colour &H00C0C0C0&
'gt(170)= Pokemach text 5s colour
'gt(171)= text 4s colour
'gt(172)= text 3s colour
'gt(173)= text 2s colour
'gt(174)= text 1s colour
'gt(175)= Prize Money colour
'gt(176)= Highlight colour
'gt(177)= Highlight Text colour
'gt(178)= Winning Lines colour
'gt(179)= Spin button forecolour
'gt(180)= Spin button style
'gt(181)= Bet button forecolour
'gt(182)= Bet button style
'gt(183)= 1 to List only currently assigned sounds
'gt(184)= FSFG bonuses restricted to payline where feature is awarded.
'gt(185)= Midi port if >= 0; < 0 midi disabled
'gt(186)= 0 All Sound off, 1 Sound no Thumbnail, 2 No Sound but keep Thumbnail setting, 3 All sound
'gt(187)= {-1 to -8} to Randomise order of midi playback, {1 to 8}: current midi position
'gt(188)= current quote position (gt(195) = 0))
'gt(189)= 1 to write to text after Configure Thumbnail
'gt(190)= Shortcuts enabled
'gt(191)= No of quotes in db
'gt(192)= Slotdata.s$t version counter  0 for ver 1.x, 1 for 2.0X, 2.1, 2.11, 2 for 2.12-2.14, 3 for 2.2.0-2, 4 for 2.2.3-19, 5 for 2.2.20 or later
'gt(193)= 1 to use quotes.s$t in base directory
'gt(194)= FastPC 1 slowest 59 fastest
'gt(195)= 1 Randomise quotes
'gt(196)= No of different pics in DB
'gt(197)= spare zeros for 198, 199
'gt(198)= spare for outputforformat
'gt(199)= same as gt(198)
'gt(200)= Installation trigger


'Stringvars:
'(1)= directory + fname of form tiled bmp
'(2)= directory + fname titlearea tiled bmp
'(3)= PATH name of quotes file
'(4)= PATH of Hall of Fame db
'(5)= Game title
'(6)= Spin button title
'(7)= Change bet button title
'(8)= Players name
'(9)= Title Font
'(10)= General Font
'(11)= Spin Button Font
'(12)= Spin Button Font
'(13)= Welcome Prompt
'(14)= Spin Prompt
'(15)= MB message
'(16)= RJ message
'(17)= 1 - 4
'(18)= 5 - 9
'(19)= 10 - 24
'(20)= 25 - 49
'(21)= 50 - 99
'(22)= 100  - 249
'(23)= 250 - 999
'(24)= 1000 - 4999
'(25)= 5000 +
'(26)= Spin           sound
'(27)= Change bet     sound
'(28)= MB             sound
'(29)= RJ             sound
'(30)= 1 - 4          sound
'(31)= 5 - 9          sound
'(32)= 10 - 24        sound
'(33)= 25 - 49        sound
'(34)= 50 - 99        sound
'(35)= 100  - 249     sound
'(36)= 250 - 999      sound
'(37)= 1000 - 4999    sound
'(38)= 5000 +         sound
'(39)= intro          mid
'(40)= pz 25 - 99     mid
'(41)= pz 100 - 249   mid
'(42)= pz 250 - 999   mid
'(43)= pz 1000 +      mid
'(44)=        mid
'(45)=        mid
'(46)=        mid
'(47)=        mid
'(48)=        mid
'(49)=        mid
'(50)=        mid


remembchkjackpot = False
remembchkmonback = False
confirmnewgame = False

setformpos Me

gametype.Caption = "MyReels: Profile Configuration Manager"      'Space(51)

getthumbspiccount thumbs, piccount, wheelvec, wheelorder


For pct = 1 To piccount
Thumblist.ListImages.Add (pct), , thumbs(pct)

picspare.PaintPicture Thumblist.ListImages(pct).Picture, 0, 0, 915, 915

imgsymbolsel(pct - 1).BorderStyle = 0
imgsymbolsel(pct - 1).Picture = picspare.Image
Set picspare = Nothing

Next
If piccount < 14 Then
For pct = piccount To 13
imgsymbolsel(pct).Visible = False
Next
End If

'Stopgap on old versions
If gt(152) > 0 And gt(152) < 50000 Then gt(152) = 50000

loadvalues

If gt(0) < 0 Then gt(0) = 0 'if game tolerance overflow

Form_Activate
End Sub
Public Sub loaddefaultz()
loadvalues
End Sub
Private Sub loadvalues()
procend = False



If gt(0) > 4 Then
cmdgametype(1).Enabled = False 'have to save config!
cmdgametype(6).Enabled = False
End If

If gt(49) + gt(50) + gt(51) + gt(150) = 0 Then
newpiccalcreturn = 2
Else
newpiccalcreturn = 0
End If

lblgametype(12) = CStr(gt(31)) & decsep & CStr(gt(32))

'Bet multiplier
For ct = 7 To 11    'Important - must zero these first
lblgametype(ct).Caption = 0
Next

For ct = 7 To 11
If gt(14 + ct) > 0 Then
If ct < 11 Then
spngametype(ct + 1).Enabled = True
lblgametype(ct + 1).Enabled = True
lblgametype(ct + 1).Caption = 0
End If
lblgametype(ct).Caption = gt(14 + ct)
End If
Next


'substitute multiplier care with the captions always >= 1 but gametypegen can be 0
If gt(4) > 0 Then
optgametype(2).Value = True
spngametype(5).Enabled = True
lblgametype(5).Enabled = True
lblgametype(6).Caption = "1"
lblgametype(5).Caption = gt(4)
Else
optgametype(3).Value = True
spngametype(6).Enabled = True
lblgametype(6).Enabled = True
lblgametype(5).Caption = "1"
lblgametype(6).Caption = gt(5)
End If

'paymode all or highest
If gt(9) = 0 Then
optgametype(0).Value = True
Else
optgametype(1).Value = True
End If


'Random jackpot
If gt(14) = 1 Then
chkgametype(0).Value = 1
For ct = 0 To 1
spngametype(ct).Enabled = True
Next
remembchkjackpot = True
lblgametype(1).Enabled = True
End If
'Disable label for starting amount but enable spin to reset if required
lblgametype(0).Caption = gt(19)
lblgametype(1).Caption = gt(15)



'Money back
If gt(10) > 0 Then
chkgametype(1).Value = 1
For ct = 3 To 4
spngametype(ct).Enabled = True
lblgametype(ct).Enabled = True
Next
remembchkmonback = True
End If
chkgametype(2).Value = gt(184)

lblgametype(3).Caption = gt(12)
lblgametype(4).Caption = gt(10)


gamspintot gamespintot


'save settings
If gt(152) <> 0 Then
For temp = 0 To 1
For ct = 1 To 9
tgs(temp, ct) = freegamesettings(temp, ct)
Next
For ct = 1 To 15
tss(temp, ct) = spinsettings(temp, ct)
Next
Next
For temp = 1 To 14
For ct = 1 To 14
tsub(temp, ct) = substitute(temp, ct)
Next
Next
End If

For ct = 0 To 3
tgamespinsymbol(ct) = gamespinsymbol(ct)
tgamespinkeep(ct) = gamespinkeep(ct)
Next



'Set initial value for starting money
leftgt155 = CLng(Left(CStr(gt(155)), 3))
If leftgt155 > 500 Then leftgt155 = 50
rightgt155 = CLng(Right(CStr(gt(155)), 3))
If rightgt155 < 500 Then
lblgametype(2).Caption = rightgt155
cashup rightgt155, zcolour, captval
cmdgametype(6).Caption = captval
cmdgametype(6).BackColor = zcolour
ElseIf rightgt155 = 500 Then
lblgametype(2).Caption = CStr(500)
cashup rightgt155, zcolour, captval
cmdgametype(6).Caption = captval
cmdgametype(6).BackColor = zcolour
Else
lblgametype(2).Caption = CStr(50)
cashup rightgt155, zcolour, captval
cmdgametype(6).Caption = captval
cmdgametype(6).BackColor = zcolour
End If
If zcolour = vbButtonFace Then
lblgametype(2).ForeColor = vbButtonText
Else
lblgametype(2).ForeColor = zcolour
End If
procend = True
End Sub
Private Sub Form_Activate()
Dim limit1 As Long, limit2 As Long, gamespinvec(4) As Long, oldgamespinkeep(4) As Long, oldgamespintot As Long, zmatch As Boolean


If Not ((gamespinsymbol(0) = tgamespinsymbol(0) And gamespinsymbol(1) = tgamespinsymbol(1) Or gamespinsymbol(0) = tgamespinsymbol(1) And gamespinsymbol(1) = tgamespinsymbol(0)) And (gamespinsymbol(2) = tgamespinsymbol(2) And gamespinsymbol(3) = tgamespinsymbol(3) Or gamespinsymbol(2) = tgamespinsymbol(3) And gamespinsymbol(3) = tgamespinsymbol(2)) And gamespintot > 1) Then


For ct = 0 To 3
oldgamespinkeep(ct) = gamespinkeep(ct)
Next
oldgamespintot = gamespintot


For ct = 1 To gamespintot
imggamespin(ct - 1).Visible = False
Next
For ct = 1 To 4
gamespinvec(ct) = 0
gamespinkeep(ct - 1) = 0
Next
gamespintot = 0
limit1 = 2
limit2 = 2
'merge free games spin symbols
For ct = 0 To 1
    If gamespinsymbol(ct) = 0 Then
    limit1 = ct
    Exit For
    End If
Next
For ct = 2 To 3
        If gamespinsymbol(ct) = 0 Then
        limit2 = ct - 2
        Exit For
        End If
Next
gamespintot = limit1 + limit2
For ct = 1 To limit1
gamespinvec(ct) = gamespinsymbol(ct - 1)
'set default gamespinkeep
gamespinkeep(ct - 1) = gamespinvec(ct)
Next
For ct = limit1 + 1 To gamespintot
gamespinvec(ct) = gamespinsymbol(ct + 1 - limit1)
gamespinkeep(ct - 1) = gamespinvec(ct)
Next
'outtahere if 1 - note gamespinkeep properly initialised if tot = 1
        If gamespintot < 2 Then
        labgametype(1).Visible = False
        For ct = 0 To 2
        lnegamespin(ct).Visible = False
        Next
        setvaluerange
        Exit Sub
        Else
        ShellSort gamespinvec, gamespintot
        For ct = 1 To gamespintot
        gamespinkeep(ct - 1) = gamespinvec(gamespintot - ct + 1)
        Next
        End If


'Compare with oldgamespinkeep
If gamespintot = oldgamespintot Then

For ct = 0 To gamespintot - 1
gamespinvec(ct + 1) = oldgamespinkeep(ct)
Next

ShellSort gamespinvec, gamespintot

'Perfect match or quit
For ct = 1 To gamespintot
zmatch = True
If gamespinvec(gamespintot - ct) <> gamespinkeep(ct - 1) Then
zmatch = False
Exit For
End If
Next
Else
zmatch = False
End If

If zmatch = True Then
For ct = 0 To 3
gamespinkeep(ct) = oldgamespinkeep(ct)
Next
End If

For ct = 0 To 2
lnegamespin(ct).Visible = True
Next

End If  'changes in gamespinsymbol

setvaluerange


For ct = 0 To gamespintot - 1
picspare.Height = 405
picspare.Width = 405
picspare.PaintPicture Thumblist.ListImages(gamespinkeep(ct)).Picture, 0, 0, 400, 400

imggamespin(ct).Visible = True
imggamespin(ct).Picture = picspare.Image
Next
labgametype(1).Visible = True

End Sub
Private Sub imggamespin_Click(Index As Integer)

'To recalculate return with MonteCarlo
For ct = 0 To 3
tgamespinkeep(ct) = gamespinkeep(ct)
Next


If Index < gamespintot - 1 Then 'not the topmost
picspare.Picture = imggamespin(Index + 1).Picture
imggamespin(Index + 1).Picture = imggamespin(Index).Picture
imggamespin(Index).Picture = picspare.Image
temp = gamespinkeep(Index + 1)
gamespinkeep(Index + 1) = gamespinkeep(Index)
gamespinkeep(Index) = temp
Else
picspare.Picture = imggamespin(0).Picture
imggamespin(0).Picture = imggamespin(gamespintot - 1).Picture
imggamespin(gamespintot - 1).Picture = picspare.Image
temp = gamespinkeep(0)
gamespinkeep(0) = gamespinkeep(gamespintot - 1)
gamespinkeep(gamespintot - 1) = temp
End If
End Sub
Private Sub imgsymbolsel_Click(Index As Integer)

If confirmnewgame = True Then
cmdgametype(6).ToolTipText = ""
cmdgametype(6).Caption = captval
confirmnewgame = False
End If


If newpiccalcreturn = 1 Then newpiccalcreturn = 2

savegamevars

If gt(0) > 5 Then gt(0) = 5

symbolselect = Index + 1

cmdgametype(4).ToolTipText = "This can take time a little time " 'Clear Scatter warning
cmdgametype(4).Caption = "&Recalculate Return Now"
cmdgametype(0).ToolTipText = "" 'Clear Scatter warning
cmdgametype(0).Caption = "Save && &Quit"


Load Sstab
gametype.Enabled = False
Sstab.Show
End Sub
Private Sub cmdgametype_Click(Index As Integer)
Dim VOGerror As Boolean


savegamevars

Select Case Index
Case 0

'nossubspingam NOT redundant here
If cmdgametype(4).BackColor = &H80FF& Then Exit Sub

If nossubspingam = True Or gt(152) <> 0 Then
    If gt(152) >= 0 Then
        If CLng(lblgametype(12).Caption) = 0 Then
            recalcnow
            Exit Sub
        ElseIf newpiccalcreturn = 2 Then
            recalcnow
            newpiccalcreturn = 1
            Exit Sub
        ElseIf changecalculated = False Then
            recalcnow
            Exit Sub
        ElseIf newpiccalcreturn = 1 Then
            'save current vars
            sstgtsav
            'OK, Don't want to execute Anychanges at all
        Else
            If Anychanges = True Then
            recalcnow
            Exit Sub
            End If
        End If
    Else
        recalcnow
        Exit Sub
    End If
Else  'zero VOG vars
    For ct = 29 To 34
    gt(ct) = 0
    Next
End If

changecalculated = True
On Error GoTo Notvaliddirectory 'in case it's deleted by user

ChDrive (Left$(loaddirectory, 1))
ChDir (loaddirectory)

LoadFrmSplsh 440

gamspintot gamespintot  'save fgame,fspin totals


gt(0) = 2
Unload gametype
Set gametype = Nothing
    If outputvars = True Then
    procend = False
    Load Pokemach
        If procend = True Then
        Pokemach.Show
        Dotaskwindow Pokemach
        Else
        Dotaskwindow , , True
        Unload Pokemach
        Set Pokemach = Nothing
        End If
    procend = False
    Else
    Unload frmSplsh
    End If
Case 1
On Error GoTo Notvaliddirectory 'in case it's deleted by user
ChDrive (Left$(loaddirectory, 1))
ChDir (loaddirectory)
LoadFrmSplsh 440
gt(0) = 1
Unload gametype
Set gametype = Nothing
procend = False
Load Pokemach
If procend = True Then
Pokemach.Show
Dotaskwindow Pokemach
Else
Unload Pokemach
Set Pokemach = Nothing
End If
procend = False
Unload frmSplsh
Case 2


If gt(0) >= 5 Then
'force back to new pics if succeeded in anychanges
gt(0) = 6 'as specified above
Else
gt(0) = 3
End If
Unload Me
Set gametype = Nothing
Load cornfig
Case 3
'force back to new pics if succeeded in anychanges
If gt(0) = 5 Then gt(0) = 6 'as specified above
Unload Me
Set gametype = Nothing
Load cfgthumb 'change gt(0) if gt(0) < 5 here
Case 4

Screen.MousePointer = vbHourglass
gametype.Enabled = False

cmdgametype(4).Caption = "Processing ...."

    If gt(152) = 0 Then
    calculatgamepercent wheelvec, piccount, VOGerror

    If VOGerror = True Then response = MsgBox("VOG Generation failed with negative probability or negative sum most likely due to excessive variation in scatter estimate. If not using scatters, please contact Author by email with attached Slotdata.s$t", vbOKOnly)

    Else
    If gt(152) < 0 Then gt(152) = -gt(152)

    If Montcarlo = False Then
    Screen.MousePointer = vbDefault
    gametype.Enabled = True
    cmdgametype(4).Caption = "&Recalculate Return Now"
    Exit Sub
    End If

For temp = 0 To 1
For ct = 1 To 9
tgs(temp, ct) = freegamesettings(temp, ct)
Next
For ct = 1 To 15
tss(temp, ct) = spinsettings(temp, ct)
Next
Next
For temp = 1 To 14
For ct = 1 To 14
tsub(temp, ct) = substitute(temp, ct)
Next
Next
For ct = 0 To 3
tgamespinsymbol(ct) = gamespinsymbol(ct)
tgamespinkeep(ct) = gamespinkeep(ct)
Next


End If

'Reset any warnings
cmdgametype(4).BackColor = vbButtonFace
cmdgametype(0).ToolTipText = ""
If cmdgametype(4).Caption <> "Well done!" Then cmdgametype(4).Caption = "&Recalculate Return Now"
gametype.Enabled = True
Screen.MousePointer = vbDefault
setvaluerange


If newpiccalcreturn = 2 Then newpiccalcreturn = 1
If True = True Then Anychanges

changecalculated = True

If gt(0) = 3 Then gt(0) = 4

Case 5
If confirmnewgame = True Then
cmdgametype(6).ToolTipText = ""
cmdgametype(6).Caption = captval
confirmnewgame = False
End If
Load Genopts
gametype.Enabled = False
Case 6
If confirmnewgame = False Then
cmdgametype(6).Caption = "Sure ?"
cmdgametype(6).ToolTipText = "Click to confirm & Scores will be written to Hall_of_Fame"
confirmnewgame = True
Else
'must clear STATS & Write to Hall of Fame etc
If OpenDb(sDatabaseName, 1) = False Then GoTo NotvalidHalldirectory


Set dbsCurrent = gdbCurrentDB

Set rectemp = dbsCurrent.OpenRecordset("Hall")

With rectemp

.MoveLast

.AddNew

On Error GoTo DBWriteerror

![Player] = Stringvars(8)

![Game] = Stringvars(5)

![Startcash] = leftgt155

![Endcash] = gt(2)

If gt(156) = 1 Then
![Reason] = 2
ElseIf gt(2) = 0 Then
![Reason] = 1
Else
![Reason] = 3
End If

![Date] = CDate(Date)

![Spintotal] = gt(49) + gt(50) + gt(51)


![ReelChange] = gt(26)

![SymbolChange] = gt(27)

![GeneralChange] = gt(28)

![StatsResets] = gt(150)


If gt(55) < 0 Then gt(55) = 0
![GameReplays] = gt(55)


If gt(133) > 0 Then
![SVOG] = ((gt(59) + gt(63) + gt(67) + gt(71) + gt(75) + gt(79) + gt(83) + gt(87)) / gt(133)) * 100
Else
![SVOG] = 0
End If

If lblgametype(12).Caption = "N/A" Then
![VOG] = 0
Else
    If VOGchg = 1 Then
    ![VOG] = -CSng(lblgametype(12).Caption)
    Else
    ![VOG] = CSng(lblgametype(12).Caption)
    End If
End If

If gt(152) < 0 Then gt(152) = -gt(152)

![MonteCarlo] = gt(152)


.Update


.Close
End With
Set rectemp = Nothing
Set dbsCurrent = Nothing
killdb sDatabaseName

If compactdb(1) = False Then GoTo NotvalidHalldirectory

NewGameNow

End If
End Select

Exit Sub

DEBUGsavequit:
response = MsgBox("Unknown SaveQuit Error.", vbOKOnly)
ShowError
Exit Sub

Notvaliddirectory:
response = MsgBox("Game directory: " & loaddirectory & " deleted or Slotdata.s$t corrupt or missing, please configure New Reel Pictures  ... ", vbOKOnly)
cmdgametype(0).Enabled = False
cmdgametype(1).Enabled = False
cmdgametype(2).Enabled = False
Unload frmSplsh
Exit Sub
NotvalidHalldirectory:
Set rectemp = Nothing
Set dbsCurrent = Nothing
killdb sDatabaseName

response = MsgBox("Hall_of_Fame corrupt or missing in chosen Directory: " & Stringvars(4) & ". Please use General Options Tab to set correct location. Game not restarted.", vbOKOnly)
confirmnewgame = False
cmdgametype(6).ToolTipText = ""
cmdgametype(6).Caption = captval
Exit Sub
DBWriteerror:
ShowError
response = MsgBox("As your locale settings seem to inhibit the writing of settings to the database, the scores on the old game cannot be saved.", vbOKOnly)
Set rectemp = Nothing
Set dbsCurrent = Nothing
killdb sDatabaseName
NewGameNow
End Sub
Private Sub NewGameNow()
gt(0) = 5 'no going back

For ct = 11 To 156
Select Case ct
Case 11, 13, 16 To 18, 26 To 34, 48 To 151, 154, 156
gt(ct) = 0
End Select
Next

'leave gt(191) alone also gt(18) = 0 (secret RJ)


genrandomseed True

VOGchg = 0
gt(55) = -1     'restarts with same RS
gt(56) = -1



gt(2) = rightgt155

On Error GoTo DEBUGRestore
Restoredefaults False

On Error GoTo DEBUGLoadvalues
loadvalues
changecalculated = False
cmdgametype(6).Caption = "New Game Started"
cmdgametype(4).Caption = "&Recalculate Return Now"
cmdgametype(4).BackColor = vbButtonFace
cmdgametype(6).ToolTipText = ""
lblgametype(12).ForeColor = &HFFFF00
confirmnewgame = False


Exit Sub
DEBUGRestore:
response = MsgBox("Restore error.", vbOKOnly)
ShowError
Exit Sub
DEBUGLoadvalues:
response = MsgBox("LoadValue error.", vbOKOnly)
ShowError
End Sub
Private Sub chkgametype_Click(Index As Integer)
If procend = False Then Exit Sub

If newpiccalcreturn = 1 Then newpiccalcreturn = 2
Select Case Index
Case 0
If remembchkjackpot = False Then
gt(16) = CLng(lblgametype(0).Caption)
gt(15) = CLng(lblgametype(1).Caption)

For ct = 0 To 1
spngametype(ct).Enabled = True
lblgametype(ct).Enabled = True
Next
remembchkjackpot = True
Else
gt(15) = 5
gt(16) = 500
gt(17) = 0
For ct = 0 To 1
spngametype(ct).Enabled = False
lblgametype(ct).Enabled = False
Next
remembchkjackpot = False
End If
Case 1
lblgametype(3).Caption = 1
If remembchkmonback = False Then
'Starts only on a payout of lblgametype(4) or less
For ct = 3 To 4
spngametype(ct).Enabled = True
lblgametype(ct).Enabled = True
Next
remembchkmonback = True
'set defaults
lblgametype(4).Caption = 10
Else
lblgametype(4).Caption = 0
gt(31) = 0
gt(32) = 0
gt(11) = 0 'current spin status
For ct = 3 To 4
spngametype(ct).Enabled = False
lblgametype(ct).Enabled = False
Next
remembchkmonback = False
End If
Case 2
gt(184) = chkgametype(2).Value
End Select
End Sub
Private Sub optgametype_Click(Index As Integer)
If procend = False Then Exit Sub

If newpiccalcreturn = 1 Then newpiccalcreturn = 2
cmdgametype(0).ToolTipText = ""
cmdgametype(0).Caption = "Save && &Quit"
Select Case Index
Case 0
If gt(152) = 0 Then
lblgametype(12).Caption = "N/A"
lblgametype(12).Enabled = False
labgametype(5).Enabled = False
cmdgametype(4).Enabled = False
End If
Case 1

setvaluerange

Case 2
spngametype(5).Enabled = True
lblgametype(5).Enabled = True
lblgametype(5).Caption = "1"
spngametype(6).Enabled = False
lblgametype(6).Caption = "1"
Case 3
spngametype(6).Enabled = True
lblgametype(6).Enabled = True
lblgametype(6).Caption = "1"
spngametype(5).Enabled = False
lblgametype(5).Caption = "1"
End Select
End Sub
Private Sub spngametype_SpinUp(Index As Integer)
temp = CLng(lblgametype(Index).Caption)
If newpiccalcreturn = 1 Then newpiccalcreturn = 2

Select Case Index
Case 0
'RJ influences MB!



If temp < 1000 Then
If temp > 249 Then
lblgametype(0).Caption = CStr(temp + 250)
ElseIf temp > 49 Then
lblgametype(0).Caption = CStr(temp + 50)
ElseIf temp > 9 Then
lblgametype(0).Caption = CStr(temp + 10)
Else
lblgametype(0).Caption = CStr(temp + 1)
End If
gt(16) = CLng(lblgametype(0).Caption)
gt(17) = 0 'Set RH component 0
End If
Case 1
If temp < 10 Then lblgametype(1).Caption = CStr(temp + 1)
Case 2
If temp < 500 Then

confirmnewgame = False  'need to reconfirm new game
cmdgametype(6).ToolTipText = ""
rightgt155 = temp + 50
lblgametype(2).Caption = rightgt155

cashup rightgt155, zcolour, captval
cmdgametype(6).Caption = captval
cmdgametype(6).BackColor = zcolour
If zcolour = vbButtonFace Then
lblgametype(2).ForeColor = vbButtonText
Else
lblgametype(2).ForeColor = zcolour
End If
End If

Case 3
If temp < 10 Then
lblgametype(3).Caption = CStr(temp + 1)

End If
Case 4
If temp < 15 Then
lblgametype(4).Caption = CStr(temp + 1)

End If
Case 5
If temp < 3 Then lblgametype(5).Caption = CStr(temp + 1)
Case 6
If temp < 3 Then lblgametype(6).Caption = CStr(temp + 1)
Case 7 To 10
If lblgametype(Index + 1).Caption > 0 And lblgametype(Index).Caption = lblgametype(Index + 1).Caption - 1 Then Exit Sub
gt(20) = 0  'Fixed on config exit
Select Case temp 'ensure values forming a range
Case 0
        If Index > 7 Then
        If lblgametype(Index - 1).Caption > 9 Then Exit Sub
        lblgametype(Index).Caption = CStr(lblgametype(Index - 1).Caption + 1)
        Else
        lblgametype(7).Caption = CStr(temp + 1)
        End If
        spngametype(Index + 1).Enabled = True
        lblgametype(Index + 1).Enabled = True
        lblgametype(Index + 1).Caption = CStr(0)
Case 1 To 9
lblgametype(Index).Caption = CStr(temp + 1)
End Select

Case 11
If temp = 0 Then
If CLng(lblgametype(10).Caption) < 10 Then
gt(20) = 0  'Fixed on config exit
lblgametype(11).Caption = CStr(lblgametype(10).Caption + 1)

End If
Else
If temp < 10 Then
gt(20) = 0  'Fixed on config exit
lblgametype(11).Caption = CStr(temp + 1)

End If
End If
End Select
End Sub
Private Sub spngametype_SpinDown(Index As Integer)
temp = CLng(lblgametype(Index).Caption)
If newpiccalcreturn = 1 Then newpiccalcreturn = 2

Select Case Index
Case 0

If temp > 0 Then
If temp > 250 Then
lblgametype(0).Caption = CStr(temp - 250)
ElseIf temp > 50 Then
lblgametype(0).Caption = CStr(temp - 50)
ElseIf temp > 10 Then
lblgametype(0).Caption = CStr(temp - 10)
ElseIf temp > 1 Then
lblgametype(0).Caption = CStr(temp - 1)
End If
gt(16) = CLng(lblgametype(0).Caption)
gt(17) = 0
End If

Case 1
If temp > 1 Then lblgametype(1).Caption = CStr(temp - 1)
Case 2

If temp > 50 Then

confirmnewgame = False  'need to reconfirm new game
cmdgametype(6).ToolTipText = ""
rightgt155 = temp - 50
lblgametype(2).Caption = rightgt155

cashup rightgt155, zcolour, captval
cmdgametype(6).Caption = captval
cmdgametype(6).BackColor = zcolour
If zcolour = vbButtonFace Then
lblgametype(2).ForeColor = vbButtonText
Else
lblgametype(2).ForeColor = zcolour
End If


End If
Case 3
If temp > 0 Then
lblgametype(3).Caption = CStr(temp - 1)

End If
Case 4
If temp > 2 Then
lblgametype(4).Caption = CStr(temp - 1)

End If
Case 5
If temp > 1 Then lblgametype(5).Caption = CStr(temp - 1)
Case 6
If temp > 1 Then lblgametype(6).Caption = CStr(temp - 1)
Case 7
If temp > 1 Then
lblgametype(7).Caption = CStr(temp - 1)

gt(20) = 0  'Fixed on config exit
End If
Case 8 To 11
'ensure values forming a range
If lblgametype(Index).Caption = lblgametype(Index - 1).Caption + 1 Then
gt(20) = 0  'Fixed on config exit
For ct = Index + 1 To 11
spngametype(ct).Enabled = False
lblgametype(ct).Enabled = False
lblgametype(ct).Caption = CStr(0)
Next
lblgametype(Index).Caption = CStr(0)

Else
If temp > 1 Then
gt(20) = 0  'Fixed on config exit
lblgametype(Index).Caption = CStr(temp - 1)

End If
End If
End Select
End Sub
Private Sub setvaluerange()
Dim argu1 As Long, argu2 As Long, captval1 As String

If nossubspingam = True Or gt(152) <> 0 Then
valuecomments argu1, argu2, zcolour, captval1
With lblgametype(12)
.Caption = inputformat(argu1)
.ForeColor = zcolour
.ToolTipText = "%Chance of no prize : " & gt(33) & decsep & gt(34)
End With
labgametype(5).Caption = captval1
End If
End Sub
Private Function nossubspingam()
Dim pct1 As Long

'highest win pays
If optgametype(1).Value = False Then
greyoutVOG
nossubspingam = False
Exit Function
End If

'Keep first game symbol for now
For ct = 0 To 3
If gamespinsymbol(ct) > 0 Then
greyoutVOG
nossubspingam = False
Exit Function
End If
Next
For pct = 1 To piccount
For pct1 = 1 To piccount
If substitute(pct, pct1) = True Then
greyoutVOG
nossubspingam = False
Exit Function
End If
Next
Next
lblgametype(12).Enabled = True
labgametype(5).Enabled = True
cmdgametype(4).Enabled = True
nossubspingam = True

End Function
Private Sub savegamevars()

'paymode
If optgametype(1).Value = True Then
gt(9) = 1
Else
gt(9) = 0
End If


'Bet multiplier

For ct = 7 To 11
gt(14 + ct) = CLng(lblgametype(ct).Caption)
Next
If gt(20) = 0 Then gt(20) = gt(21)      'bet mult

'substitute multiplier
If optgametype(2).Value = True Then 'subs
gt(4) = CLng(lblgametype(5).Caption)
gt(5) = 0
Else
gt(5) = CLng(lblgametype(6).Caption) 'nats
gt(4) = 0
End If

gt(14) = CLng(chkgametype(0).Value)
If gt(14) = 1 Then  'set random jackpot values
'only if user has clicked but not spun
gt(15) = CLng(lblgametype(1).Caption)
gt(19) = CLng(lblgametype(0).Caption)

'Get a new random jackpot value if just selected
If gt(16) = 0 Then gt(16) = gt(19)
If gt(19) = gt(16) And gt(17) = 0 Then        'Generate secret Random jackpot value
gt(18) = CLng(Int(20 * gt(16) * Rnd + 1))
End If
End If


'Money back
gt(10) = CLng(lblgametype(4).Caption)
gt(12) = CLng(lblgametype(3).Caption)
If gt(10) = 0 Then  'zero moneybackvog
gt(31) = 0
gt(32) = 0
End If

'new game
If gt(49) + gt(50) + gt(51) + gt(150) = 0 Then
gt(2) = CLng(lblgametype(2).Caption)
gt(155) = CLng(CStr(gt(2)) & CStr(gt(2)))
Else    'Always update starting money
gt(155) = CLng(CStr(leftgt155) & lblgametype(2).Caption)
End If


End Sub
Private Sub greyoutVOG()
Dim Dum As Long, neggt152 As Boolean
neggt152 = False

'No VOG as yet
labgametype(5).Caption = "Select Sample Size in General Options"

If gt(152) = 0 Then
    With lblgametype(12)
    .Caption = "N/A"
    .Enabled = False
    .ToolTipText = ""
    End With
    labgametype(5).Enabled = False
    With cmdgametype(4)
    .BackColor = vbButtonFace
    .Enabled = False
    End With
Else

    'gt(150) Stats reset
    If (gt(150) > 0 Or gt(49) + gt(50) + gt(51) > 0) Then
        If gamespinsymbol(0) = tgamespinsymbol(0) And gamespinsymbol(1) = tgamespinsymbol(1) Then
            For temp = 0 To 1
            For ct = 1 To 9
            If tgs(temp, ct) <> freegamesettings(temp, ct) Then neggt152 = True
            Next
            Next
        ElseIf gamespinsymbol(0) = tgamespinsymbol(1) And gamespinsymbol(1) = tgamespinsymbol(0) Then
            For temp = 0 To 1
    
                If temp = 0 Then
                Dum = 1
                Else
                Dum = 0
                End If
    
            For ct = 1 To 9
            If tgs(Dum, ct) <> freegamesettings(temp, ct) Then neggt152 = True
            Next
            Next

        Else
            neggt152 = True
        End If


        If gamespinsymbol(2) = tgamespinsymbol(2) And gamespinsymbol(3) = tgamespinsymbol(3) Then
            For temp = 0 To 1
            For ct = 1 To 15
            If tss(temp, ct) <> spinsettings(temp, ct) Then neggt152 = True
            Next
            Next
        ElseIf gamespinsymbol(2) = tgamespinsymbol(3) And gamespinsymbol(3) = tgamespinsymbol(2) Then
            For temp = 0 To 1
                If temp = 0 Then
                Dum = 1
                Else
                Dum = 0
                End If
            For ct = 1 To 15
            If tss(Dum, ct) <> spinsettings(temp, ct) Then neggt152 = True
            Next
            Next
        Else
            neggt152 = True
        End If


        For temp = 1 To 14
        For ct = 1 To 14
        If tsub(temp, ct) <> substitute(temp, ct) Then neggt152 = True
        Next
        Next
        For ct = 0 To 3
        If tgamespinkeep(ct) <> gamespinkeep(ct) Then neggt152 = True
        Next
        If neggt152 = True And gt(152) > 0 Then gt(152) = -gt(152)
    End If

labgametype(5).Enabled = True
lblgametype(12).Enabled = True
cmdgametype(4).Enabled = True
End If

End Sub
Private Sub recalcnow()
labgametype(5).Caption = "You need to Recalculate Return to continue"
cmdgametype(0).ToolTipText = "You need to Recalculate Return to continue"
cmdgametype(4).BackColor = &H80FF&
changecalculated = False
End Sub
Private Function Montcarlo()
Dim intreel As Long, ct1 As Long, prizetotal As Long, prizecount(2) As Long, bettotal As Long, totalprizevalue As Single
Dim currprize As Long, spinztemp(4) As Long, multiprize(2, 9, 3) As Long, picno As Long, multtemp As Long
Dim spincount As Long, freegamecount As Long, kept1or2 As Long, gamenotspin As Boolean, gamesaved As Boolean, spinsaved As Boolean, featurereset As Boolean
Dim bettest(126, 15) As Long, bmax(5) As Long, betvec(15) As Long, pzgtequalthan(10) As Long, oldspinz(4)    As Long
Dim ct2 As Long, ctrec As Long, VOG As Single, moneybackVOG As Single, Btotal As Long, ymax As Single, Y As Single, oldgt18 As Long, oldgt153 As Long, LH As Long, RH As Long
'Generate win()
totalprizevalue = 0
bettotal = 0
ctrec = 1
RH = gt(16)
LH = gt(17)
'Reserve response exclusively in this routine for RJ
response = vbNo

If gt(14) = 1 Then  'check Random jackpot limit
oldgt18 = gt(18)
gt(18) = Int(20 * RH * Rnd + 1)
On Error GoTo Erroverflowjackink
If RH + 1 + (gt(152) / 10) * gt(15) > gt(47) Then
response = MsgBox("Random Jackpot prize too high for simulation. Press Yes to add 5% to VOG - changes to the probability of no prizes are in most cases miniscule, so ignored at least in this version. To attempt simulation with Random Jackpot, reset Random jackpot prize or reduce Random Jackpot increment or deselect Random Jackpot feature or increase Game Limits or reduce MonteCarlo value", vbYesNo)
If response = vbNo Then GoTo Erroverflowjackink
End If
End If



zerogamspnvars

For ct = 1 To 10
pzgtequalthan(ct) = 0
Next

'Generate new seed
Randomize CSng((1 + 2 * CSng(Time)) * CSng(Date) ^ (1 + 2 * CSng(Time)))
dirofspin = 1

For ct = 0 To 4
hreel(ct) = False
oldspinz(ct) = spinz(ct)
dirspin(ct) = 1
Next
oldgt153 = gt(153)  'must be 0 for sortsubs

gt(153) = 0



For ct = 1 To gt(152)

If gt(14) = 1 Then  'Random jackpot
RH = RH + gt(15)
        If RH > 9 Then
        LH = LH + Int(RH / 10)
        RH = RH - 10 * Int(RH / 10)
        End If
End If



For intreel = 0 To 4
'subshere(intreel) = False generate random nos

'Now figure out what goes on the held reels
If hreel(intreel) = True Then
spinz(intreel) = hq(0, intreel)
Else


'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
spinz(intreel) = Int(24 * Rnd + 1)

End If

'snapshot the top line
hq(0, intreel) = spinz(intreel)


Next



sortsubstitutes prizecount, multiprize

multtemp = 1
prizetotal = 0
currprize = 0

    For ct1 = 1 To prizecount(0)
        picno = multiprize(0, ct1, 2)
        ct2 = sst(picno, multiprize(0, ct1, 1))
        If multiprize(0, ct1, 3) = 1 Then
                If gt(4) > 0 Then
                multtemp = gt(4)
                Else    'gt(5) > 0
                multtemp = gt(5)
                End If
        currprize = multtemp * ct2
        Else    'not in a substitute family
        currprize = ct2
        End If
        prizetotal = currprize + prizetotal
    Next




'Non - paying free games & spins
If prizetotal = 0 Then
    
    If spincount > 0 Then
        If hq(0, 5) = gamespinsymbol(2) Then
        prizetotal = spinsettings(0, 13)
        Else
        prizetotal = spinsettings(1, 13)
        End If
    ElseIf freegamecount > 0 Then
        If hq(0, 5) = gamespinsymbol(0) Then
        prizetotal = freegamesettings(0, 7)
        Else
        prizetotal = freegamesettings(1, 7)
        End If
    End If


Else
ct1 = 1
    If spincount > 0 Then
        If hq(0, 5) = gamespinsymbol(2) Then
        ct1 = spinsettings(0, 12)
        Else
        ct1 = spinsettings(1, 12)
        End If
    ElseIf freegamecount > 0 Then
        If hq(0, 5) = gamespinsymbol(0) Then
        ct1 = freegamesettings(0, 6)
        Else
        ct1 = freegamesettings(1, 6)
        End If
    End If
prizetotal = ct1 * prizetotal
End If

If freegamecount > 0 Then
freegamecount = freegamecount + 1
ElseIf spincount > 0 Then
spincount = spincount + 1
Else
bettotal = bettotal + 1
End If

If gt(14) = 1 And response = vbNo Then 'Random jackpot
    If gt(18) = CLng(2 * Rnd * (10 * LH + RH) + 1) Then
    prizetotal = LH + prizetotal
 
    'To try to be as fair as possible with the middle digit of RH
        If prizetotal < gt(47) Then
            If RH > 5 Then
            prizetotal = prizetotal + 1
            ElseIf RH = 5 Then
            If Right(gt(49) + gt(50) + gt(51), 1) > 5 Then prizetotal = prizetotal + 1
            End If
        End If
    LH = gt(19)
    RH = 0
    End If
End If

If prizetotal >= 1 Then pzgtequalthan(1) = pzgtequalthan(1) + 1
For ct1 = 2 To gt(12)
If prizetotal >= ct1 Then pzgtequalthan(ct1) = pzgtequalthan(ct1) + 1
Next


totalprizevalue = totalprizevalue + prizetotal

eligiblespingame freegamecount, spincount, kept1or2, prizetotal, gamenotspin, gamesaved, spinsaved, featurereset

For ct1 = 0 To 1
If ct1 = kept1or2 Then
        If freegamecount > freegamesettings(ct1, 8) Then
                If clearq(kept1or2, freegamecount, spincount, gamenotspin) = True Then Exit For
        ElseIf spincount > spinsettings(ct1, 14) Then
                For intreel = 0 To 4
                hreel(intreel) = False
                Next
                If clearq(kept1or2, freegamecount, spincount, gamenotspin) = True Then Exit For
        End If
End If
Next


If spincount > 0 Then

For intreel = 0 To 4
       
spinztemp(intreel) = hq(0, intreel) 'top line going down
        
'For odd spin dups
        If spinztemp(intreel) >= 0 Then
        
        ct1 = wheelorder(intreel, Advanz(spinztemp(intreel), -1))
        
            If ct1 = hq(0, 5) Or (substitute(ct1, hq(0, 5)) = True And reelcheck(ct1, intreel + 1) = True) Then
            hreel(intreel) = True
            ElseIf sst(hq(0, 5), 2) = 1 Then
            'Do extra test for scatter
        
                If wheelorder(intreel, Advanz(spinztemp(intreel), 1)) = hq(0, 5) Then
                hreel(intreel) = True
                ElseIf wheelorder(intreel, Advanz(spinztemp(intreel), -2 * 1)) = hq(0, 5) Then
                'Advanz to row 3
                hreel(intreel) = True
                Else
                hreel(intreel) = False
                End If
        
            Else
            hreel(intreel) = False
            End If
        Else
        hreel(intreel) = False
        End If
Next
End If



If ct = 5000 * ctrec Then
cmdgametype(4).Caption = "Processed " & ctrec * 5000
ctrec = ctrec + 1
End If
Next





For ct = 0 To 4
spinz(ct) = oldspinz(ct)
Next

If gt(152) > 5000000 Then cmdgametype(4).Caption = "Well done!"


estp0 = 100 * (1 - pzgtequalthan(1) / (gt(152)))
VOG = 100 * totalprizevalue / bettotal

gt(153) = oldgt153  'return proper gt153
If gt(14) = 1 Then
gt(18) = oldgt18  'restore secret value
If response = vbYes Then VOG = VOG + 0.05   'Note response yes if over limit
'p0 needs to be calculated on expectation
End If



gt(154) = pzgtequalthan(gt(12))


getBmax bmax

'Moneyback
If gt(10) > 0 And estp0 < 100 Then


If VOG < 100 Then

'Find Btotal here

For Btotal = gt(10) - 1 + bmax(5) To bmax(5) * gt(10)
If moneybackvalue(VOG, estp0, Y, bettest, Btotal) = True Then
If Y > ymax Then ymax = Y
End If
Next

Else
'if VOG > 100, always bet MAX
For ct = 1 To gt(10)
bettest(0, ct) = bmax(5)
Next
If moneybackvalue(VOG, estp0, Y, bettest, gt(10) * bmax(5)) = True Then ymax = Y
End If

End If



If VOG > 0 Then
outputforformat VOG, 29, 137

If gt(12) > 0 And gt(10) > 0 Then
'Find mean time before getting prize is 1/probg12
    If gt(152) = 0 Then 'No Monte Carlo
    moneybackVOG = VOG * ((1 - 2 * (estp0 / 100) ^ (24 ^ 5 / pzgtequalthan(gt(12))) + ymax * 2 * ((estp0 / 100) ^ (24 ^ 5 / pzgtequalthan(gt(12))))))
    Else
    moneybackVOG = VOG * ((1 - 2 * (estp0 / 100) ^ (gt(152) / pzgtequalthan(gt(12))) + ymax * 2 * ((estp0 / 100) ^ (gt(152) / pzgtequalthan(gt(12))))))
    End If
Else
moneybackVOG = ymax * VOG
End If

outputforformat moneybackVOG, 31, 138


outputforformat estp0, 33, 139
End If

zerogamspnvars
Montcarlo = True

Exit Function
Erroverflowjackink:
VOG = 0
Montcarlo = False
End Function
