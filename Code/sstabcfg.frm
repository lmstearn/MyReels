VERSION 5.00
Object = "{10AB04E3-3AEA-101C-96E6-0020AF38F4BB}#1.0#0"; "TEGSPIN3.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Sstab 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyReels: Picture Properties"
   ClientHeight    =   4335
   ClientLeft      =   2205
   ClientTop       =   2670
   ClientWidth     =   8070
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
   HelpContextID   =   8
   Icon            =   "sstabcfg.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8070
   Visible         =   0   'False
   Begin TabDlg.SSTab sstaboption 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "sstabcfg.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Free Games"
      TabPicture(1)   =   "sstabcfg.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Free Spins (1)"
      TabPicture(2)   =   "sstabcfg.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frame1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Free Spins (2)"
      TabPicture(3)   =   "sstabcfg.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frame1(3)"
      Tab(3).ControlCount=   1
      Begin VB.Frame frame1 
         BorderStyle     =   0  'None
         Height          =   3735
         Index           =   3
         Left            =   -74880
         TabIndex        =   32
         Top             =   360
         Width           =   7815
         Begin VB.CheckBox mischkgamespin 
            Caption         =   "Current Free Spin/Game sequence not reset on this new Free Spin combination. "
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   35
            Top             =   1200
            Width           =   5655
         End
         Begin VB.CheckBox mischkgamespin 
            Caption         =   "Free Spin sequence begins only if held Pictures do not contribute to a prize"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   34
            Top             =   120
            Width           =   5535
         End
         Begin VB.CheckBox mischkgamespin 
            Caption         =   "Only Natural Pictures (no Substituters in combination) will generate a Free Spin sequence"
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   33
            Top             =   600
            Width           =   4455
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   420
            Index           =   12
            Left            =   4440
            TabIndex        =   77
            Top             =   1680
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   732
            _StockProps     =   64
            Enabled         =   0   'False
            BevelWidth      =   3
            Interval        =   125
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   420
            Index           =   11
            Left            =   3360
            TabIndex        =   79
            Top             =   3120
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   732
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   125
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   10
            Left            =   4300
            TabIndex        =   84
            Top             =   2640
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   125
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   9
            Left            =   3960
            TabIndex        =   88
            Top             =   2160
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   125
         End
         Begin VB.Label lblfreespin 
            Caption         =   "The number of cumulative Free Spin sequences does not exceed"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   90
            Top             =   1800
            Width           =   4455
         End
         Begin VB.Label lblgamspn 
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
            Left            =   4320
            TabIndex        =   87
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label lblgamspn 
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
            Index           =   6
            Left            =   3720
            TabIndex        =   85
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label lblgamspn 
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
            Index           =   5
            Left            =   4670
            TabIndex        =   83
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label lblfreespin 
            Caption         =   "During a Free Spin sequence, the Reels are spun"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   76
            Top             =   3240
            Width           =   3255
         End
         Begin VB.Label lblfreespin 
            Caption         =   "times original bet"
            Height          =   255
            Index           =   3
            Left            =   5160
            TabIndex        =   75
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblfreespin 
            Caption         =   "Every non winning combination in a Free Spin sequence pays"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   74
            Top             =   2760
            Width           =   4240
         End
         Begin VB.Label lblfreespin 
            Caption         =   "times usual prize"
            Height          =   255
            Index           =   1
            Left            =   4800
            TabIndex        =   73
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblgamspn 
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
            Index           =   7
            Left            =   4800
            TabIndex        =   72
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblfreespin 
            Caption         =   "Every winning combination in a Free Spin sequence pays"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   2280
            Width           =   3975
         End
         Begin VB.Label lblnumberofspins2 
            Caption         =   "times"
            Height          =   255
            Left            =   4200
            TabIndex        =   36
            Top             =   3240
            Width           =   495
         End
         Begin ComctlLib.ImageList tinythumb 
            Left            =   6480
            Top             =   3000
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
      End
      Begin VB.Frame frame1 
         BorderStyle     =   0  'None
         Height          =   3735
         Index           =   2
         Left            =   -75000
         TabIndex        =   31
         Top             =   360
         Width           =   7935
         Begin VB.Frame frafreespin3 
            Height          =   1575
            Left            =   3960
            TabIndex        =   66
            Top             =   0
            Width           =   3900
            Begin VB.OptionButton optfoursnofreespin 
               Caption         =   "  No Free Spin on fours combinations"
               Height          =   210
               Left            =   120
               TabIndex        =   70
               Top             =   165
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optfoursany 
               Caption         =   "Free Spin for fours on any reel"
               Height          =   255
               Left            =   120
               TabIndex        =   69
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CheckBox chkspin 
               Caption         =   "Free Spin for fours on reels 2 && 3 && 4 && 5 "
               Enabled         =   0   'False
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   68
               Top             =   840
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.OptionButton optfoursleft 
               Caption         =   "Free Spin for fours on reels 1 && 2 && 3 && 4"
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   480
               Width           =   3615
            End
         End
         Begin VB.Frame frafreespin2 
            Height          =   1815
            Left            =   0
            TabIndex        =   60
            Top             =   1920
            Width           =   3735
            Begin VB.OptionButton opttriplesany 
               Caption         =   "Free Spin for triples on any reel"
               Height          =   210
               Left            =   120
               TabIndex        =   65
               Top             =   1550
               Width           =   2655
            End
            Begin VB.OptionButton opttriplesleft 
               Caption         =   "Free Spin for triples on reels 1 && 2 && 3"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   480
               Width           =   3255
            End
            Begin VB.OptionButton opttriplesnofreespin 
               Caption         =   "No Free Spin on triple combinations"
               Height          =   210
               Left            =   120
               TabIndex        =   63
               Top             =   165
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.CheckBox chkspin 
               Caption         =   "Free Spin for triples on reels 3 && 4 && 5 "
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   62
               Top             =   840
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.CheckBox chkspin 
               Caption         =   "Free Spin for triples together in middle"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   61
               Top             =   1200
               Visible         =   0   'False
               Width           =   3015
            End
         End
         Begin VB.Frame frafreespin 
            Height          =   1815
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   3735
            Begin VB.OptionButton optpairnofreespin 
               Caption         =   "No Free Spin on pair combinations"
               Height          =   210
               Left            =   120
               TabIndex        =   59
               Top             =   165
               Value           =   -1  'True
               Width           =   2775
            End
            Begin VB.OptionButton optpairleft 
               Caption         =   "Free Spin for pairs on reels 1 && 2"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   480
               Width           =   2775
            End
            Begin VB.CheckBox chkspin 
               Caption         =   "Free Spin for pairs on reels 4 && 5"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   57
               Top             =   840
               Visible         =   0   'False
               Width           =   2730
            End
            Begin VB.OptionButton optpairsany 
               Caption         =   "Free Spin for pairs on any reel"
               Height          =   210
               Left            =   120
               TabIndex        =   56
               Top             =   1550
               Width           =   2655
            End
            Begin VB.CheckBox chkspin 
               Caption         =   "Free Spin for pairs together in middle"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   55
               Top             =   1200
               Visible         =   0   'False
               Width           =   2895
            End
         End
      End
      Begin VB.Frame frame1 
         BorderStyle     =   0  'None
         Height          =   3735
         Index           =   1
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   7815
         Begin VB.Frame frafreegame 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   5055
            Begin VB.CheckBox mischkgamespin 
               Caption         =   "Free Games for triples together in the middle"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   53
               Top             =   720
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.CheckBox mischkgamespin 
               Caption         =   "Free Games for triples or more ending on reels 3 && 4 && 5"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Visible         =   0   'False
               Width           =   4335
            End
            Begin VB.OptionButton optfreegame123 
               Caption         =   "Free Games for triples or more starting on reels 1 && 2 && 3"
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Visible         =   0   'False
               Width           =   4455
            End
            Begin VB.OptionButton optfreegameany 
               Caption         =   "Free Games for triples or more on any reel"
               Enabled         =   0   'False
               Height          =   255
               Left            =   0
               TabIndex        =   50
               Top             =   1080
               Visible         =   0   'False
               Width           =   3375
            End
         End
         Begin VB.CheckBox mischkgamespin 
            Caption         =   $"sstabcfg.frx":007C
            Enabled         =   0   'False
            Height          =   615
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   2640
            Width           =   5535
         End
         Begin VB.CheckBox mischkgamespin 
            Caption         =   "Only Natural Pictures (no Substituters in combination) will generate a Free Game sequence"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   3240
            Width           =   6975
         End
         Begin VB.CheckBox chkenablefreegame 
            Caption         =   "Enable Free Game Feature"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   2175
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   0
            Left            =   2880
            TabIndex        =   40
            Top             =   1800
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            Enabled         =   0   'False
            BevelWidth      =   3
            Interval        =   125
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   1
            Left            =   5760
            TabIndex        =   42
            Top             =   1800
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            Enabled         =   0   'False
            BevelWidth      =   3
            Interval        =   125
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   2
            Left            =   3480
            TabIndex        =   80
            Top             =   2205
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            Enabled         =   0   'False
            BevelWidth      =   3
            Interval        =   125
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   3
            Left            =   5760
            TabIndex        =   89
            Top             =   2760
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            Enabled         =   0   'False
            BevelWidth      =   3
            Interval        =   125
         End
         Begin VB.Label lblgamspn 
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
            Index           =   3
            Left            =   6120
            TabIndex        =   86
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lblgamspn 
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
            Index           =   2
            Left            =   3840
            TabIndex        =   82
            Top             =   2205
            Width           =   375
         End
         Begin VB.Label lblfreegame 
            Caption         =   "times original bet"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   6600
            TabIndex        =   71
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblfreegame 
            Caption         =   "During a Free Games sequence, the reels are spun"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   48
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblfreegame 
            Caption         =   "During a Free Game sequence, a win pays"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   47
            Top             =   1920
            Width           =   2820
         End
         Begin VB.Label lblfreegame 
            Caption         =   "times usual prize, nowin pays"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   46
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblfreegame 
            Caption         =   "times"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   45
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label lblgamspn 
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
            Index           =   0
            Left            =   3240
            TabIndex        =   44
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label lblgamspn 
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
            Index           =   1
            Left            =   6120
            TabIndex        =   43
            Top             =   1800
            Width           =   375
         End
      End
      Begin VB.Frame frame1 
         BorderStyle     =   0  'None
         Height          =   3855
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7815
         Begin VB.CheckBox chkgeneral 
            Caption         =   "Picture is a Scatter"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2775
         End
         Begin VB.Frame optpaypattern 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   0
            TabIndex        =   24
            Top             =   600
            Width           =   3855
            Begin VB.OptionButton optgeneral 
               Height          =   210
               Index           =   0
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox chkgeneral 
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   27
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox chkgeneral 
               Height          =   255
               Index           =   2
               Left            =   1920
               TabIndex        =   26
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton optgeneral 
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   25
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.PictureBox Pictinythumb 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   2520
            ScaleHeight     =   975
            ScaleWidth      =   975
            TabIndex        =   8
            Top             =   1560
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdok 
            Cancel          =   -1  'True
            Caption         =   "&Accept"
            Default         =   -1  'True
            Height          =   615
            Left            =   2880
            TabIndex        =   7
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CheckBox chkreels 
            Caption         =   "reel1"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   6
            Top             =   3600
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkreels 
            Caption         =   "reel2"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   4680
            TabIndex        =   5
            Top             =   3600
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkreels 
            Caption         =   "reel3"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   5400
            TabIndex        =   4
            Top             =   3600
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkreels 
            Caption         =   "reel4"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   6120
            TabIndex        =   3
            Top             =   3600
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkreels 
            Caption         =   "reel5"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   6840
            TabIndex        =   2
            Top             =   3600
            Value           =   1  'Checked
            Width           =   735
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   420
            Index           =   4
            Left            =   1680
            TabIndex        =   11
            Top             =   1440
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   732
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   100
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   420
            Index           =   5
            Left            =   1680
            TabIndex        =   12
            Top             =   1920
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   732
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   100
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   420
            Index           =   6
            Left            =   1680
            TabIndex        =   13
            Top             =   2400
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   732
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   100
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   8
            Left            =   1680
            TabIndex        =   78
            Top             =   3360
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   100
         End
         Begin TegspinLibCtl.TegoSpin spngamspn 
            Height          =   360
            Index           =   7
            Left            =   1680
            TabIndex        =   81
            Top             =   2880
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   635
            _StockProps     =   64
            BevelWidth      =   3
            Interval        =   100
         End
         Begin VB.Label labmultiplies 
            Caption         =   "X 5 ="
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   23
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label labmultiplies 
            Caption         =   "X 4 ="
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   22
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label labmultiplies 
            Caption         =   "X 3 ="
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   21
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label labmultiplies 
            Caption         =   "X 2 ="
            Height          =   375
            Index           =   3
            Left            =   720
            TabIndex        =   20
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label labmultiplies 
            Caption         =   "X 1 ="
            Height          =   375
            Index           =   4
            Left            =   720
            TabIndex        =   19
            Top             =   3360
            Width           =   855
         End
         Begin VB.Image imgthumbprice 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Index           =   0
            Left            =   120
            Top             =   1440
            Width           =   405
         End
         Begin VB.Image imgthumbprice 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Index           =   1
            Left            =   120
            Top             =   1920
            Width           =   405
         End
         Begin VB.Image imgthumbprice 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Index           =   2
            Left            =   120
            Top             =   2400
            Width           =   405
         End
         Begin VB.Image imgthumbprice 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Index           =   3
            Left            =   120
            Top             =   2880
            Width           =   405
         End
         Begin VB.Image imgthumbprice 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   405
            Index           =   4
            Left            =   120
            Top             =   3360
            Width           =   405
         End
         Begin VB.Label lblprize 
            Height          =   420
            Index           =   0
            Left            =   2040
            TabIndex        =   18
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblprize 
            Height          =   420
            Index           =   1
            Left            =   2040
            TabIndex        =   17
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lblprize 
            Height          =   420
            Index           =   2
            Left            =   2040
            TabIndex        =   16
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lblprize 
            Height          =   420
            Index           =   3
            Left            =   2040
            TabIndex        =   15
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label lblprize 
            Height          =   420
            Index           =   4
            Left            =   2040
            TabIndex        =   14
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label labsubstitute 
            Caption         =   "substitutes for "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   10
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label labsubstitute 
            Caption         =   "on"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   9
            Top             =   2520
            Width           =   495
         End
         Begin VB.Image imgsymbolsel 
            Appearance      =   0  'Flat
            Height          =   915
            Left            =   3960
            Top             =   240
            Width           =   915
         End
         Begin VB.Line lnesubstitute 
            BorderWidth     =   2
            Index           =   0
            Tag             =   " "
            X1              =   3960
            X2              =   4800
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line lnesubstitute 
            BorderWidth     =   2
            Index           =   1
            Tag             =   " "
            X1              =   4560
            X2              =   4800
            Y1              =   1800
            Y2              =   2040
         End
         Begin VB.Line lnesubstitute 
            BorderWidth     =   2
            Index           =   2
            Tag             =   " "
            X1              =   4560
            X2              =   4800
            Y1              =   2280
            Y2              =   2040
         End
         Begin VB.Line lnesubstitute 
            BorderWidth     =   2
            Index           =   3
            Tag             =   " "
            X1              =   4800
            X2              =   4800
            Y1              =   2520
            Y2              =   3240
         End
         Begin VB.Line lnesubstitute 
            BorderWidth     =   2
            Index           =   4
            Tag             =   " "
            X1              =   4560
            X2              =   4800
            Y1              =   3000
            Y2              =   3240
         End
         Begin VB.Line lnesubstitute 
            BorderWidth     =   2
            Index           =   5
            Tag             =   " "
            X1              =   4800
            X2              =   5040
            Y1              =   3240
            Y2              =   3000
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   0
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   360
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   1
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   840
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   2
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   3
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   1800
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   4
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   2280
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   5
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   2760
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   6
            Left            =   5160
            Shape           =   3  'Circle
            Top             =   3240
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   7
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   360
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   8
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   840
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   9
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   1320
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   10
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   1800
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   11
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   2280
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   12
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   2760
            Width           =   135
         End
         Begin VB.Shape tinyshape 
            Height          =   255
            Index           =   13
            Left            =   6360
            Shape           =   3  'Circle
            Top             =   3240
            Width           =   135
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   0
            Left            =   5520
            Top             =   240
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   1
            Left            =   5520
            Top             =   720
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   2
            Left            =   5520
            Top             =   1200
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   3
            Left            =   5520
            Top             =   1680
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   4
            Left            =   5520
            Top             =   2160
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   5
            Left            =   5520
            Top             =   2640
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   6
            Left            =   5520
            Top             =   3120
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   7
            Left            =   6720
            Top             =   240
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   8
            Left            =   6720
            Top             =   720
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   9
            Left            =   6720
            Top             =   1200
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   10
            Left            =   6720
            Top             =   1680
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   11
            Left            =   6720
            Top             =   2160
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   12
            Left            =   6720
            Top             =   2640
            Width           =   405
         End
         Begin VB.Image Imgtinythumb 
            Appearance      =   0  'Flat
            Height          =   405
            Index           =   13
            Left            =   6720
            Top             =   3120
            Width           =   405
         End
      End
   End
End
Attribute VB_Name = "Sstab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thumbs(14) As StdPicture, wheelvec(5, 14) As Long, wheelorder(4, 24) As Long
Dim tgs(1, 9) As Long, tsp(1, 15) As Long
Dim piccount As Long, ct As Long, intreel As Long, frstsec As Long
Dim remembersubstitute As Long, boodisablefreetab As Boolean, intspinchoice As Long
Dim intoptionchoice As Long, hadagame As Long, boofirstorsec(3) As Boolean
Dim hadaspin As Long, wantaspinno As Long, temp As Long
Dim boojustchecked As Boolean, nomorespinup As Boolean
Dim scatterspintemp1 As Long, scatterspintemp2 As Long, spinstatus As Long
Dim symbLR As Long, symbRL As Long, symbANY As Long, symbMID As Long
'sst(pct,0) = 0 5 symbols
'sst(pct,0) = 1 1 symbol
'sst(pct,0) = 2 Pair together on left
'sst(pct,0) = 3 Pair together on right
'sst(pct,0) = 4 Pair together in middle
'sst(pct,0) = 5 Pair in any other combo
'sst(pct,0) = 6 triple together on left
'sst(pct,0) = 7 triple together on right
'sst(pct,0) = 8 triple together in middle
'sst(pct,0) = 9 triple with pair together on left
'sst(pct,0) = 10 triple with pair together on right
'sst(pct,0) = 11 triple with pair together in middle
'sst(pct,0) = 12 triple in any other combo
'sst(pct,0) = 13 fours together on left
'sst(pct,0) = 14 fours together on right
'sst(pct,0) = 15 fours split in middle
'sst(pct,0) = 16 fours 3 from left
'sst(pct,0) = 17 fours 3 from right
'sst(pct,1) = 1 L-R
'sst(pct,2) = 1 Scatter
'sst(pct,3) = 1 R-L
'sst(pct,4) = 1 Middle Threes
'sst(pct,5) = 1 Any


Private Sub Form_Load()
procend = False

setformpos Me

boojustchecked = False
nomorespinup = False
spinstatus = 0

Sstab.Caption = "Picture Properties"    'Space(35)

getthumbspiccount thumbs, piccount, wheelvec, wheelorder


If sst(symbolselect, 6) > 9999 Then nomorespinup = True

For ct = 0 To 3
boofirstorsec(ct) = False
Next
hadaspin = 0
hadagame = 0
'restore old free game settings

If gamespinsymbol(0) = symbolselect Then boofirstorsec(0) = True
If gamespinsymbol(1) = symbolselect Then boofirstorsec(1) = True
If boofirstorsec(0) = False And boofirstorsec(1) = False Then
For frstsec = 0 To 1
For ct = 1 To 9
tgs(frstsec, ct) = 0
Next
Next
If gamespinsymbol(1) > 0 Then
hadagame = 2
ElseIf gamespinsymbol(0) > 0 Then
hadagame = 1
End If
ElseIf boofirstorsec(1) = True Then
For ct = 1 To 9
tgs(1, ct) = freegamesettings(1, ct)
Next
Else
If gamespinsymbol(1) > 0 Then
'swap symbolselect
For ct = 1 To 9
temp = freegamesettings(0, ct)
freegamesettings(0, ct) = freegamesettings(1, ct)
freegamesettings(1, ct) = temp
tgs(1, ct) = freegamesettings(1, ct)
Next
'swap games as well
temp = gamespinsymbol(0)
gamespinsymbol(0) = gamespinsymbol(1)
gamespinsymbol(1) = temp
boofirstorsec(0) = False
boofirstorsec(1) = True
Else
For ct = 1 To 9
tgs(0, ct) = freegamesettings(0, ct)
Next
End If
End If

'now restore spinsettings
If gamespinsymbol(2) = symbolselect Then boofirstorsec(2) = True
If gamespinsymbol(3) = symbolselect Then boofirstorsec(3) = True

        If boofirstorsec(2) = False And boofirstorsec(3) = False Then
                Sstab.sstaboption.TabEnabled(3) = False
                For frstsec = 0 To 1
                For ct = 1 To 15
                tsp(frstsec, ct) = 0
                Next
                Next
                If gamespinsymbol(3) > 0 Then
                hadaspin = 2
                ElseIf gamespinsymbol(2) > 0 Then
                hadaspin = 1
                End If

        ElseIf boofirstorsec(3) = True Then
                spinstatus = 2  'don't need to confirm tab 3 etc
                For ct = 1 To 15
                tsp(1, ct) = spinsettings(1, ct)
                Next
        Else
                spinstatus = 2
                If gamespinsymbol(3) > 0 Then
                'swap symbolselect
                For ct = 1 To 15
                temp = spinsettings(0, ct)
                spinsettings(0, ct) = spinsettings(1, ct)
                spinsettings(1, ct) = temp
                tsp(1, ct) = spinsettings(1, ct)
                Next
                'swap spins as well
                temp = gamespinsymbol(2)
                gamespinsymbol(2) = gamespinsymbol(3)
                gamespinsymbol(3) = temp
                boofirstorsec(2) = False
                boofirstorsec(3) = True
                Else
                For ct = 1 To 15
                tsp(0, ct) = spinsettings(0, ct)
                Next
                End If
        End If

If disablegamespintabs(0) = False And disablegamespintabs(1) = False Then
    If boofirstorsec(2) = True Then
    Sstab.sstaboption.TabEnabled(1) = False
    ElseIf boofirstorsec(0) = True Then
    Sstab.sstaboption.TabEnabled(2) = False
    Else
    End If
ElseIf disablegamespintabs(0) = True And disablegamespintabs(1) = True Then
    If boofirstorsec(2) = True Or boofirstorsec(3) = True Then
    Sstab.sstaboption.TabEnabled(1) = False
    ElseIf boofirstorsec(0) = True Or boofirstorsec(1) = True Then
    Sstab.sstaboption.TabEnabled(2) = False
    Else
    Sstab.sstaboption.TabEnabled(1) = False
    Sstab.sstaboption.TabEnabled(2) = False
    End If
ElseIf disablegamespintabs(1) = True Or disablegamespintabs(0) = False Then
    If boofirstorsec(2) = True Or boofirstorsec(3) = True Then
    Sstab.sstaboption.TabEnabled(1) = False
    Else
    Sstab.sstaboption.TabEnabled(2) = False
    End If
Else
    If boofirstorsec(0) = True Or boofirstorsec(1) = True Then
    Sstab.sstaboption.TabEnabled(2) = False
    Else
    Sstab.sstaboption.TabEnabled(1) = False
    End If
End If



Imgtinythumb(symbolselect - 1).Visible = False

tinythumb.ListImages.Add (1), , thumbs(symbolselect)
Pictinythumb.PaintPicture tinythumb.ListImages(1).Picture, 0, 0, 975, 975
imgsymbolsel.Picture = Pictinythumb.Image
tinythumb.ListImages.Clear

Set Pictinythumb = Nothing
Pictinythumb.Width = 400
Pictinythumb.Height = 400


For pct = 1 To piccount
If substitute(symbolselect, pct) = True Then
tinyshape(pct - 1).FillStyle = 0
Else
tinyshape(pct - 1).FillStyle = 1
End If
tinyshape(pct - 1).BorderStyle = 0
tinyshape(pct - 1).FillColor = &HFFFF00


tinythumb.ListImages.Add (pct), , thumbs(pct)
Pictinythumb.PaintPicture tinythumb.ListImages(pct).Picture, 0, 0, 400, 400
Imgtinythumb(pct - 1).Picture = Pictinythumb.Image
Next
For intreel = 0 To 4

With labmultiplies(intreel)
.fontname = "Times New Roman"
.Fontsize = 18
.Font.Charset = 0
.Font.Weight = 700
.FontUnderline = 0           'False
.FontItalic = 0              'False
.FontStrikethru = 0       'False
End With

With lblprize(intreel)
.fontname = "Times New Roman"
.Fontsize = 14.25
.Font.Charset = 0
.Font.Weight = 700
.FontUnderline = 0           'False
.FontItalic = 0              'False
.FontStrikethru = 0       'False
.BackColor = &H8000000E
.Caption = "0"
End With

Pictinythumb.PaintPicture tinythumb.ListImages(symbolselect).Picture, 0, 0, 400, 400
imgthumbprice(intreel).Picture = Pictinythumb.Image
Next
For pct = piccount + 1 To 14
Imgtinythumb(pct - 1).Visible = False
tinyshape(pct - 1).Visible = False
Next


    For pct = 1 To piccount
    If substitute(symbolselect, pct) = True Then
    
    For intreel = 0 To 4
    If wheelvec(intreel + 1, symbolselect) > 0 Then chkreels(intreel).Enabled = True
    Next
    End If
    Next
For intreel = 1 To 5
If reelcheck(symbolselect, intreel) = False Then chkreels(intreel - 1).Value = 0
lblprize(intreel - 1).Caption = sst(symbolselect, 5 + intreel)
Next


'If lt 5 symbols
If sst(symbolselect, 0) > 0 Then
chkgeneral(0).Enabled = False 'No scatters allowed
labmultiplies(0).Enabled = False
imgthumbprice(0).Enabled = False
spngamspn(4).Enabled = False
lblprize(0).Enabled = False
If sst(symbolselect, 0) < 13 Then
labmultiplies(1).Enabled = False
imgthumbprice(1).Enabled = False
spngamspn(5).Enabled = False
lblprize(1).Enabled = False
If sst(symbolselect, 0) < 6 Then
labmultiplies(2).Enabled = False
imgthumbprice(2).Enabled = False
spngamspn(6).Enabled = False
lblprize(2).Enabled = False
If sst(symbolselect, 0) < 2 Then
labmultiplies(3).Enabled = False
imgthumbprice(3).Enabled = False
spngamspn(7).Enabled = False
lblprize(3).Enabled = False
End If
End If
End If

'Disable free games if only a pair or less, free spin if single
If sst(symbolselect, 0) < 6 Then Sstab.sstaboption.TabEnabled(1) = False
If sst(symbolselect, 0) = 1 Then Sstab.sstaboption.TabEnabled(2) = False

symbLR = sst(symbolselect, 1)
symbRL = sst(symbolselect, 3)
symbANY = sst(symbolselect, 5)
symbMID = sst(symbolselect, 4)

optgeneral(0).Caption = "Any position pays"
Select Case sst(symbolselect, 0)
Case 2, 6
optgeneral(1).Caption = "Left to right pays"
chkgeneral(1).Visible = False
chkgeneral(2).Visible = False
If symbLR = 1 Then
optgeneral(1).Value = True
Else
optgeneral(0).Value = True
End If
Case 3, 7
optgeneral(1).Caption = "Right to left pays"
chkgeneral(1).Visible = False
chkgeneral(2).Visible = False
If symbRL = 1 Then
optgeneral(1).Value = True
Else
optgeneral(0).Value = True
End If
Case 8
optgeneral(1).Caption = "Middle threes pay"
chkgeneral(1).Visible = False
chkgeneral(2).Visible = False
If symbMID = 1 Then
optgeneral(1).Value = True
Else
optgeneral(0).Value = True
End If
Case 13
optgeneral(1).Caption = "Left to right pays"
chkgeneral(1).Caption = "Middle threes pay"
chkgeneral(2).Visible = False
If symbLR = 1 Then
optgeneral(1).Value = True
If symbMID = 1 Then chkgeneral(1).Value = 1
Else
chkgeneral(1).Enabled = False
optgeneral(0).Value = True
End If
Case 14
optgeneral(1).Caption = "Right to left pays"
chkgeneral(1).Caption = "Middle threes pay"
chkgeneral(2).Visible = False
If symbRL = 1 Then
optgeneral(1).Value = True
If symbMID = 1 Then chkgeneral(1).Value = 1
Else
chkgeneral(1).Enabled = False
optgeneral(0).Value = True
End If
Case Else
optgeneral(1).Visible = False
chkgeneral(1).Visible = False
chkgeneral(2).Visible = False
optgeneral(0).Value = True
End Select
Else
For ct = 0 To 2
If sst(symbolselect, ct + 2) = 1 Then chkgeneral(ct).Value = 1
Next

If sst(symbolselect, 1) = 1 Then
optgeneral(1).Value = True
ElseIf sst(symbolselect, 5) = 1 Then
optgeneral(0).Value = True
chkgeneral(1).Enabled = False
chkgeneral(2).Enabled = False
End If

optgeneral(0).Caption = "Any position pays"
optgeneral(1).Caption = "Left to right pays"
chkgeneral(1).Caption = "Right to left pays"
chkgeneral(2).Caption = "Middle threes pay"
End If

'Scatters
If enablescatter = True Then chkgeneral(0).Enabled = True
scatterspintemp1 = intscattervec(1, 2)
scatterspintemp2 = intscattervec(2, 2)
If scatterspintemp1 > 0 And scatterspintemp2 > 0 Then 'only 2 scatters allowed
If scatterspintemp1 <> symbolselect And scatterspintemp2 <> symbolselect Then
chkgeneral(0).Enabled = False
procend = True
Exit Sub
End If
End If
'Now disable prize buttons for scatters
If scatterspintemp1 = symbolselect Or scatterspintemp2 = symbolselect Then
scatterstart False
End If

procend = True

End Sub
Private Sub Imgtinythumb_Click(Index As Integer)
Dim indexplusone As Long
For ct = 0 To 1
chkgeneral(ct).ToolTipText = ""
optgeneral(ct).ToolTipText = ""
Next


With Imgtinythumb(Index)

If procend = False Then Exit Sub
procend = False
temp = 0
indexplusone = Index + 1




If substitute(symbolselect, indexplusone) = True Then   'no substitute
  'restore L-R option if the old substituted symbol was "any"
  If sst(indexplusone, 5) = 1 Then optgeneral(1).Enabled = True

  substitute(symbolselect, indexplusone) = False
  tinyshape(Index).FillStyle = 1
    
  'lt 5 Restore reelcheck for non substitute
  For intreel = 0 To 4
  reelcheck(indexplusone, intreel + 1) = True
  Next


  remembersubstitute = remembersubstitute - 1
    If remembersubstitute = 0 Then
    If scatterspintemp2 = 0 And enablescatter = True Then chkgeneral(0).Enabled = True 'can have scatter option again

    For intreel = 0 To 4
    chkreels(intreel).Enabled = False
    Next
    End If
Else    'get a substitute


  For pct = 1 To piccount
    If substitute(pct, symbolselect) = True Then
    .ToolTipText = "Cannot use a substituted picture as a new substituter"
    procend = True
    Exit Sub
    End If

    If substitute(pct, indexplusone) = True And pct <> symbolselect Then
    .ToolTipText = "Sorry, No picture will have more than 1 substituter"
    procend = True
    Exit Sub
    End If
  Next


  If sst(symbolselect, 2) = 1 Or sst(indexplusone, 2) = 1 Then
  .ToolTipText = "A scatter may not be used as a substituter"
  procend = True
  Exit Sub
  End If


  'The range of substituter exceeds range of substituted

  If sst(indexplusone, 5) = 1 And sst(symbolselect, 1) = 1 Then
  .ToolTipText = "A substituter paying ""left to right"" cannot substitute a picture paying ""any"""
  procend = True
  Exit Sub
  End If


  If sst(indexplusone, 4) = 1 And sst(symbolselect, 1) = 1 And sst(symbolselect, 4) = 0 Then
  .ToolTipText = "A substituter neither paying ""any"" nor ""middle threes"" cannot substitute pictures paying ""middle threes"""
  procend = True
  Exit Sub
  End If

  If sst(indexplusone, 3) = 1 And sst(symbolselect, 1) = 1 And sst(symbolselect, 3) = 0 Then
  .ToolTipText = "A substituter neither paying ""any"" nor ""right to left"" cannot substitute pictures paying ""right to left"""
  procend = True
  Exit Sub
  End If



  'The substituted prizes in correct range
  Select Case sst(symbolselect, 0)
  Case 0
  If checkprize(indexplusone, 7) = False Then
  procend = True
  Exit Sub
  End If
  Case 1
  If sst(indexplusone, 10) > sst(symbolselect, 10) Or sst(indexplusone, 9) < sst(symbolselect, 10) Then
  .ToolTipText = "Prizes of substituted picture not in correct range"
  procend = True
  Exit Sub
  End If

  Case 2 To 5
  If checkprize(indexplusone, 10) = False Then
  procend = True
  Exit Sub
  End If
  Case 6 To 12
  If checkprize(indexplusone, 9) = False Then
  procend = True
  Exit Sub
  End If
  Case 13 To 17
  If checkprize(indexplusone, 8) = False Then
  procend = True
  Exit Sub
  End If
  End Select


  For pct = 1 To piccount

  For ct = 1 To piccount
    If substitute(pct, ct) = True And symbolselect <> pct Then
    temp = temp + 1
    Exit For
    End If
    Next

    If temp = 4 Then
    .ToolTipText = "Sorry - Cannot use more than 4 substituters"
    procend = True
    Exit Sub
    End If



    If substitute(indexplusone, pct) = True Then
    .ToolTipText = "This picture is already used as a substituter"
    procend = True
    Exit Sub
    End If
    Next



    substitute(symbolselect, indexplusone) = True
    tinyshape(Index).FillStyle = 0
    remembersubstitute = remembersubstitute + 1
    chkgeneral(0).Enabled = False   'no scatters allowed
    For intreel = 0 To 4
    'If symbols lt 5
    If wheelvec(intreel + 1, symbolselect) > 0 Then chkreels(intreel).Enabled = True
  Next
  .ToolTipText = ""
End If
End With
procend = True
End Sub
Private Sub cmdOK_Click()
'Need pay mode selected

If spinstatus = 1 Then
cmdok.Caption = "Spin(2) settings?"
cmdok.ToolTipText = "Please click Spin(2) tab to verify your settings"
Exit Sub
End If

If boofirstorsec(1) = True Then
For ct = 1 To 9
freegamesettings(1, ct) = tgs(1, ct)
Next
ElseIf boofirstorsec(0) = True Then
For ct = 1 To 9
freegamesettings(0, ct) = tgs(0, ct)
Next
End If


If boofirstorsec(3) = True Then
For ct = 1 To 15
spinsettings(1, ct) = tsp(1, ct)
Next
ElseIf boofirstorsec(2) = True Then
For ct = 1 To 15
spinsettings(0, ct) = tsp(0, ct)
Next
End If
intscattervec(1, 2) = scatterspintemp1
intscattervec(2, 2) = scatterspintemp2
For pct = piccount + 1 To 14
Imgtinythumb(pct - 1).Visible = True
Next
Imgtinythumb(symbolselect - 1).Visible = True
gametype.Enabled = True
Unload Me
Set Sstab = Nothing
End Sub
Private Sub optfreegame123_Click()
If procend = False Then Exit Sub
freegameset False
updategamesetting 1, 1
procend = True
End Sub
Private Sub optfreegameany_Click()
If procend = False Then Exit Sub
freegameset True
updategamesetting 1, 2
procend = True
End Sub
Private Sub mischkgamespin_Click(Index As Integer)
Dim mischval As Long
If procend = False Then Exit Sub
procend = False
mischval = mischkgamespin(Index).Value
If boofirstorsec(1) = True Or boofirstorsec(3) = True Then
    frstsec = 1
Else
    frstsec = 0
End If
Select Case Index

Case 1
spngamspn(3).Enabled = CBool(mischval)
If mischval = 1 Then
spngamspn(3).BevelWidth = 3
lblgamspn(3).Caption = "5"
Else
lblgamspn(3).Caption = "0"
End If

Case 2, 3

If sst(symbolselect, 0) > 0 Then 'Lt 5 symbols

Select Case sst(symbolselect, 0)      '2 checkboxes involved
Case 13
tgs(frstsec, Index) = mischval
Case 14
    If mischval = 0 Then
                'This is case at most one checkbox checked
                If mischkgamespin(2).Value = 1 Or mischkgamespin(3).Value = 1 Then
                tgs(frstsec, Index) = mischval
                procend = True
                Exit Sub        '1 chkbox still enabled
                End If

        resetgamevals
    Else
                If mischkgamespin(2).Value = 1 And mischkgamespin(3).Value = 1 Then
                'This is case two checkboxes checked
                tgs(frstsec, Index) = mischval
                procend = True
                Exit Sub
                End If
 
        freegameset False
        updategamesetting Index, 1
    End If
Case Else
    If mischval = 1 Then
    freegameset False
    updategamesetting Index, 1
    Else
    resetgamevals
    End If
End Select
End If
Case 6
With spngamspn(12)
If mischval = 1 Then
.Enabled = True
.BevelWidth = 3
lblgamspn(7).Caption = "5"
Else
.Enabled = False
lblgamspn(7).Caption = "0"
End If
End With

End Select
procend = True
End Sub
Private Sub chkenablefreegame_Click()
If procend = False Then Exit Sub
procend = False
If chkenablefreegame.Value = 1 Then
    If sst(symbolselect, 0) = 0 Then
    optfreegame123.Enabled = True
    optfreegameany.Enabled = True
    Else
    'This for LT 5 symbols
    Freegamecheck
    End If

Sstab.sstaboption.TabEnabled(2) = False

Else
optfreegame123.Enabled = False
optfreegameany.Enabled = False
optfreegame123.Value = False
optfreegameany.Value = False

For ct = 2 To 3
mischkgamespin(ct).Value = 0
mischkgamespin(ct).Enabled = False
Next
If Not (hadaspin = 2 Or boofirstorsec(3) = True) Then Sstab.sstaboption.TabEnabled(2) = True
disablegamespintabs(0) = False
resetgamevals
End If

procend = True
End Sub
Private Sub sstaboption_Click(PreviousTab As Integer)
Dim tempbool As Boolean
If procend = False Then Exit Sub
procend = False
temp = 0
Sstab.HelpContextID = 8 + sstaboption.Tab
Select Case sstaboption.Tab
Case 0

'The following confirms the gamesettings
'naturals 4
'broken sequence 5
'prize bonus 6
'noprize bonus 7
'no of spins 8
'cumulative max 9
If PreviousTab = 1 Then
If optfreegame123.Value = True Or optfreegameany.Value = True Or mischkgamespin(2).Value = 1 Or mischkgamespin(3).Value = 1 Then

If boofirstorsec(1) = True Then
frstsec = 1
Else
frstsec = 0
End If


If optfreegame123.Value = True Then tgs(frstsec, 1) = 1
If optfreegameany.Value = True Then tgs(frstsec, 1) = 2

For ct = 0 To 1
tgs(frstsec, 2 + ct) = mischkgamespin(2 + ct).Value
tgs(frstsec, 4 + ct) = mischkgamespin(ct).Value
tgs(frstsec, 6 + ct) = CLng(lblgamspn(ct).Caption)
tgs(frstsec, 8 + ct) = CLng(lblgamspn(2 + ct).Caption)
Next
    'Clean unnecessary values
    If tgs(frstsec, 1) = 2 Then
    tgs(frstsec, 2) = 0
    tgs(frstsec, 3) = 0
    End If

ElseIf chkenablefreegame.Value = 1 Then
'Not selected a free game!
procend = True
chkenablefreegame.Value = 0
End If
End If

'Now confirm Tab settings from tab 3
'4 spin on nonpaying combination 9
'5 only naturals 10
'6 broken sequence 11
'4 prize bonus 12
'5 noprize bonus 13
'6 no of spins 14
'7 cumulative max 15

If spinstatus = 3 Then

If boofirstorsec(3) = True Then
frstsec = 1
Else
frstsec = 0
End If

For ct = 4 To 7
tsp(frstsec, ct + 8) = CLng(lblgamspn(ct).Caption)
Next
For ct = 4 To 6
tsp(frstsec, ct + 5) = mischkgamespin(ct).Value
Next

End If


Case 1

Select Case sst(symbolselect, 0)
Case 13
mischkgamespin(3).Visible = True
optfreegame123.Visible = True
optfreegameany.Visible = True
Case 14
mischkgamespin(2).Visible = True
mischkgamespin(3).Visible = True
Case 7, 17
mischkgamespin(2).Visible = True
Case 8
mischkgamespin(3).Visible = True
Case 6, 16
optfreegame123.Visible = True
Case 9, 10, 11, 12, 15
optfreegameany.Visible = True
Case Else
optfreegame123.Visible = True
optfreegameany.Visible = True
For ct = 0 To 1
mischkgamespin(2 + ct).Visible = True
Next
End Select

If boofirstorsec(1) = True Then
frstsec = 1
ElseIf boofirstorsec(0) = True Then
frstsec = 0
Else
procend = True
Exit Sub
End If


'Gets here only if already a game here
chkenablefreegame.Value = 1

If sst(symbolselect, 0) > 0 Then       'If lt than 5 symbols
        Freegamecheck
        Select Case sst(symbolselect, 0)
        Case 13
        If tgs(frstsec, 1) = 1 Then optfreegame123.Value = True
                If tgs(frstsec, 3) = 0 Then
                mischkgamespin(3).Value = 0
                Else
                mischkgamespin(3).Value = 1
                End If
        Case 14
                For ct = 0 To 1
                If tgs(frstsec, 2 + ct) = 0 Then
                mischkgamespin(2 + ct).Value = 0
                Else
                mischkgamespin(2 + ct).Value = 1
                End If
                Next
        Case 7, 17
                If tgs(frstsec, 2) = 0 Then
                mischkgamespin(2).Value = 0
                Else
                mischkgamespin(2).Value = 1
                End If
        Case 8
                If tgs(frstsec, 3) = 0 Then
                mischkgamespin(3).Value = 0
                Else
                mischkgamespin(3).Value = 1
                End If
        Case 6, 16
        If tgs(frstsec, 1) = 1 Then optfreegame123.Value = True
        Case 9, 10, 11, 12, 15
        If tgs(frstsec, 1) = 2 Then optfreegameany.Value = True
        End Select

Else    'If 5 symbols
        optfreegame123.Visible = True
        optfreegame123.Enabled = True
        optfreegameany.Visible = True
        optfreegameany.Enabled = True
        
        If tgs(frstsec, 1) = 1 Then
        optfreegame123.Value = True
        Else
        optfreegameany.Value = True
        End If

        For ct = 0 To 1
        mischkgamespin(2 + ct).Visible = True
        If tgs(frstsec, 1) = 1 Then
        mischkgamespin(2 + ct).Enabled = True
                If tgs(frstsec, 2 + ct) = 0 Then
                mischkgamespin(2 + ct).Value = 0
                Else
                mischkgamespin(2 + ct).Value = 1
                End If
        End If
        Next


End If


For ct = 0 To 1
lblfreegame(ct).Enabled = True
lblfreegame(ct + 2).Enabled = True
lblgamspn(ct).Enabled = True
lblgamspn(ct + 2).Enabled = True
lblgamspn(ct).Caption = tgs(frstsec, 6 + ct)
lblgamspn(ct + 2).Caption = tgs(frstsec, 8 + ct)
spngamspn(ct).Enabled = True
mischkgamespin(ct).Enabled = True
mischkgamespin(ct).Value = tgs(frstsec, 4 + ct)
Next

spngamspn(2).Enabled = True
'cumulative totals
If tgs(frstsec, 5) = 1 Then spngamspn(3).Enabled = True

lblfreegame(4).Enabled = True


'The following avoids useless non zeros in the arrays
If lblgamspn(0).Caption = "0" Then lblgamspn(0).Caption = "2"
If lblgamspn(2).Caption = "0" Then lblgamspn(2).Caption = "5"

Case 2

'If lt 5 symbols
If sst(symbolselect, 0) > 0 Then
optpairnofreespin.Caption = "Select option on pair combinations"
opttriplesnofreespin.Caption = "Select option on triple combinations"
optfoursnofreespin.Caption = "Select option on four combinations"
Select Case sst(symbolselect, 0)
Case 2
optpairsany.Visible = False
opttriplesleft.Visible = False
opttriplesany.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 3
chkspin(0).Visible = True
chkspin(0).Enabled = True
optpairleft.Visible = False
optpairsany.Visible = False
opttriplesleft.Visible = False
opttriplesany.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 4, 5
optpairleft.Visible = False
opttriplesleft.Visible = False
opttriplesany.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 6
opttriplesany.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 7
chkspin(0).Visible = True
chkspin(0).Enabled = True
chkspin(2).Visible = True
chkspin(2).Enabled = True
optpairleft.Visible = False
opttriplesleft.Visible = False
opttriplesany.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 8
chkspin(1).Visible = True
chkspin(1).Enabled = True
chkspin(3).Visible = True
chkspin(3).Enabled = True
optpairleft.Visible = False
opttriplesleft.Visible = False
opttriplesany.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 9
opttriplesleft.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 10
chkspin(0).Visible = True
chkspin(0).Enabled = True
optpairleft.Visible = False
opttriplesleft.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 11
chkspin(1).Visible = True
chkspin(1).Enabled = True
optpairleft.Visible = False
opttriplesleft.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 12
optpairleft.Visible = False
opttriplesleft.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 13
chkspin(1).Visible = True
chkspin(1).Enabled = True
chkspin(3).Visible = True
chkspin(3).Enabled = True
optfoursany.Visible = False
Case 14
chkspin(0).Visible = True
chkspin(0).Enabled = True
chkspin(1).Visible = True
chkspin(1).Enabled = True
chkspin(2).Visible = True
chkspin(2).Enabled = True
chkspin(3).Visible = True
chkspin(3).Enabled = True
chkspin(4).Visible = True
chkspin(4).Enabled = True
optpairleft.Visible = False
opttriplesleft.Visible = False
optfoursleft.Visible = False
optfoursany.Visible = False
Case 15
chkspin(0).Visible = True
chkspin(0).Enabled = True
opttriplesleft.Visible = False
optfoursleft.Visible = False
Case 16
chkspin(1).Visible = True
chkspin(1).Enabled = True
optfoursleft.Visible = False
Case 17
chkspin(0).Enabled = True
chkspin(0).Visible = True
chkspin(1).Visible = True
chkspin(1).Enabled = True
chkspin(2).Enabled = True
chkspin(2).Visible = True
optpairleft.Visible = False
opttriplesleft.Visible = False
optfoursleft.Visible = False
End Select
Else
chkspin(0).Visible = True
chkspin(1).Visible = True
chkspin(2).Visible = True
chkspin(3).Visible = True
chkspin(4).Visible = True
End If

If boofirstorsec(3) = True Then
frstsec = 1
ElseIf boofirstorsec(2) = True Then
frstsec = 0
Else
procend = True
Exit Sub
End If

If tsp(frstsec, 1) = 0 Then optpairnofreespin.Value = True
If tsp(frstsec, 1) = 1 Then optpairleft.Value = True
If tsp(frstsec, 1) = 2 Then optpairsany.Value = True
If tsp(frstsec, 4) = 0 Then opttriplesnofreespin.Value = True
If tsp(frstsec, 4) = 1 Then opttriplesleft.Value = True
If tsp(frstsec, 4) = 2 Then opttriplesany.Value = True
If tsp(frstsec, 7) = 0 Then optfoursnofreespin.Value = True
If tsp(frstsec, 7) = 1 Then optfoursleft.Value = True
If tsp(frstsec, 7) = 2 Then optfoursany.Value = True

wantaspinno = 0

For ct = 1 To 8

    If ct <> 1 And ct <> 4 And ct <> 7 Then

        If tsp(frstsec, ct) = 1 Then
        
        Select Case ct
        Case 2, 3
        chkspin(0).Enabled = True
        chkspin(1).Enabled = True
        Case 5, 6
        chkspin(2).Enabled = True
        chkspin(3).Enabled = True
        Case 8
        chkspin(4).Enabled = True
        End Select

        wantaspinno = wantaspinno + 1
        chkspin(temp).Value = 1

        End If
    temp = temp + 1
    Else
    If tsp(frstsec, ct) > 0 Then wantaspinno = wantaspinno + 1

        If tsp(frstsec, ct) = 1 Then

        Select Case ct
        Case 1
        chkspin(0).Enabled = True
        chkspin(1).Enabled = True
        Case 4
        chkspin(2).Enabled = True
        chkspin(3).Enabled = True
        Case 7
        chkspin(4).Enabled = True
        End Select
        End If

    End If
Next

Case 3

If spinstatus = 3 Then
procend = True
Exit Sub
End If

If spinstatus = 1 Then
cmdok.Caption = "Accept"
cmdok.ToolTipText = ""
End If

spinstatus = 3

If boofirstorsec(3) = True Then
frstsec = 1
Else
frstsec = 0
End If
For ct = 4 To 6
mischkgamespin(ct).Value = tsp(frstsec, 5 + ct)
Next

If tsp(frstsec, 11) = 1 Then
spngamspn(12).Enabled = True
spngamspn(12).BevelWidth = 3
End If

For ct = 4 To 7
lblgamspn(ct).Caption = tsp(frstsec, ct + 8)
spngamspn(ct).BevelWidth = 3
Next

If lblgamspn(6).Caption = "0" Then lblgamspn(6).Caption = "3"
If lblgamspn(4).Caption = "0" Then lblgamspn(4).Caption = "1"
End Select
procend = True
End Sub
Private Sub optpairnofreespin_Click()
If procend = False Then Exit Sub
intoptionchoice = 1
intspinchoice = 0
nospinplease
End Sub
Private Sub optpairleft_Click()
If procend = False Then Exit Sub
intoptionchoice = 1
intspinchoice = 1
wantaspin
End Sub
Private Sub optpairsany_Click()
If procend = False Then Exit Sub
intoptionchoice = 1
intspinchoice = 2
wantaspin
End Sub
Private Sub opttriplesnofreespin_Click()
If procend = False Then Exit Sub
intoptionchoice = 4
intspinchoice = 0
nospinplease
End Sub
Private Sub opttriplesleft_Click()
If procend = False Then Exit Sub
intoptionchoice = 4
intspinchoice = 1
wantaspin
End Sub
Private Sub opttriplesany_Click()
If procend = False Then Exit Sub
intoptionchoice = 4
intspinchoice = 2
wantaspin
End Sub
Private Sub optfoursnofreespin_Click()
If procend = False Then Exit Sub
intoptionchoice = 7
intspinchoice = 0
nospinplease
End Sub
Private Sub optfoursleft_Click()
If procend = False Then Exit Sub
intoptionchoice = 7
intspinchoice = 1
wantaspin
End Sub
Private Sub optfoursany_Click()
If procend = False Then Exit Sub
intoptionchoice = 7
intspinchoice = 2
wantaspin
End Sub
Private Sub chkspin_Click(Index As Integer)
If procend = False Then Exit Sub
procend = False
temp = chkspin(Index).Value
boojustchecked = True

intspinchoice = Index

If temp = 1 Then
wantaspin
Else
wantaspinno = wantaspinno - 1
nospinplease
End If

Select Case Index
Case 0
tsp(frstsec, 2) = temp   'Rl pairs
Case 1
tsp(frstsec, 3) = temp   'middle pairs
Case 2
tsp(frstsec, 5) = temp   'RL triples
Case 3
tsp(frstsec, 6) = temp   'Middle triples
Case 4
tsp(frstsec, 8) = temp   'RL fours
End Select
procend = True
End Sub
Private Sub wantaspin()
procend = False
If spinstatus = 0 Then spinstatus = 1

If boofirstorsec(3) = True Then
frstsec = 1
Else
frstsec = 0
End If

If boojustchecked = False Then

If tsp(frstsec, intoptionchoice) = 0 Then wantaspinno = wantaspinno + 1

If sst(symbolselect, 0) = 0 Then


If intspinchoice = 1 Then
        Select Case intoptionchoice
        Case 1
        chkspin(0).Enabled = True
        chkspin(1).Enabled = True
        Case 4
        chkspin(2).Enabled = True
        chkspin(3).Enabled = True
        Case 7
        chkspin(4).Enabled = True
        End Select
Else    'any or 0
        Select Case intoptionchoice
        Case 1
        'Must zero the checkboxes and their load values
        tsp(frstsec, 2) = 0
        tsp(frstsec, 3) = 0
        If chkspin(0).Value = 1 Then wantaspinno = wantaspinno - 1
        chkspin(0).Value = 0
        If chkspin(1).Value = 1 Then wantaspinno = wantaspinno - 1
        chkspin(1).Value = 0
        chkspin(0).Enabled = False
        chkspin(1).Enabled = False
        Case 4
        tsp(frstsec, 5) = 0
        tsp(frstsec, 6) = 0
        If chkspin(2).Value = 1 Then wantaspinno = wantaspinno - 1
        chkspin(2).Value = 0
        If chkspin(3).Value = 1 Then wantaspinno = wantaspinno - 1
        chkspin(3).Value = 0
        chkspin(2).Enabled = False
        chkspin(3).Enabled = False
        Case 7
        tsp(frstsec, 8) = 0
        If chkspin(4).Value = 1 Then wantaspinno = wantaspinno - 1
        chkspin(4).Value = 0
        chkspin(4).Enabled = False
        End Select
    End If

Else    'sstab(0) > 0

'ONLY if we have the zero option selected


'If we didn't check then initialise tsp

tsp(frstsec, intoptionchoice) = intspinchoice

Select Case sst(symbolselect, 0)

Case 8, 11, 13
        If intspinchoice = 1 Then
        Select Case intoptionchoice
        Case 1
        chkspin(1).Enabled = True
        Case 4
        chkspin(3).Enabled = True
        End Select
        Else
                Select Case intoptionchoice
                Case 1
                'Must zero the checkboxes and their load values
                tsp(frstsec, 3) = 0
                With chkspin(1)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
                Case 4
                'Must zero the checkboxes and their load values
                tsp(frstsec, 6) = 0
                With chkspin(3)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
                End Select

        End If

Case 14, 17
    If intspinchoice = 1 Then
        Select Case intoptionchoice
        Case 1
        chkspin(0).Enabled = True
        chkspin(1).Enabled = True
        Case 4
        chkspin(2).Enabled = True
        chkspin(3).Enabled = True
        End Select
    Else
        Select Case intoptionchoice
        Case 1
                'Must zero the checkboxes and their load values
                tsp(frstsec, 2) = 0
                tsp(frstsec, 3) = 0
                With chkspin(0)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
                With chkspin(1)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
        Case 4
                'Must zero the checkboxes and their load values
                tsp(frstsec, 5) = 0
                tsp(frstsec, 6) = 0
                With chkspin(2)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
                With chkspin(3)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
        End Select

    End If
Case 7, 10, 15
        If intspinchoice = 1 Then
        Select Case intoptionchoice
        Case 1
        chkspin(0).Enabled = True
        Case 4
        chkspin(2).Enabled = True
        End Select
        Else
                Select Case intoptionchoice
                Case 1
                'Must zero the checkboxes and their load values
                tsp(frstsec, 2) = 0
                With chkspin(0)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
                Case 4
                'Must zero the checkboxes and their load values
                tsp(frstsec, 5) = 0
                With chkspin(2)
                If .Value = 1 Then wantaspinno = wantaspinno - 1
                .Value = 0
                .Enabled = False
                End With
                End Select

        End If

End Select
End If

End If


If boofirstorsec(3) = True Or hadaspin = 1 Then
If boojustchecked = False Then tsp(1, intoptionchoice) = intspinchoice
gamespinsymbol(3) = symbolselect
disablegamespintabs(1) = True
boofirstorsec(3) = True
Else
If boojustchecked = False Then tsp(0, intoptionchoice) = intspinchoice
gamespinsymbol(2) = symbolselect
boofirstorsec(2) = True
End If


If boojustchecked = True Then   'need this here for lt 5 chkspin

If boofirstorsec(3) = True Then
frstsec = 1
Else
frstsec = 0
End If


Select Case intspinchoice   'here, this is chkspin(index)

Case 0, 1

If optpairsany.Value = True Then
optpairsany.Value = False
    
    If optpairleft.Visible = True Then
    tsp(frstsec, 1) = 1
    optpairleft.Value = True
    Else
    optpairnofreespin.Value = True
    tsp(frstsec, 1) = 0
    End If
Else
wantaspinno = wantaspinno + 1
End If

Case 2, 3

If opttriplesany.Value = True Then
opttriplesany.Value = False
    
    If opttriplesleft.Visible = True Then
    tsp(frstsec, 4) = 1
    opttriplesleft.Value = True
    Else
    opttriplesnofreespin.Value = True
    tsp(frstsec, 4) = 0
    End If
Else
wantaspinno = wantaspinno + 1
End If

Case 4

If optfoursany.Value = True Then
optfoursany.Value = False
    
    If optfoursleft.Visible = True Then
    tsp(frstsec, 7) = 1
    optfoursleft.Value = True
    Else
    optfoursnofreespin.Value = True
    tsp(frstsec, 7) = 0
    End If
Else
wantaspinno = wantaspinno + 1
End If
End Select

End If

If wantaspinno = 1 Then
Sstab.sstaboption.TabEnabled(3) = True
Sstab.sstaboption.TabEnabled(1) = False
End If

boojustchecked = False
procend = True
End Sub
Private Sub nospinplease()
Dim bootempcheck As Boolean
procend = False

bootempcheck = False
If boofirstorsec(3) = True Then
frstsec = 1
Else
frstsec = 0
End If


If boojustchecked = False Then

If tsp(frstsec, intoptionchoice) > 0 Then
wantaspinno = wantaspinno - 1
tsp(frstsec, intoptionchoice) = 0
End If

Select Case intoptionchoice
        Case 1
            'Must zero the checkboxes and their load values
                If chkspin(0).Value = 1 Then
                wantaspinno = wantaspinno - 1
                chkspin(0).Value = 0
                End If
            tsp(frstsec, 2) = 0
            chkspin(0).Enabled = False
                If chkspin(1).Value = 1 Then
                wantaspinno = wantaspinno - 1
                chkspin(1).Value = 0
                End If
            tsp(frstsec, 3) = 0
            chkspin(1).Enabled = False
        Case 4
            'Must zero the checkboxes and their load values
                If chkspin(2).Value = 1 Then
                wantaspinno = wantaspinno - 1
                chkspin(2).Value = 0
                End If
            tsp(frstsec, 5) = 0
            chkspin(2).Enabled = False
                If chkspin(3).Value = 1 Then
                wantaspinno = wantaspinno - 1
                chkspin(3).Value = 0
                End If
            tsp(frstsec, 6) = 0
            chkspin(3).Enabled = False
        Case 7
                If chkspin(4).Value = 1 Then
                wantaspinno = wantaspinno - 1
                chkspin(4).Value = 0
                End If
            tsp(frstsec, 8) = 0
            chkspin(4).Enabled = False
End Select


End If

boojustchecked = False  'Must be reset

If wantaspinno > 0 Then
procend = True
Exit Sub
End If


spinstatus = 0
'Zero out tab3
For ct = 4 To 6
mischkgamespin(ct).Value = 0
tsp(frstsec, 5 + ct) = 0
tsp(frstsec, 8 + ct) = 0
lblgamspn(ct - 1).Caption = ""
Next
tsp(frstsec, 15) = 0
Sstab.sstaboption.TabEnabled(3) = False
        If disablegamespintabs(0) = False Then
        If sst(symbolselect, 0) = 0 Or sst(symbolselect, 0) > 5 Then Sstab.sstaboption.TabEnabled(1) = True
        End If
disablegamespintabs(1) = False
If frstsec = 1 Then
gamespinsymbol(3) = 0
boofirstorsec(3) = False
hadaspin = 1
Else
gamespinsymbol(2) = 0
boofirstorsec(2) = False
hadaspin = 0
End If
procend = True
End Sub
Private Sub chkgeneral_Click(Index As Integer)
If procend = False Then Exit Sub
procend = False

With chkgeneral(Index)

nomorespinup = False
optgeneral(0).ToolTipText = ""
'note special arrangement in sstabgeneral

If .Value = 1 Then
  If Index = 0 Then

  getscatter
  sst(symbolselect, 2) = 1

  ElseIf Index = 1 And sst(symbolselect, 2) = 1 Then

    If sst(symbolselect, 9) > 0 Then
    .ToolTipText = "Please ensure pair - prize value is zero"
    .Value = 0
    procend = True
    Exit Sub
    End If
    
  sst(symbolselect, 3) = 1

  scatterstart True

  Else
        
    If sst(symbolselect, 0) = 0 Then
      If Index = 1 Then
      'The range of substituter exceeds range of substituted
      For pct = 1 To piccount
        Select Case sst(pct, 0)
        Case 0, 3, 7, 10, 14, 15, 17
        If sst(pct, 3) = 0 And sst(pct, 5) = 0 Then
        If substitute(pct, symbolselect) = True Then
        .ToolTipText = "A substituter neither paying ""any"" nor ""right to left"" cannot substitute pictures paying ""right to left"""
        .Value = 0
        procend = True
        Exit Sub
        ElseIf substitute(symbolselect, pct) = True Then
        .ToolTipText = "A substituter paying ""right to left"" cannot substitute pictures not paying ""right to left"""
        .Value = 0
        procend = True
        Exit Sub
        End If
        End If
        End Select
      Next
      End If
        
      If Index = 2 Then
      For pct = 1 To piccount
        Select Case sst(pct, 0)
        Case 0, 8, 13, 14
        If sst(pct, 4) = 0 And sst(pct, 5) = 0 Then
        If substitute(pct, symbolselect) = True Then
        .ToolTipText = "A substituter neither paying ""any"" nor ""middle threes"" cannot substitute pictures paying ""middle threes"""
        .Value = 0
        procend = True
        Exit Sub
        ElseIf substitute(symbolselect, pct) = True Then
        .ToolTipText = "A substituter paying ""middle threes"" cannot substitute pictures not paying ""middle threes"""
        .Value = 0
        procend = True
        Exit Sub
        End If
        End If
        End Select
      Next
      End If
      sst(symbolselect, Index + 2) = 1

    Else 'lt 5
      'Event triggered *only* for sst(symbolselect, 0) 13 or 14 - not 8 as chk is greyed
      For pct = 1 To piccount
        If sst(pct, 4) = 0 And sst(pct, 5) = 0 Then
        If substitute(pct, symbolselect) = True Then
        .ToolTipText = "A substituter neither paying ""any"" nor ""middle threes"" cannot substitute pictures paying ""middle threes"""
        .Value = 0
        procend = True
        Exit Sub
        End If
        If substitute(symbolselect, pct) = True Then
        .ToolTipText = "A substituter paying ""middle threes"" cannot substitute pictures not paying ""middle threes"""
        .Value = 0
        procend = True
        Exit Sub
        End If
        End If
        Next
      sst(symbolselect, 4) = 1
      End If
    End If


Else 'chkgeneral.value = 0

    If Index = 0 Then   'deselect scatters
    For pct = 0 To piccount - 1
    Imgtinythumb(pct).ToolTipText = ""
    Next

    'If not lt 5 , re-enable middle threes
    If sst(symbolselect, 0) = 0 And sst(symbolselect, 1) = 1 Then chkgeneral(2).Enabled = True

    For ct = 6 To 10 'Reenable prize spins and set default prizes
    spngamspn(ct - 2).Enabled = True
    
    resetprize ct, symbolselect
    
    lblprize(ct - 6).Caption = sst(symbolselect, ct)
    Next

      If scatterspintemp2 = symbolselect Then
      scatterspintemp2 = 0    'note resetting other scatter vars not needed here
      intscatternumber = 1
      ElseIf scatterspintemp2 > 0 Then 'scatterspintemp1 is selected
      scatterspintemp1 = scatterspintemp2
      scatterspintemp2 = 0
      intscatternumber = 1
      Else
      scatterspintemp1 = 0
      intscatternumber = 0
      zeroscatter
      End If

      sst(symbolselect, 2) = 0
    ElseIf Index = 1 And sst(symbolselect, 2) = 1 Then

    sst(symbolselect, 3) = 0
    spngamspn(7).Enabled = True
    scatterstart True

    Else    'no scatters

    If sst(symbolselect, 0) = 0 Then
      If Index = 1 Then
      'The range of substituter exceeds range of substituted
        For pct = 1 To piccount
        Select Case sst(pct, 0)
        Case 0, 3, 7, 10, 14, 15, 17
          If sst(pct, 3) = 1 And substitute(symbolselect, pct) = True Then
          .ToolTipText = "A substituter neither paying ""any"" nor ""right to left"" cannot substitute pictures paying ""right to left"""
          .Value = 1
          procend = True
          Exit Sub
          End If
          If sst(pct, 3) = 1 And substitute(pct, symbolselect) = True Then
          .ToolTipText = "A substituter paying ""right to left"" cannot substitute pictures not paying ""right to left"""
          .Value = 1
          procend = True
          Exit Sub
          End If
        End Select
        Next
      End If
        
      If Index = 2 Then
        For pct = 1 To piccount
        Select Case sst(pct, 0)
        Case 0, 8, 13, 14
          If sst(pct, 4) = 1 And substitute(symbolselect, pct) = True Then
          .ToolTipText = "A substituter neither paying ""any"" nor ""middle threes"" cannot substitute pictures paying ""middle threes"""
          .Value = 1
          procend = True
          Exit Sub
          End If
          If sst(pct, 4) = 1 And substitute(pct, symbolselect) = True Then
          .ToolTipText = "A substituter paying ""middle threes"" cannot substitute pictures not paying ""middle threes"""
          .Value = 1
          procend = True
          Exit Sub
          End If
        End Select
        Next
      End If
      sst(symbolselect, Index + 2) = 0


      Else 'lt 5

      For pct = 1 To piccount
        If sst(pct, 4) = 1 Then
        If substitute(pct, symbolselect) = True Then
        .ToolTipText = "A substituter paying ""middle threes"" cannot substitute pictures not paying ""middle threes"""
        .Value = 0
        procend = True
        Exit Sub
        End If
        If substitute(symbolselect, pct) = True Then
        .ToolTipText = "A substituter neither paying ""any"" nor ""middle threes"" cannot substitute pictures paying ""middle threes"""
        .Value = 0
        procend = True
        Exit Sub
        End If
        End If
      Next
      sst(symbolselect, 4) = 0
      End If

    End If




End If
End With
procend = True
End Sub
Private Sub optgeneral_Click(Index As Integer)

'Re-initialise scatterspin variable

If procend = False Then Exit Sub
procend = False

nomorespinup = False

For ct = 0 To 1
chkgeneral(ct).ToolTipText = ""
optgeneral(ct).ToolTipText = ""
Next


If Index = 1 Then   'L-R

For pct = 1 To piccount     'Cannot have a LR substituting an "any"
If sst(pct, 5) = 1 Then
  If substitute(symbolselect, pct) = True Then
    optgeneral(1).ToolTipText = "A substituter paying ""left to right"" cannot substitute a picture paying ""any"""
    optgeneral(0).Value = True
    procend = True
    Exit Sub
  End If

ElseIf substitute(pct, symbolselect) = True Then
  'sst(pct, 1) = 1
  If sst(pct, 3) = 0 then
    Select Case sst(pct, 0)
    Case 3, 7, 10, 14, 15, 17
    .ToolTipText = "The substituter of this picture must pay ""right to left"""
    .Value = 0
    procend = True
    Exit Sub
    End Select
  ElseIf sst(pct, 4) = 0 then

    Select Case sst(pct, 0)
    Case 8, 13, 14
    If sst(pct, 4) = 0 And sst(pct, 5) = 0 Then
    If substitute(pct, symbolselect) = True Then
    .ToolTipText = "The substituter of this picture must pay ""middle threes"""
    .Value = 0
    procend = True
    Exit Sub
    End Select
  End If
End If
Next






sst(symbolselect, 1) = 1
sst(symbolselect, 5) = 0

Select Case sst(symbolselect, 0)       'lt 5 case
Case 0
chkgeneral(1).Enabled = True
        If sst(symbolselect, 2) = 0 Then
        chkgeneral(2).Enabled = True 'if not a scatter
        Else
        'Need to change the scatter prizes as well
        scatterstart True
        End If
Case 3, 7
sst(symbolselect, 3) = 1
Case 8
sst(symbolselect, 4) = 1
Case 13
chkgeneral(1).Enabled = True
Case 14
chkgeneral(1).Enabled = True
sst(symbolselect, 3) = 1
End Select



Else 'Index = 0


For pct = 1 To piccount     'Cannot have a LR substituting an "any"
  If sst(pct, 1) = 1 Then
    If substitute(pct, symbolselect) = True Then
    optgeneral(0).ToolTipText = "A substituter paying one or more of ""left to right"", ""right to left"", ""middle threes"" cannot substitute a picture paying ""any"""
    optgeneral(1).Value = True
    procend = True
    Exit Sub
    End If
  End If
Next


spngamspn(7).Enabled = True 'in case it was disabled in LR routine

sst(symbolselect, 1) = 0
sst(symbolselect, 5) = 1

Select Case sst(symbolselect, 0)       'lt 5 case
Case 0
For ct = 1 To 2
chkgeneral(ct).Enabled = False
chkgeneral(ct).Value = 0
sst(symbolselect, ct + 2) = 0
Next
'Need to change the scatter prizes as well
If sst(symbolselect, 2) = 1 Then scatterstart True
Case 3, 7
sst(symbolselect, 3) = 0
Case 8
sst(symbolselect, 4) = 0
Case 13
sst(symbolselect, 4) = 0
chkgeneral(1).Enabled = False
chkgeneral(1).Value = 0
Case 14
chkgeneral(1).Enabled = False
chkgeneral(1).Value = 0
sst(symbolselect, 3) = 0
sst(symbolselect, 4) = 0
End Select


End If
procend = True
End Sub
Private Sub spngamspn_SpinUp(Index As Integer)
Dim testval As Long

With lblgamspn(Index)
If Index < 4 Then
temp = CLng(.Caption)
ElseIf Index > 8 Then
temp = CLng(lblgamspn(Index - 5).Caption)
Else
    temp = CLng(lblprize(Index - 4).Caption)

    'No low level prize exceeds a higher level one
    If chkprizelevel(True, Index + 2) = False Then Exit Sub
    'Substituted price < substitute price but substituted price > substitute price - 1
        For pct = 1 To piccount
        If pct <> symbolselect Then
        testval = -1
        If sortprizes(Index + 2, pct, testval) = True And substitute(pct, symbolselect) = True And temp >= sst(pct, Index + 2) Then Exit Sub
        If Index > 4 Then
        If sortprizes(Index + 1, pct, testval) = True And substitute(symbolselect, pct) = True And temp >= sst(pct, Index + 1) Then Exit Sub
        End If
        End If
        Next

End If

Select Case Index
Case 0
If temp < 3 Then .Caption = CStr(temp + 1)
Case 1
If temp < 2 Then .Caption = CStr(temp + 1)
Case 2
If temp < 10 Then .Caption = CStr(temp + 1)
Case 3
If temp < 10 Then .Caption = CStr(temp + 1)
Case 4
If temp < 5000 Then
If temp >= 100 Then
lblprize(0).Caption = CStr(temp + 50)
ElseIf temp >= 50 Then
lblprize(0).Caption = CStr(temp + 10)
Else
lblprize(0).Caption = CStr(temp + 5)
End If
sst(symbolselect, 6) = lblprize(0).Caption
End If
Case 5
If temp < 1000 Then
If temp >= 100 Then
lblprize(1).Caption = CStr(temp + 10)
ElseIf temp >= 10 Then
lblprize(1).Caption = CStr(temp + 5)
Else
lblprize(1).Caption = CStr(temp + 1)
End If
sst(symbolselect, 7) = lblprize(1).Caption
End If
Case 6
If sst(symbolselect, 2) = 1 Then
scatterspin True, Index + 2, lblprize(2).Caption
Else
If temp < 200 Then
    If temp >= 20 Then
    lblprize(2).Caption = CStr(temp + 5)
    Else
    lblprize(2).Caption = CStr(temp + 1)
    End If
sst(symbolselect, 8) = lblprize(2).Caption
End If
End If
Case 7
If sst(symbolselect, 2) = 1 Then
scatterspin True, Index, temp
Else
If temp < 10 Then
lblprize(3).Caption = CStr(temp + 1)
sst(symbolselect, 9) = lblprize(3).Caption
End If
End If
Case 8
If temp < 5 Then
lblprize(4).Caption = CStr(temp + 1)
sst(symbolselect, 10) = lblprize(4).Caption
End If
Case 9
If temp < 3 Then
lblgamspn(4).Caption = CStr(temp + 1)
End If
Case 10
If temp < 2 Then
lblgamspn(5).Caption = CStr(temp + 1)
End If
Case 11
If temp < 5 Then
lblgamspn(6).Caption = CStr(temp + 1)
End If
Case 12
If temp < 10 Then
lblgamspn(7).Caption = CStr(temp + 1)
End If

End Select
End With
End Sub
Private Sub spngamspn_SpinDown(Index As Integer)
Dim testval As Long
With lblgamspn(Index)

If Index < 4 Then    ' free games or free spins
temp = CLng(.Caption)
ElseIf Index > 8 Then
temp = CLng(lblgamspn(Index - 5).Caption)
Else
    temp = CLng(lblprize(Index - 4).Caption)

    chkgeneral(1).ToolTipText = ""  'no RL scatter warning
    'No low level prize exceeds a higher level one

    If chkprizelevel(False, Index + 2) = False Then Exit Sub

    'Substituted price < substitute price but substituted price > substitute price - 1
        For pct = 1 To piccount
        If pct <> symbolselect Then
        testval = -1
        If sortprizes(Index + 2, pct, testval) = True And substitute(symbolselect, pct) = True And temp <= sst(pct, Index + 2) Then Exit Sub
            If Index < 8 Then
            If sortprizes(Index + 3, pct, testval) = True And substitute(pct, symbolselect) = True And temp <= sst(pct, Index + 3) Then Exit Sub
            End If
        End If
        Next
End If


Select Case Index
Case 0
If temp > 1 Then .Caption = CStr(temp - 1)
Case 1
If temp > 0 Then .Caption = CStr(temp - 1)
Case 2
If temp > 1 Then .Caption = CStr(temp - 1)
Case 3
If temp > 0 Then .Caption = CStr(temp - 1)
Case 4
If temp > 100 Then
lblprize(0).Caption = CStr(temp - 50)
ElseIf temp > 50 Then
lblprize(0).Caption = CStr(temp - 10)
ElseIf temp > 25 Then
lblprize(0).Caption = CStr(temp - 5)
End If
sst(symbolselect, 6) = lblprize(0).Caption
Case 5
If temp > 100 Then
lblprize(1).Caption = CStr(temp - 10)
ElseIf temp > 10 Then
lblprize(1).Caption = CStr(temp - 5)
ElseIf temp > 5 Then
lblprize(1).Caption = CStr(temp - 1)
End If
sst(symbolselect, 7) = lblprize(1).Caption
Case 6
If sst(symbolselect, 2) = 1 Then
scatterspin False, Index + 2, lblprize(2).Caption
Else
If temp > 1 Then
    If temp > 20 Then
    lblprize(2).Caption = CStr(temp - 5)
    Else
    lblprize(2).Caption = CStr(temp - 1)
    End If
sst(symbolselect, 8) = lblprize(2).Caption
End If
End If
Case 7
If sst(symbolselect, 2) = 1 Then
scatterspin False, Index, lblprize(3).Caption
Else
If temp > sst(symbolselect, 10) Then
lblprize(3).Caption = CStr(temp - 1)
sst(symbolselect, 9) = lblprize(3).Caption
End If
End If
Case 8
If temp > 0 Then
lblprize(4).Caption = CStr(temp - 1)
sst(symbolselect, 10) = lblprize(4).Caption
End If
Case 9
If temp > 1 Then lblgamspn(4).Caption = CStr(temp - 1)
Case 10
If temp > 0 Then lblgamspn(5).Caption = CStr(temp - 1)
Case 11
If temp > 1 Then lblgamspn(6).Caption = CStr(temp - 1)
Case 12
If temp > 0 Then lblgamspn(7).Caption = CStr(temp - 1)
End Select
End With
End Sub
Private Sub chkreels_Click(Index As Integer)
If procend = False Then Exit Sub
procend = False
If chkreels(Index).Value = 1 Then
    'Remember reelcheck default value is true
    reelcheck(symbolselect, Index + 1) = True
    'Have to change reelcheck for substituted symbols as well
    For pct = 1 To piccount
    If pct <> symbolselect And substitute(symbolselect, pct) = True Then reelcheck(pct, Index + 1) = True
    Next
Else
        temp = 0
        For intreel = 1 To 5
        If reelcheck(symbolselect, intreel) = True And wheelvec(intreel, symbolselect) > 0 Then temp = temp + 1
        Next
        If temp = 1 Then
        chkreels(Index).Value = 1
        procend = True
        Exit Sub
        Else
        reelcheck(symbolselect, Index + 1) = False
            'Have to change reelcheck for substituted symbols as well
            For pct = 1 To piccount
            If pct <> symbolselect And substitute(symbolselect, pct) = True Then reelcheck(pct, Index + 1) = False
            Next
        End If
End If
procend = True
End Sub
Private Sub Freegamecheck()
procend = False
If boofirstorsec(1) = True Then
frstsec = 1
Else
frstsec = 0
End If
'tgs(1) = 123 / any, tgs(2)= RL, tgs(3) = middle

Select Case sst(symbolselect, 0)
Case 13
optfreegame123.Visible = True
optfreegameany.Visible = True
optfreegame123.Enabled = True
optfreegameany.Enabled = True
If tgs(frstsec, 1) = 1 Then
optfreegame123.Value = True
mischkgamespin(3).Enabled = True
Else
optfreegameany.Value = True
End If
        If tgs(frstsec, 3) = 0 Then
        mischkgamespin(3).Value = 0
        Else
        mischkgamespin(3).Value = 1
        End If
Case 14
mischkgamespin(2).Visible = True
mischkgamespin(3).Visible = True
mischkgamespin(2).Enabled = True
mischkgamespin(3).Enabled = True
        For ct = 0 To 1
        If tgs(frstsec, 2 + ct) = 0 Then
        mischkgamespin(2 + ct).Value = 0
        Else
        mischkgamespin(2 + ct).Value = 1
        End If
        Next
Case 7, 17
With mischkgamespin(2)
.Visible = True
.Enabled = True
        If tgs(frstsec, 2) = 0 Then
        .Value = 0
        Else
        .Value = 1
        End If
End With
Case 8
With mischkgamespin(3)
.Visible = True
.Enabled = True
        If tgs(frstsec, 3) = 0 Then
        .Value = 0
        Else
        .Value = 1
        End If
End With
Case 6, 16
optfreegame123.Visible = True
optfreegame123.Enabled = True
If tgs(frstsec, 1) = 1 Then optfreegame123.Value = True
Case 9, 10, 11, 12, 15
optfreegameany.Visible = True
optfreegameany.Enabled = True
If tgs(frstsec, 1) = 2 Then optfreegameany.Value = True
End Select
procend = True
End Sub
Private Sub updategamesetting(intindex As Integer, tempval As Long)
If hadagame = 1 Or boofirstorsec(1) = True Then
gamespinsymbol(1) = symbolselect
disablegamespintabs(0) = True
boofirstorsec(1) = True
tgs(1, intindex) = tempval
Else
gamespinsymbol(0) = symbolselect
boofirstorsec(0) = True
tgs(0, intindex) = tempval
End If
End Sub
Private Sub resetgamevals()
procend = False
For ct = 0 To 1
mischkgamespin(ct).Value = 0
mischkgamespin(ct).Enabled = False
lblgamspn(ct).Enabled = False
lblgamspn(ct + 2).Enabled = False
lblfreegame(ct + 2).Enabled = False
lblfreegame(ct).Enabled = False
spngamspn(ct).Enabled = False
spngamspn(ct + 2).Enabled = False
Next

lblfreegame(4).Enabled = False

If boofirstorsec(0) = True Then
boofirstorsec(0) = False
gamespinsymbol(0) = 0
For ct = 1 To 9
tgs(0, ct) = 0
Next
hadagame = 0
ElseIf boofirstorsec(1) = True Then
For ct = 1 To 9
tgs(1, ct) = 0
Next
gamespinsymbol(1) = 0
boofirstorsec(1) = False
hadagame = 1
Else
gamespinsymbol(0) = 0
boofirstorsec(0) = False
hadagame = 0
End If
procend = True
End Sub
Private Sub freegameset(optany As Boolean)
procend = False

If lblgamspn(0).Enabled = False Then    'the first time here

For ct = 0 To 3
lblgamspn(ct).Enabled = True
If ct <> 3 Then spngamspn(ct).Enabled = True
spngamspn(ct).BevelWidth = 3
If ct <> 2 Then mischkgamespin(ct).Enabled = True
Next
For intreel = 0 To 4
lblfreegame(intreel).Enabled = True
Next

lblgamspn(0).Caption = "1"
lblgamspn(1).Caption = "0"
lblgamspn(2).Caption = "5"
lblgamspn(3).Caption = "0"

lblfreegame(4).Enabled = True

End If

Select Case sst(symbolselect, 0)
Case 0, 13
For ct = 2 To 3
mischkgamespin(ct).Enabled = Not (optany)
Next
Case 14
If optany = True Then
For ct = 2 To 3
mischkgamespin(ct).Value = 0
Next
End If
Case 7, 17
If optany = False Then
mischkgamespin(2).Value = 1
Else
mischkgamespin(2).Value = 0
End If
Case 8
If optany = False Then
mischkgamespin(3).Value = 1
Else
mischkgamespin(3).Value = 0
End If
End Select


procend = True
End Sub
Private Sub scatterspin(upup As Boolean, baseprizeno As Integer, currentval As Long)
Dim sstabany As Boolean, pnum As Long, currval As Long

currval = currentval

If sst(symbolselect, 5) = 1 Then
sstabany = True
Else
sstabany = False
End If

If baseprizeno = 8 And lblprize(3).Caption > 0 Then Exit Sub 'exit if the spin below is > 0

If upup = True Then
    If currval < 9 And nomorespinup = False Then  'going up
    If sst(symbolselect, 3) = 0 Then spngamspn(7).Enabled = True
    If currval = 8 Then Exit Sub
    
        If currval > 0 Then
    
        currval = currval + 1
            For ct = baseprizeno To 10
    
            If ct = baseprizeno Then
            sst(symbolselect, 16 - ct) = currval
            
            Else
            temp = currval * lblprize(10 - ct).Caption / (currval - 1)
            If temp > 10975 Then nomorespinup = True
            lblprize(10 - ct).Caption = temp
            sst(symbolselect, 16 - ct) = temp
            End If
            Next
            
        Else    'currval <= 0, starts here *only* if baseprizeno = 7 as bspn 8 <> 0
        'Exit if not correct prize multiple
            If sstabany = False Then
            If lblprize(2).Caption <> 24 / (3 * wheelvec(1, symbolselect)) Then Exit Sub
            Else
            If lblprize(2).Caption <> (24 - 3 * wheelvec(1, symbolselect)) / (3 * wheelvec(1, symbolselect)) Then Exit Sub
            End If
        currval = 1
        sst(symbolselect, 9) = 1
        End If
        
        scatterstart False

    End If
Else    'upup = false
    nomorespinup = False
    If currval > 0 Then   '=0 then disable spngamspn(7) but makes the control go crazy
    
    If currval = 1 Then

    If baseprizeno = 8 Then Exit Sub
    sst(symbolselect, 9) = 0
    End If
        
    currval = currval - 1
    For ct = baseprizeno + 1 To 10
    
    If currval > 0 Then 'Otherwise zeros everything
    
    
    temp = currval * lblprize(10 - ct).Caption / (currval + 1)
    lblprize(10 - ct).Caption = temp
    sst(symbolselect, 16 - ct) = temp
    End If
    
    Next

    End If

End If
lblprize(10 - baseprizeno).Caption = currval
sst(symbolselect, 16 - baseprizeno) = currval

End Sub
Private Sub scatterstart(calcprize As Boolean)
Dim baseno As Long

sst(symbolselect, 10) = 0      'no prize for single scatter
'disable middle threes
If sst(symbolselect, 1) = 1 Then chkgeneral(2).Enabled = False


'Now disable prize buttons for scatters
For ct = 10 To 6 Step -1 'First find baseprizeno

Select Case ct
Case 8
spngamspn(6).Enabled = True
Case 9
    If sst(symbolselect, 3) = 1 Then
    spngamspn(7).Enabled = False
    baseno = 3
    sst(symbolselect, 9) = 0
    Else
        spngamspn(7).Enabled = True
        If sst(symbolselect, 9) > 0 Then
        baseno = 2
        Else
        baseno = 3
        End If
    End If
Case Else
spngamspn(ct - 2).Enabled = False
End Select

Next

If calcprize = True Then

CalcScatterprize symbolselect, wheelvec(1, symbolselect), baseno
End If

For intreel = 0 To 4    'refresh prizes
lblprize(intreel).Caption = sst(symbolselect, intreel + 6)
Next

End Sub
Private Sub getscatter()
    If scatterspintemp1 > 0 Then
    scatterspintemp2 = symbolselect
    intscatternumber = 2
    scatterstart True
    Else    'scatterspintemp1 = 0
    scatterspintemp1 = symbolselect
    intscatternumber = 1
    scatterstart True
    End If
sst(symbolselect, 4) = 0 'middle threes always 0
End Sub
Private Function checkprize(picselect As Long, maxitn As Long)
Dim testpz As Long
checkprize = True

Select Case sst(picselect, 0)
Case 0
testpz = maxY(maxitn, 7)
Case 2 To 5
testpz = 10
Case 6 To 12
testpz = maxY(maxitn, 9)
Case 13 To 17
testpz = maxY(maxitn, 8)
End Select

For ct = testpz To 10
If sst(picselect, ct - 1) < sst(symbolselect, ct) Then checkprize = False
Next


For ct = testpz To 10
If sst(picselect, ct) > sst(symbolselect, ct) Then checkprize = False
Next


If checkprize = False Then Imgtinythumb(picselect - 1).ToolTipText = "Prizes of substituted picture not in correct range"

End Function
Private Function enablescatter()
'Disable scatter where apt
remembersubstitute = 0


For pct = 1 To piccount
If substitute(pct, symbolselect) = True Then
enablescatter = False  'No scatters with substitutes
Exit Function
ElseIf substitute(symbolselect, pct) = True Then
remembersubstitute = remembersubstitute + 1
End If
Next


If remembersubstitute > 0 Then
enablescatter = False  'No scatters with substitutes
Exit Function
End If

'Must be configured in cornfig
If reelcheck(symbolselect, 0) = False Then
enablescatter = False
Exit Function
End If


'Test wheelvec

If testscatter(wheelvec(1, symbolselect), wheelvec(2, symbolselect), wheelvec(3, symbolselect), wheelvec(4, symbolselect), wheelvec(5, symbolselect)) = False Then
enablescatter = False
Exit Function
End If

enablescatter = True
End Function
