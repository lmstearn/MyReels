VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{C3CBD80D-C8D1-11D2-9F8E-0080C7CE5CDC}#4.1#0"; "ActCndy2.ocx"
Begin VB.Form Genopts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyReels: General Options"
   ClientHeight    =   6945
   ClientLeft      =   2745
   ClientTop       =   1545
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
   HelpContextID   =   12
   Icon            =   "genopts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8070
   Visible         =   0   'False
   Begin VB.Timer MidiPlay 
      Enabled         =   0   'False
      Interval        =   18
      Left            =   6240
      Top             =   5160
   End
   Begin VB.CommandButton cmdgenopts 
      Cancel          =   -1  'True
      Caption         =   "&Accept"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   10
      Top             =   6400
      Width           =   975
   End
   Begin TabDlg.SSTab sstaboption 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "genopts.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sound"
      TabPicture(1)   =   "genopts.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Text"
      TabPicture(2)   =   "genopts.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Appearance"
      TabPicture(3)   =   "genopts.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   " "
         Height          =   6015
         Index           =   2
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   7815
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   975
            Index           =   14
            Left            =   240
            TabIndex        =   110
            Top             =   4800
            Width           =   7575
            Begin VB.CheckBox chkgenopt 
               Caption         =   "Use Base Dir."
               Height          =   255
               Index           =   10
               Left            =   3840
               TabIndex        =   153
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Configure &Thumbnails ..."
               Height          =   375
               Index           =   5
               Left            =   5400
               TabIndex        =   152
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox chkgenopt 
               Caption         =   "Randomise"
               Height          =   255
               Index           =   7
               Left            =   3840
               TabIndex        =   138
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "&Select Text file ..."
               Height          =   375
               Index           =   35
               Left            =   1560
               TabIndex        =   112
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox chkgenopt 
               Caption         =   "Show Quotes"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   111
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   735
            Index           =   13
            Left            =   240
            TabIndex        =   106
            Top             =   4000
            Width           =   7575
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   4080
               TabIndex        =   117
               Top             =   240
               Width           =   3375
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   9
               ItemData        =   "genopts.frx":007C
               Left            =   840
               List            =   "genopts.frx":009B
               TabIndex        =   107
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Caption         =   "Prizes"
               Height          =   255
               Index           =   48
               Left            =   240
               TabIndex        =   109
               Top             =   240
               Width           =   495
            End
            Begin VB.Label labgenopts 
               Caption         =   "Winning caption is :"
               Height          =   255
               Index           =   47
               Left            =   2520
               TabIndex        =   108
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   2415
            Index           =   12
            Left            =   240
            TabIndex        =   102
            Top             =   1600
            Width           =   5895
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   2280
               TabIndex        =   118
               Top             =   1800
               Width           =   3375
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   2280
               TabIndex        =   115
               Top             =   1320
               Width           =   3375
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   2280
               TabIndex        =   114
               Top             =   840
               Width           =   3375
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   2280
               TabIndex        =   113
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Welcome Message"
               Height          =   255
               Index           =   49
               Left            =   240
               TabIndex        =   116
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label labgenopts 
               Caption         =   "Random Jackpot Prize"
               Height          =   255
               Index           =   46
               Left            =   240
               TabIndex        =   105
               Top             =   1800
               Width           =   1695
            End
            Begin VB.Label labgenopts 
               Caption         =   "Money Back Prize"
               Height          =   255
               Index           =   45
               Left            =   240
               TabIndex        =   104
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label labgenopts 
               Caption         =   "Spin Prompt"
               Height          =   255
               Index           =   40
               Left            =   240
               TabIndex        =   103
               Top             =   840
               Width           =   975
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   735
            Index           =   11
            Left            =   240
            TabIndex        =   97
            Top             =   800
            Width           =   7575
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   1560
               TabIndex        =   100
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   5520
               TabIndex        =   98
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label labgenopts 
               Caption         =   "Spin button title"
               Height          =   255
               Index           =   42
               Left            =   120
               TabIndex        =   101
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label labgenopts 
               Caption         =   "Bet choice button title"
               Height          =   255
               Index           =   43
               Left            =   3720
               TabIndex        =   99
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   735
            Index           =   10
            Left            =   240
            TabIndex        =   92
            Top             =   0
            Width           =   7575
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   1080
               TabIndex        =   95
               Top             =   240
               Width           =   3375
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   5520
               TabIndex        =   93
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label labgenopts 
               Caption         =   "Player Name"
               Height          =   255
               Index           =   44
               Left            =   120
               TabIndex        =   96
               Top             =   240
               Width           =   975
            End
            Begin VB.Label labgenopts 
               Caption         =   "Game title "
               Height          =   255
               Index           =   41
               Left            =   4560
               TabIndex        =   94
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   " "
         Height          =   6015
         Index           =   1
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   7815
         Begin VB.CheckBox chkgenopt 
            Caption         =   "Enable Thumbnail Sound"
            Height          =   255
            Index           =   9
            Left            =   5040
            TabIndex        =   159
            Top             =   60
            Width           =   2055
         End
         Begin VB.CheckBox chkgenopt 
            Caption         =   "Enable Sound"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   139
            Top             =   60
            Width           =   1335
         End
         Begin VB.Frame fragenopts 
            Caption         =   "Wave Files"
            Height          =   3075
            Index           =   16
            Left            =   0
            TabIndex        =   123
            Top             =   360
            Width           =   7815
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Browse ..."
               Height          =   375
               Index           =   31
               Left            =   6840
               TabIndex        =   140
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkgenopt 
               Caption         =   "List only currently assigned System sounds"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   135
               Top             =   2760
               Width           =   3375
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Clear"
               Height          =   375
               Index           =   30
               Left            =   6840
               TabIndex        =   134
               ToolTipText     =   "Clears the selected entry"
               Top             =   1200
               Width           =   855
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Add"
               Height          =   375
               Index           =   29
               Left            =   6840
               TabIndex        =   133
               Top             =   720
               Width           =   855
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Test"
               Height          =   375
               Index           =   28
               Left            =   6840
               TabIndex        =   132
               Top             =   240
               Width           =   855
            End
            Begin VB.ListBox Lstgenopts 
               Height          =   1530
               Index           =   1
               ItemData        =   "genopts.frx":00F3
               Left            =   1800
               List            =   "genopts.frx":00F5
               TabIndex        =   131
               Top             =   600
               Width           =   4935
            End
            Begin VB.ListBox Lstgenopts 
               Height          =   2370
               Index           =   0
               ItemData        =   "genopts.frx":00F7
               Left            =   120
               List            =   "genopts.frx":0122
               TabIndex        =   130
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   1800
               TabIndex        =   125
               Top             =   240
               Width           =   4935
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   13
               ItemData        =   "genopts.frx":01EB
               Left            =   1800
               List            =   "genopts.frx":01ED
               TabIndex        =   124
               ToolTipText     =   "Choose any system sound"
               Top             =   2280
               Width           =   4935
            End
            Begin VB.Label labgenopts 
               Caption         =   "System"
               Height          =   255
               Index           =   50
               Left            =   6960
               TabIndex        =   136
               Top             =   2280
               Width           =   615
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   "Midi Files"
            Height          =   2295
            Index           =   15
            Left            =   0
            TabIndex        =   121
            Top             =   3480
            Width           =   7815
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   12
               ItemData        =   "genopts.frx":01EF
               Left            =   6360
               List            =   "genopts.frx":01FF
               TabIndex        =   157
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Browse ..."
               Height          =   375
               Index           =   34
               Left            =   6840
               TabIndex        =   122
               Top             =   1200
               Width           =   855
            End
            Begin VB.CheckBox chkgenopt 
               Caption         =   "Randomise playing order of background midi files"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   137
               Top             =   1920
               Width           =   3735
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Clear"
               Height          =   375
               Index           =   33
               Left            =   6840
               TabIndex        =   129
               ToolTipText     =   "Clears the selected entry"
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox gametextvars 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   10
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   128
               Top             =   240
               Width           =   4935
            End
            Begin VB.ListBox Lstgenopts 
               Height          =   1530
               Index           =   2
               ItemData        =   "genopts.frx":020F
               Left            =   120
               List            =   "genopts.frx":0237
               TabIndex        =   127
               Top             =   240
               Width           =   1575
            End
            Begin VB.ListBox Lstgenopts 
               Height          =   900
               Index           =   3
               ItemData        =   "genopts.frx":02ED
               Left            =   1800
               List            =   "genopts.frx":02EF
               TabIndex        =   126
               Top             =   720
               Width           =   4935
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Test"
               Height          =   375
               Index           =   32
               Left            =   6840
               TabIndex        =   141
               Top             =   240
               Width           =   855
            End
            Begin VB.Label labgenopts 
               Height          =   255
               Index           =   54
               Left            =   5280
               TabIndex        =   158
               Top             =   1800
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   " "
         Height          =   6015
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   7815
         Begin VB.ComboBox Cmbgenopts 
            Height          =   330
            Index           =   10
            ItemData        =   "genopts.frx":02F1
            Left            =   2760
            List            =   "genopts.frx":0313
            TabIndex        =   142
            ToolTipText     =   "Please refer to the Monte Carlo selection in Help (F1) before making selection"
            Top             =   5160
            Width           =   1335
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   975
            Index           =   9
            Left            =   240
            TabIndex        =   86
            Top             =   3600
            Width           =   7335
            Begin VB.CommandButton cmdgenopts 
               Height          =   280
               Index           =   26
               Left            =   5280
               TabIndex        =   91
               Top             =   200
               Width           =   1815
            End
            Begin VB.CommandButton cmdgenopts 
               Height          =   280
               Index           =   27
               Left            =   5280
               TabIndex        =   90
               Top             =   600
               Width           =   1815
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   8
               ItemData        =   "genopts.frx":0369
               Left            =   600
               List            =   "genopts.frx":0376
               TabIndex        =   87
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label labgenopts 
               Caption         =   "Use "
               Height          =   255
               Index           =   38
               Left            =   240
               TabIndex        =   89
               Top             =   360
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Default profile on New Game Generation"
               Height          =   255
               Index           =   39
               Left            =   1800
               TabIndex        =   88
               Top             =   360
               Width           =   3255
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   2055
            Index           =   8
            Left            =   240
            TabIndex        =   74
            Top             =   1560
            Width           =   7335
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   11
               ItemData        =   "genopts.frx":0393
               Left            =   6480
               List            =   "genopts.frx":03B2
               TabIndex        =   155
               Top             =   1200
               Width           =   615
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   1
               ItemData        =   "genopts.frx":03D5
               Left            =   1440
               List            =   "genopts.frx":03E5
               TabIndex        =   79
               Top             =   720
               Width           =   855
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   2
               ItemData        =   "genopts.frx":03F9
               Left            =   1440
               List            =   "genopts.frx":040C
               TabIndex        =   78
               Top             =   1200
               Width           =   855
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   3
               ItemData        =   "genopts.frx":0424
               Left            =   1440
               List            =   "genopts.frx":0437
               TabIndex        =   77
               Top             =   1680
               Width           =   855
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   5
               ItemData        =   "genopts.frx":044F
               Left            =   2160
               List            =   "genopts.frx":045C
               TabIndex        =   76
               Top             =   240
               Width           =   975
            End
            Begin VB.Label labgenopts 
               AutoSize        =   -1  'True
               Height          =   210
               Index           =   53
               Left            =   5400
               TabIndex        =   156
               Top             =   1200
               Width           =   45
            End
            Begin VB.Label labgenopts 
               Caption         =   "Spin reels up"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   85
               Top             =   720
               Width           =   975
            End
            Begin VB.Label labgenopts 
               Caption         =   "Percent of the time"
               Height          =   255
               Index           =   3
               Left            =   2400
               TabIndex        =   84
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label labgenopts 
               Caption         =   "Percent of the time"
               Height          =   255
               Index           =   5
               Left            =   2400
               TabIndex        =   83
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label labgenopts 
               Caption         =   "Shorter spins"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   82
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label labgenopts 
               Caption         =   "Percent of the time"
               Height          =   255
               Index           =   7
               Left            =   2400
               TabIndex        =   81
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label labgenopts 
               Caption         =   "Faster spins"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   80
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label labgenopts 
               Caption         =   "Free spin / game interval"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   75
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   1455
            Index           =   7
            Left            =   240
            TabIndex        =   68
            Top             =   120
            Width           =   7335
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   0
               ItemData        =   "genopts.frx":0472
               Left            =   2040
               List            =   "genopts.frx":0488
               TabIndex        =   73
               Top             =   960
               Width           =   975
            End
            Begin VB.CheckBox chkgenopt 
               Caption         =   "Game is restarted with same random seed after leaving Profile Configuration"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   71
               Top             =   600
               Width           =   6135
            End
            Begin VB.CheckBox chkgenopt 
               Caption         =   "Modify selection of random sequence when toggling bet quantity in the Game Screen"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   70
               ToolTipText     =   "This doesn't affect the selection of paylines in a multiline game."
               Top             =   240
               Width           =   6255
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "&New Random Seed Now"
               Height          =   280
               Index           =   3
               Left            =   5400
               TabIndex        =   69
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label labgenopts 
               Caption         =   " "
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   120
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label labgenopts 
               Caption         =   "New random seed every"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   72
               Top             =   960
               Width           =   1695
            End
         End
         Begin VB.CommandButton cmdgenopts 
            Caption         =   "&Change current Hall of Fame MDB directory ..."
            Height          =   280
            Index           =   4
            Left            =   4320
            TabIndex        =   17
            Top             =   5520
            Width           =   3495
         End
         Begin VB.ComboBox Cmbgenopts 
            Height          =   330
            Index           =   4
            ItemData        =   "genopts.frx":04B4
            Left            =   1320
            List            =   "genopts.frx":04CD
            TabIndex        =   8
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CheckBox chkgenopt 
            Caption         =   "Show About page on exit"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   5520
            Width           =   2175
         End
         Begin MSComDlg.CommonDialog Cdlg 
            Left            =   6840
            Top             =   4800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label labgenopts 
            Height          =   255
            Index           =   52
            Left            =   4200
            TabIndex        =   144
            Top             =   5160
            Width           =   3400
         End
         Begin VB.Label labgenopts 
            Caption         =   "Calculate Return with Monte Carlo"
            Height          =   255
            Index           =   51
            Left            =   240
            TabIndex        =   143
            Top             =   5160
            Width           =   2415
         End
         Begin VB.Label labgenopts 
            Caption         =   "Game Limits"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   9
            Top             =   4680
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   " "
         Height          =   6015
         Index           =   3
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   7815
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   1335
            Index           =   6
            Left            =   0
            TabIndex        =   145
            Top             =   4680
            Width           =   5055
            Begin VB.OptionButton optgenopts 
               Caption         =   "Bet Button"
               Height          =   210
               Index           =   5
               Left            =   120
               TabIndex        =   150
               Top             =   840
               Width           =   1095
            End
            Begin VB.OptionButton optgenopts 
               Caption         =   "Spin Button"
               Height          =   210
               Index           =   4
               Left            =   120
               TabIndex        =   149
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Font"
               Height          =   375
               Index           =   24
               Left            =   1440
               TabIndex        =   148
               Top             =   840
               Width           =   735
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Text Colour"
               Height          =   375
               Index           =   25
               Left            =   1440
               TabIndex        =   147
               Top             =   240
               Width           =   1095
            End
            Begin VB.ComboBox Cmbgenopts 
               Height          =   330
               Index           =   7
               ItemData        =   "genopts.frx":0516
               Left            =   3240
               List            =   "genopts.frx":0535
               TabIndex        =   146
               Top             =   240
               Width           =   1695
            End
            Begin ActiveCandy.CandyCommand Candy 
               CausesValidation=   0   'False
               Height          =   375
               Left            =   2400
               TabIndex        =   154
               Top             =   840
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
            Begin VB.Label labgenopts 
               Caption         =   "Style"
               Height          =   255
               Index           =   37
               Left            =   2760
               TabIndex        =   151
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.CheckBox chkgenopt 
            Caption         =   "Big Pictures"
            Height          =   255
            Index           =   3
            Left            =   5640
            TabIndex        =   119
            ToolTipText     =   "Quotes file display is disabled if this option is selected"
            Top             =   5520
            Width           =   1215
         End
         Begin VB.ComboBox Cmbgenopts 
            Height          =   330
            Index           =   6
            ItemData        =   "genopts.frx":05C9
            Left            =   6840
            List            =   "genopts.frx":05D9
            TabIndex        =   66
            Top             =   4920
            Width           =   855
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   735
            Index           =   5
            Left            =   4440
            TabIndex        =   62
            Top             =   3840
            Width           =   3375
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   23
               Left            =   2520
               TabIndex        =   63
               Top             =   240
               Width           =   615
            End
            Begin VB.Label labgenopts 
               Caption         =   "Select General Font"
               Height          =   255
               Index           =   33
               Left            =   120
               TabIndex        =   65
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   35
               Left            =   1800
               TabIndex        =   64
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   735
            Index           =   4
            Left            =   4440
            TabIndex        =   58
            Top             =   3000
            Width           =   3375
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   22
               Left            =   2520
               TabIndex        =   59
               Top             =   240
               Width           =   615
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   34
               Left            =   1800
               TabIndex        =   61
               Top             =   240
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Select Title Font"
               Height          =   255
               Index           =   32
               Left            =   120
               TabIndex        =   60
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   2175
            Index           =   3
            Left            =   0
            TabIndex        =   46
            Top             =   2160
            Width           =   4335
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   21
               Left            =   3120
               TabIndex        =   55
               Top             =   1680
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   20
               Left            =   3120
               TabIndex        =   54
               Top             =   1200
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   2
               Left            =   3120
               TabIndex        =   50
               Top             =   720
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   1
               Left            =   3120
               TabIndex        =   47
               Top             =   240
               Width           =   615
            End
            Begin VB.Line liwinframe 
               BorderColor     =   &H8000000F&
               BorderWidth     =   2
               Index           =   1
               X1              =   2400
               X2              =   2640
               Y1              =   1680
               Y2              =   2040
            End
            Begin VB.Line liwinframe 
               BorderColor     =   &H8000000F&
               BorderWidth     =   2
               Index           =   0
               X1              =   2400
               X2              =   2160
               Y1              =   1680
               Y2              =   2040
            End
            Begin VB.Label labgenopts 
               Caption         =   "Winning Lines Colour"
               Height          =   255
               Index           =   31
               Left            =   240
               TabIndex        =   57
               Top             =   1800
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Caption         =   "Text Highlight Colour"
               Height          =   255
               Index           =   30
               Left            =   240
               TabIndex        =   56
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   26
               Left            =   2160
               TabIndex        =   53
               Top             =   1200
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Prize Highlight Colour"
               Height          =   255
               Index           =   29
               Left            =   240
               TabIndex        =   52
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   25
               Left            =   2160
               TabIndex        =   51
               Top             =   720
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Prize - Money Colour"
               Height          =   255
               Index           =   28
               Left            =   240
               TabIndex        =   49
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   24
               Left            =   2160
               TabIndex        =   48
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   "Prize Label Colours"
            Height          =   2895
            Index           =   2
            Left            =   4440
            TabIndex        =   18
            Top             =   0
            Width           =   3375
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   15
               Left            =   2520
               TabIndex        =   28
               Top             =   2400
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   14
               Left            =   2520
               TabIndex        =   27
               Top             =   1920
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   13
               Left            =   2520
               TabIndex        =   26
               Top             =   1440
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   12
               Left            =   2520
               TabIndex        =   25
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   11
               Left            =   2520
               TabIndex        =   24
               Top             =   480
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   10
               Left            =   1560
               TabIndex        =   23
               Top             =   2400
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   9
               Left            =   1560
               TabIndex        =   22
               Top             =   1920
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   8
               Left            =   1560
               TabIndex        =   21
               Top             =   1440
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   7
               Left            =   1560
               TabIndex        =   20
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   6
               Left            =   1560
               TabIndex        =   19
               Top             =   480
               Width           =   615
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   21
               Left            =   960
               TabIndex        =   39
               Top             =   2400
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   20
               Left            =   960
               TabIndex        =   38
               Top             =   1920
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   19
               Left            =   960
               TabIndex        =   37
               Top             =   1440
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   18
               Left            =   960
               TabIndex        =   36
               Top             =   960
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   17
               Left            =   960
               TabIndex        =   35
               Top             =   480
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "1 Picture"
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   34
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label labgenopts 
               Caption         =   "2 Pictures"
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   33
               Top             =   2040
               Width           =   735
            End
            Begin VB.Label labgenopts 
               Caption         =   "3 Pictures"
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   32
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label labgenopts 
               Caption         =   "4 Pictures"
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label labgenopts 
               Caption         =   "5 Pictures"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Width           =   735
            End
            Begin VB.Label labgenopts 
               Caption         =   "Background        Text"
               Height          =   255
               Index           =   11
               Left            =   1440
               TabIndex        =   29
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   975
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   1080
            Width           =   4335
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   19
               Left            =   3600
               TabIndex        =   42
               Top             =   480
               Width           =   615
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   17
               Left            =   2760
               TabIndex        =   16
               ToolTipText     =   "Title Bitmap is ignored if it has the same path name as Wallpaper Bitmap"
               Top             =   480
               Width           =   615
            End
            Begin VB.OptionButton optgenopts 
               Caption         =   "Use Title Colour"
               Height          =   210
               Index           =   3
               Left            =   120
               TabIndex        =   15
               Top             =   600
               Width           =   1575
            End
            Begin VB.OptionButton optgenopts 
               Caption         =   "Use Title Bitmap"
               Height          =   210
               Index           =   2
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   23
               Left            =   2160
               TabIndex        =   45
               Top             =   480
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Background       Text"
               Height          =   255
               Index           =   27
               Left            =   2640
               TabIndex        =   43
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame fragenopts 
            Caption         =   " "
            Height          =   975
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   4335
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   18
               Left            =   3600
               TabIndex        =   40
               Top             =   480
               Width           =   615
            End
            Begin VB.OptionButton optgenopts 
               Caption         =   "Use Wallpaper Bitmap"
               Height          =   210
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   360
               Width           =   1935
            End
            Begin VB.OptionButton optgenopts 
               Caption         =   "Use Wallpaper Colour"
               Height          =   210
               Index           =   1
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   1935
            End
            Begin VB.CommandButton cmdgenopts 
               Caption         =   "Select "
               Height          =   375
               Index           =   16
               Left            =   2760
               TabIndex        =   6
               ToolTipText     =   "Wallpaper and Title Bitmaps with same name will be merged"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label labgenopts 
               Alignment       =   2  'Center
               Caption         =   "A"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   22
               Left            =   2160
               TabIndex        =   44
               Top             =   480
               Width           =   375
            End
            Begin VB.Label labgenopts 
               Caption         =   "Background       Text"
               Height          =   255
               Index           =   9
               Left            =   2640
               TabIndex        =   41
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Label labgenopts 
            Caption         =   "Flash Prizes"
            Height          =   255
            Index           =   36
            Left            =   5640
            TabIndex        =   67
            Top             =   4920
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "Genopts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim numdevices As Long, m_Media As String, m_WaveFile As String, col() As String
Dim ct As Long, ct1 As Long, response As Long, gt152 As Long, listcomb As Long, m_SysSnds() As SystemSoundDefinitions
Dim setbackbmpdir As Boolean, Spinbuttonsel As Boolean, booRestoredefaults As Boolean, boosavedefaults As Boolean, procchk As Boolean

'listcomb = 0 to 8
Private Sub Form_Load()

setformpos Me

Loadinvals      'needed for default refresh
End Sub
Private Function QuotebrsShow() As Boolean
Dim f As Form 'For DB reference
QuotebrsShow = False
For Each f In Forms
If f.Name = "Quotebrs" Then
Quotebrs.Show
QuotebrsShow = True
Exit Function
End If
Next
End Function
Private Sub Form_Activate()
QuotebrsShow
End Sub
Private Sub Loadinvals()
procend = False

setbackbmpdir = False
procchk = True
If gt(152) < 0 Then
gt152 = -gt(152)
Else
gt152 = gt(152)
End If
resetdefaultbutton

Genopts.Caption = "General Options"  'Space(40)

m_Media = Space$(260)
Call GetWindowsDirectory(m_Media, Len(m_Media))
m_Media = Left(m_Media, InStr(m_Media, vbNullChar) - 1) & "\Media\"
   
If FillList = True Then

ReDim col(Cmbgenopts(13).ListCount - 1)

'Put the values of the items in the array
For ct = 0 To Cmbgenopts(13).ListCount - 1
col(ct) = Cmbgenopts(13).List(ct)
Next

End If
  
Me.Icon = Nothing

For ct = 26 To 38
Lstgenopts(1).List(ct - 26) = Stringvars(ct)
Next
For ct = 39 To 50
Lstgenopts(3).List(ct - 39) = Stringvars(ct)
Next
For ct = 0 To 1
Lstgenopts(2 * ct).Selected(0) = True
Lstgenopts(2 * ct + 1).Selected(0) = True
gametextvars(9 + ct).Text = Stringvars(26 + 13 * ct)
Next

labgenopts(24).ForeColor = gt(175)
labgenopts(25).BackColor = gt(176)
labgenopts(26).ForeColor = gt(177)
liwinframe(0).BorderColor = gt(178)
liwinframe(1).BorderColor = gt(178)



For ct = 17 To 21
labgenopts(ct).BackColor = gt(148 + ct)
labgenopts(ct).ForeColor = gt(153 + ct)
Next

labgenopts(34).Font.Name = Stringvars(9)
labgenopts(35).Font.Name = Stringvars(10)
With Candy
.fontname = Stringvars(11)
.Fontsize = 12 * textwidthratio
.Caption = Stringvars(6)
.ForeColor = gt(179)
End With

Spinbuttonsel = True
optgenopts(4).Value = True


For ct = 0 To 3
gametextvars(ct).Text = Stringvars(ct + 5)
Next
For ct = 4 To 7
gametextvars(ct).Text = Stringvars(ct + 9)
Next


For ct = 0 To 1
labgenopts(22 + ct).ForeColor = gt(163 + ct)
If gt(161 + ct) = -1 Then
optgenopts(2 * ct).Value = True
Else
labgenopts(22 + ct).BackColor = gt(161 + ct)
optgenopts(2 * ct + 1).Value = True
End If
Next

If gt(37) = 1 Then
chkgenopt(0).Value = 1
Else
chkgenopt(0).Value = 0
End If
If gt(45) = 1 Then
chkgenopt(1).Value = 1
Else
chkgenopt(1).Value = 0
End If
If gt(46) = 1 Then
chkgenopt(2).Value = 1
Else
chkgenopt(2).Value = 0
End If
If gt(159) = 1 Then
chkgenopt(3).Value = 1
chkgenopt(6).Enabled = False
chkgenopt(7).Enabled = False
chkgenopt(10).Enabled = False
cmdgenopts(35).Enabled = False
cmdgenopts(5).Enabled = False
Else
chkgenopt(3).Value = 0
    'Allow Quotes for small pics
    If Stringvars(3) <> "" Then
    cmdgenopts(5).Enabled = True
        If gt(193) = 0 Then
        chkgenopt(10).Value = 0
        cmdgenopts(35).Enabled = True
        cmdgenopts(35).ToolTipText = "Space required :Allow for at least twice text file size. Free space : " & vbGetAvailableKBytesAsString(loaddirectory) & "k"
        Else
        cmdgenopts(35).Enabled = False
        chkgenopt(10).Value = 1
        End If
    chkgenopt(6).Value = 1
    chkgenopt(7).Enabled = True
    chkgenopt(10).Enabled = True
    Else
    chkgenopt(6).Value = 0
    chkgenopt(7).Enabled = False
    chkgenopt(10).Enabled = False
    cmdgenopts(5).Enabled = False
    cmdgenopts(35).Enabled = False
    cmdgenopts(35).ToolTipText = ""
    End If
End If


chkgenopt(4).Value = gt(183)
If gt(187) < 0 Then
chkgenopt(5).Value = 1
Else
chkgenopt(5).Value = 0
End If

If gt(195) = 1 Then
chkgenopt(7).Value = 1
Else
chkgenopt(7).Value = 0
End If


EnableNoise CBool(gt(186) > 1), True
EnableNoise CBool(gt(185))

Select Case gt(186)
Case Is < 2
chkgenopt(8).Value = 0
chkgenopt(9).Value = gt(186)
chkgenopt(9).Enabled = False
Case Else
chkgenopt(8).Value = 1
chkgenopt(9).Value = gt(186) - 2
If gt(159) = 0 Then
chkgenopt(9).Enabled = True
Else
chkgenopt(9).Enabled = False
End If
End Select


chkgenopt(10).Value = gt(193)    'Use base dir



labgenopts(1) = "starting next " & Cmbgenopts(0).List(gt(36))

listcomb = 0

For ct = 0 To 12
Randomcmbselect ct
Next

TTDefault
PCspeed

Cmbgenopts(12).ToolTipText = ""
If gt(185) = 0 Then
labgenopts(54).Caption = "Midi Unused"
Else
labgenopts(54).Caption = "Midi Device"
End If


If gt(2) = 0 Or gt(156) = 1 Then    'EOG status
Cmbgenopts(4).Enabled = False
cmdgenopts(27).Enabled = False
End If

procend = True
End Sub
Private Sub Cmbgenopts_Click(Index As Integer)
Dim teststring As String

If procend = False Then Exit Sub
procend = False
Select Case Index
Case 0
'gt(35) current seed
'gt(36) current seed-change interval index
gt(36) = Cmbgenopts(0).ListIndex

gt(38) = Second(Time)
gt(39) = Minute(Time)
gt(40) = Hour(Time)
gt(41) = Day(Date) '0 to 31 day of month
gt(42) = Month(Date)
gt(43) = Year(Date)

labgenopts(1) = "starting next " & Cmbgenopts(0).Text

Case 4
gt(47) = Val(Cmbgenopts(4).Text)
Case 5
gt(157) = Cmbgenopts(5).ListIndex
Case 6
gt(160) = Cmbgenopts(6).ListIndex
Case 7
If Spinbuttonsel = True Then
gt(180) = Cmbgenopts(7).ListIndex
Candy.BackPicture = gt(180)
Else
gt(182) = Cmbgenopts(7).ListIndex
Candy.BackPicture = gt(182)
End If
Case 8  'defaults
gt(158) = Cmbgenopts(8).ListIndex
  If Cmbgenopts(8).ListIndex = 0 And IsDefaultSaved = False Then
  gt(158) = 1
  Cmbgenopts(8).ListIndex = 1
  Cmbgenopts(8).ToolTipText = "First click Save Defaults Now!"
  Else
    If cmdgenopts(27).Caption = "<= &Saved or Generic?" Then
     If gt(158) = 1 Then
       If IsDefaultSaved = True Then
       gt(158) = 0
       Cmbgenopts(8).ListIndex = 0
       Else
       Cmbgenopts(8).ListIndex = 2
       gt(158) = 2
       End If
     End If
    Else
    TTDefault
    End If
  End If
Case 9
listcomb = Cmbgenopts(9).ListIndex
gametextvars(8).Text = Stringvars(17 + listcomb)
Case 10 'Monte Carlo
If gt152 <> Val(Cmbgenopts(10).Text) Then
    If Val(Cmbgenopts(10).Text) > 0 Then
    gametype.cmdgametype(4).BackColor = &H80FF&
    gametype.cmdgametype(0).ToolTipText = "You need to Recalculate Return to continue"
    Else
    'zero VOG vars
    For ct = 29 To 34
    gt(ct) = 0
    Next
    gametype.cmdgametype(4).BackColor = vbButtonFace
    gametype.cmdgametype(0).ToolTipText = ""
    End If
gt152 = Val(Cmbgenopts(10).Text)
gt(152) = gt152
MCarlotime
End If
Case 11 'Fast Pc
gt(194) = CLng(Cmbgenopts(11).Text)
PCspeed
Case 12
'Always turn off sound before changing
Cmbgenopts(12).ToolTipText = ""

If Cmbgenopts(12).ListIndex = 0 Then
  ct = 0
  stopsound
  If Midichg(ct) = gt(185) Then
  labgenopts(54).Caption = "Midi Unused"
  Else
  labgenopts(54).Caption = "Midi Error"
  gt(185) = 0
  End If
  EnableNoise False
Else
 If gt(185) = 0 Then 'initialise
  If Midichg = gt(185) Then 'Midichg changes gt(185) on success!
   If gt(185) = 0 Then
   labgenopts(54).Caption = "Midi Error"
   EnableNoise False
   Else
   'success: greater than 1
   If Midichg(Cmbgenopts(12).ListIndex) = 0 Then
   gt(185) = 0
   labgenopts(54).Caption = "Midi Unused"
   Cmbgenopts(12).ListIndex = 0
   Else
   labgenopts(54).Caption = "Midi Device"
   EnableNoise True
   End If
   End If
  Else
  labgenopts(54).Caption = "Midi Error"
  gt(185) = 0 ' just in case
  Cmbgenopts(12).ListIndex = 0
  EnableNoise False
  End If
  Cmbgenopts(12).ListIndex = gt(185)
 Else
 stopsound
 Stopnoise 2
   If Stopnoise(2) Then
   Cmbgenopts(12).ToolTipText = "Handle for " & gt(185) & " did not close!"
   EnableNoise False
   Cmbgenopts(12).ListIndex = 0
   Else
     If Midichg(Cmbgenopts(12).ListIndex) = 0 Then
       If gt(185) = 0 Then
       Cmbgenopts(12).ListIndex = 0
       EnableNoise False
       Else
       Cmbgenopts(12).ListIndex = gt(185)
       End If
     End If
   End If
 End If
End If
Case 13
teststring = m_SysSnds(Cmbgenopts(13).ItemData(Cmbgenopts(13).ListIndex)).Current
If InStr(teststring, ":") = 0 Then teststring = m_Media & teststring
If namelt100chars(teststring) = True Then gametextvars(9).Text = teststring
Case Else
gt(5 + Index) = Cmbgenopts(Index).ListIndex
End Select

procend = True
End Sub
Private Sub Lstgenopts_Click(Index As Integer)
If procend = False Then Exit Sub
procend = False
Select Case Index
Case 0, 2
With Lstgenopts(Index + 1)
For ct = 0 To .ListCount - 1
.Selected(ct) = False
Next
.Selected(Lstgenopts(Index).ListIndex) = True
gametextvars(9 + Index / 2).Text = .List(.ListIndex)
End With
Case 1, 3
With Lstgenopts(Index - 1)
For ct = 0 To .ListCount - 1
.Selected(ct) = False
Next
.Selected(Lstgenopts(Index).ListIndex) = True
gametextvars(9 + (Index - 1) / 2).Text = Lstgenopts(Index).List(Lstgenopts(Index).ListIndex)
End With
End Select
procend = True
End Sub
Private Sub Cmbgenopts_GotFocus(Index As Integer)
    Const CB_SHOWDROPDOWN = &H14F
    Dim tmp
    tmp = SendMessage(Cmbgenopts(Index).hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cmbgenopts_Change(Index As Integer)
If procend = True Then
Randomcmbselect CLng(Index)
Cmbgenopts(Index).Refresh
End If
End Sub
Private Sub Cmbgenopts_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim ctext As String, curpos As Integer
If KeyCode = vbKeyBack Then Exit Sub
ctext = Cmbgenopts(Index).Text
If ctext = "" Then Exit Sub
curpos = Cmbgenopts(Index).SelStart
For ct = 0 To UBound(col)
        If LCase(col(ct)) Like LCase(Cmbgenopts(Index).Text) & "*" Then
        With Cmbgenopts(Index)
        .Text = .Text & Right(col(ct), Len(col(ct)) - Len(ctext))
        .SelStart = curpos
        .SelLength = Len(.Text)
        End With
        Exit Sub
        End If
Next
End Sub
Private Sub optgenopts_Click(Index As Integer)
If procend = False Then Exit Sub
procend = False

With Candy 'Careful of with

Select Case Index
Case 0
gt(161) = -1
labgenopts(22).BackStyle = 0
Case 1
If gt(161) = -1 Then gt(161) = 0
labgenopts(22).BackStyle = 1
labgenopts(22).BackColor = gt(161)
Case 2
gt(162) = -1
labgenopts(23).BackStyle = 0
Case 3
If gt(162) = -1 Then gt(162) = 0
labgenopts(23).BackStyle = 1
labgenopts(23).BackColor = gt(162)
Case 4
.fontname = Stringvars(11)
.Fontsize = 12 * textwidthratio
.Caption = Stringvars(6)
.ForeColor = gt(179)
.BackPicture = gt(180)
Spinbuttonsel = True
Cmbgenopts(7).Text = Cmbgenopts(7).List(gt(180))
Case 5
.fontname = Stringvars(12)
.Fontsize = 12 * textwidthratio
.Caption = Stringvars(7)
.ForeColor = gt(181)
.BackPicture = gt(182)
Spinbuttonsel = False
Cmbgenopts(7).Text = Cmbgenopts(7).List(gt(182))
End Select

End With
procend = True
End Sub
Private Sub sstaboption_Click(PreviousTab As Integer)
If procend = False Then Exit Sub
Genopts.HelpContextID = 12 + sstaboption.Tab
stopsound
resetdefaultbutton
End Sub
Private Sub chkgenopt_Click(Index As Integer)
If procend = False Or procchk = False Then Exit Sub
procchk = False
resetdefaultbutton
With chkgenopt(Index)

Select Case Index
Case 0
If .Value = 1 Then
gt(37) = 1
Else
gt(37) = 0
End If
Case 1
If .Value = 1 Then
gt(45) = 1
Else
gt(45) = 0
End If
Case 2
If .Value = 1 Then
gt(46) = 1
Else
gt(46) = 0
End If
Case 3
If .Value = 1 Then
gt(159) = 1
chkgenopt(6).Enabled = False
chkgenopt(7).Enabled = False
chkgenopt(9).Enabled = False
chkgenopt(10).Enabled = False
cmdgenopts(5).Enabled = False
cmdgenopts(35).Enabled = False
cmdgenopts(35).ToolTipText = ""
chkgenopt(6).Value = 0
Else
chkgenopt(6).Enabled = True
If Stringvars(3) <> "" Then
chkgenopt(6).Value = 1
chkgenopt(7).Enabled = True

chkgenopt(9).Enabled = True
If gt(186) < 2 Then 'Quote Thumbnail Sound
chkgenopt(9).Value = gt(186)
Else
chkgenopt(9).Value = gt(186) - 2
End If

chkgenopt(10).Enabled = True
cmdgenopts(5).Enabled = True
If gt(193) = 0 Then cmdgenopts(35).Enabled = True
End If
gt(159) = 0
End If
Case 4
gt(183) = .Value
If FillList = True Then

ReDim col(Cmbgenopts(13).ListCount - 1)

'Put the values of the items in the array
For ct = 0 To Cmbgenopts(13).ListCount - 1
col(ct) = Cmbgenopts(13).List(ct)
Next

End If

Case 5
gt(187) = -gt(187)
Case 6  'Quotes File
If .Value = 1 Then
response = MsgBox("Please click OK to use quote.s$t in current directory or Cancel to Use Base Dir.", vbOKCancel)
    If response = vbOK Then
        If Getquotez(False) = True Then
        chkgenopt(10).Value = 0
        cmdgenopts(35).Enabled = True
        cmdgenopts(35).ToolTipText = "Space required : Allow for at least twice text file size. Free space : " & vbGetAvailableKBytesAsString(loaddirectory) & "k"
        Else
        procchk = True
        Exit Sub
        End If
    Else
        If Getquotez(True) = True Then
        chkgenopt(10).Value = 1
        cmdgenopts(35).Enabled = False
        cmdgenopts(35).ToolTipText = ""
        Else
        procchk = True
        Exit Sub
        End If
    End If
cmdgenopts(5).Enabled = True
chkgenopt(7).Enabled = True
chkgenopt(10).Enabled = True
Else
gt(195) = 0
Stringvars(3) = ""
chkgenopt(7).Enabled = False
chkgenopt(10).Enabled = False
chkgenopt(7).Value = 0
cmdgenopts(5).Enabled = False
cmdgenopts(35).Enabled = False
End If
Case 7  'Randomise quotes
gt(195) = .Value
Case 8
If .Value = 0 Then
gt(186) = chkgenopt(9).Value
chkgenopt(9).Enabled = False
Else
gt(186) = 2 + chkgenopt(9).Value
If gt(159) = 0 Then chkgenopt(9).Enabled = True
End If
EnableNoise CBool(.Value), True
Case 9  'Thumbnail Sound: only enabled when chkgenopt(8) enabled
gt(186) = 2 + .Value
Case 10  'base dir
If Getquotez(CBool(.Value)) = False Then
If CBool(.Value) = False Then
.Value = 1
Else
.Value = 0
End If
End If
End Select
End With
Quotezerr:
procchk = True
End Sub
Private Sub cmdgenopts_Click(Index As Integer)
procend = False

Select Case Index
Case 0
'Save sounds
For ct = 26 To 38
Stringvars(ct) = Lstgenopts(1).List(ct - 26)
Next
For ct = 39 To 50
Stringvars(ct) = Lstgenopts(3).List(ct - 39)
Next
stopsound
Unload Me
Set Genopts = Nothing
Case 1, 2
If Index = 1 Then
gt(175) = Colordlg(175)
labgenopts(24).ForeColor = gt(175)
Else
gt(176) = Colordlg(176)
labgenopts(25).BackColor = gt(176)
End If
Case 3
genrandomseed True
Rnd -1     'reset seeding
Randomize (gt(35))
Case 4
With Cdlg
.FileName = ""
.Filter = "MyReels (Hallfame.s$t)|Hallfame.s$t|All Files (*.*)|*.*"
.Flags = FileOpenConstants.cdlOFNHideReadOnly
.DialogTitle = "Open Hallfame.s$t"
' Specify default filter
.FilterIndex = 1
' Set CancelError is True
On Error GoTo NoAction
.InitDir = CurDir
If gt(0) = 0 Then .CancelError = True
.ShowOpen


Stringvars(4) = Left(.FileName, Len(.FileName) - 12)
If namelt100chars(Stringvars(4)) = False Then
Stringvars(4) = ""
ChDir (loaddirectory)
ChDrive (Left$(loaddirectory, 1))
procend = True
Exit Sub
End If
 
End With
Case 5
'Quote thumbnails
LoadFrmSplsh 430
Load Quotebrs
Quotebrs.Show
Unload frmSplsh
Case 6 To 15
gt(159 + Index) = Colordlg(159 + Index)
If Index < 11 Then
labgenopts(11 + Index).BackColor = gt(159 + Index)
Else
labgenopts(6 + Index).ForeColor = gt(159 + Index)
End If
Case 16, 17
If gt(145 + Index) > -1 Then
gt(145 + Index) = Colordlg(145 + Index)
labgenopts(Index + 6).BackStyle = 1
labgenopts(Index + 6).BackColor = gt(145 + Index)
Else
With Cdlg
.FileName = ""
.FilterIndex = 1
.Filter = "Picture (*.bmp;*.wmf;*.gif;*.jpg)|*.bmp;*.wmf;*.gif;*.jpg|All Files (*.*)|*.*"
' Specify default filter
.DialogTitle = " Open Bitmap"
    
    If setbackbmpdir = False Then

    For ct = Len(Stringvars(Index - 15)) To 2 Step -1
    If Mid(Stringvars(Index - 15), ct, 1) = "\" Then Exit For
    Next
    On Error GoTo Notvaliddirectory 'just in case dir is deleted by user
    .InitDir = Left(Stringvars(Index - 15), ct)
    setbackbmpdir = True


    Else    'setbackbmpdir true so choose curdir

    .InitDir = CurDir
    
    End If

.CancelError = True
On Error GoTo NoAction
.ShowOpen

        If namelt100chars(.FileName) = False Then
        ChDrive (Left$(loaddirectory, 1))
        ChDir (loaddirectory)
        procend = True
        Exit Sub
        End If

Stringvars(Index - 15) = CStr(.FileName)
labgenopts(6 + Index).BackStyle = 0
End With
End If


Case 18, 19
gt(145 + Index) = Colordlg(145 + Index)
labgenopts(4 + Index).ForeColor = gt(145 + Index)

Case 20, 21
gt(157 + Index) = Colordlg(157 + Index)
If Index = 20 Then
labgenopts(26).ForeColor = gt(177)
Else
liwinframe(0).BorderColor = gt(178)
liwinframe(1).BorderColor = gt(178)
End If
Case 22, 23, 24 'Fonts
With Cdlg
.CancelError = True
' Set the Flags property - true type fonts can only be rotated
.Flags = cdlCFScalableOnly Or cdlCFTTOnly Or cdlCFScreenFonts
' Display the Font dialog box
On Error GoTo NoAction
.fontname = Stringvars(Index - 13)
.ShowFont

If Len(.fontname) > 100 Then
response = MsgBox("The Fonts name is too long. Please use a different font or shorten its name.", vbOKOnly)
procend = True
Exit Sub
End If

If Index < 24 Then
labgenopts(Index + 12).fontname = .fontname
Stringvars(Index - 13) = .fontname
Else
    If Spinbuttonsel = True Then
    Stringvars(11) = .fontname
    Else
    Stringvars(12) = .fontname
    End If
Candy.Fontsize = 12 * textwidthratio
Candy.fontname = .fontname
End If

End With
Case 25
If Spinbuttonsel = True Then
gt(179) = Colordlg(179)
Else
gt(181) = Colordlg(181)
End If

Candy.ForeColor = Cdlg.Color

Case 26
If boosavedefaults = False Then
resetdefaultbutton
  If IsDefaultSaved = False Then
  savedefault
  Else
  boosavedefaults = True
  cmdgenopts(26).Caption = "S&ure?"
  cmdgenopts(26).ToolTipText = "Click to overwrite previously Saved Defaults with Current settings."
  End If
Else
  savedefault
  cmdgenopts(26).Caption = "Defaults Saved"
End If
Case 27
If booRestoredefaults = False Then
 resetdefaultbutton
 labgenopts(38).Visible = False
 labgenopts(39).Visible = False
 booRestoredefaults = True
   If gt(158) = 1 Then
     If IsDefaultSaved = True Then
     gt(158) = 0
     Cmbgenopts(8).ListIndex = 0
     Else
     Cmbgenopts(8).ListIndex = 2
     gt(158) = 2
     End If
   End If
 If IsDefaultSaved = True Then
 cmdgenopts(27).Caption = "<= &Saved or Generic?"
 cmdgenopts(27).ToolTipText = "Click to continue restoring Saved or Generic Defaults. Not applicable when current defaults are restored."
 labgenopts(38).Visible = False
 labgenopts(39).Visible = False
 booRestoredefaults = True
 Else
 cmdgenopts(27).Caption = "<&= Restore Generic?"
 cmdgenopts(27).ToolTipText = "Click to continue restoring Generic Defaults. No defaults have been saved!"
 Cmbgenopts(8).Enabled = False
 End If
Else
 If gt(158) = 1 Then
 resetdefaultbutton
 Else
 'Restore from scratch if we havn't started spinning
 Stopnoise 1
 Restoredefaults (gt(150) = 0 And gt(49) + gt(50) + gt(51) = 0)
 Loadinvals
 If gt(185) > 0 Then SndMidInit
 gametype.loaddefaultz
 cmdgenopts(27).Caption = "Defaults Restored"
 End If
End If
Case 28
m_WaveFile = gametextvars(9).Text
Call PlaySndF(m_WaveFile)
Case 29
With Lstgenopts(1)
If .ListIndex < 0 Then
procend = True
Exit Sub
End If
Stringvars(26 + .ListIndex) = gametextvars(9).Text
.List(.ListIndex) = gametextvars(9).Text
End With
Case 30
With Lstgenopts(1)
gametextvars(9).Text = ""
If .ListIndex < 0 Then
procend = True
Exit Sub
End If
Stringvars(26 + .ListIndex) = ""
.List(.ListIndex) = ""
End With
Case 31
With Cdlg
For ct = Len(Stringvars(26 + Lstgenopts(1).ListIndex)) To 2 Step -1
If Mid(Stringvars(26 + Lstgenopts(1).ListIndex), ct, 1) = "\" Then Exit For
Next
.FileName = ""
.FilterIndex = 1
.Filter = "Wave (*.wav)|*.wav|All Files (*.*)|*.*"
' Specify default filter
.DialogTitle = " Open Wave File"

On Error GoTo Notvaliddirectory 'just in case dir is deleted by user
.InitDir = Left(Stringvars(26 + Lstgenopts(1).ListIndex), ct)
.CancelError = True
On Error GoTo NoAction
.ShowOpen

If namelt100chars(.FileName) = False Then
procend = True
Exit Sub
End If

gametextvars(9).Text = CStr(.FileName)
End With
Case 32
   If gametextvars(10).Text <> "none" Then
      If cmdgenopts(32).Caption = "Test" Then
         MidiPlay.Enabled = True
      Else
         MidiPlay.Interval = 1
         DoMidiPlay
      End If
   End If
Case 33
With Lstgenopts(3)
gametextvars(10).Text = ""
If .ListIndex < 0 Then
procend = True
Exit Sub
End If
Stringvars(39 + .ListIndex) = ""
.List(.ListIndex) = ""
End With

Case 34
With Cdlg
For ct = Len(Stringvars(39 + Lstgenopts(3).ListIndex)) To 2 Step -1
If Mid(Stringvars(39 + Lstgenopts(3).ListIndex), ct, 1) = "\" Then Exit For
Next
.FileName = ""
.FilterIndex = 1
.Filter = "Midi (*.mid)|*.mid|All Files (*.*)|*.*"
' Specify default filter
.DialogTitle = " Open Midi File"

On Error GoTo Notvaliddirectory 'just in case dir is deleted by user
.InitDir = Left(Stringvars(39 + Lstgenopts(3).ListIndex), ct)

.CancelError = True
On Error GoTo NoAction
.ShowOpen

If namelt100chars(.FileName) = False Or Lstgenopts(3).ListIndex < 0 Then
procend = True
Exit Sub
End If

gametextvars(10).Text = CStr(.FileName)
Stringvars(39 + Lstgenopts(3).ListIndex) = CStr(.FileName)
Lstgenopts(3).List(Lstgenopts(3).ListIndex) = CStr(.FileName)

End With
If gt(187) < 0 Then
gt(187) = -1
Else
gt(187) = 1
End If


Case 35 'quotes
With Cdlg

.FileName = ""
.FilterIndex = 1
.Filter = "text (*.,txt)|*.txt|All Files (*.*)|*.*"
'Specify default filter
.DialogTitle = " Open Text File"

On Error GoTo Notvaliddirectory 'just in case dir is deleted by user
.InitDir = Stringvars(3)

.CancelError = True
On Error GoTo NoAction
.ShowOpen


If namelt100chars(.FileName) = False Then
procend = True
Exit Sub
End If
If App.Path & "\" = loaddirectory Then
response = MsgBox("This action will erase all thumbnails from the quotes database in the MyReels base directory. Do you wish to continue?", vbYesNo)
If response = vbNo Then
procend = True
Exit Sub
End If
End If

If .FileTitle = "Dummy.txt" Then
MsgBox "Please choose a filename other than ""Dummy.txt"""
procend = True
Exit Sub
End If

cmdgenopts(35).ToolTipText = ""
Genopts.MousePointer = vbHourglass
sDatabaseName = Stringvars(3) & "Quotes.s$t"
If Openquotes(.FileName, True) = False Then
MsgBox "Generation of quotes failed. To use the database reset the ""Show Quotes"" checkbox.", vbOKOnly
Stringvars(3) = ""
Else
sDatabaseName = Stringvars(3) & "Quotes.s$t"
If compactdb(2) = False Then Exit Sub
sDatabaseName = ""
End If

Genopts.MousePointer = vbDefault
End With

End Select

procend = True
Exit Sub

Notvaliddirectory:
With Cdlg
.InitDir = CurDir
On Error GoTo NoAction
.ShowOpen
If namelt100chars(.FileName) = False Then
procend = True
Exit Sub
End If
Select Case Index

Case 16, 17
Stringvars(Index - 15) = CStr(.FileName)
labgenopts(6 + Index).BackStyle = 0

Case 31
gametextvars(9).Text = CStr(.FileName)

Case 34
gametextvars(10).Text = CStr(.FileName)
If Lstgenopts(3).ListIndex < 0 Then
procend = True
Exit Sub
End If
Stringvars(39 + Lstgenopts(3).ListIndex) = CStr(.FileName)
Lstgenopts(3).List(Lstgenopts(3).ListIndex) = CStr(.FileName)
Case 35
Stringvars(3) = loaddirectory
End Select


Stringvars(Index) = CStr(.FileName)
setbackbmpdir = True
End With

NoAction:
procend = True
If Err.Number <> 32755 Then ShowError
End Sub
Private Sub MidiPlay_Timer()
' Share timer regarding Quotebrs show issue
If QuotebrsShow Then
  Gametype.enabled = true
  Genopts.enabled = true
  MidiPlay.Enabled = False
  MidiPlay.Interval = 18
Else
DoMidiPlay
End If
End Sub
Private Sub DoMidiPlay()
With cmdgenopts(32)
If MidiPlay.Interval = 1 Then
  Call StopMidiFile
  MidiPlay.Enabled = False
  MidiPlay.Interval = 18
  .Caption = "Test"
Else
  If .Caption = "Stop" Then Exit Sub
  If PlayMidiFile(gametextvars(10).Text, False) Then
  .Caption = "Stop"
  Else
  MidiPlay.Enabled = False
  End If
End If
End With
End Sub
Private Sub stopsound()
Call PlaySndF("", True)

  If cmdgenopts(32).Caption = "Stop" Then
  MidiPlay.Interval = 1
  DoMidiPlay
  End If
End Sub
Private Sub gametextvars_Change(Index As Integer)
If procend = False Then Exit Sub
procend = False
With gametextvars(Index)
.ToolTipText = ""
If .Text <> "" And Namevalid(2, .Text) = True Then
Select Case Index
    Case Is < 3
        If Len(.Text) > 20 Then      'no more than 20 chars
        .Text = Left(.Text, 20)
        .SetFocus
        .ToolTipText = "20 Characters or less"
        procend = True
        Exit Sub
        End If
    Case Else
        If Len(.Text) > 30 Then     'no more than 30 chars
        .Text = Left(.Text, 30)
        .SetFocus
        .ToolTipText = "30 Characters or less"
        procend = True
        Exit Sub    'Definitely no stringvars update in either case
        End If
    End Select
    
    Select Case Index
    Case Is < 4
    Stringvars(Index + 5) = .Text
    Case 4 To 7
    Stringvars(Index + 9) = .Text
    Case Else
    Stringvars(17 + listcomb) = .Text
    End Select
Else
        Beep
        If Index < 4 Then
        .Text = Stringvars(Index + 5)
        Else
        .Text = Stringvars(Index + 9)
        End If
End If
End With
procend = True
End Sub
Private Sub Randomcmbselect(Index As Long)
Dim c1 As Long, tmp As Long, errno As Long
With Cmbgenopts(Index)
Select Case Index
Case 0
'time base random seed
.Text = .List(gt(36))
Case 1  'Spin direction
.Text = .List(gt(6))
Case 2  'shorter spins
.Text = .List(gt(7))
Case 3  'faster spins
.Text = .List(gt(8))
Case 4

'Not allowed to decrease limit
For c1 = 0 To 5
response = 0    'cheap use of response
If gt(2) >= 10000 * 10 ^ c1 Then response = 1
For ct1 = 48 To 150
If gt(ct1) >= 10000 * 10 ^ c1 Then response = 1
Next
If response = 1 Then .RemoveItem 0
Next

.Refresh
.Text = CStr(gt(47))
Case 5
.Text = .List(gt(157))
Case 6
.Text = .List(gt(160))
Case 7
If Spinbuttonsel = True Then
.Text = .List(gt(180))
Candy.BackPicture = gt(180)
Else
.Text = .List(gt(182))
Candy.BackPicture = gt(182)
End If
Case 8  'defaults
.Text = .List(gt(158))
Case 9
.Text = .List(listcomb)
gametextvars(8).Text = Stringvars(17 + listcomb)
Case 10 'Monte Carlo
 .Text = CStr(gt152)
MCarlotime
'Only at beginning of game
If gt(150) > 0 Or gt(49) + gt(50) + gt(51) > 0 Then .Enabled = False
Case 11



Select Case gt(194)
Case Is < 6
.Text = .List(gt(194) - 1)
Case 6
.Text = .List(3)
Case 8
.Text = .List(4)
Case 15
.Text = .List(5)
Case 23
.Text = .List(6)
Case 58
.Text = .List(7)
Case 72
.Text = .List(8)

End Select

Case 12

errno = 0
.List(0) = 0

If gt(185) > 0 Then
  tmp = gt(185)
  If Stopnoise(2) Then
  gt(185) = 0
  labgenopts(54).Caption = "Midi Error"
  Cmbgenopts(12).ListIndex = 0
  Cmbgenopts(12).ToolTipText = "Handle for " & c1 & " did not close!"
  EnableNoise False
  Else
  ' Midichg changes gt(185)!
  For c1 = 1 To midiDevTot
    If Midichg(c1) > 0 Then
    .List(c1 - errno) = c1
      If c1 <> tmp Then
        If Stopnoise(2) Then Cmbgenopts(12).ToolTipText = "Handle for " & c1 & " did not close!"
      End If
    Else
      errno = errno + 1
    End If
  Next
  
  gt(185) = tmp

  If Midichg(tmp) = 0 Then
  labgenopts(54).Caption = "Midi Error"
  Cmbgenopts(12).ListIndex = 0
  End If
  End If
End If
.Text = .List(gt(185))

End Select
End With
End Sub
Private Function Colordlg(gtindex)
With Cdlg
.Flags = cdlCCRGBInit
.Color = gt(gtindex)
.CancelError = True
On Error GoTo NoActioncol
.ShowColor
Colordlg = .Color
Exit Function
NoActioncol:
If Err.Number <> 32755 Then
ShowError
Else
Colordlg = gt(gtindex)
End If
End With
End Function
Private Sub TTDefault()
    Select Case gt(158)
    Case 0
    Cmbgenopts(8).ToolTipText = "Saved settings as new game defaults"
    Case 1
    Cmbgenopts(8).ToolTipText = "Current settings as new game defaults"
    Case Else
    Cmbgenopts(8).ToolTipText = "Generic settings as new game defaults"
    End Select
End Sub
Private Sub resetdefaultbutton()
booRestoredefaults = False
boosavedefaults = False
Cmbgenopts(8).Enabled = True
Cmbgenopts(8).ToolTipText = ""
cmdgenopts(26).Caption = "Save &Defaults Now"
cmdgenopts(27).Caption = "&Restore Defaults"
cmdgenopts(26).ToolTipText = "Store current settings for new game profiles"
cmdgenopts(27).ToolTipText = "Restore old profile settings"
labgenopts(38).Visible = True
labgenopts(39).Visible = True
End Sub
Private Function FillList()
Dim Filter As Boolean
   
FillList = False

On Error GoTo fillerror

ReDim m_SysSnds(0) As SystemSoundDefinitions
    If SystemSoundNames(m_SysSnds) Then
    Filter = CBool(gt(183))
        With Cmbgenopts(13)
        .Text = ""
         .Clear
         For ct1 = LBound(m_SysSnds) To UBound(m_SysSnds)
            If Filter Then
               If m_SysSnds(ct1).Current <> "" Then
                  .AddItem m_SysSnds(ct1).SoundName
                  .ItemData(.NewIndex) = ct1
               End If
            Else
               .AddItem m_SysSnds(ct1).SoundName
               .ItemData(.NewIndex) = ct1
            End If
         Next ct1
         .ListIndex = 0
      End With
    End If

FillList = True
Exit Function

fillerror:
ShowError
End Function
Private Function namelt100chars(argu As String)
namelt100chars = True
If Len(argu) > 100 Then
response = MsgBox("The directory or file name is too long for this version (> 100 characters). Please use a directory with a shorter name or move to a lower directory.", vbOKOnly)
namelt100chars = False
End If
End Function
Private Sub EnableNoise(enablenow As Boolean, Optional dowave As Boolean = False)
If dowave = True Then
gametextvars(9).Enabled = enablenow
For ct = 0 To 1
Lstgenopts(ct).Enabled = enablenow
Next
Cmbgenopts(13).Enabled = enablenow
chkgenopt(4).Enabled = enablenow
For ct = 28 To 31
cmdgenopts(ct).Enabled = enablenow
Next

Else    'Midi
gametextvars(10).Enabled = enablenow
For ct = 2 To 3
Lstgenopts(ct).Enabled = enablenow
Next
chkgenopt(5).Enabled = enablenow
For ct = 32 To 34
cmdgenopts(ct).Enabled = enablenow
Next
End If
End Sub
Private Sub MCarlotime()
With labgenopts(52)
Select Case gt152
Case 0
.Caption = "(Let MyReels handle it)"
Case 50000
.Caption = "(A shortish wait)"
Case 100000
.Caption = "(A longer wait)"
Case 200000
.Caption = "(Out for a stretch)"
Case 500000
.Caption = "(Buff that +9 Garateen of Slay*)"
Case 1000000
.Caption = "(Time to sharpen that +18 Flamberge)"
Case 5000000
.Caption = "(Time to buff that +40 Steel Breastplate)"
Case 10000000
.Caption = "(Try that Winged Helm (7,11) of ""The Know"")"
Case 20000000
.Caption = "(Time to tip those 144 Silver Arrows(+3,+6))"
Case 50000000
.Caption = "(Let it go 'til that Morning Star doth shine)"
End Select
End With
End Sub
Private Function Getquotez(UseBaseDir As Boolean)
Genopts.MousePointer = vbHourglass
Getquotez = False
If UseBaseDir = True Then
If findafile(App.Path & "\", "Quotes.s$t") > 0 Then
gt(193) = 1
sDatabaseName = App.Path & "\" & "Quotes.s$t"
    If Openquotes("Dummy.txt", True) = True Then
    Stringvars(3) = App.Path & "\"
    cmdgenopts(35).Enabled = False
    Getquotez = True
    Else
    gt(193) = 0
    End If
End If
Else
If findafile(loaddirectory, "Quotes.s$t") > 0 Then
gt(193) = 0
sDatabaseName = loaddirectory & "Quotes.s$t"
    If Openquotes("Dummy.txt", True) = True Then
    Stringvars(3) = loaddirectory
    cmdgenopts(35).Enabled = True
    Getquotez = True
    Else
    gt(193) = 1
    End If
End If
End If

For ct = 1 To 1000000
Next
DoEvents
Genopts.MousePointer = vbDefault
End Function
Private Sub PCspeed()
Select Case gt(194)
Case 1
labgenopts(53).Caption = "'Slow++ PC'"
Case 2
labgenopts(53).Caption = "'Slow+ PC'"
Case 3
labgenopts(53).Caption = "'Slow PC'"
Case 5
labgenopts(53).Caption = "'Medium PC'"
Case 8
labgenopts(53).Caption = "'Fast PC'"
Case 15
labgenopts(53).Caption = "'Fast+ PC'"
Case 23
labgenopts(53).Caption = "'Fast++ PC'"
Case 58
labgenopts(53).Caption = "'Fast+++ PC'"
Case 72
labgenopts(53).Caption = "'Fast++++ PC'"
End Select
End Sub
