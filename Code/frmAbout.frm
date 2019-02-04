VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "MyReels: About this game"
   ClientHeight    =   8775
   ClientLeft      =   2835
   ClientTop       =   0
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8895
   Begin VB.CommandButton cmdfrmabout 
      Caption         =   "&Generate now"
      Height          =   345
      Index           =   0
      Left            =   1800
      TabIndex        =   130
      ToolTipText     =   "This can take some time!"
      Top             =   7320
      Width           =   1260
   End
   Begin VB.ComboBox Betsummary 
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      ItemData        =   "frmAbout.frx":000C
      Left            =   4560
      List            =   "frmAbout.frx":000E
      TabIndex        =   129
      Top             =   7680
      Width           =   4215
   End
   Begin VB.ComboBox Betsummary 
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      ItemData        =   "frmAbout.frx":0010
      Left            =   120
      List            =   "frmAbout.frx":0012
      TabIndex        =   128
      Top             =   7680
      Width           =   4215
   End
   Begin VB.ComboBox Betsummary 
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      ItemData        =   "frmAbout.frx":0014
      Left            =   840
      List            =   "frmAbout.frx":0016
      TabIndex        =   127
      Top             =   7080
      Width           =   855
   End
   Begin VB.Frame fraabout 
      Caption         =   " "
      Height          =   975
      Index           =   3
      Left            =   1200
      TabIndex        =   101
      Top             =   4680
      Width           =   7680
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   98
         Left            =   7200
         TabIndex        =   117
         Top             =   600
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   97
         Left            =   6480
         TabIndex        =   116
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   96
         Left            =   5640
         TabIndex        =   115
         Top             =   600
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   95
         Left            =   4200
         TabIndex        =   114
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   94
         Left            =   3480
         TabIndex        =   113
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   93
         Left            =   2280
         TabIndex        =   112
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   92
         Left            =   1515
         TabIndex        =   111
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   91
         Left            =   120
         TabIndex        =   110
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   90
         Left            =   7200
         TabIndex        =   109
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   89
         Left            =   6480
         TabIndex        =   108
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   88
         Left            =   5640
         TabIndex        =   107
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   87
         Left            =   4200
         TabIndex        =   106
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   86
         Left            =   3480
         TabIndex        =   105
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   85
         Left            =   2280
         TabIndex        =   104
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   84
         Left            =   1515
         TabIndex        =   103
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   83
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraabout 
      Caption         =   " "
      Height          =   1335
      Index           =   2
      Left            =   1200
      TabIndex        =   76
      Top             =   3020
      Width           =   7680
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   76
         Left            =   7200
         TabIndex        =   100
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   75
         Left            =   6480
         TabIndex        =   99
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   74
         Left            =   5640
         TabIndex        =   98
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   73
         Left            =   4200
         TabIndex        =   97
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   72
         Left            =   3480
         TabIndex        =   96
         Top             =   960
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   71
         Left            =   2280
         TabIndex        =   95
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   70
         Left            =   1515
         TabIndex        =   94
         Top             =   960
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   69
         Left            =   120
         TabIndex        =   93
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   68
         Left            =   7200
         TabIndex        =   92
         Top             =   600
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   67
         Left            =   6480
         TabIndex        =   91
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   66
         Left            =   5640
         TabIndex        =   90
         Top             =   600
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   65
         Left            =   4200
         TabIndex        =   89
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   64
         Left            =   3480
         TabIndex        =   88
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   63
         Left            =   2280
         TabIndex        =   87
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   62
         Left            =   1515
         TabIndex        =   86
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   61
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   60
         Left            =   7200
         TabIndex        =   84
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   59
         Left            =   6480
         TabIndex        =   83
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   58
         Left            =   5640
         TabIndex        =   82
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   57
         Left            =   4200
         TabIndex        =   81
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   56
         Left            =   3480
         TabIndex        =   80
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   55
         Left            =   2280
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   54
         Left            =   1515
         TabIndex        =   78
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   53
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdfrmabout 
      Caption         =   "&Hall_of_Fame"
      Height          =   345
      Index           =   5
      Left            =   3840
      TabIndex        =   43
      Top             =   360
      Width           =   1260
   End
   Begin VB.CommandButton cmdfrmabout 
      Caption         =   "&Reset Stats Now"
      Height          =   345
      Index           =   4
      Left            =   7320
      TabIndex        =   44
      Top             =   8250
      Width           =   1380
   End
   Begin VB.Frame fraabout 
      Height          =   1335
      Index           =   1
      Left            =   1200
      TabIndex        =   21
      Top             =   1400
      Width           =   7680
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   44
         Left            =   7200
         TabIndex        =   72
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   43
         Left            =   6480
         TabIndex        =   71
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   42
         Left            =   5640
         TabIndex        =   70
         Top             =   960
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   41
         Left            =   4200
         TabIndex        =   69
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   40
         Left            =   3480
         TabIndex        =   68
         Top             =   960
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   39
         Left            =   2280
         TabIndex        =   67
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   38
         Left            =   1515
         TabIndex        =   66
         Top             =   960
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   37
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   36
         Left            =   7200
         TabIndex        =   64
         Top             =   600
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   35
         Left            =   6480
         TabIndex        =   63
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   34
         Left            =   5640
         TabIndex        =   61
         Top             =   600
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   33
         Left            =   4200
         TabIndex        =   60
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   32
         Left            =   3480
         TabIndex        =   59
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   31
         Left            =   2280
         TabIndex        =   58
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   30
         Left            =   1515
         TabIndex        =   57
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   28
         Left            =   7200
         TabIndex        =   56
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   210
         Index           =   27
         Left            =   6480
         TabIndex        =   55
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   22
         Left            =   1515
         TabIndex        =   27
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "Turns ago"
         Height          =   255
         Index           =   23
         Left            =   2280
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   24
         Left            =   3480
         TabIndex        =   25
         Top             =   240
         Width           =   75
      End
      Begin VB.Label lblstats 
         Caption         =   "~  SVOG"
         Height          =   255
         Index           =   25
         Left            =   4200
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblstats 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   26
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblstats 
         Caption         =   "Largest prize"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdfrmabout 
      Caption         =   "&Sysinfo ..."
      Height          =   345
      Index           =   3
      Left            =   3240
      TabIndex        =   4
      Top             =   8250
      Width           =   1260
   End
   Begin VB.CommandButton cmdfrmabout 
      Caption         =   "&Device Info ..."
      Height          =   345
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   8250
      Width           =   1260
   End
   Begin VB.CommandButton cmdfrmabout 
      Cancel          =   -1  'True
      Caption         =   "&Back to Game"
      Default         =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   8250
      Width           =   1260
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   125
      Left            =   6600
      TabIndex        =   137
      Top             =   8250
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "Bet total"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   136
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblstats 
      Caption         =   "Possible bets"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   135
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblstats 
      Caption         =   "Money Back Strategies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   3120
      TabIndex        =   134
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label lblstats 
      Caption         =   "Optimal bets"
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   133
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblstats 
      Caption         =   "Value of game with money back option"
      Height          =   495
      Index           =   8
      Left            =   6240
      TabIndex        =   132
      ToolTipText     =   "Values with Bet Total showing an integral amount are optimal only for that Bet Total value"
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label lblstats 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Index           =   9
      Left            =   7800
      TabIndex        =   131
      Top             =   7080
      Width           =   570
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   210
      Index           =   124
      Left            =   7080
      TabIndex        =   126
      Top             =   480
      Width           =   45
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   210
      Index           =   123
      Left            =   7080
      TabIndex        =   125
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   80
      Left            =   7440
      TabIndex        =   1
      Top             =   4440
      Width           =   60
   End
   Begin VB.Label lblstats 
      AutoSize        =   -1  'True
      Height          =   210
      Index           =   79
      Left            =   6360
      TabIndex        =   124
      Top             =   4440
      Width           =   45
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   78
      Left            =   5400
      TabIndex        =   123
      Top             =   4440
      Width           =   60
   End
   Begin VB.Label lblstats 
      Caption         =   "Free Spin && Free Game no-win bonus pay ~  SVOG"
      Height          =   195
      Index           =   77
      Left            =   1200
      TabIndex        =   122
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   116
      Left            =   1560
      TabIndex        =   121
      Top             =   6650
      Width           =   60
   End
   Begin VB.Label lblstats 
      Caption         =   "Average bet size"
      Height          =   195
      Index           =   115
      Left            =   0
      TabIndex        =   120
      Top             =   6650
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Caption         =   "~  to true Value of Game"
      Height          =   195
      Index           =   121
      Left            =   6240
      TabIndex        =   119
      Top             =   6650
      Width           =   1695
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   122
      Left            =   8400
      TabIndex        =   118
      Top             =   6650
      Width           =   60
   End
   Begin VB.Label lblstats 
      Caption         =   "With Scatters"
      Height          =   255
      Index           =   50
      Left            =   120
      TabIndex        =   75
      Top             =   3260
      Width           =   975
   End
   Begin VB.Label lblstats 
      Caption         =   "With Substitutes"
      Height          =   255
      Index           =   51
      Left            =   0
      TabIndex        =   74
      Top             =   3620
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Caption         =   "With Naturals"
      Height          =   255
      Index           =   52
      Left            =   120
      TabIndex        =   73
      Top             =   3980
      Width           =   975
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   120
      Left            =   5760
      TabIndex        =   54
      Top             =   6650
      Width           =   60
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   210
      Index           =   119
      Left            =   4920
      TabIndex        =   53
      Top             =   6650
      Width           =   45
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   118
      Left            =   4200
      TabIndex        =   52
      Top             =   6650
      Width           =   60
   End
   Begin VB.Label lblstats 
      Caption         =   "Sampled Value of Game"
      Height          =   195
      Index           =   117
      Left            =   2160
      TabIndex        =   51
      Top             =   6650
      Width           =   1695
   End
   Begin VB.Label lblstats 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   114
      Left            =   7785
      TabIndex        =   50
      Top             =   6375
      Width           =   45
   End
   Begin VB.Label lblstats 
      Caption         =   "&& Bet Total"
      Height          =   195
      Index           =   113
      Left            =   6720
      TabIndex        =   49
      Top             =   6375
      Width           =   855
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   112
      Left            =   4680
      TabIndex        =   48
      Top             =   6375
      Width           =   60
   End
   Begin VB.Label lblstats 
      Caption         =   "Money Back"
      Height          =   195
      Index           =   82
      Left            =   120
      TabIndex        =   47
      Top             =   5300
      Width           =   975
   End
   Begin VB.Label lblstats 
      Caption         =   "Random Jackpot"
      Height          =   195
      Index           =   81
      Left            =   0
      TabIndex        =   46
      Top             =   4940
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Caption         =   "Strategy in last Money Back Win"
      Height          =   195
      Index           =   111
      Left            =   0
      TabIndex        =   45
      Top             =   6375
      Width           =   2415
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   110
      Left            =   8160
      TabIndex        =   42
      Top             =   6045
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "Average Lines played"
      Height          =   195
      Index           =   109
      Left            =   6120
      TabIndex        =   41
      Top             =   6045
      Width           =   1550
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   108
      Left            =   5280
      TabIndex        =   40
      Top             =   6045
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "No of Restarts with Different Random Seed"
      Height          =   375
      Index           =   107
      Left            =   3000
      TabIndex        =   39
      Top             =   5940
      Width           =   1695
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   106
      Left            =   2160
      TabIndex        =   38
      Top             =   6045
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "No of Restarts with Same Random Seed"
      Height          =   375
      Index           =   105
      Left            =   120
      TabIndex        =   37
      Top             =   5940
      Width           =   1455
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   104
      Left            =   8130
      TabIndex        =   36
      Top             =   5700
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "No of Fast Spins"
      Height          =   195
      Index           =   103
      Left            =   6240
      TabIndex        =   35
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   102
      Left            =   5400
      TabIndex        =   34
      Top             =   5700
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "No of Spins of Short Duration"
      Height          =   195
      Index           =   101
      Left            =   2760
      TabIndex        =   33
      Top             =   5700
      Width           =   2150
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   100
      Left            =   1920
      TabIndex        =   32
      Top             =   5700
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   " No of Spins Up"
      Height          =   195
      Index           =   99
      Left            =   120
      TabIndex        =   31
      Top             =   5700
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   210
      Index           =   46
      Left            =   3720
      TabIndex        =   30
      Top             =   2800
      Width           =   45
   End
   Begin VB.Label lblstats 
      Caption         =   "Free Games and"
      Height          =   255
      Index           =   47
      Left            =   4320
      TabIndex        =   29
      Top             =   2800
      Width           =   1095
   End
   Begin VB.Label lblstats 
      Caption         =   "During the"
      Height          =   195
      Index           =   45
      Left            =   2400
      TabIndex        =   28
      Top             =   2800
      Width           =   735
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   210
      Index           =   16
      Left            =   4200
      TabIndex        =   20
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label lblstats 
      Caption         =   "Free Spins"
      Height          =   255
      Index           =   49
      Left            =   6600
      TabIndex        =   19
      Top             =   2800
      Width           =   735
   End
   Begin VB.Label lblstats 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   210
      Index           =   48
      Left            =   6000
      TabIndex        =   18
      Top             =   2800
      Width           =   45
   End
   Begin VB.Label lblstats 
      Caption         =   "Non - Free Games and Spins"
      Height          =   255
      Index           =   17
      Left            =   4800
      TabIndex        =   17
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblstats 
      Caption         =   "During the"
      Height          =   195
      Index           =   15
      Left            =   2880
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblstats 
      Caption         =   "With Naturals"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   15
      Top             =   2360
      Width           =   975
   End
   Begin VB.Label lblstats 
      Caption         =   "With Substitutes"
      Height          =   255
      Index           =   19
      Left            =   0
      TabIndex        =   14
      Top             =   2000
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Caption         =   "With Scatters"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   13
      Top             =   1640
      Width           =   975
   End
   Begin VB.Label lblstats 
      Caption         =   "Chance of no prize"
      Height          =   255
      Index           =   13
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblstats 
      Caption         =   "Value of game"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblstats 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Index           =   11
      Left            =   7800
      TabIndex        =   10
      Top             =   960
      Width           =   570
   End
   Begin VB.Label lblstats 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Index           =   10
      Left            =   1440
      TabIndex        =   9
      Top             =   960
      Width           =   570
   End
   Begin VB.Label lblstats 
      Caption         =   "This program aims to promote the more relaxing and entertaining features of Slot Machine games. "
      Height          =   435
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label lblstats 
      Height          =   195
      Index           =   6
      Left            =   1440
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblstats 
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblstats 
      Caption         =   "Statistics "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Index           =   4
      Left            =   3840
      TabIndex        =   5
      Top             =   855
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   0
      X2              =   8880
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   8880
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   8880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   8880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblstats 
      Caption         =   "Stats Resets"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   5040
      TabIndex        =   2
      Top             =   8250
      Width           =   855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type


Private Type TIME_ZONE_INFORMATION
Bias As Long
StandardName(0 To 63) As Byte
StandardDate As SYSTEMTIME
StandardBias As Long
DaylightName(0 To 63) As Byte
DaylightDate As SYSTEMTIME
DaylightBias As Long
End Type

Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Const TIME_ZONE_ID_UNKNOWN = 0
Const TIME_ZONE_ID_STANDARD = 1
Const TIME_ZONE_ID_DAYLIGHT = 2


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Dim multitracker(200) As Long, bmax(5) As Long, tooltxt(2) As String, VOG As Single, moneybackVOG As Single, p0 As Single, totalnospinz As Long
Dim temp As Long, ct As Long, ct1 As Long, maxct As Long, betconcat As String, response As Boolean
Private Sub Form_Activate()
Dim f As Form
For Each f In Forms
If f.Name = "Zhidden" Then Zhidden.Show
Next
End Sub
Private Sub Form_Load()
Dim thumbs(14) As StdPicture, wheelorder(4, 24) As Long, wheelvec(5, 14) As Long, piccount As Long
Dim nRet As Long, tz As TIME_ZONE_INFORMATION


Unload Zhidden
Set Zhidden = Nothing

setformpos Me

lblstats(123).Caption = Date & " " & Time

nRet = GetTimeZoneInformation(tz)

With lblstats(124)
If nRet <> TIME_ZONE_ID_INVALID Then
Select Case nRet
Case TIME_ZONE_ID_UNKNOWN
.Caption = "Time Zone Unknown! "
Case TIME_ZONE_ID_STANDARD
.Caption = "Standard Time... "
Case TIME_ZONE_ID_DAYLIGHT
.Caption = "Daylight Savings Time... "
End Select

If tz.Bias > 0 Then
.Caption = .Caption & tz.Bias / 60 & " hrs behind UTC."
Else
.Caption = .Caption & -tz.Bias / 60 & " hrs ahead of UTC."
End If
End If
End With



If gt(10) = 0 Then cmdfrmabout(0).Enabled = False


getthumbspiccount thumbs, piccount, wheelvec, wheelorder

tooltxt(0) = "spins down"
tooltxt(1) = "long spins"
tooltxt(2) = "slow spins"


moneybackVOG = 0
VOG = 0
maxct = 0

'value of game etc
If gt(29) > 0 Then
VOG = CSng(inputformat(29))
lblstats(10).Caption = CStr(VOG)
p0 = CSng(inputformat(33))

If p0 < 1 Then
lblstats(11).Caption = "0.9999"
Else
lblstats(11).Caption = CStr(Format(p0, "00.000"))
End If

If gt(31) > 0 Then
lblstats(9).Caption = inputformat(31)
moneybackVOG = CSng(lblstats(9).Caption)
Else
lblstats(9).Visible = False
End If
Else
For ct = 9 To 11
lblstats(ct).Visible = False
Next
End If

writetolabels False

If gt(0) = -2 Then
cmdfrmabout(4).Enabled = False
cmdfrmabout(1).Caption = "Quit Game"
End If

Me.Caption = "About " & App.Title
lblstats(6).Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblstats(5).Caption = App.Title

Me.Show
Exit Sub

End Sub
Private Sub cmdfrmabout_Click(Index As Integer)
procend = False

Select Case Index
Case 0

cmdfrmabout(0).Caption = "Processing ...."

getBmax bmax

'Use truncated VOG & P0 if nothing better
If VOG = 0 Then VOG = Val(gt(29) & decsep & gt(30))
If p0 = 0 Then p0 = Val(gt(33) & decsep & gt(34))


Betsummary(2).AddItem "ANY"
For ct = gt(10) - 1 + bmax(5) To bmax(5) * gt(10)
Betsummary(2).AddItem ct    'initialise betsummary
Next
Betsummary(2).Text = Betsummary(2).List(0)


If comboinit(gt(10) - 1 + bmax(5), bmax(5) * gt(10)) = True Then
For ct = 0 To 2
Betsummary(ct).Enabled = True
Next
Betsummary(2).ToolTipText = "Select a bet total. Reselecting ""ANY"" can take some time to process!"
Betsummary(0).ToolTipText = "View all possible combinations using bet total, above"
Betsummary(1).ToolTipText = "View optimal strategy(s) using bet total, above"
cmdfrmabout(0).Caption = "&Generate now"
Else
cmdfrmabout(0).Caption = "Unavailable"
End If
cmdfrmabout(0).Enabled = False
Case 1
Unload frmAbout
Set frmAbout = Nothing
If gt(0) > -2 Then
Pokemach.Show
Pokemach.Enabled = True
End If
Case 2
zhiddnstatus = 1
Load Zhidden
Case 3
Call StartSysInfo
Case 4  'reset stat

If cmdfrmabout(4).Caption = "Done" Then Exit Sub

If cmdfrmabout(4).Caption = "Sure?" Then
writetolabels True
Else
cmdfrmabout(4).Caption = "Sure?"
End If
Case 5  'Hall of fame
zhiddnstatus = 2
Load Zhidden
End Select
procend = True
End Sub
Private Sub Betsummary_Change(Index As Integer)
If procend = True Then
Betsummary(Index).Text = Betsummary(Index).List(0)
Betsummary(Index).Refresh
End If
End Sub
Private Sub StartSysInfo()
Dim errcondition As String, SysInfoPath As String, c As New cRegistry
    
    
errcondition = ""
    
    
On Error GoTo SysInfoErr
    
'Try To Get System Info Program Path\Name From Registry...
With c
.ClassKey = HKEY_LOCAL_MACHINE
.SectionKey = gREGKEYSYSINFO
.ValueKey = gREGVALSYSINFO
SysInfoPath = .Value
End With

    If SysInfoPath = "" Then
    'Try To Get System Info Program Path Only From Registry...
    With c
    .ClassKey = HKEY_LOCAL_MACHINE
    .SectionKey = gREGKEYSYSINFOLOC
    .ValueKey = gREGVALSYSINFOLOC
    SysInfoPath = .Value
    End With
    
    If SysInfoPath = "" Then
    errcondition = "(Registry access problem)"
    GoTo SysInfoErr
    ElseIf (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
    SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    Else
    errcondition = "(missing MSINFO32.EXE)"
    GoTo SysInfoErr
    End If
        
End If

If c.EightorLater Then
If Not c.ShellOut(SysInfoPath) Then GoTo SysInfoErr
Else
Call Shell(SysInfoPath, vbNormalFocus)
End If

Exit Sub
SysInfoErr:
MsgBox "System Information Is Unavailable At This Time " & errcondition & ".", vbOKOnly
End Sub
Private Sub Betsummary_Click(Index As Integer)
If Index = 2 Then
'reinit vars
getBmax bmax
'comboinit "useless" as function here
If Betsummary(2).Text = "ANY" Then
If comboinit(gt(10) - 1 + bmax(5), bmax(5) * gt(10)) Then Exit Sub
Else
If comboinit(Val(Betsummary(2).Text), Val(Betsummary(2).Text)) Then Exit Sub
End If
End If
End Sub
Private Function comboinit(loopstart As Long, loopfin As Long)
Dim Btotal As Long, grandcombototal As Long, combosinbtotal As Long
Dim Y As Single, ymax As Single, gt31 As Long, gt32 As Long
Dim bettest(126, 15) As Long

procend = False

Screen.MousePointer = vbHourglass

ymax = 0
Y = 0
grandcombototal = 0

For ct = 1 To 200
multitracker(ct) = 0
Next

For ct = 1 To 15
For ct1 = 0 To 126
bettest(ct1, ct) = 0
Next
Next

'clean up combos

For ct = 0 To 1
Betsummary(ct).Clear
Next


'Moneyback :More than 1 bet choice amount allowed
    If VOG < 100 And VOG > 0 And gt(10) > 1 And bmax(4) > 0 Then
    'if VOG > 1, always bet MAX


    'Find Btotal here, use temp vars Xmn, Xrst for convenience

    For Btotal = loopstart To loopfin
    
        
        'Did we ever get to calculate bets?
        If moneybackvalue(VOG, p0, Y, bettest, Btotal) = True Then

        combosinbtotal = 0

        'Optimal combo for btotal store in bettest(0, )
        For ct = 0 To 126
        betconcat = ""
        If bettest(ct, 1) = 0 Then Exit For     '",1" will do
        For ct1 = 1 To gt(10) - 1
        betconcat = betconcat & CStr(bettest(ct, ct1)) & " ,"
        Next
        betconcat = betconcat & CStr(bmax(5))
            If ct = 0 Then
            Betsummary(1).Clear 'clear old entries
            Betsummary(1).AddItem betconcat
            Else
            combosinbtotal = combosinbtotal + 1
            Betsummary(0).AddItem betconcat
                If betconcat = Betsummary(1).List(0) Then
                    If Y > ymax Then
                    multitracker(1) = grandcombototal + combosinbtotal
                    For ct1 = 2 To 200
                    multitracker(ct1) = 0
                    Next
                    ymax = Y
                    If loopstart <> loopfin Then maxct = Btotal
                    ElseIf Y = ymax Then
                    For ct1 = 2 To 200
                    If multitracker(ct1) = 0 Then
                    multitracker(ct1) = grandcombototal + combosinbtotal
                    Exit For
                    End If
                    Next
                    End If


                End If
            End If
        Next
            If loopstart = loopfin Then
            outputforformat calcmonbackvog(ymax), 198, 197
                If loopstart = maxct Then
                lblstats(9).Caption = inputformat(31)
                Else    'avoid annoying roundoff
                lblstats(9).Caption = inputformat(198)
                End If
            ElseIf Btotal = loopfin Then
            lblstats(9).Caption = inputformat(31)
            End If
        
        
        grandcombototal = grandcombototal + combosinbtotal

        Else    'bets not calculated
        
            If loopstart = loopfin Then 'not "any"
       
            'load error message in combos
            'clear everything else
            For ct = 0 To 1
            Betsummary(ct).Clear
            Next
            Betsummary(0).Text = "Not possible"
            lblstats(9) = "      "
            comboinit = True
            procend = True
            Screen.MousePointer = vbDefault
            Exit Function
            End If

        End If

    Next    'btotal loop
    


   If Betsummary(2).Text = "ANY" Then
    Betsummary(1).Clear 'clear temp entries
    'Always add first item of multitracker
    Betsummary(1).AddItem CStr(Betsummary(0).List(multitracker(1) - 1))
    For ct = 2 To 200
        If multitracker(ct) > 0 Then
        Betsummary(1).AddItem CStr(Betsummary(0).List(multitracker(ct) - 1))
        Else
        Exit For
        End If
    Next
    End If

    'need to select first item
    For ct = 0 To 1
    Betsummary(ct).Text = Betsummary(ct).List(0)
    Next
    cmdfrmabout(1).SetFocus
    For ct = 0 To 1
    Betsummary(ct).Refresh
    Next
    
    comboinit = True
    Else
    comboinit = False
    End If

Screen.MousePointer = vbDefault
procend = True
End Function
Private Sub Betsummary_GotFocus(Index As Integer)
'This example automatically drops down the list portion of a ComboBox control
'whenever the ComboBox receives the focus. To try this example, create a new form
'containing a ComboBox control and an OptionButton control (used only to receive
'the focus). Create a new module using the Add Module command on the Project menu.
'Paste the Declare statement into the Declarations section of the new module,
'being sure that the statement is on one line with no break or wordwrap.
'Then paste the Sub procedure into the Declarations section of the form, and press
'F5. Use the TAB key to move the focus to and from the ComboBox.

    Const CB_SHOWDROPDOWN = &H14F
    Dim tmp
    tmp = SendMessage(Betsummary(Index).hWnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub padlabel(lblno As Long, padval As Long)
Dim looper As Long, lbltxt As String
With lblstats(lblno)
If padval = 0 Then
.Visible = False
Exit Sub
Else
.Visible = True
lbltxt = ""
For looper = Len(CStr(padval)) To Len(CStr(gt(47))) - 1
lbltxt = "0" & lbltxt
Next
.Caption = lbltxt & CStr(padval)
End If
Select Case lblno
Case 100, 102, 104
lbltxt = ""
For looper = Len(CStr(totalnospinz - padval)) To Len(CStr(gt(47))) - 1
lbltxt = "0" & lbltxt
Next
.ToolTipText = lbltxt & CStr(totalnospinz - padval) & Space(1) & tooltxt((lblno - 100) / 2)
End Select
End With
End Sub
Private Sub upordown(newstats As Single, oldstats As Single, Index As Integer)
Dim improvfactor As Single
'This routine formats & extracts changes in stats as a percentage
Select Case oldstats
Case 0.9999
If newstats <= 0.9999 Then
improvfactor = 1
Else
improvfactor = newstats / oldstats
End If
Case 999.99
If newstats >= 999.99 Then
improvfactor = 1
Else
improvfactor = newstats / oldstats
End If
Case Else
improvfactor = newstats / oldstats
End Select
If improvfactor >= 1.001 Then
lblstats(Index).Caption = "Up by "
lblstats(Index + 1).ForeColor = &HFFFF00
outputforformat 100 * (improvfactor - 1), 198, 197
lblstats(Index + 1).Caption = inputformat(198)
ElseIf improvfactor <= 0.999 Then
lblstats(Index) = "Down by"
lblstats(Index + 1).ForeColor = &H80FF&
outputforformat 100 * (1 - improvfactor), 198, 197
lblstats(Index + 1).Caption = inputformat(198)
Else
lblstats(Index).Caption = "No change"
lblstats(Index + 1).Visible = False
End If
lblstats(Index).ToolTipText = "% change from last time this form was clicked"
End Sub
Private Sub writetolabels(suretoclearstats As Boolean)
Dim overallperf As Single, sparevar1 As Single, sparevar2 As Single, Proportinnew(8) As Single, SVOGnew(8) As Single
Dim omitimprov As Boolean, oldprizetotal As Long
On Error GoTo Statserr
totalnospinz = gt(49) + gt(50) + gt(51)

If suretoclearstats = False Then

If totalnospinz = 0 Then cmdfrmabout(4).Enabled = False

    If gt(133) = gt(135) Then
    omitimprov = True
    Else
    omitimprov = False
    'save prize total here. gt(136) is now current
    oldprizetotal = gt(136)
    gt(136) = gt(59) + gt(63) + gt(67) + gt(71) + gt(75) + gt(79) + gt(83) + gt(87)
    End If
Else
'Resetting Stats
omitimprov = True
For ct = 48 To 149
gt(ct) = 0
Next

gt(150) = gt(150) + 1
totalnospinz = 0
cmdfrmabout(4).Caption = "Done"


End If

overallperf = 0

padlabel 16, gt(51)
padlabel 46, gt(49)
padlabel 48, gt(50)
padlabel 100, gt(52)
padlabel 102, gt(53)
padlabel 104, gt(54)
padlabel 106, gt(55)
padlabel 108, gt(56)
padlabel 125, gt(150)



'Calculate SVOGnew
For ct1 = 0 To 3 Step 3
For ct = 0 To 2
If gt(60 + 4 * (ct + ct1)) > 0 Then
'multiply proportion of prize winning spins with average prize win
If gt(122 + ct + ct1) > 0 Then
SVOGnew(ct + ct1) = 100 * (gt(59 + 4 * (ct + ct1)) / gt(136))   'extra brackets for overflow
Proportinnew(ct + ct1) = gt(60 + 4 * (ct + ct1)) / gt(122 + ct + ct1) 'effective bet
End If
End If
Next
Next


'Random jackpot & Money back
For ct = 0 To 1
If gt(84 + 4 * ct) > 0 And (gt(83 + 4 * ct)) > 0 Then
'compare ratios of prize types to sum of prize types
If gt(128 + ct) > 0 Then
SVOGnew(ct + 6) = 100 * (gt(83 + 4 * ct) / gt(136)) 'gt(128) always > 0 for RJ
If ct = 0 Then
Proportinnew(ct + 6) = (gt(84 + 4 * ct) + gt(48)) / gt(128 + ct) 'effective bet
Else
Proportinnew(ct + 6) = gt(84 + 4 * ct) / gt(128 + ct) 'effective bet
End If
End If
End If
Next

'FSFG no prize bonus
If gt(131) > 0 And gt(132) > 0 Then
SVOGnew(8) = 100 * (gt(131) / gt(136))
Proportinnew(8) = gt(132) / gt(130) 'effective bet
End If

'Print values
For ct1 = 22 To 54 Step 32

If ct1 = 22 Then
temp = 0
Else
temp = 3
End If

For ct = 0 To 2

If gt(60 + 4 * (ct + temp)) > 0 Then

sparevar2 = 0

padlabel ct1 + 8 * ct, gt(57 + 4 * (ct + temp))
padlabel ct1 + 2 + 8 * ct, totalnospinz - gt(58 + 4 * (ct + temp))


If SVOGnew(ct + temp) > 0 Then
outputforformat SVOGnew(ct + temp), 198, 197
lblstats(ct1 + 4 + 8 * ct).Caption = inputformat(198)
lblstats(ct1 + 4 + 8 * ct).ToolTipText = "Number of prizes " & gt(122 + ct + temp) & ". Effective bet : " & Format(Proportinnew(ct + temp), "0.0000")
lblstats(ct1 + 3 + 8 * ct).ToolTipText = "How much this prize type contributes to the Sampled Value Of Game below"
sparevar2 = Val(inputformat(104 + 2 * (ct + temp)))

If omitimprov = False Then

'Save VOGS for next FRMABOUT
gt(104 + 2 * (ct + temp)) = gt(198)
gt(105 + 2 * (ct + temp)) = gt(199)

If sparevar2 > 0 Then
upordown SVOGnew(ct + temp), sparevar2, ct1 + 5 + 8 * ct
Else
lblstats(ct1 + 5 + 8 * ct).Visible = False
lblstats(ct1 + 6 + 8 * ct).Visible = False
End If

Else
lblstats(ct1 + 5 + 8 * ct).Visible = False
lblstats(ct1 + 6 + 8 * ct).Visible = False
End If

Else
lblstats(ct1 + 4 + 8 * ct).Visible = False
lblstats(ct1 + 5 + 8 * ct).Visible = False
lblstats(ct1 + 6 + 8 * ct).Visible = False
End If  'sparevar1


Else
lblstats(ct1 + 8 * ct).Visible = False
lblstats(ct1 + 2 + 8 * ct).Visible = False
lblstats(ct1 + 4 + 8 * ct).Visible = False
lblstats(ct1 + 5 + 8 * ct).Visible = False
lblstats(ct1 + 6 + 8 * ct).Visible = False
lblstats(ct1 + 3 + 8 * ct).ToolTipText = ""
lblstats(ct1 + 5 + 8 * ct).ToolTipText = ""
End If
Next
Next


'Random jackpot & Money back

For ct = 0 To 1

If gt(84 + 4 * ct) > 0 And (gt(83 + 4 * ct)) > 0 Then

sparevar2 = 0

padlabel 84 + 8 * ct, gt(81 + 4 * ct)
padlabel 86 + 8 * ct, totalnospinz - gt(82 + 4 * ct)

If SVOGnew(ct + 6) > 0 Then
outputforformat SVOGnew(ct + 6), 198, 197
lblstats(88 + 8 * ct).Caption = inputformat(198)
lblstats(88 + 8 * ct).ToolTipText = "While feature active - Number of spins " & gt(128 + ct) & ". Effective bet " & Format(Proportinnew(ct + 6), "0.0000")
lblstats(87 + 8 * ct).ToolTipText = "How much this prize type contributes to the Sampled Value Of Game below"
sparevar2 = Val(inputformat(116 + 2 * ct))

If omitimprov = False Then

'Save VOGS for next FRMABOUT
gt(116 + 2 * ct) = gt(198)
gt(117 + 2 * ct) = gt(199)

If sparevar2 > 0 Then
upordown SVOGnew(ct + 6), sparevar2, 89 + 8 * ct
Else
lblstats(89 + 8 * ct).Visible = False
lblstats(90 + 8 * ct).Visible = False
End If

Else
lblstats(89 + 8 * ct).Visible = False
lblstats(90 + 8 * ct).Visible = False
End If


Else
lblstats(88 + 8 * ct).Visible = False
lblstats(89 + 8 * ct).Visible = False
lblstats(90 + 8 * ct).Visible = False
End If  'sparevar1


Else
lblstats(84 + 8 * ct).Visible = False
lblstats(86 + 8 * ct).Visible = False
lblstats(88 + 8 * ct).Visible = False
lblstats(89 + 8 * ct).Visible = False
lblstats(90 + 8 * ct).Visible = False
lblstats(87 + 8 * ct).ToolTipText = ""
lblstats(89 + 8 * ct).ToolTipText = ""
End If
Next


'FSFG no prize bonus
If gt(131) > 0 And gt(132) > 0 Then

sparevar2 = 0

If SVOGnew(8) > 0 Then
outputforformat SVOGnew(8), 198, 197


lblstats(78).Caption = inputformat(198)
lblstats(78).ToolTipText = "Number of spins while feature active - " & gt(130) & ". Effective Bet Total " & Format(Proportinnew(8), "0.0000")
lblstats(77).ToolTipText = "How much this prize type contributes to the Sampled Value Of Game below"
sparevar2 = Val(inputformat(120))

If omitimprov = False Then

'Save VOGS for next FRMABOUT
gt(120) = gt(198)
gt(121) = gt(199)

If sparevar2 > 0 Then
upordown SVOGnew(8), sparevar2, 79
Else
lblstats(79).Visible = False
lblstats(80).Visible = False
End If
Else
lblstats(79).Visible = False
lblstats(80).Visible = False
End If



Else
lblstats(78).Visible = False
lblstats(79).Visible = False
lblstats(80).Visible = False
End If
Else
lblstats(78).Visible = False
lblstats(79).Visible = False
lblstats(80).Visible = False
End If


If totalnospinz > 0 Then    'in case we are resettng

'MultiLine average
sparevar1 = gt(149) / totalnospinz
outputforformat sparevar1, 198, 197
lblstats(110) = inputformat(198)



'total bet quantity/ totalspinz
sparevar1 = gt(133) / totalnospinz
outputforformat sparevar1, 198, 197, True
lblstats(116) = inputformat(198)


'Total prizemoney when Frmabout last invoked / total bet
overallperf = gt(136) / gt(133)

Else
lblstats(110).Visible = False
lblstats(116).Visible = False
End If

If overallperf > 0 Then

outputforformat 100 * overallperf, 198, 197
lblstats(118) = inputformat(198)


'have an old bet total to compare
If gt(135) > 0 And oldprizetotal > 0 Then
upordown overallperf, oldprizetotal / gt(135), 119
Else
lblstats(119).Visible = False
lblstats(120).Visible = False
End If


With lblstats(122)
If moneybackVOG > 0 Then
outputforformat 10000 * overallperf / moneybackVOG, 198, 197
.Caption = inputformat(198)
.ToolTipText = "Optimised Money back VOG used for comparison"
ElseIf VOG > 0 Then
outputforformat 10000 * overallperf / VOG, 198, 197
.Caption = inputformat(198)
Else
lblstats(121).Visible = False
.Visible = False
End If
End With

Else
lblstats(118).Visible = False
lblstats(119).Visible = False
lblstats(120).Visible = False
lblstats(121).Visible = False
lblstats(122).Visible = False
End If



If gt(87) > 0 And gt(89) > 0 Then
betconcat = ""
ct1 = 0

For ct = 89 To 88 + gt(151)
If gt(ct) > 30 Then
    If Mid(CStr(gt(ct)), 3, 1) = "0" Then
    temp = CLng(Left(CStr(gt(ct)), 2))
    betconcat = betconcat & CStr(temp)
    ElseIf Mid(CStr(gt(ct)), 2, 1) = "0" Then
    temp = CLng(Left(CStr(gt(ct)), 1))
    betconcat = betconcat & CStr(temp)
    End If
Else
temp = gt(ct)
betconcat = betconcat & CStr(temp)
End If
If ct < 88 + gt(151) Then betconcat = betconcat & " ,"
ct1 = temp + ct1
Next


lblstats(112).Caption = betconcat
lblstats(114) = CStr(ct1)

Else
lblstats(112).Visible = False
lblstats(114).Visible = False
End If

If omitimprov = False Then gt(135) = gt(133)    'save old bet total between frmabout invocations

Exit Sub

Statserr:

response = MsgBox("There is an unexpected error with Statistics. Some values may not appear.", vbOKOnly)

End Sub

