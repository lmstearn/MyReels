VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form cornfig 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MyReels: Configure Reels"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   6
   Icon            =   "cornfig.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   12060
   Begin VB.PictureBox picspare 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   10800
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Chkordered 
      Caption         =   "&Ordered"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   83
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 5"
      Height          =   495
      Index           =   9
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 4"
      Height          =   495
      Index           =   8
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 3"
      Height          =   495
      Index           =   7
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 2"
      Height          =   495
      Index           =   6
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 1"
      Height          =   495
      Index           =   5
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   5880
      TabIndex        =   77
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 5"
      Height          =   495
      Index           =   4
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 4"
      Height          =   495
      Index           =   3
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 3"
      Height          =   495
      Index           =   2
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 2"
      Height          =   495
      Index           =   1
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Regen 
      Caption         =   "Reel 1"
      Height          =   495
      Index           =   0
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Continue 
      Caption         =   "&Generate"
      Default         =   -1  'True
      Height          =   285
      Left            =   5880
      TabIndex        =   71
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Regenerate 
      Caption         =   "&Regenerate All"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   70
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   13
      Left            =   6000
      Top             =   7320
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   12
      Left            =   6000
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   11
      Left            =   6000
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   10
      Left            =   6000
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   9
      Left            =   6000
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   8
      Left            =   6000
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   7
      Left            =   6000
      Top             =   840
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   6
      Left            =   120
      Top             =   7320
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   5
      Left            =   120
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   4
      Left            =   120
      Top             =   5160
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   3
      Left            =   120
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   2
      Left            =   120
      Top             =   3000
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   1
      Left            =   120
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image imgsel 
      Appearance      =   0  'Flat
      Height          =   975
      Index           =   0
      Left            =   120
      Top             =   840
      Width           =   975
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   13
      Left            =   7080
      TabIndex        =   97
      Top             =   8160
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   12
      Left            =   7080
      TabIndex        =   96
      Top             =   7080
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   11
      Left            =   7080
      TabIndex        =   95
      Top             =   6000
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   10
      Left            =   7080
      TabIndex        =   94
      Top             =   4920
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   9
      Left            =   7080
      TabIndex        =   93
      Top             =   3840
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   8
      Left            =   7080
      TabIndex        =   92
      Top             =   2760
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   7
      Left            =   7080
      TabIndex        =   91
      Top             =   1680
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   6
      Left            =   1200
      TabIndex        =   90
      Top             =   8160
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   5
      Left            =   1200
      TabIndex        =   89
      Top             =   7080
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   4
      Left            =   1200
      TabIndex        =   88
      Top             =   6000
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   3
      Left            =   1200
      TabIndex        =   87
      Top             =   4920
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   2
      Left            =   1200
      TabIndex        =   86
      Top             =   3840
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   1
      Left            =   1200
      TabIndex        =   85
      Top             =   2760
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label labscatter 
      Caption         =   "Scatter"
      Height          =   210
      Index           =   0
      Left            =   1200
      TabIndex        =   84
      Top             =   1680
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label Label15 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   69
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label15 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   68
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label15 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   67
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label15 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   66
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   65
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   64
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   63
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   62
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label13 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   61
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label13 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   60
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label13 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   59
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label13 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   58
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label12 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   57
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label12 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   56
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label12 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   55
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label12 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   54
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   53
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   52
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   51
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   50
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   49
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label10 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   48
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label10 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   47
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label10 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   46
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label9 
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   45
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label9 
      Height          =   495
      Index           =   3
      Left            =   10080
      TabIndex        =   44
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label9 
      Height          =   495
      Index           =   2
      Left            =   9240
      TabIndex        =   43
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label9 
      Height          =   495
      Index           =   1
      Left            =   8400
      TabIndex        =   42
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label8 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   41
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label8 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   40
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label8 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   39
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label8 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   38
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label7 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   37
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label7 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   36
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label7 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   35
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label7 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   34
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   33
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   32
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   31
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   30
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label5 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   29
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label5 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   28
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label5 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   27
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label5 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   25
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   24
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   23
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   22
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   21
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   20
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   19
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   18
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label15 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   17
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label14 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   16
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label13 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   15
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label12 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   14
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label11 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   13
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   12
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label9 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label8 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   10
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label Label7 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   9
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label6 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   8
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label5 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label3 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Index           =   4
      Left            =   5040
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   375
   End
   Begin ComctlLib.ImageList Thumblist 
      Left            =   11400
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "cornfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thumbs(14) As StdPicture, poscolor(4) As Long, wheelvec(5, 14) As Long, sstscat(14) As Long
Dim wheelorder(4, 24) As Long, intnotzero(13) As Long, scatterdonehere(13) As Boolean
Dim piccount As Long, tempplusone As Long, tempminusone As Long
Dim intreel As Long, poscolortemp As Long, indexplus5 As Long, ct As Long, temp As Long
Dim intlastindex As Long
'sstscat(pct)  frequency of scatters per reel; 0 if not scatter
Private Sub Form_Load()
Dim newstr As String
procend = False

If resX = 1 Then
setformpos Me
Else
Dotaskwindow Me, True

With Me
.Width = resX * .Width
.Height = resY * .Height
End With

setformpos Me

With Cancel
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With
With Continue
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With
With Chkordered
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With
With Regenerate
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
End With
With picspare
.Width = resX * .Width
.Height = resY * .Height
End With

End If


For ct = 0 To 9
With Regen(ct)
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
.fontname = "Times New Roman"
.Fontsize = 9.75 * resY
.Font.Charset = 0
.Font.Weight = 700
.FontUnderline = 0           'False
.FontItalic = 0              'False
.FontStrikethru = 0       'False
End With
Next

For intreel = 0 To 4
poscolor(intreel) = 0
Next
getthumbspiccount thumbs, piccount, wheelvec, wheelorder

For pct = 1 To piccount

Thumblist.ListImages.Add (pct), , thumbs(pct)

picspare.PaintPicture Thumblist.ListImages(pct).Picture, 0, 0, 975 * resX, 975 * resY
With imgsel(pct - 1)
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height
imgsel(pct - 1).BorderStyle = 0
imgsel(pct - 1).Picture = picspare.Image
End With
Set picspare = Nothing


For ct = 0 To 4
Select Case pct
Case 1
dofont Label2(ct)
Case 2
dofont Label3(ct)
Case 3
dofont Label4(ct), Label5(ct)
Case 4
dofont Label5(ct), Label6(ct)
Case 5
dofont Label6(ct), Label7(ct)
Case 6
dofont Label7(ct), Label8(ct)
Case 7
dofont Label8(ct), Label9(ct)
Case 8
dofont Label9(ct), Label10(ct)
Case 9
dofont Label10(ct), Label11(ct)
Case 10
dofont Label11(ct), Label12(ct)
Case 11
dofont Label12(ct), Label13(ct)
Case 12
dofont Label13(ct), Label14(ct)
Case 13
dofont Label14(ct), Label15(ct)
Case 14
dofont Label15(ct)
End Select
Next

With labscatter(pct - 1)
.Left = resX * .Left
.Top = resY * .Top
.Width = resX * .Width
.Height = resY * .Height

.fontname = "Times New Roman"
.Fontsize = 9.75 * resY
.Font.Charset = 0
.Font.Weight = 400
.FontUnderline = 0           'False
.FontItalic = -1   'True
.FontStrikethru = 0       'False
End With

Next



For pct = piccount To 13

For ct = 0 To 4
Select Case pct
Case 4
Label6(ct).Visible = False
Case 5
Label7(ct).Visible = False
Case 6
Label8(ct).Visible = False
Case 7
Label9(ct).Visible = False
Case 8
Label10(ct).Visible = False
Case 9
Label11(ct).Visible = False
Case 10
Label12(ct).Visible = False
Case 11
Label13(ct).Visible = False
Case 12
Label14(ct).Visible = False
Case 13
Label15(ct).Visible = False
End Select

Next
imgsel(pct).Visible = False
Next
If piccount < 8 Then
For intreel = 5 To 9
Regen(intreel).Visible = False
Next
End If
If gt(0) = 5 Then       'precludes wiping settings when coming back in on gt(6)
    Stopnoise 1
    Cancel.Enabled = False
    'Eliminate previous scatter, substitute settings, set default prizes for old scatters
    If piccount < 14 Then   'get rid of all irrelevant settings
    For pct = piccount + 1 To 14
    For ct = 0 To 10
    sst(pct, ct) = 0
    Next
    Next
    End If
    
    
    'reset old scatter prizes
    For pct = 1 To piccount
        If sst(pct, 2) > 0 Then
        For ct = 6 To 10
        resetprize ct, pct
        Next
        End If
    Next
    
    For intreel = 0 To 4
    'zero wheelvec for cleanliness
    For pct = 1 To 14
    wheelvec(intreel + 1, pct) = 0
    Next
    'Dummy wheelorder ;clears nuisance lt 5 symbols
    For ct = 1 To 24
    wheelorder(intreel, ct) = 1
    Next
    Next
        
    
    For pct = 1 To piccount
    'There does not exist a previous prize, possibly from previous LT 5 condition
    If sst(pct, 6) = 0 Or sst(pct, 7) = 0 Or sst(pct, 8) = 0 Then
    sst(pct, 1) = 1    'Set LR values
    sst(pct, 5) = 0
    sst(pct, 3) = 0
    sst(pct, 4) = 0
    For ct = 6 To 10
    resetprize ct, pct
    Next
    End If
    Next
    
    firstgametypeload
    
    zeroscatter
    
    'And some startup defaults
        
    Restoredefaults True
    If gt(185) > 0 Then SndMidInit
    
    genrandomseed True
    VOGchg = 0
    gt(2) = CLng(Right(CStr(gt(155)), 3))
    gt(55) = -1     'restarts with same RS
    gt(56) = -1
    If gt(158) = 1 And InStr(Stringvars(5), "MyReels_Game") Then
    If InStr(Len(loaddirectory), loaddirectory, "\") Then newstr = Left(loaddirectory, Len(loaddirectory) - 1)
    For ct = 1 To Len(newstr)
    If InStr(Len(newstr) - ct + 1, newstr, "\") Then Exit For
    Next
    newstr = Right(newstr, ct - 1)
    newstr = Left(newstr, 12)
    Stringvars(5) = newstr & "_Game!"
    Stringvars(8) = newstr & "_Player"
    Stringvars(13) = "Welcome, " & newstr & "_Player!"
    End If
    

    DoQuotes
    
    
    initwheelvec
    For intreel = 0 To 4
    initcaptions (intreel)
    Next

    For pct = 0 To piccount - 1
    intnotzero(pct) = 5
    Next

Else
    
    For pct = 1 To piccount
    If sst(pct, 2) = 1 Then
    labscatter(pct - 1).Visible = True
    For intreel = 0 To 9
    Regen(intreel).Enabled = False
    Next
    End If
    Next
    
    For intreel = 0 To 4
    For pct = 1 To piccount
    initvec(pct) = wheelvec(intreel + 1, pct)
    Next
    initcaptions (intreel)
    Next

    For pct = 0 To piccount - 1
    intnotzero(pct) = 0
    For intreel = 1 To 5
    If wheelvec(intreel, pct + 1) > 0 Then intnotzero(pct) = intnotzero(pct) + 1
    Next
    Next


End If


'Now allocate temp prizes
For pct = 1 To piccount
sstscat(pct) = sst(pct, 2) * wheelvec(1, pct)

'To enable the scatter checkbox in SSTAB
scatterdonehere(pct - 1) = reelcheck(pct, 0)

Next



If gt(1) = 1 Then Chkordered.Value = 1

procend = True

cornfig.Show
End Sub
Private Sub Form_Resize()
Dotaskwindow Me
End Sub
Private Sub Chkordered_Click()
If procend = False Then Exit Sub
If Chkordered.Value = 1 Then
gt(1) = 1
Else
gt(1) = 0
End If
End Sub
Private Sub Cancel_Click()
If gt(0) < 5 Then gt(0) = 4
Unload cornfig
Load gametype
gametype.Show
End Sub
Private Sub Continue_Click()
Dim wheeltmp(5, 14) As Long, response As Long

For intreel = 0 To 4
wheeltmp(intreel + 1, 1) = CLng(Label2(intreel).Caption)
wheeltmp(intreel + 1, 2) = CLng(Label3(intreel).Caption)
wheeltmp(intreel + 1, 3) = CLng(Label4(intreel).Caption)
Next
If piccount = 3 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 4) = CLng(Label5(intreel).Caption)
Next
If piccount = 4 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 5) = CLng(Label6(intreel).Caption)
Next
If piccount = 5 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 6) = CLng(Label7(intreel).Caption)
Next
If piccount = 6 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 7) = CLng(Label8(intreel).Caption)
Next
If piccount = 7 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 8) = CLng(Label9(intreel).Caption)
Next
If piccount = 8 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 9) = CLng(Label10(intreel).Caption)
Next
If piccount = 9 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 10) = CLng(Label11(intreel).Caption)
Next
If piccount = 10 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 11) = CLng(Label12(intreel).Caption)
Next
If piccount = 11 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 12) = CLng(Label13(intreel).Caption)
Next
If piccount = 12 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 13) = CLng(Label14(intreel).Caption)
Next
If piccount = 13 Then GoTo finish
For intreel = 0 To 4
wheeltmp(intreel + 1, 14) = CLng(Label15(intreel).Caption)
Next
finish:

'Resolve previous LT 5
For pct = 1 To piccount
If sst(pct, 0) > 0 Then
sst(pct, 1) = 1    'Set LR values
For ct = 2 To 5
sst(pct, ct) = 0
Next
For ct = 6 To 10
resetprize ct, pct
Next
End If
Next


'clear old scatters
For pct = 1 To piccount

If sst(pct, 2) * wheelvec(1, pct) <> sstscat(pct) And sst(pct, 2) > 0 Then 'any scatters changed or deleted?

For ct = 0 To 5
Select Case ct
Case 0
sst(pct, 0) = 0
Case 1
sst(pct, 1) = 1 'Set LR values
Case 2
sst(pct, 2) = 0
'Reset Prizes first
Case Else
sst(pct, ct) = 0
End Select
Next
For ct = 6 To 10
resetprize ct, pct
Next

End If
Next




For pct = 1 To piccount 'Now Allocate new scatters
reelcheck(pct, 0) = scatterdonehere(pct - 1)

If sstscat(pct) > sst(pct, 2) * wheelvec(1, pct) Then  'new scatter
sst(pct, 0) = 0 'definitely not LT 5!
sst(pct, 1) = 1
sst(pct, 2) = 1
sst(pct, 3) = 0
sst(pct, 4) = 0
sst(pct, 5) = 0
sst(pct, 10) = 0
CalcScatterprize pct, sstscat(pct), 2
End If
Next


'Clear all spin/games
For ct = 0 To 1
disablegamespintabs(ct) = False
For temp = 1 To 9
freegamesettings(ct, temp) = 0
Next
For temp = 1 To 15
spinsettings(ct, temp) = 0
Next
Next

For ct = 0 To 3
gamespinsymbol(ct) = 0
gamespinkeep(ct) = 0
Next

'Best to clear all substitutes
For ct = 1 To piccount
For intreel = 1 To 5
reelcheck(ct, intreel) = True
Next

For pct = 1 To piccount
substitute(ct, pct) = False
Next
Next

'Compare changes for stats
If gt(0) < 5 Then
For intreel = 1 To 5
For pct = 1 To piccount
If wheeltmp(intreel, pct) <> wheelvec(intreel, pct) Then gt(26) = gt(26) + 1
Next
Next
End If

For intreel = 1 To 5
For pct = 1 To 14
wheelvec(intreel, pct) = wheeltmp(intreel, pct)
Next
Next

Moretries:
If randwheelvec(piccount, wheelvec, wheelorder) = True Then   'This sets wheelorder, scatters
setthumbspiccount thumbs, piccount, wheelvec, wheelorder
Rnd -1     'reset seeding
Randomize (gt(35))
Unload Me
Set cornfig = Nothing
Load gametype
gametype.Show
Else
Cancel.Enabled = False
response = MsgBox("Generation failed - please click Ok to retry. If, after many retries, problem still persists, click Cancel and rerun program. The newly created pics will not be deleted.", vbOKCancel)
If response = 2 Then
Rnd -1     'reset seeding
Randomize (gt(35))
Unload cornfig
Else
GoTo Moretries
End If
End If
End Sub
Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tempplusone = CLng(Label2(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
    If gt(1) = 1 Then
    If tempplusone > CLng(Label3(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
If tempplusone = 1 Then intnotzero(0) = intnotzero(0) + 1
Label2(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If checkintzero(tempminusone, 0) = 0 Then Exit Sub
Label2(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 1
End Sub
Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tempplusone = CLng(Label3(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
    If gt(1) = 1 Then
    If tempplusone > CLng(Label4(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
If tempplusone = 1 Then intnotzero(1) = intnotzero(1) + 1
Label3(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label2(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 1) = 0 Then Exit Sub
Label3(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 2
End Sub
Private Sub Label4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
tempplusone = CLng(Label4(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 3 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label5(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(2) = intnotzero(2) + 1
Label4(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label3(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 2) = 0 Then Exit Sub
Label4(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 3
End Sub
Private Sub Label5_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 4 Then Exit Sub
tempplusone = CLng(Label5(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 4 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label6(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(3) = intnotzero(3) + 1
Label5(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label4(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 3) = 0 Then Exit Sub
Label5(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 4
End Sub
Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 5 Then Exit Sub
tempplusone = CLng(Label6(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 5 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label7(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(4) = intnotzero(4) + 1
Label6(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label5(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 4) = 0 Then Exit Sub
Label6(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 5
End Sub
Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 6 Then Exit Sub
tempplusone = CLng(Label7(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 6 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label8(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(5) = intnotzero(5) + 1
Label7(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label6(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 5) = 0 Then Exit Sub
Label7(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 6
End Sub
Private Sub Label8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 7 Then Exit Sub
tempplusone = CLng(Label8(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 7 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label9(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(6) = intnotzero(6) + 1
Label8(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label7(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 6) = 0 Then Exit Sub
Label8(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 7
End Sub
Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 8 Then Exit Sub
tempplusone = CLng(Label9(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 8 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label10(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(7) = intnotzero(7) + 1
Label9(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label8(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 7) = 0 Then Exit Sub
Label9(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 8
End Sub
Private Sub Label10_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 9 Then Exit Sub
tempplusone = CLng(Label10(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 9 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label11(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(8) = intnotzero(8) + 1
Label10(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label9(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 8) = 0 Then Exit Sub
Label10(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 9
End Sub
Private Sub Label11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 10 Then Exit Sub
tempplusone = CLng(Label11(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 10 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label12(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(9) = intnotzero(9) + 1
Label11(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label10(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 9) = 0 Then Exit Sub
Label11(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 10
End Sub
Private Sub Label12_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 11 Then Exit Sub
tempplusone = CLng(Label12(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 11 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label13(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(10) = intnotzero(10) + 1
Label12(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label11(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 10) = 0 Then Exit Sub
Label12(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 11
End Sub
Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 12 Then Exit Sub
tempplusone = CLng(Label13(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 12 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label14(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(11) = intnotzero(11) + 1
Label13(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label12(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 11) = 0 Then Exit Sub
Label13(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 12
End Sub
Private Sub Label14_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 13 Then Exit Sub
tempplusone = CLng(Label14(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If piccount = 13 Then
If tempplusone = 24 Then Exit Sub
Else
    If gt(1) = 1 Then
    If tempplusone > CLng(Label15(Index).Caption) Then Exit Sub
    Else
    If tempplusone > 24 Then Exit Sub
    End If
End If
If tempplusone = 1 Then intnotzero(12) = intnotzero(12) + 1
Label14(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label13(Index).Caption) Then Exit Sub
If checkintzero(tempminusone, 12) = 0 Then Exit Sub
Label14(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 13
End Sub
Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If piccount < 14 Then Exit Sub
tempplusone = CLng(Label15(Index).Caption) + 1
tempminusone = tempplusone - 2
poscolortemp = poscolor(Index)
If Button = 1 Then
If tempplusone > 24 Then Exit Sub
Label15(Index).Caption = CStr(tempplusone)
poscolortemp = poscolortemp + 1
Else
If tempminusone < 0 Then Exit Sub
If gt(1) = 1 And tempminusone < CLng(Label14(Index).Caption) Then Exit Sub
Label15(Index).Caption = CStr(tempminusone)
poscolortemp = poscolortemp - 1
End If
Indexcolor Index, 14
End Sub
Private Sub Regen_Click(Index As Integer)
Dim wasitzero(15) As Boolean

If procend = False Then Exit Sub

intreel = Index

If intreel > 4 Then intreel = intreel - 5


For pct = 2 To piccount + 1
Select Case pct
Case 2
If Label2(intreel).Caption > 0 Then
wasitzero(2) = False
Else
wasitzero(2) = True
End If
Case 3
If Label3(intreel).Caption > 0 Then
wasitzero(3) = False
Else
wasitzero(3) = True
End If
Case 4
If Label4(intreel).Caption > 0 Then
wasitzero(4) = False
Else
wasitzero(4) = True
End If
Case 5
If Label5(intreel).Caption > 0 Then
wasitzero(5) = False
Else
wasitzero(5) = True
End If
Case 6
If Label6(intreel).Caption > 0 Then
wasitzero(6) = False
Else
wasitzero(6) = True
End If
Case 7
If Label7(intreel).Caption > 0 Then
wasitzero(7) = False
Else
wasitzero(7) = True
End If
Case 8
If Label8(intreel).Caption > 0 Then
wasitzero(8) = False
Else
wasitzero(8) = True
End If
Case 9
If Label9(intreel).Caption > 0 Then
wasitzero(9) = False
Else
wasitzero(9) = True
End If
Case 10
If Label10(intreel).Caption > 0 Then
wasitzero(10) = False
Else
wasitzero(10) = True
End If
Case 11
If Label11(intreel).Caption > 0 Then
wasitzero(11) = False
Else
wasitzero(11) = True
End If
Case 12
If Label12(intreel).Caption > 0 Then
wasitzero(12) = False
Else
wasitzero(12) = True
End If
Case 13
If Label13(intreel).Caption > 0 Then
wasitzero(13) = False
Else
wasitzero(13) = True
End If
Case 14
If Label14(intreel).Caption > 0 Then
wasitzero(14) = False
Else
wasitzero(14) = True
End If
Case 15
If Label15(intreel).Caption > 0 Then
wasitzero(15) = False
Else
wasitzero(15) = True
End If
End Select
Next

initwheelvec
initcaptions (intreel)
poscolor(intreel) = 0

Regen(Index).ToolTipText = ""
Regen(indexplus5).ToolTipText = ""
Continue.Enabled = True
For ct = 0 To 4
If poscolor(ct) > 0 Then
Continue.Enabled = False
Exit For
End If
Next

Regen(intreel).BackColor = vbButtonFace
Regen(intreel + 5).BackColor = vbButtonFace

For pct = 0 To piccount - 1
If wasitzero(pct + 2) = True Then intnotzero(pct) = intnotzero(pct) + 1
Next

End Sub
Private Sub Regenerate_click()
For pct = 1 To piccount

sstscat(pct) = 0
scatterdonehere(pct - 1) = False

labscatter(pct - 1).Caption = ""
Next
initwheelvec
For intreel = 0 To 4
initcaptions (intreel)
poscolor(intreel) = 0
Regen(intreel).Enabled = True
Regen(intreel).BackColor = vbButtonFace
Regen(intreel).ToolTipText = ""
Regen(intreel + 5).ToolTipText = ""
If piccount > 7 Then
Regen(intreel + 5).Enabled = True
Regen(intreel + 5).BackColor = vbButtonFace
End If
Next

For pct = 0 To piccount - 1
intnotzero(pct) = 0
For intreel = 1 To 5
intnotzero(pct) = 5
Next
Next

Continue.Enabled = True
End Sub
Private Sub imgsel_Click(Index As Integer)
Dim inttest(4) As Long, okay As Boolean
procend = False
If intlastindex > 0 Then
If labscatter(intlastindex - 1).Caption <> "Scatter" Then labscatter(intlastindex - 1).Visible = False
End If
intlastindex = Index + 1
ct = 0
'This is to select scratch symbols

If sstscat(Index + 1) > 0 Then
noscatterplease Index + 1
Else    'request for scatter
For pct = 1 To piccount
If sstscat(pct) > 0 Then ct = ct + 1
Next
If ct > 1 Then
labscatter(Index).Visible = True
labscatter(Index).Caption = "Only 2 scatter symbols allowed"
procend = True
Exit Sub
End If

For pct = 1 To piccount
If substitute(Index + 1, pct) = True Or substitute(pct, Index + 1) = True Then
labscatter(Index).Visible = True
labscatter(Index).Caption = "No scatter in a substitute combo"
procend = True
Exit Sub
End If
Next

If intnotzero(Index) < 5 Then
labscatter(Index).Visible = True
labscatter(Index).Caption = "Scatters must be on every reel"
procend = True
Exit Sub
End If

Select Case Index
Case 0
inttest(0) = Label2(0).Caption
inttest(1) = Label2(1).Caption
inttest(2) = Label2(2).Caption
inttest(3) = Label2(3).Caption
inttest(4) = Label2(4).Caption
Case 1
inttest(0) = Label3(0).Caption
inttest(1) = Label3(1).Caption
inttest(2) = Label3(2).Caption
inttest(3) = Label3(3).Caption
inttest(4) = Label3(4).Caption
Case 2
inttest(0) = Label4(0).Caption
inttest(1) = Label4(1).Caption
inttest(2) = Label4(2).Caption
inttest(3) = Label4(3).Caption
inttest(4) = Label4(4).Caption
Case 3
inttest(0) = Label5(0).Caption
inttest(1) = Label5(1).Caption
inttest(2) = Label5(2).Caption
inttest(3) = Label5(3).Caption
inttest(4) = Label5(4).Caption
Case 4
inttest(0) = Label6(0).Caption
inttest(1) = Label6(1).Caption
inttest(2) = Label6(2).Caption
inttest(3) = Label6(3).Caption
inttest(4) = Label6(4).Caption
Case 5
inttest(0) = Label7(0).Caption
inttest(1) = Label7(1).Caption
inttest(2) = Label7(2).Caption
inttest(3) = Label7(3).Caption
inttest(4) = Label7(4).Caption
Case 6
inttest(0) = Label8(0).Caption
inttest(1) = Label8(1).Caption
inttest(2) = Label8(2).Caption
inttest(3) = Label8(3).Caption
inttest(4) = Label8(4).Caption
Case 7
inttest(0) = Label9(0).Caption
inttest(1) = Label9(1).Caption
inttest(2) = Label9(2).Caption
inttest(3) = Label9(3).Caption
inttest(4) = Label9(4).Caption
Case 8
inttest(0) = Label10(0).Caption
inttest(1) = Label10(1).Caption
inttest(2) = Label10(2).Caption
inttest(3) = Label10(3).Caption
inttest(4) = Label10(4).Caption
Case 9
inttest(0) = Label11(0).Caption
inttest(1) = Label11(1).Caption
inttest(2) = Label11(2).Caption
inttest(3) = Label11(3).Caption
inttest(4) = Label11(4).Caption
Case 10
inttest(0) = Label12(0).Caption
inttest(1) = Label12(1).Caption
inttest(2) = Label12(2).Caption
inttest(3) = Label12(3).Caption
inttest(4) = Label12(4).Caption
Case 11
inttest(0) = Label13(0).Caption
inttest(1) = Label13(1).Caption
inttest(2) = Label13(2).Caption
inttest(3) = Label13(3).Caption
inttest(4) = Label13(4).Caption
Case 12
inttest(0) = Label14(0).Caption
inttest(1) = Label14(1).Caption
inttest(2) = Label14(2).Caption
inttest(3) = Label14(3).Caption
inttest(4) = Label14(4).Caption
Case 13
inttest(0) = Label15(0).Caption
inttest(1) = Label15(1).Caption
inttest(2) = Label15(2).Caption
inttest(3) = Label15(3).Caption
inttest(4) = Label15(4).Caption
End Select
okay = testscatter(inttest(0), inttest(1), inttest(2), inttest(3), inttest(4))
If okay = False Then
labscatter(Index).Visible = True
labscatter(Index).Caption = "11111, 22222, 44444 for scatter only"
procend = True
Exit Sub
End If

labscatter(Index).Visible = True
labscatter(Index).Caption = "Scatter"
sstscat(Index + 1) = inttest(0)   'Assign frequency of scatters to this value
scatterdonehere(Index) = True


For intreel = 0 To 4
Regen(intreel).Enabled = False
Next
If piccount > 7 Then
For intreel = 0 To 4
Regen(intreel + 5).Enabled = False
Next
End If
End If
procend = True
End Sub
Private Function checkintzero(tempminusone As Long, intwhichpic As Long)
'if old caption was 1
If tempminusone = 0 Then
intnotzero(intwhichpic) = intnotzero(intwhichpic) - 1
Select Case intnotzero(intwhichpic)
Case Is <= 0
intnotzero(intwhichpic) = 1
checkintzero = 0
Case Is > 0
checkintzero = 1
End Select
Else
checkintzero = 1
End If
End Function
Private Sub Indexcolor(Index As Integer, picindex As Long)

noscatterplease picindex

indexplus5 = Index + 5
If poscolortemp > 0 Then
Regen(Index).BackColor = &H80FF&
Regen(indexplus5).BackColor = &H80FF&
ElseIf poscolortemp = 0 Then
Regen(Index).BackColor = vbButtonFace
Regen(indexplus5).BackColor = vbButtonFace
Else
Regen(Index).BackColor = &HFFFF00
Regen(indexplus5).BackColor = &HFFFF00
End If
poscolor(Index) = poscolortemp

Regen(Index).ToolTipText = ""
Regen(indexplus5).ToolTipText = ""
Continue.Enabled = True
For ct = 0 To 4
If poscolor(ct) <> 0 Then
Continue.Enabled = False
Regen(Index).ToolTipText = "Reel buttons appear orange or blue if the columns don't add to 24"
Regen(indexplus5).ToolTipText = "Reel buttons appear orange or blue if the columns don't add to 24"
Exit For
End If
Next

End Sub
Private Sub initcaptions(intreel)
Label2(intreel).Caption = initvec(1)
Label3(intreel).Caption = initvec(2)
Label4(intreel).Caption = initvec(3)
If piccount = 3 Then Exit Sub
Label5(intreel).Caption = initvec(4)
If piccount = 4 Then Exit Sub
Label6(intreel).Caption = initvec(5)
If piccount = 5 Then Exit Sub
Label7(intreel).Caption = initvec(6)
If piccount = 6 Then Exit Sub
Label8(intreel).Caption = initvec(7)
If piccount = 7 Then Exit Sub
Label9(intreel).Caption = initvec(8)
If piccount = 8 Then Exit Sub
Label10(intreel).Caption = initvec(9)
If piccount = 9 Then Exit Sub
Label11(intreel).Caption = initvec(10)
If piccount = 10 Then Exit Sub
Label12(intreel).Caption = initvec(11)
If piccount = 11 Then Exit Sub
Label13(intreel).Caption = initvec(12)
If piccount = 12 Then Exit Sub
Label14(intreel).Caption = initvec(13)
If piccount = 13 Then Exit Sub
Label15(intreel).Caption = initvec(14)
End Sub
Private Function noscatterplease(picindex As Long)
ct = 0

scatterdonehere(picindex - 1) = False
labscatter(picindex - 1).Visible = False
sstscat(picindex) = 0

For pct = 1 To piccount
If sstscat(pct) > 0 Then ct = ct + 1
Next
If ct = 0 Then 'This is when the last scatter has been deselected
    For intreel = 0 To 4
        Regen(intreel).Enabled = True
        If piccount > 7 Then
        Regen(intreel + 5).Enabled = True
        End If
    Next
End If

End Function
Private Sub dofont(objekt As Label, Optional nextobj As Label)
With objekt
.Left = resX * .Left
.Top = resY * .Top
.Height = resY * .Height
.fontname = "Times New Roman"
.Fontsize = resY * 14.25
.Font.Charset = 0
.Font.Weight = 400
.FontUnderline = 0           'False
.FontItalic = 0              'False
.FontStrikethru = 0       'False
End With

If pct < 14 And pct = piccount Then nextobj.Visible = False

End Sub
Private Sub DoQuotes()
'Need to clear Quotes.s$t + Filesize warnings!
On Error GoTo quoterror

If olddirectory = loaddirectory Then
    If findafile(loaddirectory, "Quotes.s$t") = 0 Then GoTo quoterror
    If Stringvars(3) <> "" Then Stringvars(3) = loaddirectory

Else
    If findafile(olddirectory, "Quotes.s$t") > 0 Then
        If findafile(loaddirectory, "Quotes.s$t") = 0 Then FileCopy olddirectory & "Quotes.s$t", loaddirectory & "Quotes.s$t"
            If olddirectory = App.Path & "\" Then 'condense quotes
                DoEvents
                sDatabaseName = loaddirectory & "Quotes.s$t"
                DoEvents
                If Openquotes("Clear.txt", True, , True) = False Then GoTo quoterror
            End If
        If Stringvars(3) <> "" Then
        If gt(193) = 0 Then
        Stringvars(3) = loaddirectory
        Else
        Stringvars(3) = App.Path & "\"
        End If
        End If
    Else
        If findafile(loaddirectory, "Quotes.s$t") = 0 Then
            If loaddirectory = App.Path & "\" Then GoTo quoterror
            
            If findafile(App.Path & "\", "Quotes.s$t") > 0 Then
            FileCopy App.Path & "\" & "Quotes.s$t", loaddirectory & "Quotes.s$t"
            If Stringvars(3) <> "" Then Stringvars(3) = App.Path & "\"
            DoEvents
            sDatabaseName = loaddirectory & "Quotes.s$t"
            DoEvents
            'Condense spare quote file to save space and use basedir
            If Openquotes("Clear.txt", True, , True) = False Then GoTo quoterror
            sDatabaseName = App.Path & "\" & "Quotes.s$t"
            DoEvents
            If Openquotes("Dummy.txt", True) = False Then GoTo quoterror
            gt(195) = 1
            gt(193) = 1
            Else
            GoTo quoterror
            End If
        Else
            If Stringvars(3) <> "" Then
            If gt(193) = 0 Then
            Stringvars(3) = loaddirectory
            Else
            Stringvars(3) = App.Path & "\"
            End If
            End If
        End If
    End If
End If

If Len(Stringvars(3)) > 100 Then GoTo quoterror
olddirectory = ""
Exit Sub
quoterror:
ShowError
MsgBox "Quotes.s$t pathname > 100 characters or corrupt or inaccessible from its original directory. Continuing ...", vbOKOnly
Stringvars(3) = ""
olddirectory = ""
End Sub
