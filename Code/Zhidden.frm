VERSION 5.00
Begin VB.Form Zhidden 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   5175
   ClientLeft      =   2115
   ClientTop       =   3135
   ClientWidth     =   10035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Zhidden.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Cellz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   1670
      Width           =   9800
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   70
         Left            =   0
         TabIndex        =   98
         Top             =   340
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   71
         Left            =   0
         TabIndex        =   97
         Top             =   580
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   72
         Left            =   0
         TabIndex        =   96
         Top             =   820
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   73
         Left            =   0
         TabIndex        =   95
         Top             =   1060
         Width           =   1095
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   74
         Left            =   0
         TabIndex        =   94
         Top             =   1300
         Width           =   1095
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   75
         Left            =   0
         TabIndex        =   93
         Top             =   1540
         Width           =   1095
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   76
         Left            =   0
         TabIndex        =   92
         Top             =   1780
         Width           =   1095
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   77
         Left            =   0
         TabIndex        =   91
         Top             =   2020
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   78
         Left            =   0
         TabIndex        =   90
         Top             =   2260
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   79
         Left            =   0
         TabIndex        =   89
         Top             =   2500
         Width           =   1100
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   88
         Top             =   0
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   87
         Top             =   340
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   86
         Top             =   580
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   85
         Top             =   820
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   84
         Top             =   1060
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   83
         Top             =   1300
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   82
         Top             =   1540
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   81
         Top             =   1780
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   80
         Top             =   2020
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   79
         Top             =   2260
         Width           =   2200
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   78
         Top             =   2500
         Width           =   2200
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   2200
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2200
         TabIndex        =   76
         Top             =   0
         Width           =   2190
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   4400
         TabIndex        =   75
         Top             =   0
         Width           =   1000
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   74
         Top             =   0
         Width           =   1100
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   6500
         TabIndex        =   73
         Top             =   0
         Width           =   1105
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   7600
         TabIndex        =   72
         Top             =   0
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   2200
         TabIndex        =   71
         Top             =   340
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   2200
         TabIndex        =   70
         Top             =   580
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   2200
         TabIndex        =   69
         Top             =   820
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   2200
         TabIndex        =   68
         Top             =   1060
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   2200
         TabIndex        =   67
         Top             =   1300
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   2200
         TabIndex        =   66
         Top             =   1540
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   2200
         TabIndex        =   65
         Top             =   1780
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   2200
         TabIndex        =   64
         Top             =   2020
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   2200
         TabIndex        =   63
         Top             =   2260
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2200
         TabIndex        =   62
         Top             =   2500
         Width           =   2195
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   4400
         TabIndex        =   61
         Top             =   340
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   4400
         TabIndex        =   60
         Top             =   580
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   4400
         TabIndex        =   59
         Top             =   820
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   4400
         TabIndex        =   58
         Top             =   1060
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   4400
         TabIndex        =   57
         Top             =   1300
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   4400
         TabIndex        =   56
         Top             =   1540
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4400
         TabIndex        =   55
         Top             =   1780
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4400
         TabIndex        =   54
         Top             =   2020
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   4400
         TabIndex        =   53
         Top             =   2260
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   4400
         TabIndex        =   52
         Top             =   2500
         Width           =   1000
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   5400
         TabIndex        =   51
         Top             =   340
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   5400
         TabIndex        =   50
         Top             =   580
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   5400
         TabIndex        =   49
         Top             =   820
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   5400
         TabIndex        =   48
         Top             =   1060
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   5400
         TabIndex        =   47
         Top             =   1300
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   5400
         TabIndex        =   46
         Top             =   1540
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   5400
         TabIndex        =   45
         Top             =   1780
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   5400
         TabIndex        =   44
         Top             =   2020
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   5400
         TabIndex        =   43
         Top             =   2260
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   5400
         TabIndex        =   42
         Top             =   2500
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   6500
         TabIndex        =   41
         Top             =   340
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   6500
         TabIndex        =   40
         Top             =   580
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   6500
         TabIndex        =   39
         Top             =   820
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   6500
         TabIndex        =   38
         Top             =   1060
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   6500
         TabIndex        =   37
         Top             =   1300
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   6500
         TabIndex        =   36
         Top             =   1540
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   6500
         TabIndex        =   35
         Top             =   1780
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   6500
         TabIndex        =   34
         Top             =   2020
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   6500
         TabIndex        =   33
         Top             =   2260
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   49
         Left            =   6500
         TabIndex        =   32
         Top             =   2500
         Width           =   1105
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   50
         Left            =   7600
         TabIndex        =   31
         Top             =   340
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   51
         Left            =   7600
         TabIndex        =   30
         Top             =   580
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   52
         Left            =   7600
         TabIndex        =   29
         Top             =   820
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   53
         Left            =   7600
         TabIndex        =   28
         Top             =   1060
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   54
         Left            =   7600
         TabIndex        =   27
         Top             =   1300
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   55
         Left            =   7600
         TabIndex        =   26
         Top             =   1540
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   56
         Left            =   7600
         TabIndex        =   25
         Top             =   1780
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   57
         Left            =   7600
         TabIndex        =   24
         Top             =   2020
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   58
         Left            =   7600
         TabIndex        =   23
         Top             =   2260
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   59
         Left            =   7600
         TabIndex        =   22
         Top             =   2500
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   60
         Left            =   8700
         TabIndex        =   21
         Top             =   340
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   61
         Left            =   8700
         TabIndex        =   20
         Top             =   580
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   62
         Left            =   8700
         TabIndex        =   19
         Top             =   820
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   63
         Left            =   8700
         TabIndex        =   18
         Top             =   1060
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   64
         Left            =   8700
         TabIndex        =   17
         Top             =   1300
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   65
         Left            =   8700
         TabIndex        =   16
         Top             =   1540
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   66
         Left            =   8700
         TabIndex        =   15
         Top             =   1780
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   67
         Left            =   8700
         TabIndex        =   14
         Top             =   2020
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   68
         Left            =   8700
         TabIndex        =   13
         Top             =   2260
         Width           =   1100
      End
      Begin VB.Label lblCell 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   69
         Left            =   8700
         TabIndex        =   12
         Top             =   2500
         Width           =   1095
      End
      Begin VB.Label lblCelltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   8700
         TabIndex        =   11
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2775
      Left            =   9820
      TabIndex        =   8
      Top             =   1670
      Width           =   200
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   205
      Left            =   0
      TabIndex        =   7
      Top             =   4410
      Width           =   9820
   End
   Begin VB.TextBox HiddenTitle 
      Height          =   475
      Left            =   3800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdok 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   4600
      Width           =   10095
   End
   Begin VB.Timer Cellcolour 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   3600
   End
   Begin VB.PictureBox Spare 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "Zhidden.frx":000C
      ScaleHeight     =   1575
      ScaleWidth      =   9975
      TabIndex        =   1
      Top             =   -10
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Label lblzhidd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   6600
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblzhidd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fair"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblzhidd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Good"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   4560
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblzhidd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Great"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblzhidd 
         Alignment       =   2  'Center
         Height          =   450
         Index           =   0
         Left            =   3780
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.Menu OptionZ 
      Caption         =   "Optionz"
      Visible         =   0   'False
      Begin VB.Menu Spinner 
         Caption         =   "Spin [Space]"
      End
      Begin VB.Menu Bar 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu ChangeBet 
         Caption         =   "Change Bet [Ctrl]"
      End
      Begin VB.Menu Changeln1 
         Caption         =   "Bet 1 line [1]"
      End
      Begin VB.Menu Changeln2 
         Caption         =   "Bet 2 lines [2]"
      End
      Begin VB.Menu Changeln3 
         Caption         =   "Bet 3 lines [3]"
      End
      Begin VB.Menu Cheat 
         Caption         =   "Cheat [C]"
      End
      Begin VB.Menu Bar1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu Soundd 
         Caption         =   ""
      End
      Begin VB.Menu Thumbb 
         Caption         =   ""
      End
      Begin VB.Menu Musicc 
         Caption         =   ""
      End
      Begin VB.Menu Bar2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu Configurationn 
         Caption         =   "Configuration [Enter]"
      End
      Begin VB.Menu Bar3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu ChangeDirectory 
         Caption         =   "Change Directory [D]"
      End
      Begin VB.Menu Helpp 
         Caption         =   "Help [F1]"
      End
      Begin VB.Menu Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu Quitt 
         Caption         =   "Quit [Esc]"
      End
      Begin VB.Menu Bar5 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About [A]"
      End
      Begin VB.Menu Bar6 
         Caption         =   "-"
      End
      Begin VB.Menu Titleadjust 
         Caption         =   "Adjust Title [T]"
      End
      Begin VB.Menu ThisMenu 
         Caption         =   "This Menu [Shift]"
      End
   End
End
Attribute VB_Name = "Zhidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DRIVERVERSION = 0
Const TECHNOLOGY = 2
Const HORZSIZE = 4
Const VERTSIZE = 6
Const HORZRES = 8
Const VERTRES = 10
Const BITSPIXEL = 12
Const PLANES = 14
Const NUMBRUSHES = 16
Const NUMPENS = 18
Const NUMMARKERS = 20
Const NUMFONTS = 22
Const NUMCOLORS = 24
Const PDEVICESIZE = 26
Const CURVECAPS = 28
Const LINECAPS = 30
Const POLYGONALCAPS = 32
Const TEXTCAPS = 34
Const CLIPCAPS = 36
Const RASTERCAPS = 38
Const ASPECTX = 40
Const ASPECTY = 42
Const ASPECTXY = 44
Const SIZEPALETTE = 104
Const NUMRESERVED = 106
Const COLORRES = 108
Const DT_PLOTTER = 0
Const DT_RASDISPLAY = 1
Const DT_RASPRINTER = 2
Const DT_RASCAMERA = 3
Const DT_CHARSTREAM = 4
Const DT_METAFILE = 5
Const DT_DISPFILE = 6
Const CC_NONE = 0
Const CC_CIRCLES = 1
Const CC_PIE = 2
Const CC_CHORD = 4
Const CC_ELLIPSES = 8
Const CC_WIDE = 16
Const CC_STYLED = 32
Const CC_WIDESTYLED = 64
Const CC_INTERIORS = 128
Const LC_NONE = 0
Const LC_POLYLINE = 2
Const LC_MARKER = 4
Const LC_POLYMARKER = 8
Const LC_WIDE = 16
Const LC_STYLED = 32
Const LC_WIDESTYLED = 64
Const LC_INTERIORS = 128
Const PC_NONE = 0
Const PC_POLYGON = 1
Const PC_RECTANGLE = 2
Const PC_WINDPOLYGON = 4
Const PC_TRAPEZOID = 4
Const PC_SCANLINE = 8
Const PC_WIDE = 16
Const PC_STYLED = 32
Const PC_WIDESTYLED = 64
Const PC_INTERIORS = 128
Const CP_NONE = 0
Const CP_RECTANGLE = 1
Const TC_OP_CHARACTER = &H1
Const TC_OP_STROKE = &H2
Const TC_CP_STROKE = &H4
Const TC_CR_90 = &H8
Const TC_CR_ANY = &H10
Const TC_SF_X_YINDEP = &H20
Const TC_SA_DOUBLE = &H40
Const TC_SA_INTEGER = &H80
Const TC_SA_CONTIN = &H100
Const TC_EA_DOUBLE = &H200
Const TC_IA_ABLE = &H400
Const TC_UA_ABLE = &H800
Const TC_SO_ABLE = &H1000
Const TC_RA_ABLE = &H2000
Const TC_VA_ABLE = &H4000
Const TC_RESERVED = &H8000
Const RC_BITBLT = 1
Const RC_BANDING = 2
Const RC_SCALING = 4
Const RC_BITMAP64 = 8
Const RC_GDI20_OUTPUT = &H10
Const RC_DI_BITMAP = &H80
Const RC_PALETTE = &H100
Const RC_DIBTODEV = &H200
Const RC_BIGFONT = &H400
Const RC_STRETCHBLT = &H800
Const RC_FLOODFILL = &H1000
Const RC_STRETCHDIB = &H2000
Const PC_RESERVED = &H1
Const PC_EXPLICIT = &H2
Const PC_NOCOLLAPSE = &H4
'vbBlack &H0 Black
'vbRed   &HFF    Red
'vbGreen &HFF00  Green
'vbYellow    &HFFFF  Yellow
'vbBlue  &HFF0000    Blue
'vbMagenta   &HFF00FF    Magenta
'vbCyan  &HFFFF00    Cyan
'vbWhite &HFFFFFF    White
'Brown &H80FF&
'Private Enum InterfaceColors
'icMistyRose = &HE1E4FF
'icSlateGray = &H908070
'icDodgerBlue = &HFF901E
'icDeepSkyBlue = &HFFBF00
'icSpringGreen = &H7FFF00
'icForestGreen = &H228B22
'icGoldenrod = &H20A5DA
'icFirebrick = &H2222B2
'End Enum
Private maxrowz As Long, lblmax As Long, lblDoffset As Long, lblRoffset As Long, lastclicked As Long, lastcolour As Long
Private zhiddenloading As Boolean, ct As Long, ct1 As Long, ct2 As Long, Titlez(14) As String
Private Sub Form_Load()
Dim f As Form
ct = 0
For Each f In Forms
If f.Name = "Zhidden" Then ct = ct + 1
Next
If ct = 2 Then GoTo zerror


Dim a$, B$, Coltextwidth As Long, realpicsel As Long
cmdok.Default = True
zhiddenloading = True



Select Case zhiddnstatus
Case Is < 0
'gt(196) no of different DBpics
With Zhidden
.HelpContextID = 19
.Width = .Width + 2000
.Caption = "MyReels: Bitmap Finder Screen"
End With
setformpos Zhidden
HScroll1.Visible = False
VScroll1.Visible = False
With Cellz
.Width = Zhidden.Width
.Height = Zhidden.Height - 2 * cmdok.Height
.Top = 0
.Fontsize = 18
.ToolTipText = "Click to Sort"
End With
With cmdok
.FontBold = True
.Fontsize = 18
.Caption = "Previous"
If gt(196) < 71 Then Caption = "Back to Quote Thumbnails"
.Width = Cellz.Width - 120
.Cancel = False
.Top = Zhidden.Height - .Height - 350
End With

lblDoffset = gt(196) - 70
If lblDoffset < 0 Then lblDoffset = 0

lblCelltitle(7).Visible = False
For ct = 70 To 79
lblCell(ct).Visible = False
Next
For ct1 = 0 To 6
lblCelltitle(ct1).Visible = False
lblCell(ct1).Top = lblCell(ct1 + 1).Top - lblCell(ct1).Height
For ct = 0 To 9
ct2 = ct1 * 10 + ct
lblCell(ct2).Width = Cellz.Width / 7
lblCell(ct2).Left = ct1 * Cellz.Width / 7
If ct2 < gt(196) Then lblCell(ct2).Caption = Docaptions(ct2 + lblDoffset + 1)
Next
Next

HiddenTitle.Text = ""

With Zhidden
.Enabled = True
.Show
End With


Case 0
Soundd_Click
Musicc_Click
Thumbb_Click
'Natural state
Case 1
Cellz.Visible = False
HScroll1.Visible = False
With HiddenTitle
.Width = 4094
.Height = 7355
.Left = 0
.Top = 0
End With
With cmdok
.Width = 4185
.Top = 7350
.Left = 0
End With
Zhidden.Width = 4185
Zhidden.Height = 8315
setformpos Me

With Zhidden
.HelpContextID = 4
.Caption = "MyReels: Graphics Device Properties"         'Space(17)
.Enabled = True
.Show
End With

LoadInfo a$, Zhidden.hDC
On Error GoTo Errnoprinter
'Assume no printer if error
LoadInfo B$, Printer.hDC
HiddenTitle.Text = a$ & CStr(Chr$(13) + Chr$(10)) & B$



Case 2  'Hallfame

HiddenTitle.Visible = False

With lblzhidd(0)
.Font.Bold = True
.BackColor = &H20A5DA
.Fontsize = 20
.Visible = True
End With

If OpenDb(sDatabaseName, 1) = False Then
Zhidden.HelpContextID = 3
Spare.Visible = True
Set Spare.Picture = Nothing
lblzhidd(0).Caption = "Not Found"
zhiddnstatus = 0
Else

lastcolour = 0
lblDoffset = 0
lblRoffset = 0
lblzhidd(0).Caption = "Hall_Of_Fame"
Set dbsCurrent = gdbCurrentDB

Set rectemp = dbsCurrent.OpenRecordset("Hall")

maxrowz = rectemp.RecordCount


With Spare

Coltextwidth = .TextWidth("XXXXXXXXXXXXXXXXXXXX")

.Height = 1965
.Width = .ScaleWidth

For ct = -80 To .ScaleWidth Step 2040
.PaintPicture .Picture, ct, -230, 2040, 2200
Next


.Visible = True

End With

For ct = 1 To 3
lblzhidd(ct).Fontsize = 12
Next

On Error GoTo HallNotvalid


HScroll1.Max = 7
HScroll1.LargeChange = 7
HScroll1.SmallChange = 1

Cellz.Height = Cellz.Height - HScroll1.Height
    
    If maxrowz < 11 Then  'Blank Column
    lblmax = maxrowz - 1
    VScroll1.Visible = False
    Cellz.Left = 120
    HScroll1.Left = 100
    Else
    
    lblmax = 9
    With VScroll1
    .Max = maxrowz - 10
    .LargeChange = findhcmplusone(CInt(maxrowz) - 10, 10) / 10
    .SmallChange = 1
    End With
    End If

On Error GoTo zerror

'Set Cell Properties
For ct = 0 To lblmax
lblCell(70 + ct).Visible = False
For ct1 = 0 To 7
With lblCell(ct1 * 10 + ct)
.Alignment = 2
.fontname = "Courier New"
.Fontsize = 8
End With
Next
Next



If lblmax < 9 Then
For ct = maxrowz To 9
For ct1 = 0 To 7
lblCell(ct1 * 10 + ct).Visible = False
Next
Next
End If


PrintValues

Titlez(0) = "Player"          '2200
Titlez(1) = "Game"            '2200
Titlez(2) = "Start Cash"     '1000
Titlez(3) = "End Cash"       '1100
Titlez(4) = "Reason"          '"
Titlez(5) = "Date"            '"
Titlez(6) = "Spin Total"
Titlez(7) = "Reel Chg"
Titlez(8) = "Picture Chg"
Titlez(9) = "General Chg"
Titlez(10) = "Stats Reset"
Titlez(11) = "Game Replay"
Titlez(12) = "SVOG"
Titlez(13) = "VOG"
Titlez(14) = "MonteCarlo"


For ct = 0 To 6
lblCelltitle(ct).Alignment = 2
lblCelltitle(ct).Caption = Titlez(ct)
Next
lblCelltitle(7).Visible = False
lblCelltitle(7).Alignment = 2

End If  'Db found ?



With Zhidden
.HelpContextID = 3
.Caption = "MyReels: Hall_Of_Fame"        'Space(74)
.Enabled = True
.Show
End With

setformpos Me

Case Else  'Prize display



For ct = 1 To 14
If ct = zhiddnstatus - 9 Then
realpicsel = zhiddnstatus - 9
If intscatternumber = 2 Then
    If ct < intscattervec(2, 2) Then
        If ct >= intscattervec(1, 2) Then
        realpicsel = zhiddnstatus - 8
        If realpicsel = intscattervec(2, 2) Then realpicsel = realpicsel + 1
        End If
    ElseIf ct = 13 Then
    realpicsel = intscattervec(1, 2)
    ElseIf ct = 14 Then
    realpicsel = intscattervec(2, 2)
    ElseIf ct >= intscattervec(2, 2) Then
    realpicsel = zhiddnstatus - 7
    End If
ElseIf intscatternumber = 1 Then
    If ct = 14 Or (ct = 13 And Pokemach.imgprizethumb(13).Visible = False) Then
    realpicsel = intscattervec(1, 2)
    ElseIf ct >= intscattervec(1, 2) Then
    realpicsel = zhiddnstatus - 8
    End If
End If
Exit For
End If
Next



HiddenTitle.Left = -20000
lblDoffset = Pokemach.imgprizethumb(zhiddnstatus - 10).Top
lblRoffset = Pokemach.imgprizethumb(zhiddnstatus - 10).Left
For ct = 0 To 3
lblzhidd(ct).Caption = ""
Next
lblzhidd(0).BackColor = &H8FA2E0
lblzhidd(1).BackColor = &HCECAA2
lblzhidd(2).BackColor = &H8FA2E0
lblzhidd(3).BackColor = &HCECAA2
lblzhidd(4).BackColor = &H8FA2E0
'zhiddnstatus index of Pokemach.imgprizethumb


a$ = "Pays "

If sst(realpicsel, 1) = 1 Then
a$ = a$ & "Left to Right"
    If sst(realpicsel, 3) = 1 And sst(realpicsel, 4) = 1 Then
    a$ = a$ & " Right to Left Middle Threes"
    ElseIf sst(realpicsel, 3) = 1 Then
    a$ = a$ & " Right to Left"
    ElseIf sst(realpicsel, 4) = 1 Then
    a$ = a$ & " Middle Threes"
    End If
Else
a$ = a$ & "any"
End If


lblzhidd(0).Caption = a$

a$ = ""

a$ = "Pic " & realpicsel & " Substituting for"
ct = 0
For ct1 = 1 To 14
If substitute(realpicsel, ct1) = True Then
ct = ct + 1
a$ = a$ & " " & ct1
End If
Next


If ct = 0 Then a$ = a$ & " no pictures"

lblzhidd(1).Caption = a$
a$ = ""



If ct > 0 Then
a$ = a$ & " On reels "
For ct1 = 1 To 5
If reelcheck(realpicsel, ct1) = True Then a$ = a$ & (ct1) & " "
Next
End If

lblzhidd(2).Caption = a$
a$ = ""


For ct = 0 To 3
If gamespinsymbol(ct) = realpicsel Then
If ct = 0 Or ct = 1 Then
    a$ = freegamesettings(ct, 8) & " Free "
        If freegamesettings(ct, 9) = 0 Then
        a$ = a$ & "resettable Games, operating:"
        Else
        a$ = a$ & "Games, store of " & freegamesettings(ct, 9) & ", operating:"
        End If
    lblzhidd(3).Caption = a$
    a$ = ""
    If freegamesettings(ct, 1) = 2 Then
    a$ = "Any Three, Four or Five"
    ElseIf freegamesettings(ct, 1) = 1 Then
    a$ = a$ & "Left to Right"
        If freegamesettings(ct, 2) > 0 And freegamesettings(ct, 3) > 0 Then
        a$ = a$ & " Right to Left Middle Threes"
        ElseIf freegamesettings(ct, 2) > 0 Then
        a$ = a$ & " Right to Left"
        ElseIf freegamesettings(ct, 3) > 0 Then
        a$ = a$ & " Middle Threes"
        End If
    End If
ElseIf ct = 2 Or ct = 3 Then
    a$ = spinsettings(ct - 2, 14) & " Free "
        If spinsettings(ct - 2, 15) = 0 Then
        a$ = a$ & "resettable Spins, operating:"
        Else
        a$ = a$ & "Spins, store of " & spinsettings(ct - 2, 15) & ", operating:"
        End If
        lblzhidd(3).Caption = a$
    a$ = ""
    If spinsettings(ct - 2, 1) = 2 Then
    a$ = "Any Two"
    ElseIf spinsettings(ct - 2, 1) = 1 Then
    a$ = "(12)"
        If spinsettings(ct - 2, 3) > 0 And spinsettings(ct - 2, 2) > 0 Then
        a$ = a$ & " (23) (34) (45)"
        ElseIf spinsettings(ct - 2, 3) > 0 Then
        a$ = a$ & " (23) (34)"
        ElseIf spinsettings(ct - 2, 2) > 0 Then
        a$ = a$ & " (45)"
        End If
    End If
    If spinsettings(ct - 2, 4) = 2 Then
    a$ = a$ & " Any Three"
    ElseIf spinsettings(ct - 2, 4) = 1 Then
    a$ = a$ & " (123)"
        If spinsettings(ct - 2, 6) > 0 And spinsettings(ct - 2, 5) > 0 Then
        a$ = a$ & " (234) (345)"
        ElseIf spinsettings(ct - 2, 6) > 0 Then
        a$ = a$ & " (234)"
        ElseIf spinsettings(ct - 2, 5) > 0 Then
        a$ = a$ & " (345)"
        End If
    End If
    If spinsettings(ct - 2, 7) = 2 Then
    a$ = a$ & " Any Four"
    ElseIf spinsettings(ct - 2, 7) = 1 Then
    a$ = a$ & " (1234)"
    If spinsettings(ct - 2, 8) > 0 Then a$ = a$ & " (2345)"
    End If
End If
End If
Next

If lblzhidd(3).Caption = "" Then
lblzhidd(3).Caption = "No Free Games or Spins"
Else
lblzhidd(4).Caption = a$
End If

With Zhidden
.Width = resX * 4000
If resX = 1 Then
.Height = resX * 1300
Else
.Height = resX * 1250
End If
If lblDoffset < Pokemach.Height / 2 Then
.Top = lblDoffset - 250 * resX
Else
.Top = lblDoffset - resX * 700
End If
If lblRoffset < Pokemach.Width / 2 Then
.Left = lblRoffset - 50 * resX
Else
.Left = lblRoffset - resX * 3750
End If

.HelpContextID = 17
.Caption = "Picture Info"
.Enabled = True
End With


With Spare
.Move resX * -25, resX * -25, resX * 4000, resX * 1300
.Visible = True
Set .Picture = Nothing
End With

On Error GoTo zerror

For ct = 0 To 4
With lblzhidd(ct)
.Move 0, ct * resX * 200, resX * 4000, resX * 195
.fontname = "Courier New"
.Fontsize = 8
.Visible = True
End With
Next


Me.Show
End Select



zhiddenloading = False
Exit Sub
zerror:
ShowError
zhiddenloading = False
Exit Sub
Errnoprinter:
HiddenTitle.Text = a$ & CStr(Chr$(13) + Chr$(10)) & "No printer Installed"
zhiddenloading = False
Exit Sub
HallNotvalid:
Zhidden.HelpContextID = 3
ShowError
Spare.Visible = True
Set Spare.Picture = Nothing
lblzhidd(0).Caption = "Corrupted"
zhiddenloading = False
zhiddnstatus = 0
With Zhidden
.HelpContextID = 3
.Caption = "MyReels: Hall_Of_Fame"        'Space(74)
.Enabled = True
.Show
End With
setformpos Me
End Sub
Private Sub PrintValues()
Dim Vogval As Single, vogchgtmp As Boolean, zcolour As Long, startval As Long
ct1 = 0

If lblzhidd(0).Caption = "Not Found" Then Exit Sub

With rectemp


.MoveFirst
.Move lblDoffset

For ct = 0 To lblmax
zcolour = vbWhite
Vogval = ![VOG]     'set vogval
    If Vogval < 0 Then
    Vogval = -Vogval
    vogchgtmp = True
    Else
    vogchgtmp = False
    End If

Select Case lblRoffset
Case 0
Case 1
GoTo Lgame
Case 2
GoTo LStartcash
Case 3
GoTo LEndcash
Case 4
GoTo LReason
Case 5
GoTo LDate
Case 6
GoTo LSpintotal
Case Else
GoTo LReelChange
End Select

lblCell(ct).Caption = ![Player]
lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = 1


Lgame:
lblCell(ct + 10 * ct1).Caption = ![Game]
lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1


LStartcash:
lblCell(ct + 10 * ct1).Caption = ![Startcash]
Select Case ![Startcash]
Case 50, 100, 150
zcolour = &HFFFF00
Case 400, 450, 500
zcolour = &H80FF&
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1


LEndcash:
lblCell(ct + 10 * ct1).Caption = ![Endcash]
Select Case Val(lblCell(0).Caption) - Val(![Startcash])
Case Is > startval + 50
zcolour = &HFFFF00
Case Is > Val(![Startcash]) - 50
zcolour = &H80FF&
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1

LReason:
lblCell(ct + 10 * ct1).Caption = ![Reason]
Select Case Val(![Reason])
Case 1
zcolour = &H80FF&
Case 2
zcolour = &HFFFF00
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1


LDate:
lblCell(ct + 10 * ct1).Caption = ![Date]
lblCell(ct + 10 * ct1).BackColor = vbWhite
ct1 = ct1 + 1


LSpintotal:
lblCell(ct + 10 * ct1).Caption = ![Spintotal]
Select Case Val(![Spintotal])
Case Is >= 1000000
zcolour = &HFFFF00
Case Is <= 50
zcolour = &H80FF&
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1


If ct1 = 7 Then GoTo DoneRow

LReelChange:
lblCell(ct + 10 * ct1).Caption = ![ReelChange]
Select Case Val(![ReelChange])
Case Is > 1
zcolour = &H80FF&
Case Is = 0
zcolour = &HFFFF00
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour

ct1 = ct1 + 1
If lblRoffset = 1 Then GoTo DoneRow


lblCell(ct + 10 * ct1).Caption = ![SymbolChange]
Select Case Val(![SymbolChange])
Case Is > 1
zcolour = &H80FF&
Case Is = 0
zcolour = &HFFFF00
Case Else
zcolour = vbWhite
End Select
lblCell(ct + 10 * ct1).BackColor = zcolour

ct1 = ct1 + 1


lblCell(ct + 10 * ct1).Caption = ![GeneralChange]
Select Case Val(![GeneralChange])
Case Is > 1
zcolour = &H80FF&
Case Is = 0
zcolour = &HFFFF00
Case Else
zcolour = vbWhite
End Select
lblCell(ct + 10 * ct1).BackColor = zcolour

ct1 = ct1 + 1
If ct1 = 8 Then GoTo DoneRow

lblCell(ct + 10 * ct1).Caption = ![StatsResets]
Select Case Val(![StatsResets])
Case Is > 1
zcolour = &H80FF&
Case Is = 0
zcolour = &HFFFF00
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1
If ct1 = 8 Then GoTo DoneRow

lblCell(ct + 10 * ct1).Caption = ![GameReplays]
Select Case Val(![GameReplays])
Case Is > 2
zcolour = &H80FF&
Case Is = 0
zcolour = &HFFFF00
Case Else
zcolour = vbWhite
End Select

lblCell(ct + 10 * ct1).BackColor = zcolour
ct1 = ct1 + 1
If ct1 = 8 Then GoTo DoneRow

lblCell(ct + 10 * ct1).Caption = ![SVOG]
If Vogval <> 0 Then
Select Case Val(lblCell(ct + 10 * ct1).Caption) / Vogval
Case Is >= 1.1
zcolour = &HFFFF00
Case Is <= 0.9
zcolour = &H80FF&
Case Else
zcolour = vbWhite
End Select
Else
zcolour = vbWhite
End If
lblCell(ct + 10 * ct1).BackColor = zcolour

ct1 = ct1 + 1
If ct1 = 8 Then GoTo DoneRow

lblCell(ct + 10 * ct1).Caption = Vogval
Select Case Vogval
Case 0
zcolour = vbWhite
Case Is < 80
zcolour = &HFFFF00
Case Is >= 100
zcolour = &H80FF&
Case Else
zcolour = vbWhite
End Select

'VOG has been changed in that game
If vogchgtmp = True Then zcolour = &H80FF&
lblCell(ct + 10 * ct1).BackColor = zcolour

ct1 = ct1 + 1
If ct1 = 8 Then GoTo DoneRow

lblCell(ct + 10 * ct1).Caption = ![MonteCarlo]
Select Case Val(![MonteCarlo])
Case Is > 5000000
zcolour = &HFFFF00
Case 500000, 1000000, 5000000
zcolour = vbWhite
Case 0
    If Vogval > 0 Then
    zcolour = &HFFFF00
    Else
    zcolour = vbWhite
    End If
Case Else
zcolour = &H80FF&
End Select
lblCell(ct + 10 * ct1).BackColor = zcolour

DoneRow:

.MoveNext
ct1 = 0
zcolour = vbWhite


Next


'.FormatString = "^Player|^Game|>Startcash|^VOG|^Reason|^SpinTotal|^EndCash|^Date|^ReelChange|^SymbolChange|^GeneralChange|^GameReplays|^SVOG|"
End With


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If zhiddnstatus > 9 Then
Zhidden.Enabled = False
Unload Zhidden
Set Zhidden = Nothing
Pokemach.imgprizethumb(zhiddnstatus - 10).Left = lblRoffset
zhiddnstatus = 0
End If
End Sub

Private Sub VScroll1_Change()
If lastcolour > 0 Then
lblCell(lastclicked).BackColor = lastcolour
Cellcolour.Enabled = False
lastcolour = 0
End If
If maxrowz - VScroll1.Value > 10 Then
lblDoffset = VScroll1.Value
Else
lblDoffset = maxrowz - 10
End If
PrintValues
End Sub
Private Sub HScroll1_Change()
Dim actRoffset As Long
Cellz.Width = 9820

If lastcolour > 0 Then
lblCell(lastclicked).BackColor = lastcolour
Cellcolour.Enabled = False
lastcolour = 0
End If

lblRoffset = HScroll1.Value
Select Case lblRoffset
Case 0
For ct = 0 To 6
For ct1 = 0 To lblmax
    'Adds to 9800
    With lblCell(10 * ct + ct1)
    Select Case ct
    Case 0
    .Left = ct * 2200
    .Width = 2200
    Case 1
    .Left = ct * 2200
    .Width = 2195 'Twip crap
    Case 2
    .Left = 4400
    .Width = 1000
    Case Else
    .Left = 5400 + (ct - 3) * 1100
    If ct = 4 Then  'twips crap
    .Width = 1105
    Else
    .Width = 1100
    End If
    End Select
    End With
Next
Next
For ct1 = 0 To lblmax
lblCell(70 + ct1).Visible = False
Next
Case 1
For ct = 0 To 6
For ct1 = 0 To lblmax
    'Adds to 9800
    With lblCell(10 * ct + ct1)
    Select Case ct
    Case 0
    .Left = 0
    .Width = 2500
    Case 1
    .Left = 2500
    .Width = 1100
    Case Else
    .Left = 3600 + (ct - 2) * 1240
    If ct = 3 Then
    .Width = 1235
    Else
    .Width = 1240
    End If
    End Select
    End With
Next
Next
For ct1 = 0 To lblmax
lblCell(70 + ct1).Visible = False
Next
Case Else
    For ct = 0 To 7
    For ct1 = 0 To lblmax
    With lblCell(10 * ct + ct1)
    'Adds to 9800
    .Left = 1225 * ct
    Select Case ct
    Case 1, 4
    .Width = 1220
    Case Else
    .Width = 1225
    End Select
    End With
    Next
    Next
For ct1 = 0 To lblmax
lblCell(70 + ct1).Visible = True
Next
End Select


If lblRoffset + 8 > 14 Then
actRoffset = 7
Else
actRoffset = lblRoffset
End If
For ct = 0 To 6
With lblCelltitle(ct)
.Left = lblCell(10 * ct).Left
.Width = lblCell(10 * ct).Width
.Caption = Titlez(ct + actRoffset)
End With
Next


With lblCelltitle(7)
If actRoffset > 1 Then
.Visible = True
.Width = lblCell(70).Width
.Left = lblCell(70).Left
.Caption = Titlez(7 + actRoffset)
Else
.Visible = False
End If
End With


PrintValues
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If zhiddnstatus < 0 Then
    If KeyCode = vbKeyEscape Then
    Zhidden.Enabled = False
    Unload Zhidden
    Set Zhidden = Nothing
    zhiddnstatus = 0
    End If
ElseIf zhiddnstatus = 2 Then
Dim Hally As Long, Hallx As Long
Cellcolour.Enabled = False
If lastcolour > 0 Then lblCell(lastclicked).BackColor = lastcolour
On Error Resume Next
Hally = VScroll1.Value
Select Case KeyCode
Case vbKeyPageUp, vbKeyNumpad8
If VScroll1.Visible = False Then Exit Sub
Hally = Hally - VScroll1.LargeChange
If Hally < 1 Then Hally = 0
VScroll1.Value = Hally
Case vbKeyPageDown, vbKeyNumpad2
If VScroll1.Visible = False Then Exit Sub
Hally = Hally + VScroll1.LargeChange
If Hally > maxrowz Then Hally = maxrowz
VScroll1.Value = Hally
Case vbKeyDelete, vbKeyNumpad4
HScroll1.Value = 0
Case vbKeyEnd, vbKeyNumpad6
HScroll1.Value = 7
Case vbKeyHome, vbKeyNumpad5
HScroll1.Value = 0
If VScroll1.Visible = False Then Exit Sub
VScroll1.Value = 0
Case vbKeyLeft
Hallx = HScroll1.Value  'up fires hscroll if maxrowz<10
Hallx = Hallx - 1
If Hallx < 0 Then Hallx = 0
HScroll1.Value = Hallx
Case vbKeyRight
Hallx = HScroll1.Value  'up fires hscroll if maxrowz<10
Hallx = Hallx + 1
If Hallx > 7 Then Hallx = 7
HScroll1.Value = Hallx
Case vbKeyDown
If VScroll1.Visible = False Then Exit Sub
Hally = Hally + 1
If Hally > maxrowz Then Hally = maxrowz
VScroll1.Value = Hally
Case vbKeyUp
If VScroll1.Visible = False Then Exit Sub
Hally = Hally - 1
If Hally < 1 Then Hally = 0
VScroll1.Value = Hally
End Select
End If
End Sub
Private Sub lblzhidd_Click(Index As Integer)
If zhiddnstatus > 9 Then
Pokemach.imgprizethumb(zhiddnstatus - 10).Left = lblRoffset
Zhidden.Enabled = False
Unload Zhidden
Set Zhidden = Nothing
zhiddnstatus = 0
End If
End Sub
Private Sub HiddenTitle_Change()
If zhiddenloading = False Then cmdOK_Click
End Sub
Public Sub Musicc_Click()
ct = 0
Select Case zhiddenloading

Case True
If gt(185) = 0 Then
Musicc.Caption = "Music On [M]"
Else
Musicc.Caption = "Music Off [M]"
End If
Case Else

If Musicc.Caption = "Music Off [M]" Then
Midichg ct
If gt(185) = 0 Then Musicc.Caption = "Music On [M]"
Else
  If Midichg <> gt(185) Then
  gt(185) = 0
  Musicc.Caption = "Music On [M]"
  Exit Sub
  Else
  If gt(185) > 0 Then Musicc.Caption = "Music Off [M]"
  End If
End If

End Select
Pokemach.midiplay.Enabled = CBool(gt(185) > 0)
End Sub
Public Sub Soundd_Click()

Select Case zhiddenloading

Case True
If gt(186) < 2 Then
Soundd.Caption = "Sound On [S]"
Else
Soundd.Caption = "Sound Off [S]"
End If
Case Else
If gt(186) > 1 Then
gt(186) = gt(186) - 2
Soundd.Caption = "Sound On [S]"
Else
gt(186) = gt(186) + 2
Soundd.Caption = "Sound Off [S]"
End If
End Select
End Sub
Public Sub Thumbb_Click()
Select Case zhiddenloading

Case True
Select Case gt(186)
Case 0, 2
Thumbb.Caption = "QuoteSound On [Q]"
Case Else
Thumbb.Caption = "QuoteSound Off [Q]"
End Select
Case Else
Select Case gt(186)
Case 1, 3
gt(186) = gt(186) - 1
Thumbb.Caption = "QuoteSound On [Q]"
Case Else
gt(186) = gt(186) + 1
Thumbb.Caption = "QuoteSound Off [Q]"
End Select
End Select

End Sub
Private Sub About_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
.Frmaboutt
End With
End Sub
Private Sub ChangeBet_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
.Cyclemultibets
.activatectrls
End With
End Sub
Private Sub Changeln1_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
PlaySndF Stringvars(27)
gt(153) = 0
.indicln
.activatectrls
End With
End Sub
Private Sub Changeln2_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
PlaySndF Stringvars(27)
gt(153) = 1
.indicln
.activatectrls
End With
End Sub
Private Sub Changeln3_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
PlaySndF Stringvars(27)
gt(153) = 2
.indicln
.activatectrls
End With
End Sub
Private Sub Spinner_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
chgchtr
.Preparetospin
End With
End Sub
Private Sub Cheat_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
chgchtr True
.Preparetospin
End With
End Sub
Private Sub Configurationn_Click()
DoEvents
Pokemach.Configurationn
Unload Zhidden
Set Zhidden = Nothing
End Sub
Private Sub ChangeDirectory_Click()
DoEvents
Pokemach.Changedir
Unload Zhidden
Set Zhidden = Nothing
End Sub
Private Sub Quitt_Click()
DoEvents
Pokemach.Quitt
Unload Zhidden
Set Zhidden = Nothing
End Sub
Private Sub Helpp_Click()
PlaySndF App.Path & "\help.wav"
Shell ("winhlp32.exe  -N 1 " & App.Path & "\MyReels.hlp"), vbNormalFocus
End Sub
Private Sub Titleadjust_Click()
With Pokemach
If .waitimer.Enabled = True Or .Prizemeter.Enabled = True Or .gamespinwait.Enabled = True Or .jackpot.Enabled = True Or .timoneyback.Enabled = True Or waitimmarker > 0 Then Exit Sub
.TA_Click
End With
End Sub
Private Sub Cellz_Click()
If zhiddnstatus < 0 Then
If Cellz.ToolTipText = "sorted" Then Exit Sub
lblDoffset = gt(196) - 70
If lblDoffset < 0 Then lblDoffset = 0
Cellz.ToolTipText = "sorted"
lblCell(0).Caption = Docaptions(0, True) 'sort picnames
For ct1 = 0 To 6
For ct = 0 To 9
ct2 = ct1 * 10 + ct
lblCell(ct2).Width = Cellz.Width / 7
lblCell(ct2).Left = ct1 * Cellz.Width / 7
If ct2 < gt(196) Then lblCell(ct2).Caption = Docaptions(ct2 + lblDoffset + 1, , Cellz.ToolTipText = "sorted")
Next
Next
End If
End Sub
Private Sub cmdOK_Click()
If zhiddnstatus > 9 Then
'In case of Enter or Esc
Pokemach.imgprizethumb(zhiddnstatus - 10).Left = lblRoffset
ElseIf zhiddnstatus < 0 Then 'for quotebrs db
    lblDoffset = lblDoffset - 70
    If lblDoffset > -70 Then
        If lblDoffset < 0 Then
        lblDoffset = 0
        cmdok.Caption = "Back to Quote Thumbnails"
        End If
        For ct1 = 0 To 6
        For ct = 0 To 9
        ct2 = ct1 * 10 + ct
        lblCell(ct2).Left = ct1 * Cellz.Width / 7
        lblCell(ct2).Caption = Docaptions(ct2 + lblDoffset + 1, , Cellz.ToolTipText = "sorted")
        Next
        Next
        Exit Sub
    End If
Else
    If zhiddnstatus = 2 Then
    Set rectemp = Nothing
    Set dbsCurrent = Nothing
    killdb sDatabaseName
    End If
    frmAbout.Enabled = True
End If


Zhidden.Enabled = False
Unload Zhidden
Set Zhidden = Nothing
zhiddnstatus = 0
End Sub
Private Sub Cellcolour_Timer()
lblCell(lastclicked).BackColor = lastcolour
Cellcolour.Enabled = False
End Sub
Private Sub lblCell_Click(Index As Integer)
If zhiddnstatus < 0 Then
zhiddnstatus = -(Index + lblDoffset + 1) 'TEMPORARY zhiddenstatus
If Cellz.ToolTipText = "sorted" Then Getzhiddnstat
Zhidden.Enabled = False
Unload Zhidden
Set Zhidden = Nothing
Else
With Cellcolour
If .Enabled = True Then .Enabled = False    'Reset
.Enabled = True
End With
If lastcolour > 0 Then lblCell(lastclicked).BackColor = lastcolour
lastcolour = lblCell(Index).BackColor
lastclicked = Index
lblCell(Index).BackColor = vbYellow
End If
End Sub
Private Sub LoadInfo(a$, usehDC&)
    #If Win32 Then
    Dim r As Long
    Dim nhDC As Long
    #Else
    Dim r As Integer
    Dim nhDC As Integer
    #End If
    nhDC = usehDC
    Dim crlf$
    
    crlf$ = Chr$(13) + Chr$(10)

    r = GetDeviceCaps(nhDC, TECHNOLOGY)
    If r And DT_RASPRINTER Then a$ = "Raster Printer"
    If r And DT_RASDISPLAY Then a$ = "Raster Display"
    ' You can detect other technology types here - see the
    ' GetDeviceCaps function description for technology types
    If a$ = "" Then a$ = "Other technology"
    a$ = a$ + crlf$
    a$ = a$ + "X,Y Dimensions in pixels:" + Str$(GetDeviceCaps(nhDC, HORZRES)) + "," + Str$(GetDeviceCaps(nhDC, VERTRES)) + crlf$
    a$ = a$ + "X,Y Pixels/Logical Inch:" + Str$(GetDeviceCaps(nhDC, LOGPIXELSX)) + "," + Str$(GetDeviceCaps(nhDC, LOGPIXELSY)) + crlf$
    a$ = a$ + "Bits/Pixel:" + Str$(GetDeviceCaps(nhDC, BITSPIXEL)) + crlf$
    a$ = a$ + "Color Planes:" + Str$(GetDeviceCaps(nhDC, PLANES)) + crlf$
    a$ = a$ + "Color Table Entries:" + Str$(GetDeviceCaps(nhDC, NUMCOLORS)) + crlf$
    a$ = a$ + "Aspect X,Y,XY:" + Str$(GetDeviceCaps(nhDC, ASPECTX)) + "," + Str$(GetDeviceCaps(nhDC, ASPECTY)) + "," + Str$(GetDeviceCaps(nhDC, ASPECTXY)) + crlf$
    r = GetDeviceCaps(nhDC, RASTERCAPS)
    a$ = a$ + crlf$ + "Device Capabilities:" + crlf$
    If r And RC_BANDING Then a$ = a$ + "Banding" + crlf$
    If r And RC_BIGFONT Then a$ = a$ + "Fonts >64K" + crlf$
    If r And RC_BITBLT Then a$ = a$ + "BitBlt" + crlf$
    If r And RC_BITMAP64 Then a$ = a$ + "Bitmaps >64k" + crlf$
    If r And RC_DI_BITMAP Then a$ = a$ + "Device Independent Bitmaps" + crlf$
    If r And RC_DIBTODEV Then a$ = a$ + "DIB to device" + crlf$
    If r And RC_FLOODFILL Then a$ = a$ + "Flood fill" + crlf$
    If r And RC_SCALING Then a$ = a$ + "Scaling" + crlf$
    If r And RC_STRETCHBLT Then a$ = a$ + "StretchBlt" + crlf$
    If r And RC_STRETCHDIB Then a$ = a$ + "StretchDIB" + crlf$
End Sub
