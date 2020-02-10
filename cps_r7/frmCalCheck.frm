VERSION 5.00
Begin VB.Form frmCalCheck 
   Caption         =   "Calibration Check Function"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmCalCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   6480
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   104
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "CalCheck History "
      DisabledPicture =   "frmCalCheck.frx":1CCA
      DownPicture     =   "frmCalCheck.frx":290C
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   4070
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":354E
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "Save Values"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   9600
      TabIndex        =   101
      Top             =   2880
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   9600
      TabIndex        =   100
      Top             =   3165
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   9600
      TabIndex        =   99
      Top             =   3450
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   9600
      TabIndex        =   98
      Top             =   3735
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   9600
      TabIndex        =   97
      Top             =   4020
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   9600
      TabIndex        =   96
      Top             =   4305
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   9600
      TabIndex        =   95
      Top             =   4590
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   9600
      TabIndex        =   94
      Top             =   4875
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   9600
      TabIndex        =   93
      Top             =   5160
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   9600
      TabIndex        =   92
      Top             =   5445
      Width           =   900
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   9600
      TabIndex        =   91
      Top             =   5730
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   4080
      TabIndex        =   86
      Top             =   5685
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   4080
      TabIndex        =   85
      Top             =   5400
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   4080
      TabIndex        =   84
      Top             =   5115
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   4080
      TabIndex        =   83
      Top             =   4830
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   82
      Top             =   4545
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   81
      Top             =   4260
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   4080
      TabIndex        =   80
      Top             =   3975
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   79
      Top             =   3690
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   78
      Top             =   3405
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   77
      Top             =   3120
      Width           =   900
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   76
      Top             =   2835
      Width           =   900
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   75
      Top             =   3690
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   2280
      TabIndex        =   74
      Top             =   5685
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   2280
      TabIndex        =   73
      Top             =   5400
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   2280
      TabIndex        =   72
      Top             =   5115
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   71
      Top             =   4830
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   2280
      TabIndex        =   70
      Top             =   4545
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   69
      Top             =   4260
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   68
      Top             =   3975
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   67
      Top             =   3405
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   66
      Top             =   3120
      Width           =   780
   End
   Begin VB.TextBox txtDesired 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   65
      Top             =   2835
      Width           =   780
   End
   Begin VB.TextBox txtActual 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   7080
      TabIndex        =   63
      Top             =   2835
      Width           =   900
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      DisabledPicture =   "frmCalCheck.frx":4190
      DownPicture     =   "frmCalCheck.frx":4DD2
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":5A14
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Save Values"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      DisabledPicture =   "frmCalCheck.frx":6656
      DownPicture     =   "frmCalCheck.frx":7298
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   5085
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":7EDA
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6120
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      CausesValidation=   0   'False
      DisabledPicture =   "frmCalCheck.frx":8B1C
      DownPicture     =   "frmCalCheck.frx":975E
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   3055
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":A3A0
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Print a Hard Copy"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Text            =   "sec"
      Top             =   4575
      Width           =   945
   End
   Begin VB.CommandButton cmdDown 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":AFE2
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Down"
      Top             =   3420
      Width           =   930
   End
   Begin VB.CommandButton cmdUp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":107C4
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Up"
      Top             =   2880
      Width           =   930
   End
   Begin VB.CommandButton cmdAuto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCalCheck.frx":15FA6
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Run Automatic Check Cycle"
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   930
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   11
      Left            =   1320
      TabIndex        =   13
      Top             =   5715
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   10
      Left            =   1320
      TabIndex        =   12
      Top             =   5430
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   9
      Left            =   1320
      TabIndex        =   11
      Top             =   5145
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   8
      Left            =   1320
      TabIndex        =   10
      Top             =   4860
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   1320
      TabIndex        =   9
      Top             =   4575
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   1320
      TabIndex        =   8
      Top             =   4290
      Width           =   255
   End
   Begin VB.Frame frmMfcId 
      ForeColor       =   &H80000002&
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdMfcUp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCalCheck.frx":1B788
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Up"
         Top             =   1080
         Width           =   450
      End
      Begin VB.CommandButton cmdMfcDn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCalCheck.frx":1BACA
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Down"
         Top             =   1080
         Width           =   450
      End
      Begin VB.CommandButton cmdStnUp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCalCheck.frx":1BE0C
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Up"
         Top             =   840
         Width           =   450
      End
      Begin VB.CommandButton cmdStnDn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCalCheck.frx":1C14E
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Down"
         Top             =   840
         Width           =   450
      End
      Begin VB.TextBox txtDateTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "dts"
         Top             =   240
         Width           =   5505
      End
      Begin VB.TextBox txtMethod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "method"
         Top             =   1500
         Visible         =   0   'False
         Width           =   5505
      End
      Begin VB.TextBox txtStation 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "stn"
         Top             =   780
         Width           =   5505
      End
      Begin VB.TextBox txtMFC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "mfc"
         Top             =   1140
         Width           =   5505
      End
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   4
      Top             =   4005
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   1320
      TabIndex        =   3
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   1320
      Picture         =   "frmCalCheck.frx":1C490
      TabIndex        =   2
      Top             =   3435
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   1
      Top             =   3150
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1320
      MaskColor       =   &H000000FF&
      Picture         =   "frmCalCheck.frx":1C7D2
      TabIndex        =   0
      Top             =   2865
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   960
      Top             =   5760
   End
   Begin VB.Label lblOutput 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   9600
      TabIndex        =   102
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   6960
      TabIndex        =   64
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   59
      Top             =   2835
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   58
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   57
      Top             =   3405
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   56
      Top             =   3690
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   55
      Top             =   3975
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   54
      Top             =   4260
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   53
      Top             =   4545
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   52
      Top             =   4830
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   51
      Top             =   5115
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   50
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   11
      Left            =   1680
      TabIndex        =   49
      Top             =   5685
      Width           =   495
   End
   Begin VB.Label lblCurrPoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Point"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   48
      Top             =   2595
      Width           =   495
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      Caption         =   "Check Current Calibration"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   120
      TabIndex        =   47
      Top             =   2280
      Width           =   5895
   End
   Begin VB.Label lblDesireds 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desired"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2280
      TabIndex        =   46
      Top             =   2595
      Width           =   780
   End
   Begin VB.Label lblActuals 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4080
      TabIndex        =   45
      Top             =   2595
      Width           =   900
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   11
      Left            =   3060
      TabIndex        =   44
      Top             =   5685
      Width           =   1020
   End
   Begin VB.Label lblCalibrateds 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calibrated"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3060
      TabIndex        =   43
      Top             =   2595
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   3060
      TabIndex        =   42
      Top             =   2835
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   3060
      TabIndex        =   41
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   3060
      TabIndex        =   40
      Top             =   3405
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   3060
      TabIndex        =   39
      Top             =   3690
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   3060
      TabIndex        =   38
      Top             =   3975
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   3060
      TabIndex        =   37
      Top             =   4260
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   7
      Left            =   3060
      TabIndex        =   36
      Top             =   4545
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   8
      Left            =   3060
      TabIndex        =   35
      Top             =   4830
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   9
      Left            =   3060
      TabIndex        =   34
      Top             =   5115
      Width           =   1020
   End
   Begin VB.Label lblCalibrated 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   10
      Left            =   3060
      TabIndex        =   33
      Top             =   5400
      Width           =   1020
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   11
      Left            =   4980
      TabIndex        =   32
      Top             =   5685
      Width           =   900
   End
   Begin VB.Label lblPercDiffs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% Diff"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4980
      TabIndex        =   31
      Top             =   2595
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   4980
      TabIndex        =   30
      Top             =   2835
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   4980
      TabIndex        =   29
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   4980
      TabIndex        =   28
      Top             =   3405
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   4
      Left            =   4980
      TabIndex        =   27
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   4980
      TabIndex        =   26
      Top             =   3975
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   6
      Left            =   4980
      TabIndex        =   25
      Top             =   4260
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   7
      Left            =   4980
      TabIndex        =   24
      Top             =   4545
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   8
      Left            =   4980
      TabIndex        =   23
      Top             =   4830
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   9
      Left            =   4980
      TabIndex        =   22
      Top             =   5115
      Width           =   900
   End
   Begin VB.Label lblPercDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   10
      Left            =   4980
      TabIndex        =   21
      Top             =   5400
      Width           =   900
   End
   Begin VB.Label lblDelay 
      Alignment       =   2  'Center
      Caption         =   "Delay"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   945
   End
End
Attribute VB_Name = "frmCalCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUMROWS = 11                                  ' Number of calibration check rows
Private aryDesFlow(NUMROWS) As Single               ' the desired flow (from the input table)
Private aryActualFlowSLPM(NUMROWS) As Single        ' the actual flow in Calibrated SLPM (from the input table)
Private aryActualFlowUncal(NUMROWS) As Single       ' the actual flow in Uncalibrated SLPM (from the input table)
Private aryPercDiff(NUMROWS) As Single              ' the difference between desired and actual flow in percent (from the input table)
Private bChanged(NUMROWS) As Boolean                ' the actual flow has been changed since the last save (or load)
Private bAllChanged As Boolean                      ' all the actual flows have been changed since the last save (or load)

Private aryOutputFS(NUMROWS) As Single              ' the output value found using Newton's method

Private SelectedFunc As Integer                     ' the Station Analog Function index for the selected mass flow controller (from frmMassFlowCal)
Private SelectedMFC As Integer                      ' the selected mass flow controller (from frmMassFlowCal)
Private SelectedRow As Integer                      ' the selected calibration date entry row
Private SelectedStation As Integer                  ' the selected station (from frmMassFlowCal)
Private MfcSpan As Single                           ' the selected Mfc's span in Engr Units (max - min)
Private MfcMin As Single                            ' the selected Mfc's minimum value in Engr Units

Private AutoCycleOn As Boolean
Private AutoStepNext As Date
Private AutoStepInterval As Long
Private AutoCycleSP As Single

Private Curr_MfcCal As MfcCalibration
Private curDTS As Date
Private CalReadOnly As Boolean
Private bReadyToSaveSetup As Boolean
Private bReadyToSave As Boolean
Private bHistoryExists As Boolean
Private bFormLoaded As Boolean           ' Flag - Whether the form has loaded(can't SetFocus on a TextBox until form is Loaded


Public Sub SetupCalCheck(ByVal iStn As Integer, ByVal iMfc As Integer)
'
    SelectedMFC = iMfc
    SelectedStation = iStn
    txtStation.text = "Station " & Format(SelectedStation, "#")
    ' Mass Flow Controllers
    txtMFC.text = Mfc_FunDef(SelectedMFC).desc
    Select Case SelectedMFC
        Case MFCBUTANE
            txtStation.ForeColor = DK2GREEN
            SelectedFunc = asButaneFlow
        Case MFCNITROGEN
            txtStation.ForeColor = DK2GREEN
            SelectedFunc = asNitrogenFlow
        Case MFCPURGEAIR
            txtStation.ForeColor = DKPURPLE
            SelectedFunc = asPurgeAirFlow
        Case MFCORVRBUT
            txtStation.ForeColor = DKGREEN
            SelectedFunc = asButaneORVRFlow
        Case MFCORVRNIT
            txtStation.ForeColor = DKGREEN
            SelectedFunc = asNitrogenORVRFlow
'        Case MFCORVRPRG
'            txtStation.ForeColor = DKPURPLE
'            Selectedfunc = asButaneFlow
        Case MFCLIVEFUEL
            txtStation.ForeColor = MEDORANGE
            SelectedFunc = asLiveFuelVaporFlow
        Case MFCORVRLIVE
            txtStation.ForeColor = MEDORANGE
            SelectedFunc = asLiveFuelVaporFlow
        Case Else
            ' Close the form and switch back to the Mass Flow Calibration form
            Unload Me
            Set frmCalCheck = Nothing
            frmMfcCal.Enabled = True
            frmMfcCal.txtMsg = vbCrLf & "Invalid MFC for Cal Check"
'            frmMfcCal.SetFocus
    End Select
    txtMFC.ForeColor = txtStation.ForeColor
    txtDateTime.ForeColor = txtStation.ForeColor
    MfcMin = Stn_AIO(SelectedStation, SelectedFunc).EuMin
    MfcSpan = Stn_AIO(SelectedStation, SelectedFunc).EuMax - MfcMin
End Sub

Private Sub cmdAuto_Click()
    ' toggle auto cycling
    If Not AutoCycleOn Then
        ' not cycling; start cycling
        AutoStep_Init
    Else
        ' cycling; stop cycling
        AutoStep_Done
    End If
End Sub

Private Sub cmdHistory_Click()
    With frmCalCheckHistory
        .Show
        .SelectedStation = SelectedStation
        .SelectedMFC = SelectedMFC
        .SelectedCalCheck = 0
        ' Update the Display
        .UpdateMfcSelection
        .DisplayMfcAll
        .Refresh
    End With
End Sub

Private Sub cmdMfcDn_Click()
    If (Not CalReadOnly) Then CalValves False
    SelectedMFC = SelectedMFC - 1
    If SelectedMFC < 0 Then SelectedMFC = MAXMFC
    SetupCalCheck SelectedStation, SelectedMFC
    SelectedRow = 1
    UpdateSelection
    If (Not CalReadOnly) Then CalValves True
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
'    If bFormLoaded Then txtDesired(SelectedRow).SetFocus
End Sub

Private Sub cmdMFCUp_Click()
    If (Not CalReadOnly) Then CalValves False
    SelectedMFC = SelectedMFC + 1
    If SelectedStation > MAXMFC Then SelectedMFC = 0
    SetupCalCheck SelectedStation, SelectedMFC
    SelectedRow = 1
    UpdateSelection
    If (Not CalReadOnly) Then CalValves True
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
'    If bFormLoaded Then txtDesired(SelectedRow).SetFocus
End Sub

Private Sub cmdSave_Click()
    SaveCurrCalCheckData SelectedStation, SelectedMFC
    ResetChanged
    bReadyToSave = False
    bReadyToSaveSetup = False
    bHistoryExists = True
End Sub

Private Sub cmdStnDn_Click()
    If (Not CalReadOnly) Then CalValves False
    SelectedStation = SelectedStation - 1
    If SelectedStation < 1 Then SelectedStation = LAST_STN
    SetupCalCheck SelectedStation, SelectedMFC
    SelectedRow = 1
    UpdateSelection
    If (Not CalReadOnly) Then CalValves True
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
    If bFormLoaded Then txtDesired(SelectedRow).SetFocus
End Sub

Private Sub cmdStnUp_Click()
    If (Not CalReadOnly) Then CalValves False
    SelectedStation = SelectedStation + 1
    If SelectedStation > LAST_STN Then SelectedStation = 1
    SetupCalCheck SelectedStation, SelectedMFC
    SelectedRow = 1
    UpdateSelection
    If (Not CalReadOnly) Then CalValves True
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
    If bFormLoaded Then txtDesired(SelectedRow).SetFocus
End Sub

Private Sub ResetChanged()
Dim idx As Integer
    For idx = 1 To NUMROWS
        bChanged(idx) = False
    Next idx
    bAllChanged = False
    curDTS = Now()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub cmdCancel_Click()
    ' Close the form and switch to the Mass Flow Calibration form
    Xit
End Sub

Private Sub cmdDown_Click()
    ' Selects the row below the currently selected row in the
    ' MFC calibration table, for editing
    
    txtDesired(SelectedRow).text = Format(Curr_MfcCal.PointData(SelectedRow).RawValue, "####0.0000")
    If SelectedRow >= NUMROWS Then Exit Sub
    
    SelectedRow = SelectedRow + 1
    UpdateSelection
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
    txtDesired(SelectedRow).SetFocus
    If ((SelectedRow = 1) And bReadyToSaveSetup) Then
        bReadyToSave = True
    ElseIf (SelectedRow = Curr_MfcCal.NumPoints) Then
        bReadyToSaveSetup = True
    End If
End Sub

Private Sub cmdPrint_Click()
    ' Print the calibration check data
    PrintData
End Sub

Private Sub cmdUp_Click()
    ' Selects the row above the currently selected row in the
    ' MFC calibration table, for editing
    
    txtDesired(SelectedRow).text = Format(Curr_MfcCal.PointData(SelectedRow).RawValue, "####0.0000")
    If SelectedRow <= 1 Then Exit Sub
    
    SelectedRow = SelectedRow - 1
    UpdateSelection
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
    txtDesired(SelectedRow).SetFocus
    If ((SelectedRow = 1) And bReadyToSaveSetup) Then
        bReadyToSave = True
    ElseIf (SelectedRow = Curr_MfcCal.NumPoints) Then
        bReadyToSaveSetup = True
    End If
End Sub

Private Sub Form_Load()
    Dim row As Integer
    
    bFormLoaded = False

    'Disable the MFC Calibration form
    frmMfcCal.Enabled = False
    
    ' temp setup
    SelectedStation = 1
    SelectedMFC = 0
    SelectedFunc = asButaneFlow
    txtStation.text = "Station " & Format(SelectedStation, "#")
    txtMFC.text = "Butane MFC"
    txtStation.ForeColor = DKGREEN
    txtMFC.ForeColor = txtStation.ForeColor
    txtDateTime.ForeColor = txtStation.ForeColor
    MfcMin = Stn_AIO(SelectedStation, SelectedFunc).EuMin
    MfcSpan = Stn_AIO(SelectedStation, SelectedFunc).EuMax - MfcMin
    curDTS = Now()
    
    ' enable the cells
    For row = 1 To NUMROWS
        txtDesired(row).Enabled = True
        txtDesired(row).BackColor = Entry_BackColor
        txtActual(row).Enabled = True
        txtActual(row).BackColor = Entry_BackColor
    Next row
    
    ' Current DateTime
    txtDateTime.text = Format(Now(), "YYYY-MMM-DD   HH:MM:SS")
    ' Default Delay
    txtSeconds.text = Format((2 * MFC_Settle_Time), "####0")
    txtSeconds.ForeColor = DK3RED

    ' Set up the form for its initial settings
    For row = 1 To NUMROWS
        txtDesired(row).text = "0.0"
    Next row
    
    SelectedRow = 1
 
    DisplayData
    txtDesired(SelectedRow).TabIndex = 0
    bReadyToSaveSetup = False
    bReadyToSave = False
    bFormLoaded = True

End Sub

Private Sub SelectRow()
    ' Enables the radio button for the current MFC
    Dim CurrentRow As Integer
    If SelectedRow < 1 Then SelectedRow = 1
    CurrentRow = SelectedRow
    optCalibTable(CurrentRow).Value = True
End Sub

Private Sub DisableRow(RowToDisable)
    ' Unselects every element in this row
    Dim IndexForRow As Integer

    IndexForRow = RowToDisable
    txtDesired(RowToDisable).Enabled = False
    txtActual(RowToDisable).Enabled = False
End Sub

Private Sub UpdateSelection()
    ' Updates the MFC calibration table settings based on the
    ' current row selected in SelectedMFC()
    Dim row_num As Integer
    Curr_MfcCal = Stn_MfcCal(SelectedStation, SelectedMFC)
    bHistoryExists = CalCheckHistoryFound
    FillCalPoints
    For row_num = 1 To NUMROWS
        If row_num <> SelectedRow Then DisableRow (row_num)
    Next row_num
    SelectRow
    ResetChanged
End Sub

Private Function CalCheckHistoryFound() As Boolean
'
Dim rsCrit As String
Dim dbDbase As Database
Dim rsTable  As Recordset
Dim flag As Boolean
    
    flag = False
    ' open data table MfcCalCheck
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    rsCrit = "SELECT * FROM [MfcCalCheckData] "
    rsCrit = rsCrit & "WHERE [Station] = " & SelectedStation & " "
    rsCrit = rsCrit & "AND [Mfc] = " & SelectedMFC & " "
    rsCrit = rsCrit & "AND [CalDTS] = #" & Curr_MfcCal.Dts & "#  "
    rsCrit = rsCrit & " ORDER BY [MfcCalCheckData].[CalCheckDTS] ASC"
    
    ' open recordset
    Set rsTable = dbDbase.OpenRecordset(rsCrit, dbOpenDynaset)
    If rsTable.BOF Then
        ' no data
        flag = False
    Else
        ' at least one data record
        flag = True
    End If
    rsTable.Close
    dbDbase.Close
    CalCheckHistoryFound = flag
End Function
        
Private Sub ResetSelection()
    ' Points the MFC Calibration table to the first row
    SelectedRow = 1
    UpdateSelection
End Sub

Public Sub SaveCurrCalCheckData(ByVal iStation As Integer, ByVal iMfc As Integer)
Dim dbDbase As Database
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim calDTS As Date
Dim iPoint As Integer
Dim idx As Integer

'    ' delete existing cal check
'    ClearMfcCalRecords iStation, iMfc

    ' open calibration database
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Save MFC Cal Check Point Data
    calDTS = Curr_MfcCal.Dts
    
    ' update the cal check point information
    For iPoint = 1 To NUMROWS
        
        CriteriaPts = "SELECT * FROM [MfcCalCheckData] WHERE [Station] = " & iStation & " AND [Mfc] = " & iMfc & " AND [Point] = " & iPoint & " AND [CalDTS] = #" & calDTS & "#  AND [CalCheckDTS] = #" & curDTS & "#  "
        Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
                
        ' any records found ??
        If rsRecordPts.BOF Then
            ' no records; add a new one
            rsRecordPts.AddNew
            rsRecordPts("Station") = iStation
            rsRecordPts("MFC") = iMfc
            rsRecordPts("Point") = iPoint
            rsRecordPts("CalDTS") = calDTS
            rsRecordPts("CalCheckDTS") = curDTS
        Else
            ' record exists; edit it
            rsRecordPts.Edit
        End If
        rsRecordPts("FlowPct") = aryDesFlow(iPoint)
        rsRecordPts("FlowSP") = aryDesFlow(iPoint)
        rsRecordPts("FlowPV") = aryActualFlowSLPM(iPoint)
        rsRecordPts("CalCheckFlow") = aryActualFlowUncal(iPoint)
        ' update the record
        rsRecordPts.Update
        
    Next iPoint
    
    ' done with points
    rsRecordPts.Close
                        
                
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub txtActual_Change(Index As Integer)
    If IsNumeric(txtActual(Index).text) Then
        ' Store the value as Single and display the formatted value
        If Index = SelectedRow Then
            aryActualFlowUncal(Index) = CSng(txtActual(Index).text)
            bChanged(Index) = True
        End If
    End If
End Sub

Private Sub optCalibTable_Click(Index As Integer)
' Selects the row that was clicked on, for editing (but don't repeat up/dn cmd)
    If SelectedRow <> Index Then
        SelectedRow = Index
        UpdateSelection
        DesFlowSLPM_Validate SelectedRow
        DisplayData
        txtActual(SelectedRow).Enabled = True
        txtDesired(SelectedRow).Enabled = True
        txtDesired(SelectedRow).SetFocus
        If ((SelectedRow = 1) And bReadyToSaveSetup) Then
            bReadyToSave = True
        ElseIf (SelectedRow = Curr_MfcCal.NumPoints) Then
            bReadyToSaveSetup = True
        End If
    End If
End Sub

Private Sub tmrUpdate_Timer()
    Dim InputEng, MaxEng, MinEng, SpanEng As Single
    Dim DesEng, MaxVdc, MinVdc, SpanVdc As Single
    Dim RawVal, MaxCnt, MinCnt, SpanCnt As Single
    Dim RawEng As Single
    Dim ActualEng As Single
    Dim InputCal As Single
    Dim adr, chn As Integer
    
    ' Current DateTime
    txtDateTime.text = Format(Now(), "YYYY MMMM DD   HH:MM:SS")
    
    ' Purge Air Piab
    If ((SelectedMFC = MFCPURGEAIR) Or (SelectedMFC = MFCORVRPRG)) Then
        PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = True
    End If
    
    DesEng = Curr_MfcCal.PointData(SelectedRow).RawValue
    adr = Stn_AIO(SelectedStation, SelectedFunc).addr
    chn = Stn_AIO(SelectedStation, SelectedFunc).chan
    RawVal = Map_AIO(adr, chn).RawValue
    InputEng = Stn_AIO(SelectedStation, SelectedFunc).EUValue
    MaxEng = Stn_AIO(SelectedStation, SelectedFunc).EuMax
    MinEng = Stn_AIO(SelectedStation, SelectedFunc).EuMin
    SpanEng = MaxEng - MinEng
    MaxVdc = Stn_AIO(SelectedStation, SelectedFunc).VdcMax
    MinVdc = Stn_AIO(SelectedStation, SelectedFunc).VdcMin
    SpanVdc = MaxVdc - MinVdc
    MaxCnt = CSng(FULLSCALE) * (MaxVdc / 10#)
    MinCnt = CSng(FULLSCALE) * (MinVdc / 10#)
    SpanCnt = MaxCnt - MinCnt
    If SpanCnt <> 0# Then
        RawEng = MinEng + (SpanEng * ((RawVal - MinCnt) / SpanCnt))
    Else
        RawEng = 0#
    End If
'    aryActualFlowUncal(SelectedRow) = RawEng
    If (SpanEng <> 0) Then
        InputCal = Cal_MfcInput((InputEng / SpanEng), SelectedStation, SelectedMFC, Stn_MfcCal(SelectedStation, SelectedMFC))
    Else
        InputCal = 0
    End If
    lblCalibrated(SelectedRow).Caption = Format(InputCal, "##0.000")
    ActualFlowSLPM_Validate (SelectedRow)
    If IsNumeric(txtActual(SelectedRow).text) Then
        ActualEng = CSng(txtActual(SelectedRow).text)
    Else
        ActualEng = 0#
    End If
    If (SpanEng <> 0) Then
        aryPercDiff(SelectedRow) = 100# * (InputCal - ActualEng) / SpanEng
    Else
        aryPercDiff(SelectedRow) = 0
    End If

    DisplayData

    If AutoCycleOn Then
        If Now() > AutoStepNext Then
            AutoStep_Next
        Else
            txtSeconds.text = Format(DateDiff("s", Now(), AutoStepNext), "####0")
        End If
    End If
    
    cmdSave.Enabled = IIf(bReadyToSave, True, False)
    cmdHistory.Enabled = IIf(bHistoryExists, True, False)

End Sub

Private Sub ActualFlowSLPM_Validate(ByVal Index As Integer)
    ' Store and display the entered value if its numeric
    Dim sngTemp As Single
    If IsNumeric(lblCalibrated(Index).Caption) Then
        ' Store the value as Single and display the formatted value
        aryActualFlowSLPM(Index) = CSng(lblCalibrated(Index).Caption)
    End If
End Sub

Private Sub ActualFlowUncal(ByVal func As Integer)
Dim craw, clim, neweu As Single
Dim cmax, cmin, cspan As Single
Dim emax, emin, espan As Single

    cmax = clim * (Com_AIO(func).VdcMax / 10#)
    cmin = clim * (Com_AIO(func).VdcMin / 10#)
    cspan = cmax - cmin
    emax = Com_AIO(func).EuMax
    emin = Com_AIO(func).EuMin
    espan = emax - emin

    If cspan <> 0# Then neweu = emin + (espan * ((craw - cmin) / cspan))
End Sub

Private Sub DesFlowSLPM_Validate(ByVal Index As Integer)
'
    Dim sngTemp As Single
    If Not IsNumeric(txtDesired(Index).text) Then txtDesired(Index).text = "0.0"
    sngTemp = Stn_AIO(SelectedStation, SelectedFunc).EuMax
'    If (CSng(txtDesired(Index).text) > (CSng(2) * sngTemp)) Then txtDesired(Index).text = CStr(sngTemp)
    ' Store the value as Single and display the formatted value
    aryDesFlow(Index) = MfcMin + ((Curr_MfcCal.PointData(Index).RawPercent / CSng(100)) * MfcSpan)
    ' Output the value to the Mass Flow Controller
    SendFlow aryDesFlow(Index)
End Sub

Private Sub txtDesired_Change(Index As Integer)
    DesFlowSLPM_Validate Index
End Sub

Private Sub DisplayData()
    ' Displays actual flow data to the main MFC table
    Dim row As Integer
    For row = 1 To NUMROWS
        lblPointNum(row).Caption = Format(row, "##0")
        lblCalibrated(row).Caption = Format(aryActualFlowSLPM(row), "###0.0000")
        If row <> SelectedRow Then txtActual(row).text = Format(aryActualFlowUncal(row), "###0.0000")
        lblPercDiff(row).Caption = Format(aryPercDiff(row), "###0.00") & "%"
    Next row
End Sub

Private Sub PrintData()
    ' Print the values found for the current and previous calibrations
    Dim TempString, mfcname As String
    Dim strDesFlow, strDiffPress, strInletTemp, strInletPress, strBaroPress, strActualFlow, strPercDiff As String
    Dim row, intNumRows As Integer
    Dim sngDesiredFlow, sngActualFlow, sngPercDiff As Single
' ************  NEW  ****************
Dim oldFont As New StdFont
' Save current printer font
oldFont = Printer.Font
Printer.Font = FILEFONT
Printer.Font.Size = FILEFONTSIZE
Printer.Font.Bold = False
Printer.Font.Italic = False
    ' Print a title / header
    Print_Center Trim(SysConfig.Heading)
    Print_Center Trim(SysConfig.Heading2)
    Print_Center "CANISTER PRECONDITIONING SYSTEM"
    Print_Line ""
    Print_Center "Mass Flow Controller Calibration Check Function"
    Print_Center ("Date: " & Format(Now, "mmm d, yyyy"))
    Print_Line ""
    ' print the the name of the mass flow controller
    Select Case SelectedMFC
        Case MFCBUTANE
            mfcname = "Butane"
        Case MFCNITROGEN
            mfcname = "Nitrogen"
        Case MFCPURGEAIR
            mfcname = "PurgeAir"
        Case MFCORVRBUT
            mfcname = "ORVRBut"
        Case MFCORVRNIT
            mfcname = "ORVRNit"
        Case MFCLIVEFUEL
            mfcname = "LiveFuel"
        Case MFCORVRLIVE
            mfcname = "ORVRLiveFuel"
    End Select
    
    Print_Line "Station #" & SelectedStation
    Print_Line "Mass Flow Controller: " & mfcname
    Print_Line ""
    Print_Line ""
    Print_Line "Calibration Check"
    
    ' Print "Calibration date" then aryCurrCalibDate
    Print_Line "Date: " & Now

    
    ' Print a new line
    Print_Line ""
    
    
    ' Print the calbration formula text
    ' Print_Line "Y = " & coefX6(SelectedMFC(SelectedStation), SelectedStation) & "X6" & IIf(coefX5(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(coefX5(SelectedMFC(SelectedStation), SelectedStation)) & "X5" & IIf(coefX4(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(coefX4(SelectedMFC(SelectedStation), SelectedStation)) & "X4" & IIf(coefX3(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(coefX3(SelectedMFC(SelectedStation), SelectedStation)) & "X3" & IIf(coefX2(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(coefX2(SelectedMFC(SelectedStation), SelectedStation)) & "X2" & IIf(coefX(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(coefX(SelectedMFC(SelectedStation), SelectedStation)) & "X" & "      R2=" & coefR2(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line ""
    ' Print "Calibration points" centered
    Print_Center "Calibration Points"
    ' For each calibration point

    Print_Line "Des. Flow   LFE Dif. Pres.  LFE Inl. Pres.  LFE Inl. Temp.  Baro. Pres.  Actual Flow  % Diff."
    For row = 1 To NUMROWS
        ' Print data from the input table
        strDesFlow = Format(Curr_MfcCal.PointData(row).RawValue, "000.0000")
        sngActualFlow = Curr_MfcCal.PointData(row).ActualValue
        strActualFlow = Format(sngActualFlow, "000.0000")

        ' The percent difference
        sngPercDiff = percDiff(row)
        
        strPercDiff = Format(sngPercDiff, "0000.0") & "%"
        Print_Line strDesFlow & "    " & strDiffPress & "          " & strInletPress & "           " & strInletTemp & "           " & strBaroPress & "        " & strActualFlow & "     " & strPercDiff

    Next row
    

    Print_Footer
    Printer.EndDoc
    Printer.Font = oldFont

End Sub

Private Sub CalValves(ByVal outFlag As Boolean)
    ' Outputs the valves for Calibration of
    ' the selected Mass Flow Controller
Dim outRequest As Integer

    outRequest = IIf(outFlag, cON, cOFF)
'    If (Not outFlag) Then PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRdy = outFlag
    PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = outFlag
    
    Select Case SelectedMFC
        Case MFCBUTANE
            Stn_OutDigital SelectedStation, isButaneSol, outRequest
            Stn_OutDigital SelectedStation, isAuxDirectionSol, outRequest
            
        Case MFCNITROGEN
            Stn_OutDigital SelectedStation, isNitrogenSol, outRequest
            Stn_OutDigital SelectedStation, isAuxDirectionSol, outRequest
         
         Case MFCPURGEAIR
            PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRdy = outFlag
            Stn_OutDigital SelectedStation, isPurgeSol, outRequest
            Stn_OutDigital SelectedStation, isPriDirectionSol, outRequest
             
        Case MFCORVRBUT
            Stn_OutDigital SelectedStation, isButaneOrvrSol, outRequest
            Stn_OutDigital SelectedStation, isAuxDirectionSol, outRequest
         
         Case MFCORVRNIT
            Stn_OutDigital SelectedStation, isNitrogenOrvrSol, outRequest
            Stn_OutDigital SelectedStation, isAuxDirectionSol, outRequest
             
        Case MFCLIVEFUEL
            Stn_OutDigital SelectedStation, isLiveFuelSol, outRequest
            Stn_OutDigital SelectedStation, isAuxDirectionSol, outRequest
                 
        Case MFCORVRLIVE
            Stn_OutDigital SelectedStation, isLiveFuelOrvrSol, outRequest
            Stn_OutDigital SelectedStation, isAuxDirectionSol, outRequest
                 
    End Select
'
End Sub

Private Sub SendFlow(OutputSLPM As Single)
    ' Outputs the value specified in OutputSLPM to
    ' the selected Mass Flow Controller
    
    Dim OutputEng, span As Single
    
'    aryOutputFS(SelectedRow) = frmMassFlowCal.SolveCalibFor(OutputSLPM)

    Select Case SelectedMFC
        Case MFCBUTANE
            span = Stn_AIO(SelectedStation, asButaneFlowSP).EuMax - Stn_AIO(SelectedStation, asButaneFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asButaneFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCBUTANE, Stn_MfcCal(SelectedStation, MFCBUTANE)))
            Stn_OutAnalog SelectedStation, asButaneFlowSP, CSng(OutputEng), outNORMAL
        
        Case MFCNITROGEN
            span = Stn_AIO(SelectedStation, asNitrogenFlowSP).EuMax - Stn_AIO(SelectedStation, asNitrogenFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCNITROGEN, Stn_MfcCal(SelectedStation, MFCNITROGEN)))
            Stn_OutAnalog SelectedStation, asNitrogenFlowSP, CSng(OutputEng), outNORMAL
            
        Case MFCPURGEAIR
            span = Stn_AIO(SelectedStation, asPurgeAirFlowSP).EuMax - Stn_AIO(SelectedStation, asPurgeAirFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asPurgeAirFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCPURGEAIR, Stn_MfcCal(SelectedStation, MFCPURGEAIR)))
            Stn_OutAnalog SelectedStation, asPurgeAirFlowSP, CSng(OutputEng), outNORMAL
            
        Case MFCORVRBUT
            span = Stn_AIO(SelectedStation, asButaneORVRFlowSP).EuMax - Stn_AIO(SelectedStation, asButaneORVRFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asButaneORVRFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCORVRBUT, Stn_MfcCal(SelectedStation, MFCORVRBUT)))
            Stn_OutAnalog SelectedStation, asButaneORVRFlowSP, CSng(OutputEng), outNORMAL
    
        Case MFCORVRNIT
            span = Stn_AIO(SelectedStation, asNitrogenORVRFlowSP).EuMax - Stn_AIO(SelectedStation, asNitrogenORVRFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asNitrogenORVRFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCORVRNIT, Stn_MfcCal(SelectedStation, MFCORVRNIT)))
            Stn_OutAnalog SelectedStation, asNitrogenORVRFlowSP, CSng(OutputEng), outNORMAL
            
        Case MFCLIVEFUEL
            span = Stn_AIO(SelectedStation, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(SelectedStation, asLiveFuelVaporFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCLIVEFUEL, Stn_MfcCal(SelectedStation, MFCLIVEFUEL)))
            Stn_OutAnalog SelectedStation, asLiveFuelVaporFlowSP, CSng(OutputEng), outNORMAL
                
        Case MFCORVRLIVE
            span = Stn_AIO(SelectedStation, asLiveFuelVaporORVRFlowSP).EuMax - Stn_AIO(SelectedStation, asLiveFuelVaporORVRFlowSP).EuMin
            OutputEng = Stn_AIO(SelectedStation, asLiveFuelVaporORVRFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelectedStation, MFCORVRLIVE, Stn_MfcCal(SelectedStation, MFCORVRLIVE)))
            Stn_OutAnalog SelectedStation, asLiveFuelVaporORVRFlowSP, CSng(OutputEng), outNORMAL
                
    End Select
    aryOutputFS(SelectedRow) = CSng(OutputEng)
    If (SelectedRow <> 0) Then txtOutput(SelectedRow) = Format(CSng(OutputEng), "#####0.000")
End Sub

Private Function percDiff(Index)
    ' Calculates the percent difference value between the
    ' desired flow and the actual flow
    ' This function applies to a single row, specified by Index
    
    If aryDesFlow(Index) > 0 And aryActualFlowSLPM(Index) > 0 Then
        percDiff = ((aryActualFlowSLPM(Index) - aryDesFlow(Index)) / aryDesFlow(Index)) * 100
    Else
        percDiff = 0!
    End If
End Function

Private Sub AutoStep_Init()
    ' set auto step interval
    If Not IsNumeric(CLng(txtSeconds.text)) Then txtSeconds.text = Format((2 * MFC_Settle_Time), "####0")
    AutoStepInterval = CLng(txtSeconds.text)
    ' set next timeout first just in case the timer is about to tick
    AutoStepNext = DateAdd("s", CDbl(AutoStepInterval), Now())
    ' turn on auto cycle
    AutoCycleOn = True
    ' change icon
    cmdAuto.Picture = LoadResPicture(101, vbResBitmap)
    ' change tooltip
    cmdAuto.ToolTipText = "Stop Auto Check Cycle"
    ' disable up/down buttons
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    ' show Interval in Countdown box
    txtSeconds.text = Format((AutoStepInterval), "####0")
    txtSeconds.ForeColor = DK2GREEN
    ' start with first row
    SelectedRow = 1
    txtActual(SelectedRow).Enabled = True
    txtDesired(SelectedRow).Enabled = True
    txtDesired(SelectedRow).SetFocus
    UpdateSelection
    ' set desired flow for current row
    AutoCycleSP = MfcMin + (MfcSpan * (0.1 * (SelectedRow - 1)))
'    txtDesired(selectedRow).text = Format(AutoCycleSP, "####0.0000")
'    UpdateSelection
    optCalibTable(SelectedRow).Value = True
    DesFlowSLPM_Validate SelectedRow
    DisplayData
End Sub

Private Sub AutoStep_Next()
    ' set next timeout first, just in case
    AutoStepNext = DateAdd("s", CDbl(AutoStepInterval), Now())
    ' update Seconds Countdown box
    txtSeconds.text = Format(AutoStepInterval, "####0")
    ' Format current desired value
    txtDesired(SelectedRow).text = Format(aryDesFlow(SelectedRow), "####0.0000")
    ' done yet
    If SelectedRow >= NUMROWS Then
        AutoStep_Done
    Else
        ' increment row#
        SelectedRow = SelectedRow + 1
        txtActual(SelectedRow).Enabled = True
        txtDesired(SelectedRow).Enabled = True
        txtDesired(SelectedRow).SetFocus
        UpdateSelection
        ' set desired flow for current row
        AutoCycleSP = MfcMin + (MfcSpan * (0.1 * (SelectedRow - 1)))
'        txtDesired(selectedRow).text = Format(AutoCycleSP, "####0.0000")
'        UpdateSelection
        optCalibTable(SelectedRow).Value = True
        DesFlowSLPM_Validate SelectedRow
        DisplayData
    End If
End Sub

Private Sub AutoStep_Done()
    ' set next timeout first, just in case
    AutoStepNext = DateAdd("s", CDbl(AutoStepInterval), Now())
    ' turn off auto cycle
    AutoCycleOn = False
    ' change icon
    cmdAuto.Picture = LoadResPicture(102, vbResBitmap)
    ' change tooltip
    cmdAuto.ToolTipText = "Run Auto Check Cycle"
    ' enable up/down buttons
    cmdUp.Enabled = True
    cmdDown.Enabled = True
    ' show Interval in Countdown box
    txtSeconds.text = Format((AutoStepInterval), "####0")
    txtSeconds.ForeColor = DK3RED
    ' return to first row
    SelectedRow = 1
'    UpdateSelection
    optCalibTable(SelectedRow).Value = True
    DesFlowSLPM_Validate SelectedRow
    DisplayData
    txtDesired(SelectedRow).Enabled = True
    txtDesired(SelectedRow).SetFocus
End Sub

Private Sub FillCalPoints()
'
Dim iPoint As Integer
Dim idx As Integer
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim sVdcMax As Single
Dim sVdcMin As Single
Dim sVdcSpan As Single
    
    Select Case NUMROWS
        Case 3
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(50)
            Curr_MfcCal.PointData(3).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(50)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(100)
        Case 4
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(30)
            Curr_MfcCal.PointData(3).RawPercent = CSng(70)
            Curr_MfcCal.PointData(4).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(30)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(70)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(100)
        Case 5
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(25)
            Curr_MfcCal.PointData(3).RawPercent = CSng(50)
            Curr_MfcCal.PointData(4).RawPercent = CSng(75)
            Curr_MfcCal.PointData(5).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(25)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(50)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(75)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(100)
        Case 6
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(20)
            Curr_MfcCal.PointData(3).RawPercent = CSng(40)
            Curr_MfcCal.PointData(4).RawPercent = CSng(60)
            Curr_MfcCal.PointData(5).RawPercent = CSng(80)
            Curr_MfcCal.PointData(6).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(20)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(40)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(60)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(80)
            Curr_MfcCal.PointData(6).ActualPercent = CSng(100)
        Case 7
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(16.7)
            Curr_MfcCal.PointData(3).RawPercent = CSng(33.3)
            Curr_MfcCal.PointData(4).RawPercent = CSng(50)
            Curr_MfcCal.PointData(5).RawPercent = CSng(66.7)
            Curr_MfcCal.PointData(6).RawPercent = CSng(83.3)
            Curr_MfcCal.PointData(7).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(16.7)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(33.3)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(50)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(66.7)
            Curr_MfcCal.PointData(6).ActualPercent = CSng(83.3)
            Curr_MfcCal.PointData(7).ActualPercent = CSng(100)
        Case 8
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(14.3)
            Curr_MfcCal.PointData(3).RawPercent = CSng(28.6)
            Curr_MfcCal.PointData(4).RawPercent = CSng(42.9)
            Curr_MfcCal.PointData(5).RawPercent = CSng(57.1)
            Curr_MfcCal.PointData(6).RawPercent = CSng(71.4)
            Curr_MfcCal.PointData(7).RawPercent = CSng(85.7)
            Curr_MfcCal.PointData(8).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(14.3)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(28.6)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(42.9)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(57.1)
            Curr_MfcCal.PointData(6).ActualPercent = CSng(71.4)
            Curr_MfcCal.PointData(7).ActualPercent = CSng(85.7)
            Curr_MfcCal.PointData(8).ActualPercent = CSng(100)
        Case 9
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(12.5)
            Curr_MfcCal.PointData(3).RawPercent = CSng(25)
            Curr_MfcCal.PointData(4).RawPercent = CSng(37.5)
            Curr_MfcCal.PointData(5).RawPercent = CSng(50)
            Curr_MfcCal.PointData(6).RawPercent = CSng(62.5)
            Curr_MfcCal.PointData(7).RawPercent = CSng(75)
            Curr_MfcCal.PointData(8).RawPercent = CSng(87.5)
            Curr_MfcCal.PointData(9).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(12.5)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(25)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(37.5)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(50)
            Curr_MfcCal.PointData(6).ActualPercent = CSng(62.5)
            Curr_MfcCal.PointData(7).ActualPercent = CSng(75)
            Curr_MfcCal.PointData(8).ActualPercent = CSng(87.5)
            Curr_MfcCal.PointData(9).ActualPercent = CSng(100)
        Case 10
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(11.1)
            Curr_MfcCal.PointData(3).RawPercent = CSng(22.2)
            Curr_MfcCal.PointData(4).RawPercent = CSng(33.3)
            Curr_MfcCal.PointData(5).RawPercent = CSng(44.4)
            Curr_MfcCal.PointData(6).RawPercent = CSng(55.5)
            Curr_MfcCal.PointData(7).RawPercent = CSng(66.6)
            Curr_MfcCal.PointData(8).RawPercent = CSng(77.7)
            Curr_MfcCal.PointData(9).RawPercent = CSng(88.8)
            Curr_MfcCal.PointData(10).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(11.1)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(22.2)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(33.3)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(44.4)
            Curr_MfcCal.PointData(6).ActualPercent = CSng(55.5)
            Curr_MfcCal.PointData(7).ActualPercent = CSng(66.6)
            Curr_MfcCal.PointData(8).ActualPercent = CSng(77.7)
            Curr_MfcCal.PointData(9).ActualPercent = CSng(88.8)
            Curr_MfcCal.PointData(10).ActualPercent = CSng(100)
        Case 11
            Curr_MfcCal.PointData(1).RawPercent = CSng(0)
            Curr_MfcCal.PointData(2).RawPercent = CSng(10)
            Curr_MfcCal.PointData(3).RawPercent = CSng(20)
            Curr_MfcCal.PointData(4).RawPercent = CSng(30)
            Curr_MfcCal.PointData(5).RawPercent = CSng(40)
            Curr_MfcCal.PointData(6).RawPercent = CSng(50)
            Curr_MfcCal.PointData(7).RawPercent = CSng(60)
            Curr_MfcCal.PointData(8).RawPercent = CSng(70)
            Curr_MfcCal.PointData(9).RawPercent = CSng(80)
            Curr_MfcCal.PointData(10).RawPercent = CSng(90)
            Curr_MfcCal.PointData(11).RawPercent = CSng(100)
            Curr_MfcCal.PointData(1).ActualPercent = CSng(0)
            Curr_MfcCal.PointData(2).ActualPercent = CSng(10)
            Curr_MfcCal.PointData(3).ActualPercent = CSng(20)
            Curr_MfcCal.PointData(4).ActualPercent = CSng(30)
            Curr_MfcCal.PointData(5).ActualPercent = CSng(40)
            Curr_MfcCal.PointData(6).ActualPercent = CSng(50)
            Curr_MfcCal.PointData(7).ActualPercent = CSng(60)
            Curr_MfcCal.PointData(8).ActualPercent = CSng(70)
            Curr_MfcCal.PointData(9).ActualPercent = CSng(80)
            Curr_MfcCal.PointData(10).ActualPercent = CSng(90)
            Curr_MfcCal.PointData(11).ActualPercent = CSng(100)
    End Select
    
    ' get min/max EU & Raw  for appropriate mfc
    ' Station MFC Calibration Parameters
    idx = SelectedStation
    sEuMax = Stn_AIO(idx, SelectedFunc).EuMax
    sEuMin = Stn_AIO(idx, SelectedFunc).EuMin
    sVdcMax = Stn_AIO(idx, SelectedFunc).VdcMax
    sVdcMin = Stn_AIO(idx, SelectedFunc).VdcMin
    ' calc EU & Vdc spans
    sEuSpan = sEuMax - sEuMin
    sVdcSpan = sVdcMax - sVdcMin
    
    For iPoint = 1 To MAXLSQCALPOINTS
        Select Case Curr_MfcCal.RawInputType
            Case CalRawAsMa
                Curr_MfcCal.PointData(iPoint).RawValue = CSng(4) * (sVdcMin + (sVdcSpan * (Curr_MfcCal.PointData(iPoint).RawPercent / CSng(100))))
            Case CalRawAsVolts
                Curr_MfcCal.PointData(iPoint).RawValue = sVdcMin + (sVdcSpan * (Curr_MfcCal.PointData(iPoint).RawPercent / CSng(100)))
            Case CalRawAsDegC
                Curr_MfcCal.PointData(iPoint).RawValue = sVdcMin + (sVdcSpan * (Curr_MfcCal.PointData(iPoint).RawPercent / CSng(100)))
        End Select
        Curr_MfcCal.PointData(iPoint).RawValue = sEuMin + (sEuSpan * (Curr_MfcCal.PointData(iPoint).RawPercent / CSng(100)))
        Curr_MfcCal.PointData(iPoint).ActualValue = sEuMin + (sEuSpan * (Curr_MfcCal.PointData(iPoint).ActualPercent / CSng(100)))
        txtDesired(iPoint).text = Format(Curr_MfcCal.PointData(iPoint).RawValue, "####0.###")
        txtActual(iPoint).text = Format(Curr_MfcCal.PointData(iPoint).ActualValue, "####0.###")
    Next iPoint
End Sub

Private Sub Xit()
    ' Close the form and switch to the Mass Flow Calibration form
    SendFlow 0#
    If (Not CalReadOnly) Then PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRdy = False
    If (Not CalReadOnly) Then CalValves False
    Unload Me
    Set frmCalCheck = Nothing
    frmMfcCal.Enabled = True
    frmMfcCal.Show
End Sub

