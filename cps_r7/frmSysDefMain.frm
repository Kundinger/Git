VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysDefMain 
   Caption         =   "System Definition Main"
   ClientHeight    =   11970
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   14865
   Icon            =   "frmSysDefMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11970
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   15960
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   114
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrint 
      DisabledPicture =   "frmSysDefMain.frx":57E2
      DownPicture     =   "frmSysDefMain.frx":6424
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefMain.frx":7066
      Style           =   1  'Graphical
      TabIndex        =   113
      ToolTipText     =   "Print Screen"
      Top             =   11100
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdNext 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13275
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefMain.frx":7CA8
      Style           =   1  'Graphical
      TabIndex        =   98
      ToolTipText     =   "Next Screen"
      Top             =   11100
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.Timer tmrscreen 
      Interval        =   50
      Left            =   4200
      Top             =   11300
   End
   Begin VB.Frame frmHighlight 
      BackColor       =   &H8000000D&
      Caption         =   "highlight"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1200
      TabIndex        =   60
      Top             =   11280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frmNotHighlight 
      BackColor       =   &H80000005&
      Caption         =   "NOT highlight"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   61
      Top             =   11040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmSystemSettings 
      Caption         =   "System Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   10900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.TextBox txtWcmRecipe 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         MaxLength       =   3
         TabIndex        =   168
         Text            =   "10"
         ToolTipText     =   "Enter Recipe# of TOM Default Working Capacity Multiplier Recipe"
         Top             =   5640
         Width           =   735
      End
      Begin VB.TextBox txt2GmRecipe 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         MaxLength       =   3
         TabIndex        =   167
         Text            =   "10"
         ToolTipText     =   "Enter Recipe# of TOM Default 2 Gram Breakthrough Recipe"
         Top             =   5400
         Width           =   735
      End
      Begin VB.CheckBox chkTomCanLoad 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Old Remote Load? TOM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   164
         ToolTipText     =   "Using TOM CAN Load ?"
         Top             =   5060
         Width           =   3645
      End
      Begin VB.CheckBox chkHighTempPas 
         Caption         =   " Using High Temperature PAS?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   162
         ToolTipText     =   " Using High Temperature Capable PAS ?"
         Top             =   8520
         Width           =   3645
      End
      Begin VB.CheckBox chkRemEnableChgs 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow Operator Changes?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6570
         TabIndex        =   161
         ToolTipText     =   " Using AVL File Interface to Task Order Manager ?"
         Top             =   6435
         Width           =   3045
      End
      Begin VB.CheckBox chkRemStatusMon 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Remote Status Monitoring?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   160
         ToolTipText     =   " Using AVL File Interface to Task Order Manager ?"
         Top             =   5925
         Width           =   3645
      End
      Begin VB.CheckBox chkRemAvlFiles 
         Alignment       =   1  'Right Justify
         Caption         =   "Using AVL File Interface?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6570
         TabIndex        =   159
         ToolTipText     =   " Using AVL File Interface to Task Order Manager ?"
         Top             =   6180
         Width           =   3045
      End
      Begin VB.Frame frmPAScontrol 
         Caption         =   "Purge Air Sources"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   1515
         Left            =   30
         TabIndex        =   149
         Top             =   2760
         Width           =   5775
         Begin VB.OptionButton optClient 
            Alignment       =   1  'Right Justify
            Caption         =   "AK Client"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   157
            ToolTipText     =   "PurgeAir Generator Control via AK Client"
            Top             =   1200
            Width           =   1365
         End
         Begin VB.OptionButton optMaster 
            Alignment       =   1  'Right Justify
            Caption         =   "AK Master"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   156
            ToolTipText     =   "AK Master PurgeAir Generator Control"
            Top             =   960
            Width           =   1365
         End
         Begin VB.OptionButton optAlone 
            Alignment       =   1  'Right Justify
            Caption         =   "Stand-Alone"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   155
            ToolTipText     =   "Stand-Alone PurgeAir Generator Control"
            Top             =   1200
            Width           =   1365
         End
         Begin VB.OptionButton optNone 
            Alignment       =   1  'Right Justify
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   154
            ToolTipText     =   "No PurgeAir Generator Control"
            Top             =   960
            Width           =   1365
         End
         Begin VB.TextBox txtServerIp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3990
            MaxLength       =   15
            TabIndex        =   152
            Text            =   "255.255.255.255"
            ToolTipText     =   "This AK Client's AK Server's IP Address"
            Top             =   1140
            Width           =   1575
         End
         Begin VB.TextBox txtNrPrg 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2790
            MaxLength       =   1
            TabIndex        =   150
            Text            =   "0"
            ToolTipText     =   "Enter 1 to 19 PurgeAir Supplies"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblPagType 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PAG Control Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   210
            Left            =   600
            TabIndex        =   158
            ToolTipText     =   "FID's AK Server IP Address"
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblServerIp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Server  IP Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   210
            Left            =   3930
            TabIndex        =   153
            Top             =   930
            Width           =   1695
         End
         Begin VB.Label lblPurgeAir 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of PurgeAir Sources"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   151
            ToolTipText     =   "Number of PurgeAir Sources"
            Top             =   360
            Width           =   2430
         End
      End
      Begin VB.TextBox txtWbEuMax 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8895
         MaxLength       =   5
         TabIndex        =   148
         Text            =   "100"
         ToolTipText     =   "WaterBath Engr Units Range Maximum Value"
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txtWbEuMin 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         MaxLength       =   5
         TabIndex        =   146
         Text            =   "0"
         ToolTipText     =   "WaterBath Engr Units Range Minimum Value"
         Top             =   3960
         Width           =   735
      End
      Begin VB.CheckBox chkUsingDryPurgeAir 
         Caption         =   " Using Dry Purge Air?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   145
         ToolTipText     =   " Using Positive Pressure Purge?"
         Top             =   10560
         Width           =   3315
      End
      Begin VB.CheckBox chkWaterBath 
         Alignment       =   1  'Right Justify
         Caption         =   "Using WaterBath Chiller?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   144
         ToolTipText     =   " Using DP during Purge?"
         Top             =   3030
         Width           =   3645
      End
      Begin VB.CheckBox chkPurgeOven 
         Caption         =   " Using Purge Oven?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   143
         ToolTipText     =   " Using DP during Purge?"
         Top             =   10320
         Width           =   3315
      End
      Begin VB.TextBox txtChillerComPort 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8880
         MaxLength       =   1
         TabIndex        =   139
         Text            =   "6"
         ToolTipText     =   "Com Port for WaterBath Heater communications"
         Top             =   3270
         Width           =   735
      End
      Begin VB.TextBox txtChillerCommTimeout 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   138
         Text            =   "750"
         ToolTipText     =   "Number of milliseconds of no Heater comm before a Timeout occurs."
         Top             =   3615
         Width           =   735
      End
      Begin VB.TextBox txtMFCSettleTime 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   137
         Text            =   "3"
         ToolTipText     =   "Number of seconds to allow for MFC's to settle"
         Top             =   795
         Width           =   735
      End
      Begin VB.TextBox txtOPTOComPort 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   136
         Text            =   "6"
         ToolTipText     =   "Com Port with Opto communications card"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtMsgDelay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   135
         Text            =   "750"
         ToolTipText     =   "Number of milliseconds that the DelayBox should remain open."
         Top             =   2685
         Width           =   735
      End
      Begin VB.TextBox txtLoadMfcDelay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   134
         Text            =   "3"
         ToolTipText     =   "Number of seconds after valves are opened that the MFC SetPoint is updateded"
         Top             =   1110
         Width           =   735
      End
      Begin VB.TextBox txtLoadEqlDelay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   133
         Text            =   "3"
         ToolTipText     =   "Number of seconds to allow for Scales & Butane Flow rate to settle"
         Top             =   1425
         Width           =   735
      End
      Begin VB.TextBox txtGramsPerLiter 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   132
         Text            =   "3"
         ToolTipText     =   "Density of the Butane in use (note: gm/l = kg/m3, default=2.40633)"
         Top             =   2055
         Width           =   735
      End
      Begin VB.TextBox txtMinMfcSp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         MaxLength       =   5
         TabIndex        =   131
         Text            =   "3"
         ToolTipText     =   "Minimum MFC SetPoint in Percent (1-5 % of fullscale)"
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox txtMinDataLogInterval 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   130
         Text            =   "3"
         ToolTipText     =   "Minimum seconds allowed for data logging (0.1 - 60)"
         Top             =   2370
         Width           =   735
      End
      Begin VB.TextBox txtDefScaleMax 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8880
         TabIndex        =   129
         Text            =   "10,000"
         ToolTipText     =   "Default Scale Range Max Reading in grams (100 to 100,000)"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtWeakLiveFuelDensity 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   127
         Text            =   "3"
         ToolTipText     =   "Threshold for ""Weak"" fuel density  (in gm/liter)"
         Top             =   9690
         Width           =   735
      End
      Begin VB.TextBox txtDeadLiveFuelDensity 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   125
         Text            =   "3"
         ToolTipText     =   "Threshold for ""Dead"" fuel density  (in gm/liter)"
         Top             =   9405
         Width           =   735
      End
      Begin VB.TextBox txtMaxSheathTempForADF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         TabIndex        =   123
         Text            =   "3"
         ToolTipText     =   "Maximum Sheath Temp For LiveFuel ADF Drain (in deg C)"
         Top             =   9120
         Width           =   735
      End
      Begin VB.CheckBox chkUseFuelLevelOot 
         Alignment       =   1  'Right Justify
         Caption         =   "Using LiveFuel Tank Level OOT?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   122
         ToolTipText     =   "Using Continue on LeakCheck Fail?"
         Top             =   10125
         Width           =   3645
      End
      Begin VB.CheckBox chkPurgeDP 
         Caption         =   " Using DP during Purge?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   121
         ToolTipText     =   " Using DP during Purge?"
         Top             =   9840
         Width           =   3315
      End
      Begin VB.CheckBox chkPurgeSeries 
         Caption         =   " Using Purge Canisters in Series?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   120
         ToolTipText     =   " Using System Vacuum Switch?"
         Top             =   10080
         Width           =   3315
      End
      Begin VB.CheckBox chkSystemVacSw 
         Caption         =   " Using System Vacuum Switch?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   119
         ToolTipText     =   " Using System Vacuum Switch?"
         Top             =   5400
         Width           =   3315
      End
      Begin VB.TextBox txtScaleWeightQueue 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   116
         Text            =   "0"
         ToolTipText     =   "number of elements in the weights queue used for calculating the Weight Running Average"
         Top             =   1335
         Width           =   615
      End
      Begin VB.CheckBox chkHardPipedScales 
         Caption         =   " Using Hard-Piped Scales?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   115
         ToolTipText     =   " Are Scales hardpiped (2 scales per station) ?"
         Top             =   4320
         Width           =   2685
      End
      Begin VB.CheckBox chkAuxLeakChk 
         Caption         =   " Using Aux. Leak Check?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   112
         ToolTipText     =   " Using Leakcheck of Aux plumbing; Leakcheck Aux Only and Leakcheck Both (pri & aux)  ?"
         Top             =   6120
         Width           =   3315
      End
      Begin VB.CheckBox chkRemTaskOrder 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Remote TaskOrdering?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   111
         ToolTipText     =   " Using Task Order Manager Tasks ?"
         Top             =   4800
         Width           =   3645
      End
      Begin VB.CheckBox chkLogTempRh 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Air Temp and Humidity Log?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   110
         ToolTipText     =   "  Log Air Temperature and Humidity to a DB file ?"
         Top             =   4320
         Width           =   3645
      End
      Begin VB.TextBox txtAuxOutDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5880
         MaxLength       =   24
         TabIndex        =   105
         Text            =   "123456789012345678901234"
         ToolTipText     =   "Aux Output #1 Description"
         Top             =   7560
         Width           =   3735
      End
      Begin VB.TextBox txtAuxOutDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5880
         MaxLength       =   24
         TabIndex        =   104
         Text            =   "desc"
         ToolTipText     =   "Aux Output #2 Description"
         Top             =   7920
         Width           =   3735
      End
      Begin VB.TextBox txtAuxOutDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5880
         MaxLength       =   24
         TabIndex        =   103
         Text            =   "123456789012345678901234"
         ToolTipText     =   "Aux Output #3 Description"
         Top             =   8280
         Width           =   3735
      End
      Begin VB.TextBox txtAuxOutDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5880
         MaxLength       =   24
         TabIndex        =   102
         Text            =   "desc"
         ToolTipText     =   "Aux Output #4 Description"
         Top             =   8640
         Width           =   3735
      End
      Begin VB.CheckBox chkEstopInput 
         Caption         =   " Using E-Stop Intput?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   101
         ToolTipText     =   " Using E-Stop Input (12vdc or Dry Contact) ?"
         Top             =   4680
         Width           =   2685
      End
      Begin VB.TextBox txtScaleMaxUnstableReads 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   96
         Text            =   "0"
         ToolTipText     =   "Max number of consecutive unstable reads. (0-9999)"
         Top             =   900
         Width           =   615
      End
      Begin VB.TextBox txtAuxOutputs 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   94
         Text            =   "10"
         ToolTipText     =   "Enter 1 to 4 Aux Ouputs/Station"
         Top             =   7560
         Width           =   735
      End
      Begin VB.CheckBox chkAuxOutputs 
         Caption         =   " Using Aux. Outputs?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   93
         ToolTipText     =   " Using Aux Outputs (12vdc or Dry Contact) ?"
         Top             =   7560
         Width           =   2415
      End
      Begin VB.Frame frmDebug 
         Caption         =   "Debug Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   5745
         Left            =   10080
         TabIndex        =   28
         ToolTipText     =   "Program Debug Options"
         Top             =   4320
         Width           =   4385
         Begin VB.CheckBox chkVerboseStartup 
            Caption         =   "Verbose Messages during Startup?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   163
            ToolTipText     =   "Verbose Messages during Program Startup ?"
            Top             =   1200
            Width           =   3705
         End
         Begin VB.CheckBox chkChillerComm 
            Caption         =   " No Comm with Chiller?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   142
            ToolTipText     =   "Turn Off communication with the Chiller"
            Top             =   790
            Width           =   2505
         End
         Begin VB.CommandButton cmdDisplayProperties 
            DisabledPicture =   "frmSysDefMain.frx":83AA
            DownPicture     =   "frmSysDefMain.frx":8AAC
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefMain.frx":91AE
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Open the Display Properties screen"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.CheckBox chkPasDebug 
            Caption         =   " PAS Debug?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   87
            ToolTipText     =   "PAS Debug"
            Top             =   2340
            Width           =   1995
         End
         Begin VB.CommandButton cmdMonTimers 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefMain.frx":98B0
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Monitor System Timers"
            Top             =   2520
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.CheckBox chkSclDebug 
            Caption         =   " Scale Debug?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   84
            ToolTipText     =   "Scale Debug"
            Top             =   2820
            Width           =   1995
         End
         Begin VB.CommandButton cmdSimCntrlPnl 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefMain.frx":9FB2
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Open the Simulation Control Panel"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   285
            Index           =   9
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   57
            Text            =   "250"
            ToolTipText     =   "unused Timer  interval in milliseconds"
            Top             =   5280
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   285
            Index           =   8
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   55
            Text            =   "250"
            ToolTipText     =   "Report Generation interval in milliseconds"
            Top             =   5040
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   54
            Text            =   "250"
            ToolTipText     =   "TestTimer update interval in milliseconds"
            Top             =   4800
            Width           =   735
         End
         Begin VB.CheckBox chkPrgDebug 
            Caption         =   " Purge Debug?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   46
            ToolTipText     =   "Purge Debug"
            Top             =   2580
            Width           =   1995
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   45
            Text            =   "250"
            ToolTipText     =   "Alarm/OOT scan interval in milliseconds"
            Top             =   3690
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   43
            Text            =   "250"
            ToolTipText     =   "DataLogger scan interval in milliseconds"
            Top             =   3975
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   41
            Text            =   "250"
            ToolTipText     =   "Scale scan interval in milliseconds"
            Top             =   3420
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   39
            Text            =   "250"
            ToolTipText     =   "Stations scan interval in milliseconds"
            Top             =   4530
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   37
            Text            =   "250"
            ToolTipText     =   "Controllers scan interval in milliseconds"
            Top             =   4245
            Width           =   735
         End
         Begin VB.TextBox txtTmrInterval 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   35
            Text            =   "250"
            ToolTipText     =   "IO scan interval in milliseconds"
            Top             =   3135
            Width           =   735
         End
         Begin VB.CheckBox chkReadScales 
            Caption         =   " Don't Read Scales?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   33
            ToolTipText     =   "Turn Off IO Scanning"
            Top             =   545
            Width           =   4155
         End
         Begin VB.CheckBox chkMmwDebug 
            Caption         =   " MMW Debug?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   32
            ToolTipText     =   "MMW Debug"
            Top             =   2100
            Width           =   1995
         End
         Begin VB.CheckBox chkADFdebug 
            Caption         =   " ADF Debug?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   31
            ToolTipText     =   "ADF Debug"
            Top             =   1860
            Width           =   1995
         End
         Begin VB.CheckBox chkDebug 
            Caption         =   " No Error Handler?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   30
            ToolTipText     =   "Debug"
            Top             =   1620
            Width           =   1995
         End
         Begin VB.CheckBox chkScanIO 
            Caption         =   " Don't Scan IO?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   29
            ToolTipText     =   "Turn Off IO Scanning"
            Top             =   300
            Width           =   4155
         End
         Begin VB.Label lblTmr9 
            BackStyle       =   0  'Transparent
            Caption         =   "unused Timer Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   5280
            Width           =   3015
         End
         Begin VB.Label lblRptGenInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "unused Timer Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000011&
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   5040
            Width           =   3015
         End
         Begin VB.Label lblTimerInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "TestTimer Update Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   4800
            Width           =   3015
         End
         Begin VB.Label lblAlarmOOTInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Alarm/OOT Logic Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   3705
            Width           =   3015
         End
         Begin VB.Label lblDataLogInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "DataLogger Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   3990
            Width           =   2895
         End
         Begin VB.Label lblScalesInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Scan Scales Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   3420
            Width           =   2895
         End
         Begin VB.Label lblStationInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Stations Logic Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   4545
            Width           =   3015
         End
         Begin VB.Label lblControlInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Controllers Logic Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   4260
            Width           =   2895
         End
         Begin VB.Label lblIOInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Scan IO Interval in msec"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   3135
            Width           =   2895
         End
      End
      Begin VB.Frame frmMoistUnits 
         Caption         =   "Engr Units for Moisture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   825
         Left            =   10080
         TabIndex        =   89
         Top             =   1800
         Width           =   4385
         Begin VB.OptionButton optMoistRH 
            Alignment       =   1  'Right Justify
            Caption         =   "Using Relative Humidity"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   91
            ToolTipText     =   "Using Relative Humidity (% rH)"
            Top             =   240
            Width           =   3885
         End
         Begin VB.OptionButton optMoistGrains 
            Alignment       =   1  'Right Justify
            Caption         =   "Using Grains per Pound"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   90
            ToolTipText     =   "Using Grains per pound (grains/lb)"
            Top             =   480
            Width           =   3885
         End
      End
      Begin VB.TextBox txtNrSeq 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   4
         TabIndex        =   85
         Text            =   "0"
         ToolTipText     =   "Enter 1  to 999 sequences; 1=OnlySequenceIsTheDefaultSequence"
         Top             =   2460
         Width           =   735
      End
      Begin VB.CheckBox chkErrorMsgBypass 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Bypass-Program-Error-Msgs?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   80
         ToolTipText     =   "Using Bypass-Program-Error-Msgs?"
         Top             =   10455
         Width           =   3645
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   555
         Left            =   9720
         TabIndex        =   77
         Top             =   9000
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   979
         _Version        =   393216
         LargeChange     =   32
         SmallChange     =   8
         Min             =   32768
         Max             =   98301
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   2048
         Value           =   32768
      End
      Begin VB.CheckBox chkLeakChkExhSol 
         Caption         =   " Using Leak Check Exhaust Sol?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   76
         ToolTipText     =   " Using LeakCheck Exhaust Solenoid?"
         Top             =   5880
         Width           =   3315
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefMain.frx":A6B4
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Save System Definition"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   4380
      End
      Begin VB.CheckBox chkContLCFail 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Continue on Leak Check Fail?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   74
         ToolTipText     =   "Using Continue on LeakCheck Fail?"
         Top             =   6930
         Width           =   3645
      End
      Begin VB.TextBox txtNrRcp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   4
         TabIndex        =   72
         Text            =   "0"
         ToolTipText     =   "Enter 1 to 999 recipes"
         Top             =   2175
         Width           =   735
      End
      Begin VB.TextBox txtNrCan 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   4
         TabIndex        =   70
         Text            =   "0"
         ToolTipText     =   "Enter 1 to 999 canisters"
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         MaxLength       =   60
         TabIndex        =   68
         Text            =   "desc"
         ToolTipText     =   "External Alarm Description"
         Top             =   7200
         Width           =   3735
      End
      Begin VB.CheckBox chkUPS 
         Caption         =   " Using UPS Monitoring?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   65
         ToolTipText     =   " Using UPS Monitoring?"
         Top             =   5160
         Width           =   2595
      End
      Begin VB.TextBox txtUPSType 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   64
         Text            =   "0"
         ToolTipText     =   "Type of UPS (0=none, 1=large; timed shutdown, 2=small; immediate shutdown)"
         Top             =   5130
         Width           =   735
      End
      Begin VB.Frame frmLogon 
         Caption         =   "Auto Logon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   705
         Left            =   10080
         TabIndex        =   62
         Top             =   3600
         Width           =   4385
         Begin VB.ComboBox AutoLogonOptions 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmSysDefMain.frx":ADB6
            Left            =   1680
            List            =   "frmSysDefMain.frx":ADC3
            Style           =   2  'Dropdown List
            TabIndex        =   63
            ToolTipText     =   "Alphanumeric Entry"
            Top             =   240
            Width           =   2445
         End
         Begin VB.Label lblAutoLogon 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   67
            Top             =   270
            Width           =   1815
         End
      End
      Begin VB.Frame frmLVolUnits 
         Caption         =   "Engr Units for Line Volume Calcs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   825
         Left            =   10080
         TabIndex        =   47
         Top             =   2700
         Width           =   4385
         Begin VB.OptionButton optLVolUnitsSI 
            Alignment       =   1  'Right Justify
            Caption         =   "Using SI Units"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   49
            ToolTipText     =   "Using SI Units (mm, meter))"
            Top             =   240
            Width           =   3885
         End
         Begin VB.OptionButton optLVolUnitsEngl 
            Alignment       =   1  'Right Justify
            Caption         =   "Using English Units"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   48
            ToolTipText     =   "Using English Units (inch, feet)"
            Top             =   480
            Width           =   3885
         End
      End
      Begin VB.Frame frmUnits 
         Caption         =   "Engr Units for Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   825
         Left            =   10080
         TabIndex        =   25
         Top             =   900
         Width           =   4385
         Begin VB.OptionButton optUnitsEngl 
            Alignment       =   1  'Right Justify
            Caption         =   "Using Degrees Fahrenheit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   27
            ToolTipText     =   "Using English Units (deg F)"
            Top             =   480
            Width           =   3885
         End
         Begin VB.OptionButton optUnitsSI 
            Alignment       =   1  'Right Justify
            Caption         =   "Using Degrees Centigrade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   26
            ToolTipText     =   "Using SI Units (deg C)"
            Top             =   240
            Width           =   3885
         End
      End
      Begin VB.CheckBox chkOOTPause 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Pause/Stop Station on OOT?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   24
         ToolTipText     =   "Using Pause Station on OOT?"
         Top             =   6690
         Width           =   3645
      End
      Begin VB.CheckBox chkStationTC 
         Caption         =   " Using 2 TC's per Station?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   " Using 2 TC's per Station?"
         Top             =   8295
         Width           =   3315
      End
      Begin VB.CheckBox chkPressPurge 
         Caption         =   " Using Positive Pressure Purge?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   " Using Positive Pressure Purge?"
         Top             =   9600
         Width           =   3315
      End
      Begin VB.CheckBox chkDoorOpen 
         Caption         =   " Using Door Open Sw?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   " Using Door Open Sw?"
         Top             =   4920
         Width           =   2685
      End
      Begin VB.CheckBox chkButaneMassLimit 
         Caption         =   " Using Butane Mass Limit?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   " Using Butane Mass Limit?"
         Top             =   9000
         Width           =   3315
      End
      Begin VB.CheckBox chkLoadPressure 
         Caption         =   " Using Load Pressure Limit?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   " Using Load Pressure Limit as a Station Alarm?"
         Top             =   8760
         Width           =   3315
      End
      Begin VB.CheckBox chkLineVolume 
         Caption         =   " Using Line Volume?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   " Using Line Volume?"
         Top             =   6360
         Width           =   3195
      End
      Begin VB.CheckBox chkCanVentAlarm 
         Caption         =   " Using Can Vent Alarm?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   " Using Can Vent Alarm Flow Switch?"
         Top             =   7800
         Width           =   3315
      End
      Begin VB.CheckBox chkCust12Contacts 
         Caption         =   " Using External Alarm Contacts?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   " Using Customer 1_2 Contacts?"
         Top             =   7200
         Width           =   3195
      End
      Begin VB.CheckBox chkCustLowGas 
         Caption         =   " Using Customer Low Gas Sw?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   " Using Customer Low Gas Sw?"
         Top             =   5640
         Width           =   3315
      End
      Begin VB.CheckBox chkLoadTimeLimit 
         Caption         =   " Using Load Time Limit?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   " Using Load Time Limit?"
         Top             =   9240
         Width           =   3315
      End
      Begin VB.CheckBox chkCommonTC 
         Caption         =   " Using Common TC's?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   " Using Common TC's?"
         Top             =   8055
         Width           =   3315
      End
      Begin VB.TextBox txtNrDummyStn 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "0"
         ToolTipText     =   "Enter 1 to 9 dummy (i.e. IO Only) stations"
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox txtNrRemoteScales 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "Enter 1 to 19 remote scales"
         Top             =   1185
         Width           =   735
      End
      Begin VB.TextBox txtNrScales 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Enter 1 to 19 scales"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtNrShift 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "1"
         ToolTipText     =   "Enter 1 to 4 shifts"
         Top             =   615
         Width           =   735
      End
      Begin VB.TextBox txtNrStn 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         MaxLength       =   1
         TabIndex        =   1
         Text            =   "1"
         ToolTipText     =   "Enter 1 to 9 stations"
         Top             =   330
         Width           =   735
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   555
         Left            =   9720
         TabIndex        =   79
         Top             =   9480
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   979
         _Version        =   393216
         Min             =   -1024
         Max             =   1024
         TickFrequency   =   100
      End
      Begin VB.Label lblWcmRecipe 
         BackStyle       =   0  'Transparent
         Caption         =   "WCM Recipe#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6570
         TabIndex        =   166
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label lbl2GmRecipe 
         BackStyle       =   0  'Transparent
         Caption         =   "2 Gram Recipe#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6570
         TabIndex        =   165
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label lblWBRange 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "degC Range min/max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6270
         TabIndex        =   147
         Top             =   4020
         Width           =   1875
      End
      Begin VB.Label lblHeaterComPort 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "WaterBath Chiller Com Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6270
         TabIndex        =   141
         Top             =   3330
         Width           =   2535
      End
      Begin VB.Label lblChillerTimeOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Chiller CommTimeout  in ms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6270
         TabIndex        =   140
         Top             =   3645
         Width           =   2535
      End
      Begin VB.Label lblWeakLiveFuelDensity 
         BackStyle       =   0  'Transparent
         Caption         =   """Weak"" LiveFuel density in gm/l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   128
         Top             =   9735
         Width           =   2910
      End
      Begin VB.Label lblDeadLiveFuelDensity 
         BackStyle       =   0  'Transparent
         Caption         =   """Dead"" LiveFuel density in gm/l"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   126
         Top             =   9450
         Width           =   2910
      End
      Begin VB.Label lblMaxSheathTempForADF 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Sheath Temp for ADF Drain"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   124
         Top             =   9165
         Width           =   2910
      End
      Begin VB.Label lblDefScaleMax 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Scale Max in grams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   118
         Top             =   510
         Width           =   2790
      End
      Begin VB.Label lblScaleWeightQueue 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Weights in Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   117
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Label lblAuxOutNum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   3
         Left            =   5715
         TabIndex        =   109
         Top             =   8325
         Width           =   135
      End
      Begin VB.Label lblAuxOutNum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   1
         Left            =   5715
         TabIndex        =   108
         Top             =   7605
         Width           =   135
      End
      Begin VB.Label lblAuxOutNum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   4
         Left            =   5715
         TabIndex        =   107
         Top             =   8685
         Width           =   135
      End
      Begin VB.Label lblAuxOutNum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   240
         Index           =   2
         Left            =   5715
         TabIndex        =   106
         Top             =   7965
         Width           =   135
      End
      Begin VB.Label lblScaleMaxUnstableReads 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Max Unstable Reads"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   97
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblAuxOutputs 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Outputs/Station"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3000
         TabIndex        =   95
         Top             =   7605
         Width           =   1815
      End
      Begin VB.Label lblMinDataLogInterval 
         BackStyle       =   0  'Transparent
         Caption         =   "Min Data Log Interval in sec"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   92
         Top             =   2400
         Width           =   2790
      End
      Begin VB.Label lblMinMfcSp 
         BackStyle       =   0  'Transparent
         Caption         =   "Min MFC SetPoint in %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   88
         Top             =   1770
         Width           =   2790
      End
      Begin VB.Label lblSequences 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Sequences"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   86
         ToolTipText     =   "Number of Master Job Sequences"
         Top             =   2460
         Width           =   2430
      End
      Begin VB.Label lblGramsPerLiter 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane Grams/Liter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   83
         Top             =   2085
         Width           =   2790
      End
      Begin VB.Label lblLoadEqlDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Eql Delay in sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   82
         ToolTipText     =   "Load Equalibrium Delay in seconds"
         Top             =   1455
         Width           =   2790
      End
      Begin VB.Label lblLoadMfcDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Load MFC Delay in sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   81
         Top             =   1140
         Width           =   2790
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H0000C6FE&
         Height          =   255
         Left            =   9720
         TabIndex        =   78
         Top             =   8760
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label lblRecipes 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Recipes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   73
         ToolTipText     =   "Number of Master Recipes"
         Top             =   2175
         Width           =   2430
      End
      Begin VB.Label lblCanisters 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Canisters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   71
         ToolTipText     =   "Number of Master Canister Definitions"
         Top             =   1890
         Width           =   2430
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "External Alarm Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   69
         Top             =   7215
         Width           =   2415
      End
      Begin VB.Label lblUPSType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "UPS Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   66
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblMsgDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "DelayBox Open Time in ms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   50
         Top             =   2715
         Width           =   2790
      End
      Begin VB.Label lblOPTOComPort 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OPTO Com Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   12
         Top             =   165
         Width           =   2790
      End
      Begin VB.Label lblNrDummyStn 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Dummy Stations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Number of Dummy (i.e. I/O Only) Stations"
         Top             =   1470
         Width           =   2430
      End
      Begin VB.Label lblNrRemoteScales 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Remote Scales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Number of Remote Scales (0-19)"
         Top             =   1185
         Width           =   2430
      End
      Begin VB.Label lblNrScales 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Scales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Number of Scales (0-19)"
         Top             =   900
         Width           =   2430
      End
      Begin VB.Label lblNrShift 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Shifts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Number of Shifts"
         Top             =   615
         Width           =   2430
      End
      Begin VB.Label lblNrStn 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Stations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Number of Stations"
         Top             =   330
         Width           =   2425
      End
      Begin VB.Label lblMFCSettleTime 
         BackStyle       =   0  'Transparent
         Caption         =   "MFC Settle Time in sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5970
         TabIndex        =   2
         Top             =   825
         Width           =   2790
      End
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   1800
      TabIndex        =   99
      Top             =   11100
      Width           =   11295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stations Logic Interval in msec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   51
      Top             =   9255
      Width           =   3015
   End
End
Attribute VB_Name = "frmSysDefMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private tmr As Integer
Private tmpVal As Single
Private GoingToAnotherSysdef As Boolean

Sub Refresh_SysDef()
Dim Idx As Integer
Dim flag As Boolean
    flag = IIf((chkAuxOutputs.Value = cYES), True, False)
    lblAuxOutputs.Enabled = flag
    txtAuxOutputs.Enabled = flag
    For Idx = txtAuxOutDesc.LBound To txtAuxOutDesc.UBound
        txtAuxOutDesc(Idx).Enabled = flag
        lblAuxOutNum(Idx).Enabled = flag
    Next Idx
    Select Case chkUPS.Value
        Case cNO
            lblUPSType.Enabled = False
            txtUPSType.Enabled = False
            txtUPSType.text = "0"
        Case cYES
            lblUPSType.Enabled = True
            txtUPSType.Enabled = True
    End Select
    Select Case chkCust12Contacts
        Case cNO
            lblDescription.Enabled = False
            txtDescription.Enabled = False
            If (Len(txtDescription.text) < 3) Then txtDescription.text = "External Alarm"
        Case cYES
            lblDescription.Enabled = True
            txtDescription.Enabled = True
    End Select
    
    Select Case chkRemTaskOrder.Value
        Case cNO
            lbl2GmRecipe.Enabled = False
            txt2GmRecipe.Enabled = False
            txt2GmRecipe.text = "1"
            lblWcmRecipe.Enabled = False
            txtWcmRecipe.Enabled = False
            txtWcmRecipe.text = "1"
        Case cYES
            lbl2GmRecipe.Enabled = True
            txt2GmRecipe.Enabled = True
            lblWcmRecipe.Enabled = True
            txtWcmRecipe.Enabled = True
    End Select
    
    Select Case chkTomCanLoad.Value
        Case cNO
            lbl2GmRecipe.Enabled = False
            txt2GmRecipe.Enabled = False
            txt2GmRecipe.text = "1"
            lblWcmRecipe.Enabled = False
            txtWcmRecipe.Enabled = False
            txtWcmRecipe.text = "1"
        Case cYES
            lbl2GmRecipe.Enabled = True
            txt2GmRecipe.Enabled = True
            lblWcmRecipe.Enabled = True
            txtWcmRecipe.Enabled = True
    End Select
    
    If ((chkRemTaskOrder.Value = cNO) And (chkRemStatusMon.Value = cNO)) Then
        chkRemAvlFiles.Value = cNO
        chkRemAvlFiles.Enabled = False
    Else
        chkRemAvlFiles.Enabled = True
        chkTomCanLoad.Enabled = True
    End If
'***************************************************************************************

    If (chkRemTaskOrder.Value = cYES) Then
        chkTomCanLoad.Enabled = False
    Else
        chkTomCanLoad.Enabled = True
    End If
    
    If (chkTomCanLoad.Value = cYES) Then
        chkRemTaskOrder.Enabled = False
    Else
        chkRemTaskOrder.Enabled = True
    End If
    
    lblMsg.Caption = " "
End Sub

Sub Update_SysDef()

    txtNrSeq.text = Format(NR_JOBSEQ, "0")
    txtNrPrg.text = Format(NR_PRGAIR, "0")
    txtNrStn.text = Format(NR_STN, "0")
    txtNrShift.text = Format(NR_SHIFT, "0")
    txtNrScales.text = Format(NR_SCALES, "#0")
    txtNrRemoteScales.text = Format(NR_REMOTESCALES, "#0")
    txtNrDummyStn.text = Format(NR_DUMMYSTN, "0")
    txtNrCan.text = Format(NR_CAN, "###0")
    txtNrRcp.text = Format(NR_RCP, "###0")
    txtWbEuMax.text = Format(WB_AIO.EuMax, "###0.0##")
    txtWbEuMin.text = Format(WB_AIO.EuMin, "###0.0##")
    
    chkWaterBath.Value = IIf(USINGWATERBATH, cYES, cNO)
    txtOPTOComPort.text = Format(OPTOCOM_PORT, "0")

    txtChillerComPort.text = Format(Chiller_PORT, "0")
    txtChillerCommTimeout.text = Format(Chiller_Timeout, "####0")
    
    txtLoadMfcDelay.text = Format(LoadMfcDelayTime, "#0")
    txtLoadEqlDelay.text = Format(LoadEqlDelayTime, "##0")

    txtMinMfcSp.text = Format(MfcSpMin, "0.0##")
    txtMFCSettleTime.text = Format(MFC_Settle_Time, "#0")

    txtMsgDelay.text = Format(MSGDELAY, "###0")
    
    txtGramsPerLiter.text = Format(GramsPerLiter, "#.#####")

    txtDefScaleMax.text = Format(DefScaleMax, "###,##0")
    txtScaleMaxUnstableReads.text = Format(MAXNOTSTABLECOUNT, "###0")
    txtScaleWeightQueue.text = Format(WEIGHTQUEUESIZE, "###0")
    txtMinDataLogInterval.text = Format(MinDataLogSeconds, "##0.00")

    txtUPSType.text = Format(USINGUPS, "0")
    
    If USINGC Then optUnitsSI.Value = True
    If USINGF Then optUnitsEngl.Value = True
    
    If USINGMoist_RH Then optMoistRH.Value = True
    If USINGMoist_Grains Then optMoistGrains.Value = True
    
    If USINGLVol_SI Then optLVolUnitsSI.Value = True
    If USINGLVol_Engl Then optLVolUnitsEngl.Value = True
    
    chkEstopInput.Value = IIf(USING_ESTOP_INPUT, cYES, cNO)
    
    ' scales hard-piped ??
    chkHardPipedScales.Value = IIf(USINGHARDPIPEDSCALES, cYES, cNO)
    
    ' PAG Control
    If (LocalPagControl.Type = pagAlone) Then
        optAlone.Value = True
    ElseIf (LocalPagControl.Type = pagMaster) Then
        optMaster.Value = True
    ElseIf (LocalPagControl.Type = pagClient) Then
        optClient.Value = True
    Else
        optNone.Value = True
    End If
    txtServerIp.text = PAGSERVERIP
    
    ' Temp/Rh Logging
    chkLogTempRh.Value = IIf(LogTempRh, cYES, cNO)
    ' High Temperature PAS
    chkHighTempPas.Value = IIf(USINGHIGHTEMPPAS, cYES, cNO)
    
    ' REM Can Load Tasks
    chkRemTaskOrder.Value = IIf(USINGREMCANLOAD, cYES, cNO)
    chkTomCanLoad.Value = IIf(USINGTOMCANLOAD, cYES, cNO)
    chkRemStatusMon.Value = IIf(USINGREMSTSMON, cYES, cNO)
    chkRemEnableChgs.Value = IIf(REMCHGSENABLED, cYES, cNO)
    chkRemAvlFiles.Value = IIf((USINGREMCANLOAD Or USINGREMSTSMON), IIf(USINGREMAVLFILES, cYES, cNO), cNO)
    
    'REMCHGSENABLED
    txt2GmRecipe.text = Format(TOM_2Gm_Recipe, "##0")
    txtWcmRecipe.text = Format(TOM_Wcm_Recipe, "##0")
    ' Max Sheath Temp for ADF
    txtMaxSheathTempForADF.text = Format(MaxSheathTempForAdfDrain, "#,##0.0##")
    ' "Dead" LiveFuel density (triggers tank refill immediately)
    txtDeadLiveFuelDensity.text = Format(DeadLiveFuelDensity, "##0.0##")
    ' "Weak" LiveFuel density (triggers tank refill before next Load)
    txtWeakLiveFuelDensity.text = Format(WeakLiveFuelDensity, "##0.0##")
    
    chkDoorOpen.Value = IIf(USINGDOOROPEN, cYES, cNO)
    chkButaneMassLimit.Value = IIf(USINGBUTANEMASSLIMIT, cYES, cNO)
    chkLoadTimeLimit.Value = IIf(USINGLOADTIMELIMIT, cYES, cNO)
    chkCustLowGas.Value = IIf(USINGCUSTOMERLOWGAS, cYES, cNO)
    chkCust12Contacts.Value = IIf(USING_EXT_CONTACTS, cYES, cNO)
    txtDescription.text = IIf(Len(DESC_EXT_CONTACTS) > 0, DESC_EXT_CONTACTS, "External Alarm")
    
    chkAuxOutputs.Value = IIf(USING_AUX_OUTPUTS, cYES, cNO)
    txtAuxOutputs.text = Format(NR_AUX_OUTPUTS, "0")
    txtAuxOutDesc(1).text = IIf(Len(DESC_AUX_OUTPUT1) > 0, DESC_AUX_OUTPUT1, "Aux Output1")
    txtAuxOutDesc(2).text = IIf(Len(DESC_AUX_OUTPUT2) > 0, DESC_AUX_OUTPUT2, "Aux Output2")
    txtAuxOutDesc(3).text = IIf(Len(DESC_AUX_OUTPUT3) > 0, DESC_AUX_OUTPUT3, "Aux Output3")
    txtAuxOutDesc(4).text = IIf(Len(DESC_AUX_OUTPUT4) > 0, DESC_AUX_OUTPUT4, "Aux Output4")
        
    chkCanVentAlarm.Value = IIf(USINGCANVENTALARM, cYES, cNO)
    chkLineVolume.Value = IIf(USINGLINEVOLUME, cYES, cNO)
    chkLoadPressure.Value = IIf(USINGLOADPRESSURE, cYES, cNO)
    
    chkSystemVacSw.Value = IIf(USINGSYSTEMVACSW, cYES, cNO)
    
    chkAuxLeakChk.Value = IIf(USINGAUXLEAKCHECK, cYES, cNO)
    chkLeakChkExhSol.Value = IIf(USINGLEAKCHECKEXHAUSTSOL, cYES, cNO)

    chkPressPurge.Value = IIf(USINGPRESSUREPURGE, cYES, cNO)
    chkPurgeDP.Value = IIf(USINGPURGEDP, cYES, cNO)
    chkPurgeOven.Value = IIf(USINGPURGEOVEN, cYES, cNO)
    chkPurgeSeries.Value = IIf(USINGPURGESERIES, cYES, cNO)
    chkUsingDryPurgeAir.Value = IIf(USINGDRYPURGEAIR, cYES, cNO)

    chkCommonTC.Value = IIf(USINGCOMMONTC, cYES, cNO)
    chkStationTC.Value = IIf(USINGSTNTC, cYES, cNO)
    
    chkOOTPause.Value = IIf(USINGOOTPAUSE, cYES, cNO)
    chkContLCFail.Value = IIf(USINGCONTAFTERLCFAIL, cYES, cNO)
    chkUseFuelLevelOot.Value = IIf(USINGFUELLEVELOOT, cYES, cNO)
    
    chkErrorMsgBypass.Value = IIf(USINGERRORMSGBYPASS, cYES, cNO)
    
    AutoLogonOptions.ListIndex = AutoLogon

    
    If CheckPass("0", False) Then
        ' Show Debug Options on this screen
        frmDebug.Visible = True
        chkScanIO.Value = IIf(IoComOn, 0, 1)
        chkReadScales.Value = IIf(SclComOn, 0, 1)
        chkChillerComm.Value = IIf(ChillComOn, 0, 1)
        chkDebug.Value = IIf(UseLocalErrorHandler, 0, 1)
        chkADFdebug.Value = IIf(NotDebugADF, 0, 1)
        chkMmwDebug.Value = IIf(NotDebugMMW, 0, 1)
        chkPasDebug.Value = IIf(NotDebugPAS, 0, 1)
        chkPrgDebug.Value = IIf(NotDebugPURGE, 0, 1)
        chkSclDebug.Value = IIf(NotDebugSCALES, 0, 1)
        chkScanIO.Visible = True
        chkReadScales.Visible = True
        chkDebug.Visible = True
        chkADFdebug.Visible = True
        chkMmwDebug.Visible = True
        chkVerboseStartup.Value = IIf(STARTUPVERBOSE, 1, 0)
        chkVerboseStartup.Visible = True
        ' System Timer Settings
        For tmr = 1 To 9
            txtTmrInterval(tmr).text = Format(SystemTimers(tmr).Interval, "###0")
        Next tmr
    Else
        ' Dont show debug options
        frmDebug.Visible = False
    End If


End Sub

Private Sub chkAuxOutputs_Click()
    Refresh_SysDef
End Sub

Private Sub chkCust12Contacts_Click()
    Refresh_SysDef
End Sub

Private Sub chkHardPipedScales_Click()
    Refresh_SysDef
End Sub

Private Sub chkRemStatusMon_Click()
    Refresh_SysDef
End Sub
Private Sub chkTomCanLoad_Click()
    Refresh_SysDef
End Sub
Private Sub chkRemTaskOrder_Click()
    Refresh_SysDef
End Sub

Private Sub chkUPS_Click()
    If chkUPS.Value = cYES Then
        If Not IsNumeric(txtUPSType.text) Then txtUPSType.text = "0"
        If CInt(txtUPSType.text) = 0 Then
            txtUPSType.BackColor = frmHighlight.BackColor
            txtUPSType.ForeColor = frmHighlight.ForeColor
        End If
    Else
        txtUPSType.BackColor = frmNotHighlight.BackColor
        txtUPSType.ForeColor = frmNotHighlight.ForeColor
    End If
    Refresh_SysDef
End Sub

Private Sub cmdDisplayProperties_Click()
    frmDisplayProperties.Show
    Unload Me
End Sub

Private Sub cmdMonTimers_Click()
    frmTmrMonitor.Show
End Sub

Private Sub cmdNext_Click()
    frmSysDefStn.Show
    GoingToAnotherSysdef = True
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    lblMsg.Caption = ""
    Set pbCapture.Picture = CaptureForm(Me)
    PrintPictureToFitPage Printer, pbCapture.Picture
    Printer.EndDoc
    Set pbCapture.Picture = Nothing
    lblMsg.Font.Size = 9.5
    lblMsg.ForeColor = Message_ForeColor
    lblMsg.Caption = "Sysdef sent to" & vbCrLf & PRINTERNAME
End Sub

Private Sub cmdSave_Click()

    NR_PRGAIR = CInt(txtNrPrg.text)
    NR_STN = CInt(txtNrStn.text)
    NR_SHIFT = CInt(txtNrShift.text)
    NR_SCALES = CInt(txtNrScales.text)
    NR_REMOTESCALES = CInt(txtNrRemoteScales.text)
    NR_DUMMYSTN = CInt(txtNrDummyStn.text)
    NR_CAN = CInt(txtNrCan.text)
    NR_RCP = CInt(txtNrRcp.text)
    NR_JOBSEQ = CInt(txtNrSeq.text)
    
    OPTOCOM_PORT = CInt(txtOPTOComPort.text)
    
    USINGWATERBATH = IIf(chkWaterBath.Value = cYES, True, False)
    Chiller_PORT = CInt(txtChillerComPort.text)
    Chiller_Timeout = CInt(txtChillerCommTimeout.text)
    WB_AIO.EuMax = ValueFromText(txtWbEuMax.text)
    WB_AIO.EuMin = ValueFromText(txtWbEuMin.text)

    MSGDELAY = CInt(txtMsgDelay.text)

    LoadMfcDelayTime = CInt(txtLoadMfcDelay.text)
    LoadEqlDelayTime = CInt(txtLoadEqlDelay.text)

    If Not IsNumeric(txtMinMfcSp.text) Then txtMinMfcSp.text = "5.0"    ' default MFC SetPoint Minimum = 5.0 % of Fullscale
    If CSng(txtMinMfcSp.text) < 1# Then txtMinMfcSp.text = "5.0"        ' default MFC SetPoint Minimum = 5.0 % of Fullscale
    If CSng(txtMinMfcSp.text) > 5# Then txtMinMfcSp.text = "5.0"        ' default MFC SetPoint Minimum = 5.0 % of Fullscale
    MfcSpMin = CSng(txtMinMfcSp.text)
    MFC_Settle_Time = CInt(txtMFCSettleTime.text)

    ' Master Butane Density in Grams/Liter
    '   max     2.7
    '   default 2.40633
    '   min     2.1
    If Not IsNumeric(txtGramsPerLiter.text) Then txtGramsPerLiter.text = "2.40633"
    If CSng(txtGramsPerLiter.text) > 2.7 Then txtGramsPerLiter.text = "2.7"
    If CSng(txtGramsPerLiter.text) < 2.1 Then txtGramsPerLiter.text = "2.1"
    '   Log any Master Butane Density Changes
    If CSng(txtGramsPerLiter.text) <> GramsPerLiter Then
        Write_ELog "Master Butane Density Changed to " & txtGramsPerLiter.text & " grams/liter"
    End If
    '   Master Butane Density
    GramsPerLiter = CSng(txtGramsPerLiter.text)
    
    DefScaleMax = ValueFromText(txtDefScaleMax.text)
    MAXNOTSTABLECOUNT = CInt(txtScaleMaxUnstableReads.text)
    WEIGHTQUEUESIZE = CInt(txtScaleWeightQueue.text)
    MinDataLogSeconds = ValueFromText(txtMinDataLogInterval.text)
    
    ' Max Sheath Temp for ADF (can't drain if sheath temp too high)
    MaxSheathTempForAdfDrain = ValueFromText(txtMaxSheathTempForADF.text)
    ' "Dead" LiveFuel density for ADF (density too low triggers a tank refill immediately)
    DeadLiveFuelDensity = ValueFromText(txtDeadLiveFuelDensity.text)
    ' "Weak" LiveFuel density for ADF (density too low triggers a tank refill before next Load)
    WeakLiveFuelDensity = ValueFromText(txtWeakLiveFuelDensity.text)
    
    USINGUPS = CInt(txtUPSType.text)
    
    If optUnitsSI.Value Then
        USINGC = True
        USINGF = False
    End If
    If optUnitsEngl.Value Then
        USINGC = False
        USINGF = True
    End If
    
    If optMoistRH.Value Then
        USINGMoist_RH = True
        USINGMoist_Grains = False
    End If
    If optMoistGrains.Value Then
        USINGMoist_RH = False
        USINGMoist_Grains = True
    End If
    
    If optLVolUnitsSI.Value Then
        USINGLVol_SI = True
        USINGLVol_Engl = False
    End If
    If optLVolUnitsEngl.Value Then
        USINGLVol_SI = False
        USINGLVol_Engl = True
    End If
    
    USING_ESTOP_INPUT = IIf(chkEstopInput.Value = cYES, True, False)

    USINGHARDPIPEDSCALES = IIf(chkHardPipedScales.Value = cYES, True, False)

    If optAlone.Value Then
        LocalPagControl.Type = pagAlone
    ElseIf optMaster.Value Then
        LocalPagControl.Type = pagMaster
    ElseIf optClient.Value Then
        LocalPagControl.Type = pagClient
    Else
        LocalPagControl.Type = pagNone
    End If
    USINGPASLOCALCONTROL = IIf(((LocalPagControl.Type = pagAlone) Or (LocalPagControl.Type = pagMaster)), True, False)
    PAGSERVERIP = IIf((Len(txtServerIp.text) > 3), txtServerIp.text, " ")
    
    LogTempRh = IIf(chkLogTempRh.Value = cYES, True, False)
    ' High Temperature PAS
    USINGHIGHTEMPPAS = IIf(chkHighTempPas.Value = cYES, True, False)
    ' Remote Tasks & Status
    USINGREMCANLOAD = IIf(chkRemTaskOrder.Value = cYES, True, False)
    USINGTOMCANLOAD = IIf(chkTomCanLoad.Value = cYES, True, False)
    
    USINGREMSTSMON = IIf(chkRemStatusMon.Value = cYES, True, False)
    REMCHGSENABLED = IIf(chkRemEnableChgs.Value = cYES, True, False)
    USINGREMAVLFILES = IIf((USINGREMCANLOAD Or USINGREMCANLOAD), IIf(chkRemAvlFiles.Value = cYES, True, False), False)
    
    If Not IsNumeric(txt2GmRecipe.text) Then txt2GmRecipe.text = "92"
    If Not IsNumeric(txtWcmRecipe.text) Then txtWcmRecipe.text = "93"
    TOM_2Gm_Recipe = CInt(ValueFromText(txt2GmRecipe.text))
    TOM_Wcm_Recipe = CInt(ValueFromText(txtWcmRecipe.text))
    
    USINGDOOROPEN = IIf(chkDoorOpen.Value = cYES, True, False)
    USINGBUTANEMASSLIMIT = IIf(chkButaneMassLimit.Value = cYES, True, False)
    USINGLOADTIMELIMIT = IIf(chkLoadTimeLimit.Value = cYES, True, False)
    USINGCUSTOMERLOWGAS = IIf(chkCustLowGas.Value = cYES, True, False)
    USING_EXT_CONTACTS = IIf(chkCust12Contacts.Value = cYES, True, False)
    DESC_EXT_CONTACTS = IIf(Len(txtDescription.text) > 0, txtDescription.text, "External Alarm")

    USING_AUX_OUTPUTS = IIf(chkAuxOutputs.Value = cYES, True, False)
    NR_AUX_OUTPUTS = CInt(ValueFromText(txtAuxOutputs.text))
    DESC_AUX_OUTPUT1 = IIf(Len(txtAuxOutDesc(1).text) > 0, txtAuxOutDesc(1).text, "Aux Output1")
    DESC_AUX_OUTPUT2 = IIf(Len(txtAuxOutDesc(2).text) > 0, txtAuxOutDesc(2).text, "Aux Output2")
    DESC_AUX_OUTPUT3 = IIf(Len(txtAuxOutDesc(3).text) > 0, txtAuxOutDesc(3).text, "Aux Output3")
    DESC_AUX_OUTPUT4 = IIf(Len(txtAuxOutDesc(4).text) > 0, txtAuxOutDesc(4).text, "Aux Output4")
        
    USINGCANVENTALARM = IIf(chkCanVentAlarm.Value = cYES, True, False)
    USINGLINEVOLUME = IIf(chkLineVolume.Value = cYES, True, False)
    USINGLOADPRESSURE = IIf(chkLoadPressure.Value = cYES, True, False)

    USINGSYSTEMVACSW = IIf(chkSystemVacSw.Value = cYES, True, False)
    
    USINGAUXLEAKCHECK = IIf(chkAuxLeakChk.Value = cYES, True, False)
    USINGLEAKCHECKEXHAUSTSOL = IIf(chkLeakChkExhSol.Value = cYES, True, False)

    USINGPRESSUREPURGE = IIf(chkPressPurge.Value = cYES, True, False)
    USINGPURGEDP = IIf(chkPurgeDP.Value = cYES, True, False)
    USINGPURGEOVEN = IIf(chkPurgeOven.Value = cYES, True, False)
    USINGPURGESERIES = IIf(chkPurgeSeries.Value = cYES, True, False)
    USINGDRYPURGEAIR = IIf(chkUsingDryPurgeAir.Value = cYES, True, False)

    USINGCOMMONTC = IIf(chkCommonTC.Value = cYES, True, False)
    USINGSTNTC = IIf(chkStationTC.Value = cYES, True, False)
    
    USINGOOTPAUSE = IIf(chkOOTPause.Value = cYES, True, False)
    USINGCONTAFTERLCFAIL = IIf(chkContLCFail.Value = cYES, True, False)
    USINGFUELLEVELOOT = IIf(chkUseFuelLevelOot.Value = cYES, True, False)

    USINGERRORMSGBYPASS = IIf(chkErrorMsgBypass.Value = cYES, True, False)
    
    AutoLogon = AutoLogonOptions.ListIndex
    
    If CheckPass("0", False) Then
        ' Save Debug Options from this screen
        IoComOn = IIf(chkScanIO.Value = cYES, False, True)
        SclComOn = IIf(chkReadScales.Value = cYES, False, True)
        ChillComOn = IIf(chkChillerComm.Value = cYES, False, True)
        UseLocalErrorHandler = IIf(chkDebug.Value = cYES, False, True)
        NotDebugADF = IIf(chkADFdebug.Value = cYES, False, True)
        NotDebugMMW = IIf(chkMmwDebug.Value = cYES, False, True)
        NotDebugPAS = IIf(chkPasDebug.Value = cYES, False, True)
        NotDebugPURGE = IIf(chkPrgDebug.Value = cYES, False, True)
        NotDebugSCALES = IIf(chkSclDebug.Value = cYES, False, True)
        STARTUPVERBOSE = IIf(chkVerboseStartup.Value = cYES, True, False)
        ' System Timer Settings
        For tmr = 1 To 9
            SystemTimers(tmr).Interval = CInt(txtTmrInterval(tmr).text)
            If ReadyToRun Then
                ' Already running, so update the running timers
                frmMainMenu.tmrTimer(tmr).Interval = CInt(txtTmrInterval(tmr).text)
            End If
        Next tmr
    End If

    Save_StartupSettings
    Save_SysDef
    Load_SysDef
    Load_StartupSettings
    Update_SysDef
    
    lblMsg.Caption = vbCrLf & "System Definition File Saved"

End Sub

Private Sub cmdSimCntrlPnl_Click()
    frmSimCntrlPnl.Show
End Sub

Private Sub Form_Load()
    GoingToAnotherSysdef = False
    If Not ReadyToRun Then
        cmdPrint.Visible = False
'        lblMsg.ForeColor = lblMsg.BackColor
    Else
        cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
'        lblMsg.ForeColor = Message_ForeColor
    End If
        lblMsg.ForeColor = Message_ForeColor
    
    ' Set Title Foreground color
    frmSystemSettings.ForeColor = Titles_ForeColor
    frmUnits.ForeColor = Titles_ForeColor
    frmMoistUnits.ForeColor = Titles_ForeColor
    frmLVolUnits.ForeColor = Titles_ForeColor
    frmLogon.ForeColor = Titles_ForeColor
    frmDebug.ForeColor = Titles_ForeColor
    
    lblMsg.Caption = " "
    
    Form_Center Me

    If (USINGUPS > 0) Then chkUPS.Value = cYES
    Update_SysDef
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not GoingToAnotherSysdef Then ReadyToRun = True
    Unload Me
End Sub

Private Sub optAlone_Click()
    lblServerIp.Visible = False
    txtServerIp.Visible = False
End Sub

Private Sub optClient_Click()
    lblServerIp.Visible = True
    txtServerIp.Visible = True
End Sub

Private Sub optMaster_Click()
    lblServerIp.Visible = False
    txtServerIp.Visible = False
End Sub

Private Sub optNone_Click()
    lblServerIp.Visible = False
    txtServerIp.Visible = False
End Sub

Private Sub tmrScreen_Timer()
    lblColor.BackColor = Abs(Slider1.Value + Slider2.Value)
    lblColor.Caption = Format(Slider1.Value + Slider2.Value, "#######0")
End Sub

Private Sub txtAuxOutputs_Change()
Dim Idx As Integer
    If IsNumeric(txtAuxOutputs.text) Then
        If (CInt(txtAuxOutputs.text) < 0) Then txtAuxOutputs.text = "0"
        If (CInt(txtAuxOutputs.text) > 4) Then txtAuxOutputs.text = "4"
        For Idx = txtAuxOutDesc.LBound To txtAuxOutDesc.UBound
            txtAuxOutDesc(Idx).Enabled = IIf((CInt(txtAuxOutputs.text) < Idx), False, True)
            lblAuxOutNum(Idx).Enabled = IIf((CInt(txtAuxOutputs.text) < Idx), False, True)
        Next Idx
    End If
End Sub

Private Sub txtDescription_Change()
    If (Len(txtDescription.text) < 1) Then txtDescription.text = "External Alarm"
End Sub

Private Sub txtScaleWeightQueue_Change()
    If (ValueFromText(txtScaleWeightQueue.text) < CSng(1)) Then txtScaleWeightQueue.text = "1"
End Sub

Private Sub txtUPSType_Change()
    txtUPSType.BackColor = frmNotHighlight.BackColor
    txtUPSType.ForeColor = frmNotHighlight.ForeColor
End Sub
