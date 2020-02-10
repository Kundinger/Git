VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLeakTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LeakTest Station"
   ClientHeight    =   11130
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14970
   Icon            =   "frmLeakTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11130
   ScaleMode       =   0  'User
   ScaleWidth      =   13266.52
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMsg 
      Height          =   425
      Left            =   120
      TabIndex        =   115
      Top             =   9120
      Width           =   14655
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   120
         Width           =   14415
      End
   End
   Begin VB.Frame frmLTrcp 
      Caption         =   "Recipe"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   88
      Top             =   7560
      Width           =   5055
      Begin VB.CommandButton cmdSaveRcp 
         Height          =   455
         Left            =   4440
         Picture         =   "frmLeakTest.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   840
         Width           =   455
      End
      Begin VB.TextBox txtDeffDur 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   112
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtTargetPress 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   109
         Text            =   "0"
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label lblDeffDurUnits 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   111
         Top             =   735
         Width           =   900
      End
      Begin VB.Label lblDeffDur 
         Caption         =   "Deff Stable Duration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   735
         Width           =   2295
      End
      Begin VB.Label lblTargetPressUnits 
         Caption         =   "kPa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   108
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblTargetPress 
         Caption         =   "Target Pressure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frmLTcfg 
      Caption         =   "Configuration"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   87
      Top             =   4320
      Width           =   5055
      Begin VB.TextBox txtReportInterval 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   126
         Text            =   "300"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveCfg 
         Height          =   455
         Left            =   4440
         Picture         =   "frmLeakTest.frx":5EE4
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   2400
         Width           =   455
      End
      Begin VB.TextBox txtlIntN2Flow 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   106
         Text            =   "0"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtDeffTol 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   103
         Text            =   "0"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtStablePressDur 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   100
         Text            =   "0"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtPressTol 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   97
         Text            =   "0"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtPressTimeout 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   94
         Text            =   "300"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtTimeout 
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
         ForeColor       =   &H00C0C000&
         Height          =   285
         Left            =   2640
         TabIndex        =   91
         Text            =   "2400"
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label lblReportInterval 
         Caption         =   "Report Interval"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   128
         ToolTipText     =   "1 to 9 seconds"
         Top             =   2535
         Width           =   2295
      End
      Begin VB.Label lblReportIntervalUnits 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   127
         Top             =   2535
         Width           =   420
      End
      Begin VB.Label lbllIntN2FlowUnits 
         Caption         =   "m3/s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   105
         Top             =   2175
         Width           =   900
      End
      Begin VB.Label lblIntN2Flow 
         Caption         =   "Initial N2 Flow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   104
         Top             =   2175
         Width           =   2295
      End
      Begin VB.Label lblDeffTolUnits 
         Caption         =   "inch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   102
         Top             =   1815
         Width           =   900
      End
      Begin VB.Label lblDeffTol 
         Caption         =   "Deff Tolerance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   101
         Top             =   1815
         Width           =   2295
      End
      Begin VB.Label lblStablePressDurUnits 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   99
         Top             =   1455
         Width           =   900
      End
      Begin VB.Label lblStablePressDur 
         Caption         =   "Pressure Stable Duration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   1455
         Width           =   2295
      End
      Begin VB.Label lblPressTolUnits 
         Caption         =   "kPa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   96
         Top             =   1095
         Width           =   900
      End
      Begin VB.Label lblPressTol 
         Caption         =   "Pressure Tolerance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   95
         Top             =   1095
         Width           =   2295
      End
      Begin VB.Label lblVUnits 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   93
         Top             =   735
         Width           =   900
      End
      Begin VB.Label lblPressTimeout 
         Caption         =   "Pressurize Timeout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   735
         Width           =   2295
      End
      Begin VB.Label lblTimeoutUnits 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3960
         TabIndex        =   90
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblTimeout 
         Caption         =   "Timeout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frmLeakTestSeq 
      Caption         =   "LeakTest Sequence"
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
      Height          =   1215
      Left            =   5400
      TabIndex        =   80
      Top             =   4320
      Width           =   9375
      Begin VB.CommandButton cmdLoadControllers 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8595
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmLeakTest.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Reload Controller Setup & Tuning Parameters"
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.Label lblStepVal 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   86
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblStep 
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   85
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblTaskVal 
         Caption         =   "idle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   84
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblTask 
         Caption         =   "Task"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   83
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblDescVal 
         Caption         =   "idle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   82
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblDesc 
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   81
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame frmDeff 
      Caption         =   "Effective Diameter Calculations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   3255
      Left            =   5400
      TabIndex        =   61
      Top             =   5760
      Width           =   9375
      Begin VB.TextBox txtSgN2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4320
         TabIndex        =   62
         Text            =   "0.967"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblCalcMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "calc msg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   2280
         Width           =   9135
      End
      Begin VB.Label lblLBL 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   7680
         TabIndex        =   130
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblTXT 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Index           =   4
         Left            =   8100
         TabIndex        =   129
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblTXT 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Index           =   3
         Left            =   8100
         TabIndex        =   125
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblLBL 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   7680
         TabIndex        =   124
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lblTXT 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Index           =   2
         Left            =   8100
         TabIndex        =   123
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblLBL 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   7680
         TabIndex        =   122
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblTXT 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Index           =   1
         Left            =   8100
         TabIndex        =   121
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLBL 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   120
         Top             =   480
         Width           =   480
      End
      Begin VB.Label lblTXT 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Index           =   0
         Left            =   8100
         TabIndex        =   119
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLBL 
         Caption         =   "sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   118
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblEffLeakDia 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   4320
         TabIndex        =   79
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblDeff 
         Caption         =   "Effective Leak Diameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   720
         TabIndex        =   78
         Top             =   2640
         Width           =   3390
      End
      Begin VB.Label lblDeffUnits 
         Caption         =   "inches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   5760
         TabIndex        =   77
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label lblInletPress 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4320
         TabIndex        =   76
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label lblInletPressDesc 
         Caption         =   "Inlet Pressure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   75
         Top             =   1260
         Width           =   3390
      End
      Begin VB.Label lblInletPressUnits 
         Caption         =   "kPa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   74
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label lblPatm 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4320
         TabIndex        =   73
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lblPatmDesc 
         Caption         =   "Atmospheric Pressure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   72
         Top             =   1590
         Width           =   3390
      End
      Begin VB.Label lblPatmUnits 
         Caption         =   "kPa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   71
         Top             =   1590
         Width           =   1335
      End
      Begin VB.Label lblSgN2Desc 
         Caption         =   "Specific Gravity of N2 relative to air"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   70
         Top             =   1920
         Width           =   3390
      End
      Begin VB.Label lblSgN2Units 
         Caption         =   "at 101.325kPa and 15.5 degC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   69
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label lblN2Temp 
         Alignment       =   2  'Center
         Caption         =   "0.00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4320
         TabIndex        =   68
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label lblN2TempDesc 
         Caption         =   "Nitrogen Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   67
         Top             =   930
         Width           =   3390
      End
      Begin VB.Label lblN2TempUnits 
         Caption         =   "degK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   66
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label lblN2Flow 
         Alignment       =   2  'Center
         Caption         =   "0.000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   4320
         TabIndex        =   65
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblN2FlowDesc 
         Caption         =   "Nitrogen Vol Flow Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   64
         Top             =   600
         Width           =   3390
      End
      Begin VB.Label lblN2FlowUnits 
         Caption         =   "m3/s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   63
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.PictureBox pbxTop 
      Align           =   1  'Align Top
      Height          =   2640
      Left            =   0
      ScaleHeight     =   2580
      ScaleWidth      =   14910
      TabIndex        =   11
      Top             =   600
      Width           =   14970
      Begin VB.TextBox Text11 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   960
         Width           =   2000
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   240
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   47
         Text            =   "Text3"
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   0
         Width           =   2000
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   720
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   44
         Text            =   "Text6"
         Top             =   1200
         Width           =   1995
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   43
         Text            =   "Text7"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox txtStartOp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12480
         MaxLength       =   25
         TabIndex        =   42
         ToolTipText     =   "Alphanumeric Name"
         Top             =   1845
         Width           =   2190
      End
      Begin VB.Frame frmStnCanister 
         Caption         =   "Canister"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1200
         Left            =   10800
         TabIndex        =   33
         Top             =   480
         Width           =   4045
         Begin VB.TextBox txtCanID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   120
            MaxLength       =   25
            TabIndex        =   35
            Text            =   "vehicle"
            Top             =   210
            Width           =   3815
         End
         Begin VB.TextBox txtLeakCheckStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   60
            MaxLength       =   25
            TabIndex        =   34
            Text            =   "leakcheck results"
            ToolTipText     =   "Station Canister Description"
            Top             =   450
            Visible         =   0   'False
            Width           =   3930
         End
         Begin VB.Label lblWorkCap 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2280
            TabIndex        =   41
            ToolTipText     =   "Current Canister Size in grams"
            Top             =   915
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblBedVolume 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2280
            TabIndex        =   40
            ToolTipText     =   "Current Canister Size in liters"
            Top             =   690
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lblCanWcUnits 
            Caption         =   "grams"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3240
            TabIndex        =   39
            Top             =   915
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblCanVolUnits 
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3240
            TabIndex        =   38
            Top             =   690
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Canister Volume:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   37
            Top             =   690
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Canister Work Cap:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   36
            Top             =   915
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.Frame frmStnRecipe 
         Caption         =   "Recipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1200
         Left            =   0
         TabIndex        =   28
         Top             =   480
         Width           =   4800
         Begin VB.TextBox txtRcpDsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            MaxLength       =   70
            TabIndex        =   32
            Text            =   "cycles"
            Top             =   450
            Visible         =   0   'False
            Width           =   4560
         End
         Begin VB.TextBox txtRcpDsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   360
            MaxLength       =   70
            TabIndex        =   31
            Text            =   "load"
            Top             =   915
            Visible         =   0   'False
            Width           =   4320
         End
         Begin VB.TextBox txtRcpDsc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   360
            MaxLength       =   70
            TabIndex        =   30
            Text            =   "purge"
            Top             =   690
            Visible         =   0   'False
            Width           =   4320
         End
         Begin VB.TextBox txtRecipeName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   120
            MaxLength       =   70
            TabIndex        =   29
            Text            =   "40 cfr 1066.985 leak test"
            Top             =   210
            Width           =   4575
         End
      End
      Begin VB.TextBox txtVehicle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   27
         Text            =   "1234567890123456789012345"
         ToolTipText     =   "Alphanumeric Vehicle Identification Number"
         Top             =   1845
         Width           =   2760
      End
      Begin VB.TextBox txtEngineer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1650
         MaxLength       =   25
         TabIndex        =   26
         ToolTipText     =   "Alphanumeric Name"
         Top             =   2130
         Width           =   2760
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   25
         Text            =   "Text8"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   17880
         TabIndex        =   24
         Text            =   "Text9"
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton cmdUpStn 
         DisabledPicture =   "frmLeakTest.frx":6928
         DownPicture     =   "frmLeakTest.frx":756A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmLeakTest.frx":81AC
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdDnStn 
         DisabledPicture =   "frmLeakTest.frx":8DEE
         DownPicture     =   "frmLeakTest.frx":9A30
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmLeakTest.frx":A672
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1710
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox txtDspStn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   630
         Left            =   6150
         TabIndex        =   19
         Text            =   "8"
         Top             =   1710
         Width           =   415
      End
      Begin VB.TextBox txtStation 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   5985
         TabIndex        =   18
         Text            =   "station"
         Top             =   2295
         Width           =   735
      End
      Begin VB.Frame frmStnDtlMsg 
         Height          =   1200
         Left            =   4800
         TabIndex        =   13
         Top             =   480
         Width           =   6000
         Begin VB.TextBox txtLeakTestMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   945
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Text            =   "frmLeakTest.frx":B2B4
            Top             =   150
            Width           =   5800
         End
         Begin VB.TextBox txtDebug1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   285
            Left            =   240
            MaxLength       =   25
            TabIndex        =   14
            Text            =   "desc"
            Top             =   840
            Width           =   5490
         End
      End
      Begin VB.TextBox txtEndOp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12480
         MaxLength       =   25
         TabIndex        =   12
         ToolTipText     =   "Alphanumeric Name"
         Top             =   2160
         Width           =   2190
      End
      Begin Threed.SSPanel pnlStatusFrame 
         Height          =   480
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   4800
         _Version        =   65536
         _ExtentX        =   8467
         _ExtentY        =   847
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   16777088
         Begin Threed.SSPanel pnlStatus 
            Height          =   300
            Left            =   90
            TabIndex        =   51
            ToolTipText     =   "Station Status"
            Top             =   90
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Station Status"
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlReportFrame 
         Height          =   480
         Left            =   10800
         TabIndex        =   52
         Top             =   0
         Width           =   4045
         _Version        =   65536
         _ExtentX        =   7135
         _ExtentY        =   847
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   16777088
         Begin Threed.SSPanel pnlReport 
            Height          =   300
            Left            =   90
            TabIndex        =   53
            ToolTipText     =   "Station Report Number"
            Top             =   90
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Station Report Number"
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin Threed.SSPanel pnlNameFrame 
         Height          =   480
         Left            =   4800
         TabIndex        =   54
         Top             =   0
         Width           =   6000
         _Version        =   65536
         _ExtentX        =   10583
         _ExtentY        =   847
         _StockProps     =   15
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
         FloodColor      =   16777088
         Begin Threed.SSPanel pnlStnName 
            Height          =   300
            Left            =   90
            TabIndex        =   55
            ToolTipText     =   "Station Number"
            Top             =   90
            Width           =   5820
            _Version        =   65536
            _ExtentX        =   10266
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Station Name"
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   1
            BevelOuter      =   0
         End
      End
      Begin VB.TextBox txtShift 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   8985
         TabIndex        =   16
         Text            =   "shift"
         Top             =   2295
         Visible         =   0   'False
         Width           =   415
      End
      Begin VB.TextBox txtDspShift 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   540
         Left            =   8985
         TabIndex        =   17
         Text            =   "8"
         Top             =   1755
         Visible         =   0   'False
         Width           =   415
      End
      Begin VB.CommandButton cmdDnShift 
         DisabledPicture =   "frmLeakTest.frx":B2BE
         DownPicture     =   "frmLeakTest.frx":B9C0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   8100
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmLeakTest.frx":C0C2
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1710
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton cmdUpShift 
         DisabledPicture =   "frmLeakTest.frx":C7C4
         DownPicture     =   "frmLeakTest.frx":CEC6
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   9465
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmLeakTest.frx":D5C8
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1710
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblStartOp 
         Caption         =   "Start Operator:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   59
         Top             =   1860
         Width           =   1455
      End
      Begin VB.Label lblEndOp 
         Caption         =   "End Operator:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   58
         Top             =   2145
         Width           =   1455
      End
      Begin VB.Label lblEngineer 
         Caption         =   "Engineer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   2145
         Width           =   1335
      End
      Begin VB.Label lblVehicle 
         Caption         =   "Vehicle No.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1860
         Width           =   1335
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   600
      Top             =   7800
   End
   Begin Threed.SSPanel pnlAlarms 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   10800
      Width           =   5190
      _Version        =   65536
      _ExtentX        =   9155
      _ExtentY        =   714
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      Begin Threed.SSPanel pnlEstop 
         Height          =   255
         Left            =   80
         TabIndex        =   1
         ToolTipText     =   "EMERGENCY Stop Pressed"
         Top             =   75
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "ESTOP"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlFlow 
         Height          =   255
         Left            =   915
         TabIndex        =   2
         ToolTipText     =   "Loss of Flow"
         Top             =   75
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "FLOW"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlBtn20 
         Height          =   255
         Left            =   2595
         TabIndex        =   3
         ToolTipText     =   "20% Butane LEL Alarm"
         Top             =   75
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "LEL 20"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlDoors 
         Height          =   255
         Left            =   1755
         TabIndex        =   4
         ToolTipText     =   "Loss of Vacuum"
         Top             =   75
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "DOORS"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlComms 
         Height          =   255
         Left            =   3435
         TabIndex        =   5
         ToolTipText     =   "Communication Error"
         Top             =   75
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "COMMS"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
      Begin Threed.SSPanel pnlMix 
         Height          =   255
         Left            =   4270
         TabIndex        =   6
         ToolTipText     =   "Communication Error"
         Top             =   75
         Width           =   840
         _Version        =   65536
         _ExtentX        =   1482
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "MIX"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   0
         BevelOuter      =   1
      End
   End
   Begin Threed.SSPanel pnlPurgeAir 
      Height          =   405
      Left            =   9360
      TabIndex        =   7
      Top             =   10800
      Width           =   5595
      _Version        =   65536
      _ExtentX        =   9869
      _ExtentY        =   714
      _StockProps     =   15
      Caption         =   "purge air"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      FloodShowPct    =   0   'False
   End
   Begin Threed.SSPanel pnlMessageFrame 
      Height          =   405
      Left            =   5190
      TabIndex        =   8
      Top             =   10800
      Width           =   4200
      _Version        =   65536
      _ExtentX        =   7408
      _ExtentY        =   714
      _StockProps     =   15
      Caption         =   "message"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      FloodShowPct    =   0   'False
      Begin Threed.SSPanel pnlMessage 
         Height          =   280
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4080
         _Version        =   65536
         _ExtentX        =   7197
         _ExtentY        =   494
         _StockProps     =   15
         Caption         =   "message"
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
      End
   End
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   1058
      ButtonWidth     =   1058
      ButtonHeight    =   953
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrLeakTest 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   60
      Top             =   3240
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   1588
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.Label Index2 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   10200
         TabIndex        =   132
         Top             =   360
         Width           =   3015
      End
   End
   Begin MSComctlLib.ImageList imgLeakTestDisabled 
      Left            =   7200
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":DCCA
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":F81C
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":1136E
            Key             =   "alarmlog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":12EC0
            Key             =   "ootlog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":14A12
            Key             =   "statsum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":16564
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":184B6
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":1A008
            Key             =   "continue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":1BB5A
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":1D6AC
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":1F1FE
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":20D50
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":228A2
            Key             =   "fuelsupply"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":243F4
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":25F46
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":27A98
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":295EA
            Key             =   "opercomment"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLeakTestHot 
      Left            =   6600
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":2B13C
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":2CC8E
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":2E7E0
            Key             =   "alarmlog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":30332
            Key             =   "ootlog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":31E84
            Key             =   "statsum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":339D6
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":35928
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":3747A
            Key             =   "continue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":38FCC
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":3AB1E
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":3C670
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":3E1C2
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":3FD14
            Key             =   "fuelsupply"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":41866
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":433B8
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":44F0A
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":46A5C
            Key             =   "opercomment"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLeakTestNormal 
      Left            =   6000
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":485AE
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":4A100
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":4BC52
            Key             =   "alarmlog"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":4D7A4
            Key             =   "ootlog"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":4F2F6
            Key             =   "statsum"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":50E48
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":52D9A
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":548EC
            Key             =   "continue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":5643E
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":57F90
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":59AE2
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":5B634
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":5D186
            Key             =   "fuelsupply"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":5ECD8
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":6082A
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":6237C
            Key             =   "pause"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeakTest.frx":63ECE
            Key             =   "opercomment"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logou&t"
      End
      Begin VB.Menu mnuCopyFile 
         Caption         =   "&Copy File"
      End
      Begin VB.Menu mnuPrintFile 
         Caption         =   "&Print File"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Program"
      End
   End
   Begin VB.Menu mnuEditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCanisters 
         Caption         =   "Ca&nisters"
      End
      Begin VB.Menu mnuRecipes 
         Caption         =   "&Recipes"
      End
      Begin VB.Menu mnuCourses 
         Caption         =   "Co&urses"
      End
      Begin VB.Menu mnuPurgeProfiles 
         Caption         =   "&Purge Profiles"
      End
      Begin VB.Menu mnuConfiguration 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuSysDef 
         Caption         =   "&System Definition"
      End
      Begin VB.Menu mnuTomCanLoad 
         Caption         =   "&TOM CanLoad"
      End
   End
   Begin VB.Menu mnuViewMenu 
      Caption         =   "&View"
      Begin VB.Menu mnuAirLog 
         Caption         =   "&AirLog"
      End
      Begin VB.Menu mnuButane 
         Caption         =   "&Butane Available"
      End
      Begin VB.Menu mnuEventLog 
         Caption         =   "&Event Log"
      End
      Begin VB.Menu mnuFileLog 
         Caption         =   "File &Maintenance Log"
      End
      Begin VB.Menu mnuFuelUseLog 
         Caption         =   "&Fuel Consumption Log"
      End
      Begin VB.Menu mnuJoblist 
         Caption         =   "&List of Jobs"
      End
      Begin VB.Menu mnuOotMonitor 
         Caption         =   "&OOT Monitor"
      End
   End
   Begin VB.Menu mnuDataMenu 
      Caption         =   "&Data"
      Begin VB.Menu mnuReviewData 
         Caption         =   "&Review Data"
      End
      Begin VB.Menu mnuWatchData 
         Caption         =   "&Watch Current Data"
      End
   End
   Begin VB.Menu mnuToolsMenu 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalibration 
         Caption         =   "&Calibation"
      End
      Begin VB.Menu mnuIoMonitor 
         Caption         =   "&I/O Monitor"
      End
      Begin VB.Menu mnuScaleMonitor 
         Caption         =   "&Scale Monitor"
      End
      Begin VB.Menu mnuSimulation 
         Caption         =   "Si&mulation Control Panel"
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu mnuOperatorManual 
         Caption         =   "&Operator Manual"
      End
      Begin VB.Menu mnuFirstAid 
         Caption         =   "&FirstAid File Save"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About CPS release7"
      End
   End
End
Attribute VB_Name = "frmLeakTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   form LeakTest, error module 1066
Option Explicit
Private sStr As String
Private LT_StnMode_Last(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Private LT_StnCourse_Last(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Private LT_DispStn_Last As Integer
Private LT_DispShift_Last As Integer
Private stnDtl_StnMode_Last(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
'
'******  ADDED  *****************
'Private stnDtl_StnCourse_Last(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
'Private stnDtl_DispStn_Last As Integer
'Private stnDtl_DispShift_Last As Integer
'Private tempText As String
'

Private Sub cmdSaveCfg_Click()
    RefreshLeaktest
    Save_LeakTest
End Sub

Private Sub cmdSaveRcp_Click()
    RefreshLeaktest
    Save_LeakTest
End Sub

Private Sub RefreshLeaktest()
    Cfg_LeakTest.timeOut = ValueFromText(txtTimeout.text)
    Cfg_LeakTest.PressTimeout = ValueFromText(txtPressTimeout.text)
    Cfg_LeakTest.PressTol = ValueFromText(txtPressTol.text)
    Cfg_LeakTest.PressTolDuration = ValueFromText(txtStablePressDur.text)
    Cfg_LeakTest.DeffTol = ValueFromText(txtDeffTol.text)
    Cfg_LeakTest.InitialN2Flow = ValueFromText(txtlIntN2Flow.text)
    Cfg_LeakTest.ReportInterval = ValueFromText(txtReportInterval.text)

    Rcp_LeakTest.TargetPress = ValueFromText(txtTargetPress.text)
    Rcp_LeakTest.HoldDuration = ValueFromText(txtDeffDur.text)
End Sub

Private Sub UpdateLeaktest()
    txtTimeout.text = Format(Cfg_LeakTest.timeOut, "###0")
    txtPressTimeout.text = Format(Cfg_LeakTest.PressTimeout, "###0")
    txtPressTol.text = Format(Cfg_LeakTest.PressTol, "###0.0##")
    txtStablePressDur.text = Format(Cfg_LeakTest.PressTolDuration, "###0")
    txtDeffTol.text = Format(Cfg_LeakTest.DeffTol, "###0.0###")
    txtlIntN2Flow.text = Format(Cfg_LeakTest.InitialN2Flow, "###0.0##")
    txtReportInterval.text = Format(Cfg_LeakTest.ReportInterval, "###0")
    
    txtTargetPress.text = Format(Rcp_LeakTest.TargetPress, "###0")
    txtDeffDur.text = Format(Rcp_LeakTest.HoldDuration, "###0")
End Sub

Private Sub Form_Activate()
Dim cntr As Integer
    cntr = 0
    Do Until ((STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Or (cntr > NR_STN))
        If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
            DispStn = IIf((DispStn < NR_STN), (DispStn + 1), 1)
        End If
        cntr = cntr + 1
    Loop
End Sub

Private Sub Form_Load()
Dim clr As Long
    KeyPreview = True
    frmLeakTest.Height = frmMainMenu.Height
    frmLeakTest.Width = frmMainMenu.Width
    SGN2 = ValueFromText(txtSgN2.text)
    BuildToolbars
'    pbxVertical.Left = 6955
    txtLeakTestMsg.text = " "
'    txtLeakTestMsg.text = " "
'   Set Foreground colors
    frmStnCanister.ForeColor = Titles_ForeColor
    txtCanID.ForeColor = TitlesData_Forecolor
    frmStnRecipe.ForeColor = Titles_ForeColor
    txtRecipeName.ForeColor = TitlesData_Forecolor
'    frmComment.ForeColor = Titles_ForeColor
    txtDspStn.ForeColor = TitlesData_Forecolor
    txtStation.ForeColor = TitlesData_Forecolor
    pnlNameFrame.ForeColor = TitlesData_Forecolor
    txtLeakTestMsg.ForeColor = Message_ForeColor
'    txtComment.ForeColor = TitlesData_Forecolor
    txtEndOp.ForeColor = TitlesData_Forecolor
    txtEngineer.ForeColor = TitlesData_Forecolor
    txtStartOp.ForeColor = TitlesData_Forecolor
    txtVehicle.ForeColor = TitlesData_Forecolor
    pnlStatus.Left = 90
    pnlStatus.Width = pnlStatusFrame.Width - 210
    pnlReportFrame.Width = frmStnDetail.Width - pnlReportFrame.Left - 135
    pnlReport.Left = 90
    pnlReport.Width = pnlReportFrame.Width - 210
    pnlStnName.Left = 90
    pnlStnName.Width = pnlNameFrame.Width - 210
    pnlStnName.FontSize = 12
    pnlReport.FontSize = 12
    pnlStatus.FontSize = 12
    
'   **********   Status Bar Setup
    pnlAlarms.Left = 0
    pnlAlarms.Width = pnlAlarms.Width - pnlMix.Width
    pnlAlarms.Top = 9675
    pnlAlarms.Height = pnlEstop.Height + 150
    pnlComms.Width = 840 - 60
   
    pnlMessageFrame.Left = pnlAlarms.Left + pnlAlarms.Width
    pnlMessageFrame.Top = pnlAlarms.Top
    pnlMessageFrame.Height = pnlAlarms.Height
    pnlMessage.Left = 60
    pnlMessage.Top = 60
    pnlMessage.Height = pnlMessageFrame.Height - 120
    pnlMessage.Width = pnlMessageFrame.Width + 360
    
    pnlPurgeAir.Left = pnlMessageFrame.Left + pnlMessageFrame.Width
    pnlPurgeAir.Width = frmLeakTest.Width - pnlPurgeAir.Left - 1800
    pnlPurgeAir.Top = pnlAlarms.Top
    pnlPurgeAir.Height = pnlAlarms.Height
    
'   **********   update leaktest cfg & rcp
    UpdateLeaktest
    
'   Status Bar Update
    UpdateStatusBars
    
'   display "current" station and shift
    txtDspStn.text = DispStn
    txtDspShift.text = DispShift
    
'   update station detail screen
'    stnDtl_StnMode_Last(DispStn, DispShift) = -1        ' force update of station mode indicator
    UpdateScreen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExitScreen
End Sub

Private Sub ExitScreen()
'   unload form
    frmLeakTest.Visible = False
    Unload Me
End Sub

Private Sub UpdateScreen()
Dim sTxt As String
Dim tempText As String
Dim iPid As Integer

'   update PID info
    iPid = 20 + DispStn
    lblLBL(0).Caption = "sp"
    lblLBL(1).Caption = "pv"
    lblLBL(2).Caption = "out"
    lblLBL(3).Caption = "cumI"
    lblLBL(4).Caption = "sp_ao"
    lblLBL(0).ForeColor = IIf(PID_INFO(iPid).Enable, Good_ForeColor, Warning_ForeColor)
    lblLBL(1).ForeColor = IIf(PID_INFO(iPid).Inhibit, Warning_ForeColor, Good_ForeColor)
    lblLBL(2).ForeColor = IIf((PID_INFO(iPid).Enable And (Not PID_INFO(iPid).Inhibit)), Good_ForeColor, Warning_ForeColor)
    lblLBL(3).ForeColor = Good_ForeColor
    lblLBL(4).ForeColor = Good_ForeColor
    lblTXT(0).Caption = Format(PID_INFO(iPid).SP, "###0.000")
    lblTXT(1).Caption = Format(PID_INFO(iPid).PV, "###0.000")
    lblTXT(2).Caption = Format(PID_INFO(iPid).out, "###0.000")
    lblTXT(3).Caption = Format(PID_INFO(iPid).CumI, "###0.000")
    lblTXT(4).Caption = Format(Stn_AIO(DispStn, asNitrogenFlowSP).EUValue, "###0.000#")

    
'   Update Text fields on LeakTest screen
    lblMsg.Caption = SEQ_Message2(DispStn, DispShift)
'   Station Name Panel
    If StationControl(DispStn, DispShift).Mode <> LT_StnMode_Last(DispStn, DispShift) _
        Or DispStn <> LT_DispStn_Last _
        Or DispShift <> LT_DispShift_Last Then
            pnlStnName.BackColor = IIf(StationControl(DispStn, DispShift).Mode = VBIDLE, pnlNameFrame.BackColor, pnlStatus.BackColor)
            pnlStnName.ForeColor = IIf(StationControl(DispStn, DispShift).Mode = VBIDLE, pnlNameFrame.ForeColor, pnlStatus.ForeColor)
            pnlStnName.Caption = STN_INFO(DispStn).desc
    End If
    txtCanID.text = JobInfo(DispStn, DispShift).Vehicle
    txtEngineer.text = JobInfo(DispStn, DispShift).Engineer
    txtVehicle.text = JobInfo(DispStn, DispShift).Vehicle
    txtStartOp.text = JobInfo(DispStn, DispShift).Start_Op
    txtEndOp.text = JobInfo(DispStn, DispShift).End_Op
    Select Case StationControl(DispStn, DispShift).Mode
        Case VBIDLE, VBIDLEWAITING
            txtLeakTestMsg.ForeColor = Message_ForeColor
            txtLeakTestMsg.text = vbCrLf & "LeakTest Idle"
        Case VBCOURSEWAIT
            txtLeakTestMsg.text = vbCrLf & StationSequence(DispStn, DispShift).CourseData(StationControl(DispStn, DispShift).Course).MsgText
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBCOURSEPAUSE
            sTxt = DurationDescription(StationSequence(DispStn, DispShift).CourseData(StationControl(DispStn, DispShift).Course).PauseDuration)
            txtLeakTestMsg.text = vbCrLf & Trim(StationControl(DispStn, DispShift).Job_Description) & " Paused for " & sTxt
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBPURGEWAIT
            If (USINGPURGEOVEN And StationRecipe(DispStn, DispShift).PurgeOven And (Not PurgeControl(DispStn, DispShift).PurgeOvenTempOK)) Then
                txtLeakTestMsg.text = vbCrLf & "waiting for Purge Oven"
            Else
                txtLeakTestMsg.text = vbCrLf & "waiting"
            End If
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBSCALEWAIT, VBSHIFTWAIT, VBSTARTWAIT, VBLEAKWAIT, VBPURGEWAIT
            txtLeakTestMsg.text = vbCrLf & "waiting"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBPAUSE, VBFIDPAUSE
            txtLeakTestMsg.text = vbCrLf & "paused"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBPAUSEBYUSER
            txtLeakTestMsg.text = vbCrLf & "Paused by User"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBGASPAUSE
            txtLeakTestMsg.text = vbCrLf & "Paused for Vapor Tank Refill"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBWBPAUSE
            txtLeakTestMsg.text = vbCrLf & "Waiting for WaterBath Temp"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBPAUSEOOT
            txtLeakTestMsg.text = vbCrLf & "Paused due to an OOT"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBPAUSEALARM
            txtLeakTestMsg.text = vbCrLf & "Paused due to an Alarm"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBLEAKERROR
            txtLeakTestMsg.text = vbCrLf & "Leakcheck Error"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case VBLEAKTEST
            sTxt = Trim(StationControl(DispStn, DispShift).Job_Description)
            txtLeakTestMsg.text = vbCrLf & "Running LeakTest"
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
        Case Else
            sTxt = Trim(StationControl(DispStn, DispShift).Job_Description)
            txtLeakTestMsg.text = vbCrLf & "Running " & sTxt
            txtLeakTestMsg.ForeColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
    End Select
' **************************************************************************************
'         MODE DISPLAY
    If StationControl(DispStn, DispShift).Mode <> LT_StnMode_Last(DispStn, DispShift) _
        Or DispStn <> LT_DispStn_Last _
        Or DispShift <> LT_DispShift_Last _
        Or StationControl(DispStn, DispShift).Mode = VBLEAK _
        Or StationControl(DispStn, DispShift).Mode = VBLOAD _
        Or StationControl(DispStn, DispShift).Mode = VBPURGE _
        Or StationControl(DispStn, DispShift).Mode = VBPOSTLEAK _
        Or StationControl(DispStn, DispShift).Mode = VBPOSTLOAD _
        Or StationControl(DispStn, DispShift).Mode = VBPOSTPURGE _
        Or StationControl(DispStn, DispShift).Mode = VBSCALEWAIT _
        Or StationControl(DispStn, DispShift).Mode = VBSTARTWAIT Then
        
'    only update if mode has changed (or description has a variable in it)
        Select Case StationControl(DispStn, DispShift).Mode
            Case VBLEAKTEST
                ' LeakTest
                tempText = ModeDescShort(VBLEAKTEST)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBLEAK
                ' Leak Check - add leak check phase description
                tempText = ModeDescShort(VBLEAK) & " - " & LeakPhaseDesc(LeakCheckControl.Phase) & " " & LeakMethodDesc(LeakCheckControl.Method)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBLOAD
                ' Loading or Waiting for Scales to Settle?
                If LoadControl(DispStn, DispShift).Phase = LoadPause Then
                    ' Waiting for Scales to Settle
                    tempText = " Load Settling for "
                    tempText = tempText & Format(StationConfig(DispStn, DispShift).LoadSettleTime, "##0.0#")
                    tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                    tempText = tempText
                    tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                Else
                    ' Loading - add load method description
                    Select Case StationRecipe(DispStn, DispShift).Load_MethodSave
                        Case NOLOAD
                            tempText = LoadTypeDesc(NOLOAD)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc2(NOLOAD)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(NOLOAD)
                        Case LOADBYTIME
                            tempText = LoadTypeDesc(LOADBYTIME)
                            tempText = tempText & Format(StationRecipe(DispStn, DispShift).Load_Time, "##0")
                            tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                        Case LOADBYWC
                            tempText = LoadTypeDesc(LOADBYWC)
                            tempText = tempText & Format(StationRecipe(DispStn, DispShift).WC_MultSave, "##0.#")
                            tempText = tempText & LoadTypeDesc2(LOADBYWC)
                            tempText = tempText & Format(StationRecipe(DispStn, DispShift).EPAFill, "##0")
                            tempText = tempText & LoadTypeDesc3(LOADBYWC)
                        Case LOADBYWEIGHT
                            tempText = LoadTypeDesc(LOADBYWEIGHT)
                            If Int(StationRecipe(DispStn, DispShift).Load_Wt) = StationRecipe(DispStn, DispShift).Load_Wt Then
                                ' no digits to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(DispStn, DispShift).Load_Wt, "##0")
                            Else
                                ' digit(s) to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(DispStn, DispShift).Load_Wt, "##0.##")
                            End If
                            tempText = tempText & LoadTypeDesc2(LOADBYWEIGHT)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYWEIGHT)
                        Case LOADBYBREAKTHRU
                            tempText = LoadTypeDesc(LOADBYBREAKTHRU)
                            If Int(StationRecipe(DispStn, DispShift).LoadBreakthrough) = StationRecipe(DispStn, DispShift).LoadBreakthrough Then
                                ' no digits to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(DispStn, DispShift).LoadBreakthrough, "##0")
                            Else
                                ' digit(s) to the right of the decimal point
                                tempText = tempText & Format(StationRecipe(DispStn, DispShift).LoadBreakthrough, "##0.##")
                            End If
                            tempText = tempText & LoadTypeDesc2(LOADBYBREAKTHRU)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYBREAKTHRU)
                        Case LOADBYFID
                            tempText = LoadTypeDesc(LOADBYFID)
                            tempText = tempText & Format(StationRecipe(DispStn, DispShift).FIDmg, "#####0")
                            tempText = tempText & LoadTypeDesc2(LOADBYFID)
                            tempText = tempText
                            tempText = tempText & LoadTypeDesc3(LOADBYFID)
                    End Select
                End If
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBPURGE
                ' Purging or Waiting for Scales to Settle?
                If PurgeControl(DispStn, DispShift).Phase = PurgePause Then
                    ' Waiting for Scales to Settle
                    tempText = " Purge Settling for "
                    tempText = tempText & Format(StationConfig(DispStn, DispShift).PurgeSettleTime, "##0.0#")
                    tempText = tempText & LoadTypeDesc2(PURGEBYTIME)
                    tempText = tempText
                    tempText = tempText & LoadTypeDesc3(PURGEBYTIME)
                Else
                    ' Purge - add Purge method description
                    Select Case StationRecipe(DispStn, DispShift).Purge_Method
                        Case NOPURGE
                            tempText = "No Purge"
                        Case PURGEBYTIME
                            tempText = ModeDescShort(VBPURGE) & " for " & StationRecipe(DispStn, DispShift).Purge_Time & " Minute"
                            If StationRecipe(DispStn, DispShift).Purge_Time > 1 Then tempText = tempText & "s"
                        Case PURGEBYLITERS
                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(DispStn, DispShift).Purge_Liters & " liter"
                            If StationRecipe(DispStn, DispShift).Purge_Liters <> 1 Then tempText = tempText & "s"
                        Case PURGEBYVOLUME
                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(DispStn, DispShift).Purge_Can_Vol & " Canister Volume"
                            If StationRecipe(DispStn, DispShift).Purge_Can_Vol <> 1 Then tempText = tempText & "s"
                        Case PURGEAUXONLY
                            tempText = ModeDescShort(VBPURGE) & " Aux Can for " & StationRecipe(DispStn, DispShift).Purge_AuxTime & " Minute"
                            If StationRecipe(DispStn, DispShift).Purge_AuxTime > 1 Then tempText = tempText & "s"
                        Case PURGEBYPROFILE
                            tempText = ModeDescShort(VBPURGE) & " " & " by Profile"
                        Case PURGEBYWC
                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(DispStn, DispShift).Purge_TargetWC & " % of Work Cap"
                        Case PURGETOTARGET
                            tempText = ModeDescShort(VBPURGE) & " to " & StationRecipe(DispStn, DispShift).Purge_TargetWeight & " grams"
                        Case PURGETOUNDOLOAD
                            tempText = ModeDescShort(VBPURGE) & " to " & " Undo Load"
                    End Select
                End If
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBPOSTLEAK
                ' Post LeakCheck Pause
                tempText = ModeDescShort(VBPOSTLEAK)
                tempText = tempText & " for "
                tempText = tempText & Format(StationRecipe(DispStn, DispShift).PauseLeakTime, "##0.0#")
                tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                tempText = tempText
                tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBPOSTLOAD
                ' Post Load Pause
                tempText = ModeDescShort(VBPOSTLOAD)
                tempText = tempText & " for "
                tempText = tempText & Format(StationRecipe(DispStn, DispShift).PauseLoadTime, "##0.0#")
                tempText = tempText & LoadTypeDesc2(LOADBYTIME)
                tempText = tempText
                tempText = tempText & LoadTypeDesc3(LOADBYTIME)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBPOSTPURGE
                ' Post Purge Pause
                tempText = ModeDescShort(VBPOSTPURGE)
                tempText = tempText & " for "
                tempText = tempText & Format(StationRecipe(DispStn, DispShift).PausePurgeTime, "##0.0#")
                tempText = tempText & LoadTypeDesc2(PURGEBYTIME)
                tempText = tempText
                tempText = tempText & LoadTypeDesc3(PURGEBYTIME)
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBSCALEWAIT
                ' Waiting for Scale(s) - add which scale(s)
                tempText = ModeDescShort(VBSCALEWAIT)
                If StationRecipe(DispStn, DispShift).UsePriScale And StationRecipe(DispStn, DispShift).UseAuxScale Then
                    ' Using Two Scales
                    tempText = tempText & "s "
                    ' Scales in use ?
                    If Scale_In_Use(StationRecipe(DispStn, DispShift).PriScaleNo) And Scale_In_Use(StationRecipe(DispStn, DispShift).AuxScaleNo) Then
                        ' Both Scales in use
                        tempText = tempText & Format(StationRecipe(DispStn, DispShift).PriScaleNo, "#0") & " && " & Format(StationRecipe(DispStn, DispShift).AuxScaleNo, "#0")
                    ElseIf Scale_In_Use(StationRecipe(DispStn, DispShift).PriScaleNo) Then
                        ' Primary Scale in use
                        tempText = tempText & Format(StationRecipe(DispStn, DispShift).PriScaleNo, "#0")
                    ElseIf Scale_In_Use(StationRecipe(DispStn, DispShift).AuxScaleNo) Then
                        ' Aux Scale in use
                        tempText = tempText & Format(StationRecipe(DispStn, DispShift).AuxScaleNo, "#0")
                    End If
                ElseIf StationRecipe(DispStn, DispShift).UsePriScale Then
                    ' Using Only Primary Scale
                    tempText = tempText & " "
                    If Scale_In_Use(StationRecipe(DispStn, DispShift).PriScaleNo) Then tempText = tempText & Format(StationRecipe(DispStn, DispShift).PriScaleNo, "#0")
                ElseIf StationRecipe(DispStn, DispShift).UseAuxScale Then
                    ' Using Only Aux Scale
                    tempText = tempText & " "
                    If Scale_In_Use(StationRecipe(DispStn, DispShift).AuxScaleNo) Then tempText = tempText & Format(StationRecipe(DispStn, DispShift).AuxScaleNo, "#0")
                End If
                pnlStatus.Caption = tempText
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case VBSTARTWAIT
                ' Delayed Start - add how long
                tempText = StartTypeDesc(StationRecipe(DispStn, DispShift).StartMethod)
                ' Which Method of Delay ?
                Select Case StationRecipe(DispStn, DispShift).StartMethod
                    Case STARTNOW
                        tempText = tempText & StartTypeDesc2(STARTNOW)
                    Case STARTDELAYED
                        tempText = tempText & Format(StationRecipe(DispStn, DispShift).StartDelay, "##0")
                        tempText = tempText & StartTypeDesc2(STARTDELAYED)
                    Case STARTATDATE
                        tempText = tempText & StartTypeDesc2(STARTATDATE)
                        tempText = tempText & Format(StationRecipe(DispStn, DispShift).StartDate, "D MMM, YYYY   h:mm")
                End Select
                pnlStatus.Caption = tempText
                pnlStatus.ToolTipText = ""
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
            Case Else
                pnlStatus.Caption = ModeDescShort(StationControl(DispStn, DispShift).Mode)
                pnlStatus.BackColor = ModeBackColor(StationControl(DispStn, DispShift).Mode)
                pnlStatus.ForeColor = ModeForeColor(StationControl(DispStn, DispShift).Mode)
        End Select
    End If
    ' JobNumber Panel
    If StationControl(DispStn, DispShift).DBFile = "" Then
        pnlReport.Caption = "No Active Job File"
    Else
        pnlReport.Caption = "Job Number  " & StationControl(DispStn, DispShift).Job_Number
    End If
    If pnlReport.BackColor <> pnlStatus.BackColor Then pnlReport.BackColor = pnlStatus.BackColor
    If pnlReport.ForeColor <> pnlStatus.ForeColor Then pnlReport.ForeColor = pnlStatus.ForeColor
    ' station number
    txtDspStn.text = Format(DispStn, "0")
    ' Toolbars
    UpdateNavigateBtns
    ' Status Bar Update
    UpdateStatusBars
    If (STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Then
        UpdateSeq
        lblN2Flow.Caption = Format(QN2, "0.00000###")
        lblN2Temp.Caption = Format(TN2, "###0.0#")
        lblInletPress.Caption = Format(Pin, "####0.0#")
        lblPatm.Caption = Format(Patm, "####0.0#")
        txtSgN2.text = Format(SGN2, "#0.0####")
        lblEffLeakDia.Caption = Format(Deff, "#0.0####")
        lblEffLeakDia.ForeColor = IIf(DeffCalcFlag, DKORANGE, MEDYELLOW)
        lblEffLeakDia.ToolTipText = DeffCalcMsg
'        lblCalcMsg.Caption = IIf(DeffCalcFlag, "", DeffCalcMsg)
        lblCalcMsg.Caption = DeffCalcMsg
        lblCalcMsg.ForeColor = IIf(DeffCalcFlag, Message_ForeColor, Alarm_ForeColor)
    End If
    LT_DispStn_Last = DispStn
    LT_DispShift_Last = DispShift
    LT_StnCourse_Last(DispStn, DispShift) = StationControl(DispStn, DispShift).Course
    LT_StnMode_Last(DispStn, DispShift) = StationControl(DispStn, DispShift).Mode
End Sub

Private Sub UpdateSeq()
    If (SEQ_Step(DispStn, DispShift) = 0) Then
        ' sequence is idle
        frmLeakTestSeq.Font = DKGRAY
        lblTask.Font = MEDGRAY
        lblTaskVal.Font = MEDGRAY
        lblTaskVal.Caption = SEQ_Task(DispStn, DispShift)
        lblDesc.Font = MEDGRAY
        lblDescVal.Font = MEDGRAY
        lblDescVal.Caption = SEQ_Message(DispStn, DispShift)
        lblStep.Font = MEDGRAY
        lblStepVal.Font = MEDGRAY
        lblStepVal.Caption = Format(SEQ_Step(DispStn, DispShift), "#0")
    ElseIf ((SEQ_Step(DispStn, DispShift) > 0) And (SEQ_Step(DispStn, DispShift) < 10)) Then
        ' sequence is running
        frmLeakTestSeq.Font = DK3ORANGE
        lblTask.Font = MEDBLUE
        lblTaskVal.Font = MEDBLUE
        lblTaskVal.Caption = SEQ_Task(DispStn, DispShift)
        lblDesc.Font = MEDBLUE
        lblDescVal.Font = MEDBLUE
        lblDescVal.Caption = SEQ_Message(DispStn, DispShift)
        lblStep.Font = MEDBLUE
        lblStepVal.Font = MEDBLUE
        lblStepVal.Caption = Format(SEQ_Step(DispStn, DispShift), "#0")
    ElseIf ((SEQ_Step(DispStn, DispShift) = 10)) Then
        ' sequence is done
        frmLeakTestSeq.Font = DK3ORANGE
        lblTask.Font = DK3ORANGE
        lblTaskVal.Font = DK3ORANGE
        lblTaskVal.Caption = SEQ_Task(DispStn, DispShift)
        lblDesc.Font = DK3ORANGE
        lblDescVal.Font = DK3ORANGE
        lblDescVal.Caption = SEQ_Message(DispStn, DispShift)
        lblStep.Font = DK3ORANGE
        lblStepVal.Font = DK3ORANGE
        lblStepVal.Caption = Format(SEQ_Step(DispStn, DispShift), "#0")
    ElseIf ((SEQ_Step(DispStn, DispShift) > 10) And (SEQ_Step(DispStn, DispShift) <= 99)) Then
        ' sequence is aborted
        frmLeakTestSeq.Font = MEDRED
        lblTask.Font = MEDRED
        lblTaskVal.Font = MEDRED
        lblTaskVal.Caption = SEQ_Task(DispStn, DispShift)
        lblDesc.Font = MEDRED
        lblDescVal.Font = MEDRED
        lblDescVal.Caption = SEQ_Message(DispStn, DispShift)
        lblStep.Font = MEDRED
        lblStepVal.Font = MEDRED
        lblStepVal.Caption = Format(SEQ_Step(DispStn, DispShift), "#0")
    End If
End Sub

Private Sub tmrUpdate_Timer()
    UpdateScreen
End Sub

Private Sub BuildToolbars()
' Create object variable for the Toolbar.
Dim btnX As MSComctlLib.Button
' **********************************************************
'               NAVIGATION TOOLBAR
' **********************************************************
    ' Load the ImageLists
    tbrNavigate.ImageList = frmMainMenu.imgNavigateNormal
    tbrNavigate.DisabledImageList = frmMainMenu.imgNavigateDisabled
    tbrNavigate.HotImageList = frmMainMenu.imgNavigateHot
    
    ' Add button objects to Buttons collection using the
    ' Add method. After creating each button, set both
    ' Description and ToolTipText properties.
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)

'   **********   Canisters Screen
    Set btnX = tbrNavigate.Buttons.Add(, "canisters", , tbrDefault, "can_master")
    btnX.ToolTipText = "Master Canisters"
    btnX.Description = btnX.ToolTipText
    
'   **********   Recipes Screen
    Set btnX = tbrNavigate.Buttons.Add(, "recipes", , tbrDefault, "rcp_master")
    btnX.ToolTipText = "Master Recipes"
    btnX.Description = btnX.ToolTipText
    
'   **********   Purge Profiles Screen
    Set btnX = tbrNavigate.Buttons.Add(, "purgeprofile", , tbrDefault, "prof_master")
    btnX.ToolTipText = "Master Purge Profiles"
    btnX.Description = btnX.ToolTipText
    
'   **********   Sequence (Courses) Screen
    Set btnX = tbrNavigate.Buttons.Add(, "courses", , tbrDefault, "seq_master")
    btnX.ToolTipText = "Master Sequences"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
'   **********   TOM Can Load Tasks Screen
    Set btnX = tbrNavigate.Buttons.Add(, "tomcanload", , tbrDefault, "remotecontrol")
    btnX.ToolTipText = "Task Order Manager Tasks"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
'   **********   Configuration Screen
    Set btnX = tbrNavigate.Buttons.Add(, "configuration", , tbrDefault, "configuration")
    btnX.ToolTipText = "Configuration"
    btnX.Description = btnX.ToolTipText
    
'   **********   System Definition Screen
    Set btnX = tbrNavigate.Buttons.Add(, "sysdef", , tbrDefault, "sysdef")
    btnX.ToolTipText = "System Definition"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
'   **********   Fuel Use Log
    Set btnX = tbrNavigate.Buttons.Add(, "fueluselog", , tbrDefault, "fueluselog")
    btnX.ToolTipText = "Fuel Consumption Log"
    btnX.Description = btnX.ToolTipText
    
'   **********   Butane Available
    Set btnX = tbrNavigate.Buttons.Add(, "butane", , tbrDefault, "flammablegas")
    btnX.ToolTipText = "Butane Available"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
        
'   **********   Event Log Screen
    Set btnX = tbrNavigate.Buttons.Add(, "eventlog", , tbrDefault, "eventlog")
    btnX.ToolTipText = "Event Log"
    btnX.Description = btnX.ToolTipText
    
'   **********   Joblist Screen
    Set btnX = tbrNavigate.Buttons.Add(, "joblist", , tbrDefault, "joblist")
    btnX.ToolTipText = "List of Jobs"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
     
'   **********   Station Detail Screen
    Set btnX = tbrNavigate.Buttons.Add(, "stndetail", , tbrDefault, "stndetail")
    btnX.ToolTipText = "Station Details"
    btnX.Description = btnX.ToolTipText

'   **********   Overview Screen
    Set btnX = tbrNavigate.Buttons.Add(, "overview", , tbrDefault, "overview")
    btnX.ToolTipText = "Overview"
    btnX.Description = btnX.ToolTipText
    
'   **********   Review Screen
    Set btnX = tbrNavigate.Buttons.Add(, "reviewdata", , tbrDefault, "reviewdata")
    btnX.ToolTipText = "Review Data"
    btnX.Description = btnX.ToolTipText
    
'   **********   Watch Screen
    Set btnX = tbrNavigate.Buttons.Add(, "watchdata", , tbrDefault, "watchdata")
    btnX.ToolTipText = "Watch Data"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
       
'   **********   Calibration Screen
    Set btnX = tbrNavigate.Buttons.Add(, "calibration", , tbrDefault, "calibration")
    btnX.ToolTipText = "Calibration"
    btnX.Description = btnX.ToolTipText
    
'   **********   I/O Monitor Screen
    Set btnX = tbrNavigate.Buttons.Add(, "iomonitor", , tbrDefault, "iomonitor")
    btnX.ToolTipText = "I/O Monitor"
    btnX.Description = btnX.ToolTipText
    
'   **********   Scale Monitor Screen
    Set btnX = tbrNavigate.Buttons.Add(, "scalemonitor", , tbrDefault, "scalemonitor")
    btnX.ToolTipText = "Scale Monitor"
    btnX.Description = btnX.ToolTipText
       
    Select Case LocalPagControl.Type
    
        Case pagNone
            ' no PurgeAir Generator control
        
        Case pagAlone
            ' Stand-Alone PurgeAir Generator control
        
        Case pagMaster
            ' AK Master PurgeAir Generator control
            Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            ' Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            ' AK Server
            Set btnX = tbrNavigate.Buttons.Add(, "ak_server", , tbrDefault, "ak_server")
            btnX.ToolTipText = "AK Server"
            btnX.Description = btnX.ToolTipText
    
        Case pagClient
            ' AK Client PurgeAir Generator control
            Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            ' Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            ' AK Client
            Set btnX = tbrNavigate.Buttons.Add(, "ak_client", , tbrDefault, "ak_client")
            btnX.ToolTipText = "AK Client"
            btnX.Description = btnX.ToolTipText
            ' AK Server
            Set btnX = tbrNavigate.Buttons.Add(, "ak_server", , tbrDefault, "ak_server")
            btnX.ToolTipText = "AK Server"
            btnX.Description = btnX.ToolTipText
        
    End Select
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
       
'   **********   Simulation Control Panel
    Set btnX = tbrNavigate.Buttons.Add(, "simulation", , tbrDefault, "simulation")
    btnX.ToolTipText = "Simulation Control Panel"
    btnX.Description = btnX.ToolTipText
    
    If ((Com_DIO(icAlarmBeacon).addr <> 0) Or (Com_DIO(icAlarmBeacon).chan <> 0)) Then
'       ***********   TurnOff Beacon
        Set btnX = tbrNavigate.Buttons.Add(, "beaconoff", , tbrDefault, "beaconoff")
        btnX.ToolTipText = "Turn Off Beacon"
        btnX.Description = btnX.ToolTipText
    End If
    
    If ((Com_DIO(icAlarmHorn).addr <> 0) Or (Com_DIO(icAlarmHorn).chan <> 0)) Then
'       **********   TurnOff Horn
        Set btnX = tbrNavigate.Buttons.Add(, "hornoff", , tbrDefault, "hornoff")
        btnX.ToolTipText = "Silence Horn"
        btnX.Description = btnX.ToolTipText
    End If
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            
'   **********   Operators Manual
    Set btnX = tbrNavigate.Buttons.Add(, "opermanual", , tbrDefault, "opermanual")
    btnX.ToolTipText = "Operators Manual"
    btnX.Description = btnX.ToolTipText
    
    ' ************************
    ' LEAKTEST STATION TOOLBAR
    ' ************************
    ' Load the ImageLists
    tbrLeakTest.ImageList = imgLeakTestNormal
    tbrLeakTest.DisabledImageList = imgLeakTestDisabled
    tbrLeakTest.HotImageList = imgLeakTestHot
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
                
'  ************ Alarm Log
    Set btnX = tbrLeakTest.Buttons.Add(, "alarmlog", , tbrDefault, "alarmlog")
    btnX.ToolTipText = "Alarm Log"
    btnX.Description = btnX.ToolTipText
    
'  ************* OOT Log
    Set btnX = tbrLeakTest.Buttons.Add(, "ootlog", , tbrDefault, "ootlog")
    btnX.ToolTipText = "Out Of Tolerance Log"
    btnX.Description = btnX.ToolTipText
    
'  **************  Statistics Summary
    Set btnX = tbrLeakTest.Buttons.Add(, "statsum", , tbrDefault, "statsum")
    btnX.ToolTipText = "Statistics Summary"
    btnX.Description = btnX.ToolTipText
    
'   *************  Job Log
    Set btnX = tbrLeakTest.Buttons.Add(, "joblog", , tbrDefault, "joblog")
    btnX.ToolTipText = "Job Log"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
                
'   ********   Operator Comment
    Set btnX = tbrLeakTest.Buttons.Add(, "opercomment", , tbrDefault, "opercomment")
    btnX.ToolTipText = "Operator Comment"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
'  **********   Start
    Set btnX = tbrLeakTest.Buttons.Add(, "start", , tbrDefault, "start")
    btnX.ToolTipText = "Start Job"
    btnX.Description = btnX.ToolTipText
    
'   ********   Continue
    Set btnX = tbrLeakTest.Buttons.Add(, "continue", , tbrDefault, "continue")
    btnX.ToolTipText = "Continue Job"
    btnX.Description = btnX.ToolTipText
    btnX.Visible = False
    
'  ************     Pause
    Set btnX = tbrLeakTest.Buttons.Add(, "pause", , tbrDefault, "pause")
    btnX.ToolTipText = "Pause Job"
    btnX.Description = btnX.ToolTipText
    btnX.Visible = False
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    
'   **********      Stop
    Set btnX = tbrLeakTest.Buttons.Add(, "stop", , tbrDefault, "stop")
    btnX.ToolTipText = "Stop Job"
    btnX.Description = btnX.ToolTipText
                
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
            
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrLeakTest.Buttons.Add(, , , tbrSeparator)
                
    Show
    
    tbrLeakTest.Buttons("start").Enabled = IIf(CheckPass("R", False), True, False)
    tbrLeakTest.Buttons("stop").Enabled = IIf(CheckPass("R", False), True, False)
    tbrLeakTest.Buttons("pause").Enabled = IIf(CheckPass("R", False), True, False)

    txtShift.ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
    txtDspShift.ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
    cmdUpShift.Enabled = IIf(NR_SHIFT > 1, True, False)
    cmdDnShift.Enabled = IIf(NR_SHIFT > 1, True, False)
End Sub

Sub UpdateNavigateBtns()
'
' Routine Name:  UpdateNavigateBtns
' Description:
' Updates the Navigate toolbar buttons
'
Dim iKeyCount As Integer
Dim tempText As String
 
SetErrModule 1066, 10101
If UseLocalErrorHandler Then On Error GoTo localhandler

'Exit Sub
        
        ' Login
        If CheckPass("J", False) Then
'            tbrNavigate.Buttons("login").Enabled = True
            mnuLogin.Enabled = True
        Else
'            tbrNavigate.Buttons("login").Enabled = False
            mnuLogin.Enabled = False
        End If
        
        ' Logout
'        tbrNavigate.Buttons("logout").Enabled = True
        mnuLogout.Enabled = True
        
        ' CopyFiles
        If CheckPass("F", False) Then
'            tbrNavigate.Buttons("copyfiles").Visible = True
'            tbrNavigate.Buttons("copyfiles").Enabled = True
            mnuCopyFile.Enabled = True
       Else
'            tbrNavigate.Buttons("copyfiles").Visible = False
'            tbrNavigate.Buttons("copyfiles").Enabled = False
            mnuCopyFile.Enabled = False
        End If

        ' PrintFiles
        If CheckPass("F", False) Then
'            tbrNavigate.Buttons("printfiles").Visible = True
'            tbrNavigate.Buttons("printfiles").Enabled = True
            mnuPrintFile.Enabled = True
       Else
'            tbrNavigate.Buttons("printfiles").Visible = False
'            tbrNavigate.Buttons("printfiles").Enabled = False
            mnuPrintFile.Enabled = False
        End If
        
        ' Canisters
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("canisters").Enabled = True
            mnuCanisters.Enabled = True
       Else
            tbrNavigate.Buttons("canisters").Enabled = False
            mnuCanisters.Enabled = False
        End If
        
        ' Recipes
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("recipes").Enabled = True
            mnuRecipes.Enabled = True
        Else
            tbrNavigate.Buttons("recipes").Enabled = False
            mnuRecipes.Enabled = False
        End If
        
        ' Purge Profiles
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("purgeprofile").Enabled = True
            mnuPurgeProfiles.Enabled = True
        Else
            tbrNavigate.Buttons("purgeprofile").Enabled = False
            mnuPurgeProfiles.Enabled = False
        End If
        
        ' Courses
        If CheckPass("N", False) And (NR_JOBSEQ > 1) Then
            tbrNavigate.Buttons("courses").Visible = True
            mnuCourses.Visible = True
        Else
            tbrNavigate.Buttons("courses").Visible = False
            mnuCourses.Visible = False
        End If
        
        ' TomCanLoad
        If CheckPass("N", False) And USINGREMCANLOAD Or USINGTOMCANLOAD Then
            tbrNavigate.Buttons("tomcanload").Visible = True
            mnuTomCanLoad.Visible = True
        Else
            tbrNavigate.Buttons("tomcanload").Visible = False
            mnuTomCanLoad.Visible = False
        End If
        
        ' Configuration
        If CheckPass("B", False) Then
            tbrNavigate.Buttons("configuration").Enabled = True
            mnuConfiguration.Enabled = True
        Else
            mnuConfiguration.Enabled = False
            tbrNavigate.Buttons("configuration").Enabled = False
        End If
        
        ' System Definition
        If CheckPass("H", False) Then
            tbrNavigate.Buttons("sysdef").Visible = True
            tbrNavigate.Buttons("sysdef").ToolTipText = "System Definition"
            mnuSysDef.Visible = True
        Else
            tbrNavigate.Buttons("sysdef").Visible = False
            tbrNavigate.Buttons("sysdef").ToolTipText = ""
            mnuSysDef.Visible = False
        End If
        
        ' Butane Available
        If systemhasBUTANE Then
            tbrNavigate.Buttons("butane").Visible = True
            mnuButane.Enabled = True
        Else
            tbrNavigate.Buttons("butane").Visible = False
            mnuButane.Enabled = False
        End If
        
        ' Event Log
        If CheckPass("Z", False) Then
            tbrNavigate.Buttons("eventlog").Enabled = True
            mnuEventLog.Enabled = True
        Else
            mnuEventLog.Enabled = False
            tbrNavigate.Buttons("eventlog").Enabled = False
        End If
        
        ' Joblist Log
        If CheckPass("M", False) Then
            tbrNavigate.Buttons("joblist").Enabled = True
            mnuJoblist.Enabled = True
        Else
            mnuJoblist.Enabled = False
            tbrNavigate.Buttons("joblist").Enabled = False
        End If
        
        ' Review Previous Cycle Data
        If CheckPass("F", False) Then
            tbrNavigate.Buttons("reviewdata").Enabled = True
            mnuReviewData.Enabled = True
        Else
            mnuReviewData.Enabled = False
            tbrNavigate.Buttons("reviewdata").Enabled = False
        End If
        
        ' Watch Current Cycle Data
        If CheckPass("F", False) Then
            tbrNavigate.Buttons("watchdata").Enabled = True
            mnuWatchData.Enabled = True
        Else
            mnuWatchData.Enabled = False
            tbrNavigate.Buttons("watchdata").Enabled = False
        End If
        
        ' MFC Calibration
        If CheckPass("X", False) Then
            tbrNavigate.Buttons("calibration").Enabled = True
            mnuCalibration.Enabled = True
        Else
            mnuCalibration.Enabled = False
            tbrNavigate.Buttons("calibration").Enabled = False
        End If
        
        ' I/O Monitor
        If CheckPass("2", False) Then
            tbrNavigate.Buttons("iomonitor").Enabled = True
            mnuIoMonitor.Enabled = True
        Else
            mnuIoMonitor.Enabled = False
            tbrNavigate.Buttons("iomonitor").Enabled = False
        End If
        
        ' Scale Monitor
        If CheckPass("3", False) Then
            tbrNavigate.Buttons("scalemonitor").Enabled = True
            mnuScaleMonitor.Enabled = True
        Else
            mnuScaleMonitor.Enabled = False
            tbrNavigate.Buttons("scalemonitor").Enabled = False
        End If
       
        ' Simulation
        If Not IoComOn And USINGSIMULATION And CheckPass("H", False) Then
            tbrNavigate.Buttons("simulation").Visible = True
            tbrNavigate.Buttons("simulation").Enabled = True
            tbrNavigate.Buttons("simulation").ToolTipText = "Simulation Control Panel"
'            mnuSimulation.Enabled = True
        Else
'            mnuSimulation.Enabled = False
            tbrNavigate.Buttons("simulation").Visible = False
            tbrNavigate.Buttons("simulation").Enabled = False
            tbrNavigate.Buttons("simulation").ToolTipText = ""
        End If
        
        ' Operator Manual
        If CheckPass("H", False) Then
            tbrNavigate.Buttons("opermanual").Visible = False
            tbrNavigate.Buttons("opermanual").Enabled = False
            mnuOperatorManual.Enabled = True
        ElseIf CheckPass("D", False) Then
            tbrNavigate.Buttons("opermanual").Visible = True
            tbrNavigate.Buttons("opermanual").Enabled = True
            mnuOperatorManual.Enabled = True
        Else
            mnuOperatorManual.Enabled = False
            tbrNavigate.Buttons("opermanual").Visible = False
            tbrNavigate.Buttons("opermanual").Enabled = False
        End If
        
'        ' FirstAid
        If CheckPass("T", False) Then
'            tbrNavigate.Buttons("firstaid").Visible = True
'            tbrNavigate.Buttons("firstaid").Enabled = True
'            tbrNavigate.Buttons("firstaid").ToolTipText = "FirstAid File Save for APS"
            mnuFirstAid.Enabled = True
        Else
            mnuFirstAid.Enabled = False
'            tbrNavigate.Buttons("firstaid").Visible = False
'            tbrNavigate.Buttons("firstaid").Enabled = False
'            tbrNavigate.Buttons("firstaid").ToolTipText = ""
        End If
        
        ' Close Screen
'        tbrNavigate.Buttons("close").Enabled = True
        
        ' View AirLog
        If LogTempRh Then
            mnuAirLog.Enabled = True
        Else
            mnuAirLog.Enabled = False
        End If
        
        ' Exit Program
        If CheckPass("G", False) Then
'            tbrNavigate.Buttons("exit").Enabled = True
            mnuExit.Enabled = True
        Else
            mnuExit.Enabled = False
'            tbrNavigate.Buttons("exit").Enabled = False
        End If

        ' **********************
        ' **********************
        '    LEAKTEST TOOLBAR
        ' **********************
        ' **********************
        
        tbrLeakTest.Buttons("alarmlog").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrLeakTest.Buttons("ootlog").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrLeakTest.Buttons("statsum").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrLeakTest.Buttons("joblog").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrLeakTest.Buttons("opercomment").Enabled = IIf(Len(StationControl(DispStn, DispShift).DBFile) > 0, True, False)
        tbrLeakTest.Buttons("opercomment").Visible = True
        
        '*************************************
        '*************************************
        '*************************************
        ' START/CONTINUE, PAUSE & STOP BUTTONS
        ' START/CONTINUE, PAUSE & STOP BUTTONS
        ' START/CONTINUE, PAUSE & STOP BUTTONS
        '*************************************
        '*************************************
        '*************************************
        ChgErrModule 1066, 10102
        If Stop_In_Progress = True Then
            tbrLeakTest.Buttons("start").Enabled = False
            If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
            If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
            tbrLeakTest.Buttons("stop").Enabled = False
        Else
            Select Case StationControl(DispStn, DispShift).Mode
                Case VBIDLE
                    tbrLeakTest.Buttons("continue").Visible = False
                    tbrLeakTest.Buttons("pause").Visible = False
                    tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = False
                    tbrLeakTest.Buttons("start").Visible = True
                    tbrLeakTest.Buttons("start").ToolTipText = "Start Job"
                    If AdfControl(DispStn).Step = 0 Then
                        If Pause_Alarm = SYSTEMPAUSED Then
                            ' system is paused
                            tbrLeakTest.Buttons("start").Enabled = False
                        Else
                            ' station is not paused
                            If Stn_OperReportNameIsValid = False Then
                                If (SysConfig.ReportFileName1stPart = RPT_OPERENTER _
                                  Or SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
                                  Or SysConfig.ReportFileName3rdPart = RPT_OPERENTER) Then
                                ' operator entry of report name required before Job can be started
                                tbrLeakTest.Buttons("start").Enabled = False
                                Else
                                ' no operator entry required
                                tbrLeakTest.Buttons("start").Enabled = True
                                End If
                            Else
                                ' valid operator entry
                                tbrLeakTest.Buttons("start").Enabled = True
                            End If
                        End If
                    Else
                        tbrLeakTest.Buttons("start").Enabled = False
                    End If
            
                Case VBLEAKERROR
                    tbrLeakTest.Buttons("continue").Visible = False
                    tbrLeakTest.Buttons("pause").Visible = False
                    tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    tbrLeakTest.Buttons("start").Visible = True
                    Select Case SysConfig.LeakCheckFailResponse
                        Case MANUALCHOOSE, AUTOCONTINUE
                            tbrLeakTest.Buttons("start").Enabled = True
                            tbrLeakTest.Buttons("start").ToolTipText = "CONTINUE"
                        Case Else
                            tbrLeakTest.Buttons("start").Enabled = False
                            tbrLeakTest.Buttons("start").ToolTipText = "CONTINUE"
                    End Select
            
                Case VBPAUSEALARM
                    tbrLeakTest.Buttons("continue").Visible = False
                    tbrLeakTest.Buttons("pause").Visible = False
                    If Pause_Alarm = SYSTEMPAUSED Then
                        ' System is Paused
                        tbrLeakTest.Buttons("start").Visible = True
                        tbrLeakTest.Buttons("start").Enabled = False
                        tbrLeakTest.Buttons("stop").Visible = True
                        tbrLeakTest.Buttons("stop").Enabled = False
                    Else
                        ' Station is Paused
                        tbrLeakTest.Buttons("start").Visible = True
                        tbrLeakTest.Buttons("start").Enabled = True
                        tbrLeakTest.Buttons("stop").Visible = True
                        tbrLeakTest.Buttons("stop").Enabled = True
                        tempText = "Continue"
                        If StationControl(DispStn, DispShift).Mode_PauseSave = VBLEAK Then tempText = "Restart Leak Check"
                        If StationControl(DispStn, DispShift).Mode_PauseSave = VBLOAD Then tempText = "Resume Load"
                        If StationControl(DispStn, DispShift).Mode_PauseSave = VBPURGE Then tempText = "Resume Purge"
                        tbrLeakTest.Buttons("start").ToolTipText = tempText
                    End If
            
                Case VBPAUSEOOT
                    tbrLeakTest.Buttons("continue").Visible = False
                    tbrLeakTest.Buttons("pause").Visible = False
                    tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    tbrLeakTest.Buttons("start").Visible = True
                    Select Case StationControl(DispStn, DispShift).Mode_PauseSave
                        Case VBLEAK
                            tbrLeakTest.Buttons("start").Enabled = False
                            tbrLeakTest.Buttons("start").ToolTipText = "Continue"
                        Case VBLOAD
                            tbrLeakTest.Buttons("start").Enabled = IIf((StationControl(DispStn, DispShift).OotResponse = ootrspStop), False, True)
                            tbrLeakTest.Buttons("start").ToolTipText = "Resume Load"
                        Case VBPURGE, VBPURGECONT
                            tbrLeakTest.Buttons("start").Enabled = IIf((StationControl(DispStn, DispShift).OotResponse = ootrspStop), False, True)
                            tbrLeakTest.Buttons("start").ToolTipText = "Resume Purge"
                        Case Else
                            tbrLeakTest.Buttons("start").Enabled = IIf((StationControl(DispStn, DispShift).OotResponse = ootrspStop), False, True)
                            tbrLeakTest.Buttons("start").ToolTipText = "Continue"
                    End Select
                    
                Case VBFIDPAUSE                                  ' Pause for FID
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If Not tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = True
                    tbrLeakTest.Buttons("start").Enabled = True
                    tbrLeakTest.Buttons("start").ToolTipText = "Continue; FID is ready"

                Case VBGASPAUSE                                  ' Pause for Live Fuel
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If Not tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = True
'                    If (STN_INFO(DispStn).ADF_DEF.hasADF_WaterBath And StationRecipe(DispStn, DispShift).ADF_Heater) Then
'                        If (LoadControl(DispStn, DispShift).WaterBathTempOK) Then
'                            tbrLeakTest.Buttons("start").Enabled = True
'                            tbrLeakTest.Buttons("start").ToolTipText = "Continue; Vapor Tank is ready"
'                        Else
'                            tbrLeakTest.Buttons("start").Enabled = False
'                            tbrLeakTest.Buttons("start").ToolTipText = "Vapor Tank is Not Ready; check WaterBath"
'                        End If
'                    Else
                        tbrLeakTest.Buttons("start").Enabled = True
                        tbrLeakTest.Buttons("start").ToolTipText = "Continue; Vapor Tank is ready"
'                    End If
          
                Case VBWBPAUSE                                  ' Pause for WatterBath
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If Not tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = True
                    tbrLeakTest.Buttons("start").Enabled = True
                    tbrLeakTest.Buttons("start").ToolTipText = "Continue; WaterBath is ready"
          
                Case VBPOSTLOADOPER                              ' PostLoad Pause for Operator
                    If (Not tbrLeakTest.Buttons("continue").Visible) Then tbrLeakTest.Buttons("continue").Visible = True
                    tbrLeakTest.Buttons("continue").Enabled = True
                    tbrLeakTest.Buttons("continue").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = False
          
                Case VBPOSTPURGEOPER                              ' PostPurge Pause for Operator
                    If (Not tbrLeakTest.Buttons("continue").Visible) Then tbrLeakTest.Buttons("continue").Visible = True
                    tbrLeakTest.Buttons("continue").Enabled = True
                    tbrLeakTest.Buttons("continue").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = False
          
                Case VBPURGEWAIT
                    tbrLeakTest.Buttons("continue").Visible = False
                    tbrLeakTest.Buttons("pause").Visible = False
                    tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    tbrLeakTest.Buttons("start").Visible = True
                    If USINGPASLOCALCONTROL And PAS_INFO(pasTEMPERATURE).timeOut Then
                        ' Local PAS Temperature Control Timeout
                        tbrLeakTest.Buttons("start").Enabled = True
                        tbrLeakTest.Buttons("start").ToolTipText = "Reset PAS Temperature Timeout"
                    ElseIf USINGPASLOCALCONTROL And PAS_INFO(pasMOISTURE).timeOut Then
                        ' Local PAS Moisture Control Timeout
                        tbrLeakTest.Buttons("start").Enabled = True
                        tbrLeakTest.Buttons("start").ToolTipText = "Reset PAS Moisture Timeout"
                    Else
                        ' Not using local PAS control or no timeouts
                        tbrLeakTest.Buttons("start").Enabled = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).Enabled
                        tbrLeakTest.Buttons("start").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
                    End If
            
                Case VBPAUSEVACSW                                   ' System Vacuum Switch Off; Wait for Resume from Operator after Vacuum Switch is On
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If Not tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = True
                    tbrLeakTest.Buttons("start").Enabled = IIf(Alm_SystemVacSw, False, True)
                    tbrLeakTest.Buttons("start").ToolTipText = IIf(Alm_SystemVacSw, "Cannot Resume until System Vacuum Switch is True", "Resume; System Vacumm Switch is now True")
                
                Case VBPAUSEBYUSER
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If Not tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = True
                    tbrLeakTest.Buttons("start").Enabled = True
                    tbrLeakTest.Buttons("start").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
            
                Case VBCOURSEWAIT
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = False
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If Not tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = True
                    tbrLeakTest.Buttons("start").Enabled = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).Enabled
                    tbrLeakTest.Buttons("start").ToolTipText = Stn_ContinueBtn(StationControl(DispStn, DispShift).Mode).ToolTipText
            
                Case Else
                    If tbrLeakTest.Buttons("continue").Visible Then tbrLeakTest.Buttons("continue").Visible = False
                    If Not tbrLeakTest.Buttons("pause").Visible Then tbrLeakTest.Buttons("pause").Visible = True
                    tbrLeakTest.Buttons("pause").Enabled = True
                    If Not tbrLeakTest.Buttons("stop").Visible Then tbrLeakTest.Buttons("stop").Visible = True
                    tbrLeakTest.Buttons("stop").Enabled = True
                    If tbrLeakTest.Buttons("start").Visible Then tbrLeakTest.Buttons("start").Visible = False
                    tbrLeakTest.Buttons("start").Enabled = False
                    tbrLeakTest.Buttons("start").ToolTipText = ""
            
            End Select
            ' no toolTip if button is not enabled
            tbrLeakTest.Buttons("start").ToolTipText = IIf(tbrLeakTest.Buttons("start").Enabled, tbrLeakTest.Buttons("start").ToolTipText, "")
            
        End If
        
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub mnuAbout_Click()
    'About
    menuAbout
End Sub

Private Sub mnuAirLog_Click()
    menuViewAirLog
End Sub

Private Sub mnuButane_Click()
    menuButane
End Sub

Private Sub mnuCalibration_Click()
    menuCalibration
End Sub

Private Sub mnuCanisters_Click()
    menuCanisters
End Sub

Private Sub mnuConfiguration_Click()
    ' Configuration
    menuConfiguration
End Sub

Private Sub mnuCopyFile_Click()
    menuCopyFile
End Sub

Private Sub mnuCourses_Click()
    menuCourses
End Sub

Private Sub mnuFirstAid_Click()
    menuFirstAid
End Sub

Private Sub mnuFuelUseLog_Click()
    menuViewFuelUseLog
End Sub

Private Sub mnuOotMonitor_Click()
    menuOotMonitor
End Sub

Private Sub mnuPrintFile_Click()
    menuPrintFile
End Sub

Private Sub mnuEventLog_Click()
    menuEventLog
End Sub

Private Sub mnuExit_Click()
    ' Exit Program
    menuExit
End Sub

Private Sub mnuIOMonitor_Click()
    menuIoMonitor
End Sub

Private Sub mnuJoblist_Click()
    menuJobList
End Sub

Private Sub mnuLogin_Click()
    menuLogin
End Sub

Private Sub mnuLogout_Click()
    menuLogout
End Sub

Private Sub mnuOperatorManual_Click()
    menuOperatorManual
End Sub

Private Sub mnuPurgeProfiles_Click()
    menuPurgeProfiles
End Sub

Private Sub mnuRecipes_Click()
    menuRecipes
End Sub

Private Sub mnuReviewData_Click()
    ' Review Previous Cycle Data
    menuReview
End Sub

Private Sub mnuScaleMonitor_Click()
    menuScaleMonitor
End Sub

'Private Sub mnuLeakTest_Click()
'    menuLeakTest
'End Sub

Private Sub mnuSysdef_Click()
    ' Select System Definition
    menuSysdef
End Sub

Private Sub mnuTomCanLoad_Click()
    menuRemCanLoad
End Sub

Private Sub mnuWatchData_Click()
    ' Watch Current Cycle Data
    menuWatch
End Sub

Private Sub tbrNavigate_ButtonClick(ByVal Button As MSComctlLib.Button)
   ' Use the Key property with the SelectCase statement to specify
   ' an action.
   Select Case Button.Key
       Case Is = "overview"
            menuOverview
       Case Is = "stndetail"
            menuStnDetail
       Case Is = "reviewdata"
            menuReview
       Case Is = "watchdata"
            ' Watch Current Cycle Data
            menuWatch
       Case Is = "login"
            menuLogin
       Case Is = "logout"
            menuLogout
       Case Is = "copyfiles"
            ' CopyFiles
            menuCopyFile
       Case Is = "printfiles"
            ' PrintFiles
            menuPrintFile
       Case Is = "butane"
            menuButane
       Case Is = "fueluselog"
            menuViewFuelUseLog
       Case Is = "canisters"
            menuCanisters
       Case Is = "recipes"
            menuRecipes
       Case Is = "courses"
            menuCourses
       Case Is = "purgeprofile"
            menuPurgeProfiles
       Case Is = "tomcanload"
            menuRemCanLoad
       Case Is = "configuration"
            ' Configuration
            menuConfiguration
       Case Is = "sysdef"
            ' Select System Definition
            menuSysdef
       Case Is = "eventlog"
            menuEventLog
       Case Is = "joblist"
            menuJobList
       Case Is = "calibration"
            menuCalibration
       Case Is = "iomonitor"
            menuIoMonitor
       Case Is = "scalemonitor"
            menuScaleMonitor
       Case Is = "leaktest"
            menuLeakTest
       Case Is = "simulation"
            ' Simulation
            menuSimulation
       Case Is = "opermanual"
            ' Operators Manual
            menuOperatorManual
       Case Is = "beaconoff"
            ' TurnOff Beacon
            menuBeaconOff
       Case Is = "hornoff"
            ' TurnOff Horn
            menuHornOff
       Case Is = "firstaid"
            ' First Aid File Save
            menuFirstAid
'       Case Is = "close"
'            CloseScreen

   End Select
End Sub

Private Sub cmdUpStn_Click()
    StnUp
End Sub

Private Sub cmdDnStn_Click()
    StnDown
End Sub

Private Sub StnDown()
Dim lstStn
    lstStn = DispStn
    DispStn = IIf(DispStn <= 1, LAST_STN, DispStn - 1)
    If DispStn <> lstStn Then
        If (STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Then
            UpdateScreen
        Else
            frmStnDetail.Show
        End If
    End If
End Sub

Private Sub StnUp()
Dim lstStn
    lstStn = DispStn
    DispStn = IIf(DispStn >= LAST_STN, 1, DispStn + 1)
    If DispStn <> lstStn Then
        If (STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Then
            UpdateScreen
        Else
            frmStnDetail.Show
        End If
    End If
End Sub

Private Sub txtEndOp_Change()
    If (StationControl(DispStn, DispShift).Mode = VBIDLE) Then JobInfo(DispStn, DispShift).End_Op = txtEndOp.text
End Sub

Private Sub txtEngineer_Change()
    If (StationControl(DispStn, DispShift).Mode = VBIDLE) Then JobInfo(DispStn, DispShift).Engineer = txtEngineer.text
End Sub

Private Sub txtSgN2_Change()
    SGN2 = ValueFromText(txtSgN2.text)
End Sub

Private Sub tbrLeakTest_ButtonClick(ByVal Button As MSComctlLib.Button)
   ' Use the Key property with the SelectCase statement to specify
   ' an action.
   Select Case Button.Key
       Case Is = "prevstn"
            StnDown
       Case Is = "nextstn"
            StnUp
'       Case Is = "prevshift"
'            ShiftDown
'       Case Is = "nextshift"
'            ShiftUp
       Case Is = "alarmlog"
            AlarmLog
       Case Is = "ootlog"
            OOTLog
       Case Is = "statsum"
            StatisticsSummary
       Case Is = "joblog"
            JobLog
       Case Is = "opercomment"
            frmOperComment.WhichStn = DispStn
            frmOperComment.WhichShift = DispShift
            frmOperComment.Show
       Case Is = "start"
            If StationControl(DispStn, DispShift).Mode = VBIDLE Then
                If StationCanister(DispStn, DispShift).Validated Then
                    ' start the station
                    StationStart
                Else
                    txtLeakTestMsg.text = vbCrLf & "Must have Valid CANISTER defined first"
                End If
            Else
                StationContinue "Operator"
            End If
       Case Is = "continue"
            StationContinue "Operator"
       Case Is = "pause"
            StationPause "Operator"
       Case Is = "stop"
            StationStop
       Case Is = "canisters"
            frmCanRecipe.Show
            frmCanRecipe.ChgCanRcpMode (CInt(STATIONMODE))
            frmCanRecipe.InitCanRcp
       Case Is = "recipes"
            If StationCanister(DispStn, DispShift).Validated Then
                frmRecipe.Show
                frmRecipe.ChgRecipeMode (CInt(STATIONMODE))
                frmRecipe.InitRecipe
                frmRecipe.Hide
                frmRecipe.Show
            Else
                txtLeakTestMsg.text = vbCrLf & "Must have Valid CANISTER defined first"
            End If
       Case Is = "purgeprofile"
            If (StationRecipe(DispStn, DispShift).Purge_Method = PURGEBYPROFILE) Then
                frmPurgeProfile.Show
                frmPurgeProfile.ChgProfileMode (CInt(STATIONMODE))
                frmPurgeProfile.InitProfile
            Else
                txtLeakTestMsg.text = "Current Recipe does not use Purge-By-Profile"
            End If
       Case Is = "courses"
            JobSeqAutoEdit = True
            frmCourses.Show
            frmCourses.ChgJobSeqMode (CInt(STATIONMODE))
            frmCourses.InitSeqRcp
       Case Is = "fuelsupply"
            'Fuel Supply Screen
            If ((STN_INFO(DispStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE)) Then
                If STN_INFO(DispStn).ADF_TANKTYPE <> 0 Then
                    frmFuelSupply.Show
                Else
                    txtLeakTestMsg.text = "Current Station does not have ADF Control"
                End If
            Else
                txtLeakTestMsg.text = "Current Station does not support Live Fuel"
            End If
    
    '   Case Is = "close"
    '        CloseScreen
   End Select
End Sub

Private Sub AlarmLog()
    If StationControl(DispStn, DispShift).DBFile <> "" Then
       View_Alarm StationControl(DispStn, DispShift).Job_Number, DispStn, DispShift
    Else
       txtLeakTestMsg.text = "Station MUST be running to view STATION ALARMS"
    End If
End Sub

Private Sub OOTLog()
    If StationControl(DispStn, DispShift).DBFile <> "" Then
        View_OOT StationControl(DispStn, DispShift).Job_Number, DispStn, DispShift
    Else
        txtLeakTestMsg.text = "Station MUST be running for Tolerances"
    End If
End Sub

Private Sub JobLog()
    If StationControl(DispStn, DispShift).DBFile <> "" Then
        View_JobLog StationControl(DispStn, DispShift).Job_Number, DispStn, DispShift
    Else
        txtLeakTestMsg.text = "Station MUST be running for Job Event Log"
    End If
End Sub

Public Sub StationContinue(ByVal byWhom As String)
    If Pause_Alarm Then
        txtLeakTestMsg.text = "Can not CONTINUE while system is paused"
        Exit Sub                                    ' Wait for no alarming conditions
    End If
    If CheckPass("R", msgSHOW) Then
        '   reset (optional) Local PAS Timeout Timers
        If USINGPASLOCALCONTROL Then
            If PAS_INFO(pasTEMPERATURE).timeOut Then
                PAS_INFO(pasTEMPERATURE).TimeOutDuration = 0#
                PAS_INFO(pasTEMPERATURE).timeOut = False
            End If
            If PAS_INFO(pasMOISTURE).timeOut Then
                PAS_INFO(pasMOISTURE).TimeOutDuration = 0#
                PAS_INFO(pasMOISTURE).timeOut = False
            End If
        End If
        ' Set the Continue-Button-Has-Been-Pressed flag for this Station
        StationControl(DispStn, DispShift).ContinueRequest = True
        ' build msg for Stn Detail screen
        If InStr(tbrLeakTest.Buttons("start").ToolTipText, "Reset") > 0 Then
            sStr = Right(tbrLeakTest.Buttons("start").ToolTipText, Len(tbrLeakTest.Buttons("start").ToolTipText) - 6)
            sStr = sStr & " reset by " & byWhom
        ElseIf InStr(tbrLeakTest.Buttons("start").ToolTipText, "Resume") > 0 Then
            sStr = Right(tbrLeakTest.Buttons("start").ToolTipText, Len(tbrLeakTest.Buttons("start").ToolTipText) - 7)
            sStr = sStr & " resumed by " & byWhom
        ElseIf InStr(tbrLeakTest.Buttons("start").ToolTipText, "Cancel") > 0 Then
            sStr = Right(tbrLeakTest.Buttons("start").ToolTipText, Len(tbrLeakTest.Buttons("start").ToolTipText) - 7)
            sStr = sStr & " cancelled by " & byWhom
        Else
            sStr = "Operator continued " & ModeDescShort(StationControl(DispStn, DispShift).Mode)
        End If
        ' display msg
        txtLeakTestMsg.text = sStr
    End If
End Sub

Public Sub StationPause(ByVal byWhom As String)
'
'  Pause a station because Operator pressed Pause
'
      
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 1066, 3762
      
    ' start pause time
    StationControl(DispStn, DispShift).PausedDts = Now
    Write_ELog byWhom & " Paused Station #" & Format(DispStn, "0") & " Shift #" & Format(DispShift, "0")
    Write_JLog DispStn, DispShift, byWhom & " Paused Station"
    
    ' Save the current mode for the continue button
    StationControl(DispStn, DispShift).Mode_PauseSave = StationControl(DispStn, DispShift).Mode
    ' save elapsed hours so far
    Select Case StationControl(DispStn, DispShift).Mode
        Case VBLEAK
            LeakCheckControl.ElapsedHours_Prev = LeakCheckControl.ElapsedHours
        Case VBLOAD
            LoadControl(DispStn, DispShift).ElapsedHours_Prev = LoadControl(DispStn, DispShift).ElapsedHours
        Case VBPURGE
            PurgeControl(DispStn, DispShift).ElapsedHours_Prev = PurgeControl(DispStn, DispShift).ElapsedHours
        Case Else
    End Select
   
    If StationControl(DispStn, DispShift).Mode_PauseSave = VBLOAD Then              ' station was loading before Pause
        If StationRecipe(DispStn, DispShift).Load_Method = LOADBYTIME Or StationRecipe(DispStn, DispShift).Load_Method = LOADBYWC Then
            StationControl(DispStn, DispShift).PauseAlarmStartTime = Now            ' save pause time on load by time
        End If
    End If
   
    If StationControl(DispStn, DispShift).Mode_PauseSave = VBPURGE Then             ' station was purging before Pause
        StationControl(DispStn, DispShift).PauseAlarmStartTime = Now                ' save pause time
    End If
   
    '  Turn Off Station MFCs
    ShutdownStnMFCs DispStn, DispShift
    '   Station Valves
    Close_Stn_Valves DispStn, DispShift
    '   Scale Valves
    If StationRecipe(DispStn, DispShift).UsePriScale And StationControl(DispStn, DispShift).PriScaleStn > 0 _
            And StationControl(DispStn, DispShift).PriScaleStn < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(DispStn, DispShift).PriScaleStn, isPriAuxVentSol, cOFF
    End If
    If StationRecipe(DispStn, DispShift).PurgeAuxCan And StationControl(DispStn, DispShift).AuxScaleStn > 0 Then
        Stn_OutDigital StationControl(DispStn, DispShift).AuxScaleStn, isAuxPurgeSol, cOFF
    End If
'    ' Release Common (Leak) Pressure Transducer (if this station is using it)
    If LeakCheckControl.station = DispStn Then
        LeakCheckControl.station = 0
        LeakCheckControl.Shift = 0
        LeakCheckControl.Phase = 0
        LeakCheckControl.ElapsedHours = 0
        LeakCheckControl.ElapsedHours_Prev = 0
    End If
    
    ' set mode to Paused
    Select Case byWhom
        Case "Oper", "Operator"
            StationControl(DispStn, DispShift).Mode = VBPAUSEBYUSER
        Case "ADF"
            StationControl(DispStn, DispShift).Mode = VBGASPAUSE
        Case Else
            StationControl(DispStn, DispShift).Mode = VBPAUSEBYUSER
    End Select
          
 ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub StationStart()
    If CheckPass("R", msgSHOW) Then
    
        If Pause_Alarm <> 0 Then                                ' error, in alarm
            txtLeakTestMsg.text = "Can not START while system is paused"
            Exit Sub
        End If
        If StationControl(DispStn, DispShift).Mode <> VBIDLE Then          ' error, in use
            txtLeakTestMsg.text = "Start Button pushed when not in IDLE...Error"
            Exit Sub
        End If
        
'        ' if setup failed, notify user
'        If Not StationSequence(DispStn, DispShift).Validated Then
'            txtLeakTestMsg.text = "Invalid Job Sequence; Nothing to do"
'            Exit Sub
'        End If
        
        If Stn_OperReportNameIsValid = False _
            And _
                (SysConfig.ReportFileName1stPart = RPT_OPERENTER _
                Or SysConfig.ReportFileName2ndPart = RPT_OPERENTER _
                Or SysConfig.ReportFileName3rdPart = RPT_OPERENTER) Then
            txtLeakTestMsg.text = "Must have Validated Report File Name before Starting"
            Exit Sub
        End If
    
        
       ' set "Start Station" flag (used by Module 2)
        StationControl(DispStn, DispShift).StartRequest = True
        txtLeakTestMsg.text = vbCrLf & "Starting Station #" & Format(DispStn, "#0")
        
    End If

End Sub

Private Sub StationStop()
    If CheckPass("R", msgSHOW) Then
        frmStop.Show
        frmStop.cmdYES = False
    End If
End Sub

Private Sub UpdateStatusBars()
    ' Status Bar #1
    pnlEstop.BackColor = frmMainMenu.pnlEstop.BackColor
    pnlFlow.BackColor = frmMainMenu.pnlFlow.BackColor
    pnlBtn20.BackColor = frmMainMenu.pnlBtn20.BackColor
    pnlDoors.BackColor = frmMainMenu.pnlDoors.BackColor
    pnlComms.BackColor = frmMainMenu.pnlComms.BackColor
'    pnlPAcomm.BackColor = frmMainMenu.pnlPAcomm.BackColor
    
    pnlEstop.ToolTipText = frmMainMenu.pnlEstop.ToolTipText
    pnlFlow.ToolTipText = frmMainMenu.pnlFlow.ToolTipText
    pnlBtn20.ToolTipText = frmMainMenu.pnlBtn20.ToolTipText
    pnlDoors.ToolTipText = frmMainMenu.pnlDoors.ToolTipText
    pnlComms.ToolTipText = frmMainMenu.pnlComms.ToolTipText
'    pnlPAcomm.ToolTipText = frmMainMenu.pnlPAcomm.ToolTipText
    
    pnlMix.BackColor = frmMainMenu.pnlMix.BackColor
    pnlMix.ToolTipText = frmMainMenu.pnlMix.ToolTipText
    pnlMix.Top = frmMainMenu.pnlMix.Top
    
'    pnlMessageFrame.Width = frmMainMenu.pnlMessageFrame.Width
    pnlMessage.Font = frmMainMenu.pnlMessage.Font
    pnlMessage.FontSize = frmMainMenu.pnlMessage.FontSize
    pnlMessage.Width = frmMainMenu.pnlMessage.Width
    pnlMessage.BackColor = SysMessage_BackColor
    pnlMessage.ForeColor = SysMessage_ForeColor
    pnlMessage.Caption = SysMessage_Text
    pnlMessage.ToolTipText = SysMessage_Tooltip
    
'    pnlPurgeAir.Left = frmMainMenu.pnlPurgeAir.Left
'    pnlPurgeAir.Width = frmMainMenu.pnlPurgeAir.Width
    pnlPurgeAir.ForeColor = PurgeAirMsg_ForeColor
    pnlPurgeAir.Caption = PurgeAirMsg_Text
    pnlPurgeAir.ToolTipText = PurgeAirMsg_ToolTip
End Sub

'  *************  ADDED  ********************
'Private Sub cmdDnShift_Click()
'    ShiftDown
'End Sub
'  *************  ADDED  ***********************
'Private Sub cmdUpShift_Click()
'    ShiftUp
'End Sub

'Private Sub ShiftDown()
'Dim lstShift
'    lstShift = DispShift
'    DispShift = IIf(DispShift <= 1, NR_SHIFT, DispShift - 1)
'    If DispShift <> lstShift Then
'        txtLeakTestMsg.text = " "
'ADDED  *************
'        stnDtl_StnMode_Last(DispStn, DispShift) = -1        ' force update of station mode indicator
'        Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
'        ChartXYValues DispStn, DispShift
'        Update_Text DispStn, DispShift
'        Update_LKStn DispStn, DispShift
'        If frmRecipe.Visible Then frmRecipe.RecipeDisplay_ByStnShift               ' force update of station recipe
'************
'    End If
'End Sub

Private Sub ShiftUp()
Dim lstShift
    lstShift = DispShift
    DispShift = IIf(DispShift = NR_SHIFT, 1, DispShift + 1)
    If DispShift <> lstShift Then
        txtLeakTestMsg.text = " "
'ADDED  *************
'        stnDtl_StnMode_Last(DispStn, DispShift) = -1          ' force update of station mode indicator
'        Update_Text DispStn, DispShift
'        Update_LKStn DispStn, DispShift
'        If frmRecipe.Visible Then frmRecipe.RecipeDisplay_ByStnShift               ' force update of station recipe
'************
    End If
End Sub

Private Sub StatisticsSummary()
    frmSummary.Show
End Sub

Private Sub cmdLoadControllers_Click()
    ' Load Controllers Config Data
    Load_Controllers
End Sub

Private Sub txtStartOp_Change()
    If (StationControl(DispStn, DispShift).Mode = VBIDLE) Then JobInfo(DispStn, DispShift).Start_Op = txtStartOp.text
End Sub

Private Sub txtVehicle_Change()
    If (StationControl(DispStn, DispShift).Mode = VBIDLE) Then JobInfo(DispStn, DispShift).Vehicle = txtVehicle.text
End Sub

' ****  ADDED  ****************************
'
'Sub Update_LKStn(ByVal Index As Integer, ByVal Index2 As Integer)
'SetErrModule 1066, 1
''If UseLocalErrorHandler Then On Error GoTo localhandler
'
'Dim dMsg As String
'
'txtDebug1.Visible = False
'
'    ' Update Text
'    If StationControl(Index, Index2).Course <> stnDtl_StnCourse_Last(Index, Index2) Then UpdateStnRcpDsc Index, Index2
'    If StationControl(Index, Index2).Course <> stnDtl_StnCourse_Last(Index, Index2) Then Update_Text Index, Index2
'    If StationControl(Index, Index2).Mode <> stnDtl_StnMode_Last(Index, Index2) Then Update_Text Index, Index2
'    ' Update the Navigate Toolbar buttons
'    UpdateNavigateBtns
'
'    ' Status Bars
'    UpdateStatusBars
'
'    '**************************************************************
'    ChgErrModule 1066, 2
'    txtDspStn.text = Index
'    txtDspShift.text = Index2
'    ' Clock (PurgeAir) Panel
'    pnlPurgeAir.ForeColor = frmMainMenu.pnlPurgeAir.ForeColor
'    pnlPurgeAir.Caption = frmMainMenu.pnlPurgeAir.Caption
'
'    ChgErrModule 1066, 3
'    ' **************************************************************************************
'    ' MODE DISPLAY
'    If StationControl(Index, Index2).Mode <> stnDtl_StnMode_Last(Index, Index2) _
'        Or DispStn <> stnDtl_DispStn_Last _
'        Or DispShift <> stnDtl_DispShift_Last _
'        Or StationControl(Index, Index2).Mode = VBLEAK _
'        Or StationControl(Index, Index2).Mode = VBLOAD _
'        Or StationControl(Index, Index2).Mode = VBPURGE _
'        Or StationControl(Index, Index2).Mode = VBPOSTLEAK _
'        Or StationControl(Index, Index2).Mode = VBPOSTLOAD _
'        Or StationControl(Index, Index2).Mode = VBPOSTPURGE _
'        Or StationControl(Index, Index2).Mode = VBSCALEWAIT _
'        Or StationControl(Index, Index2).Mode = VBSTARTWAIT Then
'
'        ' only update if mode has changed (or description has a variable in it)
'        Select Case StationControl(Index, Index2).Mode
'            Case VBLEAK
'                ' Leak Check - add leak check phase description
'                tempText = ModeDescShort(VBLEAK) & " - " & LeakPhaseDesc(LeakCheckControl.Phase) & " " & LeakMethodDesc(LeakCheckControl.Method)
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBLOAD
'                ' Loading or Waiting for Scales to Settle?
'                If LoadControl(Index, Index2).Phase = LoadPause Then
'                    ' Waiting for Scales to Settle
'                    tempText = " Load Settling for "
'                    tempText = tempText & Format(StationConfig(Index, Index2).LoadSettleTime, "##0.0#")
'                    tempText = tempText & LoadTypeDesc2(LOADBYTIME)
'                    tempText = tempText
'                    tempText = tempText & LoadTypeDesc3(LOADBYTIME)
'                Else
'                    ' Loading - add load method description
'                    Select Case StationRecipe(Index, Index2).Load_MethodSave
'                        Case NOLOAD
'                            tempText = LoadTypeDesc(NOLOAD)
'                            tempText = tempText
'                            tempText = tempText & LoadTypeDesc2(NOLOAD)
'                            tempText = tempText
'                            tempText = tempText & LoadTypeDesc3(NOLOAD)
'                        Case LOADBYTIME
'                            tempText = LoadTypeDesc(LOADBYTIME)
'                            tempText = tempText & Format(StationRecipe(Index, Index2).Load_Time, "##0")
'                            tempText = tempText & LoadTypeDesc2(LOADBYTIME)
'                            tempText = tempText
'                            tempText = tempText & LoadTypeDesc3(LOADBYTIME)
'                        Case LOADBYWC
'                            tempText = LoadTypeDesc(LOADBYWC)
'                            tempText = tempText & Format(StationRecipe(Index, Index2).WC_MultSave, "##0.#")
'                            tempText = tempText & LoadTypeDesc2(LOADBYWC)
'                            tempText = tempText & Format(StationRecipe(Index, Index2).EPAFill, "##0")
'                            tempText = tempText & LoadTypeDesc3(LOADBYWC)
'                        Case LOADBYWEIGHT
'                            tempText = LoadTypeDesc(LOADBYWEIGHT)
'                            If Int(StationRecipe(Index, Index2).Load_Wt) = StationRecipe(Index, Index2).Load_Wt Then
'                                ' no digits to the right of the decimal point
'                                tempText = tempText & Format(StationRecipe(Index, Index2).Load_Wt, "##0")
'                            Else
'                                ' digit(s) to the right of the decimal point
'                                tempText = tempText & Format(StationRecipe(Index, Index2).Load_Wt, "##0.##")
'                            End If
'                            tempText = tempText & LoadTypeDesc2(LOADBYWEIGHT)
'                            tempText = tempText
'                            tempText = tempText & LoadTypeDesc3(LOADBYWEIGHT)
'                        Case LOADBYBREAKTHRU
'                            tempText = LoadTypeDesc(LOADBYBREAKTHRU)
'                            If Int(StationRecipe(Index, Index2).LoadBreakthrough) = StationRecipe(Index, Index2).LoadBreakthrough Then
'                                ' no digits to the right of the decimal point
'                                tempText = tempText & Format(StationRecipe(Index, Index2).LoadBreakthrough, "##0")
'                            Else
'                                ' digit(s) to the right of the decimal point
'                                tempText = tempText & Format(StationRecipe(Index, Index2).LoadBreakthrough, "##0.##")
'                            End If
'                            tempText = tempText & LoadTypeDesc2(LOADBYBREAKTHRU)
'                            tempText = tempText
'                            tempText = tempText & LoadTypeDesc3(LOADBYBREAKTHRU)
'                        Case LOADBYFID
'                            tempText = LoadTypeDesc(LOADBYFID)
'                            tempText = tempText & Format(StationRecipe(Index, Index2).FIDmg, "#####0")
'                            tempText = tempText & LoadTypeDesc2(LOADBYFID)
'                            tempText = tempText
'                            tempText = tempText & LoadTypeDesc3(LOADBYFID)
'                    End Select
'                End If
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBPURGE
'                ' Purging or Waiting for Scales to Settle?
'                If PurgeControl(Index, Index2).Phase = PurgePause Then
'                    ' Waiting for Scales to Settle
'                    tempText = " Purge Settling for "
'                    tempText = tempText & Format(StationConfig(Index, Index2).PurgeSettleTime, "##0.0#")
'                    tempText = tempText & LoadTypeDesc2(PURGEBYTIME)
'                    tempText = tempText
'                    tempText = tempText & LoadTypeDesc3(PURGEBYTIME)
'                Else
'                    ' Purge - add Purge method description
'                    Select Case StationRecipe(Index, Index2).Purge_Method
'                        Case NOPURGE
'                            tempText = "No Purge"
'                        Case PURGEBYTIME
'                            tempText = ModeDescShort(VBPURGE) & " for " & StationRecipe(Index, Index2).Purge_Time & " Minute"
'                            If StationRecipe(Index, Index2).Purge_Time > 1 Then tempText = tempText & "s"
'                        Case PURGEBYVOLUME
'                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(Index, Index2).Purge_Can_Vol & " Canister Volume"
'                            If StationRecipe(Index, Index2).Purge_Can_Vol <> 1 Then tempText = tempText & "s"
'                        Case PURGEAUXONLY
'                            tempText = ModeDescShort(VBPURGE) & " Aux Can for " & StationRecipe(Index, Index2).Purge_AuxTime & " Minute"
'                            If StationRecipe(Index, Index2).Purge_AuxTime > 1 Then tempText = tempText & "s"
'                        Case PURGEBYPROFILE
'                            tempText = ModeDescShort(VBPURGE) & " " & " by Profile"
'                        Case PURGEBYWC
'                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(Index, Index2).Purge_TargetWC & " % of Work Cap"
'                        Case PURGETOTARGET
'                            tempText = ModeDescShort(VBPURGE) & " to " & StationRecipe(Index, Index2).Purge_TargetWeight & " grams"
'                        Case PURGETOUNDOLOAD
'                            tempText = ModeDescShort(VBPURGE) & " to " & " Undo Load"
'                        Case PURGEBYLITERS
'                            tempText = ModeDescShort(VBPURGE) & " " & StationRecipe(Index, Index2).Purge_Liters & " liter"
'                            If StationRecipe(Index, Index2).Purge_Liters <> 1 Then tempText = tempText & "s"
'                    End Select
'                End If
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBPOSTLEAK
'                ' Post LeakCheck Pause
'                tempText = ModeDescShort(VBPOSTLEAK)
'                tempText = tempText & " for "
'                tempText = tempText & Format(StationRecipe(Index, Index2).PauseLeakTime, "##0.0#")
'                tempText = tempText & LoadTypeDesc2(LOADBYTIME)
'                tempText = tempText
'                tempText = tempText & LoadTypeDesc3(LOADBYTIME)
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBPOSTLOAD
'                ' Post Load Pause
'                tempText = ModeDescShort(VBPOSTLOAD)
'                tempText = tempText & " for "
'                tempText = tempText & Format(StationRecipe(Index, Index2).PauseLoadTime, "##0.0#")
'                tempText = tempText & LoadTypeDesc2(LOADBYTIME)
'                tempText = tempText
'                tempText = tempText & LoadTypeDesc3(LOADBYTIME)
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBPOSTPURGE
'                ' Post Purge Pause
'                tempText = ModeDescShort(VBPOSTPURGE)
'                tempText = tempText & " for "
'                tempText = tempText & Format(StationRecipe(Index, Index2).PausePurgeTime, "##0.0#")
'                tempText = tempText & LoadTypeDesc2(PURGEBYTIME)
'                tempText = tempText
'                tempText = tempText & LoadTypeDesc3(PURGEBYTIME)
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBSCALEWAIT
'                ' Waiting for Scale(s) - add which scale(s)
'                tempText = ModeDescShort(VBSCALEWAIT)
'                If StationRecipe(Index, Index2).UsePriScale And StationRecipe(Index, Index2).UseAuxScale Then
'                    ' Using Two Scales
'                    tempText = tempText & "s "
'                    ' Scales in use ?
'                    If Scale_In_Use(StationRecipe(Index, Index2).PriScaleNo) And Scale_In_Use(StationRecipe(Index, Index2).AuxScaleNo) Then
'                        ' Both Scales in use
'                        tempText = tempText & Format(StationRecipe(Index, Index2).PriScaleNo, "#0") & " && " & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
'                    ElseIf Scale_In_Use(StationRecipe(Index, Index2).PriScaleNo) Then
'                        ' Primary Scale in use
'                        tempText = tempText & Format(StationRecipe(Index, Index2).PriScaleNo, "#0")
'                    ElseIf Scale_In_Use(StationRecipe(Index, Index2).AuxScaleNo) Then
'                        ' Aux Scale in use
'                        tempText = tempText & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
'                    End If
'                ElseIf StationRecipe(Index, Index2).UsePriScale Then
'                    ' Using Only Primary Scale
'                    tempText = tempText & " "
'                    If Scale_In_Use(StationRecipe(Index, Index2).PriScaleNo) Then tempText = tempText & Format(StationRecipe(Index, Index2).PriScaleNo, "#0")
'                ElseIf StationRecipe(Index, Index2).UseAuxScale Then
'                    ' Using Only Aux Scale
'                    tempText = tempText & " "
'                    If Scale_In_Use(StationRecipe(Index, Index2).AuxScaleNo) Then tempText = tempText & Format(StationRecipe(Index, Index2).AuxScaleNo, "#0")
'                End If
'                pnlStatus.Caption = tempText
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case VBSTARTWAIT
'                ' Delayed Start - add how long
'                tempText = StartTypeDesc(StationRecipe(Index, Index2).StartMethod)
'                ' Which Method of Delay ?
'                Select Case StationRecipe(Index, Index2).StartMethod
'                    Case STARTNOW
'                        tempText = tempText & StartTypeDesc2(STARTNOW)
'                    Case STARTDELAYED
'                        tempText = tempText & Format(StationRecipe(Index, Index2).StartDelay, "##0")
'                        tempText = tempText & StartTypeDesc2(STARTDELAYED)
'                    Case STARTATDATE
'                        tempText = tempText & StartTypeDesc2(STARTATDATE)
'                        tempText = tempText & Format(StationRecipe(Index, Index2).StartDate, "D MMM, YYYY   h:mm")
'                End Select
'                pnlStatus.Caption = tempText
'                pnlStatus.ToolTipText = ""
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'            Case Else
'                pnlStatus.Caption = ModeDescShort(StationControl(Index, Index2).Mode)
'                pnlStatus.BackColor = ModeBackColor(StationControl(Index, Index2).Mode)
'                pnlStatus.ForeColor = ModeForeColor(StationControl(Index, Index2).Mode)
'        End Select
'    End If
'
'    ' JobNumber Panel
'    If StationControl(Index, Index2).DBFile = "" Then
'        pnlReport.Caption = "No Active Job File"
'    Else
'        pnlReport.Caption = "Job Number  " & StationControl(Index, Index2).Job_Number
'    End If
'    If pnlReport.BackColor <> pnlStatus.BackColor Then pnlReport.BackColor = pnlStatus.BackColor
'    If pnlReport.ForeColor <> pnlStatus.ForeColor Then pnlReport.ForeColor = pnlStatus.ForeColor
'
'    ' Station Name Panel
'    If StationControl(Index, Index2).Mode <> stnDtl_StnMode_Last(Index, Index2) _
'        Or DispStn <> stnDtl_DispStn_Last _
'        Or DispShift <> stnDtl_DispShift_Last Then
'            pnlStnName.BackColor = IIf(StationControl(Index, Index2).Mode = VBIDLE, pnlNameFrame.BackColor, pnlStatus.BackColor)
'            pnlStnName.ForeColor = IIf(StationControl(Index, Index2).Mode = VBIDLE, pnlNameFrame.ForeColor, pnlStatus.ForeColor)
'    End If
'' **************************************************************************************
'
'    ChgErrModule 1066, 445
'    ' **************************************************************************************
'    ' Is This Shift the Active Shift for this Station
'    If Stn_ActiveShift(Index) = Index2 Then
'
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ' LOAD CONCORDANCE DISPLAY
'    ' LOAD CONCORDANCE DISPLAY
'    ' LOAD CONCORDANCE DISPLAY
'    '       only display if user = APS
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ChgErrModule 1066, 1331
'    If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, Index2).LiveFuel)) Or (Not CheckPass("G", False))) Then
'        ' no concordance
'    Else
'        If StationControl(Index, Index2).Mode = VBLOAD _
'          And LoadControl(Index, Index2).Phase = LoadLoading _
'          And StationRecipe(Index, Index2).UsePriScale = True Then
'            If StationControl(Index, Index2).TestTimer > Stn_LoadEql_StartTimer(Index, Index2) Then
'                ' Open Concordance Screen
'                If (Not LoadControl(Index, Index2).ConcordanceIsOpen) Then frmConcordance.Show
'                End If
'            End If
'        End If
'    End If
'
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ' update StnDetail variables
'    ' update StnDetail variables
'    ' update StnDetail variables
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ' **************************************************************************************
'    ChgErrModule 1066, 1349
'    stnDtl_DispStn_Last = Index
'    stnDtl_DispShift_Last = Index2
'    stnDtl_StnCourse_Last(Index, Index2) = StationControl(Index, Index2).Course
'    stnDtl_StnMode_Last(Index, Index2) = StationControl(Index, Index2).Mode
'
'ResetErrModule
'End Sub
'Exit Sub

' *******  ADDED  *****************
'
'Private Sub Update_Text(Index As Integer, Index2 As Integer)
'Dim sTxt As String
'
'    'Update Text fields on detail screen
'    pnlStnName.Caption = STN_INFO(Index).desc
'    sTxt = StationSequence(Index, Index2).EstSeqDurDesc
'    txtRecipeName.text = StationRecipe(Index, Index2).Name
'    txtRcpDsc(0).text = StationRecipe(Index, Index2).desc(0)
'    txtRcpDsc(1).text = StationRecipe(Index, Index2).desc(1)
'    txtRcpDsc(2).text = StationRecipe(Index, Index2).desc(2)
'    txtCanID.text = StationCanister(Index, Index2).Description
'    Select Case StationControl(Index, Index2).LeakCheckStatus
'        Case RESULTGOOD
'            txtLeakCheckStatus.ForeColor = Good_ForeColor
'            txtLeakCheckStatus.text = "Passed LeakCheck"
''            txtLeakCheckStatus.text = StationControl(Index, index2).LcStatusDescription & " LeakCheck"
'        Case NORESULT
'            txtLeakCheckStatus.ForeColor = txtLeakCheckStatus.BackColor
''            txtLeakCheckStatus.ForeColor = MEDGRAY
'            txtLeakCheckStatus.text = StationControl(Index, Index2).LcStatusDescription
'        Case Else
'            txtLeakCheckStatus.ForeColor = Alarm_ForeColor
'            txtLeakCheckStatus.text = "Failed LeakCheck"
''            txtLeakCheckStatus.text = " LeakCheck " & StationControl(Index, index2).LcStatusDescription
'    End Select
'    lblBedVolume.Caption = Format(StationCanister(Index, Index2).WorkingVolume, "###0.00")
'    lblWorkCap.Caption = Format(StationCanister(Index, Index2).WorkingCapacity, "###0.00")
'    lblWorkCap.ForeColor = IIf((StationCanister(Index, Index2).WorkingCapacity = CSng(0)), Warning_ForeColor, Black)
'    txtEngineer.text = JobInfo(Index, Index2).Engineer
'    txtVehicle.text = JobInfo(Index, Index2).Vehicle
'    txtStartOp.text = JobInfo(Index, Index2).Start_Op
'    txtEndOp.text = JobInfo(Index, Index2).End_Op
'End Sub

