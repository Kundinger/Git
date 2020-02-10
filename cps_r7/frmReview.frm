VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReview 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Review CycleData"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15330
   ControlBox      =   0   'False
   Icon            =   "frmReview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10875
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbxBottom 
      Align           =   2  'Align Bottom
      Height          =   460
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   15270
      TabIndex        =   348
      Top             =   10410
      Width           =   15330
      Begin Threed.SSPanel pnlPurgeAir 
         Height          =   405
         Left            =   10110
         TabIndex        =   349
         Top             =   0
         Width           =   5155
         _Version        =   65536
         _ExtentX        =   9093
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "purge air"
         ForeColor       =   -2147483646
         BackColor       =   -2147483633
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
         Left            =   6030
         TabIndex        =   350
         Top             =   0
         Width           =   4080
         _Version        =   65536
         _ExtentX        =   7197
         _ExtentY        =   714
         _StockProps     =   15
         Caption         =   "message"
         ForeColor       =   -2147483646
         BackColor       =   -2147483633
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
            TabIndex        =   364
            Top             =   0
            Width           =   3960
            _Version        =   65536
            _ExtentX        =   6985
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
      Begin Threed.SSPanel pnlAlarms 
         Height          =   405
         Left            =   0
         TabIndex        =   367
         Top             =   0
         Width           =   6030
         _Version        =   65536
         _ExtentX        =   10636
         _ExtentY        =   714
         _StockProps     =   15
         BackColor       =   -2147483633
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
            TabIndex        =   368
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
            TabIndex        =   369
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
            TabIndex        =   370
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
            TabIndex        =   371
            ToolTipText     =   "A door is open"
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
            TabIndex        =   372
            ToolTipText     =   "I/O Communication Error"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "IOCOM"
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
            Left            =   5110
            TabIndex        =   373
            ToolTipText     =   "Mixture OutOfTolerance"
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
         Begin Threed.SSPanel pnlPAcomm 
            Height          =   255
            Left            =   4270
            TabIndex        =   374
            ToolTipText     =   "PurgeAirSystem Communication Not Online"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "PACOM"
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
   End
   Begin VB.Frame frmResults 
      BackColor       =   &H80000005&
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
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   14655
      Begin VB.Timer tmrClock 
         Interval        =   100
         Left            =   14040
         Top             =   240
      End
      Begin VB.Label lblDataHiLiTemplate 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data HiLi"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   13080
         TabIndex        =   346
         Top             =   1560
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblHeaderTemplate 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Left            =   13080
         TabIndex        =   345
         Top             =   840
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblDataBoldTemplate 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Bold"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   13080
         TabIndex        =   344
         Top             =   1080
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblDataRegTemplate 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Reg"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   13080
         TabIndex        =   343
         Top             =   1320
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   11640
         TabIndex        =   342
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   10200
         TabIndex        =   341
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   8760
         TabIndex        =   340
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   7320
         TabIndex        =   339
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   5880
         TabIndex        =   338
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   4440
         TabIndex        =   337
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   3000
         TabIndex        =   336
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   1560
         TabIndex        =   335
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   334
         Top             =   9120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   120
         TabIndex        =   333
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   1560
         TabIndex        =   332
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   3000
         TabIndex        =   331
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   4440
         TabIndex        =   330
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   5880
         TabIndex        =   329
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   7320
         TabIndex        =   328
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   8760
         TabIndex        =   327
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   10200
         TabIndex        =   326
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   34
         Left            =   11640
         TabIndex        =   325
         Top             =   8880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   11640
         TabIndex        =   324
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   10200
         TabIndex        =   323
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   8760
         TabIndex        =   322
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   7320
         TabIndex        =   321
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   5880
         TabIndex        =   320
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   4440
         TabIndex        =   319
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   3000
         TabIndex        =   318
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   1560
         TabIndex        =   317
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   316
         Top             =   8640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   315
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   1560
         TabIndex        =   314
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   3000
         TabIndex        =   313
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   4440
         TabIndex        =   312
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   5880
         TabIndex        =   311
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   7320
         TabIndex        =   310
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   8760
         TabIndex        =   309
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   10200
         TabIndex        =   308
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   32
         Left            =   11640
         TabIndex        =   307
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   306
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   1560
         TabIndex        =   305
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   304
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   4440
         TabIndex        =   303
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   5880
         TabIndex        =   302
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   7320
         TabIndex        =   301
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   8760
         TabIndex        =   300
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   10200
         TabIndex        =   299
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   31
         Left            =   11640
         TabIndex        =   298
         Top             =   8160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   297
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   1560
         TabIndex        =   296
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   3000
         TabIndex        =   295
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   4440
         TabIndex        =   294
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   5880
         TabIndex        =   293
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   7320
         TabIndex        =   292
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   8760
         TabIndex        =   291
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   10200
         TabIndex        =   290
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   30
         Left            =   11640
         TabIndex        =   289
         Top             =   7920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   11640
         TabIndex        =   288
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   10200
         TabIndex        =   287
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   8760
         TabIndex        =   286
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   7320
         TabIndex        =   285
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   5880
         TabIndex        =   284
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   4440
         TabIndex        =   283
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   3000
         TabIndex        =   282
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   1560
         TabIndex        =   281
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   29
         Left            =   120
         TabIndex        =   280
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   11640
         TabIndex        =   279
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   10200
         TabIndex        =   278
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   8760
         TabIndex        =   277
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   7320
         TabIndex        =   276
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   5880
         TabIndex        =   275
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   4440
         TabIndex        =   274
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   3000
         TabIndex        =   273
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   1560
         TabIndex        =   272
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   271
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   11640
         TabIndex        =   270
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   10200
         TabIndex        =   269
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   8760
         TabIndex        =   268
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   7320
         TabIndex        =   267
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   5880
         TabIndex        =   266
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   4440
         TabIndex        =   265
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   3000
         TabIndex        =   264
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   1560
         TabIndex        =   263
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   262
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   11640
         TabIndex        =   261
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   10200
         TabIndex        =   260
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   8760
         TabIndex        =   259
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   7320
         TabIndex        =   258
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   5880
         TabIndex        =   257
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   4440
         TabIndex        =   256
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   3000
         TabIndex        =   255
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   1560
         TabIndex        =   254
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   253
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   11640
         TabIndex        =   252
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   10200
         TabIndex        =   251
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   8760
         TabIndex        =   250
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   7320
         TabIndex        =   249
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   5880
         TabIndex        =   248
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   4440
         TabIndex        =   247
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   3000
         TabIndex        =   246
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   1560
         TabIndex        =   245
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   244
         Top             =   6720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   243
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   1560
         TabIndex        =   242
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   3000
         TabIndex        =   241
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   4440
         TabIndex        =   240
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   5880
         TabIndex        =   239
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   7320
         TabIndex        =   238
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   8760
         TabIndex        =   237
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   10200
         TabIndex        =   236
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   24
         Left            =   11640
         TabIndex        =   235
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   234
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   1560
         TabIndex        =   233
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   3000
         TabIndex        =   232
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   4440
         TabIndex        =   231
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   5880
         TabIndex        =   230
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   7320
         TabIndex        =   229
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   8760
         TabIndex        =   228
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   10200
         TabIndex        =   227
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   23
         Left            =   11640
         TabIndex        =   226
         Top             =   6240
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   225
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   1560
         TabIndex        =   224
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   3000
         TabIndex        =   223
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   4440
         TabIndex        =   222
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   5880
         TabIndex        =   221
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   7320
         TabIndex        =   220
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   8760
         TabIndex        =   219
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   10200
         TabIndex        =   218
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   22
         Left            =   11640
         TabIndex        =   217
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   216
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   1560
         TabIndex        =   215
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   3000
         TabIndex        =   214
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   4440
         TabIndex        =   213
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   5880
         TabIndex        =   212
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   7320
         TabIndex        =   211
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   8760
         TabIndex        =   210
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   10200
         TabIndex        =   209
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   21
         Left            =   11640
         TabIndex        =   208
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   207
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   1560
         TabIndex        =   206
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   205
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   4440
         TabIndex        =   204
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   5880
         TabIndex        =   203
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   7320
         TabIndex        =   202
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   8760
         TabIndex        =   201
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   10200
         TabIndex        =   200
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   20
         Left            =   11640
         TabIndex        =   199
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   11640
         TabIndex        =   198
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   10200
         TabIndex        =   197
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   8760
         TabIndex        =   196
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   7320
         TabIndex        =   195
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   5880
         TabIndex        =   194
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   4440
         TabIndex        =   193
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   192
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   1560
         TabIndex        =   191
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   190
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   11640
         TabIndex        =   189
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   10200
         TabIndex        =   188
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   8760
         TabIndex        =   187
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   7320
         TabIndex        =   186
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   5880
         TabIndex        =   185
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   4440
         TabIndex        =   184
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   3000
         TabIndex        =   183
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   1560
         TabIndex        =   182
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   181
         Top             =   5040
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   11640
         TabIndex        =   180
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   10200
         TabIndex        =   179
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   8760
         TabIndex        =   178
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   7320
         TabIndex        =   177
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   5880
         TabIndex        =   176
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   4440
         TabIndex        =   175
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   174
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   1560
         TabIndex        =   173
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   172
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   11640
         TabIndex        =   171
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   10200
         TabIndex        =   170
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   8760
         TabIndex        =   169
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   7320
         TabIndex        =   168
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   5880
         TabIndex        =   167
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   4440
         TabIndex        =   166
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   3000
         TabIndex        =   165
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   1560
         TabIndex        =   164
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   163
         Top             =   4560
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   11640
         TabIndex        =   162
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   10200
         TabIndex        =   161
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   8760
         TabIndex        =   160
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   7320
         TabIndex        =   159
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   5880
         TabIndex        =   158
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   157
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   156
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   155
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   154
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   11640
         TabIndex        =   153
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   10200
         TabIndex        =   152
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   8760
         TabIndex        =   151
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   7320
         TabIndex        =   150
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   5880
         TabIndex        =   149
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   4440
         TabIndex        =   148
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   147
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   146
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   145
         Top             =   4080
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   11640
         TabIndex        =   144
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   10200
         TabIndex        =   143
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   8760
         TabIndex        =   142
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   7320
         TabIndex        =   141
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   5880
         TabIndex        =   140
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   4440
         TabIndex        =   139
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   138
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   137
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   136
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   11640
         TabIndex        =   135
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   10200
         TabIndex        =   134
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   8760
         TabIndex        =   133
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   7320
         TabIndex        =   132
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   5880
         TabIndex        =   131
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   4440
         TabIndex        =   130
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   129
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   128
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   127
         Top             =   3600
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   11640
         TabIndex        =   126
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   10200
         TabIndex        =   125
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   8760
         TabIndex        =   124
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   7320
         TabIndex        =   123
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   5880
         TabIndex        =   122
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   121
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   120
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   119
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   118
         Top             =   3360
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   11640
         TabIndex        =   117
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   10200
         TabIndex        =   116
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   8760
         TabIndex        =   115
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   7320
         TabIndex        =   114
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   5880
         TabIndex        =   113
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   112
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   111
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   110
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   11640
         TabIndex        =   109
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   10200
         TabIndex        =   108
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   8760
         TabIndex        =   107
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   7320
         TabIndex        =   106
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   5880
         TabIndex        =   105
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   104
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   103
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   102
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   11640
         TabIndex        =   101
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   10200
         TabIndex        =   100
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   8760
         TabIndex        =   99
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   7320
         TabIndex        =   98
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   97
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   96
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   95
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   94
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   11640
         TabIndex        =   93
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   10200
         TabIndex        =   92
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   8760
         TabIndex        =   91
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   7320
         TabIndex        =   90
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   5880
         TabIndex        =   89
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   88
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   87
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   86
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   11640
         TabIndex        =   85
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   10200
         TabIndex        =   84
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   8760
         TabIndex        =   83
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   82
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   5880
         TabIndex        =   81
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   80
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   79
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   78
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   11640
         TabIndex        =   77
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   10200
         TabIndex        =   76
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   8760
         TabIndex        =   75
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   7320
         TabIndex        =   74
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   5880
         TabIndex        =   73
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   72
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   71
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   70
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   11640
         TabIndex        =   69
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   10200
         TabIndex        =   68
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   8760
         TabIndex        =   67
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   66
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   5880
         TabIndex        =   65
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   64
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   63
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   62
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   11640
         TabIndex        =   61
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   10200
         TabIndex        =   60
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   8760
         TabIndex        =   59
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   7320
         TabIndex        =   58
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   57
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   56
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   55
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   54
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   11640
         TabIndex        =   53
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   10200
         TabIndex        =   52
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   8760
         TabIndex        =   51
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   7320
         TabIndex        =   50
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   49
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   48
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   47
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   46
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   11640
         TabIndex        =   45
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   10200
         TabIndex        =   44
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   43
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   42
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   41
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   40
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   39
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   38
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   11640
         TabIndex        =   37
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   10200
         TabIndex        =   36
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   8760
         TabIndex        =   35
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   34
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   33
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   32
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   31
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Col 1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   30
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 18"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   11640
         TabIndex        =   29
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 17"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   10200
         TabIndex        =   28
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 16"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   27
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 15"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   26
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 14"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   25
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 13"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   24
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 12"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   23
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol08 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 08"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   11640
         TabIndex        =   22
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol07 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 07"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   10200
         TabIndex        =   21
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol06 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 06"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   8760
         TabIndex        =   20
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol05 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 05"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol04 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 04"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   18
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol03 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 03"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol02 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 02"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 01"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol01 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 11"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblDataRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "HH:MM:SS"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 10"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblHdrRowCol00 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Header 00"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.PictureBox pbxTop 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   14895
      TabIndex        =   347
      Top             =   5520
      Width           =   14895
   End
   Begin MSComctlLib.Toolbar tbrReview 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   351
      Top             =   630
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1111
      ButtonWidth     =   1058
      ButtonHeight    =   1005
      ImageList       =   "imgReviewNormal"
      DisabledImageList=   "imgReviewDisabled"
      HotImageList    =   "imgReviewHot"
      _Version        =   393216
      Begin VB.TextBox txtCourse 
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
         Height          =   245
         Left            =   6600
         TabIndex        =   366
         Text            =   "Course"
         Top             =   280
         Width           =   735
      End
      Begin VB.TextBox txtDspCourse 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6120
         TabIndex        =   365
         Text            =   "888"
         Top             =   0
         Width           =   540
      End
      Begin VB.TextBox txtMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   190
         Index           =   3
         Left            =   11760
         TabIndex        =   362
         Text            =   "msg3"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   190
         Index           =   2
         Left            =   11760
         TabIndex        =   361
         Text            =   "msg2"
         Top             =   175
         Width           =   855
      End
      Begin VB.TextBox txtMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   190
         Index           =   1
         Left            =   11760
         TabIndex        =   360
         Text            =   "msg1"
         Top             =   -10
         Width           =   855
      End
      Begin VB.ComboBox cboProcess 
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
         Height          =   345
         Left            =   8160
         TabIndex        =   359
         Text            =   "process"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtDspStn 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   358
         Text            =   "8"
         Top             =   0
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
         Height          =   245
         Left            =   1320
         TabIndex        =   357
         Text            =   "Station"
         Top             =   280
         Width           =   735
      End
      Begin VB.TextBox txtDspShift 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   356
         Text            =   "8"
         Top             =   0
         Width           =   415
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
         Height          =   245
         Left            =   2760
         TabIndex        =   355
         Text            =   "Shift"
         Top             =   280
         Width           =   735
      End
      Begin VB.TextBox txtDspCycle 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         TabIndex        =   354
         Text            =   "888"
         Top             =   0
         Width           =   540
      End
      Begin VB.TextBox txtCycle 
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
         Height          =   245
         Left            =   4680
         TabIndex        =   353
         Text            =   "Cycle"
         Top             =   280
         Width           =   735
      End
      Begin VB.TextBox txtProcess 
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
         Height          =   245
         Left            =   9360
         TabIndex        =   352
         Text            =   "Process"
         Top             =   280
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   363
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1111
      ButtonWidth     =   1058
      ButtonHeight    =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgReviewNormal 
      Left            =   12600
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":57E2
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":6434
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":7086
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":7CD8
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":892A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":957C
            Key             =   "searchprev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":A1CE
            Key             =   "searchnext"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgReviewDisabled 
      Left            =   13200
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":AE20
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":BA72
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":C6C4
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":D316
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":DF68
            Key             =   "close"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":EBBA
            Key             =   "searchprev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":F80C
            Key             =   "searchnext"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgReviewHot 
      Left            =   13800
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":1045E
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":110B0
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":11D02
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":12954
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":135A6
            Key             =   "close"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":141F8
            Key             =   "searchprev"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReview.frx":14E4A
            Key             =   "searchnext"
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
         Caption         =   "&Copy Files"
      End
      Begin VB.Menu mnuPrintFile 
         Caption         =   "&Print Files"
      End
      Begin VB.Menu beforeExit 
         Caption         =   "-"
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
         Caption         =   "&PurgeProfiles"
      End
      Begin VB.Menu mnuConfiguration 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuSysdef 
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
      Begin VB.Menu mnuFuelUseLog 
         Caption         =   "&Fuel Consumption Log"
      End
      Begin VB.Menu mnuJoblist 
         Caption         =   "&Joblist"
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
      Begin VB.Menu mnuak_client 
         Caption         =   "&AK Client"
      End
      Begin VB.Menu mnuak_server 
         Caption         =   "&AK Server"
      End
      Begin VB.Menu mnuCalibration 
         Caption         =   "&Calibration"
      End
      Begin VB.Menu mnuIOMonitor 
         Caption         =   "&I/O Monitor"
      End
      Begin VB.Menu mnuScaleMonitor 
         Caption         =   "&Scale Monitor"
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
      Begin VB.Menu beforeAbout 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About CPS release7"
      End
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 818 '''''''''''''''Form frmReview.frm ''''''''''''''''''''
Option Explicit

' Data Review Variables
Private RvwCycle, RvwFirstRecord, RvwMode, RvwRows, RvwRowsMax As Integer
Private RvwType, LstFirstRecord, RvwCharWidth As Integer
Private RvwStn As Integer
Private RvwShift As Integer
Private RvwCourse As Integer
Private RvwControl, NewCycle, AtEOF As Boolean
Private RvwMsg(1 To 3) As String
Private RvwData As Single
Private RvwFileName As String
Private RvwFileNumber As Integer
Private RvwMaxColNum, RvwNumHeaderRows As Integer
Private RvwMaxRecDsp, RvwRowIndex, RvwRecordIndex As Long
Private RvwCurRecCnt, RvwCurRowIndex As Long
Private RvwResponse As Integer
Private RvwLastClick, RvwDeltaClick As Double
Private RvwDosCmd, RvwStyle As String
Private dbDbase As Database
Private rsRecord As Recordset
Private RvwCriteria As String
Private RvwInterval As Integer
Private RvwLastReadDbFile As String
Private LastDataDTS, LastStatusDTS As Date
Private pString, pstrng2, pstrng3 As String
Private sTime, sPFlow, sPTemp, sMoist, sPvol, sPri, sAux As String
Private RvwColData(0 To 9) As ColumnsOfData

Public Sub SetRvwStn(ByVal newStn As Integer)
    RvwStn = newStn
End Sub

Public Sub JobComplete(ByVal iStn As Integer, ByVal iShift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 818, 80
    ' Review Complete; Job Complete
    If RvwLastReadDbFile = StationControl(iStn, iShift).DBFile Then
        ChgErrModule 818, 801
        rsRecord.Close
        ChgErrModule 818, 802
        dbDbase.Close
        ChgErrModule 818, 803
        RvwLastReadDbFile = " "
    End If
    ChgErrModule 818, 81
    If iStn = RvwStn Then
        ChgErrModule 818, 811
        NewMessage Message_ForeColor, " ", "Job is Complete", " "
    End If
ResetErrModule
Exit Sub

localhandler:
If err = 3420 Then
    ' Write to Event Log
    Write_ELog "vbIgnore after Error: " & err & _
      ", M" & ErrModule(0) & "-L" & ErrLevel(0) & " " & error$(err)
    ' Skip to next line, try to ignore
    Resume Next
'    ResetErrModule
'    Exit Sub
End If
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

Public Sub JobStart(ByVal iStn As Integer, ByVal iShift As Integer)
    ' A Job has Started
    If RvwStn = iStn And RvwShift = iShift Then
        RvwCycle = IIf(RvwCycle <= 1, StationRecipe(RvwStn, RvwShift).Cycles, RvwCycle - 1)
        RvwMode = StationControl(RvwStn, RvwShift).Mode
        RvwInterval = SysConfig.Load_Interval
        cboProcess.ListIndex = 0
        cboProcess.Refresh
        ChgSource
        UpdateScreen
        NewMessage Message_ForeColor, " ", "A New Job has Started", " "
    End If
End Sub

Private Sub NewMessage(ByVal iForeColor As Long, ByVal sMsg1 As String, ByVal sMsg2 As String, ByVal sMsg3 As String)
        txtMsg(1).ForeColor = iForeColor
        txtMsg(2).ForeColor = iForeColor
        txtMsg(3).ForeColor = iForeColor
        txtMsg(1).text = sMsg1
        txtMsg(2).text = sMsg2
        txtMsg(3).text = sMsg3
End Sub

Private Sub BuildToolbars()
' Create object variable for the Toolbar.
Dim btnX As MSComctlLib.Button
    
    ' ******************
    ' NAVIGATION TOOLBAR
    ' ******************
    
    ' Load the ImageLists
    tbrNavigate.ImageList = frmMainMenu.imgNavigateNormal
    tbrNavigate.DisabledImageList = frmMainMenu.imgNavigateDisabled
    tbrNavigate.HotImageList = frmMainMenu.imgNavigateHot
    
    ' Add button objects to Buttons collection using the
    ' Add method. After creating each button, set both
    ' Description and ToolTipText properties.
    
    tbrNavigate.Buttons.Add , , , tbrSeparator
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
'    'Login Screen
'    Set btnX = tbrNavigate.Buttons.Add(, "login", , tbrDefault, "login")
'    btnX.ToolTipText = "User Login"
'    btnX.Description = btnX.ToolTipText
'    'Logout
'    Set btnX = tbrNavigate.Buttons.Add(, "logout", , tbrDefault, "logout")
'    btnX.ToolTipText = "User Logout"
'    btnX.Description = btnX.ToolTipText
'
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'
'    'Copy Files
'    Set btnX = tbrNavigate.Buttons.Add(, "copyfiles", , tbrDefault, "copyfiles")
'    btnX.ToolTipText = "Copy Files"
'    btnX.Description = btnX.ToolTipText
'    'Print Files
'    Set btnX = tbrNavigate.Buttons.Add(, "printfiles", , tbrDefault, "printfiles")
'    btnX.ToolTipText = "Print Files"
'    btnX.Description = btnX.ToolTipText

'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)

    'Canisters Screen
    Set btnX = tbrNavigate.Buttons.Add(, "canisters", , tbrDefault, "can_master")
    btnX.ToolTipText = "Master Canisters"
    btnX.Description = btnX.ToolTipText
    'Recipes Screen
    Set btnX = tbrNavigate.Buttons.Add(, "recipes", , tbrDefault, "rcp_master")
    btnX.ToolTipText = "Master Recipes"
    btnX.Description = btnX.ToolTipText
    'Purge Profiles Screen
    Set btnX = tbrNavigate.Buttons.Add(, "purgeprofile", , tbrDefault, "prof_master")
    btnX.ToolTipText = "Master Purge Profiles"
    btnX.Description = btnX.ToolTipText
    'Sequence (Courses) Screen
    Set btnX = tbrNavigate.Buttons.Add(, "courses", , tbrDefault, "seq_master")
    btnX.ToolTipText = "Master Sequences"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    'Remote Tasks Screen
    Set btnX = tbrNavigate.Buttons.Add(, "tomcanload", , tbrDefault, "remotecontrol")
    btnX.ToolTipText = "Task Order Manager Tasks"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    'Configuration Screen
    Set btnX = tbrNavigate.Buttons.Add(, "configuration", , tbrDefault, "configuration")
    btnX.ToolTipText = "Configuration"
    btnX.Description = btnX.ToolTipText
    'System Definition Screen
    Set btnX = tbrNavigate.Buttons.Add(, "sysdef", , tbrDefault, "sysdef")
    btnX.ToolTipText = "System Definition"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    'Fuel Use Log
    Set btnX = tbrNavigate.Buttons.Add(, "fueluselog", , tbrDefault, "fueluselog")
    btnX.ToolTipText = "Fuel Consumption Log"
    btnX.Description = btnX.ToolTipText
    'Butane Available
    Set btnX = tbrNavigate.Buttons.Add(, "butane", , tbrDefault, "flammablegas")
    btnX.ToolTipText = "Butane Available"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
        
    'Event Log Screen
    Set btnX = tbrNavigate.Buttons.Add(, "eventlog", , tbrDefault, "eventlog")
    btnX.ToolTipText = "Event Log"
    btnX.Description = btnX.ToolTipText
    'Joblist Screen
    Set btnX = tbrNavigate.Buttons.Add(, "joblist", , tbrDefault, "joblist")
    btnX.ToolTipText = "List of Jobs"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
     
    'Station Detail Screen
    Set btnX = tbrNavigate.Buttons.Add(, "stndetail", , tbrDefault, "stndetail")
    btnX.ToolTipText = "Station Details"
    btnX.Description = btnX.ToolTipText
    'Overview Screen
    Set btnX = tbrNavigate.Buttons.Add(, "overview", , tbrDefault, "overview")
    btnX.ToolTipText = "Overview"
    btnX.Description = btnX.ToolTipText
    'Review Screen
    Set btnX = tbrNavigate.Buttons.Add(, "reviewdata", , tbrDefault, "reviewdata")
    btnX.ToolTipText = "Review Data"
    btnX.Description = btnX.ToolTipText
    'Watch Screen
    Set btnX = tbrNavigate.Buttons.Add(, "watchdata", , tbrDefault, "watchdata")
    btnX.ToolTipText = "Watch Data"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
       
    'Calibration Screen
    Set btnX = tbrNavigate.Buttons.Add(, "calibration", , tbrDefault, "calibration")
    btnX.ToolTipText = "Calibration"
    btnX.Description = btnX.ToolTipText
    'I/O Monitor Screen
    Set btnX = tbrNavigate.Buttons.Add(, "iomonitor", , tbrDefault, "iomonitor")
    btnX.ToolTipText = "I/O Monitor"
    btnX.Description = btnX.ToolTipText
    'Scale Monitor Screen
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
'            Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            'AK Server
            Set btnX = tbrNavigate.Buttons.Add(, "ak_server", , tbrDefault, "ak_server")
            btnX.ToolTipText = "AK Server"
            btnX.Description = btnX.ToolTipText
    
        Case pagClient
            ' AK Client PurgeAir Generator control
            Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'            Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            'AK Client
            Set btnX = tbrNavigate.Buttons.Add(, "ak_client", , tbrDefault, "ak_client")
            btnX.ToolTipText = "AK Client"
            btnX.Description = btnX.ToolTipText
            'AK Server
            Set btnX = tbrNavigate.Buttons.Add(, "ak_server", , tbrDefault, "ak_server")
            btnX.ToolTipText = "AK Server"
            btnX.Description = btnX.ToolTipText
        
    End Select
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
       
    'Simulation Control Panel
    Set btnX = tbrNavigate.Buttons.Add(, "simulation", , tbrDefault, "simulation")
    btnX.ToolTipText = "Simulation Control Panel"
    btnX.Description = btnX.ToolTipText
    
    If ((Com_DIO(icAlarmBeacon).addr <> 0) Or (Com_DIO(icAlarmBeacon).chan <> 0)) Then
        'TurnOff Beacon
        Set btnX = tbrNavigate.Buttons.Add(, "beaconoff", , tbrDefault, "beaconoff")
        btnX.ToolTipText = "Turn Off Beacon"
        btnX.Description = btnX.ToolTipText
    End If
    
    If ((Com_DIO(icAlarmHorn).addr <> 0) Or (Com_DIO(icAlarmHorn).chan <> 0)) Then
        'TurnOff Horn
        Set btnX = tbrNavigate.Buttons.Add(, "hornoff", , tbrDefault, "hornoff")
        btnX.ToolTipText = "Silence Horn"
        btnX.Description = btnX.ToolTipText
    End If
    
    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            
    'Operators Manual
    Set btnX = tbrNavigate.Buttons.Add(, "opermanual", , tbrDefault, "opermanual")
    btnX.ToolTipText = "Operators Manual"
    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
                            
'    'FirstAid
'    Set btnX = tbrNavigate.Buttons.Add(, "firstaid", , tbrDefault, "firstaid")
'    btnX.ToolTipText = "FirstAid File Save"
'    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
            
    ' blank space
'    Set btnX = tbrNavigate.Buttons.Add(, "fillright", , tbrPlaceholder)
'    btnX.Width = 2550 ' Placeholder width
    
    'Close Screen
'    Set btnX = tbrNavigate.Buttons.Add(, "close", , tbrDefault, "close")
'    btnX.ToolTipText = "Close Screen"
'    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    
    

    
    
    ' *******************
    ' REVIEW DATA TOOLBAR
    ' *******************
    
    ' Load the ImageLists
    tbrReview.ImageList = imgReviewNormal
    tbrReview.DisabledImageList = imgReviewDisabled
    tbrReview.HotImageList = imgReviewHot
    
    tbrReview.Buttons.Add , , , tbrSeparator
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
                
    'Station Number
    Set btnX = tbrReview.Buttons.Add(, "prevstn", , tbrDefault, "prev")
    btnX.ToolTipText = "Previous Station"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrReview.Buttons.Add(, "StnNoTxt", , tbrPlaceholder)
    btnX.Width = 750 ' Placeholder width to accommodate a textbox.
    Set btnX = tbrReview.Buttons.Add(, "nextstn", , tbrDefault, "next")
    btnX.ToolTipText = "Next Station"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
                
    'Shift Number
    Set btnX = tbrReview.Buttons.Add(, "ShiftNoTxt", , tbrPlaceholder)
    btnX.Width = 500 ' Placeholder width to accommodate a textbox.
    
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
            
    'Course Number
    If (NR_JOBSEQ > 1) Then
        Set btnX = tbrReview.Buttons.Add(, "prevcourse", , tbrDefault, "prev")
        btnX.ToolTipText = "Previous Course"
        btnX.Description = btnX.ToolTipText
        Set btnX = tbrReview.Buttons.Add(, "CourseNoTxt", , tbrPlaceholder)
        btnX.Width = 750 ' Placeholder width to accommodate a textbox.
        Set btnX = tbrReview.Buttons.Add(, "nextcourse", , tbrDefault, "next")
        btnX.ToolTipText = "Next Course"
        btnX.Description = btnX.ToolTipText
            
        Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
    Else
        txtDspCourse.Visible = False
        txtCourse.Visible = False
    End If
            
    'Cycle Number
    Set btnX = tbrReview.Buttons.Add(, "prevcycle", , tbrDefault, "prev")
    btnX.ToolTipText = "Previous Cycle"
    btnX.Description = btnX.ToolTipText
    Set btnX = tbrReview.Buttons.Add(, "CycleNoTxt", , tbrPlaceholder)
    btnX.Width = 550 ' Placeholder width to accommodate a textbox.
    Set btnX = tbrReview.Buttons.Add(, "nextcycle", , tbrDefault, "next")
    btnX.ToolTipText = "Next Cycle"
    btnX.Description = btnX.ToolTipText
        
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
            
    ' Process Combo Box
    Set btnX = tbrReview.Buttons.Add(, "comboProcess", , tbrPlaceholder)
    btnX.Width = 1500 ' Placeholder width to accommodate a combobox.
    
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
            
    'Search Previous
    Set btnX = tbrReview.Buttons.Add(, "searchprev", , tbrDefault, "searchprev")
    btnX.ToolTipText = "Search Previous"
    btnX.Description = btnX.ToolTipText
'    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
    'Search Next
    Set btnX = tbrReview.Buttons.Add(, "searchnext", , tbrDefault, "searchnext")
    btnX.ToolTipText = "Search Next"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
            
    ' Message Box
    Set btnX = tbrReview.Buttons.Add(, "msgbox", , tbrPlaceholder)
    btnX.Width = 3090                               ' Placeholder width to accommodate a label.
'    btnX.Width = IIf((NR_JOBSEQ > 1), 2000, 3090)  ' Placeholder width to accommodate a label.
    
    'OK
    Set btnX = tbrReview.Buttons.Add(, "ok", , tbrDefault, "ok")
    btnX.ToolTipText = "OK"
    btnX.Description = btnX.ToolTipText
    
    'Cancel
    Set btnX = tbrReview.Buttons.Add(, "cancel", , tbrDefault, "cancel")
    btnX.ToolTipText = "cancel"
    btnX.Description = btnX.ToolTipText
    
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
            
    'Close Screen
'    Set btnX = tbrReview.Buttons.Add(, "close", , tbrDefault, "close")
'    btnX.ToolTipText = "Close This Screen"
'    btnX.Description = btnX.ToolTipText
    
    ' blank space
'    Set btnX = tbrReview.Buttons.Add(, "fillright", , tbrPlaceholder)
'    btnX.Width = 500 ' Placeholder width
    
'    Set btnX = tbrReview.Buttons.Add(, , , tbrSeparator)
    
    
    
    ' Show form to continue configuring
    Show
    
    With txtDspStn
        .Height = 0.45 * tbrReview.Buttons("StnNoTxt").Height
        .Width = 0.5 * tbrReview.Buttons("StnNoTxt").Width
        .Top = tbrReview.Buttons("StnNoTxt").Top + 15
        .Left = tbrReview.Buttons("StnNoTxt").Left + (0.25 * tbrReview.Buttons("StnNoTxt").Width)
        .FontBold = True
        .Locked = True
    End With
    With txtStation
        .Height = 0.4 * tbrReview.Buttons("StnNoTxt").Height
        .Width = tbrReview.Buttons("StnNoTxt").Width
        .Top = tbrReview.Buttons("StnNoTxt").Top + (0.6 * tbrReview.Buttons("StnNoTxt").Height)
        .Left = tbrReview.Buttons("StnNoTxt").Left
        .FontBold = True
        .FontSize = 9
        .text = "Station"
        .Locked = True
    End With

    With txtDspShift
        .Height = txtDspStn.Height
        .Width = txtDspStn.Width
        .Top = txtDspStn.Top
        .Left = tbrReview.Buttons("ShiftNoTxt").Left + (0.5 * (tbrReview.Buttons("ShiftNoTxt").Width - txtDspShift.Width))
        .ForeColor = IIf(NR_SHIFT > 1, TitlesData_Forecolor, DKGRAY)
        .FontBold = True
        .Locked = True
    End With
    With txtShift
        .Height = txtStation.Height
        .Width = tbrReview.Buttons("ShiftNoTxt").Width
        .Top = tbrReview.Buttons("ShiftNoTxt").Top + (0.6 * tbrReview.Buttons("ShiftNoTxt").Height)
        .Left = tbrReview.Buttons("ShiftNoTxt").Left
        .ForeColor = IIf(NR_SHIFT > 1, TitlesLabel_ForeColor, DKGRAY)
        .FontBold = True
        .FontSize = 9
        .text = "Shift"
        .Locked = True
    End With

    If (NR_JOBSEQ > 1) Then
        With txtDspCourse
            .Height = txtDspStn.Height
            .Width = txtDspStn.Width
            .Top = txtDspStn.Top
            .Left = tbrReview.Buttons("CourseNoTxt").Left + (0.5 * (tbrReview.Buttons("CourseNoTxt").Width - txtDspCourse.Width))
            .FontBold = True
            .Locked = True
        End With
        With txtCourse
            .Height = txtStation.Height
            .Width = tbrReview.Buttons("CourseNoTxt").Width
            .Top = tbrReview.Buttons("CourseNoTxt").Top + (0.6 * tbrReview.Buttons("CourseNoTxt").Height)
            .Left = tbrReview.Buttons("CourseNoTxt").Left
            .FontBold = True
            .FontSize = 9
            .text = "Course"
            .Locked = True
        End With
    End If
    
    With txtDspCycle
        .Height = txtDspStn.Height
        .Width = txtDspStn.Width
        .Top = txtDspStn.Top
        .Left = tbrReview.Buttons("CycleNoTxt").Left + (0.5 * (tbrReview.Buttons("CycleNoTxt").Width - txtDspCycle.Width))
        .FontBold = True
        .Locked = True
    End With
    With txtCycle
        .Height = txtStation.Height
        .Width = tbrReview.Buttons("CycleNoTxt").Width
        .Top = tbrReview.Buttons("CycleNoTxt").Top + (0.6 * tbrReview.Buttons("CycleNoTxt").Height)
        .Left = tbrReview.Buttons("CycleNoTxt").Left
        .FontBold = True
        .FontSize = 9
        .text = "Cycle"
        .Locked = True
    End With

    With cboProcess
        .Width = tbrReview.Buttons("comboProcess").Width
        .Top = txtDspStn.Top
        .Left = tbrReview.Buttons("comboProcess").Left
        .AddItem "Load"
        .AddItem "Purge"
        .AddItem "LeakCheck"
        .ListIndex = 0
    End With
    With txtProcess
        .Height = txtStation.Height
        .Width = cboProcess.Width
        .Top = tbrReview.Buttons("comboProcess").Top + (0.6 * tbrReview.Buttons("comboProcess").Height)
        .Left = cboProcess.Left
        .FontBold = True
        .FontSize = 9
        .text = "Process"
        .Locked = True
    End With

    With txtMsg(1)
        .Height = tbrReview.Buttons("msgbox").Height
        .Width = tbrReview.Buttons("msgbox").Width
'        .Top = txtDspStn.Top
        .Left = tbrReview.Buttons("msgbox").Left
        .FontBold = True
        .Locked = True
    End With

    With txtMsg(2)
        .Height = tbrReview.Buttons("msgbox").Height
        .Width = tbrReview.Buttons("msgbox").Width
'        .Top = txtDspStn.Top
        .Left = tbrReview.Buttons("msgbox").Left
        .FontBold = True
        .Locked = True
    End With

    With txtMsg(3)
        .Height = tbrReview.Buttons("msgbox").Height
        .Width = tbrReview.Buttons("msgbox").Width
'        .Top = txtDspStn.Top
        .Left = tbrReview.Buttons("msgbox").Left
        .FontBold = True
        .Locked = True
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReview = Nothing
End Sub

Private Sub mnuAbout_Click()
    'About
    menuAbout
End Sub

Private Sub mnuAirLog_Click()
    menuViewAirLog
End Sub

Private Sub mnuAk_Client_Click()
    menuAk_Client
End Sub

Private Sub mnuAk_Server_Click()
    menuAk_Server
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

Private Sub tbrReview_ButtonClick(ByVal Button As MSComctlLib.Button)
   ' Use the Key property with the SelectCase statement to specify
   ' an action.
   Select Case Button.Key
       Case Is = "prevcourse"
            CoursePrev
       Case Is = "nextcourse"
            CourseNext
       Case Is = "prevstn"
            StationPrev
       Case Is = "nextstn"
            StationNext
       Case Is = "prevcycle"
            CyclePrev
       Case Is = "nextcycle"
            CycleNext
       Case Is = "searchprev"
            SearchPrev
       Case Is = "searchnext"
            SearchNext
       Case Is = "ok"
            OKClick
       Case Is = "cancel"
            CancelClick
    '   Case Is = "close"
    '        CloseScreen
   End Select
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
       Case Is = "ak_client"
            ' AK Client
            menuAk_Client
       Case Is = "ak_server"
            ' AK Server
            menuAk_Server
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
            ' Close Screen
'            CloseScreen
   End Select
End Sub

Private Sub Form_Activate()
SetErrModule 818, 0
If UseLocalErrorHandler Then On Error GoTo localhandler
    If Not IsNumeric(RvwFirstRecord) Then RvwFirstRecord = 0
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

Private Sub Form_Load()
SetErrModule 818, 1
If UseLocalErrorHandler Then On Error GoTo localhandler
    frmReview.Height = frmMainMenu.Height
    frmReview.Width = frmMainMenu.Width
    ' Set Title Foreground color
    frmResults.ForeColor = Titles_ForeColor
    txtDspStn.ForeColor = TitlesData_Forecolor
    txtDspCourse.ForeColor = TitlesData_Forecolor
    txtDspCycle.ForeColor = TitlesData_Forecolor
    cboProcess.ForeColor = TitlesData_Forecolor
    txtStation.ForeColor = TitlesLabel_ForeColor
    txtCourse.ForeColor = TitlesLabel_ForeColor
    txtCycle.ForeColor = TitlesLabel_ForeColor
    txtProcess.ForeColor = TitlesLabel_ForeColor
    ' Set Data Foreground color
    lblDataBoldTemplate.ForeColor = Data_ForeColor
    lblDataRegTemplate.ForeColor = Data_ForeColor
    lblDataHiLiTemplate.ForeColor = DataHiLite_ForeColor

    RvwLastReadDbFile = " "
    RvwStn = IIf(RvwStn < 1, 1, IIf(RvwStn > NR_STN, NR_STN, RvwStn))
    RvwShift = IIf((Stn_ActiveShift(RvwStn) > 0), Stn_ActiveShift(RvwStn), 1)

    BuildToolbars
    If StationControl(RvwStn, RvwShift).DBFile = "" Then
        ' Test not running
        RvwCourse = 1
        RvwCycle = 1
        RvwMode = StationControl(RvwStn, RvwShift).Mode
        RvwInterval = SysConfig.Load_Interval
        cboProcess.ListIndex = 0
        cboProcess.Refresh
    Else
        ' Test in progress
        RvwCourse = StationControl(RvwStn, RvwShift).Course
        RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
        Select Case StationControl(RvwStn, RvwShift).Mode
            Case VBLOAD
               RvwMode = VBLOAD
               RvwInterval = SysConfig.Load_Interval
               cboProcess.ListIndex = 0
               cboProcess.Refresh
            Case VBPURGE
               RvwMode = VBPURGE
               RvwInterval = SysConfig.Purge_Interval
               cboProcess.ListIndex = 1
               cboProcess.Refresh
            Case VBLEAK
               RvwMode = VBLEAK
               RvwInterval = SysConfig.LeakCheck_Interval
               cboProcess.ListIndex = 2
               cboProcess.Refresh
            Case Else
               RvwMode = StationControl(RvwStn, RvwShift).Mode
               RvwInterval = SysConfig.Load_Interval
               cboProcess.ListIndex = 0
               cboProcess.Refresh
        End Select
    End If
    NewCycle = True
    AtEOF = False
    RvwCharWidth = 144
    RvwFirstRecord = 0
    LstFirstRecord = 0
    RvwRowsMax = 36
    RvwRows = 32
    frmResults.Height = 8610
    frmResults.Top = 1220
    frmResults.Left = 60
    frmResults.Width = frmReview.Width - 220
    RvwMaxRecDsp = 1200
    RvwMaxColNum = 9
    RvwNumHeaderRows = 3
    RvwCurRecCnt = 0
    RvwCurRowIndex = 0
    RvwResponse = vbIgnore
    LastDataDTS = Now()
    LastStatusDTS = Now()
    RvwDosCmd = "0"
    tmrClock.Interval = 350
    tmrClock.Enabled = True
    tbrReview.Buttons("searchprev").ToolTipText = "Previous; DoubleClick to BOF"
    tbrReview.Buttons("searchnext").ToolTipText = "Next; DoubleClick to EOF"
    
    ' Status Bar Setup
    frmReview.pbxBottom.Top = 9885
    pnlAlarms.Left = 0
    pnlAlarms.Width = pnlAlarms.Width - pnlMix.Width
    pnlAlarms.Top = 0
    pnlAlarms.Height = pnlEstop.Height + 150
    pnlMessageFrame.Left = pnlAlarms.Left + pnlAlarms.Width
    pnlMessageFrame.Top = pnlAlarms.Top
    pnlMessageFrame.Height = pnlAlarms.Height
    pnlMessage.Left = 60
    pnlMessage.Top = 60
    pnlMessage.Height = pnlMessageFrame.Height - 120
    pnlMessage.Width = pnlMessageFrame.Width - 120
    pnlPurgeAir.Left = pnlMessageFrame.Left + pnlMessageFrame.Width
    pnlPurgeAir.Width = frmReview.Width - pnlPurgeAir.Left - 150
    pnlPurgeAir.Top = pnlAlarms.Top
    pnlPurgeAir.Height = pnlAlarms.Height
    ' Status Bar Update
    UpdateStatusBars
    
    ' Clear Header Rows
    ClearHeader
    ' Clear All Data Rows
        ' unused data rows
        If RvwRows < RvwRowsMax Then
            For RvwRowIndex = RvwRows To (RvwRowsMax - 1)
               lblDataRowCol00(RvwRowIndex).Caption = " "
               lblDataRowCol01(RvwRowIndex).Caption = " "
               lblDataRowCol02(RvwRowIndex).Caption = " "
               lblDataRowCol03(RvwRowIndex).Caption = " "
               lblDataRowCol04(RvwRowIndex).Caption = " "
               lblDataRowCol05(RvwRowIndex).Caption = " "
               lblDataRowCol06(RvwRowIndex).Caption = " "
               lblDataRowCol07(RvwRowIndex).Caption = " "
               lblDataRowCol08(RvwRowIndex).Caption = " "
            Next RvwRowIndex
        End If
        ' data rows in use
        ClearData
    ' Set screen to Normal Style
    Style_Normal
    ' Update Screen
    UpdateScreen
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

Private Sub CancelClick()
    ' User pressed Cancel
    RvwResponse = vbCancel
    ' Return screen to Normal Style
    Style_Normal
    ' Close Open Data File
    Reader2
End Sub

Private Sub OKClick()
    ' User pressed OK
    RvwResponse = vbOK
    ' Return screen to Normal Style
    Style_Normal
    ' Read(send to file) data from Open Data File
    Reader2
End Sub

Private Sub CloseScreen()
    Unload Me
    Set frmReview = Nothing
End Sub

Private Sub SearchPrev()
SetErrModule 818, 10
If UseLocalErrorHandler Then On Error GoTo localhandler
     Select Case RvwFirstRecord
        Case 0
            ' AtEOF doesnot change
        Case 1 To 30
            If RvwCurRecCnt > RvwRows Then AtEOF = False
        Case Is > 30
            If RvwCurRecCnt > (RvwFirstRecord - 30 + RvwRows) Then AtEOF = False
    End Select
   If NewCycle Then
        NewCycle = False
        RvwFirstRecord = 0
    Else
        ' Double Click to goto beginning of records
        RvwDeltaClick = StationControl(RvwStn, RvwShift).TestTimer - RvwLastClick
        If RvwDeltaClick < 0.25 Then RvwFirstRecord = 0
        RvwLastClick = StationControl(RvwStn, RvwShift).TestTimer
        RvwFirstRecord = IIf(RvwFirstRecord <= 30, 0, RvwFirstRecord - 30)
    End If
    UpdateScreen
    If StationControl(RvwStn, RvwShift).DBFile <> "" Then Reader StationControl(RvwStn, RvwShift).DBFile
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

Private Sub SearchNext()
SetErrModule 818, 11
If UseLocalErrorHandler Then On Error GoTo localhandler
    If NewCycle Then
        NewCycle = False
        RvwFirstRecord = 0
    Else
        ' Double Click to goto end of records
        RvwDeltaClick = StationControl(RvwStn, RvwShift).TestTimer - RvwLastClick
        If RvwDeltaClick < 0.25 Then RvwFirstRecord = RvwCurRecCnt - 10
        RvwLastClick = StationControl(RvwStn, RvwShift).TestTimer
        If Not AtEOF Then RvwFirstRecord = IIf(RvwFirstRecord < 0, 0, RvwFirstRecord + 30)
    End If
    UpdateScreen
    If StationControl(RvwStn, RvwShift).DBFile <> "" Then Reader StationControl(RvwStn, RvwShift).DBFile
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

Private Sub CyclePrev()
SetErrModule 818, 12
If UseLocalErrorHandler Then On Error GoTo localhandler
    If (StationControl(RvwStn, RvwShift).Mode = VBIDLE) Then
        RvwCycle = 1
        UpdateScreen
    Else
        ChgSource
        RvwCycle = IIf(RvwCycle <= 1, StationControl(RvwStn, RvwShift).CurrCycle, RvwCycle - 1)
        UpdateScreen
    End If
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

Private Sub CycleNext()
SetErrModule 818, 13
If UseLocalErrorHandler Then On Error GoTo localhandler
    If (StationControl(RvwStn, RvwShift).Mode = VBIDLE) Then
        RvwCycle = 1
        UpdateScreen
    Else
        ChgSource
        RvwCycle = IIf(RvwCycle >= StationControl(RvwStn, RvwShift).CurrCycle, 1, RvwCycle + 1)
        UpdateScreen
    End If
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

Private Sub StationPrev()
SetErrModule 818, 14
If UseLocalErrorHandler Then On Error GoTo localhandler
    If (LAST_STN > 1) Then
        RvwStn = IIf(RvwStn <= 1, LAST_STN, RvwStn - 1)
        RvwShift = IIf((Stn_ActiveShift(RvwStn) > 0), Stn_ActiveShift(RvwStn), 1)

        If StationControl(RvwStn, RvwShift).DBFile = "" Then
            ' Test not running
            RvwCourse = 1
            RvwCycle = 1
        Else
            ' Test in progress
            RvwCourse = StationControl(RvwStn, RvwShift).Course
            Select Case StationControl(RvwStn, RvwShift).Mode
             Case VBLOAD
                RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
                If RvwMode <> VBLOAD Then
                    RvwMode = VBLOAD
                    RvwInterval = SysConfig.Load_Interval
                    cboProcess.ListIndex = 0
                    cboProcess.Refresh
                End If
             Case VBPURGE
                RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
                If RvwMode <> VBPURGE Then
                    RvwMode = VBPURGE
                    RvwInterval = SysConfig.Purge_Interval
                    cboProcess.ListIndex = 1
                    cboProcess.Refresh
                End If
             Case VBLEAK
                RvwCycle = 0
                If RvwMode <> VBLEAK Then
                    RvwMode = VBLEAK
                    RvwInterval = SysConfig.LeakCheck_Interval
                    cboProcess.ListIndex = 2
                    cboProcess.Refresh
                End If
            End Select
        End If
        ChgSource
        UpdateScreen
    End If
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

Private Sub StationNext()
SetErrModule 818, 15
If UseLocalErrorHandler Then On Error GoTo localhandler
    If (LAST_STN > 1) Then
        RvwStn = IIf(RvwStn = LAST_STN, 1, RvwStn + 1)
        RvwShift = IIf((Stn_ActiveShift(RvwStn) > 0), Stn_ActiveShift(RvwStn), 1)
        If StationControl(RvwStn, RvwShift).DBFile = "" Then
            ' Test not running
            RvwCourse = 1
            RvwCycle = 1
        Else
            ' Test in progress
            RvwCourse = StationControl(RvwStn, RvwShift).Course
            Select Case StationControl(RvwStn, RvwShift).Mode
             Case VBLOAD
                RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
                If RvwMode <> VBLOAD Then
                    RvwMode = VBLOAD
                    RvwInterval = SysConfig.Load_Interval
                    cboProcess.ListIndex = 0
                    cboProcess.Refresh
                End If
             Case VBPURGE
                RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
                If RvwMode <> VBPURGE Then
                    RvwMode = VBPURGE
                    RvwInterval = SysConfig.Purge_Interval
                    cboProcess.ListIndex = 1
                    cboProcess.Refresh
                End If
             Case VBLEAK
                RvwCycle = 0
                If RvwMode <> VBLEAK Then
                    RvwMode = VBLEAK
                    RvwInterval = SysConfig.LeakCheck_Interval
                    cboProcess.ListIndex = 2
                    cboProcess.Refresh
                End If
            End Select
        End If
        ChgSource
        UpdateScreen
    End If
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

Private Sub CoursePrev()
SetErrModule 818, 1441
If UseLocalErrorHandler Then On Error GoTo localhandler
    If (StationControl(RvwStn, RvwShift).Mode = VBIDLE) Then
        RvwCourse = 1
        UpdateScreen
    Else
        RvwCourse = IIf(RvwCourse <= 1, StationSequence(RvwStn, RvwShift).NumCourses, RvwCourse - 1)
        ChgSource
        UpdateScreen
    End If
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

Private Sub CourseNext()
SetErrModule 818, 1551
If UseLocalErrorHandler Then On Error GoTo localhandler
    If (StationControl(RvwStn, RvwShift).Mode = VBIDLE) Then
        RvwCourse = 1
        UpdateScreen
    Else
        RvwCourse = IIf(RvwCourse >= StationSequence(RvwStn, RvwShift).NumCourses, 1, RvwCourse + 1)
        ChgSource
        UpdateScreen
    End If
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

Private Sub cboProcess_Click()
SetErrModule 818, 16
If UseLocalErrorHandler Then On Error GoTo localhandler
    Select Case cboProcess.ListIndex
     Case 0
        RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
        If RvwMode <> VBLOAD Then
            RvwMode = VBLOAD
            RvwInterval = SysConfig.Load_Interval
        End If
     Case 1
        RvwCycle = StationControl(RvwStn, RvwShift).CurrCycle
        If RvwMode <> VBPURGE Then
            RvwMode = VBPURGE
            RvwInterval = SysConfig.Purge_Interval
        End If
     Case 2
        RvwCycle = 0
        If RvwMode <> VBLEAK Then
            RvwMode = VBLEAK
            RvwInterval = SysConfig.LeakCheck_Interval
        End If
    End Select
    ChgSource
    UpdateScreen
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

Private Sub ChgSource()
SetErrModule 818, 20
If UseLocalErrorHandler Then On Error GoTo localhandler
    NewCycle = True
    AtEOF = False
    RvwResponse = vbIgnore
    tbrReview.Buttons("searchprev").ToolTipText = "Previous"
    tbrReview.Buttons("searchnext").ToolTipText = "Next"
    ' reset new data flag
    StationControl(RvwStn, RvwShift).NewDataInDB = False
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

Private Sub UpdateScreen()
SetErrModule 818, 21
If UseLocalErrorHandler Then On Error GoTo localhandler
    txtDspStn.text = STN_INFO(RvwStn).Abrev
    txtDspShift.text = Format(RvwShift, "0")
    txtDspCourse.text = Format(RvwCourse, "0")
    If StationControl(RvwStn, RvwShift).DBFile = "" Then
        ' Test not running
        ClearHeader
    Else
        ' Test in progress
        UpdateHeader
    End If
    Select Case StationRecipe(RvwStn, RvwShift).EndMethod
        Case ENDCYCLES
            RvwCycle = IIf(RvwCycle > StationRecipe(RvwStn, RvwShift).Cycles, StationRecipe(RvwStn, RvwShift).Cycles, RvwCycle)
        Case ENDWEIGHTCHG
            RvwCycle = RvwCycle
        Case Else
            RvwCycle = IIf(RvwCycle > StationRecipe(RvwStn, RvwShift).Cycles, StationRecipe(RvwStn, RvwShift).Cycles, RvwCycle)
    End Select
    txtDspCycle.text = Format(RvwCycle, "##0")
    ClearData
    UpdateStatus
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

Private Sub UpdateStatus()
SetErrModule 818, 22
If UseLocalErrorHandler Then On Error GoTo localhandler
    LastStatusDTS = Now()
    If RvwStyle = "Normal" Then
        If StationControl(RvwStn, RvwShift).DBFile = "" Then
            NewMessage Black, " ", "No Open DB File", " "
        Else
            RvwMsg(1) = "Using DB File:  "
            RvwMsg(1) = RvwMsg(1) & Mid(StationControl(RvwStn, RvwShift).DBFile, (Len(StationControl(RvwStn, RvwShift).DBFile) - 10), 7)
            If AtEOF And Not NewCycle Then
                RvwMsg(2) = "at End of Cycle " & Format(RvwCycle, "##")
                If RvwMode = VBLOAD Then RvwMsg(2) = RvwMsg(2) & " Load Data"
                If RvwMode = VBPURGE Then RvwMsg(2) = RvwMsg(2) & " Purge Data"
                If RvwMode = VBLEAK Then RvwMsg(2) = RvwMsg(2) & " Leak Check Data"
            Else
                RvwMsg(2) = " "
            End If
            NewMessage Message_ForeColor, " ", RvwMsg(1), RvwMsg(2)
        End If
    End If
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

Private Sub UpdateHeader()
Dim iCol, alignCode As Integer
SetErrModule 818, 30
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' Setup Data Column Information Array
    SetupColData
    
    ' Align Header Text
    For iCol = 0 To 8
        Select Case RvwColData(iCol).ColAlign
            Case "LEFT"
                alignCode = 0
            Case "CENTER"
                alignCode = 2
            Case "RIGHT"
                alignCode = 1
        End Select
        Select Case iCol
            Case 0
                lblHdrRowCol00(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol00(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol00(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol00(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol00(0).Alignment = alignCode
                lblHdrRowCol00(1).Alignment = alignCode
                lblHdrRowCol00(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol00(1).ForeColor = ModeBackColor(RvwMode)
            Case 1
                lblHdrRowCol01(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol01(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol01(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol01(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol01(0).Alignment = alignCode
                lblHdrRowCol01(1).Alignment = alignCode
                lblHdrRowCol01(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol01(1).ForeColor = ModeBackColor(RvwMode)
            Case 2
                lblHdrRowCol02(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol02(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol02(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol02(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol02(0).Alignment = alignCode
                lblHdrRowCol02(1).Alignment = alignCode
                lblHdrRowCol02(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol02(1).ForeColor = ModeBackColor(RvwMode)
            Case 3
                lblHdrRowCol03(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol03(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol03(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol03(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol03(0).Alignment = alignCode
                lblHdrRowCol03(1).Alignment = alignCode
                lblHdrRowCol03(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol03(1).ForeColor = ModeBackColor(RvwMode)
            Case 4
                lblHdrRowCol04(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol04(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol04(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol04(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol04(0).Alignment = alignCode
                lblHdrRowCol04(1).Alignment = alignCode
                lblHdrRowCol04(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol04(1).ForeColor = ModeBackColor(RvwMode)
            Case 5
                lblHdrRowCol05(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol05(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol05(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol05(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol05(0).Alignment = alignCode
                lblHdrRowCol05(1).Alignment = alignCode
                lblHdrRowCol05(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol05(1).ForeColor = ModeBackColor(RvwMode)
            Case 6
                lblHdrRowCol06(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol06(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol06(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol06(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol06(0).Alignment = alignCode
                lblHdrRowCol06(1).Alignment = alignCode
                lblHdrRowCol06(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol06(1).ForeColor = ModeBackColor(RvwMode)
            Case 7
                lblHdrRowCol07(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol07(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol07(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol07(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol07(0).Alignment = alignCode
                lblHdrRowCol07(1).Alignment = alignCode
                lblHdrRowCol07(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol07(1).ForeColor = ModeBackColor(RvwMode)
            Case 8
                lblHdrRowCol08(0).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol08(1).Left = RvwColData(iCol).ColLeft
                lblHdrRowCol08(0).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol08(1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                lblHdrRowCol08(0).Alignment = alignCode
                lblHdrRowCol08(1).Alignment = alignCode
                lblHdrRowCol08(0).ForeColor = ModeBackColor(RvwMode)
                lblHdrRowCol08(1).ForeColor = ModeBackColor(RvwMode)
        End Select
    Next iCol
    
    ' Display Header Text
        ' only display 'Next' Header if the next Data Write Time will be displayed
    lblHdrRowCol00(0).Caption = IIf(RvwInterval > 2, RvwColData(0).Header1, " ")
    lblHdrRowCol01(0).Caption = RvwColData(1).Header1
    lblHdrRowCol02(0).Caption = RvwColData(2).Header1
    lblHdrRowCol03(0).Caption = RvwColData(3).Header1
    lblHdrRowCol04(0).Caption = RvwColData(4).Header1
    lblHdrRowCol05(0).Caption = RvwColData(5).Header1
    lblHdrRowCol06(0).Caption = RvwColData(6).Header1
    lblHdrRowCol07(0).Caption = RvwColData(7).Header1
    lblHdrRowCol08(0).Caption = RvwColData(8).Header1
    lblHdrRowCol00(1).Caption = RvwColData(0).Header3
    lblHdrRowCol01(1).Caption = RvwColData(1).Header3
    lblHdrRowCol02(1).Caption = RvwColData(2).Header3
    lblHdrRowCol03(1).Caption = RvwColData(3).Header3
    lblHdrRowCol04(1).Caption = RvwColData(4).Header3
    lblHdrRowCol05(1).Caption = RvwColData(5).Header3
    lblHdrRowCol06(1).Caption = RvwColData(6).Header3
    lblHdrRowCol07(1).Caption = RvwColData(7).Header3
    lblHdrRowCol08(1).Caption = RvwColData(8).Header3
    
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
Private Sub Style_Normal()
SetErrModule 818, 40
If UseLocalErrorHandler Then On Error GoTo localhandler
    RvwStyle = "Normal"
    tbrReview.Buttons("ok").Enabled = False
    tbrReview.Buttons("cancel").Enabled = False
    tbrReview.Buttons("ok").Visible = False
    tbrReview.Buttons("cancel").Visible = False
    txtDspStn.Enabled = True
    txtDspShift.Enabled = True
    txtDspCycle.Enabled = True
    cboProcess.Enabled = True
    tbrReview.Buttons("searchprev").Enabled = True
    tbrReview.Buttons("searchprev").ToolTipText = "Previous"
    tbrReview.Buttons("searchnext").Enabled = True
    tbrReview.Buttons("searchnext").ToolTipText = "Next"
'    tbrReview.Buttons("close").Enabled = True
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

Private Sub Style_WaitingAnswer()
SetErrModule 818, 41
If UseLocalErrorHandler Then On Error GoTo localhandler
    RvwStyle = "WaitingAnswer"
    tbrReview.Buttons("ok").Enabled = True
    tbrReview.Buttons("cancel").Enabled = True
    tbrReview.Buttons("ok").Visible = True
    tbrReview.Buttons("cancel").Visible = True
    txtDspStn.Enabled = False
    txtDspShift.Enabled = False
    txtDspCycle.Enabled = False
    cboProcess.Enabled = False
    tbrReview.Buttons("searchprev").Enabled = False
    tbrReview.Buttons("searchnext").Enabled = False
'    tbrReview.Buttons("close").Enabled = False
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

Private Sub Reader(ByVal datafile As String)
' Module Name:  Reader
' Description:  This routine opens the  'Data' recordset
'               of the specified datafile.
'
Dim sOrder As String
SetErrModule 818, 100
If UseLocalErrorHandler Then On Error GoTo localhandler

' Check to see if the Job is Still Active
If RvwMode = VBLOAD _
    Or RvwMode = VBPURGE _
    Or RvwMode = VBLEAK Then
    
    If RvwMode = VBLEAK Then RvwCycle = 0

    ' Open Database
    Set dbDbase = OpenDatabase(datafile)
        
    
    ' Open dynaset style recordset
    sOrder = "ORDER BY [TestTime] ASC"
    If RvwMode = VBLEAK Then sOrder = "ORDER BY [TestTime] ASC, [ReportCode] ASC"
    RvwCriteria = "SELECT * FROM [Data] WHERE ([Course] = " & Format(RvwCourse, "##0") & _
                   " AND [Cycle] = " & Format(RvwCycle, "##0") & _
                   " AND ([Data].[Mode] = " & Format(RvwMode, "#0") & ")) " & sOrder
    
    Set rsRecord = dbDbase.OpenRecordset(RvwCriteria, dbOpenDynaset)
    RvwCurRecCnt = rsRecord.RecordCount
    If rsRecord.RecordCount > 0 Then
    
        If rsRecord.RecordCount > RvwMaxRecDsp Then
            ' Too many records for scrolling
            RvwMsg(1) = "Too Many Records (" & Format(rsRecord.RecordCount, "######0") & ") !"
            RvwMsg(2) = "Data can be sent to a Text File"
            RvwMsg(3) = "Notepad will then open the file."
            NewMessage Message_ForeColor, RvwMsg(1), RvwMsg(2), RvwMsg(3)
            ' Clear Data Rows
            ClearData
            ' Clear Header
            ClearHeader
            ' Set screen to Waiting for Answer
            Style_WaitingAnswer
            ' Note: After user answers OK or Cancel, Reader2 is called)
            
        Else
            ' Read the data (and display it)
            If rsRecord.BOF = False Then  ' Do if mode and cycle match
                rsRecord.MoveFirst
                ' Display the Data
                DisplayData
            End If
          
        End If
        
    Else
        NewMessage Black, " ", "No Data to Display", " "
    End If

    RvwLastReadDbFile = datafile

End If

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

Private Sub Reader2()
'
SetErrModule 818, 102
If UseLocalErrorHandler Then On Error GoTo localhandler
    Select Case RvwResponse
        Case vbOK
            ' OK to Proceed
            NewMessage Message_ForeColor, " ", "Sending Data to File", " "
            ' Read the Data (& send it to a temp. file)
            If rsRecord.BOF = False Then  ' Do if mode and cycle match
               rsRecord.MoveFirst
               ' Send the Data to File
               SendData
            End If
            ' Open File in Notepad
            RvwDosCmd = FILEPATH_reports & "TestData.txt"
            RvwDosCmd = "notepad " & RvwDosCmd
            ' Shell to DOS
            Shell RvwDosCmd
            NewMessage Message_ForeColor, " ", "Data sent to Notepad", " "
            
       Case vbCancel
            ' Review Canceled
            NewMessage Black, " ", "Review Canceled", " "
            
    End Select
    rsRecord.Close
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

Private Sub DisplayData()
'
' Module Name:  DisplayData
' Description:  This routine reads the report data values
'               from the 'Data' recordset
'               of the database specified as dBase
'
'
Dim iCol As Integer
Dim alignCode As Integer
Dim dataVal As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 818, 210
    
    tbrReview.Buttons("searchprev").ToolTipText = "Previous; DoubleClick to BOF"
    tbrReview.Buttons("searchnext").ToolTipText = "Next; DoubleClick to EOF"
    
    ' Don't let the FirstRecordToDisplay pointer exceed the record count
    If RvwCurRecCnt < RvwFirstRecord Then
        If RvwFirstRecord > 5 Then RvwFirstRecord = (RvwCurRecCnt - 5)
'    ElseIf RvwCurRecCnt = (RvwFirstRecord + RvwRows + 1) Then
    ElseIf RvwCurRecCnt > (RvwFirstRecord + RvwRows) And RvwCurRecCnt < (RvwFirstRecord + RvwRows + 5) Then
        RvwFirstRecord = (RvwFirstRecord + 20)
        AtEOF = False
    End If
    
    ' Clear Data Rows (except refresh only)
    If Not StationControl(RvwStn, RvwShift).NewDataInDB Or LstFirstRecord <> RvwFirstRecord Then ClearData
    ' Remember the First Record index
    LstFirstRecord = RvwFirstRecord

    ' Write data lines
    RvwRecordIndex = 0
    RvwRowIndex = 0
    Do While rsRecord.EOF = False
        LastDataDTS = rsRecord("Time")
        If (RvwRecordIndex >= RvwFirstRecord) And (RvwRecordIndex < (RvwFirstRecord + RvwRows)) Then
            RvwCurRowIndex = RvwRowIndex
            If rsRecord("ReportCode") = OOTPAUSEBEGIN Or rsRecord("ReportCode") = OOTPAUSECLEAR Then
            
                ' OOT Start/End Data Line
                If rsRecord("ReportCode") = OOTPAUSEBEGIN Then
                    lblDataRowCol01(RvwRowIndex).Caption = Format(rsRecord("Time"), RvwColData(0).DataFormat)
                    lblDataRowCol02(RvwRowIndex).Caption = "********"
                    lblDataRowCol03(RvwRowIndex).Caption = "OOT Alarm"
                    lblDataRowCol04(RvwRowIndex).Caption = "********"
                Else
                    lblDataRowCol01(RvwRowIndex).Caption = Format(rsRecord("Time"), RvwColData(0).DataFormat)
                    lblDataRowCol02(RvwRowIndex).Caption = "Operator"
                    lblDataRowCol03(RvwRowIndex).Caption = "Cleared"
                    lblDataRowCol04(RvwRowIndex).Caption = "OOT Alarm"
                End If
            
            Else
            
                ' Normal Data Line     (not OOT start or OOT end)
                For iCol = 0 To RvwMaxColNum
                    If RvwColData(iCol).InUse Then
                    
                        ' Get Data Value as a String
                        Select Case RvwMode
                            Case VBLOAD
                                Select Case RvwColData(iCol).DataName
                                    Case "Time"
                                        dataVal = Format(rsRecord("Time"), RvwColData(iCol).DataFormat)
                                    Case "NitFlow"
                                        dataVal = Format(rsRecord("NitFlow"), RvwColData(iCol).DataFormat)
                                    Case "Mix"
                                        dataVal = Format(rsRecord("Mix"), RvwColData(iCol).DataFormat)
                                    Case "BtnFlow"
                                        dataVal = Format(rsRecord("BtnFlow"), RvwColData(iCol).DataFormat)
                                    Case "LoadRate"
                                        dataVal = Format(rsRecord("LoadRate"), RvwColData(iCol).DataFormat)
                                    Case "LoadGrams"
                                        dataVal = Format(rsRecord("LoadTotalGrams"), RvwColData(iCol).DataFormat)
                                    Case "BtnFlow"
                                        dataVal = Format(rsRecord("BtnFlow"), RvwColData(iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(rsRecord("priscale"), RvwColData(iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(rsRecord("auxscale"), RvwColData(iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(rsRecord("TotalWtChg"), RvwColData(iCol).DataFormat)
                                    Case "WtChgRate"
                                        dataVal = Format(rsRecord("TotalWtChgRate"), RvwColData(iCol).DataFormat)
                                    Case "LiveFuelCycles"
                                        dataVal = Format(rsRecord("LiveFuelCycles"), RvwColData(iCol).DataFormat)
                                    Case "LiveFuelTemp"
                                        dataVal = Format(rsRecord("LiveFuelTemp"), RvwColData(iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(rsRecord("TestTime"), RvwColData(iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            Case VBPURGE
                                Select Case RvwColData(iCol).DataName
                                    Case "Time"
                                        dataVal = Format(rsRecord("Time"), RvwColData(iCol).DataFormat)
                                    Case "PurgeFlow"
                                        dataVal = Format(rsRecord("PurgeFlow"), RvwColData(iCol).DataFormat)
                                    Case "PurgeTemp"
                                        dataVal = Format(rsRecord("PATemp"), RvwColData(iCol).DataFormat)
                                    Case "PurgeMoist"
                                        dataVal = Format(rsRecord("Moisture"), RvwColData(iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(rsRecord("priscale"), RvwColData(iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(rsRecord("auxscale"), RvwColData(iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(rsRecord("TotalWtChg"), RvwColData(iCol).DataFormat)
                                    Case "WtChgRate"
                                        dataVal = Format(rsRecord("TotalWtChgRate"), RvwColData(iCol).DataFormat)
                                    Case "PurgeVol"
                                        dataVal = Format(rsRecord("PurgeVol"), RvwColData(iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(rsRecord("TestTime"), RvwColData(iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            Case VBLEAK
                                Select Case RvwColData(iCol).DataName
                                    Case "Time"
                                        dataVal = Format(rsRecord("Time"), RvwColData(iCol).DataFormat)
                                    Case "Pressure"
                                        If rsRecord("Pressure") < 0 Then
                                            ' show negative values as = zero
                                            dataVal = Format(CSng(0), RvwColData(iCol).DataFormat)
                                        Else
                                            dataVal = Format(rsRecord("Pressure"), RvwColData(iCol).DataFormat)
                                        End If
                                    Case "Comment"
                                        Select Case CInt(rsRecord("ReportCode"))
                                            Case LCBEGINPHASE0
                                                dataVal = "Leak Check - Begin Purging"
                                            Case LCBEGINPHASE1
                                                dataVal = "Leak Check - Begin Pressurize"
                                            Case LCBEGINPHASE2
                                                dataVal = "Leak Check - Begin Testing"
                                            Case LCTESTRESULT
                                                Select Case CInt(rsRecord("LeakCheckResult"))
                                                    Case 0 To 9
                                                        ' valid result codes
                                                        dataVal = LeakResultDesc(CInt(rsRecord("LeakCheckResult")))
                                                        If dataVal = "undefined" Then dataVal = "Leak Check - Undefined Result" & CInt(rsRecord("LeakCheckResult"))
                                                    Case Else
                                                        dataVal = "Leak Check - Unknown Result" & CInt(rsRecord("LeakCheckResult"))
                                                End Select
                                            Case LCOPERCONTINUE
                                                dataVal = "Operator CONTINUE after a Leak Check Failure"
                                            Case LCAUTOCONTINUE
                                                dataVal = "Automatic Continue after a Leak Check Failure"
                                            Case Else
                                                ' normal writes
                                                dataVal = " "
                                        End Select
                                    Case "TestTime"
                                        dataVal = Format(rsRecord("TestTime"), RvwColData(iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            Case Else
                                dataVal = " "
                        End Select
                        ' Get Alignment Code for this Col
                        Select Case RvwColData(iCol).ColAlign
                            Case "LEFT"
                                alignCode = 0
                            Case "CENTER"
                                alignCode = 2
                            Case "RIGHT"
                                alignCode = 1
                        End Select
                    
                        ' Display(& align) the Data Value
                        Select Case iCol
                            Case 0
                                lblDataRowCol00(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol00(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol00(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol00(RvwRowIndex).FontItalic = False
                                lblDataRowCol00(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol00(RvwRowIndex).Caption = dataVal
                            Case 1
                                lblDataRowCol01(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol01(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol01(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol01(RvwRowIndex).FontItalic = False
                                lblDataRowCol01(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol01(RvwRowIndex).Caption = dataVal
                            Case 2
                                lblDataRowCol02(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol02(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol02(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol02(RvwRowIndex).FontItalic = False
                                lblDataRowCol02(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol02(RvwRowIndex).Caption = dataVal
                            Case 3
                                lblDataRowCol03(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol03(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol03(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol03(RvwRowIndex).FontItalic = False
                                lblDataRowCol03(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol03(RvwRowIndex).Caption = dataVal
                            Case 4
                                lblDataRowCol04(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol04(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol04(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol04(RvwRowIndex).FontItalic = False
                                lblDataRowCol04(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol04(RvwRowIndex).Caption = dataVal
                            Case 5
                                lblDataRowCol05(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol05(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol05(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol05(RvwRowIndex).FontItalic = False
                                lblDataRowCol05(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol05(RvwRowIndex).Caption = dataVal
                            Case 6
                                lblDataRowCol06(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol06(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol06(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol06(RvwRowIndex).FontItalic = False
                                lblDataRowCol06(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol06(RvwRowIndex).Caption = dataVal
                            Case 7
                                lblDataRowCol07(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol07(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol07(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol07(RvwRowIndex).FontItalic = False
                                lblDataRowCol07(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol07(RvwRowIndex).Caption = dataVal
                            Case 8
                                lblDataRowCol08(RvwRowIndex).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol08(RvwRowIndex).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol08(RvwRowIndex).Alignment = alignCode
                                lblDataRowCol08(RvwRowIndex).FontItalic = False
                                lblDataRowCol08(RvwRowIndex).ForeColor = lblDataBoldTemplate.ForeColor
                                lblDataRowCol08(RvwRowIndex).Caption = dataVal
                            Case Else
                                ' only nine data columns on display
                        End Select
                        
                    End If
                Next iCol
            
            End If
        
            RvwRowIndex = RvwRowIndex + 1
                                                
        End If
        
        RvwRecordIndex = RvwRecordIndex + 1
        rsRecord.MoveNext
    Loop
    
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

Private Sub SendData()
'
' Module Name:  SendData
' Description:  This routine sends the report data
'               to a text file.
'
'
Dim iCol, iRow As Integer
Dim ColWidth, ItemWidth, lpad, rpad As Integer
Dim dataVal As String


If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 818, 310
    
    ' Setup Data Column Information Array
'    SetupColData (don't print the "Next" column)
    
    ' open Text File
    RvwFileName = FILEPATH_reports & "TestData.txt"
    RvwFileNumber = FreeFile
    Open RvwFileName For Output As #RvwFileNumber
    
    File_Line RvwFileNumber, " "
    File_Line RvwFileNumber, " "
    If RvwMode = VBLOAD Then File_Line RvwFileNumber, "   LOAD CYCLE " & Format(RvwCycle, "##0")
    If RvwMode = VBPURGE Then File_Line RvwFileNumber, "   PURGE CYCLE " & Format(RvwCycle, "##0")
    File_Line RvwFileNumber, " "
    File_Line RvwFileNumber, "   STATION " & Format(RvwStn, "0")
    File_Line RvwFileNumber, " "
    File_Line RvwFileNumber, "   SHIFT " & Format(RvwShift, "0")
    File_Line RvwFileNumber, " "
    File_Line RvwFileNumber, " "
    
    
    ' Build Header Rows
    For iRow = 1 To RvwNumHeaderRows
        pString = ""
        For iCol = 1 To RvwMaxColNum
            If RvwColData(iCol).InUse Then
            
                ' Get Header String for this Row/Column
                Select Case iRow
                    Case 1
                        dataVal = RvwColData(iCol).Header1
                        '   (note: scale # is appended to Header2)
                        If RvwColData(iCol).DataName = "PriScale" Then dataVal = "Primary"
                        If RvwColData(iCol).DataName = "AuxScale" Then dataVal = "Aux"
                    Case 2
                        dataVal = RvwColData(iCol).Header2
                    Case 3
                        dataVal = RvwColData(iCol).Header3
                End Select
                ' Width of this column
                ColWidth = RvwColData(iCol).ColWidth
                ' Add this Column to the Header Row
                Select Case RvwColData(iCol).ColAlign
                    Case "LEFT"
                        lpad = 0
                        rpad = ColWidth - Len(dataVal)
                    Case "CENTER"
                        rpad = Int((ColWidth - Len(dataVal)) / 2)
                        lpad = ColWidth - (Len(dataVal) + rpad)
                    Case "RIGHT"
                        lpad = ColWidth - Len(dataVal)
                        rpad = 0
                End Select
                pString = pString + Space(lpad) + dataVal + Space(rpad)
                                
            End If
        Next iCol
        ' Send the Header Row to the File
        File_Line RvwFileNumber, pString
    Next iRow
    
       
    ' Write data lines
    Do While rsRecord.EOF = False
        If rsRecord("ReportCode") = OOTPAUSEBEGIN Or rsRecord("ReportCode") = OOTPAUSECLEAR Then
        
            ' OOT Start/End Data Line
            If rsRecord("ReportCode") = OOTPAUSEBEGIN Then
                pString = Format(rsRecord("Time"), "HH:MM:SS ") & "***************** OOT Alarm ******************"
            Else
                pString = Format(rsRecord("Time"), "HH:MM:SS ") & "********* Operator Cleared OOT Alarm *********"
            End If
            ' Print the String
            File_Line RvwFileNumber, pString
        
        Else
        
            ' Normal Data Line     (not OOT start or OOT end)
   
            pString = ""
            For iCol = 1 To RvwMaxColNum
                If RvwColData(iCol).InUse Then
                
                    ' Get Data Value as a String
                    Select Case RvwMode
                        Case VBLOAD
                            Select Case RvwColData(iCol).DataName
                                Case "Time"
                                    dataVal = Format(rsRecord("Time"), RvwColData(iCol).DataFormat)
                                Case "NitFlow"
                                    dataVal = Format(rsRecord("NitFlow"), RvwColData(iCol).DataFormat)
                                Case "Mix"
                                    dataVal = Format(rsRecord("Mix"), RvwColData(iCol).DataFormat)
                                Case "BtnFlow"
                                    dataVal = Format(rsRecord("BtnFlow"), RvwColData(iCol).DataFormat)
                                Case "LoadRate"
                                    dataVal = Format(rsRecord("LoadRate"), RvwColData(iCol).DataFormat)
                                Case "LoadGrams"
                                    dataVal = Format(rsRecord("LoadTotalGrams"), RvwColData(iCol).DataFormat)
                                Case "PriScale"
                                    dataVal = Format(rsRecord("PriScale"), RvwColData(iCol).DataFormat)
                                Case "AuxScale"
                                    dataVal = Format(rsRecord("AuxScale"), RvwColData(iCol).DataFormat)
                                Case "WtChange"
                                    dataVal = Format(rsRecord("TotalWtChg"), RvwColData(iCol).DataFormat)
                                Case "LiveFuelCycles"
                                    dataVal = Format(rsRecord("LiveFuelCycles"), RvwColData(iCol).DataFormat)
                                Case "LiveFuelTemp"
                                    dataVal = Format(rsRecord("LiveFuelTemp"), RvwColData(iCol).DataFormat)
                                Case "TestTime"
                                    dataVal = Format(rsRecord("TestTime"), RvwColData(iCol).DataFormat)
                                Case Else
                                    dataVal = "unknown"
                            End Select
                        Case VBPURGE
                            Select Case RvwColData(iCol).DataName
                                Case "Time"
                                    dataVal = Format(rsRecord("Time"), RvwColData(iCol).DataFormat)
                                Case "PurgeFlow"
                                    dataVal = Format(rsRecord("PurgeFlow"), RvwColData(iCol).DataFormat)
                                Case "PurgeTemp"
                                    dataVal = Format(rsRecord("PATemp"), RvwColData(iCol).DataFormat)
                                Case "PurgeMoist"
                                    dataVal = Format(rsRecord("Moisture"), RvwColData(iCol).DataFormat)
                                Case "PriScale"
                                    dataVal = Format(rsRecord("PriScale"), RvwColData(iCol).DataFormat)
                                Case "AuxScale"
                                    dataVal = Format(rsRecord("AuxScale"), RvwColData(iCol).DataFormat)
                                Case "WtChange"
                                    dataVal = Format(rsRecord("TotalWtChg"), RvwColData(iCol).DataFormat)
                                Case "PurgeVol"
                                    dataVal = Format(rsRecord("PurgeVol"), RvwColData(iCol).DataFormat)
                                Case "TestTime"
                                    dataVal = Format(rsRecord("TestTime"), RvwColData(iCol).DataFormat)
                                Case Else
                                    dataVal = "unknown"
                            End Select
                        Case VBLEAK
                            Select Case RvwColData(iCol).DataName
                                Case "Time"
                                    dataVal = Format(rsRecord("Time"), RvwColData(iCol).DataFormat)
                                Case "Pressure"
                                    If rsRecord("Pressure") < 0 Then
                                        ' show negative values as = zero
                                        dataVal = Format(CSng(0), RvwColData(iCol).DataFormat)
                                    Else
                                        dataVal = Format(rsRecord("Pressure"), RvwColData(iCol).DataFormat)
                                    End If
                                Case "Comment"
                                    Select Case CInt(rsRecord("ReportCode"))
                                        Case LCBEGINPHASE0
                                            dataVal = "Leak Check - Begin Purging"
                                        Case LCBEGINPHASE1
                                            dataVal = "Leak Check - Begin Pressurize"
                                        Case LCBEGINPHASE2
                                            dataVal = "Leak Check - Begin Testing"
                                        Case LCTESTRESULT
                                            Select Case CInt(rsRecord("LeakCheckResult"))
                                                Case 0 To 9
                                                    ' valid result codes
                                                    dataVal = LeakResultDesc(CInt(rsRecord("LeakCheckResult")))
                                                    If dataVal = "undefined" Then dataVal = "Leak Check - Undefined Result" & CInt(rsRecord("LeakCheckResult"))
                                                Case Else
                                                    dataVal = "Leak Check - Unknown Result" & CInt(rsRecord("LeakCheckResult"))
                                            End Select
                                        Case LCOPERCONTINUE
                                            dataVal = "Operator CONTINUE after a Leak Check Failure"
                                        Case LCAUTOCONTINUE
                                            dataVal = "Automatic Continue after a Leak Check Failure"
                                        Case Else
                                            ' normal writes
                                            dataVal = " "
                                    End Select
                                Case "TestTime"
                                    dataVal = Format(rsRecord("TestTime"), RvwColData(iCol).DataFormat)
                                Case Else
                                    dataVal = " "
                            End Select
                        Case Else
                            dataVal = "unknown"
                    End Select
                    ' Width of this column
                    ColWidth = RvwColData(iCol).ColWidth
                    ' Align this Column
                    Select Case RvwColData(iCol).ColAlign
                        Case "LEFT"
                            lpad = 0
                            rpad = ColWidth - Len(dataVal)
                        Case "CENTER"
                            rpad = Int((ColWidth - Len(dataVal)) / 2)
                            lpad = ColWidth - (Len(dataVal) + rpad)
                        Case "RIGHT"
                            lpad = ColWidth - Len(dataVal)
                            rpad = 0
                    End Select
                    pString = pString + Space(lpad) + dataVal + Space(rpad)
                             
                End If
            Next iCol
            ' Send the Data Row to the File
            File_Line RvwFileNumber, pString
            
        End If
        rsRecord.MoveNext
    Loop
    
    Close #RvwFileNumber


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

Private Sub ClearData()
SetErrModule 818, 60
If UseLocalErrorHandler Then On Error GoTo localhandler
    ' Clear Data Rows
    For RvwRowIndex = 0 To (RvwRows - 1)
       lblDataRowCol00(RvwRowIndex).Caption = " "
       lblDataRowCol01(RvwRowIndex).Caption = " "
       lblDataRowCol02(RvwRowIndex).Caption = " "
       lblDataRowCol03(RvwRowIndex).Caption = " "
       lblDataRowCol04(RvwRowIndex).Caption = " "
       lblDataRowCol05(RvwRowIndex).Caption = " "
       lblDataRowCol06(RvwRowIndex).Caption = " "
       lblDataRowCol07(RvwRowIndex).Caption = " "
       lblDataRowCol08(RvwRowIndex).Caption = " "
    Next RvwRowIndex
    RvwCurRowIndex = 0
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

Private Sub ClearHeader()
SetErrModule 818, 61
If UseLocalErrorHandler Then On Error GoTo localhandler
    ' Clear Header Rows
    lblHdrRowCol00(0).Caption = " "
    lblHdrRowCol01(0).Caption = " "
    lblHdrRowCol02(0).Caption = " "
    lblHdrRowCol03(0).Caption = " "
    lblHdrRowCol04(0).Caption = " "
    lblHdrRowCol05(0).Caption = " "
    lblHdrRowCol06(0).Caption = " "
    lblHdrRowCol07(0).Caption = " "
    lblHdrRowCol08(0).Caption = " "
    lblHdrRowCol00(1).Caption = " "
    lblHdrRowCol01(1).Caption = " "
    lblHdrRowCol02(1).Caption = " "
    lblHdrRowCol03(1).Caption = " "
    lblHdrRowCol04(1).Caption = " "
    lblHdrRowCol05(1).Caption = " "
    lblHdrRowCol06(1).Caption = " "
    lblHdrRowCol07(1).Caption = " "
    lblHdrRowCol08(1).Caption = " "
    
    lblHdrRowCol00(1).ForeColor = lblHdrRowCol00(0).ForeColor
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

Private Sub SetupColData()
Dim iCol As Integer
Dim iShift As Integer
SetErrModule 818, 80
If UseLocalErrorHandler Then On Error GoTo localhandler

    iShift = IIf((Stn_ActiveShift(RvwStn) > 0), Stn_ActiveShift(RvwStn), 1)
    iCol = 0
    ' set column 0
    RvwColData(iCol).ColAlign = "CENTER"
    RvwColData(iCol).ColLeft = 120
    RvwColData(iCol).ColWidth = 10
    RvwColData(iCol).DataFormat = "HH:MM:SS"
    RvwColData(iCol).DataName = "Next"
    RvwColData(iCol).Header1 = "Next"
    RvwColData(iCol).Header2 = " "
    RvwColData(iCol).Header3 = " "
    RvwColData(iCol).InUse = True
    
    Select Case RvwMode
        Case VBLOAD
            ' LOAD DATA COLUMNS
            ' Time
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "HH:MM:SS"
            RvwColData(iCol).DataName = "Time"
            RvwColData(iCol).Header1 = "Time"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = " "
            RvwColData(iCol).InUse = True
            
            Select Case STN_INFO(RvwStn).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE, STN_ORVR2_TYPE
                    ' Nitrogen Flow
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 10
                    RvwColData(iCol).DataFormat = "#0.000"
                    RvwColData(iCol).DataName = "NitFlow"
                    RvwColData(iCol).Header1 = "Nit Flow"
                    RvwColData(iCol).Header2 = " "
                    RvwColData(iCol).Header3 = "(slpm)"
                    RvwColData(iCol).InUse = True
                    ' Mix
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 10
                    RvwColData(iCol).DataFormat = "##0.0"
                    RvwColData(iCol).DataName = "Mix"
                    RvwColData(iCol).Header1 = "Mix %"
                    RvwColData(iCol).Header2 = " "
                    RvwColData(iCol).Header3 = "(% Btn)"
                    RvwColData(iCol).InUse = True
                    ' Butane Flow
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 10
                    RvwColData(iCol).DataFormat = "#0.000"
                    RvwColData(iCol).DataName = "BtnFlow"
                    RvwColData(iCol).Header1 = "Btn Flow"
                    RvwColData(iCol).Header2 = " "
                    RvwColData(iCol).Header3 = "(slpm)"
                    RvwColData(iCol).InUse = True
                    ' Load Rate in Grams
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 10
                    RvwColData(iCol).DataFormat = "#0.0"
                    RvwColData(iCol).DataName = "LoadRate"
                    RvwColData(iCol).Header1 = "Load Rate"
                    RvwColData(iCol).Header2 = " "
                    RvwColData(iCol).Header3 = "(gm/hr)"
                    RvwColData(iCol).InUse = True
                    ' Butane Flow Totalized
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 10
                    RvwColData(iCol).DataFormat = "##0.00"
                    RvwColData(iCol).DataName = "LoadGrams"
                    RvwColData(iCol).Header1 = "Total"
                    RvwColData(iCol).Header2 = " "
                    RvwColData(iCol).Header3 = "(grams)"
                    RvwColData(iCol).InUse = True
                
                Case STN_LIVEFUEL_TYPE
                    ' Vapor Carrier Flow
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 20
                    RvwColData(iCol).DataFormat = "#0.000"
                    RvwColData(iCol).DataName = "NitFlow"
                    RvwColData(iCol).Header1 = "Vapor Carrier Flow"
                    RvwColData(iCol).Header2 = " "
                    RvwColData(iCol).Header3 = "(slpm)"
                    RvwColData(iCol).InUse = True
                    If StationRecipe(RvwStn, iShift).UsePriScale Then
                        ' load rate
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 10
                        RvwColData(iCol).DataFormat = "##0.00"
                        RvwColData(iCol).DataName = "LoadRate"
                        RvwColData(iCol).Header1 = "LoadRate"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "gm/hr"
                        RvwColData(iCol).InUse = True
                    End If
                    
                Case STN_LIVEREG_TYPE, STN_LIVEORVR2_TYPE
                    If StationRecipe(RvwStn, RvwShift).LiveFuel Then
                        ' Vapor Carrier Flow
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 20
                        RvwColData(iCol).DataFormat = "#0.000"
                        RvwColData(iCol).DataName = "NitFlow"
                        RvwColData(iCol).Header1 = "Vapor Carrier Flow"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "(slpm)"
                        RvwColData(iCol).InUse = True
                        If StationRecipe(RvwStn, iShift).UsePriScale Then
                            ' load rate
                            iCol = iCol + 1
                            RvwColData(iCol).ColAlign = "CENTER"
                            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                            RvwColData(iCol).ColWidth = 10
                            RvwColData(iCol).DataFormat = "##0.00"
                            RvwColData(iCol).DataName = "LoadRate"
                            RvwColData(iCol).Header1 = "LoadRate"
                            RvwColData(iCol).Header2 = " "
                            RvwColData(iCol).Header3 = "gm/hr"
                            RvwColData(iCol).InUse = True
                        End If
                    Else
                        ' Nitrogen Flow
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 10
                        RvwColData(iCol).DataFormat = "#0.000"
                        RvwColData(iCol).DataName = "NitFlow"
                        RvwColData(iCol).Header1 = "Nit Flow"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "(slpm)"
                        RvwColData(iCol).InUse = True
                        ' Mix
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 10
                        RvwColData(iCol).DataFormat = "##0.0"
                        RvwColData(iCol).DataName = "Mix"
                        RvwColData(iCol).Header1 = "Mix %"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "(% Btn)"
                        RvwColData(iCol).InUse = True
                        ' Butane Flow
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 10
                        RvwColData(iCol).DataFormat = "#0.000"
                        RvwColData(iCol).DataName = "BtnFlow"
                        RvwColData(iCol).Header1 = "Btn Flow"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "(slpm)"
                        RvwColData(iCol).InUse = True
                        ' Load Rate in Grams
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 10
                        RvwColData(iCol).DataFormat = "#0.0"
                        RvwColData(iCol).DataName = "LoadRate"
                        RvwColData(iCol).Header1 = "Load Rate"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "(gm/hr)"
                        RvwColData(iCol).InUse = True
                        ' Butane Flow Totalized
                        iCol = iCol + 1
                        RvwColData(iCol).ColAlign = "CENTER"
                        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                        RvwColData(iCol).ColWidth = 10
                        RvwColData(iCol).DataFormat = "##0.00"
                        RvwColData(iCol).DataName = "LoadGrams"
                        RvwColData(iCol).Header1 = "Total"
                        RvwColData(iCol).Header2 = " "
                        RvwColData(iCol).Header3 = "(gm Btn)"
                        RvwColData(iCol).InUse = True
                    End If
                                                            
                Case STN_COMBO3_TYPE
                    ' future
                    
            End Select
            
            If StationRecipe(RvwStn, RvwShift).UsePriScale Then
                ' using a Primary Scale
                iCol = iCol + 1
                RvwColData(iCol).ColAlign = "CENTER"
                RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                RvwColData(iCol).ColWidth = 10
                RvwColData(iCol).DataFormat = "####0.00"
                RvwColData(iCol).DataName = "PriScale"
                RvwColData(iCol).Header1 = "Pri " & Format(StationRecipe(RvwStn, RvwShift).PriScaleNo, "#0")
                RvwColData(iCol).Header2 = "Scale #" & Format(StationRecipe(RvwStn, RvwShift).PriScaleNo, "#0")
                RvwColData(iCol).Header3 = "(grams)"
                RvwColData(iCol).InUse = True
            End If
            
            If StationRecipe(RvwStn, RvwShift).UseAuxScale Then
                ' using an Aux Scale
                iCol = iCol + 1
                RvwColData(iCol).ColAlign = "CENTER"
                RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                RvwColData(iCol).ColWidth = 10
                RvwColData(iCol).DataFormat = "####0.00"
                RvwColData(iCol).DataName = "AuxScale"
                RvwColData(iCol).Header1 = "Aux " & Format(StationRecipe(RvwStn, RvwShift).AuxScaleNo, "#0")
                RvwColData(iCol).Header2 = "Scale #" & Format(StationRecipe(RvwStn, RvwShift).AuxScaleNo, "#0")
                RvwColData(iCol).Header3 = "(grams)"
                RvwColData(iCol).InUse = True
            End If
            
            If ((STN_INFO(RvwStn).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(RvwStn).Type = STN_LIVEREG_TYPE) And StationRecipe(RvwStn, RvwShift).LiveFuel) Or ((STN_INFO(RvwStn).Type = STN_LIVEORVR2_TYPE) And StationRecipe(RvwStn, RvwShift).LiveFuel)) Then
            
                '   Cycles
                iCol = iCol + 1
                RvwColData(iCol).ColAlign = "CENTER"
                RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                RvwColData(iCol).ColWidth = 13
                RvwColData(iCol).DataFormat = "#0"
                RvwColData(iCol).DataName = "LiveFuelCycles"
                RvwColData(iCol).Header1 = "Cycles"
                RvwColData(iCol).Header2 = " "
                RvwColData(iCol).Header3 = "SinceRefill"
                RvwColData(iCol).InUse = True
                
                If ((STN_INFO(RvwStn).ADF_TANKTYPE > 10) And (STN_INFO(RvwStn).ADF_TANKTYPE <= 20)) Then
                    ' Live Fuel Station - AutoDrainFill & Heater
                    '   Fuel Temp
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 10
                    RvwColData(iCol).DataFormat = "##0.0"
                    RvwColData(iCol).DataName = "LiveFuelTemp"
                    RvwColData(iCol).Header1 = "Fuel Temp"
                    RvwColData(iCol).Header2 = " "
                    If USINGC Then RvwColData(iCol).Header3 = "(deg C)"
                    If USINGF Then RvwColData(iCol).Header3 = "(deg F)"
                    RvwColData(iCol).InUse = True
                End If
                
            Else
            
                If StationRecipe(RvwStn, RvwShift).UsePriScale Then
                    ' weight change
                    iCol = iCol + 1
                    RvwColData(iCol).ColAlign = "CENTER"
                    RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                    RvwColData(iCol).ColWidth = 12
                    RvwColData(iCol).DataFormat = "####0.00"
                    RvwColData(iCol).DataName = "WtChange"
                    RvwColData(iCol).Header1 = "WtChange"
                    RvwColData(iCol).Header2 = "(this Load)"
                    RvwColData(iCol).Header3 = "(grams)"
                    RvwColData(iCol).InUse = True
                End If
            
            
                
            End If
           
           
        Case VBPURGE
            ' PURGE DATA COLUMNS
            '   Time
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "HH:MM:SS"
            RvwColData(iCol).DataName = "Time"
            RvwColData(iCol).Header1 = "Time"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = " "
            RvwColData(iCol).InUse = True
            '   PurgeAir Flow
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "###0.00"
            RvwColData(iCol).DataName = "PurgeFlow"
            RvwColData(iCol).Header1 = "PA Flow"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = "(slpm)"
            RvwColData(iCol).InUse = True
            '   PurgeAir Temp
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "##0.0"
            RvwColData(iCol).DataName = "PurgeTemp"
            RvwColData(iCol).Header1 = "PA Temp"
            RvwColData(iCol).Header2 = " "
            If USINGC Then RvwColData(iCol).Header3 = "(deg C)"
            If USINGF Then RvwColData(iCol).Header3 = "(deg F)"
            RvwColData(iCol).InUse = True
            '   PurgeAir Humidity
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "##0.0"
            RvwColData(iCol).DataName = "PurgeMoist"
            RvwColData(iCol).Header1 = "PA Moist"
            RvwColData(iCol).Header2 = " "
            If USINGMoist_RH Then RvwColData(iCol).Header3 = "(% rH)"
            If USINGMoist_Grains Then RvwColData(iCol).Header3 = "(Grn/Lb)"
            RvwColData(iCol).InUse = True
            
            If StationRecipe(RvwStn, RvwShift).UsePriScale Then
                ' using a Primary Scale
                iCol = iCol + 1
                RvwColData(iCol).ColAlign = "CENTER"
                RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                RvwColData(iCol).ColWidth = 10
                RvwColData(iCol).DataFormat = "####0.00"
                RvwColData(iCol).DataName = "PriScale"
                RvwColData(iCol).Header1 = "Pri " & Format(StationRecipe(RvwStn, RvwShift).PriScaleNo, "#0")
                RvwColData(iCol).Header2 = "Scale #" & Format(StationRecipe(RvwStn, RvwShift).PriScaleNo, "#0")
                RvwColData(iCol).Header3 = "(grams)"
                RvwColData(iCol).InUse = True
            End If
            
            If StationRecipe(RvwStn, RvwShift).UseAuxScale Then
                ' using an Aux Scale
                iCol = iCol + 1
                RvwColData(iCol).ColAlign = "CENTER"
                RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                RvwColData(iCol).ColWidth = 10
                RvwColData(iCol).DataFormat = "####0.00"
                RvwColData(iCol).DataName = "AuxScale"
                RvwColData(iCol).Header1 = "Aux " & Format(StationRecipe(RvwStn, RvwShift).AuxScaleNo, "#0")
                RvwColData(iCol).Header2 = "Scale #" & Format(StationRecipe(RvwStn, RvwShift).AuxScaleNo, "#0")
                RvwColData(iCol).Header3 = "(grams)"
                RvwColData(iCol).InUse = True
            End If
            
            If StationRecipe(RvwStn, RvwShift).UsePriScale Then
                ' weight change
                iCol = iCol + 1
                RvwColData(iCol).ColAlign = "CENTER"
                RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
                RvwColData(iCol).ColWidth = 12
                RvwColData(iCol).DataFormat = "####0.00"
                RvwColData(iCol).DataName = "WtChange"
                RvwColData(iCol).Header1 = "WtChange"
                RvwColData(iCol).Header2 = "(thisPurge)"
                RvwColData(iCol).Header3 = "(grams)"
                RvwColData(iCol).InUse = True
            End If
            
            '   Purge Volume
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 14
            RvwColData(iCol).DataFormat = "####0.0"
            RvwColData(iCol).DataName = "PurgeVol"
            RvwColData(iCol).Header1 = "Total Volume"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = "(liters)"
            RvwColData(iCol).InUse = True
                        
        Case VBLEAK
            ' LEAK CHECK DATA COLUMNS
            '   Time
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "HH:MM:SS"
            RvwColData(iCol).DataName = "Time"
            RvwColData(iCol).Header1 = "Time"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = " "
            RvwColData(iCol).InUse = True
            '   Pressure
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "CENTER"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 10
            RvwColData(iCol).DataFormat = "##0.00"
            RvwColData(iCol).DataName = "Pressure"
            RvwColData(iCol).Header1 = "Pressure"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = "(psig)"
            RvwColData(iCol).InUse = True
            '   Comment
            iCol = iCol + 1
            RvwColData(iCol).ColAlign = "LEFT"
            RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
            RvwColData(iCol).ColWidth = 28
            RvwColData(iCol).DataFormat = "text"
            RvwColData(iCol).DataName = "Comment"
            RvwColData(iCol).Header1 = "Comment"
            RvwColData(iCol).Header2 = " "
            RvwColData(iCol).Header3 = " "
            RvwColData(iCol).InUse = True
    
    End Select

    If iCol <= (RvwMaxColNum - 1) Then
       '   Test Time
       iCol = iCol + 1
       RvwColData(iCol).ColAlign = "CENTER"
       RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
       RvwColData(iCol).ColWidth = 15
       RvwColData(iCol).DataFormat = "#,###,##0.000"
       RvwColData(iCol).DataName = "TestTime"
       RvwColData(iCol).Header1 = "Test Time"
       RvwColData(iCol).Header2 = " "
       RvwColData(iCol).Header3 = "(seconds)"
       RvwColData(iCol).InUse = True
    End If
    
    ' set remaining columns to Unused
    Do While iCol < RvwMaxColNum
        iCol = iCol + 1
        RvwColData(iCol).ColAlign = "CENTER"
        RvwColData(iCol).ColLeft = RvwColData(iCol - 1).ColLeft + (RvwCharWidth * (RvwColData(iCol - 1).ColWidth))
        RvwColData(iCol).ColWidth = 0
        RvwColData(iCol).DataFormat = "0"
        RvwColData(iCol).DataName = "unused"
        RvwColData(iCol).Header1 = " "
        RvwColData(iCol).Header2 = " "
        RvwColData(iCol).Header3 = " "
        RvwColData(iCol).InUse = False
    Loop

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

Private Sub tmrClock_Timer()
Dim iCol As Integer
Dim tmp As Date
Dim dataVal As String
Dim alignCode As Integer

SetErrModule 818, 900
If UseLocalErrorHandler Then On Error GoTo localhandler

' Update the Navigate Toolbar buttons
UpdateNavigateBtns
' update status bars
UpdateStatusBars
' update status display
If (DateDiff("s", LastStatusDTS, Now()) > 1) Then UpdateStatus

If RvwMode <> VBCOMPLETE And RvwMode <> VBIDLEWAITING And RvwMode <> VBIDLE Then
    
    ' only refresh if style=Normal (i.e. Not WaitingAnswer)
    If RvwStyle = "Normal" Then
        ' only show 'current' data if previous line has a valid time and not last line
        If ((RvwCurRowIndex > 0) Or (IsDate(lblDataRowCol01(RvwCurRowIndex).Caption))) _
            And (RvwCurRowIndex < (RvwRows - 1)) Then
                ' only show 'current' data if at EOF
                If Not AtEOF Then
                
                    ' cycle through the columns
                    For iCol = 0 To 8
                    
                        ' Get Alignment Code for this Col
                        Select Case RvwColData(iCol).ColAlign
                            Case "LEFT"
                                alignCode = 0
                            Case "CENTER"
                                alignCode = 2
                            Case "RIGHT"
                                alignCode = 1
                        End Select
                    
                        ' Get Data Value as a String
                        Select Case RvwMode
                            Case VBLOAD
                                Select Case RvwColData(iCol).DataName
                                    Case "Next"
                                        If IsDate(lblDataRowCol01(RvwCurRowIndex).Caption) Then
                                            tmp = DateAdd("s", RvwInterval, CDate(lblDataRowCol01(RvwCurRowIndex).Caption))
                                            ' only show current time when the Interval between Successive Data is > 2 sec
                                            dataVal = IIf(RvwInterval > 2, Format(tmp, RvwColData(iCol).DataFormat), " ")
                                        Else
                                            dataVal = " "
                                        End If
                                    Case "Time"
                                        dataVal = Format(Now(), RvwColData(iCol).DataFormat)
                                    Case "NitFlow"
                                        dataVal = Format(Stn_Nit_Flow_PV(RvwStn, RvwShift), RvwColData(iCol).DataFormat)
                                    Case "Mix"
                                        If Stn_Btn_Flow_PV(RvwStn, RvwShift) + Stn_Nit_Flow_PV(RvwStn, RvwShift) <= 0.0001 Then
                                            RvwData = 0
                                        Else
                                            RvwData = 100 * Stn_Btn_Flow_PV(RvwStn, RvwShift) / _
                                                (Stn_Btn_Flow_PV(RvwStn, RvwShift) + Stn_Nit_Flow_PV(RvwStn, RvwShift) + 0.00001)
                                        End If
                                        dataVal = Format(RvwData, RvwColData(iCol).DataFormat)
                                    Case "BtnFlow"
                                        dataVal = Format(Stn_Btn_Flow_PV(RvwStn, RvwShift), RvwColData(iCol).DataFormat)
                                    Case "LoadRate"
                                        dataVal = Format(SlpmToGramsPerHour(Stn_Btn_Flow_PV(RvwStn, RvwShift), StationControl(RvwStn, RvwShift).BtnDensity), RvwColData(iCol).DataFormat)
                                    Case "LoadGrams"
                                        dataVal = Format(LoadControl(RvwStn, RvwShift).loadTotalGrams, RvwColData(iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).PriScaleWt, RvwColData(iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).AuxScaleWt, RvwColData(iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).PriScaleWt, RvwColData(iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(LoadControl(RvwStn, RvwShift).TotalWtChg, RvwColData(iCol).DataFormat)
                                    Case "LiveFuelTemp"
                                        dataVal = Format(Stn_AIO(RvwStn, asFuelTankTemp).EUValue, RvwColData(iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).TestTimer, RvwColData(iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            Case VBPURGE
                                Select Case RvwColData(iCol).DataName
                                    Case "Next"
                                        tmp = DateAdd("s", RvwInterval, CDate(lblDataRowCol01(RvwCurRowIndex).Caption))
                                        ' only show current time when the Interval between Successive Data is > 2 sec
                                        dataVal = IIf(RvwInterval > 2, Format(tmp, RvwColData(iCol).DataFormat), " ")
                                    Case "Time"
                                        dataVal = Format(Now(), RvwColData(iCol).DataFormat)
                                    Case "PurgeFlow"
                                        dataVal = Format(Stn_AIO(RvwStn, asPurgeAirFlow).EUValue, RvwColData(iCol).DataFormat)
                                    Case "PurgeTemp"
                                        dataVal = Format(PATemp, RvwColData(iCol).DataFormat)
                                    Case "PurgeMoist"
                                        dataVal = Format(PAMoisture, RvwColData(iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).PriScaleWt, RvwColData(iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).AuxScaleWt, RvwColData(iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(PurgeControl(RvwStn, RvwShift).TotalWtChg, RvwColData(iCol).DataFormat)
                                    Case "PurgeVol"
                                        dataVal = Format(PurgeControl(RvwStn, RvwShift).Purge_Total, RvwColData(iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).TestTimer, RvwColData(iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            Case VBLEAK
                                Select Case RvwColData(iCol).DataName
                                    Case "Next"
                                        tmp = DateAdd("s", RvwInterval, CDate(lblDataRowCol01(RvwCurRowIndex).Caption))
                                        ' only show current time when the Interval between Successive Data is > 2 sec
                                        dataVal = IIf(RvwInterval > 2, Format(tmp, RvwColData(iCol).DataFormat), " ")
                                    Case "Time"
                                        dataVal = Format(Now(), RvwColData(iCol).DataFormat)
                                    Case "Pressure"
                                        dataVal = Format(PTinvalue, RvwColData(iCol).DataFormat)
                                    Case "Comment"
                                        dataVal = " "
                                    Case "TestTime"
                                        dataVal = Format(StationControl(RvwStn, RvwShift).TestTimer, RvwColData(iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            Case Else
                                dataVal = " "
                        End Select
                        
                        ' Display(& align) the Data Value
                        Select Case iCol
                            Case 0
                                lblDataRowCol00(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol00(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol00(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol00(RvwCurRowIndex + 1).FontItalic = False
                                If RvwColData(iCol).DataName = "Next" Then
                                    lblDataRowCol00(RvwCurRowIndex + 1).ForeColor = ModeForeColor(RvwMode)
                                Else
                                    lblDataRowCol00(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                End If
                                lblDataRowCol00(RvwCurRowIndex + 1).Caption = dataVal
                            Case 1
                                lblDataRowCol01(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol01(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol01(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol01(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol01(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol01(RvwCurRowIndex + 1).Caption = dataVal
                            Case 2
                                lblDataRowCol02(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol02(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol02(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol02(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol02(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol02(RvwCurRowIndex + 1).Caption = dataVal
                            Case 3
                                lblDataRowCol03(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol03(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol03(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol03(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol03(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol03(RvwCurRowIndex + 1).Caption = dataVal
                            Case 4
                                lblDataRowCol04(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol04(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol04(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol04(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol04(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol04(RvwCurRowIndex + 1).Caption = dataVal
                            Case 5
                                lblDataRowCol05(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol05(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol05(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol05(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol05(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol05(RvwCurRowIndex + 1).Caption = dataVal
                            Case 6
                                lblDataRowCol06(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol06(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol06(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol06(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol06(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol06(RvwCurRowIndex + 1).Caption = dataVal
                            Case 7
                                lblDataRowCol07(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol07(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol07(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol07(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol07(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol07(RvwCurRowIndex + 1).Caption = dataVal
                            Case 8
                                lblDataRowCol08(RvwCurRowIndex + 1).Left = RvwColData(iCol).ColLeft
                                lblDataRowCol08(RvwCurRowIndex + 1).Width = RvwCharWidth * RvwColData(iCol).ColWidth
                                lblDataRowCol08(RvwCurRowIndex + 1).Alignment = alignCode
                                lblDataRowCol08(RvwCurRowIndex + 1).FontItalic = False
                                lblDataRowCol08(RvwCurRowIndex + 1).ForeColor = lblDataHiLiTemplate.ForeColor
                                lblDataRowCol08(RvwCurRowIndex + 1).Caption = dataVal
                            Case Else
                                ' only nine data columns on display
                        End Select
                        
                    Next iCol
        
                Else
                    ' set all columns to 'blank'
                    lblDataRowCol00(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol01(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol02(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol03(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol04(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol05(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol06(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol07(RvwCurRowIndex + 1).Caption = ""
                    lblDataRowCol08(RvwCurRowIndex + 1).Caption = ""
                End If
            End If
    
        ' only refresh if new data has been written to the db file
        If StationControl(RvwStn, RvwShift).NewDataInDB Then
        
            ' refresh the page only if the current page is showing the latest data
            If Not NewCycle _
                And RvwResponse = vbIgnore _
                And StationControl(RvwStn, RvwShift).DBFile <> "" _
                And (RvwCurRecCnt < RvwFirstRecord + RvwRows + 11) _
                And RvwCurRecCnt >= RvwFirstRecord And RvwCurRecCnt <= RvwMaxRecDsp Then
                
                        ' read data from db file & refresh the displayed data
                        Reader StationControl(RvwStn, RvwShift).DBFile
            End If
            
            ' reset new data flag
            StationControl(RvwStn, RvwShift).NewDataInDB = False

        End If
        
        ' at the end of the data
        If ((RvwMode <> StationControl(RvwStn, RvwShift).Mode) _
            Or (RvwMode <> VBLEAK And (RvwCycle <> StationControl(RvwStn, RvwShift).CurrCycle))) _
                    And _
            (RvwCurRecCnt <= RvwFirstRecord + RvwRows - 1) Then
                AtEOF = True
        End If

    End If
    
    ' Job is Ending
'    AtEOF = True
End If
    
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

Private Sub UpdateStatusBars()
    ' Status Bar #1
    pnlEstop.BackColor = frmMainMenu.pnlEstop.BackColor
    pnlFlow.BackColor = frmMainMenu.pnlFlow.BackColor
    pnlBtn20.BackColor = frmMainMenu.pnlBtn20.BackColor
    pnlDoors.BackColor = frmMainMenu.pnlDoors.BackColor
    pnlComms.BackColor = frmMainMenu.pnlComms.BackColor
    pnlPAcomm.BackColor = frmMainMenu.pnlPAcomm.BackColor
    pnlEstop.ToolTipText = frmMainMenu.pnlEstop.ToolTipText
    pnlFlow.ToolTipText = frmMainMenu.pnlFlow.ToolTipText
    pnlBtn20.ToolTipText = frmMainMenu.pnlBtn20.ToolTipText
    pnlDoors.ToolTipText = frmMainMenu.pnlDoors.ToolTipText
    pnlComms.ToolTipText = frmMainMenu.pnlComms.ToolTipText
    pnlPAcomm.ToolTipText = frmMainMenu.pnlPAcomm.ToolTipText
    pnlMix.BackColor = frmMainMenu.pnlMix.BackColor
    pnlMix.ToolTipText = frmMainMenu.pnlMix.ToolTipText
    pnlMix.Top = frmMainMenu.pnlMix.Top
    pnlMessageFrame.Width = frmMainMenu.pnlMessageFrame.Width
    pnlMessage.Font = frmMainMenu.pnlMessage.Font
    pnlMessage.FontSize = frmMainMenu.pnlMessage.FontSize
    pnlMessage.Width = frmMainMenu.pnlMessage.Width
    pnlMessage.BackColor = SysMessage_BackColor
    pnlMessage.ForeColor = SysMessage_ForeColor
    pnlMessage.Caption = SysMessage_Text
    pnlMessage.ToolTipText = SysMessage_Tooltip
    pnlPurgeAir.Left = frmMainMenu.pnlPurgeAir.Left
    pnlPurgeAir.Width = frmMainMenu.pnlPurgeAir.Width
    pnlPurgeAir.ForeColor = PurgeAirMsg_ForeColor
    pnlPurgeAir.Caption = PurgeAirMsg_Text
    pnlPurgeAir.ToolTipText = PurgeAirMsg_ToolTip
End Sub

Sub UpdateNavigateBtns()

'
' Routine Name:  UpdateNavigateBtns
' Description:
' Updates the Navigate toolbar buttons
'
Dim iKeyCount As Integer
 
SetErrModule 818, 10101
If UseLocalErrorHandler Then On Error GoTo localhandler

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
        If CheckPass("N", False) And USINGREMCANLOAD Then
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
        
        Select Case LocalPagControl.Type
        
            Case pagNone
                ' no PurgeAir Generator control
            
            Case pagAlone
                ' Stand-Alone PurgeAir Generator control
            
            Case pagMaster
                ' AK Master PurgeAir Generator control
                ' AK Server
                If CheckPass("7", False) Then
                    tbrNavigate.Buttons("ak_server").Enabled = True
                    mnuak_server.Enabled = True
                Else
                    mnuak_server.Enabled = False
                    tbrNavigate.Buttons("ak_server").Enabled = False
                End If
        
            Case pagClient
                ' AK Client PurgeAir Generator control
                ' AK Client
                If CheckPass("7", False) Then
                    tbrNavigate.Buttons("ak_client").Enabled = True
                    mnuak_client.Enabled = True
                Else
                    mnuak_client.Enabled = False
                    tbrNavigate.Buttons("ak_client").Enabled = False
                End If
                ' AK Server
                If CheckPass("7", False) Then
                    tbrNavigate.Buttons("ak_server").Enabled = True
                    mnuak_server.Enabled = True
                Else
                    mnuak_server.Enabled = False
                    tbrNavigate.Buttons("ak_server").Enabled = False
                End If
            
        End Select
    
        ' Simulation
        If Not IoComOn And USINGSIMULATION And CheckPass("H", False) Then
            tbrNavigate.Buttons("simulation").Visible = True
            tbrNavigate.Buttons("simulation").Enabled = True
            tbrNavigate.Buttons("simulation").ToolTipText = "Simulation Control Panel"
'            tbrNavigate.Buttons("fillright").Width = 2550
'            mnuSimulation.Enabled = True
        Else
'            mnuSimulation.Enabled = False
            tbrNavigate.Buttons("simulation").Visible = False
            tbrNavigate.Buttons("simulation").Enabled = False
            tbrNavigate.Buttons("simulation").ToolTipText = ""
'            tbrNavigate.Buttons("fillright").Width = 3200
        End If
        
        ' Operator Manual
        If CheckPass("D", False) Then
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


