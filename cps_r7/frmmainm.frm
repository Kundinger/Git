VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainMenu 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Canister Preconditioning System"
   ClientHeight    =   10605
   ClientLeft      =   195
   ClientTop       =   900
   ClientWidth     =   14880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmmainm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10961.83
   ScaleMode       =   0  'User
   ScaleWidth      =   14970.46
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbxBottom 
      Align           =   2  'Align Bottom
      Height          =   460
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   14820
      TabIndex        =   133
      Top             =   10140
      Width           =   14880
      Begin Threed.SSPanel pnlAlarms 
         Height          =   405
         Left            =   0
         TabIndex        =   134
         Top             =   0
         Width           =   6030
         _Version        =   65536
         _ExtentX        =   10636
         _ExtentY        =   714
         _StockProps     =   15
         ForeColor       =   -2147483630
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
            TabIndex        =   135
            ToolTipText     =   "EMERGENCY Stop Pressed"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "ESTOP"
            ForeColor       =   -2147483630
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
            TabIndex        =   136
            ToolTipText     =   "Loss of Flow"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "FLOW"
            ForeColor       =   -2147483630
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
            TabIndex        =   137
            ToolTipText     =   "20% Butane LEL Alarm"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "LEL 20"
            ForeColor       =   -2147483630
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
            TabIndex        =   138
            ToolTipText     =   "A door is open"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "DOORS"
            ForeColor       =   -2147483630
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
            TabIndex        =   139
            ToolTipText     =   "I/O Communication Error"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "IOCOM"
            ForeColor       =   -2147483630
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
            TabIndex        =   140
            ToolTipText     =   "Mixture OutOfTolerance"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "MIX"
            ForeColor       =   -2147483630
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
            TabIndex        =   217
            ToolTipText     =   "PurgeAirSystem Communication Not Online"
            Top             =   75
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "PACOM"
            ForeColor       =   -2147483630
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
         TabIndex        =   141
         Top             =   0
         Width           =   5595
         _Version        =   65536
         _ExtentX        =   9869
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
         TabIndex        =   142
         Top             =   0
         Width           =   3360
         _Version        =   65536
         _ExtentX        =   5927
         _ExtentY        =   714
         _StockProps     =   15
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
            Left            =   60
            TabIndex        =   143
            Top             =   60
            Width           =   3240
            _Version        =   65536
            _ExtentX        =   5715
            _ExtentY        =   494
            _StockProps     =   15
            Caption         =   "message"
            ForeColor       =   -2147483630
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
   End
   Begin Threed.SSCommand cmdLargeFont 
      Height          =   375
      Left            =   13200
      TabIndex        =   131
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Large Font"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdRegFont 
      Height          =   375
      Left            =   13200
      TabIndex        =   130
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Reg Font"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   500
      Left            =   9840
      Top             =   9240
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   500
      Left            =   9840
      Top             =   8760
   End
   Begin VB.Timer tmrTimer 
      Index           =   7
      Interval        =   1
      Left            =   9840
      Top             =   8280
   End
   Begin VB.Timer tmrTimer 
      Index           =   3
      Interval        =   1000
      Left            =   9840
      Top             =   6360
   End
   Begin VB.Timer tmrTimer 
      Index           =   2
      Interval        =   1100
      Left            =   9840
      Top             =   5880
   End
   Begin VB.Timer tmrTimer 
      Index           =   5
      Interval        =   1100
      Left            =   9840
      Top             =   7320
   End
   Begin VB.Timer tmrTimer 
      Index           =   4
      Interval        =   500
      Left            =   9840
      Top             =   6840
   End
   Begin VB.Timer tmrTimer 
      Index           =   6
      Interval        =   350
      Left            =   9840
      Top             =   7800
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   0
         Left            =   75
         TabIndex        =   1
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 1"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   0
               Left            =   240
               TabIndex        =   46
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   47
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   120
               TabIndex        =   49
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1725
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   1
            Left            =   0
            TabIndex        =   50
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   51
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   52
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   240
               TabIndex        =   54
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   18
         Left            =   120
         TabIndex        =   145
         Top             =   3240
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   18
            Left            =   120
            TabIndex        =   146
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   18
               Left            =   240
               TabIndex        =   147
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   18
               Left            =   240
               TabIndex        =   148
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   18
            Left            =   120
            TabIndex        =   149
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   18
               Left            =   120
               TabIndex        =   150
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   19
         Left            =   150
         TabIndex        =   151
         Top             =   4485
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   19
            Left            =   0
            TabIndex        =   152
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   19
               Left            =   240
               TabIndex        =   153
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   19
               Left            =   240
               TabIndex        =   154
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   19
            Left            =   0
            TabIndex        =   155
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   19
               Left            =   240
               TabIndex        =   156
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin VB.Timer tmrTimer 
      Index           =   1
      Interval        =   500
      Left            =   9840
      Top             =   5400
   End
   Begin VB.Timer tmrScreen 
      Interval        =   1000
      Left            =   4800
      Top             =   8640
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   1
      Left            =   4965
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   1
         Left            =   75
         TabIndex        =   5
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 2"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   2
            Left            =   0
            TabIndex        =   55
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   56
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   57
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   120
               TabIndex        =   59
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1725
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   3
            Left            =   0
            TabIndex        =   60
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   3
               Left            =   240
               TabIndex        =   61
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   62
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   3
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   120
               TabIndex        =   64
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   20
         Left            =   0
         TabIndex        =   157
         Top             =   3360
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   20
            Left            =   120
            TabIndex        =   158
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   20
               Left            =   240
               TabIndex        =   159
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   20
               Left            =   240
               TabIndex        =   160
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   20
            Left            =   120
            TabIndex        =   161
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   20
               Left            =   120
               TabIndex        =   162
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   21
         Left            =   30
         TabIndex        =   163
         Top             =   4605
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   21
            Left            =   0
            TabIndex        =   164
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   21
               Left            =   240
               TabIndex        =   165
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   21
               Left            =   240
               TabIndex        =   166
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   21
            Left            =   0
            TabIndex        =   167
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   21
               Left            =   240
               TabIndex        =   168
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   2
      Left            =   9945
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   2
         Left            =   75
         TabIndex        =   9
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 3"
         ForeColor       =   8421376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   4
            Left            =   0
            TabIndex        =   65
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   4
               Left            =   240
               TabIndex        =   66
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   67
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   120
               TabIndex        =   69
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   5
         Left            =   90
         TabIndex        =   11
         Top             =   1720
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   5
            Left            =   0
            TabIndex        =   70
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   5
               Left            =   240
               TabIndex        =   71
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   72
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   5
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   120
               TabIndex        =   74
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   22
         Left            =   0
         TabIndex        =   169
         Top             =   3360
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   22
            Left            =   120
            TabIndex        =   170
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   22
               Left            =   240
               TabIndex        =   171
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   22
               Left            =   240
               TabIndex        =   172
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   22
            Left            =   120
            TabIndex        =   173
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   22
               Left            =   120
               TabIndex        =   174
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   23
         Left            =   30
         TabIndex        =   175
         Top             =   4605
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   23
            Left            =   0
            TabIndex        =   176
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   23
               Left            =   240
               TabIndex        =   177
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   23
               Left            =   240
               TabIndex        =   178
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   23
            Left            =   0
            TabIndex        =   179
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   23
               Left            =   240
               TabIndex        =   180
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   3
      Left            =   45
      TabIndex        =   12
      Top             =   3915
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   3
         Left            =   75
         TabIndex        =   13
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 4"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   6
         Left            =   90
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   6
            Left            =   0
            TabIndex        =   75
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   6
               Left            =   240
               TabIndex        =   76
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   77
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   6
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   6
               Left            =   120
               TabIndex        =   79
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   7
         Left            =   90
         TabIndex        =   15
         Top             =   1720
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   7
            Left            =   0
            TabIndex        =   80
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   7
               Left            =   240
               TabIndex        =   81
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   82
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   7
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   7
               Left            =   120
               TabIndex        =   84
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   24
         Left            =   0
         TabIndex        =   181
         Top             =   3240
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   24
            Left            =   120
            TabIndex        =   182
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   24
               Left            =   240
               TabIndex        =   183
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   24
               Left            =   240
               TabIndex        =   184
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   24
            Left            =   120
            TabIndex        =   185
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   24
               Left            =   120
               TabIndex        =   186
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   25
         Left            =   30
         TabIndex        =   187
         Top             =   4485
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   25
            Left            =   0
            TabIndex        =   188
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   25
               Left            =   240
               TabIndex        =   189
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   25
               Left            =   240
               TabIndex        =   190
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   25
            Left            =   0
            TabIndex        =   191
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   25
               Left            =   240
               TabIndex        =   192
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   4
      Left            =   4965
      TabIndex        =   16
      Top             =   3915
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   4
         Left            =   75
         TabIndex        =   17
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 5"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   8
         Left            =   90
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   8
            Left            =   0
            TabIndex        =   85
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   8
               Left            =   240
               TabIndex        =   86
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   87
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   8
            Left            =   0
            TabIndex        =   88
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   8
               Left            =   120
               TabIndex        =   89
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   1725
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   9
            Left            =   0
            TabIndex        =   90
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   9
               Left            =   240
               TabIndex        =   91
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   9
               Left            =   240
               TabIndex        =   92
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   9
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   9
               Left            =   120
               TabIndex        =   94
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   26
         Left            =   120
         TabIndex        =   193
         Top             =   3240
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   26
            Left            =   120
            TabIndex        =   194
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   26
               Left            =   240
               TabIndex        =   195
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   26
               Left            =   240
               TabIndex        =   196
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   26
            Left            =   120
            TabIndex        =   197
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   26
               Left            =   120
               TabIndex        =   198
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   27
         Left            =   150
         TabIndex        =   199
         Top             =   4485
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   27
            Left            =   0
            TabIndex        =   200
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   27
               Left            =   240
               TabIndex        =   201
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   27
               Left            =   240
               TabIndex        =   202
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   27
            Left            =   0
            TabIndex        =   203
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   27
               Left            =   240
               TabIndex        =   204
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   5
      Left            =   9945
      TabIndex        =   20
      Top             =   3915
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   5
         Left            =   75
         TabIndex        =   21
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 6"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   10
         Left            =   90
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   10
            Left            =   0
            TabIndex        =   95
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   10
               Left            =   240
               TabIndex        =   96
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   97
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   10
            Left            =   0
            TabIndex        =   98
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   10
               Left            =   120
               TabIndex        =   99
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   11
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   11
            Left            =   0
            TabIndex        =   100
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   11
               Left            =   240
               TabIndex        =   101
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   11
               Left            =   240
               TabIndex        =   102
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   11
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   11
               Left            =   120
               TabIndex        =   104
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   28
         Left            =   120
         TabIndex        =   205
         Top             =   3120
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   28
            Left            =   120
            TabIndex        =   206
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   28
               Left            =   240
               TabIndex        =   207
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   16744448
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   28
               Left            =   240
               TabIndex        =   208
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   28
            Left            =   120
            TabIndex        =   209
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   28
               Left            =   120
               TabIndex        =   210
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   29
         Left            =   150
         TabIndex        =   211
         Top             =   4365
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   29
            Left            =   0
            TabIndex        =   212
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   29
               Left            =   240
               TabIndex        =   213
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   29
               Left            =   240
               TabIndex        =   214
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   29
            Left            =   0
            TabIndex        =   215
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   29
               Left            =   240
               TabIndex        =   216
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   6
      Left            =   45
      TabIndex        =   24
      Top             =   6985
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   6
         Left            =   75
         TabIndex        =   25
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 7"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   12
         Left            =   90
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   12
            Left            =   0
            TabIndex        =   105
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   12
               Left            =   240
               TabIndex        =   106
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   107
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   12
            Left            =   0
            TabIndex        =   108
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   12
               Left            =   120
               TabIndex        =   109
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   13
         Left            =   90
         TabIndex        =   27
         Top             =   1720
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   13
            Left            =   0
            TabIndex        =   110
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   13
               Left            =   240
               TabIndex        =   111
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   112
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   13
            Left            =   0
            TabIndex        =   113
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   13
               Left            =   120
               TabIndex        =   114
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   7
      Left            =   4970
      TabIndex        =   28
      Top             =   6985
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   7
         Left            =   75
         TabIndex        =   29
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 8"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   14
         Left            =   90
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   14
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   14
               Left            =   120
               TabIndex        =   44
               Top             =   120
               Width           =   3825
            End
         End
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   14
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   14
               Left            =   240
               TabIndex        =   42
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   14
               Left            =   240
               TabIndex        =   41
               Top             =   120
               Width           =   4080
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   15
         Left            =   90
         TabIndex        =   31
         Top             =   1720
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   15
            Left            =   0
            TabIndex        =   115
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   15
               Left            =   240
               TabIndex        =   116
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   15
               Left            =   240
               TabIndex        =   117
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   15
            Left            =   0
            TabIndex        =   118
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   15
               Left            =   120
               TabIndex        =   119
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin Threed.SSPanel pnlStn 
      Height          =   3045
      Index           =   8
      Left            =   9940
      TabIndex        =   32
      Top             =   6985
      Visible         =   0   'False
      Width           =   4875
      _Version        =   65536
      _ExtentX        =   8599
      _ExtentY        =   5371
      _StockProps     =   15
      ForeColor       =   -2147483630
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
      BevelWidth      =   2
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   6
      Begin Threed.SSCommand cmdStation 
         Height          =   435
         Index           =   8
         Left            =   75
         TabIndex        =   33
         ToolTipText     =   "Station Three Details"
         Top             =   60
         Visible         =   0   'False
         Width           =   4740
         _Version        =   65536
         _ExtentX        =   8361
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "Station 9"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         Outline         =   0   'False
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   16
         Left            =   90
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   16
            Left            =   0
            TabIndex        =   120
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   16
               Left            =   240
               TabIndex        =   121
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   16
               Left            =   240
               TabIndex        =   122
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   16
            Left            =   0
            TabIndex        =   123
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   16
               Left            =   120
               TabIndex        =   124
               Top             =   120
               Width           =   3825
            End
         End
      End
      Begin Threed.SSPanel pnlShift 
         Height          =   1230
         Index           =   17
         Left            =   90
         TabIndex        =   35
         Top             =   1720
         Visible         =   0   'False
         Width           =   4710
         _Version        =   65536
         _ExtentX        =   8308
         _ExtentY        =   2170
         _StockProps     =   15
         ForeColor       =   -2147483630
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
         BevelWidth      =   2
         BorderWidth     =   1
         BevelInner      =   1
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   615
            Index           =   17
            Left            =   0
            TabIndex        =   125
            Top             =   480
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   1085
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   2
            BorderWidth     =   4
            BevelInner      =   1
            Autosize        =   3
            Begin Threed.SSPanel pbarActual 
               Height          =   180
               Index           =   17
               Left            =   240
               TabIndex        =   126
               Top             =   360
               Visible         =   0   'False
               Width           =   4080
               _Version        =   65536
               _ExtentX        =   7197
               _ExtentY        =   317
               _StockProps     =   15
               ForeColor       =   -2147483630
               BackColor       =   -2147483646
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelWidth      =   6
               BorderWidth     =   10
               BevelOuter      =   0
               FloodType       =   1
               FloodColor      =   12583104
               FloodShowPct    =   0   'False
            End
            Begin VB.Label lblMode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "mode"
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
               Height          =   255
               Index           =   17
               Left            =   240
               TabIndex        =   127
               Top             =   120
               Width           =   4080
            End
         End
         Begin Threed.SSPanel pnlMsg 
            Height          =   495
            Index           =   17
            Left            =   0
            TabIndex        =   128
            Top             =   0
            Width           =   4455
            _Version        =   65536
            _ExtentX        =   7858
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "SSPanel1"
            BackColor       =   14215660
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Begin VB.Label lblMessage 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "message"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   17
               Left            =   120
               TabIndex        =   129
               Top             =   120
               Width           =   3825
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   132
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1058
      ButtonWidth     =   1058
      ButtonHeight    =   953
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgNavigateNormal 
      Left            =   10560
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":57E2
            Key             =   "login"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":6434
            Key             =   "logout"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":7086
            Key             =   "copyfiles"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":7CD8
            Key             =   "printfiles"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":892A
            Key             =   "reviewdata"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":957C
            Key             =   "watchdata"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":A1CE
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":AE20
            Key             =   "can_master"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":BA72
            Key             =   "configuration"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":C6C4
            Key             =   "sysdef"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":D316
            Key             =   "eventlog"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":E368
            Key             =   "joblist"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":EFBA
            Key             =   "calibration"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":FC0C
            Key             =   "scalemonitor"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1085E
            Key             =   "close"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":114B0
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":12102
            Key             =   "overview"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":12D54
            Key             =   "stndetail"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":139A6
            Key             =   "iomonitor"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":145F8
            Key             =   "simulation"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1524A
            Key             =   "flammablegas"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":15E3C
            Key             =   "opermanual"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":16A8E
            Key             =   "firstaid"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":176E0
            Key             =   "beaconoff"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":18332
            Key             =   "hornoff"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":18F84
            Key             =   "tomcanload"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":19BD6
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1A828
            Key             =   "prof_master"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1B47A
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1C0CC
            Key             =   "rcp_master"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1CD1E
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1D970
            Key             =   "seq_master"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1E5C2
            Key             =   "fueluselog"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1F214
            Key             =   "remotecontrol"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":1FE66
            Key             =   "analoginput"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":20AB8
            Key             =   "mfc"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2170A
            Key             =   "leaktest"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2235C
            Key             =   "ak_server"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":22FAE
            Key             =   "ak_client"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgNavigateDisabled 
      Left            =   11400
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":24000
            Key             =   "login"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":24C52
            Key             =   "logout"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":258A4
            Key             =   "copyfiles"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":264F6
            Key             =   "printfiles"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":27148
            Key             =   "reviewdata"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":27D9A
            Key             =   "watchdata"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":289EC
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2963E
            Key             =   "configuration"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2A290
            Key             =   "sysdef"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2AEE2
            Key             =   "eventlog"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2BF34
            Key             =   "joblist"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2CB86
            Key             =   "scalemonitor"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2D7D8
            Key             =   "calibration"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2E42A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2F07C
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":2FCCE
            Key             =   "overview"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":30920
            Key             =   "stndetail"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":31572
            Key             =   "iomonitor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":321C4
            Key             =   "simulation"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":32E16
            Key             =   "flammablegas"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":33A68
            Key             =   "opermanual"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":346BA
            Key             =   "firstaid"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3530C
            Key             =   "beaconoff"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":35F5E
            Key             =   "hornoff"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":36BB0
            Key             =   "tomcanload"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":37802
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":38454
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":390A6
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":39CF8
            Key             =   "fueluselog"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3A94A
            Key             =   "remotecontrol"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3B59C
            Key             =   "analoginput"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3C1EE
            Key             =   "mfc"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3CE40
            Key             =   "leaktest"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3DA92
            Key             =   "ak_server"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3E6E4
            Key             =   "ak_client"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgNavigateHot 
      Left            =   12120
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":3F736
            Key             =   "login"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":40388
            Key             =   "logout"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":40FDA
            Key             =   "copyfiles"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":41C2C
            Key             =   "printfiles"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4287E
            Key             =   "reviewdata"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":434D0
            Key             =   "watchdata"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":44122
            Key             =   "canisters"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":44D74
            Key             =   "recipes"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":459C6
            Key             =   "configuration"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":46618
            Key             =   "sysdef"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4726A
            Key             =   "eventlog"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":482BC
            Key             =   "joblist"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":48F0E
            Key             =   "calibration"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":49B60
            Key             =   "scalemonitor"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4A7B2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4B404
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4C056
            Key             =   "overview"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4CCA8
            Key             =   "stndetail"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4D8FA
            Key             =   "iomonitor"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4E54C
            Key             =   "simulation"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4F19E
            Key             =   "flammablegas"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":4FD90
            Key             =   "opermanual"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":509E2
            Key             =   "firstaid"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":51634
            Key             =   "beaconoff"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":52286
            Key             =   "hornoff"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":52ED8
            Key             =   "tomcanload"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":53B2A
            Key             =   "purgeprofile"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":5477C
            Key             =   "courses"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":553CE
            Key             =   "fueluselog"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":56020
            Key             =   "remotecontrol"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":56C72
            Key             =   "analoginput"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":578C4
            Key             =   "mfc"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":58516
            Key             =   "leaktest"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":59168
            Key             =   "ak_server"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmainm.frx":59DBA
            Key             =   "ak_client"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   1
      Left            =   9360
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblMsgMode 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "msg / mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   13440
      TabIndex        =   39
      Top             =   10080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblActiveTitleBarColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "ActiveTitleBarColor"
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
      Height          =   180
      Left            =   12720
      TabIndex        =   144
      Top             =   10200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDataColor 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "data"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   13920
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblLightBkgd 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "data"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   13440
      TabIndex        =   38
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblCurrDataColor 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "data"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   12480
      TabIndex        =   36
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
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
      Begin VB.Menu mnuCourses 
         Caption         =   "Co&urses"
      End
      Begin VB.Menu mnuRecipes 
         Caption         =   "&Recipes"
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
         Caption         =   "I/O &Monitor"
      End
      Begin VB.Menu mnuScaleMonitor 
         Caption         =   "&Scale Monitor"
      End
      Begin VB.Menu mnuSimulation 
         Caption         =   "Simulatio&n"
         Visible         =   0   'False
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
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 99 '''''''''''Form MAINMENU.frm ''''''''''''''''''''''
'
Option Explicit
'
Private NrStns, NrShifts As Integer
Private CharWidth, CanLength, RcpLength As Integer
Private StnPerRow, MsgRows As Integer
Private FirstColLeft, FirstRowTop As Integer
Private FirstShiftLeft, FirstShiftTop As Integer
Private StnWidth, StnHeight, ShiftHeight, ShiftWidth As Integer
Private CmdLeft, CmdTop, CmdWidth, CmdHeight As Integer
Private MsgLeft, MsgTop, MsgWidth, MsgHeight As Integer
Private ModeLeft, ModeTop, ModeWidth, ModeHeight As Integer
Private PbarLeft, PbarTop, PbarWidth, PbarHeight As Integer
Private PnlMsgLeft, PnlMsgTop, PnlMsgWidth, PnlMsgHeight As Integer
Private PnlModeLeft, PnlModeTop, PnlModeWidth, PnlModeHeight As Integer
Private StnSpaceHorz, StnSpaceVert As Integer
Private CmdFontSize, MsgFontSize, ModeFontSize As Integer
Private BevelInner, BevelOuter, BevelWidth, BorderWidth As Integer
Private TopOfStations, BottomOfStations, TopToBottomOfStations, FreeSpaceAtBottom As Integer
Private sString As String
Private ScreenIsBuilt As Boolean
Private PaleBkgd As Variant

Sub BuildMain()
'
' Routine Name:  Build Main
' Description:
' Builds the main menu screen at program load
'
Dim Idx, idx2, iStn, iShift, iRow As Integer
 
SetErrModule 99, 0
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' Set Data Foreground color
    lblDataColor.ForeColor = Data_ForeColor
    lblCurrDataColor.ForeColor = Data_ForeColor

' *************
' STATION BOXES
' *************
NrStns = LAST_STN
NrShifts = NR_SHIFT
If ((NrStns > 6) And (NrShifts > 2)) Then NrShifts = 2

'TopOfStations = tbrNavigate.Top + tbrNavigate.Height + 30
'TopToBottomOfStations = 9597
'BottomOfStations = TopOfStations + TopToBottomOfStations
'FreeSpaceAtBottom = pbxBottom.Top - BottomOfStations
'If FreeSpaceAtBottom < 0 Then
'    idx = 1
'End If
TopOfStations = tbrNavigate.Top + tbrNavigate.Height
BottomOfStations = pbxBottom.Top
TopToBottomOfStations = BottomOfStations - TopOfStations
FreeSpaceAtBottom = 0

Select Case NrStns
    Case 1
        StnPerRow = 1
        FirstColLeft = 0
        FirstRowTop = TopOfStations
        StnWidth = frmMainMenu.Width
        StnHeight = (BottomOfStations - TopOfStations)
        StnSpaceHorz = (frmMainMenu.Width - (StnWidth)) / 2
        StnSpaceVert = ((BottomOfStations - TopOfStations) - (StnHeight)) / 2
        CmdLeft = 75
        CmdTop = 60
        CmdHeight = 870
        CmdWidth = StnWidth - 240
        CmdFontSize = cmdLargeFont.FontSize
        Select Case NrShifts
            Case 1
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 240
                ShiftHeight = -210 + (StnHeight - CmdHeight - CmdTop - 180) / NrShifts
                BevelInner = 1
                BevelOuter = 2
                BevelWidth = 2
                BorderWidth = 6
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 4440
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 150
                MsgTop = 150
                MsgHeight = PnlMsgHeight - (2 * MsgTop)
                MsgWidth = PnlMsgWidth - (2 * MsgLeft)
                MsgFontSize = 16
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 150
                ModeTop = 150
                ModeHeight = 1920
                ModeWidth = PnlModeWidth - (2 * ModeLeft)
                ModeFontSize = 16
                PbarLeft = 150
                PbarTop = ModeTop + ModeHeight
                PbarHeight = PnlModeHeight - ModeHeight - 300
                PbarWidth = PnlModeWidth - (2 * PbarLeft)
                CharWidth = 72
                CanLength = 20
                RcpLength = 60
                MsgRows = 16
            Case 2
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 240
                ShiftHeight = -105 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 1
                BevelOuter = 2
                BevelWidth = 2
                BorderWidth = 6
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 3060
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 150
                MsgTop = 150
                MsgHeight = PnlMsgHeight - (2 * MsgTop)
                MsgWidth = PnlMsgWidth - (2 * MsgLeft)
                MsgFontSize = 16
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 150
                ModeTop = 150
                ModeHeight = 390
                ModeWidth = PnlModeWidth - (2 * ModeLeft)
                ModeFontSize = 16
                PbarLeft = 150
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 240
                PbarWidth = PnlModeWidth - (2 * PbarLeft)
                CharWidth = 72
                CanLength = 20
                RcpLength = 60
                MsgRows = 10
            Case 3
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 240
                ShiftHeight = -55 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 1
                BevelOuter = 1
                BevelWidth = 1
                BorderWidth = 4
                PnlMsgLeft = 120
                PnlMsgTop = 120
                PnlMsgHeight = 1690
                PnlMsgWidth = ShiftWidth - 240
                MsgLeft = 120
                MsgTop = 120
                MsgHeight = PnlMsgHeight - (2 * MsgTop)
                MsgWidth = PnlMsgWidth - (2 * MsgLeft)
                MsgFontSize = 12
                PnlModeLeft = 120
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 240
                PnlModeWidth = ShiftWidth - 240
                ModeLeft = 120
                ModeTop = 120
                ModeHeight = 385
                ModeWidth = PnlModeWidth - (2 * ModeLeft)
                ModeFontSize = 16
                PbarLeft = 120
                PbarTop = ModeTop + ModeHeight
                PbarHeight = PnlModeHeight - ModeHeight - 240
                PbarWidth = PnlModeWidth - (2 * PbarLeft)
                CharWidth = 72
                CanLength = 20
                RcpLength = 60
                MsgRows = 5
            Case 4
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 240
                ShiftHeight = -60 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 1395
                PnlMsgWidth = ShiftWidth - 240
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 120
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 60
                ModeHeight = 300
                ModeWidth = PnlModeWidth
                ModeFontSize = 12
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 240
                PbarWidth = PnlModeWidth
                CharWidth = 46
                CanLength = 20
                RcpLength = 40
                MsgRows = 5
        End Select
    Case 2
        StnPerRow = 2
        FirstColLeft = 0
        FirstRowTop = TopOfStations
        StnWidth = frmMainMenu.Width / 2
        StnHeight = (BottomOfStations - TopOfStations)  '(BottomOfStations - TopOfStations) / 2
        StnSpaceHorz = (frmMainMenu.Width - (2 * StnWidth))
        StnSpaceVert = ((BottomOfStations - TopOfStations) - (2 * StnHeight))
        CmdLeft = 75
        CmdTop = 60
        CmdHeight = 870
        CmdWidth = StnWidth - 195       ' was 180
        CmdFontSize = cmdRegFont.FontSize
        Select Case NrShifts
            Case 1
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 180
                ShiftHeight = -195 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 1
                BevelOuter = 2
                BevelWidth = 2
                BorderWidth = 6
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 2980
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 150
                MsgTop = 150
                MsgHeight = PnlMsgHeight - (2 * MsgTop)
                MsgWidth = PnlMsgWidth - (2 * MsgLeft)
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 150
                ModeTop = 150
                ModeHeight = 1920
                ModeWidth = PnlModeWidth - (2 * ModeLeft)
                ModeFontSize = 16
                PbarLeft = 150
                PbarTop = ModeTop + ModeHeight
                PbarHeight = PnlModeHeight - ModeHeight - 300
                PbarWidth = PnlModeWidth - (2 * PbarLeft)
                CharWidth = 72
                CanLength = 20
                RcpLength = 40
                MsgRows = 9
            Case 2
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 195        ' was 180
                ShiftHeight = -105 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)    'was -30
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 1390
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 330
                ModeWidth = PnlModeWidth
                ModeFontSize = 12
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 300
                PbarWidth = PnlModeWidth
                CharWidth = 46
                CanLength = 20
                RcpLength = 40
                MsgRows = 5
            Case 3
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 180
                ShiftHeight = -75 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 1390
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 400
                ModeHeight = 540
                ModeWidth = PnlModeWidth
                ModeFontSize = 16
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 480
                PbarWidth = PnlModeWidth
                CharWidth = 46
                CanLength = 20
                RcpLength = 40
                MsgRows = 5
            Case 4
                FirstShiftLeft = 75
                FirstShiftTop = 935
                ShiftWidth = StnWidth - 180
                ShiftHeight = -60 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 1390
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 60
                ModeHeight = 360
                ModeWidth = PnlModeWidth
                ModeFontSize = 12
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 300
                PbarWidth = PnlModeWidth
                CharWidth = 46
                CanLength = 20
                RcpLength = 40
                MsgRows = 5
        End Select
    Case 3, 4
        StnPerRow = 2
        FirstColLeft = 0
        FirstRowTop = TopOfStations
        StnWidth = frmMainMenu.Width / StnPerRow
        StnHeight = (BottomOfStations - TopOfStations) / 2
        StnSpaceHorz = (frmMainMenu.Width - (StnPerRow * StnWidth))
        StnSpaceVert = ((BottomOfStations - TopOfStations) - (2 * StnHeight))
        CmdLeft = 75
        CmdTop = 60
        CmdHeight = 435
        CmdWidth = StnWidth - 180
        CmdFontSize = cmdRegFont.FontSize
        Select Case NrShifts
            Case 1
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 180
                ShiftHeight = -60 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 1
                BevelOuter = 2
                BevelWidth = 2
                BorderWidth = 6
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 2980
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 150
                MsgTop = 150
                MsgHeight = PnlMsgHeight - (2 * MsgTop)
                MsgWidth = PnlMsgWidth - (2 * MsgLeft)
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 150
                ModeTop = 150
                ModeHeight = 360
                ModeWidth = PnlModeWidth - (2 * ModeLeft)
                ModeFontSize = 12
                PbarLeft = 150
                PbarTop = ModeTop + ModeHeight
'                PbarHeight = 320
                PbarHeight = PnlModeHeight - ModeHeight - 300
                PbarWidth = PnlModeWidth - (2 * PbarLeft)
                CharWidth = 66
                CanLength = 20
                RcpLength = 40
                MsgRows = 9
            Case 2
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 180
                ShiftHeight = -30 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 1390
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 330
                ModeWidth = PnlModeWidth
                ModeFontSize = 12
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 300
                PbarWidth = PnlModeWidth
                CharWidth = 46
                CanLength = 20
                RcpLength = 40
                MsgRows = 5
            Case 3
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 195
                ShiftHeight = -15 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 60
                PnlMsgHeight = 870
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 10
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 120
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 240
                ModeWidth = PnlModeWidth
                ModeFontSize = 10
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 180
                PbarWidth = PnlModeWidth
                CharWidth = 42
                CanLength = 20
                RcpLength = 40
                MsgRows = 4
            Case 4
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 195
                ShiftHeight = -15 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 480
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 10
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 260
                ModeWidth = PnlModeWidth
                ModeFontSize = 10
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 210
                PbarWidth = PnlModeWidth
                CharWidth = 42
                CanLength = 20
                RcpLength = 40
                MsgRows = 3
        End Select
    Case 5, 6
        StnPerRow = 3
        FirstColLeft = 0
        FirstRowTop = TopOfStations
        StnWidth = frmMainMenu.Width / StnPerRow
        StnHeight = (BottomOfStations - TopOfStations) / 2
        StnSpaceHorz = (frmMainMenu.Width - (StnPerRow * StnWidth))
        StnSpaceVert = ((BottomOfStations - TopOfStations) - (2 * StnHeight))
        CmdLeft = 75
        CmdTop = 60
        CmdHeight = 435
        CmdWidth = StnWidth - 180
        CmdFontSize = cmdRegFont.FontSize
        Select Case NrShifts
            Case 1
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 180
                ShiftHeight = (StnHeight - CmdHeight - CmdTop - 180) / NrShifts
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = (690 * 2) + 150 + 120
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = 800 + 690 + 150 + 30 + 60
                PnlModeHeight = 250 + 180 + 145 + 180 + 30
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 250 + 180
                ModeWidth = PnlModeWidth
                ModeFontSize = 12
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 145 + 180 + 30 - 60
                PbarWidth = PnlModeWidth
                CharWidth = 31
                CanLength = 15
                RcpLength = 30
                MsgRows = 6
            Case 2
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 180
                ShiftHeight = (StnHeight - CmdHeight - CmdTop - 195) / NrShifts
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 720
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 10
                PnlModeLeft = 75
                PnlModeTop = 800
                PnlModeHeight = 395
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 250
                ModeWidth = PnlModeWidth
                ModeFontSize = 10
                PbarLeft = 0
                PbarTop = ModeHeight
                PbarHeight = 145
                PbarWidth = PnlModeWidth
                CharWidth = 39
                CanLength = 15
                RcpLength = 36
                MsgRows = 3
            Case 3
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 195
                ShiftHeight = -15 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 60
                PnlMsgHeight = 870
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 10
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 120
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 240
                ModeWidth = PnlModeWidth
                ModeFontSize = 10
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 180
                PbarWidth = PnlModeWidth
                CharWidth = 42
                CanLength = 20
                RcpLength = 40
                MsgRows = 4
            Case 4
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 195
                ShiftHeight = -15 + ((StnHeight - CmdHeight - CmdTop - 180) / NrShifts)
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 480
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 10
                PnlModeLeft = 75
                PnlModeTop = PnlMsgTop + PnlMsgHeight
                PnlModeHeight = ShiftHeight - PnlMsgHeight - 150
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 260
                ModeWidth = PnlModeWidth
                ModeFontSize = 10
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 210
                PbarWidth = PnlModeWidth
                CharWidth = 42
                CanLength = 20
                RcpLength = 40
                MsgRows = 3
            End Select
    Case 7, 8, 9
        StnPerRow = 3
        FirstColLeft = 0
        FirstRowTop = TopOfStations
        StnWidth = frmMainMenu.Width / 3
        StnHeight = (BottomOfStations - TopOfStations) / 3
        StnSpaceHorz = (frmMainMenu.Width - (3 * StnWidth)) / 2
        StnSpaceVert = ((BottomOfStations - TopOfStations) - (3 * StnHeight)) / 2
        CmdLeft = 75
        CmdTop = 60
        CmdHeight = 435
        CmdWidth = StnWidth - 180
        CmdFontSize = cmdRegFont.FontSize
        Select Case NrShifts
            Case 1
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 180
                ShiftHeight = (StnHeight - CmdHeight - CmdTop - 180) / NrShifts
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = (690 * 2) + 150 + 120
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 12
                PnlModeLeft = 75
                PnlModeTop = 800 + 690 + 150 + 30 + 60
                PnlModeHeight = 250 + 180 + 145 + 180 + 30
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 250 + 180
                ModeWidth = PnlModeWidth
                ModeFontSize = 12
                PbarLeft = 0
                PbarTop = ModeTop + ModeHeight
                PbarHeight = 145 + 180 + 30 - 60
                PbarWidth = PnlModeWidth
                CharWidth = 31
                CanLength = 15
                RcpLength = 30
                MsgRows = 6
            Case 2
                FirstShiftLeft = 75
                FirstShiftTop = 500
                ShiftWidth = StnWidth - 180
                ShiftHeight = (StnHeight - CmdHeight - CmdTop - 180) / NrShifts
                BevelInner = 0
                BevelOuter = 0
                BevelWidth = 0
                BorderWidth = 0
                PnlMsgLeft = 75
                PnlMsgTop = 75
                PnlMsgHeight = 720
                PnlMsgWidth = ShiftWidth - 150
                MsgLeft = 0
                MsgTop = 0
                MsgHeight = PnlMsgHeight
                MsgWidth = PnlMsgWidth
                MsgFontSize = 10
                PnlModeLeft = 75
                PnlModeTop = 800
                PnlModeHeight = 395
                PnlModeWidth = ShiftWidth - 150
                ModeLeft = 0
                ModeTop = 0
                ModeHeight = 250
                ModeWidth = PnlModeWidth
                ModeFontSize = 10
                PbarLeft = 0
                PbarTop = ModeHeight
                PbarHeight = 145
                PbarWidth = PnlModeWidth
                CharWidth = 39
                CanLength = 15
                RcpLength = 36
                MsgRows = 3
        End Select

End Select

PaleBkgd = lblLightBkgd.BackColor
'PaleBkgd = pnlStn(0).BackColor

For iStn = 1 To NrStns

    Idx = iStn - 1
    
    Select Case StnPerRow
        Case 1
            iRow = 1
        Case 2
            Select Case iStn
                Case 1, 2
                    iRow = 1
                Case 3, 4
                    iRow = 2
            End Select
        Case 3
            Select Case iStn
                Case 1, 2, 3
                    iRow = 1
                Case 4, 5, 6
                    iRow = 2
                Case 7, 8, 9
                    iRow = 3
            End Select
    End Select
    
    Select Case iRow
        Case 1
            pnlStn(Idx).Left = FirstColLeft + ((iStn - 1) * (StnWidth + StnSpaceHorz))
            pnlStn(Idx).Top = FirstRowTop
        Case 2
            pnlStn(Idx).Left = FirstColLeft + ((iStn - 1 - (1 * StnPerRow)) * (StnWidth + StnSpaceHorz))
            pnlStn(Idx).Top = FirstRowTop + StnHeight + StnSpaceVert
        Case 3
            pnlStn(Idx).Left = FirstColLeft + ((iStn - 1 - (2 * StnPerRow)) * (StnWidth + StnSpaceHorz))
            pnlStn(Idx).Top = FirstRowTop + (2 * (StnHeight + StnSpaceVert))
    End Select
    pnlStn(Idx).Width = StnWidth
    pnlStn(Idx).Height = StnHeight
    pnlStn(Idx).ToolTipText = ""
    pnlStn(Idx).Visible = IIf(iStn > NrStns, False, True)
       
    cmdStation(Idx).Left = CmdLeft
    cmdStation(Idx).Top = CmdTop
    cmdStation(Idx).Width = CmdWidth
    cmdStation(Idx).Height = CmdHeight
    cmdStation(Idx).ForeColor = Titles_ForeColor
    cmdStation(Idx).FontSize = CmdFontSize
'    cmdStation(idx).Caption = "Station " & Format(iStn, "0")
    cmdStation(Idx).Caption = STN_INFO(iStn).desc
    cmdStation(Idx).ToolTipText = "Click for Station Details"
    cmdStation(Idx).Visible = True
       
    For iShift = 1 To NrShifts
           
        Select Case iShift
            Case 1, 2
                idx2 = (2 * (iStn - 1)) + (iShift - 1)
            Case 3, 4
                idx2 = 18 + ((2 * (iStn - 1)) + (iShift - 3))
        End Select

        pnlShift(idx2).Left = FirstShiftLeft
        pnlShift(idx2).Top = FirstShiftTop + (ShiftHeight * (iShift - 1))
        pnlShift(idx2).Width = ShiftWidth
        pnlShift(idx2).Height = ShiftHeight
        pnlShift(idx2).BackColor = pnlStn(0).BackColor
        pnlShift(idx2).Visible = True
        
        pnlMsg(idx2).Left = PnlMsgLeft
        pnlMsg(idx2).Top = PnlMsgTop
        pnlMsg(idx2).Width = PnlMsgWidth
        pnlMsg(idx2).Height = PnlMsgHeight
        pnlMsg(idx2).BackColor = pnlStn(0).BackColor
        pnlMsg(idx2).BevelInner = BevelInner
        pnlMsg(idx2).BevelOuter = BevelOuter
        pnlMsg(idx2).BevelWidth = BevelWidth
        pnlMsg(idx2).BorderWidth = BorderWidth
        pnlMsg(idx2).Caption = ""
        pnlMsg(idx2).Visible = True
        
        lblMessage(idx2).Left = MsgLeft
        lblMessage(idx2).Top = MsgTop
        lblMessage(idx2).Width = MsgWidth
        lblMessage(idx2).Height = MsgHeight
        lblMessage(idx2).Font = lblMsgMode.Font
        lblMessage(idx2).FontSize = MsgFontSize
        lblMessage(idx2).FontBold = True
        lblMessage(idx2).Alignment = 0          ' Left Justify
        lblMessage(idx2).ToolTipText = "Click for Station Details"
        lblMessage(idx2).BackStyle = 1          ' Opaque
        lblMessage(idx2).Visible = True
        
        pnlMode(idx2).Left = PnlModeLeft
        pnlMode(idx2).Top = PnlModeTop
        pnlMode(idx2).Width = PnlModeWidth
        pnlMode(idx2).Height = PnlModeHeight
        pnlMode(idx2).BackColor = pnlShift(0).BackColor
        pnlMode(idx2).BevelInner = BevelInner
        pnlMode(idx2).BevelOuter = BevelOuter
        pnlMode(idx2).BevelWidth = BevelWidth
        pnlMode(idx2).BorderWidth = BorderWidth
        pnlMode(idx2).Caption = ""
        pnlMode(idx2).Visible = True
        
        lblMode(idx2).Left = ModeLeft
        lblMode(idx2).Top = ModeTop
        lblMode(idx2).Width = ModeWidth
        lblMode(idx2).Height = ModeHeight
        lblMode(idx2).Font = lblMsgMode.Font
        lblMode(idx2).FontSize = ModeFontSize
        lblMode(idx2).FontBold = True
        lblMode(idx2).BackStyle = 1             ' Opaque
        lblMode(idx2).Visible = True
        
        pbarActual(idx2).Left = PbarLeft
        pbarActual(idx2).Top = PbarTop
        pbarActual(idx2).Width = PbarWidth
        pbarActual(idx2).Height = PbarHeight
        pbarActual(idx2).Visible = True

    Next iShift
Next iStn

    ' Background
    AutoRedraw = -1   ' Turn on AutoRedraw.
    frmMainMenu.BackColor = vbButtonFace
    Common_BackColor = frmMainMenu.Point(90, 90)
    frmMainMenu.BackColor = Gainsboro

ScreenIsBuilt = True

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

Sub UpdateNavigateBtns()

'
' Routine Name:  UpdateNavigateBtns
' Description:
' Updates the Navigate toolbar buttons
'
Dim iKeyCount As Integer
 
SetErrModule 99, 10101
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
        If CheckPass("N", False) And (USINGREMCANLOAD Or USINGTOMCANLOAD) Then
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
            mnuSysdef.Visible = True
        Else
            tbrNavigate.Buttons("sysdef").Visible = False
            tbrNavigate.Buttons("sysdef").ToolTipText = ""
            mnuSysdef.Visible = False
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
            mnuIOMonitor.Enabled = True
        Else
            mnuIOMonitor.Enabled = False
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

Sub BuildStatusBars()

'
' Routine Name:  BuildStatusBars
' Description:
' Builds the main menu screen Status Bars at program load
'
 
SetErrModule 99, 4
If UseLocalErrorHandler Then On Error GoTo localhandler


    ' STATUS BAR #1
    ' Alarm Buttons
    pnlAlarms.Left = 0

ChgErrModule 99, 5
    Select Case LocalPagControl.Type

        Case pagNone, pagAlone
ChgErrModule 99, 9
            pnlAlarms.Width = pnlAlarms.Width - (pnlPAcomm.Width + pnlMix.Width)
            
ChgErrModule 99, 8
            pnlPAcomm.Top = OutOfSight

        Case pagMaster, pagClient
            pnlAlarms.Width = pnlAlarms.Width - (pnlMix.Width)
            pnlPAcomm.Top = 75
            pnlPAcomm.BackColor = IIf(PaComm_Flag, Good_ForeColor, MEDYELLOW)
            pnlPAcomm.ToolTipText = IIf(PaComm_Flag, "PurgeAirSystem Communications OK", "PurgeAirSystem Communications Not Online")

    End Select

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
    pnlPurgeAir.Width = frmMainMenu.Width - pnlPurgeAir.Left - 150
    pnlPurgeAir.Top = pnlAlarms.Top
    pnlPurgeAir.Height = pnlAlarms.Height
    pnlEstop.BackColor = IIf(Alm_Estop, Alarm_ForeColor, Good_ForeColor)
    pnlEstop.ToolTipText = IIf(Alm_Estop, "ESTOP Pressed", "ESTOP OK")
    pnlFlow.BackColor = IIf(Alm_Flow, Alarm_ForeColor, Good_ForeColor)
    pnlFlow.ToolTipText = IIf(Alm_Flow, "Loss of Exhaust Flow", "Exhaust Flow OK")
    pnlBtn20.BackColor = IIf(Alm_Btn20, Alarm_ForeColor, Good_ForeColor)
    pnlBtn20.ToolTipText = IIf(Alm_Btn20, "LEL20 Alarm", "LEL20 OK")
    
    If Alm_Doors Then
        pnlDoors.BackColor = Alarm_ForeColor
        pnlDoors.ToolTipText = "Door Open Alarm"
    ElseIf Not Com_DIO(icDoorSw).Value And USINGDOOROPEN Then
        pnlDoors.BackColor = MEDYELLOW
        pnlDoors.ToolTipText = "One or more Doors are Open"
    Else
        pnlDoors.BackColor = Good_ForeColor
        pnlDoors.ToolTipText = "Doors Closed"
    End If
    
    pnlComms.BackColor = IIf(IoComm_Flag, Good_ForeColor, MEDYELLOW)
    pnlComms.ToolTipText = IIf(IoComm_Flag, "IO Communications OK", "IO Communications Alarm")
    pnlMix.Top = OutOfSight
    pnlMessage.Caption = " "
    pnlPurgeAir.Caption = " "
        
    
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

Public Sub UpdateMain()
'
' Routine Name:  Update Main
' Author:        Analytical Process Programmer / APS
' Description:
' Updates the main menu screen with current status
'

Dim Idx, idx2, gap As Integer
Dim iStn, iShift, iCycle As Integer
Dim sDbf, sRcp, sCan, sCyc, sSft, sMsg, sMoistUnits, sTempUnits As String
Dim fillVal As Single
Dim temptime As Date
Dim tempSec, pauseSec As Long
Dim tempMin, delayMin As Long
Dim PurgeAir_Warning As Boolean

 
SetErrModule 99, 100
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' Screen Built ?
    If Not ScreenIsBuilt Then BuildMain
    
    ' STATUS BAR
    ' Alarm Buttons
    pnlEstop.BackColor = IIf(Alm_Estop, Alarm_ForeColor, Good_ForeColor)
    pnlEstop.ToolTipText = IIf(Alm_Estop, "ESTOP Pressed", "ESTOP OK")
    pnlFlow.BackColor = IIf(Alm_Flow, Alarm_ForeColor, Good_ForeColor)
    pnlFlow.ToolTipText = IIf(Alm_Flow, "Loss of Exhaust Flow", "Exhaust Flow OK")
    pnlBtn20.BackColor = IIf(Alm_Btn20, Alarm_ForeColor, Good_ForeColor)
    pnlBtn20.ToolTipText = IIf(Alm_Btn20, "LEL20 Alarm", "LEL20 OK")
    If Alm_Doors Then
        pnlDoors.BackColor = Alarm_ForeColor
        pnlDoors.ToolTipText = "Door Open Alarm"
    ElseIf USINGDOOROPEN Then
        If Not Com_DIO(icDoorSw).Value Then
            pnlDoors.BackColor = MEDYELLOW
            pnlDoors.ToolTipText = "One or more Doors are Open"
        Else
            pnlDoors.BackColor = Good_ForeColor
            pnlDoors.ToolTipText = "Doors Closed"
        End If
    ElseIf Not USINGDOOROPEN Then
        pnlDoors.BackColor = Good_ForeColor
        pnlDoors.ToolTipText = "Doors Not Monitored"
    End If
    pnlComms.BackColor = IIf(IoComm_Flag, Good_ForeColor, MEDYELLOW)
    pnlComms.ToolTipText = IIf(IoComm_Flag, "IO Communications OK", "IO Communications Alarm")
    Select Case LocalPagControl.Type
        Case pagNone, pagAlone
            ' nothing to update
        Case pagMaster, pagClient
            pnlPAcomm.BackColor = IIf(PaComm_Flag, Good_ForeColor, MEDYELLOW)
            pnlPAcomm.ToolTipText = IIf(PaComm_Flag, "PurgeAirSystem Communications OK", "PurgeAirSystem Communications Not Online")
    End Select

    ' Message Box
    If (Pause_Alarm = SYSTEMPAUSED) Then
        ' System Paused
        SysMessage_BackColor = Alarm_ForeColor
        SysMessage_ForeColor = White
        SysMessage_Text = "SYSTEM PAUSED"
        SysMessage_Tooltip = "System is Paused for a Major Alarm"
    ElseIf (USINGUPS <> 0 And Alm_Ups_Count > 0) Then
        ' UPS Shutdown Countdown is Active
        SysMessage_BackColor = PALERED
        SysMessage_ForeColor = Black
        If SysConfig.UPSOpenDelay = Alm_Ups_Count Then
            ' 1 minute (or less) to go
            sMsg = "UPS shutdown in 1 minute"
        Else
            ' more than 1 minute  to go
            sMsg = "UPS shutdown in " & Format((SysConfig.UPSOpenDelay - Alm_Ups_Count + 1), "##0") & " minutes"
        End If
        SysMessage_Text = sMsg
        SysMessage_Tooltip = "Ups Shutdown Countdown is Active"
    ElseIf (USINGDOOROPEN And Not Alm_Doors And Alm_Doors_Count > 0) Then
        ' Door(s) Open Countdown is Active
        SysMessage_BackColor = EntryInvalid_BackColor
        SysMessage_ForeColor = Alarm_ForeColor
        If SysConfig.DoorOpenDelay = Alm_Doors_Count Then
            ' 1 minute (or less) to go
            sMsg = "All Jobs will be Paused in 1 minute"
        Else
            ' more than 1 minute  to go
            sMsg = "All Jobs will be Paused in " & Format((SysConfig.DoorOpenDelay - Alm_Doors_Count + 1), "##0") & " minutes"
        End If
        SysMessage_Text = sMsg
        SysMessage_Tooltip = "Door Open Shutdown Countdown is Active"
    ElseIf (USINGSYSTEMVACSW And Alm_SystemVacSw) Then
        ' System Vacuum Switch is off
        SysMessage_BackColor = PALERED
        SysMessage_ForeColor = Alarm_ForeColor
        SysMessage_Text = "System Vacuum Switch is OFF"
        SysMessage_Tooltip = "Check Purge Air Supply"
    ElseIf MaintMode Then
        ' Maintenance Mode is Active
        SysMessage_BackColor = EntryInvalid_BackColor
        SysMessage_ForeColor = DKORANGE
            sMsg = "Maintenance Mode is Active"
        SysMessage_Text = sMsg
        SysMessage_Tooltip = "Maintenance Mode Switch DI is True"
    ElseIf (systemhasBUTANE And ButaneSupply.WarningActive) Then
        ' Low Butane Warning
        SysMessage_BackColor = EntryInvalid_BackColor
        SysMessage_ForeColor = DKORANGE
        SysMessage_Text = "Low Butane Warning"
        SysMessage_Tooltip = "Change Butane Cylinder Soon"
    ElseIf (USINGERRORMSGBYPASS And ShortTermErrorCounter > 1) Then
        ' Multiple Recent Program Errors Warning
        SysMessage_BackColor = EntryInvalid_BackColor
        SysMessage_ForeColor = DKORANGE
        SysMessage_Text = "Program Errors - Check Event Log"
        SysMessage_Tooltip = "There have been recent program errors. Please check the Events Log"
    ElseIf (USINGERRORMSGBYPASS And UnreadProgramErrorMessage) Then
        ' Recent Program Error Warning
        SysMessage_BackColor = EntryInvalid_BackColor
        SysMessage_ForeColor = DKORANGE
        SysMessage_Text = "Program Error - Check Event Log"
        SysMessage_Tooltip = "There has been a recent program error. Please check the Events Log"
    ElseIf (USINGSIMULATION And ((Not IoComOn) Or (Not SclComOn))) Then
        ' Simulation is Active
        SysMessage_BackColor = pnlAlarms.BackColor
        SysMessage_ForeColor = Warning_ForeColor
        If ((Not IoComOn) And (Not SclComOn)) Then
            SysMessage_Text = "SIMULATED I/O AND SCALES"
            SysMessage_Tooltip = "I/O and Scale Values are Simulated"
        ElseIf (Not IoComOn) Then
            SysMessage_Text = "SIMULATED I/O"
            SysMessage_Tooltip = "I/O Values are Simulated"
        ElseIf (Not SclComOn) Then
            SysMessage_Text = "SIMULATED SCALES"
            SysMessage_Tooltip = "Scale Values are Simulated"
        End If
    ElseIf InStr(USINGRELEASEDATE, "Release Version") = 0 Then
        ' Debug Version
        SysMessage_BackColor = pnlAlarms.BackColor
        SysMessage_ForeColor = Warning_ForeColor
        SysMessage_Text = "DEBUG VERSION"
        SysMessage_Tooltip = "This is a debug version of the program."
    Else
        SysMessage_BackColor = pnlAlarms.BackColor
        SysMessage_ForeColor = Black
        SysMessage_Text = " "
        SysMessage_Tooltip = " "
    End If
    With pnlMessage
        .FontSize = 10
        .BackColor = SysMessage_BackColor
        .ForeColor = SysMessage_ForeColor
        .Caption = SysMessage_Text
        .ToolTipText = SysMessage_Tooltip
    End With
    

    PurgeAir_Warning = False

    For iStn = 1 To NrStns
    
        Idx = iStn - 1
        
        cmdStation(Idx).Caption = STN_INFO(iStn).desc
        
        For iShift = 1 To NrShifts
        
            ChgErrModule 99, 101
            Select Case iShift
                Case 1, 2
                    idx2 = (2 * (iStn - 1)) + (iShift - 1)
                Case 3, 4
                    idx2 = 18 + ((2 * (iStn - 1)) + (iShift - 3))
            End Select
  
            If StationControl(iStn, iShift).Mode = VBIDLE _
              Or StationControl(iStn, iShift).Mode = VBIDLEWAITING _
              Or StationControl(iStn, iShift).Mode = VBCOMPLETE Then
                iCycle = StationControl(iStn, iShift).CompletedCycles
            Else
                iCycle = StationControl(iStn, iShift).CurrCycle
            End If
        
            Select Case StationControl(iStn, iShift).Mode
                Case VBLEAK
                    ' Leak Check - add leak check phase description
                    sMsg = ModeDescShort(VBLEAK) & " - " & LeakPhaseDesc(LeakCheckControl.Phase)
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).ToolTipText = ""
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBLOAD
                    ' Loading or Waiting for Scales to Settle?
                    If LoadControl(iStn, iShift).Phase = LoadPause Then
                        ' Waiting for Scales to Settle
                        sMsg = " Load Settling for "
                        sMsg = sMsg & Format(StationConfig(iStn, iShift).LoadSettleTime, "##0.0#")
                        sMsg = sMsg & LoadTypeDesc2(LOADBYTIME)
                        sMsg = sMsg
                        sMsg = sMsg & LoadTypeDesc3(LOADBYTIME)
                    Else
                        ' Load - add load method description
                        Select Case StationRecipe(iStn, iShift).Load_MethodSave
                            Case NOLOAD
                                sMsg = LoadTypeDesc(NOLOAD)
                                sMsg = sMsg
                                sMsg = sMsg & LoadTypeDesc2(NOLOAD)
                                sMsg = sMsg
                                sMsg = sMsg & LoadTypeDesc3(NOLOAD)
                            Case LOADBYTIME
                                sMsg = LoadTypeDesc(LOADBYTIME)
                                sMsg = sMsg & Format(StationRecipe(iStn, iShift).Load_Time, "##0")
                                sMsg = sMsg & LoadTypeDesc2(LOADBYTIME)
                                sMsg = sMsg
                                sMsg = sMsg & LoadTypeDesc3(LOADBYTIME)
                            Case LOADBYWC
                                sMsg = LoadTypeDesc(LOADBYWC)
                                sMsg = sMsg & Format(StationRecipe(iStn, iShift).WC_MultSave, "##0.#")
                                sMsg = sMsg & LoadTypeDesc2(LOADBYWC)
                                sMsg = sMsg & Format(StationRecipe(iStn, iShift).EPAFill, "##0")
                                sMsg = sMsg & LoadTypeDesc3(LOADBYWC)
                            Case LOADBYWEIGHT
                                sMsg = LoadTypeDesc(LOADBYWEIGHT)
                                If Int(StationRecipe(iStn, iShift).Load_Wt) = StationRecipe(iStn, iShift).Load_Wt Then
                                    ' no digits to the right of the decimal point
                                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).Load_Wt, "##0")
                                Else
                                    ' digit(s) to the right of the decimal point
                                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).Load_Wt, "##0.##")
                                End If
                                sMsg = sMsg & LoadTypeDesc2(LOADBYWEIGHT)
                                sMsg = sMsg
                                sMsg = sMsg & LoadTypeDesc3(LOADBYWEIGHT)
                            Case LOADBYBREAKTHRU
                                sMsg = LoadTypeDesc(LOADBYBREAKTHRU)
                                If Int(StationRecipe(iStn, iShift).LoadBreakthrough) = StationRecipe(iStn, iShift).LoadBreakthrough Then
                                    ' no digits to the right of the decimal point
                                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).LoadBreakthrough, "##0")
                                Else
                                    ' digit(s) to the right of the decimal point
                                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).LoadBreakthrough, "##0.##")
                                End If
                                sMsg = sMsg & LoadTypeDesc2(LOADBYBREAKTHRU)
                                sMsg = sMsg
                                sMsg = sMsg & LoadTypeDesc3(LOADBYBREAKTHRU)
                            Case LOADBYFID
                                sMsg = LoadTypeDesc(LOADBYFID)
                                sMsg = sMsg & Format(StationRecipe(iStn, iShift).Load_Time, "#####0")
                                sMsg = sMsg & LoadTypeDesc2(LOADBYFID)
                                sMsg = sMsg
                                sMsg = sMsg & LoadTypeDesc3(LOADBYFID)
                        End Select
                    End If
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).ToolTipText = "Click to ReviewData"
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBPURGE
                    ' Purging or Waiting for Scales to Settle?
                    If PurgeControl(iStn, iShift).Phase = PurgePause Then
                        ' Waiting for Scales to Settle
                        sMsg = " Purge Settling for "
                        sMsg = sMsg & Format(StationConfig(iStn, iShift).PurgeSettleTime, "##0.0#")
                        sMsg = sMsg & LoadTypeDesc2(PURGEBYTIME)
                        sMsg = sMsg
                        sMsg = sMsg & LoadTypeDesc3(PURGEBYTIME)
                    Else
                        ' Purge - add Purge method description
                        Select Case StationRecipe(iStn, iShift).Purge_Method
                            Case NOPURGE
                                sMsg = "No Purge"
                            Case PURGEBYTIME
                                sMsg = ModeDescShort(VBPURGE) & " for " & StationRecipe(iStn, iShift).Purge_Time & " Minute"
                                If StationRecipe(iStn, iShift).Purge_Time > 1 Then sMsg = sMsg & "s"
                            Case PURGEBYLITERS
                                sMsg = ModeDescShort(VBPURGE) & " " & StationRecipe(iStn, iShift).Purge_Liters & " liter"
                                If StationRecipe(iStn, iShift).Purge_Liters <> 1 Then sMsg = sMsg & "s"
                            Case PURGEBYVOLUME
                                sMsg = ModeDescShort(VBPURGE) & " " & StationRecipe(iStn, iShift).Purge_Can_Vol & " Canister Volume"
                                If StationRecipe(iStn, iShift).Purge_Can_Vol <> 1 Then sMsg = sMsg & "s"
                            Case PURGEAUXONLY
                                sMsg = ModeDescShort(VBPURGE) & " Aux Can for " & StationRecipe(iStn, iShift).Purge_AuxTime & " Minute"
                                If StationRecipe(iStn, iShift).Purge_AuxTime > 1 Then sMsg = sMsg & "s"
                            Case PURGEBYPROFILE
                                sMsg = ModeDescShort(VBPURGE) & " " & " by Profile"
                            Case PURGEBYWC
                                sMsg = ModeDescShort(VBPURGE) & " " & StationRecipe(iStn, iShift).Purge_TargetWC & " % of Work Cap"
                            Case PURGETOTARGET
                                sMsg = ModeDescShort(VBPURGE) & " " & "to " & StationRecipe(iStn, iShift).Purge_TargetWeight & " grams"
                            Case PURGETOUNDOLOAD
                                sMsg = ModeDescShort(VBPURGE) & " to " & " Undo Load"
                        End Select
                    End If
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).ToolTipText = "Click to ReviewData"
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBPOSTLEAK
                    ' Post LeakCheck Pause
                    sMsg = ModeDescShort(VBPOSTLEAK)
                    sMsg = sMsg & " for "
                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).PauseLeakTime, "##0.0#")
                    sMsg = sMsg & LoadTypeDesc2(LOADBYTIME)
                    sMsg = sMsg
                    sMsg = sMsg & LoadTypeDesc3(LOADBYTIME)
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBPOSTLOAD
                    ' Post Load Pause
                    sMsg = ModeDescShort(VBPOSTLOAD)
                    sMsg = sMsg & " for "
                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).PauseLoadTime, "##0.0#")
                    sMsg = sMsg & LoadTypeDesc2(LOADBYTIME)
                    sMsg = sMsg
                    sMsg = sMsg & LoadTypeDesc3(LOADBYTIME)
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBPOSTPURGE
                    ' Post Purge Pause
                    sMsg = ModeDescShort(VBPOSTPURGE)
                    sMsg = sMsg & " for "
                    sMsg = sMsg & Format(StationRecipe(iStn, iShift).PausePurgeTime, "##0.0#")
                    sMsg = sMsg & LoadTypeDesc2(PURGEBYTIME)
                    sMsg = sMsg
                    sMsg = sMsg & LoadTypeDesc3(PURGEBYTIME)
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBSCALEWAIT
                    ' Waiting for Scale(s) - add which scale(s)
                    sMsg = ModeDescShort(VBSCALEWAIT)
                    If StationRecipe(iStn, iShift).UsePriScale And StationRecipe(iStn, iShift).UseAuxScale Then
                        ' Using Two Scales
                        sMsg = sMsg & "s "
                        ' Scales in use ?
                        If Scale_In_Use(StationRecipe(iStn, iShift).PriScaleNo) And Scale_In_Use(StationRecipe(iStn, iShift).AuxScaleNo) Then
                            ' Both Scales in use
                            sMsg = sMsg & Format(StationRecipe(iStn, iShift).PriScaleNo, "#0") & " && " & Format(StationRecipe(iStn, iShift).AuxScaleNo, "#0")
                        ElseIf Scale_In_Use(StationRecipe(iStn, iShift).PriScaleNo) Then
                            ' Primary Scale in use
                            sMsg = sMsg & Format(StationRecipe(iStn, iShift).PriScaleNo, "#0")
                        ElseIf Scale_In_Use(StationRecipe(iStn, iShift).AuxScaleNo) Then
                            ' Aux Scale in use
                            sMsg = sMsg & Format(StationRecipe(iStn, iShift).AuxScaleNo, "#0")
                        End If
                    ElseIf StationRecipe(iStn, iShift).UsePriScale Then
                        ' Using Only Primary Scale
                        sMsg = sMsg & " "
                        If Scale_In_Use(StationRecipe(iStn, iShift).PriScaleNo) Then sMsg = sMsg & Format(StationRecipe(iStn, iShift).PriScaleNo, "#0")
                    ElseIf StationRecipe(iStn, iShift).UseAuxScale Then
                        ' Using Only Aux Scale
                        sMsg = sMsg & " "
                        If Scale_In_Use(StationRecipe(iStn, iShift).AuxScaleNo) Then sMsg = sMsg & Format(StationRecipe(iStn, iShift).AuxScaleNo, "#0")
                    End If
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).ToolTipText = ""
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBSTARTWAIT
                    ' Delayed Start - add how long
                    sMsg = StartTypeDesc(StationRecipe(iStn, iShift).StartMethod)
                    ' Which Method of Delay ?
                    Select Case StationRecipe(iStn, iShift).StartMethod
                        Case STARTNOW
                            sMsg = sMsg & StartTypeDesc2(STARTNOW)
                        Case STARTDELAYED
                            sMsg = sMsg & Format(StationRecipe(iStn, iShift).StartDelay, "##0")
                            sMsg = sMsg & StartTypeDesc2(STARTDELAYED)
                        Case STARTATDATE
                            sMsg = sMsg & StartTypeDesc2(STARTATDATE)
                            sMsg = sMsg & Format(StationRecipe(iStn, iShift).StartDate, "D MMM YYYY   h:mm")
                    End Select
                    lblMode(idx2).Caption = sMsg
                    lblMode(idx2).ToolTipText = ""
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                Case VBIDLE
                    Select Case NrShifts
                        Case 1
'                           lblMode(idx2).Caption = "Station " & Format(iStn, "0") & " is " & ModeDescShort(StationControl(iStn, iShift).Mode)
                           lblMode(idx2).Caption = "Station is " & ModeDescShort(StationControl(iStn, iShift).Mode)
                        Case 2, 3, 4
                            lblMode(idx2).Caption = "Shift " & Format(iShift, "0") & " is " & ModeDescShort(StationControl(iStn, iShift).Mode)
                    End Select
                    lblMode(idx2).ToolTipText = ""
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = DKGRAY
                Case Else
                    lblMode(idx2).Caption = ModeDescShort(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ToolTipText = ""
                    lblMode(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    lblMode(idx2).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            End Select
            If MsgRows = 16 Then lblMode(idx2).Caption = vbCrLf & vbCrLf & lblMode(idx2).Caption
        
            ChgErrModule 99, 102
            
            If StationControl(iStn, iShift).DBFile = "" Then
                Select Case MsgRows
                    Case 3
'                        sDbf = "no open db file"
                        sDbf = "no active job"
                    Case Else
'                        sDbf = Space(12) & "no open db file"
                        sDbf = Space(14) & "no active job"
                End Select
                Select Case StationControl(iStn, iShift).Mode
                    Case VBIDLE
'                        pnlShift(idx2).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        lblMessage(idx2).BackColor = pnlStn(Idx).BackColor
                        lblMessage(idx2).ForeColor = MEDGRAY
                        pbarActual(idx2).FloodPercent = 100
                    Case VBCOURSEPAUSE
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        tempMin = CLng(StationSequence(iStn, iShift).CourseData(StationControl(iStn, iShift).Course).PauseDuration)
                        temptime = StationSequence(iStn, iShift).CourseData(StationControl(iStn, iShift).Course).DtsStart + TimeSerial(0, CInt(tempMin), 0) - Now()
                        tempSec = CLng((3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime))
                        pauseSec = CLng(60# * tempMin)
                        fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).BackColor = PaleBkgd
                        pbarActual(idx2).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case Else
                        lblMessage(idx2).BackColor = pnlStn(Idx).BackColor
                        lblMessage(idx2).ForeColor = DKGRAY
                        pbarActual(idx2).FloodPercent = 0
                End Select
            Else
                Select Case MsgRows
                    Case 3
'                        sDbf = "db file " & Mid(StationControl(iStn, iShift).DBFile, (Len(StationControl(iStn, iShift).DBFile) - 10), 7)
                        sDbf = "job number " & StationControl(iStn, iShift).Job_Number
                    Case Else
'                        sDbf = "  db file:  " & Mid(StationControl(iStn, iShift).DBFile, (Len(StationControl(iStn, iShift).DBFile) - 10), 7)
                        sDbf = "  job number:   " & StationControl(iStn, iShift).Job_Number
                End Select
                Select Case StationControl(iStn, iShift).Mode
                    Case VBLEAK
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        Select Case LeakCheckControl.Phase
                            Case 0
                                fillVal = 0
                            Case 1
                                If StationControl(iStn, iShift).Target > 0 Then
                                    fillVal = 100 * (StationControl(iStn, iShift).Actual / StationControl(iStn, iShift).Target)
                                Else
                                    fillVal = 0
                                End If
                            Case 2
                                fillVal = (100 * (DateDiff("s", StationControl(iStn, iShift).Mode_StartDts, Now()) / SysConfig.LCTime))
                        End Select
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case VBLOAD
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        If LoadControl(iStn, iShift).Phase = LoadPause Then
                            ' Waiting for Scales to Settle
                            temptime = LoadControl(iStn, iShift).PhaseDts - Now()
                            tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                            pauseSec = CLng(60 * StationConfig(iStn, iShift).LoadSettleTime)
                            If pauseSec = 0 Then
                                ' duration = zero seconds
                                fillVal = 0
                            Else
                                ' normal duration
                                fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                            End If
                        Else
                            ' Loading
                            If StationControl(iStn, iShift).Target > 0 Then
                                fillVal = 100 * (StationControl(iStn, iShift).Actual / StationControl(iStn, iShift).Target)
                            Else
                                fillVal = 0
                            End If
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = "Click to ReviewData"
                    Case VBPURGE
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        If PurgeControl(iStn, iShift).Phase = PurgePause Then
                            ' Waiting for Scales to Settle
                            temptime = PurgeControl(iStn, iShift).PhaseDts - Now()
                            tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                            pauseSec = CLng(60 * StationConfig(iStn, iShift).PurgeSettleTime)
                            If pauseSec = 0 Then
                                ' duration = zero seconds
                                fillVal = 0
                            Else
                                ' normal duration
                                fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                            End If
                        Else
                            ' Purging
                            If StationControl(iStn, iShift).Target > 0 Then
                                fillVal = 100 * (StationControl(iStn, iShift).Actual / StationControl(iStn, iShift).Target)
                            Else
                                fillVal = 0
                            End If
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = "Click to ReviewData"
                    Case VBPOSTLOAD
                        ChgErrModule 99, 10221
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        temptime = StationControl(iStn, iShift).End_Time - Now()
                        tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                        pauseSec = CLng(60 * StationRecipe(iStn, iShift).PauseLoadTime)
                        If pauseSec = 0 Then
                            ' duration = zero seconds
                            fillVal = 0
                        Else
                            ' normal duration
                            fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case VBPOSTPURGE
                        ChgErrModule 99, 10222
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        temptime = StationControl(iStn, iShift).End_Time - Now()
                        tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                        pauseSec = CLng(60 * StationRecipe(iStn, iShift).PausePurgeTime)
                        If pauseSec = 0 Then
                            ' duration = zero seconds
                            fillVal = 0
                        Else
                            ' normal duration
                            fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case VBPOSTLEAK
                        ChgErrModule 99, 11221
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        temptime = StationControl(iStn, iShift).End_Time - Now()
                        tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
                        pauseSec = CLng(60 * StationRecipe(iStn, iShift).PauseLeakTime)
                        If pauseSec = 0 Then
                            ' duration = zero seconds
                            fillVal = 0
                        Else
                            ' normal duration
                            fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case VBPRELOAD
                        ChgErrModule 99, 10222
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        temptime = StationControl(iStn, iShift).End_Time - Now()
                        tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
'    tempSec = StationConfig(station, Shift).NitrogenPurgeTime Mod 60
'    tempMin = CInt((StationConfig(station, Shift).NitrogenPurgeTime - tempSec) / 60)
                        pauseSec = CLng(StationConfig(iStn, iShift).NitrogenPurgeTime)
                        If pauseSec = 0 Then
                            ' duration = zero seconds
                            fillVal = 0
                        Else
                            ' normal duration
                            fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case VBPURGEWAIT
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        If PRG_INFO(STN_INFO(iStn).AspiratorNum).Ready Then
                            ' waiting for another station to finish purge or 5 minutes to expire
                            tempSec = CLng(IIf(Now > LastPurgeStart, DateDiff("s", LastPurgeStart, Now), 0))
                            pauseSec = CLng(300)
                            fillVal = 100 * (tempSec / pauseSec)
                        Else
                            fillVal = 0
                        End If
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).BackColor = PaleBkgd
                        pbarActual(idx2).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""

                    Case VBSTARTWAIT
                        ChgErrModule 99, 10223
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        Select Case StationRecipe(iStn, iShift).StartMethod
                            Case STARTNOW
                                fillVal = 100
                            Case STARTDELAYED, STARTATDATE
                                pauseSec = CLng(StationControl(iStn, iShift).DelaySeconds)
                                tempSec = CLng(StationControl(iStn, iShift).TestTimer)
                                fillVal = 100 * (tempSec / pauseSec)
                        End Select
                        If fillVal < 0 Then fillVal = 0
                        If fillVal > 100 Then fillVal = 100
                        pbarActual(idx2).FloodPercent = fillVal
                        pbarActual(idx2).ToolTipText = ""
                    Case VBLEAKWAIT
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                        pbarActual(idx2).FloodPercent = 0
                        pbarActual(idx2).ToolTipText = ""
                    Case VBSCALEWAIT, VBSHIFTWAIT
                        lblMessage(idx2).BackColor = PaleBkgd
                        lblMessage(idx2).ForeColor = DKGRAY
                        pbarActual(idx2).FloodPercent = 0
                        pbarActual(idx2).ToolTipText = ""
                    Case VBLEAKERROR, VBPAUSEALARM
                        lblMessage(idx2).BackColor = MEDRED
                        lblMessage(idx2).ForeColor = Black
                        pbarActual(idx2).FloodPercent = 0
                        pbarActual(idx2).ToolTipText = ""
                    Case VBPAUSEOOT
                        lblMessage(idx2).BackColor = PALERED
                        lblMessage(idx2).ForeColor = Black
                        pbarActual(idx2).BackColor = lblMessage(idx2).BackColor
                        pbarActual(idx2).FloodPercent = 0
                        pbarActual(idx2).ToolTipText = ""
                    Case Else
                        lblMessage(idx2).BackColor = pnlStn(Idx).BackColor
                        lblMessage(idx2).ForeColor = MEDGRAY
                        pbarActual(idx2).FloodPercent = 0
                        pbarActual(idx2).ToolTipText = ""
                End Select
            End If
            pbarActual(idx2).BackColor = lblMessage(idx2).BackColor
            pbarActual(idx2).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
            ChgErrModule 99, 10229
            
            Select Case MsgRows
                Case 3
                    ' 9 stations, 2 shifts; 4 stations, 4 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
                    sSft = " shift " & Format(iShift, "0")
                    gap = CharWidth - Len(sSft) - Len(sCan) - 1
                    sMsg = sSft & Space(gap) & sCan & vbCrLf
                    gap = ((CharWidth - Len(sRcp)) / 2)
                    sMsg = sMsg & Space(1 + gap) & sRcp & vbCrLf
                    gap = CharWidth - Len(sCyc) - Len(sDbf)
                    sMsg = sMsg & sCyc & Space(gap) & sDbf
                    lblMessage(idx2).Caption = sMsg
                Case 4
                    ' 4 stations, 3 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle:    " & Space(11) & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
                    sSft = " shift " & Format(iShift, "0")
                    sMsg = sSft & " recipe:     " & sRcp & vbCrLf
                    sMsg = sMsg & "  canister: " & Space(9) & sCan & vbCrLf
                    sMsg = sMsg & sDbf & vbCrLf
                    sMsg = sMsg & sCyc
                    lblMessage(idx2).Caption = sMsg
                Case 5
                    ' 4 stations, 2 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle:    " & Space(11) & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
                    sSft = " shift " & Format(iShift, "0")
                    sMsg = sSft & " recipe: " & vbCrLf
                    gap = 1 + (CharWidth - Len(sRcp)) / 2
                    sMsg = sMsg & Space(gap) & sRcp & vbCrLf
                    sMsg = sMsg & "  canister: " & Space(9) & sCan & vbCrLf
                    sMsg = sMsg & sDbf & vbCrLf
                    sMsg = sMsg & sCyc
                    lblMessage(idx2).Caption = sMsg
                Case 6
                    ' 9 stations, 1 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle:    " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
'                    sSft = " shift " & Format(iShift, "0")
                    sMsg = vbCrLf
                    sMsg = sMsg & "  recipe: " & vbCrLf
                    gap = 1 + (CharWidth - Len(sRcp)) / 2
                    sMsg = sMsg & Space(gap) & sRcp & vbCrLf
                    sMsg = sMsg & "  canister: " & sCan & vbCrLf
                    sMsg = sMsg & sDbf & vbCrLf
'                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sCyc
                    lblMessage(idx2).Caption = sMsg
                Case 9
                    ' 4 stations, 1 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle:    " & Space(11) & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
'                    sSft = " shift " & Format(iShift, "0")
                    sMsg = vbCrLf
                    sMsg = sMsg & "  recipe: " & vbCrLf
                    gap = 1 + (CharWidth - Len(sRcp)) / 2
                    sMsg = sMsg & Space(gap) & sRcp & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & "  canister: " & Space(9) & sCan & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sDbf & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sCyc
                    lblMessage(idx2).Caption = sMsg
                Case 10
                    ' 1 station, 2 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle:    " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
                    sSft = " shift " & Format(iShift, "0")
'                    sMsg = vbCrLf
                    sMsg = sSft & " recipe: " & vbCrLf
                    gap = 1 + (CharWidth - Len(sRcp)) / 2
                    sMsg = sMsg & Space(gap) & sRcp & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & "  canister: " & sCan & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sDbf & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sCyc
                    lblMessage(idx2).Caption = sMsg
                Case 16
                    ' 1 station, 1 shifts
                    sCan = Mid(StationCanister(iStn, iShift).Description, 1, CanLength)
                    sRcp = Mid(StationRecipe(iStn, iShift).Name, 1, RcpLength)
                    sCyc = "  cycle:    " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
'                    sSft = " shift " & Format(iShift, "0")
                    sMsg = vbCrLf
                    sMsg = sMsg & "  recipe: " & vbCrLf
                    sMsg = sMsg & vbCrLf
                    gap = 1 + (CharWidth - Len(sRcp)) / 2
                    sMsg = sMsg & Space(gap) & sRcp & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & "  canister: " & sCan & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sDbf & vbCrLf
                    sMsg = sMsg & vbCrLf
                    sMsg = sMsg & sCyc
                    lblMessage(idx2).Caption = sMsg
            End Select
            
            ' Any Purge Air Supply Problems
            If OOTs(iStn, iShift).AirTempOOT Then PurgeAir_Warning = True
            If OOTs(iStn, iShift).AirMoistOOT Then PurgeAir_Warning = True
            
        Next iShift
    Next iStn

    ' Purge Air Panel
    If PurgeAir_Warning Then
        ' Purge Air is OOT
        PurgeAirMsg_ForeColor = Alarm_ForeColor
        PurgeAirMsg_ToolTip = "Purge Air is OutOfTolerance"
    ElseIf (USINGPASLOCALCONTROL And Com_DIO(icPASisRunningIn).Value) And PAS_INFO(pasTEMPERATURE).timeOut Then
        ' Local PAS Temperature Timeout
        PurgeAirMsg_ForeColor = Alarm_ForeColor
        PurgeAirMsg_ToolTip = "Local PAS Temperature Timeout"
    ElseIf (USINGPASLOCALCONTROL And Com_DIO(icPASisRunningIn).Value) And PAS_INFO(pasMOISTURE).timeOut Then
        ' Local PAS Moisture Timeout
        PurgeAirMsg_ForeColor = Alarm_ForeColor
        PurgeAirMsg_ToolTip = "Local PAS Moisture Timeout"
    ElseIf LogTempRh And Not PAS_INFO(pasTEMPERATURE).Ok Then
        ' Local PAS Temperature Not OK
        PurgeAirMsg_ForeColor = Alarm_ForeColor
        PurgeAirMsg_ToolTip = "Air Temperature is OutOfTolerance"
    ElseIf LogTempRh And Not PAS_INFO(pasMOISTURE).Ok Then
        ' Local PAS Moisture Not OK
        PurgeAirMsg_ForeColor = Alarm_ForeColor
        PurgeAirMsg_ToolTip = "Air Moisture is OutOfTolerance"
    ElseIf (USINGPASLOCALCONTROL And PRG_INFO(1).UsingPrgReqHdw And (Not SysConfig.PosPressPurge)) Then
        If Com_DIO(icPurgeRequestOut).Value Then
            ' Request is True
            If Com_DIO(icPurgeReadyIn).Value Then
                ' Ready is True
                PurgeAirMsg_ForeColor = DK2GREEN
                PurgeAirMsg_ToolTip = "Purge Air Ready"
            Else
                ' Ready is False
                PurgeAirMsg_ForeColor = Warning_ForeColor
                PurgeAirMsg_ToolTip = "Waiting for Purge Air Ready"
            End If
        Else
            ' Request is False
            PurgeAirMsg_ForeColor = TitlesData_Forecolor
            PurgeAirMsg_ToolTip = "Waiting for Purge Air Request"
        End If
    ElseIf ((LocalPagControl.Type = pagClient) And (Not SysConfig.PosPressPurge)) Then
        If (PAG_Request Or MasterPagData.ReqIn) Then
            ' Request is True
            If MasterPagData.RdyOut Then
                ' Ready is True
                PurgeAirMsg_ForeColor = DK2GREEN
                PurgeAirMsg_ToolTip = "Purge Air Ready"
            Else
                ' Ready is False
                PurgeAirMsg_ForeColor = Warning_ForeColor
                PurgeAirMsg_ToolTip = "Waiting for Purge Air Ready"
            End If
        Else
            ' Request is False
            PurgeAirMsg_ForeColor = TitlesData_Forecolor
            PurgeAirMsg_ToolTip = "Waiting for Purge Air Request"
        End If
    Else
        ' Not Using Hardware Request/Ready & No OOT
        PurgeAirMsg_ForeColor = lblCurrDataColor.ForeColor
        PurgeAirMsg_ToolTip = "Purge Air Values"
    End If
    ' temperature units
    If USINGC Then
        sTempUnits = "C   "
    ElseIf USINGF Then
        sTempUnits = "F   "
    End If
    ' moisture units
    If USINGMoist_RH Then
        sMoistUnits = " % rH      "
    ElseIf USINGMoist_Grains Then
        sMoistUnits = " grains/lb      "
    End If
    ' purge air text
    PurgeAirMsg_Text = Format(PATemp, "##0.0") & Chr$(160) & Chr$(176) & sTempUnits & _
        Format(PAMoisture, "##0.0") & sMoistUnits & _
        Format(Now(), "D MMMM YYYY   h:mm:ss")
        
    pnlPurgeAir.ForeColor = PurgeAirMsg_ForeColor
    pnlPurgeAir.ToolTipText = PurgeAirMsg_ToolTip
    pnlPurgeAir.Caption = PurgeAirMsg_Text

    
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

Private Sub cmdStation_Click(Index As Integer)

    DispStn = Index + 1
    If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
        ' Station Detail screen
        DispShift = IIf((Stn_ActiveShift(DispStn) > 0), Stn_ActiveShift(DispStn), 1)
        frmStnDetail.Left = frmMainMenu.Left
        frmStnDetail.Top = frmMainMenu.Top
        frmStnDetail.Show
    Else
        ' LeakTest Station screen
        DispShift = 1
        frmLeakTest.Left = frmMainMenu.Left
        frmLeakTest.Top = frmMainMenu.Top
        frmLeakTest.Show
    End If
    
End Sub

Private Sub lblMessage_Click(Index As Integer)
Dim iStn As Integer
Dim iShift As Integer

    If IsEven(Index) Then
        ' even index = odd-numbered shift (1 or 3)
        If (Index < 18) Then
            iShift = 1
            iStn = 1 + (Index / 2)
        Else
            iShift = 3
            iStn = 1 + ((Index - 18) / 2)
        End If
    ElseIf IsOdd(Index) Then
        ' odd index = even-numbered shift (2 or 4)
        If (Index < 18) Then
            iShift = 2
            iStn = 1 + ((Index - 1) / 2)
        Else
            iShift = 4
            iStn = 1 + ((Index - 19) / 2)
        End If
    Else
        iShift = 1
        iStn = 1
    End If
    If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
        ' Station Detail screen
        DispStn = iStn
        DispShift = iShift
        frmStnDetail.Left = frmMainMenu.Left
        frmStnDetail.Top = frmMainMenu.Top
        frmStnDetail.Show
    Else
        ' LeakTest Station screen
        DispStn = iStn
        DispShift = 1
        frmLeakTest.Left = frmMainMenu.Left
        frmLeakTest.Top = frmMainMenu.Top
        frmLeakTest.Show
    End If
End Sub

Private Sub lblMode_Click(Index As Integer)
Dim iStn, iShift As Integer
Dim count As Long

    If IsEven(Index) Then
        ' even index = odd-numbered shift (1 or 3)
        If (Index < 18) Then
            iShift = 1
            iStn = 1 + (Index / 2)
        Else
            iShift = 3
            iStn = 1 + ((Index - 18) / 2)
        End If
    ElseIf IsOdd(Index) Then
        ' odd index = even-numbered shift (2 or 4)
        If (Index < 18) Then
            iShift = 2
            iStn = 1 + ((Index - 1) / 2)
        Else
            iShift = 4
            iStn = 1 + ((Index - 19) / 2)
        End If
    Else
        iShift = 1
        iStn = 1
    End If
    If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then

        Select Case StationControl(iStn, iShift).Mode
            Case VBLOAD, VBPURGE, VBLEAK
                ' set ReviewData to desired Station
    '            RvwStn = CInt(iStn)
                ' ReviewData screen position
    '            frmReview.Left = frmMainMenu.Left
    '            frmReview.Top = frmMainMenu.Top
    '            frmReview.Show
    '            DoEvents
                ' delay  some
    '                For count = 0 To 500000
    '                    count = count + 1
    '                Next count
                ' begin observing data
    '            frmReview.cmdFirstRecUp.DoClick
    '            DoEvents
                ' delay  some
    '                For count = 0 To 500000
    '                    count = count + 1
    '                Next count
                ' goto latest data
    '            frmReview.cmdFirstRecUp.DoClick
    '            DoEvents
    '            frmReview.cmdFirstRecUp.DoClick
    '            frmDataWatcher.Left = frmMainMenu.Left
    '            frmDataWatcher.Top = frmMainMenu.Top
                frmDataWatcher.Show
            Case Else
                ' Go to Station Detail
                DispStn = iStn
                DispShift = iShift
                ' Station Detail screen
                frmStnDetail.Left = frmMainMenu.Left
                frmStnDetail.Top = frmMainMenu.Top
                frmStnDetail.Show
        End Select
    Else
        ' LeakTest Station screen
        DispStn = iStn
        DispShift = 1
        frmLeakTest.Left = frmMainMenu.Left
        frmLeakTest.Top = frmMainMenu.Top
        frmLeakTest.Show
    End If
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

Private Sub MSComm_OnComm(Index As Integer)
Dim X As Integer
   Select Case MSComm(Index).CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak        ' A Break was received.
        X = X
      Case comEventFrame        ' Framing Error
        X = X
      Case comEventOverrun      ' Data Lost.
        X = X
      Case comEventRxOver       ' Receive buffer overflow.
        X = X
      Case comEventRxParity     ' Parity Error.
        X = X
      Case comEventTxFull       ' Transmit buffer full.
        X = X
      Case comEventDCB          ' Unexpected error retrieving DCB]
        X = X

   ' Events
      Case comEvCD      ' Change in the CD line.
        X = X
      Case comEvCTS     ' Change in the CTS line.
        X = X
      Case comEvDSR     ' Change in the DSR line.
        X = X
      Case comEvRing    ' Change in the Ring Indicator.
        X = X
      Case comEvReceive ' Received RThreshold # of chars.
        If Index = mscommChiller Then ChillerPortRead
      Case comEvSend    ' There are SThreshold number of characters in the transmit buffer.
        If Index = mscommChiller Then LF_Chiller.CmdSentFlag = True
      Case comEvEOF     ' An EOF charater was found in the input stream
        X = X
   End Select
End Sub

Private Sub pbarActual_Click(Index As Integer)
Dim iStn, iShift As Integer
Dim count As Long
    If IsEven(Index) Then
        ' even index = odd-numbered shift (1 or 3)
        If (Index < 18) Then
            iShift = 1
            iStn = 1 + (Index / 2)
        Else
            iShift = 3
            iStn = 1 + ((Index - 18) / 2)
        End If
    ElseIf IsOdd(Index) Then
        ' odd index = even-numbered shift (2 or 4)
        If (Index < 18) Then
            iShift = 2
            iStn = 1 + ((Index - 1) / 2)
        Else
            iShift = 4
            iStn = 1 + ((Index - 19) / 2)
        End If
    Else
        iShift = 1
        iStn = 1
    End If
    If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
        Select Case StationControl(iStn, iShift).Mode
            Case VBLOAD, VBPURGE, VBLEAK
                ' set ReviewData to desired Station
    '            RvwStn = CInt(iStn)
                ' ReviewData screen position
    '            frmReview.Left = frmMainMenu.Left
    '            frmReview.Top = frmMainMenu.Top
    '            frmReview.Show
    '            DoEvents
                ' delay  some
    '                For count = 0 To 500000
    '                    count = count + 1
    '                Next count
                ' begin observing data
    '            frmReview.cmdFirstRecUp.DoClick
    '            DoEvents
                ' delay  some
    '                For count = 0 To 500000
    '                    count = count + 1
    '                Next count
                ' goto latest data
    '            frmReview.cmdFirstRecUp.DoClick
    '            DoEvents
    '            frmReview.cmdFirstRecUp.DoClick
                frmDataWatcher.Left = frmMainMenu.Left
                frmDataWatcher.Top = frmMainMenu.Top
                frmDataWatcher.Show
            Case Else
                ' Go to Station Detail
                DispStn = iStn
                DispShift = iShift
                ' Station Detail screen
                frmStnDetail.Left = frmMainMenu.Left
                frmStnDetail.Top = frmMainMenu.Top
                frmStnDetail.Show
        End Select
    Else
        ' LeakTest Station screen
        DispStn = iStn
        DispShift = 1
        frmLeakTest.Left = frmMainMenu.Left
        frmLeakTest.Top = frmMainMenu.Top
        frmLeakTest.Show
    End If
End Sub

Private Sub pnlMode_Click(Index As Integer)
Dim iStn As Integer
Dim iShift As Integer

    If IsEven(Index) Then
        ' even index = odd-numbered shift (1 or 3)
        If (Index < 18) Then
            iShift = 1
            iStn = 1 + (Index / 2)
        Else
            iShift = 3
            iStn = (1 + ((Index - 18) / 2))
        End If
    ElseIf IsOdd(Index) Then
        ' odd index = even-numbered shift (2 or 4)
        If (Index < 18) Then
            iShift = 2
            iStn = 1 + ((Index - 1) / 2)
        Else
            iShift = 4
            iStn = 1 + ((Index - 19) / 2)
        End If
    Else
        ' default
        iShift = 1
        iStn = 1
    End If
    If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
        ' Station Detail screen
        DispStn = iStn
        DispShift = iShift
        frmStnDetail.Left = frmMainMenu.Left
        frmStnDetail.Top = frmMainMenu.Top
        frmStnDetail.Show
    Else
        ' LeakTest Station screen
        DispStn = iStn
        DispShift = 1
        frmLeakTest.Left = frmMainMenu.Left
        frmLeakTest.Top = frmMainMenu.Top
        frmLeakTest.Show
    End If
End Sub

Private Sub pnlShift_Click(Index As Integer)
Dim iStn As Integer
Dim iShift As Integer

    If IsEven(Index) Then
        ' even index = odd-numbered shift (1 or 3)
        If (Index < 18) Then
            iShift = 1
            iStn = 1 + (Index / 2)
        Else
            iShift = 3
            iStn = (1 + ((Index - 18) / 2))
        End If
    ElseIf IsOdd(Index) Then
        ' odd index = even-numbered shift (2 or 4)
        If (Index < 18) Then
            iShift = 2
            iStn = 1 + ((Index - 1) / 2)
        Else
            iShift = 4
            iStn = 1 + ((Index - 19) / 2)
        End If
    Else
        ' default
        iShift = 1
        iStn = 1
    End If
    If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
        ' Station Detail screen
        DispStn = iStn
        DispShift = iShift
        frmStnDetail.Left = frmMainMenu.Left
        frmStnDetail.Top = frmMainMenu.Top
        frmStnDetail.Show
    Else
        ' LeakTest Station screen
        DispStn = iStn
        DispShift = 1
        frmLeakTest.Left = frmMainMenu.Left
        frmLeakTest.Top = frmMainMenu.Top
        frmLeakTest.Show
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF2 Then frmPassword.Show               'F12 = show passwords
End Sub

Private Sub Form_Load()

SetErrModule 99, 3
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim tmr As Integer
Dim iKeyCount As Integer
Dim BtnWidth As Integer
Dim BtnHeight As Integer
Dim BtnSpace As Integer
Dim NrBtns As Integer

KeyPreview = True

'ActiveTitleBar_Color = lblActiveTitleBarColor.ForeColor

' Position TopLeft of frame
'Top = 0
'Left = 0

' Fudge Overall Height for Menu & Toolbar
'   & allow exactly 9575 in height for the stations
Height = Height - 78 - 120 + 5

' Set Current Time
currDTS = Now()
CurrTimer = Timer()
' Set Logic Timer Intervals
For tmr = 1 To 9
    tmrTimer(tmr).Interval = SystemTimers(tmr).Interval
    SystemTimers(tmr).LastTimer = Timer()
Next tmr

' Program Error Msg Bypass Setup
ShortTermErrorMax = 32
ShortTermErrorCounter = 0

' Set Screen Update Timer Interval
tmrScreen.Interval = 1000

IOForceActive = False
ScreenIsBuilt = False

' Build Toolbars
BuildToolbars
' Update Toolbar Buttons
UpdateNavigateBtns
' Build Status Bars
BuildStatusBars

' Update Screen
UpdateMain

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not CheckPass("G", True) Then Cancel = 1
End Sub

Private Sub pnlStn_Click(Index As Integer)
    DispStn = Index + 1
    DispShift = IIf((Stn_ActiveShift(DispStn) > 0), Stn_ActiveShift(DispStn), 1)

    ' Station Detail screen
    frmStnDetail.Left = frmMainMenu.Left
    frmStnDetail.Top = frmMainMenu.Top
    frmStnDetail.Show
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

Private Sub tmrScreen_Timer()
    ' Update the Navigate Toolbar buttons
    UpdateNavigateBtns
    ' Update Main Screen
    UpdateMain
    ' ask for a username & password, if not already logged in
    If CurrentUser.USER = "DEFAULT" Then
        frmPassword.Show
    End If
End Sub

Private Sub tmrTimer_Timer(Index As Integer)

' cooperative multi-tasking executive timer
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 99, 110
Dim address, prg, shft, stn, prt, scl As Integer
Dim sSclMax As Single
Dim sSclMin As Single
Dim sSclSpan As Single
Dim calweight As Single
Dim calscaleread As String

    TimerData CInt(Index)
    
    ChgErrModule 99, CInt(110 + Index)
    
    Select Case Index
        Case tmrScanIO
        
            ' request to turn opto22 communications on ??
            If OptoCommOn_Request Then frmMainForm.OptoCommOn
            
            If IoComOn Then
                ' IO SCAN
                ChgErrModule 99, CInt(1100 + SystemTimers(tmrScanIO).Phase)
                Select Case SystemTimers(tmrScanIO).Phase
                    Case 0
                        ' Read Common Board
                        If Node_Info(0) > 0 Then
                            OPTO_ReadDigital 0
                            OPTO_ReadDigital 1
                            OPTO_ReadAnalog 2
                            If Node_Info(0) > 8 Then OPTO_ReadAnalog 3
                        End If
                        
                        ' Map Common
                        Map_ComDigitals
                        Map_ComAnalogs
            
                    
                    Case 1 To 9
                        ' Read Station Board
                        stn = SystemTimers(tmrScanIO).Phase
                        If Node_Info(stn) > 0 Then
                            address = stn * 4                   ' station base + 0
                            OPTO_ReadDigital CInt(address)
                            address = address + 1               ' station base + 1
                            OPTO_ReadDigital CInt(address)
                            address = address + 1               ' station base + 2
                            OPTO_ReadAnalog CInt(address)
                            address = address + 1               ' station base + 3
                            If Node_Info(stn) > 8 Then OPTO_ReadAnalog CInt(address)
                        End If
                    
                        ' Map Station
                        Map_StnDigitals CInt(stn)
                        Map_StnAnalogs CInt(stn)
                    
                    Case Else
                        ' Map Purge IO
                        For prg = 1 To NR_PRGAIR
                            Map_PrgDigitals CInt(prg)
                            Map_PrgAnalogs CInt(prg)
                        Next prg
                        ' Comm Flags
                        OptoReadAllOnce = True
                        IoComm_Flag = True
            
                End Select
            
                ' request to turn opto22 communications off ??
                If (OptoCommOff_Request And (Not OptoCommOn_Request)) Then frmMainForm.OptoCommOff
            
            Else
                
                    ' Simulation of I/O
                    ChgErrModule 99, 1111
                    If USINGSIMULATION Then SimulateIO
        
            End If
            SystemTimers(tmrScanIO).Phase = SystemTimers(tmrScanIO).Phase + 1
            If (SystemTimers(tmrScanIO).Phase < 10 And SystemTimers(tmrScanIO).Phase > NR_STN) Then SystemTimers(tmrScanIO).Phase = 10
            If SystemTimers(tmrScanIO).Phase > 10 Then SystemTimers(tmrScanIO).Phase = 0
            
            
        Case tmrScales
            ' READ SCALES
            If NR_SCALES > 0 Then
                
                ' request to turn scale communications on ??
                If ScaleCommOn_Request Then frmComm8Card.ScaleCommOn
            
                If SclComOn Then
                    ChgErrModule 99, 1111
                    ' cycle thru ports with scales
                    For prt = 1 To MAX_COMM
                        ' Read the Port
                        If Port_In_Use(prt) Then frmComm8Card.Read_Comm_Port CInt(prt)
                    Next prt
                 
                    ChgErrModule 99, 1112
                     ' cycle thru scales
                    For scl = 1 To NR_SCALES
                    ChgErrModule 99, 1700 + CInt(scl)
                        ' Map the Port Values to the Scale
                        prt = Scale_Port(scl)
                        sSclMax = Scale_Cal(scl).CalRangeMax
                        sSclMin = Scale_Cal(scl).CalRangeMin
                        sSclSpan = sSclMax - sSclMin
                        If (sSclSpan > 0) Then
                            calweight = Cal_Scale(((Port_Weight(prt) - sSclMin) / sSclSpan), scl, Scale_Cal(scl))
                        Else
                            calweight = Port_Weight(prt)
                        End If
                        calscaleread = Format(calweight, "###,##0.00#")
                        Scale_OK(scl) = Port_OK(prt)
                        Scale_Weight(scl) = calweight
                        Scale_Value(scl) = calscaleread
                    Next scl
                
                    ' request to turn scale communications off ??
                    If (ScaleCommOff_Request And (Not ScaleCommOn_Request)) Then frmComm8Card.ScaleCommOff
            
                Else
                
                    ' Simulation of Scales
                    ChgErrModule 99, 1112
                    If USINGSIMULATION Then SimulateScales
        
                End If
               
                 ' Map Scale Values to StationScales
                ChgErrModule 99, 1122
                For stn = 1 To NR_STN
                    For shft = 1 To NR_SHIFT
                        If Not StationRecipe(stn, shft).UseAuxScale And Not StationRecipe(stn, shft).UsePriScale Then
                            StationControl(stn, shft).Scale_OK = False
                        Else
                            If StationRecipe(stn, shft).UseAuxScale Then
                                ' using aux scale
                                StationControl(stn, shft).Scale_OK = Scale_OK(StationRecipe(stn, shft).AuxScaleNo)
                                StationControl(stn, shft).AuxScaleWt = Scale_Weight(StationRecipe(stn, shft).AuxScaleNo)
                            Else
                                StationControl(stn, shft).AuxScaleWt = CSng(0)
                            End If
                            If StationRecipe(stn, shft).UsePriScale Then
                                If StationRecipe(stn, shft).UseAuxScale Then
                                    ' using both scales
                                    StationControl(stn, shft).Scale_OK = IIf(Scale_OK(StationRecipe(stn, shft).PriScaleNo), StationControl(stn, shft).Scale_OK, False)
                                Else
                                    ' using only primary scale
                                    StationControl(stn, shft).Scale_OK = Scale_OK(StationRecipe(stn, shft).PriScaleNo)
                                End If
                                StationControl(stn, shft).PriScaleWt = Scale_Weight(StationRecipe(stn, shft).PriScaleNo)
                            Else
                                StationControl(stn, shft).PriScaleWt = CSng(0)
                            End If
                        End If
                    Next shft
                Next stn
                
            End If
                   
        Case tmrAlmOOT
            ' ALARM/OOT
            If OptoReadAllOnce Then             ' On Startup, wait for initial IoReads before checking alarms and OOT's
                Select Case SystemTimers(tmrAlmOOT).Phase
                    Case 0
                        Alarm_Check             ' Check Alarm conditions
                    Case 1
                        OOT_Check               ' Check Out-Of-Tolerance conditions
                End Select
            End If
            SystemTimers(tmrAlmOOT).Phase = SystemTimers(tmrAlmOOT).Phase + 1
            If SystemTimers(tmrAlmOOT).Phase > 1 Then SystemTimers(tmrAlmOOT).Phase = 0
        
        Case tmrDataLog
            ' DATA LOGGER
            Data_Writer
        
        Case tmrControl
            ' CONTROLLERS
            Select Case SystemTimers(tmrControl).Phase
                Case 0
                    Common_Valves                                   ' Control Common Valves to match Station needs
                Case 1
                    LiveFuel_Controller                             ' Control LiveFuel ADF Sequence & Heater(s)
                Case 2
                    Select Case LocalPagControl.Type
                        Case pagNone
                            ' no pag control
                            If LogTempRh Then
                                '   Check AIR Temperature
                                AIR_Check pasTEMPERATURE                    ' Tolerance check for PAS Temp for AirLog
                                '   Check AIR Moisture
                                AIR_Check pasMOISTURE                       ' Tolerance check for PAS Moisture for AirLog
                            End If
                        Case pagAlone
                            ' Local Control of PAS Temp & Moisture
                            PAS_LocalControl
                        Case pagMaster
                            ' Local Control of PAS Temp & Moisture
                            PAS_LocalControl
                        Case pagClient
                            ' PAG Master Error Status
                            UpdateErrorStatus
                            If LogTempRh Then
                                '   Check AIR Temperature
                                AIR_Check pasTEMPERATURE                    ' Tolerance check for PAS Temp for AirLog
                                '   Check AIR Moisture
                                AIR_Check pasMOISTURE                       ' Tolerance check for PAS Moisture for AirLog
                            End If
                    End Select
                    PurgeAir_Controller                             ' Control PurgeAir "Source(s)" to match Station needs
                Case 3
                    Seq_Controller                                  ' Control Station Sequence(s)
                Case 4
                    If ChillComOn Then ChillerCommander             ' Chiller Control
                    PurgeOven_Controller                            ' Purge Oven Controller
            End Select
            SystemTimers(tmrControl).Phase = SystemTimers(tmrControl).Phase + 1
            If SystemTimers(tmrControl).Phase > 4 Then SystemTimers(tmrControl).Phase = 0
        
        Case tmrStnLogic
            ' STATIONS
            Check_Stations
        
        Case tmrSysTmr
            ' SYSTEM TIMER UPDATE
            ' Determine Elapsed Time (in seconds) since previous update
            prevDts = currDTS
            PrevTimer = CurrTimer
            currDTS = Now()
            CurrTimer = Timer()
            If CurrTimer < PrevTimer Then
                ' Time Interval crosses Midnight OR a Glitch ?
                Write_ELog "CurrTimer (" & Format(CurrTimer, "#####0.000") & ") < PrevTimer (" & Format(PrevTimer, "#####0.000") & ")"
                Select Case CurrTimer
                    Case Is < 1#
                        ' Current Time is shortly after Midnight
                        Select Case PrevTimer
                            Case Is > 86399#
                                ' Previous Time is just before Midnight
                                DeltTimer = (CurrTimer + 86400#) - PrevTimer
                                Write_ELog "Normal Rollover at Midnight (PrevDts - " & Format(prevDts, "YYYY-MM-DD hh:mm:ss") & ")"
                            Case Is > 86396#
                                ' Previous Time is shortly before Midnight
                                DeltTimer = (CurrTimer + 86400#) - PrevTimer
                                Write_ELog "Long-Interval Rollover at Midnight (PrevDts - " & Format(prevDts, "YYYY-MM-DD hh:mm:ss") & ")"
                            Case Else
                                ' Previous Time is not near Midnight
                                DeltTimer = CDbl(DateDiff("s", prevDts, currDTS))
                                Write_ELog "DateDiff is " & Format(DeltTimer, "#####0.000") & " (PrevDts - " & Format(prevDts, "YYYY-MM-DD hh:mm:ss") & ")"
                        End Select
                    Case Else
                        DeltTimer = CDbl(DateDiff("s", prevDts, currDTS))
                        Write_ELog "DateDiff is " & Format(DeltTimer, "#####0.000") & " (PrevDts - " & Format(prevDts, "YYYY-MM-DD hh:mm:ss") & ")"
                End Select
            Else
                ' Normal Time Interval (i.e. doesn't cross Midnight)
                DeltTimer = CurrTimer - PrevTimer
            End If
            ' log abnormal cases; fix DeltTimer
            Select Case DeltTimer
                Case Is >= 86400#
                    ' delta time too big by over a day; fix it
                    Write_ELog "DeltTimer (" & Format(DeltTimer, "#####0.000") & ") > 86400; set to " & Format(DeltTimer - 86400#, "#####0.000")
                    DeltTimer = DeltTimer - 86400#
                Case Is > 1000#
                    ' delta time is suspiciously big; fix it
                    Write_ELog "DeltTimer (" & Format(DeltTimer, "#####0.000") & ") > 1000; set to 0"
                    DeltTimer = 0#
                Case Is > 1#
                    ' delta time is unusually big
                    Write_ELog "DeltTimer (" & Format(DeltTimer, "#####0.000") & ") > 1"
                Case Is < 0#
                    ' delta time is negative; fix it
                    Write_ELog "DeltTimer (" & Format(DeltTimer, "#####0.000") & ") < 0; set to 0"
                    DeltTimer = 0#
                Case Else
                    ' normal; do nothing
            End Select
            
            ' Update running TestTimers
            For tmrStation = 1 To LAST_STN
                For tmrShift = 1 To NR_SHIFT
                    If StationControl(tmrStation, tmrShift).TestTimerIsRunning Then
                        StationControl(tmrStation, tmrShift).TestTimer = StationControl(tmrStation, tmrShift).TestTimer + DeltTimer
                    End If
                Next tmrShift
            Next tmrStation
        
            ' Error Message Bypass
            If USINGERRORMSGBYPASS Then
                ' Update the "Bypass Error Messages Active" flag
                '   if too many errors turn off errormsgbypass
                If ShortTermErrorCounter < ShortTermErrorMax Then
                    ErrorMsgBypassActive = True
                Else
                    ErrorMsgBypassActive = False
                End If
                ' Decrement the "Recent Error Message(s)" Counter
                If (SystemTimers(tmrSysTmr).Phase = 1) Then
                    If (ShortTermErrorCounter > 0) Then
                        ShortTermErrorCounter = ShortTermErrorCounter - 1
                    End If
                End If
            Else
                ShortTermErrorCounter = 0
                ErrorMsgBypassActive = False
            End If
                ' Update the "Bypass Error Messages Active" flag
            If USINGERRORMSGBYPASS Then
                If ShortTermErrorCounter < ShortTermErrorMax Then
                    ErrorMsgBypassActive = True
                Else
                    ErrorMsgBypassActive = False
                End If
            Else
                ErrorMsgBypassActive = False
            End If
            
            ' AK Client screen support
            If (LocalPagControl.Type = pagClient) Then
                If (frmAkClient.AkClientResetReq) Then
                    Select Case AkClientResetStepNum
                        Case 0
                            ' turn off
                            AkClientResetStepNumNext = 1
                        Case 1
                            ' idle; nothing to do
                            AkClientResetStepStartSecs = 0
                            AkClientResetStepDoneSecs = 0
                            If (frmAkClient.AkClientResetReq) Then AkClientResetStepNumNext = 2
                        Case 2
                            ' start; unload the form
                            If (Timer < 86390) Then
                                Unload frmAkClient
'                                frmAkClient = Nothing
                                AkClientResetStepStartSecs = Timer
                                AkClientResetStepDoneSecs = AkClientResetStepStartSecs + 2
                                AkClientResetStepNumNext = 3
                            End If
                        Case 3
                            ' short pause
                            If (Timer > AkClientResetStepDoneSecs) Then
                                AkClientResetStepNumNext = 4
                            End If
                        Case 4
                            ' re-load the form
                            frmAkClient.Show
                            AkClientResetStepNumNext = 5
                        Case 5
                            frmAkClient.AkClientResetReq = False
                            AkClientResetStepNumNext = 1
                    End Select
                    AkClientResetStepNum = AkClientResetStepNumNext
                End If
            End If
            
            ' update the phase
            SystemTimers(tmrSysTmr).Phase = SystemTimers(tmrSysTmr).Phase + 1
            If SystemTimers(tmrSysTmr).Phase > 99 Then SystemTimers(tmrSysTmr).Phase = 0
        
        Case tmrUnused8
            ' unused
            Delay_Box "Unused System Timer #" & Format(Index, "#0"), MSGDELAY, msgSHOW
            Write_ELog "Unused System Timer #" & Format(Index, "#0")
        
        Case tmrUnused9
            ' unused
            Delay_Box "Unused System Timer #" & Format(Index, "#0"), MSGDELAY, msgSHOW
            Write_ELog "Unused System Timer #" & Format(Index, "#0")
    
    End Select
    

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

Sub TimerData(tmr As Integer)
    SystemTimers(tmr).Actual = (Timer - SystemTimers(tmr).LastTimer) * 1000
    SystemTimers(tmr).delta = CDbl(SystemTimers(tmr).Interval) - SystemTimers(tmr).Actual
    If SystemTimers(tmr).delta > SystemTimers(tmr).max Then SystemTimers(tmr).max = SystemTimers(tmr).delta
    If SystemTimers(tmr).delta < SystemTimers(tmr).Min Then SystemTimers(tmr).Min = SystemTimers(tmr).delta
    SystemTimers(tmr).LastTimer = Timer
End Sub

Private Sub BuildToolbars()

    ' Load the ImageLists
    tbrNavigate.ImageList = imgNavigateNormal
    tbrNavigate.DisabledImageList = imgNavigateDisabled
    tbrNavigate.HotImageList = imgNavigateHot
    
    ' Create object variable for the Toolbar.
    Dim btnX As MSComctlLib.Button
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
    
    'TOM Can Load Tasks Screen
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
    
    'Exit Program
'    Set btnX = tbrNavigate.Buttons.Add(, "exit", , tbrDefault, "exit")
'    btnX.ToolTipText = "Exit Program"
'    btnX.Description = btnX.ToolTipText
    
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
'    Set btnX = tbrNavigate.Buttons.Add(, , , tbrSeparator)
    
    
    
'    Show
    
End Sub

Public Sub ChillerCommInit()
Dim errors As Integer
    MSComm(mscommChiller).CommPort = Chiller_PORT
    MSComm(mscommChiller).PortOpen = True
    MSComm(mscommChiller).Settings = "9600,N,8,1"
    MSComm(mscommChiller).InputLen = 0
    MSComm(mscommChiller).InputMode = comInputModeText
    MSComm(mscommChiller).RThreshold = 1
    MSComm(mscommChiller).SThreshold = 1
'    MSComm(mscommChiller).Handshaking = comNone
'    MSComm(mscommChiller).Handshaking = comXOnXOff
    MSComm(mscommChiller).Handshaking = comRTS
'    MSComm(mscommChiller).Handshaking = comRTSXOnXOff
    errors = Not MSComm(mscommChiller).PortOpen
    LF_Chiller.InitComplete = IIf(errors = 0, True, False)
End Sub

Private Sub ChillerPortRead()
    LF_Chiller.BufferIn = LF_Chiller.BufferIn + MSComm(mscommChiller).Input
    '   EXIT (if new char is not CrLf)
    If InStr(LF_Chiller.BufferIn, vbCrLf) = 0 Then Exit Sub         ' no CrLf
     
    ' Read Complete (includes CrLf)
    LF_Chiller.ErrorCount = 0
    LF_Chiller.CmdRecChars = Mid(LF_Chiller.BufferIn, 1, InStr(LF_Chiller.BufferIn, vbCrLf) - 1)
    LF_Chiller.BufferIn = LF_Chiller.BufferIn & MSComm(mscommChiller).Input
    LF_Chiller.BufferIn = ""
    
    ' What have we read ?
    If InStr(LF_Chiller.CmdRecChars, "OK") > 0 Then
        ' Read is a cmd ack
        LF_Chiller.CmdRecAckFlag = True
        LF_Chiller.CurCmdComplete = True
    
    ElseIf InStr(LF_Chiller.CmdRecChars, "ERR_") > 0 Then
        ' Read is an error response
        LF_Chiller.CmdRecErrorFlag = True
        LF_Chiller.CmdRecErrorNumber = CInt(Mid(LF_Chiller.CmdRecChars, 5, Len(LF_Chiller.CmdRecChars) - 4))
    
    ElseIf IsNumeric(LF_Chiller.CmdRecChars) Then
        ' Read is a numeric value
        LF_Chiller.CurCmdComplete = True
        LF_Chiller.CmdRecValueChars = LF_Chiller.CmdRecChars
        
        Select Case LF_Chiller.CurCmdIdx
            Case chillerIn_PV
                LF_Chiller.PvIn = CSng(LF_Chiller.CmdRecValueChars)
            Case chillerIn_SP
                LF_Chiller.SpIn = CSng(LF_Chiller.CmdRecValueChars)
            Case chillerIn_Out
                LF_Chiller.OutIn = CInt(LF_Chiller.CmdRecValueChars)
            Case chillerIn_OperMode
                LF_Chiller.OperModeIn = CInt(LF_Chiller.CmdRecValueChars)
            Case chillerIn_OvrTmpSp
                LF_Chiller.OvrTmpSpIn = CSng(LF_Chiller.CmdRecValueChars)
            Case chillerIn_P
                LF_Chiller.P_In = CSng(LF_Chiller.CmdRecValueChars)
            Case chillerIn_I
                LF_Chiller.I_In = CSng(LF_Chiller.CmdRecValueChars)
            Case chillerIn_Mode
                LF_Chiller.ModeIn = CInt(LF_Chiller.CmdRecValueChars)
            Case chillerIn_Type
                LF_Chiller.Type = LF_Chiller.CmdRecChars
            Case chillerIn_Version
                LF_Chiller.Version = LF_Chiller.CmdRecChars
            Case chillerIn_Status
                LF_Chiller.StatusIn = CInt(LF_Chiller.CmdRecValueChars)
            Case chillerIn_Stat
                LF_Chiller.StatIn = LF_Chiller.CmdRecValueChars
                LF_Chiller.Overtemp = IIf(Mid(LF_Chiller.StatIn, 1, 1) = "1", True, False)
                LF_Chiller.LowLevel = IIf(Mid(LF_Chiller.StatIn, 2, 1) = "1", True, False)
                LF_Chiller.PumpBlocked = IIf(Mid(LF_Chiller.StatIn, 3, 1) = "1", True, False)
                LF_Chiller.IntFaultMc1 = IIf(Mid(LF_Chiller.StatIn, 4, 1) = "1", True, False)
                LF_Chiller.IntFaultMc2 = IIf(Mid(LF_Chiller.StatIn, 5, 1) = "1", True, False)
        End Select
                
    ElseIf LF_Chiller.CurCmdIdx = chillerIn_Type Then
        LF_Chiller.CurCmdComplete = True
        LF_Chiller.Type = LF_Chiller.CmdRecChars
    ElseIf LF_Chiller.CurCmdIdx = chillerIn_Version Then
        LF_Chiller.CurCmdComplete = True
        LF_Chiller.Version = LF_Chiller.CmdRecChars
        
    Else
        ' Read is unknown type
    End If
End Sub


