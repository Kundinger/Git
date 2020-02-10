VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form frmViewThermo 
   Caption         =   "View Thermocouples"
   ClientHeight    =   15240
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   21360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   21360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEngineer 
      Height          =   285
      Left            =   2850
      MaxLength       =   25
      TabIndex        =   36
      ToolTipText     =   "Alphanumeric Name"
      Top             =   8085
      Width           =   3240
   End
   Begin VB.TextBox txtVehicle 
      Height          =   285
      Left            =   2850
      MaxLength       =   25
      TabIndex        =   35
      ToolTipText     =   "Alphanumeric Vehicle Identification Number"
      Top             =   8370
      Width           =   3240
   End
   Begin VB.TextBox txtEndOp 
      Height          =   285
      Left            =   2850
      MaxLength       =   25
      TabIndex        =   34
      ToolTipText     =   "Alphanumeric Name"
      Top             =   8940
      Width           =   3240
   End
   Begin VB.TextBox txtStartOp 
      Height          =   285
      Left            =   2850
      MaxLength       =   25
      TabIndex        =   33
      ToolTipText     =   "Alphanumeric Name"
      Top             =   8655
      Width           =   3240
   End
   Begin VB.Frame frmComment 
      Caption         =   "Comment"
      ForeColor       =   &H00004080&
      Height          =   975
      Left            =   1320
      TabIndex        =   31
      Top             =   9360
      Width           =   4890
      Begin VB.TextBox txtComment 
         Height          =   675
         Left            =   90
         MaxLength       =   360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   32
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   180
         Width           =   4665
      End
   End
   Begin Threed.SSPanel pnlXYGraph 
      Height          =   5700
      Left            =   3600
      TabIndex        =   19
      Top             =   1680
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   10054
      _StockProps     =   15
      ForeColor       =   65280
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
      Begin MSChart20Lib.MSChart chtStnChart 
         Height          =   5475
         Left            =   120
         OleObjectBlob   =   "frmViewThermo.frx":0000
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   5625
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit Temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2250
      TabIndex        =   0
      ToolTipText     =   "Close Thermocouple display"
      Top             =   960
      Width           =   4335
   End
   Begin Threed.SSPanel pnlTC2 
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Thermocouple Two Value"
      Top             =   600
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "199.9"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC1 
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Thermocouple One Value"
      Top             =   135
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "97.4"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC4 
      Height          =   345
      Left            =   4200
      TabIndex        =   3
      ToolTipText     =   "Thermocouple Four Value"
      Top             =   600
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "98.6"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC3 
      Height          =   345
      Left            =   4200
      TabIndex        =   4
      ToolTipText     =   "Thermocouple Three Value"
      Top             =   135
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "97.4"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC6 
      Height          =   345
      Left            =   7290
      TabIndex        =   13
      ToolTipText     =   "Thermocouple Six Value"
      Top             =   630
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "98.6"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC5 
      Height          =   345
      Left            =   7305
      TabIndex        =   14
      ToolTipText     =   "Thermocouple Five Value"
      Top             =   165
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "97.4"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlScale 
      Height          =   1500
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Current Scale Reading"
      Top             =   5280
      Width           =   2925
      _Version        =   65536
      _ExtentX        =   5159
      _ExtentY        =   2646
      _StockProps     =   15
      Caption         =   "Scale Weight in Grams"
      ForeColor       =   4210816
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   8
      Begin Threed.SSPanel pnlScaleWt 
         Height          =   375
         Left            =   90
         TabIndex        =   22
         ToolTipText     =   "Scale Display"
         Top             =   675
         Width           =   1495
         _Version        =   65536
         _ExtentX        =   2637
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "00000000.00"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlscalePRI 
         Height          =   375
         Left            =   90
         TabIndex        =   23
         ToolTipText     =   "Scale Display"
         Top             =   240
         Width           =   1495
         _Version        =   65536
         _ExtentX        =   2637
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "00000000.00"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aux. #"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1660
         TabIndex        =   25
         Top             =   742
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Primary #"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1660
         TabIndex        =   24
         Top             =   307
         Width           =   1200
      End
   End
   Begin Threed.SSPanel pnlDelay 
      Height          =   1215
      Left            =   360
      TabIndex        =   26
      ToolTipText     =   "Current Time Remaining"
      Top             =   3120
      Width           =   2925
      _Version        =   65536
      _ExtentX        =   5159
      _ExtentY        =   2143
      _StockProps     =   15
      Caption         =   "Delay"
      ForeColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   8
      Begin Threed.SSPanel pnlToGo 
         Height          =   375
         Left            =   1660
         TabIndex        =   27
         ToolTipText     =   "Delay Remaining in Seconds"
         Top             =   120
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BackColor       =   -2147483626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlTotal 
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         ToolTipText     =   "Total Delay in Seconds"
         Top             =   480
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1764
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BackColor       =   -2147483626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   90
         TabIndex        =   29
         Top             =   545
         Width           =   1560
      End
   End
   Begin Threed.SSPanel pnlStnDtlMsg 
      Height          =   885
      Left            =   1320
      TabIndex        =   41
      Top             =   10560
      Width           =   8520
      _Version        =   65536
      _ExtentX        =   15028
      _ExtentY        =   1561
      _StockProps     =   15
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.TextBox txtStnDtlMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   705
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Text            =   "frmViewThermo.frx":2AD5
         Top             =   90
         Width           =   8325
      End
   End
   Begin VB.Label lblVehicle 
      Caption         =   "Vehicle No.:"
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
      Left            =   1440
      TabIndex        =   40
      Top             =   8385
      Width           =   1335
   End
   Begin VB.Label lblEngineer 
      Caption         =   "Engineer:"
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
      Left            =   1440
      TabIndex        =   39
      Top             =   8100
      Width           =   1335
   End
   Begin VB.Label lblEndOp 
      Caption         =   "End Operator:"
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
      Left            =   1440
      TabIndex        =   38
      Top             =   8955
      Width           =   1335
   End
   Begin VB.Label lblStartOp 
      Caption         =   "Start Operator:"
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
      Left            =   1440
      TabIndex        =   37
      Top             =   8670
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6225
      TabIndex        =   18
      Top             =   165
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6225
      TabIndex        =   17
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   8295
      TabIndex        =   16
      Top             =   165
      Width           =   855
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   8295
      TabIndex        =   15
      Top             =   630
      Width           =   855
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   5205
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   5205
      TabIndex        =   11
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   10
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3135
      TabIndex        =   9
      Top             =   135
      Width           =   1005
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   2200
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   2200
      TabIndex        =   7
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC  #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   135
      Width           =   1000
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   5
      Top             =   600
      Width           =   1005
   End
End
Attribute VB_Name = "frmViewThermo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'no error mod ''''' Form Expand.FRM ''' filler form '''''''''''''''''''''''''
' 6/1/2000
Option Explicit

Const NumPoints = 600
' Note: the number of points displayed on the graph is the number of elements
' allocated in the first dimension of the Graph array
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Update()
    pnlTC1 = Format(Com_AIO(acCommonTC1).EUValue, "#00.0#")
    pnlTC2 = Format(Com_AIO(acCommonTC2).EUValue, "#00.0#")
    pnlTC3 = Format(Com_AIO(acCommonTC3).EUValue, "#00.0#")
    pnlTC4 = Format(Com_AIO(acCommonTC4).EUValue, "#00.0#")
    pnlTC5 = Format(Com_AIO(acCommonTC5).EUValue, "#00.0#")
    pnlTC6 = Format(Com_AIO(acCommonTC6).EUValue, "#00.0#")
End Sub

