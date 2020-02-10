VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmSysDefStn 
   Caption         =   "Station Definition"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   Icon            =   "frmSysDefStn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmComdef 
      Caption         =   "Common Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1095
      Left            =   120
      TabIndex        =   45
      Top             =   6605
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton cmdScaleCfg 
         Caption         =   "Configure Scales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Configure Scales"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   3360
      End
      Begin VB.CommandButton cmdPurgeCfg 
         Caption         =   "ConfigurePurgeAir Sources"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":5EE4
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Configure Purge Air Sources"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   3360
      End
   End
   Begin VB.ComboBox BoardType 
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
      Height          =   315
      ItemData        =   "frmSysDefStn.frx":65E6
      Left            =   9000
      List            =   "frmSysDefStn.frx":65F8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Number of Slots on Opto Board"
      Top             =   1280
      Width           =   2565
   End
   Begin VB.Frame frmStnDef 
      Caption         =   "Station Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   5175
      Left            =   120
      TabIndex        =   48
      Top             =   1440
      Width           =   8535
      Begin VB.TextBox txtAbreviation 
         Alignment       =   2  'Center
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
         Left            =   7920
         MaxLength       =   1
         TabIndex        =   89
         Text            =   "W"
         ToolTipText     =   "Station (one letter) Abreviation"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSysID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         MaxLength       =   60
         TabIndex        =   87
         Text            =   "WXWXWXWX"
         ToolTipText     =   "Station Unique System ID (max 8 alphnumeric characters)"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame frmPurge 
         Caption         =   "Purge Options"
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
         Height          =   855
         Left            =   4920
         TabIndex        =   84
         Top             =   2760
         Width           =   3375
         Begin VB.CheckBox chkPurgeOven 
            Alignment       =   1  'Right Justify
            Caption         =   "Purge Oven"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "This station has a Purge Oven"
            Top             =   480
            Width           =   3135
         End
      End
      Begin VB.Frame frmStnType 
         Caption         =   "Station Type"
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
         Height          =   3020
         Left            =   240
         TabIndex        =   70
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton optLeakTest 
            Caption         =   "LeakTest"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   86
            ToolTipText     =   "Dummy Station; IO Only"
            Top             =   2300
            Width           =   1500
         End
         Begin VB.OptionButton optCombo3 
            Caption         =   "Combo3"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   80
            ToolTipText     =   "future"
            Top             =   2040
            Width           =   1500
         End
         Begin VB.OptionButton optOrvr2Live 
            Caption         =   "ORVR2/LiveFuel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   79
            ToolTipText     =   "future"
            Top             =   1780
            Width           =   1740
         End
         Begin VB.OptionButton optRegLive 
            Caption         =   "Reg/LiveFuel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   78
            ToolTipText     =   "Butane and Live Fuel Station"
            Top             =   1520
            Width           =   1500
         End
         Begin VB.OptionButton optRegular 
            Caption         =   "Regular"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   75
            ToolTipText     =   "Regular Station"
            Top             =   480
            Width           =   1500
         End
         Begin VB.OptionButton optORVR 
            Caption         =   "ORVR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   74
            ToolTipText     =   "ORVR Station"
            Top             =   740
            Width           =   1500
         End
         Begin VB.OptionButton optLiveFuel 
            Caption         =   "Live Fuel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   73
            ToolTipText     =   "Live Fuel Station"
            Top             =   1260
            Width           =   1500
         End
         Begin VB.OptionButton optDummy 
            Caption         =   "Dummy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   72
            ToolTipText     =   "Dummy Station; IO Only"
            Top             =   2560
            Width           =   1500
         End
         Begin VB.OptionButton optORVR2 
            Caption         =   "ORVR2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   71
            ToolTipText     =   "ORVR2 Station (has 2 Nitrogen & 2 Butane MFC's"
            Top             =   1000
            Width           =   1500
         End
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   4560
         MaxLength       =   60
         TabIndex        =   69
         ToolTipText     =   "Station Description (alphnumeric)"
         Top             =   240
         Width           =   3735
      End
      Begin VB.Frame frmLiveFuel 
         Caption         =   "Live Fuel  Options"
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
         Height          =   2420
         Left            =   2520
         TabIndex        =   63
         Top             =   1200
         Width           =   2175
         Begin VB.CheckBox chkWaterBath 
            Alignment       =   1  'Right Justify
            Caption         =   "WaterBath"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "This LiveFuel Tank has a WaterBath Heater/Chiller"
            Top             =   1020
            Width           =   1695
         End
         Begin VB.CheckBox chkLevelXmtr 
            Alignment       =   1  'Right Justify
            Caption         =   "Level Xmtr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "This LiveFuel Tank has a Heater"
            Top             =   1240
            Width           =   1695
         End
         Begin VB.CheckBox chkFuelStorage 
            Alignment       =   1  'Right Justify
            Caption         =   "Storage Tank"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   81
            ToolTipText     =   "This LiveFuel system has a Fuel Storage Tank"
            Top             =   1460
            Width           =   1695
         End
         Begin VB.CheckBox chkHeater 
            Alignment       =   1  'Right Justify
            Caption         =   "InTank Heater"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "This LiveFuel Tank has a Heater"
            Top             =   800
            Width           =   1695
         End
         Begin VB.CheckBox chkADF 
            Alignment       =   1  'Right Justify
            Caption         =   "AutoDrainFill"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "This LiveFuel Tank has Auto Drain & Fill"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkVapor 
            Alignment       =   1  'Right Justify
            Caption         =   "Vapor Valve"
            BeginProperty Font 
               Name            =   "Arial"
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
            ToolTipText     =   "This LiveFuel Tank has a Vapor Valve"
            Top             =   580
            Width           =   1695
         End
         Begin VB.TextBox txtTank 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   64
            ToolTipText     =   "Live Fuel Tank Number (1-9)"
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblTank 
            BackStyle       =   0  'Transparent
            Caption         =   "Fuel Tank #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   2085
            Width           =   1155
         End
      End
      Begin VB.Frame frmAspirator 
         Caption         =   "Purge Aspirator Selection"
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
         Height          =   1335
         Left            =   240
         TabIndex        =   59
         Top             =   3720
         Width           =   4455
         Begin VB.CommandButton cmdPrgDn 
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
            Picture         =   "frmSysDefStn.frx":6632
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "previous"
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrgUp 
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
            Left            =   3840
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefStn.frx":6D34
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "next"
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin Threed.SSPanel pnlPrgDesc 
            Height          =   595
            Left            =   630
            TabIndex        =   62
            ToolTipText     =   "PurgeAir Source Description"
            Top             =   480
            Width           =   3205
            _Version        =   65536
            _ExtentX        =   5653
            _ExtentY        =   1050
            _StockProps     =   15
            Caption         =   "a very very very long purgeair description"
            ForeColor       =   -2147483646
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.19
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
         End
      End
      Begin VB.Frame frmScales 
         Caption         =   "Scale Options"
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
         Height          =   1335
         Left            =   4920
         TabIndex        =   54
         Top             =   3720
         Width           =   3375
         Begin VB.TextBox txtDefPri 
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
            Left            =   2550
            MaxLength       =   1
            TabIndex        =   56
            ToolTipText     =   "Scale Number (1-19)"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtDefAux 
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
            Left            =   2550
            MaxLength       =   1
            TabIndex        =   55
            ToolTipText     =   "Scale Number (1-19)"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblDefPri 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Default Primary Scale #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   525
            Width           =   2300
         End
         Begin VB.Label lblDefAux 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Default Aux Scale #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   885
            Width           =   2300
         End
      End
      Begin VB.Frame frmButane 
         Caption         =   "Butane  Options"
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
         Height          =   1455
         Left            =   4920
         TabIndex        =   49
         Top             =   1200
         Width           =   3375
         Begin VB.TextBox txtBtnMfcDensityMult 
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
            Index           =   1
            Left            =   2640
            TabIndex        =   51
            ToolTipText     =   " MFC Butane Density Multiplier (0.9 - 1.1)"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtBtnMfcDensityMult 
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
            Index           =   2
            Left            =   2640
            TabIndex        =   50
            ToolTipText     =   "ORVR2 MFC Butane Density Multiplier (0.9 - 1.1)"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblBtnMfcDensityMult 
            BackStyle       =   0  'Transparent
            Caption         =   "Mfc Density Multiplier"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   405
            Width           =   2235
         End
         Begin VB.Label lblBtnMfcDensityMult 
            BackStyle       =   0  'Transparent
            Caption         =   "ORVR Mfc Density Multiplier"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   765
            Width           =   2475
         End
      End
      Begin VB.Label lblAbreviation 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Station Abreviation"
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
         Height          =   255
         Left            =   5880
         TabIndex        =   90
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label lblSysID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Station System ID"
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
         Height          =   255
         Left            =   2520
         TabIndex        =   88
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Station Description"
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
         Height          =   255
         Left            =   2520
         TabIndex        =   76
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.Frame frmOptodef 
      Caption         =   "OPTO Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   7580
      Left            =   8760
      TabIndex        =   10
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox ModuleType 
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
         Index           =   0
         ItemData        =   "frmSysDefStn.frx":7436
         Left            =   1200
         List            =   "frmSysDefStn.frx":745A
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   1560
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   1
         ItemData        =   "frmSysDefStn.frx":749C
         Left            =   1200
         List            =   "frmSysDefStn.frx":74C0
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   1920
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   2
         ItemData        =   "frmSysDefStn.frx":7502
         Left            =   1200
         List            =   "frmSysDefStn.frx":7526
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   2280
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   3
         ItemData        =   "frmSysDefStn.frx":7568
         Left            =   1200
         List            =   "frmSysDefStn.frx":758C
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   2640
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   4
         ItemData        =   "frmSysDefStn.frx":75CE
         Left            =   1200
         List            =   "frmSysDefStn.frx":75F2
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   3000
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   5
         ItemData        =   "frmSysDefStn.frx":7634
         Left            =   1200
         List            =   "frmSysDefStn.frx":7658
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   3360
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   6
         ItemData        =   "frmSysDefStn.frx":769A
         Left            =   1200
         List            =   "frmSysDefStn.frx":76BE
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   3720
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   7
         ItemData        =   "frmSysDefStn.frx":7700
         Left            =   1200
         List            =   "frmSysDefStn.frx":7724
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   4080
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   8
         ItemData        =   "frmSysDefStn.frx":7766
         Left            =   1200
         List            =   "frmSysDefStn.frx":778A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   4440
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   9
         ItemData        =   "frmSysDefStn.frx":77CC
         Left            =   1200
         List            =   "frmSysDefStn.frx":77F0
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   4800
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   10
         ItemData        =   "frmSysDefStn.frx":7832
         Left            =   1200
         List            =   "frmSysDefStn.frx":7856
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   5160
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   11
         ItemData        =   "frmSysDefStn.frx":7898
         Left            =   1200
         List            =   "frmSysDefStn.frx":78BC
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   5520
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   12
         ItemData        =   "frmSysDefStn.frx":78FE
         Left            =   1200
         List            =   "frmSysDefStn.frx":7922
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   5880
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   13
         ItemData        =   "frmSysDefStn.frx":7964
         Left            =   1200
         List            =   "frmSysDefStn.frx":7988
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   6240
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   14
         ItemData        =   "frmSysDefStn.frx":79CA
         Left            =   1200
         List            =   "frmSysDefStn.frx":79EE
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   6600
         Width           =   1600
      End
      Begin VB.ComboBox ModuleType 
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
         Index           =   15
         ItemData        =   "frmSysDefStn.frx":7A30
         Left            =   1200
         List            =   "frmSysDefStn.frx":7A54
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   6960
         Width           =   1600
      End
      Begin Threed.SSPanel pnlAddr 
         Height          =   585
         Left            =   1200
         TabIndex        =   27
         ToolTipText     =   "Board Base Address"
         Top             =   480
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "9"
         ForeColor       =   -2147483646
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   3
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   510
         TabIndex        =   44
         Top             =   1605
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   510
         TabIndex        =   43
         Top             =   1965
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   510
         TabIndex        =   42
         Top             =   2325
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   510
         TabIndex        =   41
         Top             =   2685
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   510
         TabIndex        =   40
         Top             =   3045
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   510
         TabIndex        =   39
         Top             =   3405
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   510
         TabIndex        =   38
         Top             =   3765
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   510
         TabIndex        =   37
         Top             =   4125
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   510
         TabIndex        =   36
         Top             =   4485
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   510
         TabIndex        =   35
         Top             =   4845
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   510
         TabIndex        =   34
         Top             =   5205
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   510
         TabIndex        =   33
         Top             =   5565
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   510
         TabIndex        =   32
         Top             =   5925
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   510
         TabIndex        =   31
         Top             =   6285
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   510
         TabIndex        =   30
         Top             =   6645
         Width           =   600
      End
      Begin VB.Label lblSlot 
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   510
         TabIndex        =   29
         Top             =   7005
         Width           =   600
      End
      Begin VB.Label lblAddr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Base Address"
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
         Height          =   555
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame frmSelector 
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdUp 
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
         Left            =   5805
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":7A96
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Next"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDown 
         DisabledPicture =   "frmSysDefStn.frx":8198
         DownPicture     =   "frmSysDefStn.frx":889A
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
         Left            =   3360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":8F9C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Previous"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   840
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
         Left            =   7080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":969E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Function Definition"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.CommandButton cmdSetDefaults 
         Caption         =   "Defaults"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":9DA0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Set Station and Opto to Default Values"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSysDefStn.frx":A292
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Clear Station and Opto Values"
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin Threed.SSPanel pnlStn 
         Height          =   600
         Left            =   4200
         TabIndex        =   9
         ToolTipText     =   "Station Number Displayed"
         Top             =   360
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "Station 9"
         ForeColor       =   -2147483646
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   3
      End
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
      Left            =   10440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefStn.frx":A784
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Next Screen"
      Top             =   8165
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdBack 
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
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefStn.frx":AE86
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Previous Screen"
      Top             =   8165
      UseMaskColor    =   -1  'True
      Width           =   1500
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
      Height          =   1005
      Left            =   1680
      TabIndex        =   77
      Top             =   7805
      Width           =   8655
   End
End
Attribute VB_Name = "frmSysDefStn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
Private GoingToAnotherSysdef As Boolean

Sub Refresh_StationDef()
Dim Idx, slot, maxslot As Integer

    OptoMaxNodeNum = HighestNodeNumber

    For Idx = 0 To 3
        If BoardType.ItemData(Idx) = Node_Info(DefStn) Then BoardType.ListIndex = Idx
    Next Idx
    maxslot = Node_Info(DefStn) - 1

    For slot = 0 To 15
        ModuleType(slot).ListIndex = Opto_Info((DefStn * 4), slot)
    Next slot
    If DefStn > 0 Then
        ' Station Def
        Def_Stn = STN_INFO(DefStn)
        txtDescription.text = Def_Stn.desc
        txtAbreviation.text = Def_Stn.Abrev
        txtSysID.text = Def_Stn.SysID
        If (Def_Stn.Type = STN_LEAKTEST_TYPE) Then
            ' leaktest station
            optLeakTest = True
        Else
            ' canister prg/load station
            DefStnPrg = IIf((Def_Stn.AspiratorNum > 0 And Def_Stn.AspiratorNum < NR_PRGAIR), Def_Stn.AspiratorNum, 1)
            pnlPrgDesc.Caption = PRG_INFO(Def_Stn.AspiratorNum).desc
            txtDefPri.text = Format(Def_Stn.DefPriScale, "0")
            txtDefAux.text = Format(Def_Stn.DefAuxScale, "0")
            txtBtnMfcDensityMult(1).text = Format(Def_Stn.ButMfcDensityMult, "0.0##")
            txtBtnMfcDensityMult(2).text = Format(Def_Stn.ButMfc2DensityMult, "0.0##")
            chkPurgeOven.Value = IIf(Def_Stn.USINGPURGEOVEN, cYES, cNO)
            If Def_Stn.Type = STN_REGULAR_TYPE Then optRegular = True
            If Def_Stn.Type = STN_ORVR_TYPE Then optORVR = True
            If Def_Stn.Type = STN_ORVR2_TYPE Then optORVR2 = True
            If Def_Stn.Type = STN_LIVEFUEL_TYPE Then optLiveFuel = True
            If Def_Stn.Type = STN_LIVEREG_TYPE Then optRegLive = True
            If Def_Stn.Type = STN_LIVEORVR2_TYPE Then optOrvr2Live = True
            If Def_Stn.Type = STN_COMBO3_TYPE Then optCombo3 = True
            If Def_Stn.Type = STN_LEAKTEST_TYPE Then optLeakTest = True
            If Def_Stn.Type = STN_DUMMY_TYPE Then optDummy = True
            If Def_Stn.Type = STN_LIVEFUEL_TYPE _
                Or Def_Stn.Type = STN_LIVEREG_TYPE _
                Or Def_Stn.Type = STN_LIVEORVR2_TYPE Then
                    txtTank.text = Format(Def_Stn.ADF_StnNum, "0")
                    chkADF.Value = IIf(Def_Stn.ADF_DEF.hasAUTODRAINFILL, cYES, cNO)
                    chkVapor.Value = IIf(Def_Stn.ADF_DEF.hasADF_VaporValve, cYES, cNO)
                    chkHeater.Value = IIf(Def_Stn.ADF_DEF.hasADF_Heater, cYES, cNO)
                    chkLevelXmtr.Value = IIf(Def_Stn.ADF_DEF.hasADF_LT, cYES, cNO)
                    chkFuelStorage.Value = IIf(Def_Stn.ADF_DEF.hasADF_FST, cYES, cNO)
                    chkWaterBath.Value = IIf(Def_Stn.ADF_DEF.hasADF_WaterBath, cYES, cNO)
            End If
       End If
    Else
        ' Common Def
    End If
    
    lblMsg.Caption = " "
End Sub

Sub Update_StationDef()
'
Dim slot As Integer

    pnlAddr.Caption = Format((DefStn * 4), "#0")
    
    For slot = 0 To 15
        If (BoardType.ItemData(BoardType.ListIndex)) > slot Then
            lblSlot(slot).Visible = True
            ModuleType(slot).Visible = True
        Else
            lblSlot(slot).Visible = False
            ModuleType(slot).Visible = False
        End If
    Next slot
    
    
    If DefStn = 0 Then
    
        pnlStn.Caption = "Common"
        frmComdef.Visible = True
        cmdPurgeCfg.Enabled = True
        cmdPurgeCfg.Visible = True
        cmdScaleCfg.Enabled = True
        cmdScaleCfg.Visible = True
        frmStnDef.Visible = False
    
    Else
    
        pnlStn.Caption = "Station " & Format(DefStn, "0")
        frmComdef.Visible = False
        cmdPurgeCfg.Enabled = False
        cmdPurgeCfg.Visible = False
        cmdScaleCfg.Enabled = False
        cmdScaleCfg.Visible = False
        frmStnDef.Visible = True
        If optLeakTest Then
            frmLiveFuel.Visible = False
            frmAspirator.Visible = False
            frmScales.Visible = False
            frmButane.Visible = False
            frmPurge.Visible = False
        End If
        If optDummy Then
            frmLiveFuel.Visible = False
            frmAspirator.Visible = False
            frmScales.Visible = False
            frmButane.Visible = False
            frmPurge.Visible = False
        End If
        If optLiveFuel Then
            frmLiveFuel.Visible = True
            frmAspirator.Visible = True
            frmScales.Visible = True
            frmButane.Visible = False
            frmPurge.Visible = True
        End If
        If optOrvr2Live Then
            frmLiveFuel.Visible = True
            frmAspirator.Visible = True
            frmScales.Visible = True
            frmButane.Visible = False
            frmPurge.Visible = True
            lblBtnMfcDensityMult(2).Visible = True
            txtBtnMfcDensityMult(2).Visible = True
        End If
        If optORVR Then
            frmLiveFuel.Visible = False
            frmAspirator.Visible = True
            frmScales.Visible = True
            frmButane.Visible = True
            frmPurge.Visible = True
            lblBtnMfcDensityMult(2).Visible = False
            txtBtnMfcDensityMult(2).Visible = False
        End If
        If optORVR2 Then
            frmLiveFuel.Visible = False
            frmAspirator.Visible = True
            frmScales.Visible = True
            frmButane.Visible = True
            frmPurge.Visible = True
            lblBtnMfcDensityMult(2).Visible = True
            txtBtnMfcDensityMult(2).Visible = True
        End If
        If optRegular Then
            frmLiveFuel.Visible = False
            frmAspirator.Visible = True
            frmScales.Visible = True
            frmButane.Visible = True
            frmPurge.Visible = True
            lblBtnMfcDensityMult(2).Visible = False
            txtBtnMfcDensityMult(2).Visible = False
        End If
        If optRegLive Then
            frmLiveFuel.Visible = True
            frmAspirator.Visible = True
            frmScales.Visible = True
            frmButane.Visible = True
            frmPurge.Visible = True
            lblBtnMfcDensityMult(2).Visible = False
            txtBtnMfcDensityMult(2).Visible = False
        End If
        If optCombo3 Then
            ' future
        End If
        
    End If
    
    lblMsg.Caption = " "

End Sub

Private Function HighestNodeNumber() As Integer
Dim Idx As Integer
Dim iNode As Integer
    Idx = 0
    For iNode = 0 To MAX_NODE
        If (Node_Info(iNode) > 0) Then
            Idx = iNode
        End If
    Next iNode
    HighestNodeNumber = (Idx * 4) + 3
End Function

Private Function EncodeAdfDef() As Integer
Dim whichADF As Integer
    whichADF = 0
    If chkADF.Value = cYES Then
        If (chkVapor.Value = cNO And chkHeater.Value = cNO And chkWaterBath.Value = cNO And chkLevelXmtr.Value = cNO And chkFuelStorage.Value = cNO) Then
            whichADF = 1        '(Mark IV)
        ElseIf (chkVapor.Value = cYES And chkHeater.Value = cYES And chkWaterBath.Value = cNO And chkLevelXmtr.Value = cNO And chkFuelStorage.Value = cNO) Then
            whichADF = 12       '(Mahle)
        ElseIf (chkVapor.Value = cYES And chkHeater.Value = cYES And chkWaterBath.Value = cNO And chkLevelXmtr.Value = cYES And chkFuelStorage.Value = cNO) Then
            whichADF = 20       '(Stant)
        ElseIf (chkVapor.Value = cYES And chkHeater.Value = cNO And chkWaterBath.Value = cNO And chkLevelXmtr.Value = cYES And chkFuelStorage.Value = cYES) Then
            whichADF = 22       '(Chrysler)
        End If
    ElseIf (chkADF.Value = cNO And chkVapor.Value = cNO And chkHeater.Value = cNO And chkWaterBath.Value = cYES And chkLevelXmtr.Value = cNO And chkFuelStorage.Value = cNO) Then
        whichADF = 90           '(Honda R&D)
    End If
    EncodeAdfDef = whichADF
End Function

Private Sub BoardType_Click()
    Update_StationDef
End Sub

Private Sub cmdBack_Click()
    frmSysDefMain.Show
    GoingToAnotherSysdef = True
    Unload Me
End Sub

Private Sub cmdClear_Click()

    If CheckPass("9", True) Then
    
        If DefStn = 0 Then
        
            ' Set Common Defaults
            BoardType.ListIndex = 2
            ModuleType(0).ListIndex = 0
            ModuleType(1).ListIndex = 0
            ModuleType(2).ListIndex = 0
            ModuleType(3).ListIndex = 0
            ModuleType(4).ListIndex = 0
            ModuleType(5).ListIndex = 0
            ModuleType(6).ListIndex = 0
            ModuleType(7).ListIndex = 0
            ModuleType(8).ListIndex = 0
            ModuleType(9).ListIndex = 0
            ModuleType(10).ListIndex = 0
            ModuleType(11).ListIndex = 0
            ModuleType(12).ListIndex = 0
            ModuleType(13).ListIndex = 0
            ModuleType(14).ListIndex = 0
            ModuleType(15).ListIndex = 0
                    
            ' Save the Default Values
            cmdSave_Click
            Delay_Box "Common Info Cleared", MSGDELAY, msgSHOW
            lblMsg.Caption = "Common Info Cleared"
            
        Else
        
            ' Set Station Defaults
            txtDescription = "Station #" & Format(DefStn, "0")
            optDummy = True
            txtDefPri.text = Format(DefStn, "0")
            txtDefAux.text = Format(DefStn, "0")
            txtBtnMfcDensityMult(1).text = "1.0"
            txtBtnMfcDensityMult(2).text = "1.0"
            DefStnPrg = 1
            BoardType.ListIndex = 2
            ModuleType(0).ListIndex = 0
            ModuleType(1).ListIndex = 0
            ModuleType(2).ListIndex = 0
            ModuleType(3).ListIndex = 0
            ModuleType(4).ListIndex = 0
            ModuleType(5).ListIndex = 0
            ModuleType(6).ListIndex = 0
            ModuleType(7).ListIndex = 0
            ModuleType(8).ListIndex = 0
            ModuleType(9).ListIndex = 0
            ModuleType(10).ListIndex = 0
            ModuleType(11).ListIndex = 0
            ModuleType(12).ListIndex = 0
            ModuleType(13).ListIndex = 0
            ModuleType(14).ListIndex = 0
            ModuleType(15).ListIndex = 0
                    
            ' Save the Default Values
            cmdSave_Click
            Delay_Box "Station Info Cleared", MSGDELAY, msgSHOW
            lblMsg.Caption = vbCrLf & "Station Info Cleared"
            
        End If
            
    Else
        Delay_Box "Insufficient Access", MSGDELAY, msgSHOW
        lblMsg.Caption = vbCrLf & "Insufficient Access"
    End If

End Sub

Private Sub cmdDown_Click()
    If DefStn > 0 Then
        DefStn = DefStn - 1
    Else
        DefStn = NR_STN
    End If
    Refresh_StationDef
    Update_StationDef
End Sub

Private Sub cmdNext_Click()
    frmSysDefFunc.Show
    GoingToAnotherSysdef = True
    Unload Me
End Sub

Private Sub cmdPrgDn_Click()
    If DefStnPrg > 1 Then
        DefStnPrg = DefStnPrg - 1
    Else
        DefStnPrg = NR_PRGAIR
    End If
    pnlPrgDesc.Caption = PRG_INFO(DefStnPrg).desc
End Sub

Private Sub cmdPrgUp_Click()
    If DefStnPrg < NR_PRGAIR Then
        DefStnPrg = DefStnPrg + 1
    Else
        DefStnPrg = 1
    End If
    pnlPrgDesc.Caption = PRG_INFO(DefStnPrg).desc
End Sub

Private Sub cmdPurgeCfg_Click()
    If CheckPass("5", True) Then
        frmPurgeAir.Show
    End If
End Sub

Private Sub cmdSave_Click()
Dim baseaddr As Integer
Dim chan As Integer
Dim slot As Integer

If CheckPass("9", True) Then

    Node_Info(DefStn) = BoardType.ItemData(BoardType.ListIndex)
    baseaddr = DefStn * 4
    For slot = 0 To MAX_SLOT
        Opto_Info(baseaddr, slot) = ModuleType(slot).ListIndex
    Next slot
    If DefStn > 0 Then
        Def_Stn.desc = ((Trim(txtDescription.text)) & " ")
        Def_Stn.Abrev = Trim(txtAbreviation.text)
        Def_Stn.SysID = Trim(txtSysID.text)
        If optLeakTest Then
            Def_Stn.Type = STN_LEAKTEST_TYPE
            Def_Stn.ADF_StnNum = 0
            Def_Stn.ADF_TANKTYPE = 0
            Def_Stn.ButMfcDensityMult = 1#
            Def_Stn.ButMfc2DensityMult = 1#
            Def_Stn.ADF_DEF.hasLIVEFUEL = False
            Def_Stn.AspiratorNum = 0
            Def_Stn.USINGPURGEOVEN = False
        ElseIf optDummy Then
            Def_Stn.Type = STN_DUMMY_TYPE
            Def_Stn.ADF_StnNum = 0
            Def_Stn.ADF_TANKTYPE = 0
            Def_Stn.ButMfcDensityMult = 1#
            Def_Stn.ButMfc2DensityMult = 1#
            Def_Stn.ADF_DEF.hasLIVEFUEL = False
            Def_Stn.AspiratorNum = 0
            Def_Stn.USINGPURGEOVEN = False
        Else
            Def_Stn.AspiratorNum = DefStnPrg
            Def_Stn.USINGPURGEOVEN = IIf((chkPurgeOven.Value = cYES), True, False)
            If optLiveFuel Then
                Def_Stn.Type = STN_LIVEFUEL_TYPE
                If IsNumeric(txtTank.text) Then
                    Def_Stn.ADF_StnNum = CInt(txtTank.text)
                Else
                    Def_Stn.ADF_StnNum = 0
                End If
                Def_Stn.ADF_TANKTYPE = EncodeAdfDef
                Def_Stn.ButMfcDensityMult = 1#
                Def_Stn.ButMfc2DensityMult = 1#
                Def_Stn.ADF_DEF.hasLIVEFUEL = True
            End If
            If optRegLive Then
                Def_Stn.Type = STN_LIVEREG_TYPE
                If IsNumeric(txtTank.text) Then
                    Def_Stn.ADF_StnNum = CInt(txtTank.text)
                Else
                    Def_Stn.ADF_StnNum = 0
                End If
                Def_Stn.ADF_TANKTYPE = EncodeAdfDef
                ' Butane MFC Density Multiplier
                '   max     1.1
                '   default 1.0
                '   min     0.9
                If IsNumeric(txtBtnMfcDensityMult(1).text) Then
                    If CSng(txtBtnMfcDensityMult(1).text) < 0.9 Then txtBtnMfcDensityMult(1).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(1).text) > 1.1 Then txtBtnMfcDensityMult(1).text = "1.1"
                Else
                    txtBtnMfcDensityMult(1).text = "1.0"
                End If
                Def_Stn.ButMfcDensityMult = CSng(txtBtnMfcDensityMult(1).text)
                Def_Stn.ButMfc2DensityMult = 1#
                Def_Stn.ADF_DEF.hasLIVEFUEL = True
            End If
            If optOrvr2Live Then
                Def_Stn.Type = STN_LIVEORVR2_TYPE
                If IsNumeric(txtTank.text) Then
                    Def_Stn.ADF_StnNum = CInt(txtTank.text)
                Else
                    Def_Stn.ADF_StnNum = 0
                End If
                Def_Stn.ADF_TANKTYPE = EncodeAdfDef
                If IsNumeric(txtBtnMfcDensityMult(1).text) Then
                    If CSng(txtBtnMfcDensityMult(1).text) < 0.9 Then txtBtnMfcDensityMult(1).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(1).text) > 1.1 Then txtBtnMfcDensityMult(1).text = "1.1"
                Else
                    txtBtnMfcDensityMult(1).text = "1.0"
                End If
                Def_Stn.ButMfcDensityMult = CSng(txtBtnMfcDensityMult(1).text)
                ' Butane ORVR2 MFC Density Multiplier
                '   max     1.1
                '   default 1.0
                '   min     0.9
                If IsNumeric(txtBtnMfcDensityMult(2).text) Then
                    If CSng(txtBtnMfcDensityMult(2).text) < 0.9 Then txtBtnMfcDensityMult(2).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(2).text) > 1.1 Then txtBtnMfcDensityMult(2).text = "1.1"
                Else
                    txtBtnMfcDensityMult(2).text = "1.0"
                End If
                Def_Stn.ADF_DEF.hasLIVEFUEL = True
                
            End If
            If optORVR Then
                Def_Stn.Type = STN_ORVR_TYPE
                Def_Stn.ADF_StnNum = 0
                Def_Stn.ADF_TANKTYPE = 0
                ' Butane MFC Density Multiplier
                '   max     1.1
                '   default 1.0
                '   min     0.9
                If IsNumeric(txtBtnMfcDensityMult(1).text) Then
                    If CSng(txtBtnMfcDensityMult(1).text) < 0.9 Then txtBtnMfcDensityMult(1).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(1).text) > 1.1 Then txtBtnMfcDensityMult(1).text = "1.1"
                Else
                    txtBtnMfcDensityMult(1).text = "1.0"
                End If
                Def_Stn.ButMfcDensityMult = CSng(txtBtnMfcDensityMult(1).text)
                Def_Stn.ButMfc2DensityMult = 1#
                Def_Stn.ADF_DEF.hasLIVEFUEL = False
            End If
            If optORVR2 Then
                Def_Stn.Type = STN_ORVR2_TYPE
                Def_Stn.ADF_StnNum = 0
                Def_Stn.ADF_TANKTYPE = 0
                ' Butane MFC Density Multiplier
                '   max     1.1
                '   default 1.0
                '   min     0.9
                If IsNumeric(txtBtnMfcDensityMult(1).text) Then
                    If CSng(txtBtnMfcDensityMult(1).text) < 0.9 Then txtBtnMfcDensityMult(1).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(1).text) > 1.1 Then txtBtnMfcDensityMult(1).text = "1.1"
                Else
                    txtBtnMfcDensityMult(1).text = "1.0"
                End If
                Def_Stn.ButMfcDensityMult = CSng(txtBtnMfcDensityMult(1).text)
                ' Butane ORVR2 MFC Density Multiplier
                '   max     1.1
                '   default 1.0
                '   min     0.9
                If IsNumeric(txtBtnMfcDensityMult(2).text) Then
                    If CSng(txtBtnMfcDensityMult(2).text) < 0.9 Then txtBtnMfcDensityMult(2).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(2).text) > 1.1 Then txtBtnMfcDensityMult(2).text = "1.1"
                Else
                    txtBtnMfcDensityMult(2).text = "1.0"
                End If
                Def_Stn.ButMfc2DensityMult = CSng(txtBtnMfcDensityMult(2).text)
                Def_Stn.ADF_DEF.hasLIVEFUEL = False
            End If
            If optRegular Then
                Def_Stn.Type = STN_REGULAR_TYPE
                Def_Stn.ADF_StnNum = 0
                Def_Stn.ADF_TANKTYPE = 0
                ' Butane MFC Density Multiplier
                '   max     1.1
                '   default 1.0
                '   min     0.9
                If IsNumeric(txtBtnMfcDensityMult(1).text) Then
                    If CSng(txtBtnMfcDensityMult(1).text) < 0.9 Then txtBtnMfcDensityMult(1).text = "0.9"
                    If CSng(txtBtnMfcDensityMult(1).text) > 1.1 Then txtBtnMfcDensityMult(1).text = "1.1"
                Else
                    txtBtnMfcDensityMult(1).text = "1.0"
                End If
                Def_Stn.ButMfcDensityMult = CSng(txtBtnMfcDensityMult(1).text)
                Def_Stn.ButMfc2DensityMult = 1#
                Def_Stn.ADF_DEF.hasLIVEFUEL = False
            End If
            If optCombo3 Then
                Def_Stn.Type = STN_COMBO3_TYPE
                Def_Stn.ADF_StnNum = 0
                Def_Stn.ADF_TANKTYPE = 0
                ' Butane MFC Density Multiplier
                Def_Stn.ButMfcDensityMult = 1#
                Def_Stn.ButMfc2DensityMult = 1#
                Def_Stn.ADF_DEF.hasLIVEFUEL = False
            End If
        End If
        If USINGHARDPIPEDSCALES Then
            ' two scales per station, fixed assignments for pri & aux for each station; stn#1 pri = 1, stn#1 aux = 2, etc.
            Def_Stn.DefPriScale = 1 + (2 * (DefStn - 1))
            Def_Stn.DefAuxScale = 1 + Def_Stn.DefPriScale
        Else
            ' one scale per station, selectable as pri or aux by any station
            If IsNumeric(txtDefPri.text) Then
                If CInt(txtDefPri.text) > 0 And CInt(txtDefPri.text) <= NR_SCALES Then
                    Def_Stn.DefPriScale = CInt(txtDefPri.text)
                Else
                    Def_Stn.DefPriScale = DefStn
                End If
            Else
                Def_Stn.DefPriScale = DefStn
            End If
            If IsNumeric(txtDefAux.text) Then
                If CInt(txtDefAux.text) > 0 And CInt(txtDefAux.text) <= NR_SCALES Then
                    Def_Stn.DefAuxScale = CInt(txtDefAux.text)
                Else
                    Def_Stn.DefAuxScale = DefStn
                End If
            Else
                Def_Stn.DefAuxScale = DefStn
            End If
        End If
        ' Log any Butane Density Multiplier Changes
        Select Case Def_Stn.Type
            Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                If Def_Stn.ButMfcDensityMult <> STN_INFO(DefStn).ButMfcDensityMult Then
                    Write_ELog "Station #" & Format(DefStn, "0") & " Butane MFC Density Multiplier changed to " & Format(Def_Stn.ButMfcDensityMult, "0.0##")
                End If
            Case STN_LIVEREG_TYPE
                If Def_Stn.ButMfcDensityMult <> STN_INFO(DefStn).ButMfcDensityMult Then
                    Write_ELog "Station #" & Format(DefStn, "0") & " Butane MFC Density Multiplier changed to " & Format(Def_Stn.ButMfcDensityMult, "0.0##")
                End If
            Case STN_ORVR2_TYPE
                If Def_Stn.ButMfcDensityMult <> STN_INFO(DefStn).ButMfcDensityMult Then
                    Write_ELog "Station #" & Format(DefStn, "0") & " Butane MFC Density Multiplier changed to " & Format(Def_Stn.ButMfcDensityMult, "0.0##")
                End If
                If Def_Stn.ButMfc2DensityMult <> STN_INFO(DefStn).ButMfc2DensityMult Then
                    Write_ELog "Station #" & Format(DefStn, "0") & " Butane ORVR2 MFC Density Multiplier changed to " & Format(Def_Stn.ButMfc2DensityMult, "0.0##")
                End If
            Case Else
                ' nothing to do
        End Select
        
        ' Decode LiveFuel ADF Options
        If Def_Stn.ADF_DEF.hasLIVEFUEL Then
            Def_Stn.ADF_DEF.hasADF_FST = IIf(chkFuelStorage = cYES, True, False)
            Def_Stn.ADF_DEF.hasADF_Heater = IIf(chkHeater = cYES, True, False)
            Def_Stn.ADF_DEF.hasADF_LT = IIf(((Def_Stn.ADF_TANKTYPE >= 20) And (Def_Stn.ADF_TANKTYPE < 90)), True, False)
            Def_Stn.ADF_DEF.hasADF_PS = IIf(Def_Stn.ADF_TANKTYPE = 12, True, False)
            Def_Stn.ADF_DEF.hasADF_VaporValve = IIf(chkVapor = cYES, True, False)
            Def_Stn.ADF_DEF.hasADF_WaterBath = IIf(chkWaterBath = cYES, True, False)
            Def_Stn.ADF_DEF.hasAUTODRAINFILL = IIf(chkADF = cYES, True, False)
            Def_Stn.ADF_DEF.TankNum = CInt(ValueFromText(txtTank.text))
        Else
            Def_Stn.ADF_DEF.hasADF_FST = False
            Def_Stn.ADF_DEF.hasADF_Heater = False
            Def_Stn.ADF_DEF.hasADF_LT = False
            Def_Stn.ADF_DEF.hasADF_PS = False
            Def_Stn.ADF_DEF.hasADF_VaporValve = False
            Def_Stn.ADF_DEF.hasADF_WaterBath = False
            Def_Stn.ADF_DEF.hasAUTODRAINFILL = False
            Def_Stn.ADF_DEF.TankNum = CInt(0)
        End If
        ' Apply Changes
        STN_INFO(DefStn) = Def_Stn
        AdfDef(DefStn) = STN_INFO(DefStn).ADF_DEF
        AdfControl(DefStn).AdfDefinition = STN_INFO(DefStn).ADF_DEF
    End If
    
    ' Apply Opto Changes
    SetupOpto
    ' Save Station Information
    Save_StationInfo
    ' Save Opto Information
    Save_OptoInfo
    ' Save Node Information
    Save_NodeInfo
    ' Refresh screen values
    Refresh_StationDef
    ' message to user
'    Delay_Box "Station and Opto Info Files Saved", MSGDELAY, msgSHOW
    lblMsg.Caption = vbCrLf & "Station and Opto Info Files Saved"

End If

End Sub

Private Sub cmdScaleCfg_Click()
    If CheckPass("5", True) Then
        frmComm8Card.Show
    End If
End Sub

Private Sub cmdSetDefaults_Click()
'
    If CheckPass("9", True) Then
    
        If DefStn = 0 Then
        
            ' Set Common Defaults
            BoardType.ListIndex = 1
            ModuleType(0).ListIndex = 1
            ModuleType(1).ListIndex = 1
            ModuleType(2).ListIndex = 0
            ModuleType(3).ListIndex = 2
            ModuleType(4).ListIndex = 2
            ModuleType(5).ListIndex = 0
            ModuleType(6).ListIndex = 0
            ModuleType(7).ListIndex = 0
            ModuleType(8).ListIndex = 0
            ModuleType(9).ListIndex = 0
            ModuleType(10).ListIndex = 3
            ModuleType(11).ListIndex = 3
            ModuleType(12).ListIndex = 0
            ModuleType(13).ListIndex = 0
            ModuleType(14).ListIndex = 0
            ModuleType(15).ListIndex = 0
                   
            ' Save the Default Values
            cmdSave_Click
            Delay_Box "Common Info set to Defaults", MSGDELAY, msgSHOW
            lblMsg.Caption = vbCrLf & "Common Info set to Defaults"
            
        Else
        
            ' Set Station Defaults
            txtDescription = "Station #" & Format(DefStn, "0")
            optRegular = True
            txtDefPri.text = Format(DefStn, "0")
            txtDefAux.text = Format(DefStn, "0")
            txtBtnMfcDensityMult(1).text = "1.0"
            txtBtnMfcDensityMult(2).text = "1.0"
            DefStnPrg = 1
            BoardType.ListIndex = 2
            ModuleType(0).ListIndex = 2
            ModuleType(1).ListIndex = 2
            ModuleType(2).ListIndex = 2
            ModuleType(3).ListIndex = 2
            ModuleType(4).ListIndex = 1
            ModuleType(5).ListIndex = 0
            ModuleType(6).ListIndex = 4
            ModuleType(7).ListIndex = 4
            ModuleType(8).ListIndex = 0
            ModuleType(9).ListIndex = 3
            ModuleType(10).ListIndex = 3
            ModuleType(11).ListIndex = 0
            ModuleType(12).ListIndex = 0
            ModuleType(13).ListIndex = 0
            ModuleType(14).ListIndex = 0
            ModuleType(15).ListIndex = 0
            
            ' Save the Default Values
            cmdSave_Click
            Delay_Box "Station Info set to Defaults", MSGDELAY, msgSHOW
            lblMsg.Caption = vbCrLf & "Station Info set to Defaults"
            
        End If
    End If
End Sub

Private Sub cmdUp_Click()
    If DefStn < NR_STN Then
        DefStn = DefStn + 1
    Else
        DefStn = 0
    End If
    Refresh_StationDef
    Update_StationDef
End Sub

Private Sub Form_Load()
    GoingToAnotherSysdef = False
'    If Not ReadyToRun Then
'        lblMsg.ForeColor = lblMsg.BackColor
'    Else
        lblMsg.ForeColor = Message_ForeColor
'    End If

    chkWaterBath.Visible = IIf(USINGWATERBATH, cYES, cNO)

    ' Set Title Foreground color
    frmSelector.ForeColor = Titles_ForeColor
    pnlStn.ForeColor = TitlesData_Forecolor
    frmStnDef.ForeColor = Titles_ForeColor
    frmStnType.ForeColor = TitlesLabel_ForeColor
    frmAspirator.ForeColor = TitlesLabel_ForeColor
    pnlPrgDesc.ForeColor = TitlesData_Forecolor
    frmLiveFuel.ForeColor = TitlesLabel_ForeColor
    frmButane.ForeColor = TitlesLabel_ForeColor
    frmScales.ForeColor = TitlesLabel_ForeColor
    frmComdef.ForeColor = Titles_ForeColor
    frmOptodef.ForeColor = Titles_ForeColor
    pnlAddr.ForeColor = TitlesData_Forecolor

    ' Set Background colors
    pnlAddr.BackColor = EntryNotChangeable_BackColor
    pnlPrgDesc.BackColor = EntryNotChangeable_BackColor
    pnlStn.BackColor = EntryNotChangeable_BackColor

    lblMsg.Caption = " "

    Form_Center Me

    Refresh_StationDef
    Update_StationDef
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not GoingToAnotherSysdef Then ReadyToRun = True
    Unload Me
End Sub

Private Sub optDummy_Click()
    Update_StationDef
End Sub

Private Sub optLeakTest_Click()
    Update_StationDef
End Sub

Private Sub optLiveFuel_Click()
    Update_StationDef
End Sub

Private Sub optORVR_Click()
    Update_StationDef
End Sub

Private Sub optORVR2_Click()
    Update_StationDef
End Sub

Private Sub optORVR2Live_Click()
    Update_StationDef
End Sub

Private Sub optRegLive_Click()
    Update_StationDef
End Sub

Private Sub optRegular_Click()
    Update_StationDef
End Sub

Private Sub optCombo3_Click()
    Update_StationDef
End Sub



