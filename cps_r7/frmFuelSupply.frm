VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmFuelSupply 
   Caption         =   "Fuel Supply"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "frmFuelSupply.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFstTask 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
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
      Height          =   360
      Left            =   3750
      TabIndex        =   56
      Text            =   "Storage Tank Drain & Fill Task Desc."
      Top             =   1515
      Width           =   2365
   End
   Begin VB.TextBox txtFstMessage 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   3750
      TabIndex        =   55
      Text            =   "Storage Tank Drain & Fill Message Window"
      Top             =   1875
      Width           =   2365
   End
   Begin VB.Frame frmADF_Controls 
      Caption         =   "Vapor Generator Drain && Fill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2055
      Left            =   240
      TabIndex        =   47
      Top             =   6720
      Width           =   5775
      Begin VB.CommandButton cmdIgnore 
         Caption         =   "Ignore"
         DisabledPicture =   "frmFuelSupply.frx":57E2
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":5B24
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Proceed Anyway"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         DisabledPicture =   "frmFuelSupply.frx":5E66
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":61A8
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Pause the Sequence"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         DisabledPicture =   "frmFuelSupply.frx":64EA
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":682C
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Stop the Sequence"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtMessage 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   360
         Left            =   120
         TabIndex        =   53
         Text            =   "Auto Drain & Fill Message Window"
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox txtTask 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Height          =   360
         Left            =   120
         TabIndex        =   52
         Text            =   "Auto Drain & Fill Task Desc."
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Stop"
         DisabledPicture =   "frmFuelSupply.frx":6B6E
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":6EB0
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Stop the Sequence"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdRetry 
         Caption         =   "Retry"
         DisabledPicture =   "frmFuelSupply.frx":71F2
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":7534
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Restart the Sequence"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
         DisabledPicture =   "frmFuelSupply.frx":7876
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":7BB8
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Manual Refill Complete"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdDrain 
         Caption         =   "Drain"
         DisabledPicture =   "frmFuelSupply.frx":7EFA
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":823C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Drain Vapor Tank"
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtPipe_2Can1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3375
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   3120
      Width           =   1440
   End
   Begin VB.TextBox txtPipe_ADF2Waste2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   34
      ToolTipText     =   "High Level Switch"
      Top             =   6120
      Width           =   405
   End
   Begin VB.TextBox txtPipe_FS2Waste2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3120
      Width           =   405
   End
   Begin VB.TextBox txtPipe_FromSupply1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1905
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   360
      Width           =   2070
   End
   Begin Threed.SSPanel pnlADF_Tank 
      Height          =   1965
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Live Fuel Tank Display Indicators"
      Top             =   3840
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   3466
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      FloodColor      =   4210816
      Alignment       =   2
      Begin VB.Frame frmHeater 
         BackColor       =   &H00808080&
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
         Height          =   300
         Left            =   2760
         TabIndex        =   41
         Top             =   1230
         Width           =   2580
         Begin VB.Label ADF_Temp 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "888.8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1365
            TabIndex        =   45
            Top             =   0
            Width           =   645
         End
         Begin VB.Label ADF_TempSP 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "888.8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   720
            TabIndex        =   44
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lbl_ADF_TempUnits 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "deg C"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2040
            TabIndex        =   43
            Top             =   30
            Width           =   555
         End
         Begin VB.Label lblADF_Temp 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Temp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   42
            Tag             =   "deg"
            ToolTipText     =   "Fuel Temperature"
            Top             =   30
            Width           =   600
         End
      End
      Begin VB.TextBox txtADF_Safety 
         BackColor       =   &H0000C000&
         Height          =   315
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "Safety Level Switch"
         Top             =   863
         Width           =   400
      End
      Begin VB.TextBox txtADF_Heater 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Heater On"
         ToolTipText     =   "Heater Status"
         Top             =   863
         Width           =   2205
      End
      Begin VB.TextBox txtADF_Fill 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Fill Open"
         ToolTipText     =   "Fill Valve Status"
         Top             =   60
         Width           =   1365
      End
      Begin VB.TextBox txtADF_HiHi 
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "HiHi Level Switch"
         Top             =   60
         Width           =   400
      End
      Begin VB.TextBox txtADF_Vapor 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "VaporClosed"
         ToolTipText     =   "Vapor Valve Status"
         Top             =   60
         Width           =   2205
      End
      Begin VB.TextBox txtN2PS 
         BackColor       =   &H0000C000&
         Height          =   315
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "N2 Pressure Switch"
         Top             =   60
         Width           =   400
      End
      Begin VB.TextBox txtADF_Circ 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Circulating"
         ToolTipText     =   "Fuel Pump Status"
         Top             =   1560
         Width           =   2205
      End
      Begin VB.TextBox txtADF_Drain 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "DrainClosed"
         ToolTipText     =   "Drain Valve Status"
         Top             =   1590
         Width           =   1365
      End
      Begin VB.TextBox txtADF_Low 
         BackColor       =   &H0000C000&
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Low Level Switch"
         Top             =   1215
         Width           =   400
      End
      Begin VB.Label lblFuelStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel is Weak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   780
         TabIndex        =   65
         Top             =   1230
         Width           =   1365
      End
      Begin VB.Label ADF_Sheath 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "888.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         TabIndex        =   63
         Top             =   900
         Width           =   645
      End
      Begin VB.Label ADF_MAX 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "888.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         ToolTipText     =   "configured cycles between refills"
         Top             =   510
         Width           =   645
      End
      Begin VB.Label ADF_ACTUAL 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "888.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4125
         TabIndex        =   39
         ToolTipText     =   "cycles since last refill"
         Top             =   510
         Width           =   645
      End
      Begin VB.Label lblADF_Cycles 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Load Cycles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2760
         TabIndex        =   38
         Top             =   390
         Width           =   600
      End
      Begin VB.Label lbl_ADF_CyclesUnits 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "since refill"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4800
         TabIndex        =   37
         Top             =   390
         Width           =   600
      End
      Begin VB.Label lblADF_TankDesc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " VAPOR GENERATOR TANK "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1050
         TabIndex        =   23
         Tag             =   "deg"
         ToolTipText     =   "Fuel Temperature"
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblADF_Level 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "188.8 %"
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
         Left            =   60
         TabIndex        =   21
         Top             =   900
         Width           =   900
      End
   End
   Begin Threed.SSPanel pnlFS_Tank 
      Height          =   1755
      Left            =   480
      TabIndex        =   2
      Top             =   1020
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   3096
      _StockProps     =   15
      ForeColor       =   65280
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   0
      FloodColor      =   49152
      FloodShowPct    =   0   'False
      Alignment       =   1
      Begin VB.CommandButton cmdStop2 
         DisabledPicture =   "frmFuelSupply.frx":857E
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":88C0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   705
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtFS_Fill 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Fill Open"
         ToolTipText     =   "Fill Valve Status"
         Top             =   60
         Width           =   1365
      End
      Begin VB.TextBox txtFS_HiHi 
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "HiHi Level Switch"
         Top             =   60
         Width           =   400
      End
      Begin VB.TextBox txtFS_Drain 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "DrainClosed"
         ToolTipText     =   "Drain Valve Status"
         Top             =   1380
         Width           =   1365
      End
      Begin VB.TextBox txtFS_Low 
         BackColor       =   &H0000C000&
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Low Level Switch"
         Top             =   1065
         Width           =   400
      End
      Begin VB.CommandButton cmdDrain2 
         DisabledPicture =   "frmFuelSupply.frx":8C02
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":8F44
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1380
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdFill2 
         DisabledPicture =   "frmFuelSupply.frx":9286
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFuelSupply.frx":95C8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFS_Level 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "188.8 %"
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
         Left            =   60
         TabIndex        =   20
         Top             =   750
         Width           =   900
      End
      Begin VB.Label lblFS_TankDesc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " FUEL STORAGE TANK "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   1055
         TabIndex        =   12
         Tag             =   "deg"
         ToolTipText     =   "Fuel Temperature"
         Top             =   555
         Width           =   960
      End
   End
   Begin VB.TextBox txtPipe_FS2ADF 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   1905
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   195
   End
   Begin Threed.SSPanel pnlFS_Level 
      Height          =   1755
      Left            =   180
      TabIndex        =   3
      Top             =   1020
      Width           =   240
      _Version        =   65536
      _ExtentX        =   423
      _ExtentY        =   3096
      _StockProps     =   15
      Caption         =   "88.8"
      ForeColor       =   65280
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      FloodType       =   4
      FloodColor      =   49152
      FloodShowPct    =   0   'False
   End
   Begin VB.TextBox txtPipe_FromSupply2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   1905
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   195
   End
   Begin VB.Timer tmrScreen 
      Interval        =   150
      Left            =   480
      Top             =   360
   End
   Begin Threed.SSPanel pnlADF_Level 
      Height          =   1965
      Left            =   180
      TabIndex        =   22
      Top             =   3840
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   3466
      _StockProps     =   15
      ForeColor       =   65280
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      FloodType       =   4
      FloodColor      =   49152
      FloodShowPct    =   0   'False
   End
   Begin VB.TextBox txtPipe_FS2Waste1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2400
      Width           =   195
   End
   Begin VB.TextBox txtPipe_ADF2Waste1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   5280
      Width           =   195
   End
   Begin VB.TextBox txtPipe_ADF2Vent1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   2790
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   3360
      Width           =   195
   End
   Begin VB.TextBox txtPipe_2Can2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   3120
      Width           =   195
   End
   Begin VB.Label lblAdfSPdescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LevelSetPoint:"
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
      Left            =   3840
      TabIndex        =   60
      Top             =   5880
      Width           =   1380
   End
   Begin VB.Label lblFstSPdescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LevelSetPoint:"
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
      Left            =   3840
      TabIndex        =   59
      Top             =   2640
      Width           =   1380
   End
   Begin VB.Label lblAdfSP 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "188.8 %"
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
      Left            =   5220
      TabIndex        =   58
      Top             =   5880
      Width           =   780
   End
   Begin VB.Label lblFstSP 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "188.8 %"
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
      Left            =   5220
      TabIndex        =   57
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label lblNotLiveFuel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "THIS STATION DOES NOT SUPPORT LIVE FUEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label lblTarget 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TARGET "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3720
      TabIndex        =   36
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblActual 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " ACTUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4575
      TabIndex        =   35
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblPipe_ADF2Waste 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TO WASTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblPipe_FS2Waste 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TO WASTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblPipe_2Can 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TO CANISTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   3120
      Width           =   1125
   End
   Begin VB.Label lblPipe_FromSupply 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FROM FACILITY SUPPLY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   360
      Width           =   1965
   End
   Begin VB.Label lblPipe_ADF2Vent 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TO VENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2520
      TabIndex        =   62
      Top             =   3120
      Width           =   735
   End
End
Attribute VB_Name = "frmFuelSupply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''' form FuelSupply
' error module 2456
Option Explicit

Private Sub SetCancelToUnenabled()
    cmdCancel.Caption = cmdPause.Caption
    cmdCancel.Picture = cmdPause.Picture
    cmdCancel.ToolTipText = cmdPause.ToolTipText
    cmdCancel.Enabled = False
End Sub

Private Sub SetCancelToPause()
    cmdCancel.Caption = cmdPause.Caption
    cmdCancel.Picture = cmdPause.Picture
    cmdCancel.ToolTipText = cmdPause.ToolTipText
    cmdCancel.Enabled = True
End Sub

Private Sub SetCancelToStop()
    cmdCancel.Caption = cmdStop.Caption
    cmdCancel.Picture = cmdStop.Picture
    cmdCancel.ToolTipText = cmdStop.ToolTipText
    cmdCancel.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    If AdfControl(DispStn).Step = 89 Then
        ' already paused; stop
        AdfControl(DispStn).Mode = 0
        AdfControl(DispStn).Step = 0
        AdfControl(DispStn).ReadyForLoad = False
        SetCancelToUnenabled
    ElseIf AdfControl(DispStn).Step = 61 Then
        ' already paused; stop
        AdfControl(DispStn).Mode = 0
        AdfControl(DispStn).Step = 0
        AdfControl(DispStn).ReadyForLoad = False
        SetCancelToUnenabled
    Else
        ' pause
        AdfControl(DispStn).StepBeforePause = AdfControl(DispStn).Step
        AdfControl(DispStn).Step = 89
        SetCancelToStop
    End If
End Sub

Private Sub cmdDone_Click()
    AdfControl(DispStn).Message = "Manual Refill Complete"
    AdfControl(DispStn).InitialFill_Complete = True
    StationControl(DispStn, 1).LiveFuelCycleCount = 0
    AdfControl(DispStn).ReadyForRefill = False
    AdfControl(DispStn).ButtonVisible_Done = False
    SetCancelToUnenabled
    If (StationRecipe(DispStn, DispShift).ADF_Heater And STN_INFO(DispStn).ADF_TANKTYPE <> 90) Then
        AdfControl(DispStn).Mode = 3
        AdfControl(DispStn).Step = 31
        AdfControl(DispStn).ReadyForLoad = False
    Else
        If LoadControl(DispStn, DispShift).WaterBathTempOK Then
            AdfControl(DispStn).ReadyForLoad = True
            ' start Load
'MsgBox "Load 22", vbInformation, "Info"
            Load_Start CInt(DispStn), CInt(1)
        End If
    End If
End Sub

Private Sub cmdDrain_Click()
    ' Start Vapor Generator Tank Drain Sequence
    AdfControl(DispStn).Mode = 1
    AdfControl(DispStn).Step = 0
    SetCancelToPause
End Sub

Private Sub cmdDrain2_Click()
    ' Start Fuel Storage Tank Drain Sequence
    FstControl(DispStn).Mode = 1
    FstControl(DispStn).Step = 0
End Sub

Private Sub cmdFill2_Click()
    ' Start Fuel Storage Tank Fill Sequence
    FstControl(DispStn).Mode = 2
    FstControl(DispStn).Step = 0
End Sub

Private Sub cmdIgnore_Click()
    ' Ignore LowLevel Switch Still Being On
    If AdfControl(DispStn).Heater Then
      AdfControl(DispStn).Heater_Enable = True
      AdfControl(DispStn).Step = 31   ' Continue; Heater in Use
    Else
      AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, 5)         ' delay for level to settle
      AdfControl(DispStn).Step = 49   ' Fill Complete;No Heater; Monitor Level
    End If
    AdfControl(DispStn).ButtonVisible_Ignore = False
    AdfControl(DispStn).ButtonVisible_Retry = False
End Sub

Private Sub cmdRetry_Click()
    Select Case AdfControl(DispStn).Step
        Case 61     ' LowLevel Switch is Still On
            ' Resume Filling
            AdfControl(DispStn).Step = 20
        Case 89     ' User Pause
            ' Reset Step Timeout
            Select Case AdfControl(DispStn).StepBeforePause
                Case 2
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).PurgeTimeout)    ' setup timeout for Purge
                Case 3
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).PurgeDrainDelay)
                Case 12
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).DrainTimeout)   ' setup timeout for Drain
                Case 13
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).DrainDelay)
                Case 16
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).DrainDelay)
                Case 22
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).FillTimeout)    ' setup timeout for Fill
                Case 24
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).FillDelay)
                Case 25
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, StationCfg_ADF(DispStn, 1).HeaterTimeout, 0)
                Case 32
                    If AdfControl(DispStn).AdfDefinition.hasADF_Heater Then
                        AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, 2 * StationCfg_ADF(DispStn, 1).PurgeTimeout)   ' longer purge timeout if Heater Only
                    Else
                        AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).PurgeTimeout)       ' setup timeout for Purge
                    End If
                Case 33
                    If AdfControl(DispStn).AdfDefinition.hasADF_Heater Then
                        AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, 2 * StationCfg_ADF(DispStn, 1).PurgeFillDelay)   ' longer purge timeout if Heater Only
                    Else
                        AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(DispStn, 1).PurgeFillDelay)       ' setup timeout for Purge
                    End If
                Case 39
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, StationCfg_ADF(DispStn, 1).HeaterTimeout, 0)
                Case 49
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, 5)         ' delay for level to settle
                Case Else
                    AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, 0, 5)         ' default = 5 seconds
            End Select
            ' Restart Sequence at last step before Pause
            AdfControl(DispStn).Step = AdfControl(DispStn).StepBeforePause
        Case 91     ' Drain Timeout
            ' Restart Drain Sequence
            AdfControl(DispStn).Step = 0
        Case 92     ' Fill Timeout
            ' Restart Fill Sequence
            AdfControl(DispStn).Step = 20
        Case 93     ' Reach Temp Timeout
            ' Restart Heat to Temp
            AdfControl(DispStn).Step_Time = Now() + TimeSerial(0, StationCfg_ADF(DispStn, 1).HeaterTimeout, 0)
            AdfControl(DispStn).Step = 39
        Case 94     ' N2 PS Timeout
            ' Restart N2 Sequence
            AdfControl(DispStn).Step = 31
        Case 96     ' N2 PS Maintain Timeout
            ' Restart Temp Control Sequence
            AdfControl(DispStn).Step = 101
        Case 97     ' Storage Tank Low Level
            ' Restart Fill Sequence
            AdfControl(DispStn).Step = 20
        Case 98     ' Vapor Tank Low Level
            ' Restart Fill Sequence
            AdfControl(DispStn).Step = 20
        Case Else
            ' Should never get here; Reset Everything
            Write_ELog "Can't Retry from ADF @ Mode" & AdfControl(DispStn).Mode & " Step" & AdfControl(DispStn).Step & " for Station" & DispStn & " Shift" & DispShift
            AdfControl(DispStn).Mode = 0
            AdfControl(DispStn).Step = 0
    End Select
    SetCancelToPause
    AdfControl(DispStn).ButtonVisible_Retry = False
End Sub

Private Sub cmdStop2_Click()
    FstControl(DispStn).Message = "Operator Pressed Stop"
    FstControl(DispStn).Mode = 0
    FstControl(DispStn).Step = 0
    FstControl(DispStn).StepBeforePause = 0
    Stn_OutDigital DispStn, isStorageDrainSol, cOFF
    Stn_OutDigital DispStn, isStorageFillSol, cOFF
    Stn_OutDigital DispStn, isStorageFillRequest, cOFF
    If AdfControl(DispStn).Mode = 0 Then
        If Stn_DIO(DispStn, isFuelPumpMotor).Value Then Stn_OutDigital DispStn, isFuelPumpMotor, cOFF               ' Turn Pump OFF
    End If
    FstControl(DispStn).ButtonVisible_Drain = True
    FstControl(DispStn).ButtonVisible_Fill = True
    FstControl(DispStn).ButtonVisible_Stop = False
End Sub

Sub UpdateScreen()
Dim tmpPercent As Single
Dim tmpEU As Single
Dim flag As Boolean

    ' **************************************************************************************
    '
    '   LIVE FUEL TANK DISPLAY
    SetErrModule 2456, 13
    frmFuelSupply.Caption = "Fuel Supply for Station #" & Format(DispStn, "#0")
    If STN_INFO(DispStn).ADF_TANKTYPE > 0 Then
            
        ' FUEL STORAGE TANK
        If STN_INFO(DispStn).ADF_DEF.hasADF_FST Then
        
            pnlFS_Level.Top = 1020
            pnlFS_Tank.Top = pnlFS_Level.Top
            lblPipe_FromSupply.Top = 360
            txtPipe_FromSupply1.Top = lblPipe_FromSupply.Top
            txtPipe_FromSupply2.Top = lblPipe_FromSupply.Top
            lblPipe_FS2Waste.Left = 120
            txtPipe_FS2Waste1.Left = 1320
            txtPipe_FS2Waste2.Top = lblPipe_FS2Waste.Top
            txtPipe_FS2ADF.Left = txtPipe_FromSupply2.Left
            lblPipe_ADF2Waste.Left = lblPipe_FS2Waste.Left
            txtPipe_ADF2Waste1.Left = txtPipe_FS2Waste1.Left
            txtPipe_ADF2Waste2.Left = txtPipe_FS2Waste2.Left
            lblPipe_2Can.Top = lblPipe_FS2Waste.Top
            txtPipe_2Can1.Top = lblPipe_2Can.Top
            txtPipe_2Can2.Top = lblPipe_2Can.Top
            pnlADF_Level.Top = 3840
            pnlADF_Tank.Top = pnlADF_Level.Top
            lblTarget.Top = 3600
            lblActual.Top = lblTarget.Top
            frmADF_Controls.Left = 240
            lblNotLiveFuel.Top = OutOfSight
            
            ' level
            lblFstSP.Caption = Format(FstControl(DispStn).LevelSP, "###0.0##")
            tmpPercent = Stn_AIO(DispStn, asStorageTankLevel).EUValue
            If tmpPercent < 0 Then tmpPercent = 0
            If tmpPercent > 100 Then tmpPercent = 100
            pnlFS_Level.FloodPercent = tmpPercent
            tmpEU = (tmpPercent / CSng(100)) * StationCfg_ADF(DispStn, 1).FuelStorageTankVol
            Select Case SysSysDef.USINGLVol_SI
                Case True
                    lblFS_Level.Caption = Format(tmpEU, "#,##0.0#") & " l"
                Case False
                    lblFS_Level.Caption = Format(tmpEU, "#,##0.0#") & " gal"
            End Select
            lblFS_Level.ToolTipText = Format(tmpPercent, "##0.0") & " %"
            lblFS_Level.ForeColor = IIf(Stn_DIO(DispStn, isStorageHiHiLevelLS).Value, LTORANGE, IIf(Stn_DIO(DispStn, isStorageLowLevelLS).Value, DKGREEN, LTORANGE))
            ' level switches
            txtFS_HiHi.BackColor = IIf(Stn_DIO(DispStn, isStorageHiHiLevelLS).Value, Warning_ForeColor, MEDGRAY)
            txtFS_Low.BackColor = IIf(Stn_DIO(DispStn, isStorageLowLevelLS).Value, DKGREEN, Warning_ForeColor)
            ' drain/fill messages
            If (Stn_DIO(DispStn, isStorageFillSol).Value And Stn_DIO(DispStn, isStorageFillRequest).Value) Then
                lblPipe_FromSupply.ForeColor = DKGREEN
                txtPipe_FromSupply1.BackColor = pnlFS_Level.FloodColor
                txtPipe_FromSupply2.BackColor = pnlFS_Level.FloodColor
                txtFS_Fill.ForeColor = BarActual_ForeColor
                txtFS_Fill.text = "Filling"
            Else
                lblPipe_FromSupply.ForeColor = MEDGRAY
                txtPipe_FromSupply1.BackColor = MEDGRAY
                txtPipe_FromSupply2.BackColor = MEDGRAY
                txtFS_Fill.ForeColor = MEDGRAY
                txtFS_Fill.text = "FillClosed"
            End If
            If Stn_DIO(DispStn, isFuelFillSol).Value Then
                txtPipe_FS2ADF.BackColor = pnlFS_Level.FloodColor
                txtFS_Drain.ForeColor = BarActual_ForeColor
                txtFS_Drain.text = "Transferring"
            ElseIf (Stn_DIO(DispStn, isStorageDrainSol).Value And Stn_DIO(DispStn, isFuelPumpMotor).Value) Then
                lblPipe_FS2Waste.ForeColor = DKGREEN
                txtPipe_FS2Waste1.BackColor = pnlFS_Level.FloodColor
                txtPipe_FS2Waste2.BackColor = pnlFS_Level.FloodColor
                txtFS_Drain.ForeColor = BarActual_ForeColor
                txtFS_Drain.text = "Draining"
            Else
                lblPipe_FS2Waste.ForeColor = MEDGRAY
                txtPipe_FS2ADF.BackColor = MEDGRAY
                txtPipe_FS2Waste1.BackColor = MEDGRAY
                txtPipe_FS2Waste2.BackColor = MEDGRAY
                txtFS_Drain.ForeColor = MEDGRAY
                txtFS_Drain.text = "DrainClosed"
            End If
            ' **************************************************************************************
            '
            ' Storage Tank Mode & Task Messages
            If (FstControl(DispStn).Mode <> 0) Then
                txtFstTask.Left = 3750
                txtFstMessage.Left = 3750
                txtFstTask.text = FstControl(DispStn).Task
                txtFstMessage.text = FstControl(DispStn).Message
            Else
                txtFstTask.Left = OutOfSight
                txtFstMessage.Left = OutOfSight
            End If
        ElseIf (STN_INFO(DispStn).ADF_TANKTYPE < 90) Then
            ' no fuel storage tank but has ADF
            lblPipe_FromSupply.Top = 360
            txtPipe_FromSupply1.Top = lblPipe_FromSupply.Top
            txtPipe_FromSupply2.Top = lblPipe_FromSupply.Top
            txtPipe_FS2ADF.Left = txtPipe_FromSupply2.Left
            pnlFS_Level.Left = OutOfSight
            pnlFS_Tank.Left = OutOfSight
            txtFstTask.Left = OutOfSight
            txtFstMessage.Left = OutOfSight
            lblPipe_FS2Waste.Left = OutOfSight
            txtPipe_FS2Waste2.Left = OutOfSight
            txtPipe_FS2Waste1.Left = OutOfSight
            lblFstSPdescription.Left = OutOfSight
            lblFstSP.Left = OutOfSight
            lblNotLiveFuel.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
        Else
            ' no fuel storage tank, no ADF
            lblPipe_FromSupply.Top = OutOfSight
            txtPipe_FromSupply1.Top = OutOfSight
            txtPipe_FromSupply2.Top = OutOfSight
            txtPipe_FS2ADF.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Tank.Left = OutOfSight
            txtFstTask.Left = OutOfSight
            txtFstMessage.Left = OutOfSight
            lblPipe_FS2Waste.Left = OutOfSight
            txtPipe_FS2Waste2.Left = OutOfSight
            txtPipe_FS2Waste1.Left = OutOfSight
            lblFstSPdescription.Left = OutOfSight
            lblFstSP.Left = OutOfSight
            lblNotLiveFuel.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
            pnlFS_Level.Left = OutOfSight
        End If
        
        
        ' VAPOR GENERATOR TANK
        lblAdfSP.Caption = Format(AdfControl(DispStn).LevelSP, "###0.0##")
        tmpPercent = Stn_AIO(DispStn, asFuelTankLevel).EUValue
        If tmpPercent < 0 Then tmpPercent = 0
        If tmpPercent > 100 Then tmpPercent = 100
        pnlADF_Level.FloodPercent = tmpPercent
        tmpEU = (tmpPercent / CSng(100)) * StationCfg_ADF(DispStn, 1).VaporGenTankVol
        Select Case SysSysDef.USINGLVol_SI
            Case True
                lblADF_Level.Caption = Format(tmpEU, "#,##0.0#") & " l"
            Case False
                lblADF_Level.Caption = Format(tmpEU, "#,##0.0#") & " gal"
        End Select
        lblADF_Level.ToolTipText = Format(tmpPercent, "##0.0") & " %"
        lblADF_Level.ForeColor = IIf(Stn_DIO(DispStn, isFuelHiHiLevelLS).Value, Alarm_ForeColor, IIf(Stn_DIO(DispStn, isFuelLowLevelLS).Value, DKGREEN, DK2GRAY))
        
        ' Optional Fuel Heater
        If (STN_INFO(DispStn).ADF_DEF.hasADF_Heater) Then
            txtN2PS.BackColor = IIf(Stn_DIO(DispStn, isFuelPressPS).Value, DKGREEN, MEDGRAY)
            txtN2PS.Top = txtADF_Vapor.Top
            txtADF_Safety.Left = txtN2PS.Left
            txtADF_Heater.Left = txtADF_Vapor.Left
            txtADF_Circ.Left = txtADF_Vapor.Left
            frmHeater.Left = txtADF_Vapor.Left
            If USINGC Then
                lbl_ADF_TempUnits.Caption = "deg C"
            Else
                lbl_ADF_TempUnits.Caption = "deg F"
            End If
            ADF_Temp.Visible = True
            ADF_Sheath.Top = lblADF_Level.Top
            txtADF_Heater.ForeColor = IIf(Stn_DIO(DispStn, isFuelHeaterSSR).Value, MEDRED, MEDGRAY)
            txtADF_Heater.text = IIf(Stn_DIO(DispStn, isFuelHeaterSSR).Value, "Heater On", "HeaterOff")
            If Stn_DIO(DispStn, isFuelSafetyLevelLS).Value Then
                txtADF_Safety.BackColor = IIf(Stn_DIO(DispStn, isFuelOverTempSw).Value, MEDRED, DKGREEN)
            Else
                txtADF_Safety.BackColor = IIf(Stn_DIO(DispStn, isFuelOverTempSw).Value, MEDRED, MEDGRAY)
            End If
            If Stn_DIO(DispStn, isFuelRecircSol).Value And Stn_DIO(DispStn, isFuelDrainSol).Value _
                    And Stn_DIO(DispStn, isFuelPumpMotor).Value Then
                txtADF_Circ.ForeColor = BarActual_ForeColor
                txtADF_Circ.text = "Circulating"
            Else
                txtADF_Circ.ForeColor = MEDGRAY
                txtADF_Circ.text = "Circulate Off"
            End If
        Else
            txtN2PS.Top = OutOfSight
            txtADF_Safety.Left = OutOfSight
            txtADF_Heater.Left = OutOfSight
            txtADF_Circ.Left = OutOfSight
            frmHeater.Left = OutOfSight
            ADF_Sheath.Top = OutOfSight
            txtN2PS.BackColor = MEDGRAY
            txtADF_Heater.ForeColor = MEDGRAY
            txtADF_Heater.text = "No Heater"
            txtADF_Safety.BackColor = MEDGRAY
            txtADF_Circ.ForeColor = MEDGRAY
            txtADF_Circ.text = "Circulate Off"
        End If
        
        lblActual.ForeColor = DKCYAN
        ADF_MAX.ForeColor = DKGREEN
        ADF_TempSP.ForeColor = DKGREEN
        ADF_ACTUAL.ForeColor = DKCYAN
        ADF_Temp.ForeColor = DKCYAN
        ADF_Sheath.ForeColor = DKCYAN
        ADF_ACTUAL.Caption = StationControl(DispStn, 1).LiveFuelCycleCount
        ADF_MAX.Caption = IIf(AdfControl(DispStn).AdfDefinition.hasAUTODRAINFILL, AdfControl(DispStn).LiveFuelChgFreq, StationRecipe(DispStn, 1).LiveFuelChgFreq)
        ADF_Sheath.Caption = Format(Stn_AIO(DispStn, asFuelHeaterTemp).EUValue, "##0.0")
        ADF_Temp.Caption = Format(Stn_AIO(DispStn, asFuelTankTemp).EUValue, "##0.0")
        ADF_TempSP.Caption = Format(StationRecipe(DispStn, 1).ADF_HeaterSP, "#00.0")
        txtADF_HiHi.BackColor = IIf(Stn_DIO(DispStn, isFuelHiHiLevelLS).Value, MEDRED, MEDGRAY)
        txtADF_Low.BackColor = IIf(Stn_DIO(DispStn, isFuelLowLevelLS).Value, DKGREEN, MEDGRAY)
        If Stn_DIO(DispStn, isFuelVentSol).Value Then
            lblPipe_ADF2Vent.Caption = "Vent Open"
            txtPipe_ADF2Vent1.BackColor = pnlADF_Level.FloodColor
            lblPipe_ADF2Vent.ForeColor = DKGREEN
        Else
            lblPipe_ADF2Vent.Caption = "Vent Closed"
            txtPipe_ADF2Vent1.BackColor = MEDGRAY
            lblPipe_ADF2Vent.ForeColor = MEDGRAY
        End If
        If Stn_DIO(DispStn, isFuelVaporSol).Value Then
            txtADF_Vapor.text = "Vapor Open"
            txtADF_Vapor.ForeColor = BarActual_ForeColor
            txtPipe_2Can1.BackColor = pnlADF_Level.FloodColor
            txtPipe_2Can2.BackColor = pnlADF_Level.FloodColor
            lblPipe_2Can.ForeColor = DKGREEN
        Else
            txtADF_Vapor.text = "VaporClosed"
            txtADF_Vapor.ForeColor = MEDGRAY
            txtPipe_2Can1.BackColor = MEDGRAY
            txtPipe_2Can2.BackColor = MEDGRAY
            lblPipe_2Can.ForeColor = MEDGRAY
        End If
        If Stn_DIO(DispStn, isFuelFillSol).Value Then
            txtADF_Fill.ForeColor = BarActual_ForeColor
            txtPipe_FromSupply1.BackColor = DKGREEN
            txtPipe_FromSupply2.BackColor = DKGREEN
            txtPipe_FS2ADF.BackColor = DKGREEN
            lblPipe_FromSupply.ForeColor = DKGREEN
            txtADF_Fill.text = "Filling"
        Else
            txtADF_Fill.ForeColor = MEDGRAY
            txtPipe_FromSupply1.BackColor = MEDGRAY
            txtPipe_FromSupply2.BackColor = MEDGRAY
            txtPipe_FS2ADF.BackColor = MEDGRAY
            lblPipe_FromSupply.ForeColor = MEDGRAY
            txtADF_Fill.text = "FillClosed"
        End If
        If Stn_DIO(DispStn, isFuelDrainSol).Value And Stn_DIO(DispStn, isFuelPumpMotor).Value And Stn_DIO(DispStn, isFuelRecircSol).Value Then
            lblPipe_ADF2Waste.ForeColor = MEDGRAY
            txtPipe_ADF2Waste1.BackColor = MEDGRAY
            txtPipe_ADF2Waste2.BackColor = MEDGRAY
            txtADF_Drain.ForeColor = MEDGRAY
            txtADF_Drain.text = "Circulating"
        ElseIf Stn_DIO(DispStn, isFuelDrainSol).Value And Stn_DIO(DispStn, isFuelPumpMotor).Value Then
            lblPipe_ADF2Waste.ForeColor = DKGREEN
            txtPipe_ADF2Waste1.BackColor = pnlADF_Level.FloodColor
            txtPipe_ADF2Waste2.BackColor = pnlADF_Level.FloodColor
            txtADF_Drain.ForeColor = BarActual_ForeColor
            txtADF_Drain.text = "Draining"
        Else
            lblPipe_ADF2Waste.ForeColor = MEDGRAY
            txtPipe_ADF2Waste1.BackColor = MEDGRAY
            txtPipe_ADF2Waste2.BackColor = MEDGRAY
            txtADF_Drain.ForeColor = MEDGRAY
            txtADF_Drain.text = "DrainClosed"
        End If
        Select Case AdfControl(DispStn).LiveFuelState
            Case fuelDead
                lblFuelStatus.Caption = "Fuel is Dead"
                lblFuelStatus.ForeColor = Alarm_ForeColor
            Case fuelWeak
                lblFuelStatus.Caption = "Fuel is Weak"
                lblFuelStatus.ForeColor = Warning_ForeColor
            Case Else
                lblFuelStatus.Caption = " "
                lblFuelStatus.ForeColor = Good_ForeColor
        End Select
        
        
        ' **************************************************************************************
        '
        ' Vapor Generator Tank Mode & Task Messages
        txtTask.text = AdfControl(DispStn).Task
        txtMessage.text = AdfControl(DispStn).Message
'        txtTask.Enabled = IIf(AdfControl(DispStn).Enable, True, False)
'        txtMessage.Enabled = IIf(AdfControl(DispStn).Enable, True, False)
        
        ' **************************************************************************************
        '
        ' Make Buttons Enabled, or Not
        If ((StationControl(DispStn, 1).Mode = VBIDLE) And (AdfControl(DispStn).Step = 0)) Then
'          AdfControl(DispStn).ButtonVisible_Stop = False
'        Else
'          AdfControl(DispStn).ButtonVisible_Stop = True
        End If
        
        If AdfControl(DispStn).ButtonVisible_Done Then
            cmdDone.Enabled = True
        Else
            cmdDone.Enabled = False
        End If
        If ((Not AdfControl(DispStn).ButtonVisible_Done) And (AdfControl(DispStn).ButtonVisible_Ignore)) Then
            cmdIgnore.Enabled = True
            cmdIgnore.Left = cmdDone.Left
        Else
            cmdIgnore.Enabled = False
            cmdIgnore.Left = OutOfSight
        End If
        If AdfControl(DispStn).ButtonVisible_Retry Then
            cmdRetry.Enabled = True
        Else
            cmdRetry.Enabled = False
        End If
        If AdfControl(DispStn).ButtonVisible_Stop Then
            If ((AdfControl(DispStn).Mode = 2) And (AdfControl(DispStn).Step = 61)) Then SetCancelToStop
            cmdCancel.Enabled = True
            AdfControl(DispStn).ButtonVisible_Stop = False
        ElseIf ((AdfControl(DispStn).Mode = 0) And (AdfControl(DispStn).Step = 0)) Then
            cmdCancel.Enabled = False
        End If
        ' Storage Tank Stop Button
        If FstControl(DispStn).ButtonVisible_Stop Then
            cmdStop2.Enabled = True
        Else
            cmdStop2.Enabled = False
        End If
        ' Enable Storage Tank Fill Button if room for Fuel in Storage Tank
        If ((Not Stn_DIO(DispStn, isStorageHiHiLevelLS).Value) And (FstControl(DispStn).ButtonVisible_Fill)) Then
            cmdFill2.Enabled = True
        Else
            cmdFill2.Enabled = False
        End If
        ' Enable Storage Tank Drain Button if ADF is Idle & Fuel in Storage Tank
        If (((AdfControl(DispStn).Mode = 0) Or (AdfControl(DispStn).Step = 100)) And (Stn_DIO(DispStn, isStorageLowLevelLS).Value) And (FstControl(DispStn).ButtonVisible_Drain)) Then
            cmdDrain2.Enabled = True
        Else
            cmdDrain2.Enabled = False
        End If
        ' Enable Vapor Drain Button if Idle & Fuel in Tank
        If AdfControl(DispStn).AdfDefinition.hasADF_LT Then
            ' use level transmitter
            If (StationControl(DispStn, DispShift).Mode = VBIDLE And AdfControl(DispStn).Mode = 0 And (Stn_AIO(DispStn, asFuelTankLevel).EUValue >= StationCfg_ADF(DispStn, 1).DrainShutOff) And Not Stn_DIO(DispStn, isFuelDrainSol).Value) Or (Alm_LiveFuelLevel(DispStn, 1)) Then
                cmdDrain.Enabled = True
            Else
                cmdDrain.Enabled = False
            End If
        Else
            ' use level switches
            If (StationControl(DispStn, DispShift).Mode = VBIDLE And AdfControl(DispStn).Mode = 0 And (Stn_DIO(DispStn, isFuelLowLevelLS).Value) And Not Stn_DIO(DispStn, isFuelDrainSol).Value) Or (Alm_LiveFuelLevel(DispStn, 1)) Then
                cmdDrain.Enabled = True
            Else
                cmdDrain.Enabled = False
            End If
        End If
    Else
        ' not a Live Fuel Station
        pnlFS_Level.Top = OutOfSight
        pnlFS_Tank.Top = OutOfSight
        lblPipe_FromSupply.Top = OutOfSight
        txtPipe_FromSupply1.Top = OutOfSight
        txtPipe_FromSupply2.Top = OutOfSight
        lblPipe_FS2Waste.Left = OutOfSight
        txtPipe_FS2Waste1.Left = OutOfSight
        txtPipe_FS2Waste2.Top = OutOfSight
        txtPipe_FS2ADF.Left = OutOfSight
        lblPipe_ADF2Waste.Left = OutOfSight
        txtPipe_ADF2Waste1.Left = OutOfSight
        txtPipe_ADF2Waste2.Left = OutOfSight
        lblPipe_2Can.Top = OutOfSight
        txtPipe_2Can1.Top = OutOfSight
        txtPipe_2Can2.Top = OutOfSight
        pnlADF_Level.Top = OutOfSight
        pnlADF_Tank.Top = OutOfSight
        lblTarget.Top = OutOfSight
        lblActual.Top = OutOfSight
        frmADF_Controls.Left = OutOfSight
        lblNotLiveFuel.Top = 3000
        lblNotLiveFuel.Caption = "STATION #" & Format(DispStn, "#0") & " DOES NOT SUPPORT LIVE FUEL"
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

Private Sub Form_Load()

    KeyPreview = True
    
    ' Set Foreground colors
    lblActual.ForeColor = BarActual_ForeColor
    ADF_Temp.ForeColor = BarActual_ForeColor
    ADF_ACTUAL.ForeColor = BarActual_ForeColor
    txtTask.ForeColor = TitlesLabel_ForeColor
    
    SetCancelToUnenabled
    
    ' debug displays
    txtFstTask.Left = OutOfSight
    txtFstMessage.Left = OutOfSight

End Sub

Private Sub tmrScreen_Timer()
    UpdateScreen
End Sub
