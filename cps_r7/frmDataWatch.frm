VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataWatcher 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Watcher"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDataWatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11835
   ScaleMode       =   0  'User
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   13680
      Top             =   3120
   End
   Begin Threed.SSPanel pbarActual2 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   3300
      _Version        =   65536
      _ExtentX        =   5821
      _ExtentY        =   503
      _StockProps     =   15
      BackColor       =   -2147483633
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
   Begin VB.PictureBox pbxBottom 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   15270
      TabIndex        =   8
      Top             =   10395
      Width           =   15330
      Begin Threed.SSPanel pnlPurgeAir 
         Height          =   405
         Left            =   10110
         TabIndex        =   9
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
            Size            =   9.74
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
         TabIndex        =   10
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
            Size            =   11.99
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
            TabIndex        =   12
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
               Size            =   8.11
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
         TabIndex        =   13
         Top             =   0
         Width           =   6030
         _Version        =   65536
         _ExtentX        =   10636
         _ExtentY        =   714
         _StockProps     =   15
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.74
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   20
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
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1058
      ButtonWidth     =   1058
      ButtonHeight    =   953
      _Version        =   393216
   End
   Begin VB.Label lblRecipe 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  recipe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
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
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2550
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   4
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label lblCycle 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cycle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   3
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   180
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   180
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   14760
      X2              =   14760
      Y1              =   1176.961
      Y2              =   2375.718
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      Index           =   0
      Tag             =   "Line"
      X1              =   120
      X2              =   14760
      Y1              =   2353.923
      Y2              =   2353.923
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   1176.961
      Y2              =   2375.718
   End
   Begin VB.Label lblStn 
      BackStyle       =   0  'Transparent
      Caption         =   "Station 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   0
      Tag             =   "Line"
      X1              =   1800
      X2              =   14760
      Y1              =   1176.961
      Y2              =   1176.961
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   0
      Tag             =   "Line"
      X1              =   120
      X2              =   840
      Y1              =   1176.961
      Y2              =   1176.961
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
      Begin VB.Menu mnuak_client 
         Caption         =   "&AK Client"
      End
      Begin VB.Menu mnuak_server 
         Caption         =   "&AK Server"
      End
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
      Begin VB.Menu beforeAbout 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About CPS release7"
      End
   End
End
Attribute VB_Name = "frmDataWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 33 '''''''''''''''''Form DataWatcher.frm '''''''''''''''''''
Option Explicit

Private StnNum, StnSpace As Integer
Private StnWidth, StnLblWidth, ColWidth As Integer
Private StnHeight, StnLblHeight, RowHeight As Integer
Private StnLblLineDeltaHgt, StnLblLineDeltaWid, StnLblFactor, StnLblMargin As Integer
Private MsgHeight, MsgWidth As Integer
Private MsgLeft, MsgTop As Integer
Private TxtHeight, TxtLeft As Integer
Private TxtRecipeTop, TxtCycleTop, TxtModeTop As Integer
Private TxtRecipeWidth, TxtCycleWidth, TxtModeWidth As Integer
Private TxtProgTop, TxtProgLeft, TxtProgHeight, TxtProgWidth As Integer
Private StnLeft, StnLblLeft, FirstColLeft, LastColLeft As Integer
Private NrStns, NrCols, NrRows As Integer
Private NrSpaces, TotalHeight As Integer
Private NrHdrPerStn, NrDataPerStn As Integer
Private StnTop(1 To MAX_STN), TopRowTop As Integer
Private StnLblLineDelta As Integer
Private WtcMode(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Private WtcColData(1 To MAX_STN, 0 To 9) As ColumnsOfData


Private Sub BuildToolbars()
' Create object variable for the Toolbar.
Dim btnX As MSComctlLib.Button
    
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
    
End Sub

Private Sub lblStn_Click(Index As Integer)
    DispStn = Index + 1
    DispShift = IIf((Stn_ActiveShift(DispStn) > 0), Stn_ActiveShift(DispStn), 1)
    ' Station Detail screen position
    frmStnDetail.Show
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

Private Sub CloseScreen()
    Unload Me
    Set frmDataWatcher = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmDataWatcher = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key codes
End Sub

Private Sub Form_Load()
Dim iStn, iShift, iHeight As Integer

SetErrModule 33, 0
If UseLocalErrorHandler Then On Error GoTo localhandler

KeyPreview = True
frmDataWatcher.Height = frmMainMenu.Height
frmDataWatcher.Width = frmMainMenu.Width

BuildToolbars

' Status Bar Setup
frmDataWatcher.pbxBottom.Top = 9900
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
pnlPurgeAir.Width = frmDataWatcher.Width - pnlPurgeAir.Left - 150
pnlPurgeAir.Top = pnlAlarms.Top
pnlPurgeAir.Height = pnlAlarms.Height

' Status Bar Update
UpdateStatusBars


NrCols = 10

NrStns = LAST_STN
NrSpaces = NrStns
StnTop(1) = 960
StnSpace = 90
RowHeight = 170

' Determine Number of Rows & Vertical Space between Stations
TotalHeight = StnTop(1) + (MAX_STN * StnSpace) + (MAX_STN * (80 + (5 * RowHeight)))
Select Case NrStns
    Case 1
        NrRows = 12
        StnHeight = 80 + (NrRows * RowHeight)
        StnSpace = StnSpace
        StnTop(1) = 1400
    Case 2
        NrRows = 12
        StnHeight = 80 + (NrRows * RowHeight)
        StnSpace = 800
        StnTop(1) = 1400
    Case 3
        NrRows = 12
        StnHeight = 80 + (NrRows * RowHeight)
        StnSpace = 640
        StnTop(1) = 1400
    Case 4
        NrRows = 12
        StnHeight = 80 + (NrRows * RowHeight)
        iHeight = StnTop(1) + (NrSpaces * StnSpace) + (NrStns * (80 + (NrRows * RowHeight)))
        StnSpace = StnSpace + ((TotalHeight - iHeight) / NrSpaces)
    Case Else
        NrRows = CInt((((TotalHeight - StnTop(1) - (NrSpaces * StnSpace)) / NrStns) - 80) / RowHeight)
        StnHeight = 80 + (NrRows * RowHeight)
        iHeight = StnTop(1) + (NrSpaces * StnSpace) + (NrStns * (80 + (NrRows * RowHeight)))
        If ((TotalHeight - iHeight) / NrSpaces) = -20 Then
            StnTop(1) = StnTop(1) - 60
            iHeight = StnTop(1) + (NrSpaces * StnSpace) + (NrStns * (80 + (NrRows * RowHeight)))
        ElseIf ((TotalHeight - iHeight) / NrSpaces) < -20 Then
            StnTop(1) = StnTop(1) - 90
            iHeight = StnTop(1) + (NrSpaces * StnSpace) + (NrStns * (80 + (NrRows * RowHeight)))
        End If
        StnSpace = StnSpace + ((TotalHeight - iHeight) / NrSpaces)
End Select

NrHdrPerStn = NrCols
NrDataPerStn = NrCols * (NrRows - 1)

ColWidth = 1100
FirstColLeft = 3500
LastColLeft = 13400
TopRowTop = 120
StnLeft = 170
StnWidth = 14655
StnLblLeft = 1550
StnLblHeight = 360
StnLblFactor = 88
StnLblMargin = 100  '40
StnLblWidth = 750
StnLblLineDeltaHgt = 90
StnLblLineDeltaWid = 60

MsgLeft = 240
MsgHeight = 270
MsgTop = 180
MsgWidth = 3300

TxtLeft = 1020
TxtHeight = 180
TxtRecipeTop = 340
TxtCycleTop = 500
TxtModeTop = 700
TxtProgTop = 900
TxtProgLeft = 240
TxtRecipeWidth = 3300
TxtCycleWidth = 3300
TxtModeWidth = 3300
TxtProgWidth = 3300
TxtProgHeight = 120


' Build Screen
BuildScreen

' Initial Screen Update
For iStn = 1 To NrStns
    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    WtcMode(iStn, iShift) = StationControl(iStn, iShift).Mode
    ClearHeader CInt(iStn)
    ClearData CInt(iStn)
    Select Case StationControl(iStn, iShift).Mode
        Case VBLEAK
            UpdateHeader CInt(iStn)
            UpdateDataLayout CInt(iStn)
            UpdateDataValues CInt(iStn)
        Case VBLOAD
            UpdateHeader CInt(iStn)
            UpdateDataLayout CInt(iStn)
            UpdateDataValues CInt(iStn)
        Case VBPURGE
            UpdateHeader CInt(iStn)
            UpdateDataLayout CInt(iStn)
            UpdateDataValues CInt(iStn)
    End Select
Next iStn

' update timer
tmrUpdate.Interval = 350
tmrUpdate.Enabled = True


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

Private Sub lblMessage_Click(Index As Integer)
    DispStn = Index + 1
    DispShift = IIf((Stn_ActiveShift(DispStn) > 0), Stn_ActiveShift(DispStn), 1)
    ' Station Detail screen position
    frmStnDetail.Show
End Sub

Private Sub lblMode_Click(Index As Integer)
Dim iStn, iShift As Integer
Dim count As Long
    iStn = Index + 1
    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    Select Case StationControl(iStn, iShift).Mode
        Case VBLOAD, VBPURGE
            ' set ReviewData to desired Station
            frmReview.SetRvwStn (CInt(iStn))
            ' ReviewData screen position
'            frmReview.Left = frmMainMenu.Left
'            frmReview.Top = frmMainMenu.Top
            frmReview.Show
            DoEvents
            ' delay  some
                For count = 0 To PAUSEDELAY
                    count = count + 1
                Next count
            ' begin observing data
            frmReview.tbrReview.Buttons("searchnext").Value = tbrPressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrUnpressed
            ' delay  some
                For count = 0 To PAUSEDELAY
                    count = count + 1
                Next count
            ' goto latest data
            frmReview.tbrReview.Buttons("searchnext").Value = tbrPressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrUnpressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrPressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrUnpressed
        Case Else
            ' Nothing to do
    End Select
End Sub

Private Sub lblRecipe_Click(Index As Integer)
    DispStn = Index + 1
    DispShift = IIf((Stn_ActiveShift(DispStn) > 0), Stn_ActiveShift(DispStn), 1)
    ' Station Detail screen position
    frmStnDetail.Show
End Sub

Private Sub pbarActual2_Click(Index As Integer)
Dim iStn, iShift As Integer
Dim count As Long
    iStn = Index + 1
    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    Select Case StationControl(iStn, iShift).Mode
        Case VBLOAD, VBPURGE
            ' set ReviewData to desired Station
            frmReview.SetRvwStn (CInt(iStn))
            ' ReviewData screen position
            frmReview.Show
            DoEvents
            ' delay  some
                For count = 0 To PAUSEDELAY
                    count = count + 1
                Next count
            ' begin observing data
            frmReview.tbrReview.Buttons("searchnext").Value = tbrPressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrUnpressed
            DoEvents
            ' delay  some
                For count = 0 To PAUSEDELAY
                    count = count + 1
                Next count
            ' goto latest data
            frmReview.tbrReview.Buttons("searchnext").Value = tbrPressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrUnpressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrPressed
            DoEvents
            frmReview.tbrReview.Buttons("searchnext").Value = tbrUnpressed
            DoEvents
        Case Else
            ' Nothing to do
    End Select
End Sub

Private Sub tmrUpdate_Timer()
Dim iStn, iShift, iCycle, iFrm As Integer
Dim fillVal As Single
Dim sMsg As String
Dim temptime As Date
Dim tempSec, pauseSec As Long
Dim tempMin, delayMin As Long

SetErrModule 33, 1
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' Update the Navigate Toolbar buttons
    UpdateNavigateBtns
    ' Status Bar
    UpdateStatusBars
    
    For iStn = 1 To NrStns
        ChgErrModule 33, 100
        iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
        iFrm = iStn - 1                                 ' screen 'frame' index
        If StationControl(iStn, iShift).Mode = VBIDLE _
          Or StationControl(iStn, iShift).Mode = VBIDLEWAITING _
          Or StationControl(iStn, iShift).Mode = VBCOMPLETE Then
            iCycle = StationControl(iStn, iShift).CompletedCycles
        Else
            iCycle = StationControl(iStn, iShift).CurrCycle
        End If
        StnTop(iStn) = StnTop(1) + (iFrm * (StnHeight + StnSpace))
    
        lblRecipe(iFrm).Caption = IIf((STN_INFO(iStn).Type = STN_LEAKTEST_TYPE), "1066.985 LeakTest", Mid(StationRecipe(iStn, iShift).Name, 1, 31))
        ' number of cycles
        If (STN_INFO(iStn).Type = STN_LEAKTEST_TYPE) Then
            lblCycle(iFrm).Caption = ""
        Else
            Select Case StationRecipe(iStn, iShift).EndMethod
                Case ENDCYCLES
                    lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
                Case ENDWEIGHTCHG
                    lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0")
                Case Else
                    lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
            End Select
        End If
        
        ChgErrModule 33, 101
        Select Case StationControl(iStn, iShift).Mode
            Case VBLEAK
                ' Leak Check - add leak check phase description
                sMsg = ModeDescShort(VBLEAK) & " - " & LeakPhaseDesc(LeakCheckControl.Phase)
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).ToolTipText = ""
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case VBLEAKTEST
                ' Leak Test - add leak
                sMsg = ModeDescShort(VBLEAKTEST) & " - " & DeffCalcMsg
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).ToolTipText = ""
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
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
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).ToolTipText = "Click to ReviewData"
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
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
                    End Select
                End If
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).ToolTipText = "Click to ReviewData"
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case VBPOSTLEAK
                ' Post LeakCheck Pause
                sMsg = ModeDescShort(VBPOSTLEAK)
                sMsg = sMsg & " for "
                sMsg = sMsg & Format(StationRecipe(iStn, iShift).PauseLeakTime, "##0.0#")
                sMsg = sMsg & LoadTypeDesc2(LOADBYTIME)
                sMsg = sMsg
                sMsg = sMsg & LoadTypeDesc3(LOADBYTIME)
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case VBPOSTLOAD
                ' Post Load Pause
                sMsg = ModeDescShort(VBPOSTLOAD)
                sMsg = sMsg & " for "
                sMsg = sMsg & Format(StationRecipe(iStn, iShift).PauseLoadTime, "##0.0#")
                sMsg = sMsg & LoadTypeDesc2(LOADBYTIME)
                sMsg = sMsg
                sMsg = sMsg & LoadTypeDesc3(LOADBYTIME)
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case VBPOSTPURGE
                ' Post Purge Pause
                sMsg = ModeDescShort(VBPOSTPURGE)
                sMsg = sMsg & " for "
                sMsg = sMsg & Format(StationRecipe(iStn, iShift).PausePurgeTime, "##0.0#")
                sMsg = sMsg & LoadTypeDesc2(PURGEBYTIME)
                sMsg = sMsg
                sMsg = sMsg & LoadTypeDesc3(PURGEBYTIME)
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case VBSCALEWAIT
                ' Waiting for Scale(s) - add which scale(s)
                sMsg = ModeDescShort(VBSCALEWAIT)
                ' Using Two Scales ?
                If StationRecipe(iStn, iShift).UsePriScale And StationRecipe(iStn, iShift).UseAuxScale Then
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
                    If Scale_In_Use(StationRecipe(iStn, iShift).PriScaleNo) Then sMsg = sMsg & Format(StationRecipe(iStn, iShift).PriScaleNo, "#0")
                ElseIf StationRecipe(iStn, iShift).UseAuxScale Then
                    ' Using Only Aux Scale
                    If Scale_In_Use(StationRecipe(iStn, iShift).AuxScaleNo) Then sMsg = sMsg & Format(StationRecipe(iStn, iShift).AuxScaleNo, "#0")
                End If
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).ToolTipText = ""
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
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
                        sMsg = sMsg & Format(StationRecipe(iStn, iShift).StartDate, "D MMM, YYYY   h:mm")
                End Select
                lblMode(iFrm).Caption = sMsg
                lblMode(iFrm).ToolTipText = ""
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case Else
                ' only update if mode has changed
                lblMode(iFrm).Caption = ModeDescShort(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ToolTipText = ""
                lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
        End Select
        
        ChgErrModule 33, 102
        
        If StationControl(iStn, iShift).DBFile = "" Then
            lblRecipe(iFrm).ForeColor = frmDataWatcher.BackColor
            lblCycle(iFrm).ForeColor = frmDataWatcher.BackColor
            Select Case StationControl(iStn, iShift).Mode
                Case VBCOURSEPAUSE
                    lblMessage(iFrm).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    tempMin = CLng(StationSequence(iStn, iShift).CourseData(StationControl(iStn, iShift).Course).PauseDuration)
                    temptime = StationSequence(iStn, iShift).CourseData(StationControl(iStn, iShift).Course).DtsStart + TimeSerial(0, CInt(tempMin), 0) - Now()
                    tempSec = CLng((3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime))
                    pauseSec = CLng(60# * tempMin)
                    fillVal = 100 * ((pauseSec - tempSec) / pauseSec)
                    If fillVal < 0 Then fillVal = 0
                    If fillVal > 100 Then fillVal = 100
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                Case Else
                    lblMessage(iFrm).ForeColor = frmDataWatcher.BackColor
                    lblMessage(iFrm).Caption = "No Open DB File"
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = 100
                    pbarActual2(iFrm).ToolTipText = ""
            End Select
        Else
            lblRecipe(iFrm).ForeColor = Data_ForeColor
            lblCycle(iFrm).ForeColor = Data_ForeColor
            Select Case StationControl(iStn, iShift).Mode
                Case VBLEAKTEST
                    lblMessage(iFrm).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    If StationControl(iStn, iShift).Target > 0 Then
                        fillVal = 100 * (StationControl(iStn, iShift).Actual / StationControl(iStn, iShift).Target)
                    Else
                        fillVal = 0
                    End If
                    If fillVal < 0 Then fillVal = 0
                    If fillVal > 100 Then fillVal = 100
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                
                Case VBLEAK
                    lblMessage(iFrm).ForeColor = Data_ForeColor
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                
                Case VBPOSTLEAK
                    lblMessage(iFrm).ForeColor = MEDGRAY
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                
                Case VBLOAD
                    lblMessage(iFrm).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
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
                        If StationControl(iStn, iShift).Target > 0 Then
                            fillVal = 100 * (StationControl(iStn, iShift).Actual / StationControl(iStn, iShift).Target)
                        Else
                            fillVal = 0
                        End If
                    End If
                    If fillVal < 0 Then fillVal = 0
                    If fillVal > 100 Then fillVal = 100
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = "Click to ReviewData"
                Case VBPURGE
                    lblMessage(iFrm).ForeColor = ModeBackColor(StationControl(iStn, iShift).Mode)
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
                        If StationControl(iStn, iShift).Target > 0 Then
                            fillVal = 100 * (StationControl(iStn, iShift).Actual / StationControl(iStn, iShift).Target)
                        Else
                            fillVal = 0
                        End If
                    End If
                    If fillVal < 0 Then fillVal = 0
                    If fillVal > 100 Then fillVal = 100
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = "Click to ReviewData"
                
                Case VBPRELOAD
                    lblMessage(iFrm).ForeColor = MEDGRAY
                    temptime = StationControl(iStn, iShift).End_Time - Now()
                    tempSec = (3600 * Hour(temptime)) + (60 * Minute(temptime)) + Second(temptime)
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                
                Case VBPOSTLOAD
                    lblMessage(iFrm).ForeColor = MEDGRAY
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                
                Case VBPOSTPURGE
                    lblMessage(iFrm).ForeColor = MEDGRAY
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                 
                 Case VBPURGEWAIT
                    lblMessage(iFrm).ForeColor = MEDGRAY
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""

               Case VBSTARTWAIT
                    lblMessage(iFrm).ForeColor = MEDGRAY
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
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
                    pbarActual2(iFrm).FloodColor = ModeBackColor(StationControl(iStn, iShift).Mode)
                    pbarActual2(iFrm).FloodPercent = fillVal
                    pbarActual2(iFrm).ToolTipText = ""
                
                Case Else
                    lblMessage(iFrm).ForeColor = MEDGRAY
                    pbarActual2(iFrm).FloodColor = MEDGRAY
                    pbarActual2(iFrm).FloodPercent = 0
                    pbarActual2(iFrm).BackColor = frmDataWatcher.BackColor
            End Select
            sMsg = Space(0)
            If NR_SHIFT > 1 Then sMsg = "Shift " & Format(iShift, "0") & "   "
'            sMsg = sMsg & "DB File "
'            sMsg = sMsg & Mid(StationControl(iStn, iShift).DBFile, (Len(StationControl(iStn, iShift).DBFile) - 10), 7)
            sMsg = sMsg & "Job # "
            sMsg = sMsg & StationControl(iStn, iShift).Job_Number
            lblMessage(iFrm).Caption = sMsg
        End If
        
        ChgErrModule 33, 103
        Select Case StationControl(iStn, iShift).Mode
            Case VBLOAD
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateHeader CInt(iStn)
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateDataLayout CInt(iStn)
                UpdateDataValues CInt(iStn)
            Case VBPURGE
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateHeader CInt(iStn)
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateDataLayout CInt(iStn)
                UpdateDataValues CInt(iStn)
            Case VBLEAK
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateHeader CInt(iStn)
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateDataLayout CInt(iStn)
                UpdateDataValues CInt(iStn)
            Case VBLEAKTEST
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateHeader CInt(iStn)
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then UpdateDataLayout CInt(iStn)
                UpdateDataValues CInt(iStn)
            Case Else
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then ClearHeader CInt(iStn)
                If WtcMode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then ClearData CInt(iStn)
        End Select
        WtcMode(iStn, iShift) = StationControl(iStn, iShift).Mode
        
    Next iStn
    
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

Private Sub BuildScreen()
Dim iStn, iShift, iCycle, iHeight As Integer
Dim iCol, iRow, idx, iFrm, iLeft As Integer
Dim alignCode As Integer
Dim sMsg As String
Dim NrShifts As Integer
Dim txtWidth As Single

SetErrModule 33, 8
If UseLocalErrorHandler Then On Error GoTo localhandler

NrShifts = NR_SHIFT
    
' **************
' FIRST STATION
' **************
iStn = 1                                        ' station
iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
iFrm = iStn - 1                                 ' screen 'frame' index
If StationControl(iStn, iShift).Mode = VBIDLE _
  Or StationControl(iStn, iShift).Mode = VBIDLEWAITING _
  Or StationControl(iStn, iShift).Mode = VBCOMPLETE Then
    iCycle = StationControl(iStn, iShift).CompletedCycles
Else
    iCycle = StationControl(iStn, iShift).CurrCycle
End If

Printer.Font = lblStn(iFrm).Font
txtWidth = Printer.TextWidth(STN_INFO(iStn).desc)
Select Case txtWidth
    Case Is < 1000
        StnLblFactor = 100
    Case Is < 2000
        StnLblFactor = 90
    Case Else
        StnLblFactor = 80
End Select
StnLblWidth = StnLblMargin + CInt(CSng(StnLblFactor) * txtWidth * 0.01)

Line1(iFrm).X1 = StnLeft
Line1(iFrm).X2 = StnLblLeft - StnLblLineDeltaWid
Line1(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
Line1(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt
Line1(iFrm).BorderColor = DataHiLite_ForeColor

Line2(iFrm).X1 = StnLblLeft + StnLblWidth
Line2(iFrm).X2 = StnLeft + StnWidth
Line2(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
Line2(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt
Line2(iFrm).BorderColor = DataHiLite_ForeColor

Line3(iFrm).X1 = StnLeft
Line3(iFrm).X2 = StnLeft
Line3(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
Line3(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
Line3(iFrm).BorderColor = DataHiLite_ForeColor

Line4(iFrm).X1 = StnLeft
Line4(iFrm).X2 = StnLeft + StnWidth
Line4(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
Line4(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
Line4(iFrm).BorderColor = DataHiLite_ForeColor

Line5(iFrm).X1 = StnLeft + StnWidth
Line5(iFrm).X2 = StnLeft + StnWidth
Line5(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
Line5(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
Line5(iFrm).BorderColor = DataHiLite_ForeColor

lblStn(iFrm).Top = StnTop(iStn)
lblStn(iFrm).Left = StnLblLeft
lblStn(iFrm).Width = StnLblWidth
lblStn(iFrm).Caption = Mid(STN_INFO(iStn).desc, 1, 18)
lblStn(iFrm).ToolTipText = "Click for Station Detail"
lblStn(iFrm).Visible = True
lblStn(iFrm).ForeColor = TitlesLabel_ForeColor

lblRecipe(iFrm).Top = StnTop(iStn) + TxtRecipeTop
lblRecipe(iFrm).Left = TxtProgLeft
lblRecipe(iFrm).Width = TxtProgWidth
lblRecipe(iFrm).Alignment = 2             ' center
lblRecipe(iFrm).Caption = Mid(StationRecipe(iStn, iShift).Name, 1, 31)
lblRecipe(iFrm).ToolTipText = "Click for Station Detail"
lblRecipe(iFrm).Visible = True

lblCycle(iFrm).Top = StnTop(iStn) + TxtCycleTop
lblCycle(iFrm).Left = TxtProgLeft
lblCycle(iFrm).Width = TxtProgWidth
lblCycle(iFrm).Alignment = 2             ' center
' number of cycles
Select Case StationRecipe(iStn, iShift).EndMethod
    Case ENDCYCLES
        lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
    Case ENDWEIGHTCHG
        lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0")
    Case Else
        lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
End Select
lblCycle(iFrm).Visible = True

lblMode(iFrm).Top = StnTop(iStn) + TxtModeTop
lblMode(iFrm).Left = TxtProgLeft
lblMode(iFrm).Height = TxtHeight + 30
lblMode(iFrm).Width = TxtProgWidth
lblMode(iFrm).Alignment = 2             ' center
'lblMode(iFrm).FontName = "Arial"
lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
lblMode(iFrm).Caption = ModeDescShort(StationControl(iStn, iShift).Mode)
lblMode(iFrm).Visible = True

pbarActual2(iFrm).Top = StnTop(iStn) + TxtProgTop
pbarActual2(iFrm).Left = TxtProgLeft
pbarActual2(iFrm).Height = TxtProgHeight
pbarActual2(iFrm).Width = TxtProgWidth
pbarActual2(iFrm).FloodPercent = 0
pbarActual2(iFrm).Visible = True

lblMessage(iFrm).Alignment = IIf(NR_SHIFT > 1, 0, 2)       ' 0=left justify, 2=center
lblMessage(iFrm).Top = StnTop(iStn) + MsgTop
lblMessage(iFrm).Left = MsgLeft
lblMessage(iFrm).Height = MsgHeight
lblMessage(iFrm).Width = MsgWidth
lblMessage(iFrm).ToolTipText = "Click for Station Detail"
lblMessage(iFrm).Visible = True

If StationControl(iStn, iShift).DBFile = "" Then
    lblRecipe(iFrm).ForeColor = frmDataWatcher.BackColor
    lblCycle(iFrm).ForeColor = frmDataWatcher.BackColor
    lblMessage(iFrm).ForeColor = frmDataWatcher.BackColor
    lblMessage(iFrm).Caption = "No Open DB File"
Else
    lblCycle(iFrm).ForeColor = Data_ForeColor
    Select Case StationControl(iStn, iShift).Mode
        Case VBLOAD
            lblMessage(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            lblRecipe(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
        Case VBPURGE
            lblMessage(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            lblRecipe(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
        Case Else
            lblMessage(iFrm).ForeColor = Data_ForeColor
            lblRecipe(iFrm).ForeColor = Data_ForeColor
    End Select
    sMsg = Space(10)
    If NR_SHIFT > 1 Then sMsg = "Shift " & Format(iShift, "0") & "   "
    sMsg = sMsg & "DB File "
    sMsg = sMsg & Mid(StationControl(iStn, iShift).DBFile, (Len(StationControl(iStn, iShift).DBFile) - 10), 7)
    lblMessage(iFrm).Caption = sMsg
End If


' header & data rows
For iRow = 0 To (NrRows - 1)
    For iCol = 0 To (NrCols - 1)
        Select Case iRow
            Case 0
                ' Header
                If iCol = 0 Then
                    ' lblHeader(0)
                    lblHeader(iCol).Left = FirstColLeft + (iCol * ColWidth)
                    lblHeader(iCol).Top = StnTop(iStn) + TopRowTop
                    lblHeader(iCol).Caption = "Hdr " & Format(iCol, "0")
                    lblHeader(iCol).Visible = True
                Else
                    Load lblHeader(iCol)
                    lblHeader(iCol).Left = FirstColLeft + (iCol * ColWidth)
                    lblHeader(iCol).Top = StnTop(iStn) + TopRowTop
                    lblHeader(iCol).Caption = "Hdr " & Format(iCol, "0")
                    lblHeader(iCol).Visible = True
                End If
            Case 1
                ' Data - 1st row
                If iCol = 0 Then
                    ' lblData(0)
                    idx = iCol + (NrCols * (iRow - 1))
                    lblData(idx).Left = FirstColLeft + (iCol * ColWidth)
                    lblData(idx).Top = StnTop(iStn) + TopRowTop + (iRow * RowHeight)
                    lblData(idx).Caption = "Data " & Format(idx, "0000")
                    lblData(idx).Visible = True
                    lblData(idx).ForeColor = Data_ForeColor
                Else
                    idx = iCol + (NrCols * (iRow - 1))
                    Load lblData(idx)
                    lblData(idx).Left = FirstColLeft + (iCol * ColWidth)
                    lblData(idx).Top = StnTop(iStn) + TopRowTop + (iRow * RowHeight)
                    lblData(idx).Caption = "Data " & Format(idx, "0000")
                    lblData(idx).Visible = True
                    lblData(idx).ForeColor = Data_ForeColor
                End If
            Case Else
                ' Data - other rows
                idx = iCol + (NrCols * (iRow - 1))
                Load lblData(idx)
                lblData(idx).Left = FirstColLeft + (iCol * ColWidth)
                lblData(idx).Top = StnTop(iStn) + TopRowTop + (iRow * RowHeight)
                lblData(idx).Caption = "Data " & Format(idx, "0000")
                lblData(idx).Visible = True
                lblData(idx).ForeColor = Data_ForeColor
        End Select
    Next iCol
Next iRow


' **************
' OTHER STATIONS
' **************
For iStn = 2 To NrStns

    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    iFrm = iStn - 1                                 ' screen 'frame' index
    If StationControl(iStn, iShift).Mode = VBIDLE _
      Or StationControl(iStn, iShift).Mode = VBIDLEWAITING _
      Or StationControl(iStn, iShift).Mode = VBCOMPLETE Then
        iCycle = StationControl(iStn, iShift).CompletedCycles
    Else
        iCycle = StationControl(iStn, iShift).CurrCycle
    End If
    StnTop(iStn) = StnTop(1) + (iFrm * (StnHeight + StnSpace))
    txtWidth = Printer.TextWidth(STN_INFO(iStn).desc)
    Select Case txtWidth
        Case Is < 1000
            StnLblFactor = 100
        Case Is < 2000
            StnLblFactor = 90
        Case Else
            StnLblFactor = 80
    End Select
    StnLblWidth = StnLblMargin + CInt(CSng(StnLblFactor) * txtWidth * 0.01)
    
    Load Line1(iFrm)
    Line1(iFrm).X1 = StnLeft
    Line1(iFrm).X2 = StnLblLeft - StnLblLineDeltaWid
    Line1(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
    Line1(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt
    Line1(iFrm).Visible = True
    Line1(iFrm).BorderColor = DataHiLite_ForeColor
    
    Load Line2(iFrm)
    Line2(iFrm).X1 = StnLblLeft + StnLblWidth
    Line2(iFrm).X2 = StnLeft + StnWidth
    Line2(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
    Line2(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt
    Line2(iFrm).Visible = True
    Line2(iFrm).BorderColor = DataHiLite_ForeColor
    
    Load Line3(iFrm)
    Line3(iFrm).X1 = StnLeft
    Line3(iFrm).X2 = StnLeft
    Line3(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
    Line3(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
    Line3(iFrm).Visible = True
    Line3(iFrm).BorderColor = DataHiLite_ForeColor
    
    Load Line4(iFrm)
    Line4(iFrm).X1 = StnLeft
    Line4(iFrm).X2 = StnLeft + StnWidth
    Line4(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
    Line4(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
    Line4(iFrm).Visible = True
    Line4(iFrm).BorderColor = DataHiLite_ForeColor
    
    Load Line5(iFrm)
    Line5(iFrm).X1 = StnLeft + StnWidth
    Line5(iFrm).X2 = StnLeft + StnWidth
    Line5(iFrm).Y1 = StnTop(iStn) + StnLblLineDeltaHgt
    Line5(iFrm).Y2 = StnTop(iStn) + StnLblLineDeltaHgt + StnHeight
    Line5(iFrm).Visible = True
    Line5(iFrm).BorderColor = DataHiLite_ForeColor
    
    Load lblStn(iFrm)
    lblStn(iFrm).Top = StnTop(iStn)
    lblStn(iFrm).Left = StnLblLeft
    lblStn(iFrm).Width = StnLblWidth
    lblStn(iFrm).Alignment = 0         ' left justify
'    lblStn(iFrm).Caption = "Station " & Format(iStn, "0")
    lblStn(iFrm).Caption = Mid(STN_INFO(iStn).desc, 1, 18)
    lblStn(iFrm).ToolTipText = "Click for Station Detail"
    lblStn(iFrm).Visible = True
    lblStn(iFrm).ForeColor = TitlesLabel_ForeColor
    
    Load lblRecipe(iFrm)
    lblRecipe(iFrm).Top = StnTop(iStn) + TxtRecipeTop
    lblRecipe(iFrm).Left = TxtProgLeft
    lblRecipe(iFrm).Width = TxtProgWidth
    lblRecipe(iFrm).Alignment = 2             ' center
    lblRecipe(iFrm).Caption = Mid(StationRecipe(iStn, iShift).Name, 1, 31)
    lblRecipe(iFrm).ToolTipText = "Click for Station Detail"
    lblRecipe(iFrm).Visible = True

    Load lblCycle(iFrm)
    lblCycle(iFrm).Top = StnTop(iStn) + TxtCycleTop
    lblCycle(iFrm).Left = TxtProgLeft
    lblCycle(iFrm).Width = TxtProgWidth
    lblCycle(iFrm).Alignment = 2             ' center
    ' number of cycles
    Select Case StationRecipe(iStn, iShift).EndMethod
        Case ENDCYCLES
            lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
        Case ENDWEIGHTCHG
            lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0")
        Case Else
            lblCycle(iFrm).Caption = "Cycle " & Format(iCycle, "0") & " of " & Format(StationRecipe(iStn, iShift).Cycles, "0")
    End Select
    lblCycle(iFrm).Visible = True
    
    Load lblMode(iFrm)
    lblMode(iFrm).Top = StnTop(iStn) + TxtModeTop
    lblMode(iFrm).Left = TxtProgLeft
    lblMode(iFrm).Height = TxtHeight + 30
    lblMode(iFrm).Width = TxtProgWidth
    lblMode(iFrm).Alignment = 2             ' center
'    lblMode(iFrm).FontName = "Arial"
    lblMode(iFrm).BackColor = ModeBackColor(StationControl(iStn, iShift).Mode)
    lblMode(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
    lblMode(iFrm).Caption = ModeDescShort(StationControl(iStn, iShift).Mode)
    lblMode(iFrm).Visible = True
    
    Load pbarActual2(iFrm)
    pbarActual2(iFrm).Top = StnTop(iStn) + TxtProgTop
    pbarActual2(iFrm).Left = TxtProgLeft
    pbarActual2(iFrm).Height = TxtProgHeight
    pbarActual2(iFrm).Width = TxtProgWidth
    pbarActual2(iFrm).FloodPercent = 0
    pbarActual2(iFrm).Visible = True

    Load lblMessage(iFrm)
    lblMessage(iFrm).Alignment = IIf(NR_SHIFT > 1, 0, 2)       ' 0=left justify, 2=center
    lblMessage(iFrm).Top = StnTop(iStn) + MsgTop
    lblMessage(iFrm).Left = MsgLeft
    lblMessage(iFrm).Height = MsgHeight
    lblMessage(iFrm).Width = MsgWidth
    lblMessage(iFrm).ToolTipText = "Click for Station Detail"
    lblMessage(iFrm).Visible = True
    
    If StationControl(iStn, iShift).DBFile = "" Then
        lblRecipe(iFrm).ForeColor = frmDataWatcher.BackColor
        lblCycle(iFrm).ForeColor = frmDataWatcher.BackColor
        lblMessage(iFrm).ForeColor = frmDataWatcher.BackColor
        lblMessage(iFrm).Caption = "No Open DB File"
    Else
        lblCycle(iFrm).ForeColor = Data_ForeColor
        Select Case StationControl(iStn, iShift).Mode
            Case VBLOAD
                lblMessage(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                lblRecipe(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case VBPURGE
                lblMessage(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                lblRecipe(iFrm).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
            Case Else
                lblMessage(iFrm).ForeColor = Data_ForeColor
                lblRecipe(iFrm).ForeColor = Data_ForeColor
        End Select
        sMsg = Space(10)
        If NR_SHIFT > 1 Then sMsg = "Shift " & Format(iShift, "0") & "   "
        sMsg = sMsg & "DB File "
        sMsg = sMsg & Mid(StationControl(iStn, iShift).DBFile, (Len(StationControl(iStn, iShift).DBFile) - 10), 7)
        lblMessage(iFrm).Caption = sMsg
    End If


    ' header & data rows
    For iRow = 0 To (NrRows - 1)
        For iCol = 0 To (NrCols - 1)
            Select Case iRow
                Case 0
                    ' Header
                    idx = iCol + (NrHdrPerStn * iFrm)
                    Load lblHeader(idx)
                    lblHeader(idx).Left = FirstColLeft + (iCol * ColWidth)
                    lblHeader(idx).Top = StnTop(iStn) + TopRowTop
                    lblHeader(idx).Caption = "Hdr " & Format(iCol, "0")
                    lblHeader(idx).Visible = True
                    lblHeader(idx).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)

                Case 1
                    ' Data - 1st row
                    idx = iCol + (NrCols * (iRow - 1)) + (NrDataPerStn * iFrm)
                    Load lblData(idx)
                    lblData(idx).Left = FirstColLeft + (iCol * ColWidth)
                    lblData(idx).Top = StnTop(iStn) + TopRowTop + (iRow * RowHeight)
                    lblData(idx).Caption = "Data " & Format(idx, "0000")
                    lblData(idx).Visible = True
                    lblData(idx).ForeColor = DataHiLite_ForeColor
                
                Case Else
                    ' Data - all other rows
                    idx = iCol + (NrCols * (iRow - 1)) + (NrDataPerStn * iFrm)
                    Load lblData(idx)
                    lblData(idx).Left = FirstColLeft + (iCol * ColWidth)
                    lblData(idx).Top = StnTop(iStn) + TopRowTop + (iRow * RowHeight)
                    lblData(idx).Caption = "Data " & Format(idx, "0000")
                    lblData(idx).Visible = True
                    lblData(idx).ForeColor = Data_ForeColor
            End Select
        Next iCol
    Next iRow

Next iStn


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

Private Sub ClearData(iStn As Integer)
Dim iRow, iCol, idx As Integer
SetErrModule 33, 2
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' Clear Data Rows
    For iRow = 1 To (NrRows - 1)
        For iCol = 0 To (NrCols - 1)
            idx = iCol + (NrDataPerStn * (iStn - 1)) + (NrCols * (iRow - 1))
            lblData(idx).Caption = " "
        Next iCol
    Next iRow

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

Private Sub ClearHeader(iStn As Integer)
Dim iCol, idx As Integer
SetErrModule 33, 3
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' Clear Header Row
    For iCol = 0 To (NrCols - 1)
        idx = iCol + (NrHdrPerStn * (iStn - 1))
        lblHeader(idx).Caption = " "
    Next iCol

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

Private Sub SetupColData(iStn As Integer)
Dim iCol As Integer
Dim iShift As Integer
SetErrModule 33, 4
If UseLocalErrorHandler Then On Error GoTo localhandler

    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    
    iCol = 0
    ' set column 0
    WtcColData(iStn, iCol).ColAlign = "CENTER"
    WtcColData(iStn, iCol).ColWidth = 1100
    WtcColData(iStn, iCol).DataFormat = "HH:MM:SS"
    WtcColData(iStn, iCol).DataName = "Next"
    WtcColData(iStn, iCol).Header1 = "Next"
    WtcColData(iStn, iCol).Header2 = " "
    WtcColData(iStn, iCol).Header3 = " "
    WtcColData(iStn, iCol).InUse = True
    
    Select Case StationControl(iStn, iShift).Mode
        Case VBLOAD
            ' LOAD DATA COLUMNS
            ' Time
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "HH:MM:SS"
            WtcColData(iStn, iCol).DataName = "Time"
            WtcColData(iStn, iCol).Header1 = "Time"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = " "
            WtcColData(iStn, iCol).InUse = True
            
            Select Case STN_INFO(iStn).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE, STN_ORVR2_TYPE
                    ' Nitrogen Flow
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1100
                    WtcColData(iStn, iCol).DataFormat = "#0.000"
                    WtcColData(iStn, iCol).DataName = "NitFlow"
                    WtcColData(iStn, iCol).Header1 = "Nit Flow"
                    WtcColData(iStn, iCol).Header2 = " "
                    WtcColData(iStn, iCol).Header3 = "(slpm)"
                    WtcColData(iStn, iCol).InUse = True
                    ' Mix
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1100
                    WtcColData(iStn, iCol).DataFormat = "##0.0"
                    WtcColData(iStn, iCol).DataName = "Mix"
                    WtcColData(iStn, iCol).Header1 = "Mix %"
                    WtcColData(iStn, iCol).Header2 = " "
                    WtcColData(iStn, iCol).Header3 = "(% Btn)"
                    WtcColData(iStn, iCol).InUse = True
                    ' Butane Flow
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1100
                    WtcColData(iStn, iCol).DataFormat = "#0.000"
                    WtcColData(iStn, iCol).DataName = "BtnFlow"
                    WtcColData(iStn, iCol).Header1 = "Btn Flow"
                    WtcColData(iStn, iCol).Header2 = " "
                    WtcColData(iStn, iCol).Header3 = "(slpm)"
                    WtcColData(iStn, iCol).InUse = True
                    ' Load Rate in Grams
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1100
                    WtcColData(iStn, iCol).DataFormat = "#0.00"
                    WtcColData(iStn, iCol).DataName = "LoadRate"
                    WtcColData(iStn, iCol).Header1 = "Load Rate"
                    WtcColData(iStn, iCol).Header2 = " "
                    WtcColData(iStn, iCol).Header3 = "(gm/hr)"
                    WtcColData(iStn, iCol).InUse = True
                    ' Butane Flow Totalized
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1100
                    WtcColData(iStn, iCol).DataFormat = "###0.00"
                    WtcColData(iStn, iCol).DataName = "LoadGrams"
                    WtcColData(iStn, iCol).Header1 = "Total"
                    WtcColData(iStn, iCol).Header2 = " "
                    WtcColData(iStn, iCol).Header3 = "(grams)"
                    WtcColData(iStn, iCol).InUse = True
                
                Case STN_LIVEFUEL_TYPE
                    ' Vapor Carrier Flow
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1800
                    WtcColData(iStn, iCol).DataFormat = "#0.000"
                    WtcColData(iStn, iCol).DataName = "NitFlow"
                    WtcColData(iStn, iCol).Header1 = "Vapor Carrier Flow"
                    WtcColData(iStn, iCol).Header2 = " "
                    WtcColData(iStn, iCol).Header3 = "(slpm)"
                    WtcColData(iStn, iCol).InUse = True
                    If StationRecipe(iStn, iShift).UsePriScale Then
                        ' load rate
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1100
                        WtcColData(iStn, iCol).DataFormat = "##0.00"
                        WtcColData(iStn, iCol).DataName = "LoadRate"
                        WtcColData(iStn, iCol).Header1 = "LoadRate"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "gm/hr"
                        WtcColData(iStn, iCol).InUse = True
                    End If
                    
                Case STN_LIVEREG_TYPE, STN_LIVEORVR2_TYPE
                    If StationRecipe(iStn, iShift).LiveFuel Then
                        ' Vapor Carrier Flow
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1800
                        WtcColData(iStn, iCol).DataFormat = "#0.000"
                        WtcColData(iStn, iCol).DataName = "NitFlow"
                        WtcColData(iStn, iCol).Header1 = "Vapor Carrier Flow"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "(slpm)"
                        WtcColData(iStn, iCol).InUse = True
                        If StationRecipe(iStn, iShift).UsePriScale Then
                            ' load rate
                            iCol = iCol + 1
                            WtcColData(iStn, iCol).ColAlign = "CENTER"
                            WtcColData(iStn, iCol).ColWidth = 1100
                            WtcColData(iStn, iCol).DataFormat = "##0.00"
                            WtcColData(iStn, iCol).DataName = "LoadRate"
                            WtcColData(iStn, iCol).Header1 = "LoadRate"
                            WtcColData(iStn, iCol).Header2 = " "
                            WtcColData(iStn, iCol).Header3 = "gm/hr"
                            WtcColData(iStn, iCol).InUse = True
                        End If
                    Else
                        ' Nitrogen Flow
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1100
                        WtcColData(iStn, iCol).DataFormat = "#0.000"
                        WtcColData(iStn, iCol).DataName = "NitFlow"
                        WtcColData(iStn, iCol).Header1 = "Nit Flow"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "(slpm)"
                        WtcColData(iStn, iCol).InUse = True
                        ' Mix
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1100
                        WtcColData(iStn, iCol).DataFormat = "##0.0"
                        WtcColData(iStn, iCol).DataName = "Mix"
                        WtcColData(iStn, iCol).Header1 = "Mix %"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "(% Btn)"
                        WtcColData(iStn, iCol).InUse = True
                        ' Butane Flow
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1100
                        WtcColData(iStn, iCol).DataFormat = "#0.000"
                        WtcColData(iStn, iCol).DataName = "BtnFlow"
                        WtcColData(iStn, iCol).Header1 = "Btn Flow"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "(slpm)"
                        WtcColData(iStn, iCol).InUse = True
                        ' Load Rate in Grams
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1100
                        WtcColData(iStn, iCol).DataFormat = "#0.00"
                        WtcColData(iStn, iCol).DataName = "LoadRate"
                        WtcColData(iStn, iCol).Header1 = "Load Rate"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "(gm/hr)"
                        WtcColData(iStn, iCol).InUse = True
                        ' Butane Flow Totalized
                        iCol = iCol + 1
                        WtcColData(iStn, iCol).ColAlign = "CENTER"
                        WtcColData(iStn, iCol).ColWidth = 1100
                        WtcColData(iStn, iCol).DataFormat = "###0.00"
                        WtcColData(iStn, iCol).DataName = "LoadGrams"
                        WtcColData(iStn, iCol).Header1 = "Total"
                        WtcColData(iStn, iCol).Header2 = " "
                        WtcColData(iStn, iCol).Header3 = "(gm Btn)"
                        WtcColData(iStn, iCol).InUse = True
                    End If
                    
                Case STN_COMBO3_TYPE
                    ' future
                    
            End Select
            
            If StationRecipe(iStn, iShift).UsePriScale Then
                ' using a Primary Scale
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "####0.00"
                WtcColData(iStn, iCol).DataName = "PriScale"
                WtcColData(iStn, iCol).Header1 = "Pri"
                WtcColData(iStn, iCol).Header2 = "Scale"
                WtcColData(iStn, iCol).Header3 = "(grams)"
                WtcColData(iStn, iCol).InUse = True
            End If
            
            If StationRecipe(iStn, iShift).UseAuxScale Then
                ' using an Aux Scale
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "####0.00"
                WtcColData(iStn, iCol).DataName = "AuxScale"
                WtcColData(iStn, iCol).Header1 = "Aux"
                WtcColData(iStn, iCol).Header2 = "Scale"
                WtcColData(iStn, iCol).Header3 = "(grams)"
                WtcColData(iStn, iCol).InUse = True
            End If
            
            If StationRecipe(iStn, iShift).UsePriScale Then
                ' weight change rate
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "##0.00"
                WtcColData(iStn, iCol).DataName = "WtChgRate"
                WtcColData(iStn, iCol).Header1 = "WtChgRate"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(gm/hr)"
                WtcColData(iStn, iCol).InUse = True
            End If
            
            If (STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE Or ((STN_INFO(iStn).Type = STN_LIVEREG_TYPE) And StationRecipe(iStn, iShift).LiveFuel) Or ((STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE) And StationRecipe(iStn, iShift).LiveFuel)) Then
            
                '   Cycles
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "#0"
                WtcColData(iStn, iCol).DataName = "LiveFuelCycles"
                WtcColData(iStn, iCol).Header1 = "Cycles"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(SinceRefill)"
                WtcColData(iStn, iCol).InUse = True
                
                If ((STN_INFO(iStn).ADF_TANKTYPE > 10) And (STN_INFO(iStn).ADF_TANKTYPE <= 20)) Then
                    ' Live Fuel Station - AutoDrainFill & Heater
                    '   Fuel Temp
                    iCol = iCol + 1
                    WtcColData(iStn, iCol).ColAlign = "CENTER"
                    WtcColData(iStn, iCol).ColWidth = 1100
                    WtcColData(iStn, iCol).DataFormat = "##0.0"
                    WtcColData(iStn, iCol).DataName = "LiveFuelTemp"
                    WtcColData(iStn, iCol).Header1 = "Fuel Temp"
                    WtcColData(iStn, iCol).Header2 = " "
                    If USINGC Then WtcColData(iStn, iCol).Header3 = "(deg C)"
                    If USINGF Then WtcColData(iStn, iCol).Header3 = "(deg F)"
                    WtcColData(iStn, iCol).InUse = True
                End If
                
            End If
           
           
        Case VBPURGE
            ' PURGE DATA COLUMNS
            '   Time
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "HH:MM:SS"
            WtcColData(iStn, iCol).DataName = "Time"
            WtcColData(iStn, iCol).Header1 = "Time"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = " "
            WtcColData(iStn, iCol).InUse = True
            '   PurgeAir Flow
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "###0.00"
            WtcColData(iStn, iCol).DataName = "PurgeFlow"
            WtcColData(iStn, iCol).Header1 = "PA Flow"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = "(slpm)"
            WtcColData(iStn, iCol).InUse = True
            '   PurgeAir Temp
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "##0.0"
            WtcColData(iStn, iCol).DataName = "PurgeTemp"
            WtcColData(iStn, iCol).Header1 = "PA Temp"
            WtcColData(iStn, iCol).Header2 = " "
            If USINGC Then WtcColData(iStn, iCol).Header3 = "(deg C)"
            If USINGF Then WtcColData(iStn, iCol).Header3 = "(deg F)"
            WtcColData(iStn, iCol).InUse = True
            '   PurgeAir Humidity
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "##0.0"
            WtcColData(iStn, iCol).DataName = "PurgeMoist"
            WtcColData(iStn, iCol).Header1 = "PA Moist"
            WtcColData(iStn, iCol).Header2 = " "
            If USINGMoist_RH Then WtcColData(iStn, iCol).Header3 = "(% rH)"
            If USINGMoist_Grains Then WtcColData(iStn, iCol).Header3 = "(Grn/Lb)"
            WtcColData(iStn, iCol).InUse = True
            
            If StationRecipe(iStn, iShift).UsePriScale Then
                ' using a Primary Scale
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "####0.00"
                WtcColData(iStn, iCol).DataName = "PriScale"
                WtcColData(iStn, iCol).Header1 = "Pri"
                WtcColData(iStn, iCol).Header2 = "Scale"
                WtcColData(iStn, iCol).Header3 = "(grams)"
                WtcColData(iStn, iCol).InUse = True
            End If
            
            If StationRecipe(iStn, iShift).UseAuxScale Then
                ' using an Aux Scale
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "####0.00"
                WtcColData(iStn, iCol).DataName = "AuxScale"
                WtcColData(iStn, iCol).Header1 = "Aux"
                WtcColData(iStn, iCol).Header2 = "Scale"
                WtcColData(iStn, iCol).Header3 = "(grams)"
                WtcColData(iStn, iCol).InUse = True
            End If
            
            If StationRecipe(iStn, iShift).UsePriScale Then
                ' weight change
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1100
                WtcColData(iStn, iCol).DataFormat = "#,##0.00"
                WtcColData(iStn, iCol).DataName = "WtChgRate"
                WtcColData(iStn, iCol).Header1 = "WtChgRate"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(gm/hr)"
                WtcColData(iStn, iCol).InUse = True
            End If
            
            '   Purge Volume
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "####0.0"
            WtcColData(iStn, iCol).DataName = "PurgeVol"
            WtcColData(iStn, iCol).Header1 = "Total Volume"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = "(liters)"
            WtcColData(iStn, iCol).InUse = True
    
        Case VBLEAK
            ' LEAK DATA COLUMNS
            '   Time
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "HH:MM:SS"
            WtcColData(iStn, iCol).DataName = "Time"
            WtcColData(iStn, iCol).Header1 = "Time"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = " "
            WtcColData(iStn, iCol).InUse = True
            '   Pressure
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "####0.00"
            WtcColData(iStn, iCol).DataName = "Pressure"
            WtcColData(iStn, iCol).Header1 = "Pressure"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = "(psig)"
            WtcColData(iStn, iCol).InUse = True
            
        Case VBLEAKTEST
            ' Time
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "HH:MM:SS"
            WtcColData(iStn, iCol).DataName = "Time"
            WtcColData(iStn, iCol).Header1 = "Time"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = " "
            WtcColData(iStn, iCol).InUse = True
            ' Nitrogen Flow
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "#0.000"
            WtcColData(iStn, iCol).DataName = "N2Flow"
            WtcColData(iStn, iCol).Header1 = "N2 Flow"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = "(slpm)"
            WtcColData(iStn, iCol).InUse = True
            ' N2 Temp
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "##0.0"
            WtcColData(iStn, iCol).DataName = "N2Temp"
            WtcColData(iStn, iCol).Header1 = "N2 Temp"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = IIf(USINGC, "deg C", "deg F")
            WtcColData(iStn, iCol).InUse = True
            ' Inlet Pressure
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "##0.0##"
            WtcColData(iStn, iCol).DataName = "InPress"
            WtcColData(iStn, iCol).Header1 = "In Press"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = "(psi)"
            WtcColData(iStn, iCol).InUse = True
            ' Effective Leak Diameter
            iCol = iCol + 1
            WtcColData(iStn, iCol).ColAlign = "CENTER"
            WtcColData(iStn, iCol).ColWidth = 1100
            WtcColData(iStn, iCol).DataFormat = "##0.00##"
            WtcColData(iStn, iCol).DataName = "LeakDia"
            WtcColData(iStn, iCol).Header1 = "Leak Dia"
            WtcColData(iStn, iCol).Header2 = " "
            WtcColData(iStn, iCol).Header3 = "(inches)"
            WtcColData(iStn, iCol).InUse = True
                
    End Select

    ' set remaining columns to Unused
    Do While iCol < (NrCols - 2)
        ' blank
        iCol = iCol + 1
        WtcColData(iStn, iCol).ColAlign = "CENTER"
        WtcColData(iStn, iCol).ColWidth = 100
        WtcColData(iStn, iCol).DataFormat = "0"
        WtcColData(iStn, iCol).DataName = "unused"
        WtcColData(iStn, iCol).Header1 = " "
        WtcColData(iStn, iCol).Header2 = " "
        WtcColData(iStn, iCol).Header3 = " "
        WtcColData(iStn, iCol).InUse = False
    Loop

    ' last column
    If iCol <= (NrCols - 2) Then
        Select Case StationControl(iStn, iShift).Mode
            Case VBLOAD
                '   Test Time
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1400
                WtcColData(iStn, iCol).DataFormat = "#,###,##0.000"
                WtcColData(iStn, iCol).DataName = "TestTime"
                WtcColData(iStn, iCol).Header1 = "Test Time"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(seconds)"
                WtcColData(iStn, iCol).InUse = True
            Case VBPURGE
                '   Test Time
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1400
                WtcColData(iStn, iCol).DataFormat = "#,###,##0.000"
                WtcColData(iStn, iCol).DataName = "TestTime"
                WtcColData(iStn, iCol).Header1 = "Test Time"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(seconds)"
                WtcColData(iStn, iCol).InUse = True
            Case VBLEAK
                '   Test Time
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1400
                WtcColData(iStn, iCol).DataFormat = "#,###,##0.000"
                WtcColData(iStn, iCol).DataName = "TestTime"
                WtcColData(iStn, iCol).Header1 = "Test Time"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(seconds)"
                WtcColData(iStn, iCol).InUse = True
            Case VBLEAKTEST
                '   Test Time
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1400
                WtcColData(iStn, iCol).DataFormat = "#,###,##0.000"
                WtcColData(iStn, iCol).DataName = "TestTime"
                WtcColData(iStn, iCol).Header1 = "Test Time"
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = "(seconds)"
                WtcColData(iStn, iCol).InUse = True
            Case Else
                ' blank
                iCol = iCol + 1
                WtcColData(iStn, iCol).ColAlign = "CENTER"
                WtcColData(iStn, iCol).ColWidth = 1400
                WtcColData(iStn, iCol).DataFormat = "0"
                WtcColData(iStn, iCol).DataName = "unused"
                WtcColData(iStn, iCol).Header1 = " "
                WtcColData(iStn, iCol).Header2 = " "
                WtcColData(iStn, iCol).Header3 = " "
                WtcColData(iStn, iCol).InUse = False
        End Select
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

Private Sub UpdateDataLayout(iStn As Integer)
Dim alignCode, iShift, iCol, iRow, idx, iLeft As Integer
Dim dataVal As String
SetErrModule 33, 5
If UseLocalErrorHandler Then On Error GoTo localhandler


    ' Align Data
    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    For iRow = 1 To (NrRows - 1)                       ' index to rows on screen; 0=bottom (i.e. the newest)
        Select Case iRow
            Case (NrRows - 1)
                ' ******************
                ' CURRENT DATA ROW
                ' ******************
                iLeft = FirstColLeft
                For iCol = 0 To (NrCols - 1)
                    idx = iCol + (NrDataPerStn * (iStn - 1)) + (NrCols * (iRow - 1))
                    If iCol = (NrCols - 1) Then iLeft = LastColLeft
                    Select Case WtcColData(iStn, iCol).ColAlign
                        Case "LEFT"
                        alignCode = 0
                        Case "CENTER"
                        alignCode = 2
                        Case "RIGHT"
                        alignCode = 1
                    End Select
                    
                    If WtcColData(iStn, iCol).DataName = "Next" Then
                        Select Case StationControl(iStn, iShift).Mode
                            Case VBLOAD
                                lblData(idx).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                            Case VBPURGE
                                lblData(idx).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                            Case VBLEAK
                                lblData(idx).ForeColor = ModeForeColor(StationControl(iStn, iShift).Mode)
                        End Select
                    Else
                        lblData(idx).ForeColor = DataHiLite_ForeColor
                    End If
                    
                    dataVal = " "
                    
                    lblData(idx).Left = iLeft
                    lblData(idx).Width = WtcColData(iStn, iCol).ColWidth
                    lblData(idx).Alignment = alignCode
                    lblData(idx).Caption = dataVal
                    lblData(idx).FontBold = True
                    lblData(idx).FontItalic = False
                
                    iLeft = iLeft + WtcColData(iStn, iCol).ColWidth
                Next iCol
            
            Case Else
                ' ******************
                ' PAST DATA ROWS
                ' ******************
                iLeft = FirstColLeft
                For iCol = 0 To (NrCols - 1)
                    idx = iCol + (NrDataPerStn * (iStn - 1)) + (NrCols * (iRow - 1))
'                    iData = (NrRows - 1) - iRow
                    If iCol = (NrCols - 1) Then iLeft = LastColLeft
                    Select Case WtcColData(iStn, iCol).ColAlign
                        Case "LEFT"
                        alignCode = 0
                        Case "CENTER"
                        alignCode = 2
                        Case "RIGHT"
                        alignCode = 1
                    End Select
                                
                    dataVal = " "
                    
                    lblData(idx).ForeColor = Data_ForeColor
                    lblData(idx).Left = iLeft
                    lblData(idx).Width = WtcColData(iStn, iCol).ColWidth
                    lblData(idx).Alignment = alignCode
                    lblData(idx).Caption = dataVal
                    lblData(idx).FontBold = True
                    lblData(idx).FontItalic = False
                    
                    iLeft = iLeft + WtcColData(iStn, iCol).ColWidth
                Next iCol
        End Select
    Next iRow

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

Private Sub UpdateDataValues(iStn As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 33, 6
Dim iShift, iCol, iRow, idx, iData As Integer
Dim fData As Single
Dim dataVal As String
Dim dTime As Date
Dim tmp As Date

    ' Display Data
    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    
    For iRow = 1 To (NrRows - 1)                       ' index to rows on screen; 0=bottom (i.e. the newest)
        Select Case iRow
            Case (NrRows - 1)
                ' ******************
                ' CURRENT DATA ROW
                ' ******************
                For iCol = 0 To (NrCols - 1)
                    idx = iCol + (NrDataPerStn * (iStn - 1)) + (NrCols * (iRow - 1))
                    Select Case StationControl(iStn, iShift).Mode
                        Case VBLOAD
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        dTime = DateAdd("s", SysConfig.LoadTotal_Interval, CDate(StnLoadData(iStn, iShift, 0).ClkTime))
                                        dataVal = Format(dTime, WtcColData(iStn, iCol).DataFormat)
                                    Case "Time"
                                        dataVal = Format(Now(), WtcColData(iStn, iCol).DataFormat)
                                    Case "NitFlow"
                                        dataVal = Format(Stn_Nit_Flow_PV(iStn, iShift), WtcColData(iStn, iCol).DataFormat)
                                    Case "Mix"
                                        If Stn_Btn_Flow_PV(iStn, iShift) + Stn_Nit_Flow_PV(iStn, iShift) <= 0.0001 Then
                                            fData = 0
                                        Else
                                            fData = 100 * Stn_Btn_Flow_PV(iStn, iShift) / _
                                                (Stn_Btn_Flow_PV(iStn, iShift) + Stn_Nit_Flow_PV(iStn, iShift) + 0.00001)
                                        End If
                                        dataVal = Format(fData, WtcColData(iStn, iCol).DataFormat)
                                    Case "BtnFlow"
                                        dataVal = Format(Stn_Btn_Flow_PV(iStn, iShift), WtcColData(iStn, iCol).DataFormat)
                                    Case "LoadRate"
                                        dataVal = Format(LoadControl(iStn, iShift).LoadRate, WtcColData(iStn, iCol).DataFormat)
                                    Case "LoadGrams"
                                        dataVal = Format(LoadControl(iStn, iShift).loadTotalGrams, WtcColData(iStn, iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StationControl(iStn, iShift).PriScaleWt, WtcColData(iStn, iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(StationControl(iStn, iShift).AuxScaleWt, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(LoadControl(iStn, iShift).TotalWtChg, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChgRate"
                                        dataVal = Format(LoadControl(iStn, iShift).TotalWtChgRate, WtcColData(iStn, iCol).DataFormat)
                                    Case "LiveFuelCycles"
                                        dataVal = Format(StationControl(iStn, iShift).LiveFuelCycleCount, WtcColData(iStn, iCol).DataFormat)
                                    Case "LiveFuelTemp"
                                        dataVal = Format(Stn_AIO(iStn, asFuelTankTemp).EUValue, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StationControl(iStn, iShift).TestTimer, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                        Case VBPURGE
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        tmp = DateAdd("s", SysConfig.PurgeTotal_Interval, CDate(StnPurgeData(iStn, iShift, 0).ClkTime))
                                        dataVal = Format(tmp, WtcColData(iStn, iCol).DataFormat)
                                    Case "Time"
                                        dataVal = Format(Now(), WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeFlow"
                                        dataVal = Format(Stn_AIO(iStn, asPurgeAirFlow).EUValue, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeTemp"
                                        dataVal = Format(PATemp, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeMoist"
                                        dataVal = Format(PAMoisture, WtcColData(iStn, iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StationControl(iStn, iShift).PriScaleWt, WtcColData(iStn, iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(StationControl(iStn, iShift).AuxScaleWt, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(PurgeControl(iStn, iShift).TotalWtChg, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChgRate"
                                        dataVal = Format(PurgeControl(iStn, iShift).TotalWtChgRate, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeVol"
                                        dataVal = Format(PurgeControl(iStn, iShift).Purge_Total, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StationControl(iStn, iShift).TestTimer, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                        Case VBLEAK
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        tmp = DateAdd("s", SysConfig.LeakTotal_Interval, CDate(StnLeakData(iStn, iShift, 0).ClkTime))
                                        dataVal = Format(tmp, WtcColData(iStn, iCol).DataFormat)
                                    Case "Time"
                                        dataVal = Format(Now(), WtcColData(iStn, iCol).DataFormat)
                                    Case "Pressure"
                                        dataVal = Format(PTinvalue, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StationControl(iStn, iShift).TestTimer, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                        Case VBLEAKTEST
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        tmp = DateAdd("s", Cfg_LeakTest.ReportInterval, CDate(StnLT2Data(iStn, iShift, 0).ClkTime))
                                        dataVal = Format(tmp, WtcColData(iStn, iCol).DataFormat)
                                    Case "Time"
                                        dataVal = Format(Now(), WtcColData(iStn, iCol).DataFormat)
                                    Case "N2Flow"
                                        dataVal = Format(Stn_AIO(iStn, asNitrogenFlow).EUValue, WtcColData(iStn, iCol).DataFormat)
                                    Case "N2Temp"
                                        dataVal = Format(Stn_AIO(iStn, asLtN2Temp).EUValue, WtcColData(iStn, iCol).DataFormat)
                                    Case "InPress"
                                        dataVal = Format(Stn_AIO(iStn, asLtInletPress).EUValue, WtcColData(iStn, iCol).DataFormat)
                                    Case "LeakDia"
                                        dataVal = Format(Deff, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StationControl(iStn, iShift).TestTimer, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                        Case Else
                                dataVal = " "
                    End Select
                    
                    lblData(idx).Caption = dataVal
                
                Next iCol
            
            Case Else
                ' ******************
                ' PAST DATA ROWS
                ' ******************
                For iCol = 0 To (NrCols - 1)
                    idx = iCol + (NrDataPerStn * (iStn - 1)) + (NrCols * (iRow - 1))
                    iData = (NrRows - 1) - iRow - 1
                    Select Case StationControl(iStn, iShift).Mode
                        Case VBLOAD
                            If StnLoadData(iStn, iShift, iData).isBlank Then
                                dataVal = " "
                            Else
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        dataVal = " "
                                    Case "Time"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).ClkTime, WtcColData(iStn, iCol).DataFormat)
                                    Case "NitFlow"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).NitFlow, WtcColData(iStn, iCol).DataFormat)
                                    Case "Mix"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).MixPcnt, WtcColData(iStn, iCol).DataFormat)
                                    Case "BtnFlow"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).BtnFlow, WtcColData(iStn, iCol).DataFormat)
                                    Case "LoadRate"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).LoadRate, WtcColData(iStn, iCol).DataFormat)
                                    Case "LoadGrams"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).loadTotalGrams, WtcColData(iStn, iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).PriScle, WtcColData(iStn, iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).AuxScle, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).WtChange, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChgRate"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).WtChgRate, WtcColData(iStn, iCol).DataFormat)
                                    Case "LiveFuelCycles"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).LFcycls, WtcColData(iStn, iCol).DataFormat)
                                    Case "LiveFuelTemp"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).FuelTmp, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StnLoadData(iStn, iShift, iData).TstTimr, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            End If
                        Case VBPURGE
                            If StnPurgeData(iStn, iShift, iData).isBlank Then
                                dataVal = " "
                            Else
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        dataVal = " "
                                    Case "Time"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).ClkTime, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeFlow"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).PrgFlow, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeTemp"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).PrgTemp, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeMoist"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).PrgHumd, WtcColData(iStn, iCol).DataFormat)
                                    Case "PriScale"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).PriScle, WtcColData(iStn, iCol).DataFormat)
                                    Case "AuxScale"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).AuxScle, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChange"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).WtChange, WtcColData(iStn, iCol).DataFormat)
                                    Case "WtChgRate"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).WtChgRate, WtcColData(iStn, iCol).DataFormat)
                                    Case "PurgeVol"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).VolTotl, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StnPurgeData(iStn, iShift, iData).TstTimr, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            End If
                        Case VBLEAK
                            If StnLeakData(iStn, iShift, iData).isBlank Then
                                dataVal = " "
                            Else
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        dataVal = " "
                                    Case "Time"
                                        dataVal = Format(StnLeakData(iStn, iShift, iData).ClkTime, WtcColData(iStn, iCol).DataFormat)
                                    Case "Pressure"
                                        dataVal = Format(StnLeakData(iStn, iShift, iData).Pressure, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StnLeakData(iStn, iShift, iData).TstTimr, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            End If
                        Case VBLEAKTEST
                            If StnLT2Data(iStn, iShift, iData).isBlank Then
                                dataVal = " "
                            Else
                                Select Case WtcColData(iStn, iCol).DataName
                                    Case "Next"
                                        dataVal = " "
                                    Case "Time"
                                        dataVal = Format(StnLT2Data(iStn, iShift, iData).ClkTime, WtcColData(iStn, iCol).DataFormat)
                                    Case "N2Flow"
                                        dataVal = Format(StnLT2Data(iStn, iShift, iData).NitFlow, WtcColData(iStn, iCol).DataFormat)
                                    Case "N2Temp"
                                        dataVal = Format(StnLT2Data(iStn, iShift, iData).NitTemp, WtcColData(iStn, iCol).DataFormat)
                                    Case "InPress"
                                        dataVal = Format(StnLT2Data(iStn, iShift, iData).InPress, WtcColData(iStn, iCol).DataFormat)
                                    Case "LeakDia"
                                        dataVal = Format(StnLT2Data(iStn, iShift, iData).EffDia, WtcColData(iStn, iCol).DataFormat)
                                    Case "TestTime"
                                        dataVal = Format(StnLT2Data(iStn, iShift, iData).SecTimer, WtcColData(iStn, iCol).DataFormat)
                                    Case Else
                                        dataVal = " "
                                End Select
                            End If
                        Case Else
                                dataVal = " "
                    End Select
            
                    lblData(idx).Caption = dataVal
                    
                Next iCol
        End Select
    Next iRow
    
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

Private Sub UpdateHeader(iStn As Integer)
Dim alignCode, iCol, idx, iLeft, iShift As Integer
Dim sTxt As String
SetErrModule 33, 7
If UseLocalErrorHandler Then On Error GoTo localhandler


    ' Setup Data Column Information Array
    SetupColData (iStn)
    
    ' Align & Display Header Text
    iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
    iLeft = FirstColLeft
    For iCol = 0 To (NrCols - 1)
        If iCol = (NrCols - 1) Then iLeft = LastColLeft
        Select Case WtcColData(iStn, iCol).ColAlign
            Case "LEFT"
            alignCode = 0
            Case "CENTER"
            alignCode = 2
            Case "RIGHT"
            alignCode = 1
        End Select
        idx = iCol + (NrHdrPerStn * (iStn - 1))
        sTxt = IIf(Len(WtcColData(iStn, iCol).Header2) > 1, WtcColData(iStn, iCol).Header2, "")
        If WtcColData(iStn, iCol).DataName = "PriScale" And StationRecipe(iStn, iShift).PriScaleNo > 0 Then
                ' Primary Scale
            sTxt = sTxt & " " & Format(StationRecipe(iStn, iShift).PriScaleNo, "#0")
        End If
        If WtcColData(iStn, iCol).DataName = "AuxScale" And StationRecipe(iStn, iShift).AuxScaleNo > 0 Then
                ' Aux Scale
            sTxt = sTxt & " " & Format(StationRecipe(iStn, iShift).AuxScaleNo, "#0")
        End If
        lblHeader(idx).Left = iLeft
        lblHeader(idx).Width = WtcColData(iStn, iCol).ColWidth
        lblHeader(idx).Alignment = alignCode
        lblHeader(idx).Caption = WtcColData(iStn, iCol).Header1 & " " & sTxt
        lblHeader(idx).ToolTipText = Mid(WtcColData(iStn, iCol).Header3, 2, IIf(Len(WtcColData(iStn, iCol).Header3) > 2, Len(WtcColData(iStn, iCol).Header3) - 2, 1))
        lblHeader(idx).ForeColor = ModeBackColor(StationControl(iStn, Stn_ActiveShift(iStn)).Mode)
        iLeft = iLeft + WtcColData(iStn, iCol).ColWidth
    Next iCol
    
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
 
SetErrModule 33, 10101
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
'            mnuSimulation.Enabled = True
        Else
'            mnuSimulation.Enabled = False
            tbrNavigate.Buttons("simulation").Visible = False
            tbrNavigate.Buttons("simulation").Enabled = False
            tbrNavigate.Buttons("simulation").ToolTipText = ""
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



