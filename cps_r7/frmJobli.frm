VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmJoblist 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job List Viewer"
   ClientHeight    =   10830
   ClientLeft      =   90
   ClientTop       =   1830
   ClientWidth     =   14880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJobli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10830
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbrJobList 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "imgNormal"
      DisabledImageList=   "imgDisabled"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "clearjoblist"
            Description     =   "Clear Joblist"
            Object.ToolTipText     =   "Clear Joblist"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deletejob"
            Description     =   "Delete Job"
            Object.ToolTipText     =   "Delete the selected Job"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "backupfiles"
            Description     =   "Backup Files"
            Object.ToolTipText     =   "Backup Files"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "viewjoblog"
            Description     =   "View Job Event Log"
            Object.ToolTipText     =   "View Job Event Log"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previewreports"
            Description     =   "Preview Reports"
            Object.ToolTipText     =   "Preview Reports"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "generatereports"
            Description     =   "Generate Reports"
            Object.ToolTipText     =   "Generate Reports"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "printreports"
            Description     =   "Print Reports"
            Object.ToolTipText     =   "Print Reports"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin Threed.SSPanel pnlMsgJoblist 
         Height          =   630
         Left            =   6840
         TabIndex        =   7
         Top             =   0
         Width           =   6960
         _Version        =   65536
         _ExtentX        =   12277
         _ExtentY        =   1111
         _StockProps     =   15
         Caption         =   "message"
         ForeColor       =   255
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
         BevelInner      =   1
         FloodShowPct    =   0   'False
         Begin VB.CheckBox optPrintSummary 
            Caption         =   "Print Summary Report? "
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
            Left            =   1320
            TabIndex        =   13
            ToolTipText     =   "Check to Print Summary Report"
            Top             =   0
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.CheckBox optPrintDetail 
            Caption         =   "Print Detail Report? "
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
            Left            =   1320
            TabIndex        =   12
            ToolTipText     =   "Check to Print Detail Report"
            Top             =   270
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.CheckBox optRptBackup 
            Caption         =   "Backup Reports? "
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
            TabIndex        =   11
            ToolTipText     =   "File Backup Path"
            Top             =   330
            Width           =   2000
         End
         Begin VB.CheckBox optDbfBackup 
            Caption         =   "Backup DB Files? "
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
            TabIndex        =   10
            ToolTipText     =   "File Backup Path"
            Top             =   90
            Width           =   2000
         End
         Begin VB.TextBox txtRptBackupPath 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   9
            ToolTipText     =   "Backup Path for Report Files"
            Top             =   315
            Width           =   4720
         End
         Begin VB.TextBox txtDbfBackupPath 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            ToolTipText     =   "Backup Path for DB Files"
            Top             =   45
            Width           =   4720
         End
      End
   End
   Begin MSDataGridLib.DataGrid dbgJoblist 
      Align           =   1  'Align Top
      Bindings        =   "frmJobli.frx":57E2
      Height          =   8655
      Left            =   0
      TabIndex        =   6
      Top             =   1230
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   15266
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      ForeColor       =   -2147483635
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List of Jobs"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Job Number"
         Caption         =   "Job Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Description"
         Caption         =   "Job Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Vehicle"
         Caption         =   "Vehicle"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Start Time"
         Caption         =   "Start Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "YYYY MMM DD   HH:MM:SS"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Stop Time"
         Caption         =   "Stop Time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "YYYY MMM DD   HH:MM:SS"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Station"
         Caption         =   "Station"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Shift"
         Caption         =   "Shift"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Report Filename"
         Caption         =   "Report Filename"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3465.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2670.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2564.788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2564.788
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4470.236
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoJoblist 
      Height          =   495
      Left            =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=CpsMaster"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "CpsMaster"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM [Joblist] ORDER BY [Job Number] DESC"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrNavigate 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   953
      _Version        =   393216
   End
   Begin VB.PictureBox pbxBottom 
      Align           =   2  'Align Bottom
      Height          =   460
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   14820
      TabIndex        =   2
      Top             =   10365
      Width           =   14880
      Begin Threed.SSPanel pnlPurgeAir 
         Height          =   405
         Left            =   10110
         TabIndex        =   3
         Top             =   0
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
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
         TabIndex        =   4
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
            TabIndex        =   14
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
         TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   22
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
   Begin VB.Timer tmrScreen 
      Interval        =   250
      Left            =   8280
      Top             =   9960
   End
   Begin MSComctlLib.ImageList imgDisabled 
      Left            =   9960
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":57FB
            Key             =   "backup"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":644D
            Key             =   "generate"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":709F
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":7CF1
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":8943
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":9595
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":A1E7
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":AE39
            Key             =   "close"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":BA8B
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":C6DD
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":D32F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   9360
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":D8BD
            Key             =   "backup"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":E50F
            Key             =   "generate"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":F161
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":FDB3
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":10A05
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":11657
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":122A9
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":12EFB
            Key             =   "close"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":13B4D
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1479F
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":153F1
            Key             =   "compressDB"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgNormal 
      Left            =   8760
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1597F
            Key             =   "backup"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":165D1
            Key             =   "generate"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":17223
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":17E75
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":18AC7
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":19719
            Key             =   "joblog"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1A76B
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1B3BD
            Key             =   "close"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1C00F
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1CC61
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJobli.frx":1D8B3
            Key             =   "compressDB"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblJobData 
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   14280
      TabIndex        =   0
      Top             =   9840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logou&t"
      End
      Begin VB.Menu mnuBackupFiles 
         Caption         =   "&Backup Files"
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
   Begin VB.Menu mnuJobsMenu 
      Caption         =   "&Jobs"
      Begin VB.Menu mnuClearJobs 
         Caption         =   "&Clear All Jobs"
      End
      Begin VB.Menu mnuDeleteJob 
         Caption         =   "&Delete Job"
      End
      Begin VB.Menu mnuPrintJobs 
         Caption         =   "&Print Job List"
      End
   End
   Begin VB.Menu mnuReportsMenu 
      Caption         =   "&Reports"
      Begin VB.Menu mnuDisplayReports 
         Caption         =   "&Display Reports"
      End
      Begin VB.Menu mnuGenerateReports 
         Caption         =   "&Generate Reports"
      End
      Begin VB.Menu mnuPrintReports 
         Caption         =   "&Print Reports"
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
Attribute VB_Name = "frmJoblist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 95 '''''''''form JOBLIST.frm ''''''''''''''''''''''
Option Explicit
Private daodb36 As DAO.Database
Private rS As DAO.Recordset
Dim sourcename, destname As String
Dim sPath, rptPath, rptName As String
Dim rsCrit As String
Dim stopTime As Date
Dim iStn As Integer
Dim iShift As Integer
Dim dbPath, DBFile As String
Dim antiRepeatDelete As Boolean
Dim joblistMsg As String
Dim joblistMsgColor As Long
Dim joblistQuestion As Integer
Private Const qNONE = 0
Private Const qBACKUPFILES = 1
Private Const qPRINTREPORTS = 2
Private Const qCLEARJOBLIST = 3
Private Const SELECTJOBRPTPATH = 2
Private Const SELECTJOBDBFPATH = 3

Private Sub adoJoblist_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    antiRepeatDelete = False
End Sub

Function SetStr(ByVal sInput As String, ByVal iLen As Integer) As String

SetErrModule 95, 2
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim tempstr As String
tempstr = sInput
tempstr = Left(tempstr, iLen)

Do While Len(tempstr) < iLen
  tempstr = tempstr & " "
Loop

SetStr = tempstr
ResetErrModule

Exit Function

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Public Sub RefreshJoblist()
    adoJoblist.Refresh
End Sub

Private Sub dbgJoblist_Click()
    joblistMsgColor = Message_ForeColor
    joblistMsg = " "
End Sub

Private Sub dbgJoblist_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex + 1
        Case 1
            ' Job Number
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Job Number] DESC"
        Case 2
            ' Description
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Description] ASC, [Job Number] DESC"
        Case 3
            ' Vehicle Nr
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Vehicle] ASC, [Job Number] DESC"
        Case 4
            ' Start Time
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Start Time] DESC, [Job Number] DESC"
        Case 5
            ' Stop Time
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Stop Time] DESC, [Job Number] DESC"
        Case 6
            ' Station
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Station] ASC, [Shift] ASC, [Job Number] DESC"
        Case 7
            ' Shift
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Shift] ASC,  [Station] ASC, [Job Number] DESC"
        Case 8
            ' Report Filename
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Report Filename] ASC, [Job Number] DESC"
        Case Else
            ' default
            rsCrit = "SELECT * FROM [Joblist] ORDER BY [Job Number] DESC"
    End Select
    adoJoblist.RecordSource = rsCrit
    RefreshJoblist
End Sub

Private Sub Form_Activate()
Dim station As Integer
Dim Shift As Integer
    For station = 1 To LAST_STN
      For Shift = 1 To NR_SHIFT
        If StationControl(station, Shift).Mode <> VBIDLE Then
          ' No clear button
          tbrJobList.Buttons(3).Enabled = False
          mnuClearJobs.Visible = False
        End If
      Next Shift
    Next station
    joblistMsg = " "
    RefreshJoblist
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmJoblist = Nothing
    End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    dbgJoblist.ForeColor = Data_ForeColor
    adoJoblist.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & FILEPATH_sysdbf & DATAMASTER & ";" _
        & "Persist Security Info=False"
    ' default
    adoJoblist.RecordSource = "SELECT * FROM [Joblist] ORDER BY [Job Number] DESC"
    RefreshJoblist
    frmJoblist.Height = frmMainMenu.Height
    frmJoblist.Width = frmMainMenu.Width
    ' Build Toolbars
    BuildToolbars
    ' Status Bar Setup
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
    pnlPurgeAir.Width = frmJoblist.Width - pnlPurgeAir.Left - 120
    pnlPurgeAir.Top = pnlAlarms.Top
    pnlPurgeAir.Height = pnlAlarms.Height
    pnlPurgeAir.ForeColor = Titles_ForeColor
    ' initialize question/msg box
    QuestionMsg qNONE
    optDbfBackup.Value = IIf(SysConfig.DbFileBackup_Active, 1, 0)
    txtDbfBackupPath.text = SysConfig.DbFileBackup_Path
    optRptBackup.Value = IIf(SysConfig.ReportBackup_Active, 1, 0)
    txtRptBackupPath.text = SysConfig.ReportBackup_Path
    ' print reports
    optPrintDetail.Value = 0
    optPrintSummary.Value = IIf(SysConfig.RptConfig.TextEotSummary_AutoPrint, 1, 0)
    ' Status Bar Update
    UpdateStatusBars
    ' Job List Setup
    dbgJoblist.Height = 8405    '8585, 8885, 8635
    If CheckPass("T", False) Then
        dbgJoblist.AllowDelete = True
    Else
        dbgJoblist.AllowDelete = False
    End If
    dbgJoblist.AllowRowSizing = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub mnuAbout_Click()
    'About
    menuAbout
End Sub

Private Sub mnuBackupFiles_Click()
    ' Begin Backup Files; show Backup Paths & Checkboxes
    QuestionMsg qBACKUPFILES
End Sub

Private Sub mnuCopyFile_Click()
    menuCopyFile
End Sub

Private Sub mnuFirstAid_Click()
    menuFirstAid
End Sub

Private Sub mnuLogin_Click()
    menuLogin
End Sub

Private Sub mnuLogout_Click()
    menuLogout
End Sub

Private Sub mnuPrintFile_Click()
    menuPrintFile
End Sub

Private Sub mnuClearJobs_Click()
    ClearJobs
End Sub

Private Sub mnuDeleteJob_Click()
    ' Delete Job
    DeleteJob
End Sub

Private Sub mnuDisplayReports_Click()
    DisplayJobReports
End Sub

Private Sub mnuExit_Click()
    ' Exit Program
    menuExit
End Sub

Private Sub mnuGenerateReports_Click()
    GenerateJobReports
End Sub

Private Sub mnuOperatorManual_Click()
    menuOperatorManual
End Sub

Private Sub mnuPrintJobs_Click()
    PrintJobs
End Sub

Private Sub mnuPrintReports_Click()
    QuestionMsg qPRINTREPORTS
End Sub

Private Sub tbrJobList_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "clearjoblist"
        ' Ask for Confirmation
        QuestionMsg qCLEARJOBLIST
    Case Is = "deletejob"
        ' Delete Job
        DeleteJob
    Case Is = "backupfiles"
        ' Begin Backup Files; show Backup Paths & Checkboxes
        QuestionMsg qBACKUPFILES
    Case Is = "previewreports"
        ' Display (Preview) Reports
        DisplayJobReports
    Case Is = "generatereports"
        ' Generate Reports
        GenerateJobReports
    Case Is = "printreports"
        ' Print Reports
        QuestionMsg qPRINTREPORTS
    Case Is = "viewjoblog"
        ' View Job Events Log
        ViewJobLog
    Case Is = "ok"
        ' Proceed
        Select Case joblistQuestion
            Case qBACKUPFILES
                ' backup files
                BackupFiles
            Case qPRINTREPORTS
                ' print reports
                PrintJobReports
            Case qCLEARJOBLIST
                ' Clear Joblist
                ClearJobs
            Case Else
                ' nothing to do
        End Select
        ' clear msg box
        QuestionMsg qNONE
    Case Is = "cancel"
        ' Cancel
        QuestionMsg qNONE
        joblistMsgColor = Message_ForeColor
        joblistMsg = ""
    Case Is = "close"
        ' Close Screen
        CloseScreen
End Select
End Sub

Private Sub tmrScreen_Timer()
    UpdateStatusBars
    UpdateNavigateBtns
End Sub

Public Sub UpdateStatusBars()
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
    pnlMsgJoblist.ForeColor = joblistMsgColor
    pnlMsgJoblist.Caption = joblistMsg
End Sub

Private Sub BackupFiles()
Dim fs As Object
Dim filesCopied As Integer

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    filesCopied = 0
    joblistMsg = ""
    If adoJoblist.Recordset.BOF Then
        joblistMsgColor = MEDRED
        joblistMsg = "No Job Data Available"
    Else
    
        If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid Job Number"
            Exit Sub
        Else
            DBFile = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
            For iStn = 1 To LAST_STN
              For iShift = 1 To NR_SHIFT
                If (StationControl(iStn, iShift).Job_Number) = DBFile Then
                  ' Report an error
                  joblistMsgColor = MEDRED
                  joblistMsg = "That Job#" & DBFile & "  has not been completed yet by Station " & iStn & " Shift " & iShift
                  Exit Sub
                End If
              Next iShift
            Next iStn
        End If
        
        If IsNull(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid filename"
            Exit Sub
        End If
        
    
        MousePointer = vbHourglass
        ' Report Files
        If optRptBackup.Value = cYES Then
            If fs.FolderExists(txtRptBackupPath.text) Then
                
                ' report filename is column 7
                rptName = dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))
                
                ' text summary report
                sourcename = FILEPATH_reports & rptName & "Summary.RPT"
                destname = txtRptBackupPath.text & Left(rptName, 64) & "Summary.RPT"
                If fs.FileExists(sourcename) Then
                    FileCopy sourcename, destname
                    filesCopied = filesCopied + 1
                End If
                
                ' text detail report
                sourcename = FILEPATH_reports & rptName & "Detail.RPT"
                destname = txtRptBackupPath.text & Left(rptName, 64) & "Detail.RPT"
                If fs.FileExists(sourcename) Then
                    FileCopy sourcename, destname
                    filesCopied = filesCopied + 1
                End If
                
                ' xls report
                sourcename = FILEPATH_reports & rptName & ".xls"
                destname = txtRptBackupPath.text & Left(rptName, 64) & ".XLS"
                If fs.FileExists(sourcename) Then
                    FileCopy sourcename, destname
                    filesCopied = filesCopied + 1
                End If
                
                ' csv summary report
                sourcename = FILEPATH_reports & rptName & "Summary.CSV"
                destname = txtRptBackupPath.text & Left(rptName, 64) & "Summary.CSV"
                If fs.FileExists(sourcename) Then
                    FileCopy sourcename, destname
                    filesCopied = filesCopied + 1
                End If
                
                ' csv detail report
                sourcename = FILEPATH_reports & rptName & "Detail.CSV"
                destname = txtRptBackupPath.text & Left(rptName, 64) & "Detail.CSV"
                If fs.FileExists(sourcename) Then
                    FileCopy sourcename, destname
                    filesCopied = filesCopied + 1
                End If
            
            Else  ' backup path doesn't exist
                
                Delay_Box "Backup Path >" & txtRptBackupPath.text & "< Not defined; ABORTING Backup", MSGDELAY, msgSHOW
        
            End If
        End If
            
        If optDbfBackup.Value = cYES Then
            If fs.FolderExists(txtDbfBackupPath.text) Then
                
                ' db file
                DBFile = dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))
                sourcename = FILEPATH_data & "C" & DBFile & AccessDbFileExt
                destname = txtDbfBackupPath.text & "C" & DBFile & AccessDbFileExt
                If fs.FileExists(sourcename) Then
                    FileCopy sourcename, destname
                    filesCopied = filesCopied + 1
                End If
            
            Else  ' backup path doesn't exist
                
                Delay_Box "Backup Path >" & txtDbfBackupPath.text & "< Not defined; ABORTING Backup", MSGDELAY, msgSHOW
            
            End If
        
        End If
                
        DoEvents
        MousePointer = vbDefault
                
        ' Backup Complete (or is it???)
        joblistMsgColor = IIf(Len(joblistMsg) < 2, Message_ForeColor, joblistMsgColor)
        If (Len(joblistMsg) < 2) Then
            joblistMsg = IIf((filesCopied = 0), "No Files to Copy", "Job Files Backup Complete")
        End If
            
    End If
    
    joblistQuestion = 0
    QuestionMsg joblistQuestion
    tbrJobList.Buttons("fillright").Width = 3820
    tbrJobList.Buttons("ok").Visible = False
    tbrJobList.Buttons("cancel").Visible = False
End Sub

Private Sub GenerateJobReports()
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Dim JobNumStr As String
Dim cmdLine As String
Dim rptgenCmdInt As Integer
Dim rptgenCmdCode As String
Dim cfgBits(0 To 9) As Boolean
    
    joblistMsg = ""
    
    ' JOB EXISTS ??
    If adoJoblist.Recordset.BOF Then
        joblistMsgColor = MEDRED
        joblistMsg = "No Job Data Available"
    Else
    
        ' JOB OK FOR REPORTING ??
        ' JOB OK FOR REPORTING ??
        ' JOB OK FOR REPORTING ??
        '
        If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid Job Number"
            Exit Sub
        Else
            DBFile = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
            For iStn = 1 To LAST_STN
              For iShift = 1 To NR_SHIFT
                If (StationControl(iStn, iShift).Job_Number) = DBFile Then
                  ' Report an error
                  joblistMsgColor = MEDRED
                  joblistMsg = "That Job#" & DBFile & "  has not been completed yet by Station " & iStn & " Shift " & iShift
                  Exit Sub
                End If
              Next iShift
            Next iStn
        End If
        
        If Not IsDate(dbgJoblist.Columns(4).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "That test has no Stop Time. The data may be incomplete or corrupt."
            Exit Sub
        End If
    
        If Not fs.FileExists(FILEPATH_data & "C" & CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) & AccessDbFileExt) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "The Database File Does Not Exist."
            Exit Sub
        End If
        
        
        ' DELETE EXISTING XLS & CSV FILES
        ' DELETE EXISTING XLS & CSV FILES
        ' DELETE EXISTING XLS & CSV FILES
        '
        ' report filename is column 7
        rptName = dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))
        
        ' delete existing xls report file
        If SysConfig.RptConfig.XlsGenReporting Then
            ' xls report
            rptPath = FILEPATH_reports & rptName & ReportsXlsFileExt
            ' delete the existing xls report
            If fs.FileExists(rptPath) Then Kill rptPath
            DoEvents
        End If
        
        ' delete existing csv report files
        If SysConfig.RptConfig.CsvGenReporting Then
            ' report filepath
            rptPath = FILEPATH_reports & rptName & ".csv"
            ' delete the existing report
            If fs.FileExists(rptPath) Then Kill rptPath
            DoEvents
            ' summary report
            If SysConfig.RptConfig.CsvGenSummary Then
                ' report filepath
                rptPath = FILEPATH_reports & rptName & "_summary.csv"
                ' delete the existing report
                If fs.FileExists(rptPath) Then Kill rptPath
                DoEvents
            End If
            ' detail report
            If SysConfig.RptConfig.CsvGenDetail Then
                ' report filepath
                rptPath = FILEPATH_reports & rptName & "_detail.csv"
                ' delete the existing report
                If fs.FileExists(rptPath) Then Kill rptPath
                DoEvents
            End If
        End If
        
        
        ' REQUEST REPORT GENERATION
        ' REQUEST REPORT GENERATION
        ' REQUEST REPORT GENERATION
        '
        ' combine report config Bits
        '
        '   bit 0 = TextReporting
        '   bit 1 = TextSummary
        '   bit 2 = TextSummary_AutoPrint
        '   bit 3 = TextDetail
        '
        '   bit 4 = XlsReporting
        '   bit 5 = XlsSummary
        '   bit 6 = XlsDetail
        '
        '   bit 7 = XlsReporting
        '   bit 8 = XlsSummary
        '   bit 9 = XlsDetail
        '
        With SysConfig.RptConfig
            cfgBits(0) = .TextGenReporting
            cfgBits(1) = .TextGenSummary
            cfgBits(2) = False
            cfgBits(3) = .TextGenDetail
            cfgBits(4) = .XlsGenReporting
            cfgBits(5) = .XlsGenSummary
            cfgBits(6) = .XlsGenDetail
            cfgBits(7) = .CsvGenReporting
            cfgBits(8) = .CsvGenSummary
            cfgBits(9) = .CsvGenDetail
            rptgenCmdInt = Bits_Pack(cfgBits(0), cfgBits(1), cfgBits(2), cfgBits(3), cfgBits(4), cfgBits(5), cfgBits(6), cfgBits(7), cfgBits(8), cfgBits(9), False, False, False, False, False)
        End With
        rptgenCmdCode = Format(rptgenCmdInt, "000000")
    
        ' request reports be generated
        JobNumStr = Trim(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
        dbPath = FILEPATH_data & "C" & JobNumStr & AccessDbFileExt
        ReportGenerate_dbPath = dbPath
'        ReportGenerate_Request = True
        cmdLine = filepath & "\cps_r7_Reporter.exe  " & JobNumStr & "  " & rptgenCmdCode
        Shell cmdLine, vbNormalFocus
                        
    End If
End Sub

Private Sub DisplayJobReports()
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
    
    joblistMsg = ""
    
    If adoJoblist.Recordset.BOF Then
        joblistMsgColor = MEDRED
        joblistMsg = "No Job Data Available"
    Else
    
        If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid Job Number"
            Exit Sub
        Else
            DBFile = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
            For iStn = 1 To LAST_STN
              For iShift = 1 To NR_SHIFT
                If (StationControl(iStn, iShift).Job_Number) = DBFile Then
                  ' Report an error
                  joblistMsgColor = MEDRED
                  joblistMsg = "That Job#" & DBFile & "  has not been completed yet by Station " & iStn & " Shift " & iShift
                  Exit Sub
                End If
              Next iShift
            Next iStn
        End If
        
        If IsNull(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid filename"
            Exit Sub
        End If
        
        If Not IsDate(dbgJoblist.Columns(4).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "That test has no Stop Time, the data may be incomplete or corrupt"
            Exit Sub
        End If
        
        MousePointer = vbHourglass
        DoEvents
        ' report filename is column 7
        rptName = dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))
        ' summary report
        rptPath = FILEPATH_reports & rptName & "Summary.RPT"
        If Not fs.FileExists(rptPath) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "The Summary Report File Does Not Exist."
        Else
            ' display summary report
            sourcename = "notepad " & rptPath
            Shell sourcename
            DoEvents
        End If
        ' detail report
        rptPath = FILEPATH_reports & rptName & "Detail.RPT"
        If Not fs.FileExists(rptPath) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "The Detail Report File Does Not Exist."
        Else
            ' display detail report
            sourcename = "notepad " & rptPath
            Shell sourcename
            DoEvents
        End If
        ' Reports Sent to Notepad
        If Len(joblistMsg) < 2 Then joblistMsgColor = Message_ForeColor
        If Len(joblistMsg) < 2 Then joblistMsg = "Reports Sent to Notepad"
        DoEvents
        MousePointer = vbDefault
    End If
End Sub

Private Sub PrintJobReports()
' Created By:       Brunrose
' Description:      This procedure prints the selected Job Reports
'
SetErrModule 95, 111
If UseLocalErrorHandler Then On Error GoTo localhandler

If (optPrintDetail.Value + optPrintSummary.Value = 0) Then
    ' nothing to do
    joblistMsgColor = MEDRED
    joblistMsg = "No Reports Selected"
    Exit Sub
End If

Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
    
    joblistMsg = ""
    
    If adoJoblist.Recordset.BOF Then
        joblistMsgColor = MEDRED
        joblistMsg = "No Job Data Available"
    Else
    
        If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid Job Number"
            Exit Sub
        Else
            DBFile = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
            For iStn = 1 To LAST_STN
              For iShift = 1 To NR_SHIFT
                If (StationControl(iStn, iShift).Job_Number) = DBFile Then
                  ' Report an error
                  joblistMsgColor = MEDRED
                  joblistMsg = "That Job#" & DBFile & "  has not been completed yet by Station " & iStn & " Shift " & iShift
                  Exit Sub
                End If
              Next iShift
            Next iStn
        End If
        
        If IsNull(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid filename"
            Exit Sub
        End If
        
        If Not IsDate(dbgJoblist.Columns(4).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            joblistMsgColor = MEDRED
            joblistMsg = "That test has no Stop Time, the data may be incomplete or corrupt"
            Exit Sub
        End If
        
        MousePointer = vbHourglass
        DoEvents
        ' report filename is column 7
        rptName = dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))
        If optPrintSummary = 1 Then
            ' Print Summary Report
            rptPath = FILEPATH_reports & rptName & "Summary.RPT"
            If Not fs.FileExists(rptPath) Then
                ' Report an error
                joblistMsgColor = MEDRED
                joblistMsg = "The Summary Report File Does Not Exist."
            Else
                ' print summary report
                joblistMsgColor = Message_ForeColor
                joblistMsg = "                                       Sending Summary Report to the Printer"
                Print_File CStr(rptPath)
                DoEvents
            End If
        End If
        If optPrintDetail = 1 Then
            ' Print Detail Report
            rptPath = FILEPATH_reports & rptName & "Detail.RPT"
            If Not fs.FileExists(rptPath) Then
                ' Report an error
                joblistMsgColor = MEDRED
                joblistMsg = "The Detail Report File Does Not Exist."
            Else
                ' print detail report
                joblistMsgColor = Message_ForeColor
                joblistMsg = "                                       Sending Detail Report to the Printer"
                Print_File CStr(rptPath)
                DoEvents
            End If
        End If
        ' Reports Sent to Printer
        If joblistMsgColor <> MEDRED Then
            joblistMsgColor = Message_ForeColor
            If (optPrintDetail + optPrintSummary = 1) Then joblistMsg = "        Report Sent to " & PRINTERNAME
            If (optPrintDetail + optPrintSummary = 2) Then joblistMsg = "        Reports Sent to " & PRINTERNAME
        End If
        DoEvents
        MousePointer = vbDefault
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

Private Sub ClearJobs()
    SetErrModule 95, 3
    If UseLocalErrorHandler Then On Error GoTo localhandler
    For iStn = 1 To LAST_STN
      For iShift = 1 To NR_SHIFT
        If StationControl(iStn, iShift).Mode <> VBIDLE Then
          ' Report an error
          joblistMsgColor = MEDRED
          joblistMsg = "Station " & iStn & " Shift " & iShift & " MUST BE IDLE"
          Exit Sub
        End If
      Next iShift
    Next iStn
    If CheckPass("T", True) Then
       ' Clearing Job List
       joblistMsgColor = Message_ForeColor
       joblistMsg = "Clearing Job List.. Please Wait"
        adoJoblist.Recordset.MoveLast
        Do Until adoJoblist.Recordset.BOF
          adoJoblist.Recordset.MoveLast
          adoJoblist.Recordset.Delete
          adoJoblist.Recordset.MovePrevious
        Loop
       joblistMsgColor = Message_ForeColor
       joblistMsg = "Job List Cleared"
    End If
    ResetErrModule
Exit Sub

localhandler:
    joblistMsgColor = MEDRED
    joblistMsg = "Unable to Clear Job List"
End Sub

Private Sub DeleteJob()
SetErrModule 95, 31
If UseLocalErrorHandler Then On Error GoTo localhandler
    If Not antiRepeatDelete Then
        If adoJoblist.Recordset.BOF Then
            joblistMsgColor = MEDRED
            joblistMsg = "No Job Data Available"
        Else
        
            If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
                ' Report an error
                joblistMsgColor = MEDRED
                joblistMsg = "Invalid Job Number"
                Exit Sub
            Else
                DBFile = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
                For iStn = 1 To LAST_STN
                  For iShift = 1 To NR_SHIFT
                    If (StationControl(iStn, iShift).Job_Number) = DBFile Then
                      ' Report an error
                      joblistMsgColor = MEDRED
                      joblistMsg = "Job #" & DBFile & "  is still open"
                      Exit Sub
                    End If
                  Next iShift
                Next iStn
            End If
            
            adoJoblist.Recordset.Delete
            joblistMsgColor = Message_ForeColor
            joblistMsg = "Job Deleted"
            antiRepeatDelete = True
           
        End If
    End If
Exit Sub
localhandler:
    joblistMsgColor = MEDRED
    joblistMsg = "Unable to Delete Job"
End Sub

Private Sub PrintJobs()
' Created By:       Brunrose
' Description:      This procedure prints the Job List File
'
SetErrModule 95, 1
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim Idx, xStartTime, yData, numPages As Integer
Dim numJobs As Long
Dim printstring As String
Dim rs_shift, rs_stn, rs_start, rs_stop, rs_job, rs_veh, rs_desc, sdate As String
Dim oldFont As New StdFont
    ' Save current printer font
    oldFont = Printer.Font
    sPath = FILEPATH_sysdbf & DATAMASTER
    Set daodb36 = DBEngine.OpenDatabase(sPath)
    Set rS = daodb36.OpenRecordset("joblist")
    adoJoblist.Recordset.MoveFirst
    numJobs = adoJoblist.Recordset.RecordCount
    sdate = "yyyy mmm dd  hh:mm"
    ' Title
    ' set font
    Printer.Font = REPORTFONT
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    Print_Center ("JOB LISTING")
    Print_Line " "
    ' Header
    ' set font
    Printer.Font.Size = 10
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    Print_Center "Canister Preconditioning System"
    Print_Center Trim$(SysConfig.Heading)
    Print_Center Trim$(SysConfig.Heading2)
    Print_Center (Format(numJobs, "######0") & " Jobs as of " & Format(Now, "yyyy mmmm dd"))
    Print_Line " "
    Print_Line " "
    Printer.Font.Size = 8.5
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    ' Print Column Headings
    Printer.Print "Job Number", "Description", "", "", "Vehicle", "", "Start Time", "", "Stop Time", "", "Station Shift"
    printstring = _
    "____________________________________________________________________________________________________________________________"
    Print_Line printstring
    xStartTime = 6210
    numPages = (numJobs \ 50)
    If (numJobs Mod 50) > 0 Then numPages = numPages + 1
    Idx = 1
    Do Until (adoJoblist.Recordset.EOF)
'    Do Until (adoJoblist.Recordset.EOF Or idx > 39)
        ' job number
        If IsNull(adoJoblist.Recordset("Job Number")) Then
            rs_job = " "
        Else
            rs_job = Mid$(adoJoblist.Recordset("Job Number"), 1, 6)
        End If
        rs_job = SetStr(rs_job, 6)
        ' description
        If IsNull(adoJoblist.Recordset("Description")) Then
            rs_desc = " "
        Else
            rs_desc = Trim$(Mid$(adoJoblist.Recordset("Description"), 1, 38))
        End If
        rs_desc = SetStr(rs_desc, 40)
        ' vehicle
        If IsNull(adoJoblist.Recordset("Vehicle")) Then
            rs_veh = " "
        Else
            rs_veh = Trim$(Mid$(adoJoblist.Recordset("Vehicle"), 1, 20))
        End If
        rs_veh = SetStr(rs_veh, 22)
        ' start time
        If IsNull(adoJoblist.Recordset("Start Time")) Then
            rs_start = " "
        ElseIf Not IsDate(adoJoblist.Recordset("Start Time")) Then
            rs_start = " "
        Else
            rs_start = Format(adoJoblist.Recordset("Start Time"), sdate)
        End If
        rs_start = SetStr(rs_start, 19)
        ' stop time
        If IsNull(adoJoblist.Recordset("Stop Time")) Then
            rs_stop = " "
        ElseIf Not IsDate(adoJoblist.Recordset("Stop Time")) Then
            rs_stop = " "
        Else
            rs_stop = Format(adoJoblist.Recordset("Stop Time"), sdate)
        End If
        rs_stop = SetStr(rs_stop, 22)
        ' station
        If IsNull(adoJoblist.Recordset("Station")) Then
            rs_stn = " "
        ElseIf Not IsNumeric(adoJoblist.Recordset("Station")) Then
            rs_stn = " "
        Else
            rs_stn = Format(adoJoblist.Recordset("Station"), "0")
        End If
        rs_stn = SetStr(rs_stn, 11)
        ' shift
        If IsNull(adoJoblist.Recordset("Shift")) Then
            rs_shift = " "
        ElseIf Not IsNumeric(adoJoblist.Recordset("Shift")) Then
            rs_shift = " "
        Else
            rs_shift = Format(adoJoblist.Recordset("Shift"), "0")
        End If
        rs_shift = SetStr(rs_shift, 2)
        ' remember CurrentY position
        yData = Printer.CurrentY
        ' print job,desc,vehicle
        Printer.Print rs_job, rs_desc, rs_veh
        ' print starttime, stoptime, station, shift
        Printer.CurrentX = xStartTime
        Printer.CurrentY = yData
        Printer.Print rs_start, rs_stop, rs_stn & rs_shift
        ' end of page?
        If adoJoblist.Recordset.EOF Or (Idx Mod 50) = 0 Then
            ' print footer
            Print_Footer numPages
            ' more pages?
            If Not adoJoblist.Recordset.EOF Then
                ' new page
                Printer.NewPage
                ' Title
                ' set font
                Printer.Font = REPORTFONT
                Printer.Font.Size = 12
                Printer.Font.Bold = False
                Printer.Font.Italic = False
                Print_Center ("JOB LISTING")
                Print_Line " "
                ' Header
                ' set font
                Printer.Font.Size = 10
                Printer.Font.Bold = False
                Printer.Font.Italic = False
                Print_Center "Canister Preconditioning System"
                Print_Center Trim$(SysConfig.Heading)
                Print_Center Trim$(SysConfig.Heading2)
                Print_Center (Format(numJobs, "######0") & " Jobs as of " & Format(Now, "yyyy mmmm dd"))
                Print_Line " "
                Print_Line " "
                Printer.Font.Size = 8.5
                Printer.Font.Bold = False
                Printer.Font.Italic = False
                ' Print Header
                ' 23456789^123456789^123456789^123456789^123456789^123456789^123456789^1234567890^1234567890^1234567890
                'printstring = _
                '"Job Number     Vehicle             Start Time           Stop Time        Station        Shift"
                'Print_Line printstring
                Printer.Print "Job Number", "Description", "", "", "Vehicle", "", "Start Time", "", "Stop Time", "", "Station Shift"
                printstring = _
                "____________________________________________________________________________________________________________________________"
                Print_Line printstring
            End If
        End If
        adoJoblist.Recordset.MoveNext
        Idx = Idx + 1
    Loop
    
    adoJoblist.Recordset.MoveFirst
      
    Print_Footer numPages
    Printer.EndDoc
    Printer.Font = oldFont
    
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
    btnX.ToolTipText = "System Definition Screen"
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
    btnX.ToolTipText = "Watch Data Screen"
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
    

    
    
    
    
    ' ***************
    ' Joblist Toolbar
    ' ***************
    
    'Spacer
    Set btnX = tbrJobList.Buttons.Add(, "fillright", , tbrPlaceholder)
'    btnX.Width = 2650 ' Placeholder width to accommodate a textbox.
    btnX.Width = 3820 ' Placeholder width to accommodate a textbox.
    
    'OK
    Set btnX = tbrJobList.Buttons.Add(, "ok", , tbrDefault, "ok")
    btnX.ToolTipText = "OK"
    btnX.Description = btnX.ToolTipText
    btnX.Visible = False
    
    'Cancel
    Set btnX = tbrJobList.Buttons.Add(, "cancel", , tbrDefault, "cancel")
    btnX.ToolTipText = "cancel"
    btnX.Description = btnX.ToolTipText
    btnX.Visible = False
    
'    Set btnX = tbrJobList.Buttons.Add(, , , tbrSeparator)
    
    'Message Box
    Set btnX = tbrJobList.Buttons.Add(, "msgbox", , tbrPlaceholder)
    btnX.Width = 6960 ' Placeholder width to accommodate a textbox.
    
    Set btnX = tbrJobList.Buttons.Add(, , , tbrSeparator)
            

    With pnlMsgJoblist
        .Height = tbrJobList.Height + 60
        .Width = tbrJobList.Buttons("msgbox").Width
'        .Top = txtDspStn.Top
        .Left = tbrJobList.Buttons("msgbox").Left
        .FontBold = True
'        .Locked = True
    End With


    ' disable PrintReports button if no printer
    tbrJobList.Buttons("printreports").Enabled = PRINTERAVAILABLE

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
    Set frmJoblist = Nothing
End Sub

Public Sub UpdateNavigateBtns()

'
' Routine Name:  UpdateNavigateBtns
' Description:
' Updates the Navigate & JobList toolbar buttons
'
 
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 95, 10101
Dim AllStnIdle As Boolean
Dim iKeyCount As Integer
Dim iShift As Integer
Dim iStation As Integer

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
'            mnuCanisters.Enabled = True
        Else
            tbrNavigate.Buttons("canisters").Enabled = False
'            mnuCanisters.Enabled = False
        End If
        
        ' Recipes
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("recipes").Enabled = True
'            mnuRecipes.Enabled = True
        Else
            tbrNavigate.Buttons("recipes").Enabled = False
'            mnuRecipes.Enabled = False
        End If
        
        ' Purge Profiles
        If CheckPass("N", False) Then
            tbrNavigate.Buttons("purgeprofile").Enabled = True
'            mnuPurgeProfiles.Enabled = True
        Else
            tbrNavigate.Buttons("purgeprofile").Enabled = False
'            mnuPurgeProfiles.Enabled = False
        End If
        
        ' Courses
        If CheckPass("N", False) And (NR_JOBSEQ > 1) Then
            tbrNavigate.Buttons("courses").Visible = True
'            mnuCourses.Visible = True
        Else
            tbrNavigate.Buttons("courses").Visible = False
'            mnuCourses.Visible = False
        End If
        
        ' TomCanLoad
        If CheckPass("N", False) And (USINGREMCANLOAD Or USINGTOMCANLOAD) Then
            tbrNavigate.Buttons("tomcanload").Visible = True
'            mnuTomCanLoad.Visible = True
        Else
            tbrNavigate.Buttons("tomcanload").Visible = False
'            mnuTomCanLoad.Visible = False
        End If
        
        ' Configuration
        If CheckPass("B", False) Then
            tbrNavigate.Buttons("configuration").Enabled = True
'            mnuConfiguration.Enabled = True
        Else
'            mnuConfigurationEnabled = False
            tbrNavigate.Buttons("configuration").Enabled = False
        End If
        
        ' System Definition
        If CheckPass("H", False) Then
            tbrNavigate.Buttons("sysdef").Visible = True
            tbrNavigate.Buttons("sysdef").ToolTipText = "System Definition"
'            mnuSysDef.Visible = True
        Else
            tbrNavigate.Buttons("sysdef").Visible = False
            tbrNavigate.Buttons("sysdef").ToolTipText = ""
'            mnuSysDef.Visible = False
        End If
        
        ' Butane Available
        If systemhasBUTANE Then
            tbrNavigate.Buttons("butane").Visible = True
'            mnuButane.Enabled = True
        Else
            tbrNavigate.Buttons("butane").Visible = False
'            mnuButane.Enabled = False
        End If
        
        ' Event Log
        If CheckPass("Z", False) Then
            tbrNavigate.Buttons("eventlog").Enabled = True
'            mnuEventLog.Enabled = True
        Else
'            mnuEventLogEnabled = False
            tbrNavigate.Buttons("eventlog").Enabled = False
        End If
        
        ' File Maintenance Log
        If CheckPass("L", False) Then
'            tbrNavigate.Buttons("filelog").Enabled = True
'            mnuFileLog.Enabled = True
        Else
'            mnuFileLog.Enabled = False
'            tbrNavigate.Buttons("filelog").Enabled = False
        End If
        
        ' Joblist Log
        If CheckPass("M", False) Then
            tbrNavigate.Buttons("joblist").Enabled = True
'            mnuJoblist.Enabled = True
        Else
'            mnuJoblist.Enabled = False
            tbrNavigate.Buttons("joblist").Enabled = False
        End If
        
        ' Review Previous Cycle Data
        If CheckPass("F", False) Then
            tbrNavigate.Buttons("reviewdata").Enabled = True
'            mnuReviewData.Enabled = True
        Else
'            mnuReviewData.Enabled = False
            tbrNavigate.Buttons("reviewdata").Enabled = False
        End If
        
        ' Watch Current Cycle Data
        If CheckPass("F", False) Then
            tbrNavigate.Buttons("watchdata").Enabled = True
'            mnuWatchData.Enabled = True
        Else
'            mnuWatchData.Enabled = False
            tbrNavigate.Buttons("watchdata").Enabled = False
        End If
        
        ' MFC Calibration
        If CheckPass("X", False) Then
            tbrNavigate.Buttons("calibration").Enabled = True
'            mnuCalibration.Enabled = True
        Else
'            mnuCalibration.Enabled = False
            tbrNavigate.Buttons("calibration").Enabled = False
        End If
        
        ' I/O Monitor
        If CheckPass("2", False) Then
            tbrNavigate.Buttons("iomonitor").Enabled = True
'            mnuIoMonitor.Enabled = True
        Else
'            mnuIoMonitor.Enabled = False
            tbrNavigate.Buttons("iomonitor").Enabled = False
        End If
        
        ' Scale Monitor
        If CheckPass("3", False) Then
            tbrNavigate.Buttons("scalemonitor").Enabled = True
'            mnuScaleMonitor.Enabled = True
        Else
'            mnuScaleMonitor.Enabled = False
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
'                    mnuak_server.Enabled = True
                Else
'                    mnuak_server.Enabled = False
                    tbrNavigate.Buttons("ak_server").Enabled = False
                End If
        
            Case pagClient
                ' AK Client PurgeAir Generator control
                ' AK Client
                If CheckPass("7", False) Then
                    tbrNavigate.Buttons("ak_client").Enabled = True
'                   mnuAk_Client.Enabled = True
                Else
'                   mnuAk_Client.Enabled = False
                    tbrNavigate.Buttons("ak_client").Enabled = False
                End If
                ' AK Server
                If CheckPass("7", False) Then
                    tbrNavigate.Buttons("ak_server").Enabled = True
'                    mnuAk_Server.Enabled = True
                Else
'                    mnuAk_Server.Enabled = False
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
'            mnuAirLog.Enabled = True
        Else
'            mnuAirLog.Enabled = False
        End If
        
        ' Exit Program
        If CheckPass("G", False) Then
'            tbrNavigate.Buttons("exit").Enabled = True
            mnuExit.Enabled = True
        Else
            mnuExit.Enabled = False
'            tbrNavigate.Buttons("exit").Enabled = False
        End If
        
        
        
        ' *************** JOBLIST TOOLBAR ***************
        ' *************** JOBLIST TOOLBAR ***************
        ' *************** JOBLIST TOOLBAR ***************
        
        ' Generate Reports Button
        AllStnIdle = True
'        For iStation = 1 To LAST_STN
'            For iShift = 1 To NR_SHIFT
'                If (Not StationControl(iStation, iShift).ModeIsIdle_Debounced) Then AllStnIdle = False
'            Next iShift
'        Next iStation
'        tbrJobList.Buttons("generatereports").Enabled = IIf(AllStnIdle, True, False)
        tbrJobList.Buttons("generatereports").Enabled = True
    
        Select Case joblistQuestion
            Case qBACKUPFILES, qPRINTREPORTS
                ' DisableJoblistButtons
            Case Else
                ' no question
                ' Clear Joblist Button
                tbrJobList.Buttons("clearjoblist").Enabled = IIf(CheckPass("T", False), True, False)
                ' Delete Job Button
                tbrJobList.Buttons("deletejob").Enabled = IIf(CheckPass("T", False), True, False)
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

Public Sub Reconnect()
    Dim cnn As ADODB.Connection
   
    Set cnn = New ADODB.Connection
    cnn.ConnectionString = _
     "Provider=Microsoft.Jet.OLEDB.4.0;" & _
     "Data Source=" & _
      FILEPATH_sysdbf & "cpsmaster_rev05" & AccessDbFileExt
    cnn.Open
'    Debug.Print cnn.ConnectionString
    Set cnn = Nothing
End Sub

Private Sub txtDbfBackupPath_Click()
    frmBackupPath.ChangeBackupSelect SELECTJOBDBFPATH
    frmBackupPath.Show
End Sub

Private Sub txtDbfBackupPath_DblClick()
    frmBackupPath.ChangeBackupSelect SELECTJOBDBFPATH
    frmBackupPath.Show
End Sub

Private Sub txtRptBackupPath_Click()
    frmBackupPath.ChangeBackupSelect SELECTJOBRPTPATH
    frmBackupPath.Show
End Sub

Private Sub txtRptBackupPath_DblClick()
    frmBackupPath.ChangeBackupSelect SELECTJOBRPTPATH
    frmBackupPath.Show
End Sub

Private Sub QuestionMsg(ByVal whichQuestion As Integer)
    joblistQuestion = whichQuestion
    Select Case joblistQuestion
        Case qBACKUPFILES
            DisableJoblistButtonsExcept tbrJobList.Buttons("backupfiles").Index
            optDbfBackup.Left = 90
            optDbfBackup.Top = 90
            optDbfBackup.Visible = True
            txtDbfBackupPath.Visible = True
            optRptBackup.Left = 90
            optRptBackup.Top = 330
            optRptBackup.Visible = True
            txtRptBackupPath.Visible = True
            optPrintDetail.Visible = False
            optPrintSummary.Visible = False
            tbrJobList.Buttons("fillright").Width = 2650
            tbrJobList.Buttons("ok").Visible = True
            tbrJobList.Buttons("cancel").Visible = True
            joblistMsgColor = Message_ForeColor
            joblistMsg = ""
        Case qPRINTREPORTS
            DisableJoblistButtonsExcept tbrJobList.Buttons("printreports").Index
            optDbfBackup.Visible = False
            txtDbfBackupPath.Visible = False
            optRptBackup.Visible = False
            txtRptBackupPath.Visible = False
            optPrintDetail.Left = 90
            optPrintDetail.Top = 90
            optPrintDetail.Visible = True
            optPrintSummary.Left = 90
            optPrintSummary.Top = 330
            optPrintSummary.Visible = True
            tbrJobList.Buttons("fillright").Width = 2650
            tbrJobList.Buttons("ok").Visible = True
            tbrJobList.Buttons("cancel").Visible = True
            joblistMsgColor = Message_ForeColor
            joblistMsg = "                 Select Reports to Print"
        Case qCLEARJOBLIST
            DisableJoblistButtonsExcept tbrJobList.Buttons("clearjoblist").Index
            optDbfBackup.Visible = False
            txtDbfBackupPath.Visible = False
            optRptBackup.Visible = False
            txtRptBackupPath.Visible = False
            optPrintDetail.Visible = False
            optPrintSummary.Visible = False
            tbrJobList.Buttons("fillright").Width = 2650
            tbrJobList.Buttons("ok").Visible = True
            tbrJobList.Buttons("cancel").Visible = True
            joblistMsgColor = Message_ForeColor
            joblistMsg = "              Erase All Joblist Entries ?"
        Case Else
            ' no question
            EnableJoblistButtons
            optDbfBackup.Visible = False
            txtDbfBackupPath.Visible = False
            optRptBackup.Visible = False
            txtRptBackupPath.Visible = False
            optPrintDetail.Visible = False
            optPrintSummary.Visible = False
            tbrJobList.Buttons("fillright").Width = 3820
            tbrJobList.Buttons("ok").Visible = False
            tbrJobList.Buttons("cancel").Visible = False
'            joblistMsg = ""
    End Select
End Sub

Private Sub DisableJoblistButtonsExcept(ByVal Idx As Integer)
    If Idx <> tbrJobList.Buttons("clearjoblist").Index Then tbrJobList.Buttons("clearjoblist").Enabled = False
    If Idx <> tbrJobList.Buttons("deletejob").Index Then tbrJobList.Buttons("deletejob").Enabled = False
    If Idx <> tbrJobList.Buttons("backupfiles").Index Then tbrJobList.Buttons("backupfiles").Enabled = False
    If Idx <> tbrJobList.Buttons("previewreports").Index Then tbrJobList.Buttons("previewreports").Enabled = False
    If Idx <> tbrJobList.Buttons("generatereports").Index Then tbrJobList.Buttons("generatereports").Enabled = False
    If Idx <> tbrJobList.Buttons("printreports").Index Then tbrJobList.Buttons("printreports").Enabled = False
End Sub

Private Sub EnableJoblistButtons()
    tbrJobList.Buttons("clearjoblist").Enabled = True
    tbrJobList.Buttons("deletejob").Enabled = True
    tbrJobList.Buttons("backupfiles").Enabled = True
    tbrJobList.Buttons("previewreports").Enabled = True
    tbrJobList.Buttons("generatereports").Enabled = True
    tbrJobList.Buttons("printreports").Enabled = True
End Sub

Public Sub UpdateJoblistMsg(ByVal sMsg As String, ByVal iColor As Long)
        ' messages
        joblistMsgColor = iColor
        joblistMsg = sMsg
End Sub

Private Sub ViewJobLog()
Dim flag As Boolean
Dim JobNum As String
Dim JobStn As Integer
Dim JobSft As Integer
Dim baseFile As String
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
    
    joblistMsg = ""
    
    If adoJoblist.Recordset.BOF Then
        joblistMsgColor = MEDRED
        joblistMsg = "No Job Data Available"
    Else
    
        If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' report an error
            joblistMsgColor = MEDRED
            joblistMsg = "Invalid Job Number"
            Exit Sub
        Else
            ' read the joblist values for the designated job
            JobNum = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
            JobStn = CInt(dbgJoblist.Columns(5).CellValue(dbgJoblist.GetBookmark(0)))
            JobSft = CInt(dbgJoblist.Columns(6).CellValue(dbgJoblist.GetBookmark(0)))
            dbPath = FILEPATH_data & "C" & JobNum & AccessDbFileExt
            ' is the job still open ??
            If (StationControl(JobStn, JobSft).Job_Number = JobNum) Then
              ' report an error
              joblistMsgColor = MEDRED
              joblistMsg = "Job #" & JobNum & "  is still open"
              Exit Sub
            End If
            If Not fs.FileExists(dbPath) Then
                ' Report an error
                joblistMsgColor = MEDRED
                joblistMsg = "The Job Database File " & baseFile & " could not be located."
                Exit Sub
            End If
        End If
        
        ' view job events log
        View_JobLog JobNum, JobStn, JobSft
                                
    End If
    
End Sub


