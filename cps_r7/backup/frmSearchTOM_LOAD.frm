VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSearchTom 
   Caption         =   "BMW TOM LOAD"
   ClientHeight    =   11625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11625
   ScaleMode       =   0  'User
   ScaleWidth      =   14910.38
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisplayAllTasks 
      Caption         =   "All Tasks"
      DisabledPicture =   "frmSearchTOM_LOAD.frx":0000
      DownPicture     =   "frmSearchTOM_LOAD.frx":0C42
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   13800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSearchTOM_LOAD.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9500
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdDisplayActiveTasks 
      Caption         =   " Running Tasks "
      DisabledPicture =   "frmSearchTOM_LOAD.frx":24C6
      DownPicture     =   "frmSearchTOM_LOAD.frx":3108
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   13800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSearchTOM_LOAD.frx":3D4A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8150
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdDisplayReadyTasks 
      Caption         =   "Ready Tasks"
      DisabledPicture =   "frmSearchTOM_LOAD.frx":498C
      DownPicture     =   "frmSearchTOM_LOAD.frx":55CE
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   13800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSearchTOM_LOAD.frx":6210
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6800
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.OptionButton optWcm 
      Alignment       =   1  'Right Justify
      Caption         =   "W.C.Multiplier"
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
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Using English Units (inch, feet)"
      Top             =   9300
      Width           =   2370
   End
   Begin VB.OptionButton opt2Gm 
      Alignment       =   1  'Right Justify
      Caption         =   "2 gram Breakthrough"
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
      Left            =   3240
      TabIndex        =   0
      ToolTipText     =   "Using SI Units (mm, meter))"
      Top             =   8950
      Width           =   2370
   End
   Begin MSDataGridLib.DataGrid dgTomTasks 
      Bindings        =   "frmSearchTOM_LOAD.frx":6E52
      Height          =   6000
      Left            =   119
      TabIndex        =   2
      Top             =   0
      Width           =   18585
      _ExtentX        =   32782
      _ExtentY        =   10583
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Put TOM into CPS-R7"
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "TOM_TestOrderID"
         Caption         =   "TestOrderID"
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
         DataField       =   "TOM_VIN"
         Caption         =   "VIN"
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
         DataField       =   "TOM_RequestedStation"
         Caption         =   "RequestedStation"
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
         DataField       =   "TOM_RequestedShift"
         Caption         =   "RequestedShift"
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
      BeginProperty Column04 
         DataField       =   "CAN_WorkingCapacity"
         Caption         =   "CAN_WorkingCapacity"
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
      BeginProperty Column05 
         DataField       =   "CAN_Volume"
         Caption         =   "CAN_Volume"
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
         DataField       =   "RCP_Use2Gm"
         Caption         =   "RCP_Use2Gm"
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
         DataField       =   "RCP_UseWCM"
         Caption         =   "RCP_UseWCM"
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
      BeginProperty Column08 
         DataField       =   "TOM_TaskStatus"
         Caption         =   "TOM_TaskStatus"
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
      BeginProperty Column09 
         DataField       =   "TOM_Type"
         Caption         =   "TOM_Type"
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
      BeginProperty Column10 
         DataField       =   "TOM_Comment"
         Caption         =   "TOM_Comment"
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
      BeginProperty Column11 
         DataField       =   "TOM_Specialist"
         Caption         =   "TOM_Specialist"
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
      BeginProperty Column12 
         DataField       =   "TOM_ActualJobNumber"
         Caption         =   "TOM_ActualJobNumber"
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
      BeginProperty Column13 
         DataField       =   "TOM_ActualStation"
         Caption         =   "TOM_ActualStation"
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
      BeginProperty Column14 
         DataField       =   "TOM_ActualShift"
         Caption         =   "TOM_ActualShift"
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
      BeginProperty Column15 
         DataField       =   "TOM_OrderDate"
         Caption         =   "TOM_OrderDate"
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
      BeginProperty Column16 
         DataField       =   "TOM_ActualStartDate"
         Caption         =   "TOM_ActualStartDate"
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
      BeginProperty Column17 
         DataField       =   "TOM_ActualDoneDate"
         Caption         =   "TOM_ActualDoneDate"
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
      BeginProperty Column18 
         DataField       =   "TOM_PreviousResult"
         Caption         =   "TOM_PreviousResult"
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
      BeginProperty Column19 
         DataField       =   ""
         Caption         =   "Spare"
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
      BeginProperty Column20 
         DataField       =   ""
         Caption         =   "Spare2"
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
      BeginProperty Column21 
         DataField       =   ""
         Caption         =   "Spare3"
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
      BeginProperty Column22 
         DataField       =   ""
         Caption         =   "Spare4"
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
      BeginProperty Column23 
         DataField       =   ""
         Caption         =   "Spare5"
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
         BeginProperty Column00 
            ColumnWidth     =   1748.746
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1748.746
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1206.287
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1206.287
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1809.146
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1296.317
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1462.132
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1613.132
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2005.161
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1613.132
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1613.132
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1613.132
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1613.132
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1748.746
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2035.361
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2125.961
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   2005.161
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1326.517
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1748.746
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1341.902
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1748.746
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1326.517
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1748.746
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   3015.434
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoTomTasks 
      Height          =   375
      Left            =   7800
      Top             =   10920
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=cpsTomCanLoad"
      OLEDBString     =   "DSN=cpsTomCanLoad"
      OLEDBFile       =   ""
      DataSourceName  =   "cpsRemote"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM [TOM_CanLoadTasks] ORDER BY [tom_VIN] ASC"
      Caption         =   "TOM Tasks"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel pbxBottom 
      Height          =   4440
      Left            =   0
      TabIndex        =   3
      Top             =   6495
      Width           =   14625
      _Version        =   65536
      _ExtentX        =   25797
      _ExtentY        =   7832
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select Task"
         DisabledPicture =   "frmSearchTOM_LOAD.frx":6E6C
         DownPicture     =   "frmSearchTOM_LOAD.frx":7AAE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":86F0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   5925
      End
      Begin VB.Frame frmStatus 
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
         Height          =   4140
         Left            =   6120
         TabIndex        =   18
         Top             =   30
         Width           =   7515
         Begin VB.Label lblMessage 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "messages"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3795
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   7260
         End
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Test"
         DisabledPicture =   "frmSearchTOM_LOAD.frx":9332
         DownPicture     =   "frmSearchTOM_LOAD.frx":9F74
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":ABB6
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3420
         UseMaskColor    =   -1  'True
         Width           =   2925
      End
      Begin VB.CommandButton cmdShiftUp 
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
         Left            =   5325
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":B7F8
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Next"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.CommandButton cmdShiftDn 
         DisabledPicture =   "frmSearchTOM_LOAD.frx":BEFA
         DownPicture     =   "frmSearchTOM_LOAD.frx":C5FC
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
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":CCFE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Previous"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.CommandButton cmdStnUp 
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
         Left            =   2205
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":D400
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Next"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.CommandButton cmdStnDn 
         DisabledPicture =   "frmSearchTOM_LOAD.frx":DB02
         DownPicture     =   "frmSearchTOM_LOAD.frx":E204
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
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":E906
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Previous"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.Timer tmrScreen 
         Interval        =   250
         Left            =   5880
         Top             =   -120
      End
      Begin VB.Frame frmRcp 
         Caption         =   "Recipe Selection"
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
         Height          =   1065
         Left            =   3120
         TabIndex        =   12
         Top             =   2160
         Width           =   2940
      End
      Begin VB.Frame frmCan 
         Caption         =   "Canister Parameters"
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
         Height          =   1065
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   2940
         Begin VB.Label lblCanWC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Working Capacity"
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
            Left            =   0
            TabIndex        =   11
            Top             =   360
            Width           =   1605
         End
         Begin VB.Label lblCanWC_Units 
            BackStyle       =   0  'Transparent
            Caption         =   "gm"
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
            Left            =   2520
            TabIndex        =   10
            Top             =   360
            Width           =   315
         End
         Begin VB.Label lblCanVol_Units 
            BackStyle       =   0  'Transparent
            Caption         =   "L"
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
            Left            =   2520
            TabIndex        =   9
            Top             =   705
            Width           =   315
         End
         Begin VB.Label lblCanVol 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Volume"
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
            Left            =   0
            TabIndex        =   8
            Top             =   705
            Width           =   1605
         End
         Begin VB.Label lblCanWorkingCapacity 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "888"
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
            Left            =   1680
            TabIndex        =   7
            Top             =   330
            Width           =   735
         End
         Begin VB.Label lblCanVolume 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "888"
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
            Left            =   1680
            TabIndex        =   6
            Top             =   675
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Save Canister && Recipe to Station"
         DisabledPicture =   "frmSearchTOM_LOAD.frx":F008
         DownPicture     =   "frmSearchTOM_LOAD.frx":FC4A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchTOM_LOAD.frx":1088C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3420
         UseMaskColor    =   -1  'True
         Width           =   2925
      End
      Begin Threed.SSPanel pnlStn 
         Height          =   600
         Left            =   600
         TabIndex        =   21
         ToolTipText     =   "Station selected for TOM Task"
         Top             =   1200
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
      Begin Threed.SSPanel pnlShift 
         Height          =   600
         Left            =   3720
         TabIndex        =   22
         ToolTipText     =   "Shift  selected for TOM Task"
         Top             =   1200
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   1058
         _StockProps     =   15
         Caption         =   "Shift 1"
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
End
Attribute VB_Name = "frmSearchTom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 859'''''''''' Form SearchTom.frm ''''''''''''''''''''
Option Explicit
Dim exitFlag As Boolean
Dim exitCntr As Integer
Dim sPath As String
Dim curStatusCrit As String
Dim rsCrit As String
Dim RowHgt As Single
Dim InitRow As Integer

Private Sub cmdDisplayActiveTasks_Click()
    curStatusCrit = "Active"
    DisplayData 2
    Setup_DisplayTasks
End Sub

Private Sub cmdDisplayAllTasks_Click()
    curStatusCrit = "All"
    DisplayData 2
    Setup_DisplayTasks
End Sub

Private Sub cmdDisplayReadyTasks_Click()
    curStatusCrit = "Ready"
    DisplayData 2
    Setup_SelectTask
End Sub

Private Sub cmdLoad_Click()
Dim flag As Boolean

    ' save VIN to the station/shift
    JobInfo(CurTomTask.ActualStation, CurTomTask.ActualShift).Vehicle = CurTomTask.VIN
    ' save canister to the station/shift
    StationCanister(CurTomTask.ActualStation, CurTomTask.ActualShift).Description = "Canister from T.O.M."
    StationCanister(CurTomTask.ActualStation, CurTomTask.ActualShift).Number = CInt(0)
    StationCanister(CurTomTask.ActualStation, CurTomTask.ActualShift).Validated = True
    StationCanister(CurTomTask.ActualStation, CurTomTask.ActualShift).WorkingVolume = CurTomTask.CanVolume
    StationCanister(CurTomTask.ActualStation, CurTomTask.ActualShift).WorkingCapacity = CurTomTask.CanWC
    ' save station canister recipes
    Save_StationCanisters
    ' clear TOM Data for this Station/Shift
    StnTomTask(CurTomTask.ActualStation, CurTomTask.ActualShift) = TomData_Clear
    ' save recipe to the station/shift
    flag = ValidTomRecipe(CurTomTask.rcpRecipeNumber)
    If flag Then
        ' recipe validation was successful
        lblMessage.Caption = vbCrLf & "Canister and Recipe Validated for Station #" + Format(CurTomTask.ActualStation, "0") & " Shift #" + Format(CurTomTask.ActualShift, "0")
        lblMessage.Caption = lblMessage.Caption & vbCrLf
        ' copy new TOM Data to the Station/Shift
        StnTomTask(CurTomTask.ActualStation, CurTomTask.ActualShift) = CurTomTask
        ' close recipe screen
        Unload frmRecipe
        ' open station detail screen
        frmStnDetail.Refresh
        ' briefly show SearchTom screen
        frmSearchTom.Show
        'frmSearchTom.Show
        ' setup SearchTom screen as Ready-To-Start
        Setup_ReadyToStart
    Else
        ' recipe validation failed
        lblMessage.Caption = vbCrLf & "Recipe Validation Failed for Station #" + Format(CurTomTask.ActualStation, "0") & " Shift #" + Format(CurTomTask.ActualShift, "0")
        lblMessage.Caption = lblMessage.Caption & vbCrLf
    End If
End Sub

Private Sub cmdSelect_Click()
Dim errorFlag As Boolean
Dim newStn As String
Dim newShift As String
Dim newTaskID As String
Dim newVIN As String
Dim newType As String
Dim newWC As String
Dim newVol As String
Dim newUse2Gm As String
Dim newUseWCM As String
Dim newStatus As String
Dim newPreviousResult As String
    lblMessage.Caption = ""
    errorFlag = False
    newTaskID = dgTomTasks.Columns(0).CellValue(dgTomTasks.GetBookmark(0))
    newVIN = dgTomTasks.Columns(1).CellValue(dgTomTasks.GetBookmark(0))
    newStn = dgTomTasks.Columns(2).CellValue(dgTomTasks.GetBookmark(0))
    newShift = dgTomTasks.Columns(3).CellValue(dgTomTasks.GetBookmark(0))
    newWC = dgTomTasks.Columns(4).CellValue(dgTomTasks.GetBookmark(0))
    newVol = dgTomTasks.Columns(5).CellValue(dgTomTasks.GetBookmark(0))
    newUse2Gm = dgTomTasks.Columns(6).CellValue(dgTomTasks.GetBookmark(0))
    newUseWCM = dgTomTasks.Columns(7).CellValue(dgTomTasks.GetBookmark(0))
    newStatus = dgTomTasks.Columns(8).CellValue(dgTomTasks.GetBookmark(0))
    newType = dgTomTasks.Columns(9).CellValue(dgTomTasks.GetBookmark(0))
    newPreviousResult = dgTomTasks.Columns(18).CellValue(dgTomTasks.GetBookmark(0))
    newTaskID = Trim(newTaskID)
    newVIN = Trim(newVIN)
    ' check requested station number
    If (Not IsNumeric(newStn)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Station is Not a Number; set to 1" & vbCrLf
        newStn = "1"
    End If
    If ((CInt(newStn) < 1) Or (CInt(newStn) > LAST_STN)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Station is Out-Of-Range; set to 1" & vbCrLf
        newStn = "1"
    End If
    ' check requested shift number
    If (Not IsNumeric(newShift)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Shift is Not a Number; set to 1" & vbCrLf
        newShift = "1"
    End If
    If ((CInt(newShift) < 1) Or (CInt(newShift) > NR_SHIFT)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Shift is Out-Of-Range; set to 1" & vbCrLf
        newShift = "1"
    End If
    ' check Vehicle Identification Number
    If ((Len(newVIN) < 1) Or (Len(newVIN) > 128)) Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Invalid VIN" & vbCrLf
    End If
    ' check Canister Working Capacity
    If (Not IsNumeric(newWC)) Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Canister Working Capacity is Not a Number" & vbCrLf
    ElseIf CSng(newWC) < 1 Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Canister Working Capacity is too small (< 1 grams)" & vbCrLf
    ElseIf CSng(newWC) > 1000 Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Canister Working Capacity is too large (> 1000 grams)" & vbCrLf
    End If
    ' check Canister Volume
    If (Not IsNumeric(newVol)) Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Canister Volume is Not a Number" & vbCrLf
    ElseIf CSng(newVol) < 0.005 Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Canister Volume is too small (< 0.005 Liters)" & vbCrLf
    ElseIf CSng(newVol) > 10 Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Canister Volume is too large (> 10 Liters)" & vbCrLf
    End If
    ' check Recipe Selection
    If (newUse2Gm = "True") And (newUseWCM = "True") Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Recipe Selection must be unique" & vbCrLf & "Both are currently selected" & vbCrLf
    End If
    If (newUse2Gm = "False") And (newUseWCM = "False") Then
        errorFlag = True
        lblMessage.Caption = lblMessage.Caption & "Recipe Selection must be unique" & vbCrLf & "Neither choice is currently selected" & vbCrLf
    End If
    ' No errors; proceed
    If (Not errorFlag) Then
        CurTomTask.RequestedStation = CInt(newStn)
        CurTomTask.RequestedShift = CInt(newShift)
        CurTomTask.TaskID = newTaskID
        CurTomTask.VIN = newVIN
        CurTomTask.TaskType = newType
        CurTomTask.TaskStatus = newStatus
        CurTomTask.PreviousResult = newPreviousResult
        CurTomTask.CanWC = CSng(newWC)
        CurTomTask.CanVolume = CSng(newVol)
        CurTomTask.rcpUse2Gm = IIf(newUse2Gm = "True", True, False)
        CurTomTask.rcpUseWCM = IIf(newUseWCM = "True", True, False)
        CurTomTask.rcpRecipeNumber = IIf(CurTomTask.rcpUse2Gm, TOM_2Gm_Recipe, TOM_Wcm_Recipe)
        pnlStn.Caption = "Station " & Format(CurTomTask.RequestedStation, "0")
        pnlShift.Caption = "Shift " & Format(CurTomTask.RequestedShift, "0")
        lblCanWorkingCapacity.Caption = Format(CurTomTask.CanWC, "##0.0")
        lblCanVolume.Caption = Format(CurTomTask.CanVolume, "##0.0")
        opt2Gm.Value = CurTomTask.rcpUse2Gm
        optWcm.Value = CurTomTask.rcpUseWCM
        ' set the actual station/shift
        CurTomTask.ActualStation = CurTomTask.RequestedStation
        CurTomTask.ActualShift = CurTomTask.RequestedShift
    
        CheckForIdle
    End If
End Sub

Private Sub cmdStart_Click()
    If (NR_SHIFT > 1) Then
        lblMessage.Caption = vbCrLf & "Starting Test on Station " & Format(CurTomTask.ActualStation, "0") & " Shift " & Format(CurTomTask.ActualShift, "0") & vbCrLf & vbCrLf
    Else
        lblMessage.Caption = vbCrLf & "Starting Test on Station " & Format(CurTomTask.ActualStation, "0") & vbCrLf & vbCrLf
    End If
    exitFlag = True
End Sub

Private Sub cmdShiftDn_Click()
    CurTomTask.ActualShift = IIf(CurTomTask.ActualShift <= 1, NR_SHIFT, CurTomTask.ActualShift - 1)
    pnlShift.Caption = "Shift " & Format(CurTomTask.ActualShift, "0")
    CheckForIdle
End Sub

Private Sub cmdShiftUp_Click()
    CurTomTask.ActualShift = IIf(CurTomTask.ActualShift = NR_SHIFT, 1, CurTomTask.ActualShift + 1)
    pnlShift.Caption = "Shift " & Format(CurTomTask.ActualShift, "0")
    CheckForIdle
End Sub

Private Sub cmdStnDn_Click()
    CurTomTask.ActualStation = IIf(CurTomTask.ActualStation <= 1, LAST_STN, CurTomTask.ActualStation - 1)
    pnlStn.Caption = "Station " & Format(CurTomTask.ActualStation, "0")
    CheckForIdle
End Sub

Private Sub cmdStnUp_Click()
    CurTomTask.ActualStation = IIf(CurTomTask.ActualStation >= LAST_STN, 1, CurTomTask.ActualStation + 1)
    pnlStn.Caption = "Station " & Format(CurTomTask.ActualStation, "0")
    CheckForIdle
End Sub

Private Sub dgTomTasks_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex + 1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me

      Set frmSearchTom = Nothing
    End If
End Sub

' TESTING
'Private Sub Form_Load()
'    SetErrModule 859, 2
'    KeyPreview = True
'    exitFlag = False
'    curStatusCrit = "Ready"
'
'    lblMessage.ForeColor = Message_ForeColor
'
'    dgTomTasks.AllowRowSizing = False
'
'    DisplayData 2
'    Setup_SelectTask
'
'    ResetErrModule
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Testing
    Unload frmSearchTom
    Set frmSearchTom = Nothing
'    Unload frmTomLoad
'    Set frmTomLoad = Nothing
End Sub

Private Sub DisplayData(sortCol As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 859, 888

Dim statusDesc As String
   ' get Description for desired status criteria
   ' and setup beginning of criteria string
     Select Case curStatusCrit
        Case "Ready"
            statusDesc = "Ready-To-Run"
            rsCrit = "SELECT * FROM [TOM_CanLoadTasks] WHERE [TOM_TaskStatus] = '" & curStatusCrit & "'"
        Case "Active"
            statusDesc = "Active"
            rsCrit = "SELECT * FROM [TOM_CanLoadTasks] WHERE [TOM_TaskStatus] = '" & curStatusCrit & "'"
        Case "Done"
            statusDesc = "Completed"
            rsCrit = "SELECT * FROM [TOM_CanLoadTasks] WHERE [TOM_TaskStatus] = '" & curStatusCrit & "'"
        Case "All"
            statusDesc = " "
            rsCrit = "SELECT * FROM [TOM_CanLoadTasks]"
        Case Else
            statusDesc = "Ready-To-Run"
            rsCrit = "SELECT * FROM [TOM_CanLoadTasks] WHERE [TOM_TaskStatus] = '" & curStatusCrit & "'"
    End Select
   ' Select & Sort
    Select Case sortCol
        Case 1
            rsCrit = rsCrit & " ORDER BY [TOM_TestOrderID] ASC"
        Case 2
            rsCrit = rsCrit & " ORDER BY [TOM_VIN] ASC"
        Case 3
            rsCrit = rsCrit & " ORDER BY [TOM_RequestedStation] ASC"
        Case 4
            rsCrit = rsCrit & " ORDER BY [TOM_RequestedShift] ASC"
        Case Else
            rsCrit = rsCrit & " ORDER BY [TOM_VIN] ASC"
    End Select
    adoTomTasks.RecordSource = rsCrit
    adoTomTasks.Refresh

    ' Display number of TomTasks found
    If Not adoTomTasks.Recordset.BOF Then
        adoTomTasks.Recordset.GetRows
        Select Case adoTomTasks.Recordset.RecordCount
            Case 0
                dgTomTasks.Caption = " No " & statusDesc & " TomTasks"
            Case 1
                dgTomTasks.Caption = Format(adoTomTasks.Recordset.RecordCount, "###0") & " Only " & statusDesc & " TomTask"
            Case Else
                dgTomTasks.Caption = Format(adoTomTasks.Recordset.RecordCount, "###0") & " " & statusDesc & " TomTasks"
        End Select
        
        ' Set column properties
        dgTomTasks.Columns(0).Width = 1500
        dgTomTasks.Columns(1).Width = 1800
        dgTomTasks.Columns(2).Width = 1300
        dgTomTasks.Columns(3).Width = 1100
        dgTomTasks.Columns(4).Width = 2200
        dgTomTasks.Columns(5).Width = 1450
        dgTomTasks.Columns(6).Width = 1550
        dgTomTasks.Columns(7).Width = 1550
        dgTomTasks.Columns(8).Width = 1700
        
        ' move pointer to first row
        adoTomTasks.Recordset.MoveFirst
    Else
        dgTomTasks.Caption = " No " & statusDesc & " TomTasks"
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


Public Sub SetInitialRow(rownum As Integer)
    InitRow = rownum
End Sub

Private Sub Setup_DisplayTasks()
    cmdSelect.Enabled = False
    cmdStnDn.Enabled = False
    cmdStnUp.Enabled = False
    pnlStn.Enabled = False
    cmdShiftDn.Enabled = False
    cmdShiftUp.Enabled = False
    pnlShift.Enabled = False
    frmCan.Enabled = False
    lblCanWC.Enabled = False
    lblCanWorkingCapacity.Enabled = False
    lblCanWC_Units.Enabled = False
    lblCanVol.Enabled = False
    lblCanVolume.Enabled = False
    lblCanVol_Units.Enabled = False
    frmRcp.Enabled = False
    opt2Gm.Enabled = False
    optWcm.Enabled = False
    cmdLoad.Enabled = False
    cmdStart.Enabled = False
    lblMessage.Caption = vbCrLf & "View Tasks"
End Sub

Private Sub Setup_SelectTask()
    cmdSelect.Enabled = True
    cmdStnDn.Enabled = False
    cmdStnUp.Enabled = False
    pnlStn.Enabled = False
    cmdShiftDn.Enabled = False
    cmdShiftUp.Enabled = False
    pnlShift.Enabled = False
    frmCan.Enabled = False
    lblCanWC.Enabled = False
    lblCanWorkingCapacity.Enabled = False
    lblCanWC_Units.Enabled = False
    lblCanVol.Enabled = False
    lblCanVolume.Enabled = False
    lblCanVol_Units.Enabled = False
    frmRcp.Enabled = False
    opt2Gm.Enabled = False
    optWcm.Enabled = False
    cmdLoad.Enabled = False
    cmdStart.Enabled = False
    lblMessage.Caption = vbCrLf & "Select a Task"
End Sub

Private Sub Setup_ChangeStnShift()
    cmdSelect.Enabled = True
    cmdStnDn.Enabled = True
    cmdStnUp.Enabled = True
    pnlStn.Enabled = True
    If NR_SHIFT > 1 Then
        cmdShiftDn.Enabled = True
        cmdShiftUp.Enabled = True
        pnlShift.Enabled = True
    Else
        cmdShiftDn.Enabled = False
        cmdShiftUp.Enabled = False
        pnlShift.Enabled = False
    End If
    frmCan.Enabled = True
    lblCanWC.Enabled = True
    lblCanWorkingCapacity.Enabled = True
    lblCanWC_Units.Enabled = True
    lblCanVol.Enabled = True
    lblCanVolume.Enabled = True
    lblCanVol_Units.Enabled = True
    frmRcp.Enabled = True
    opt2Gm.Enabled = True
    optWcm.Enabled = True
    cmdLoad.Enabled = False
    cmdStart.Enabled = False
    lblMessage.Caption = lblMessage.Caption & "Change the Station or Shift" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Select a Different Task" & vbCrLf
End Sub

Private Sub Setup_ReadyToLoad()
    cmdSelect.Enabled = True
    cmdStnDn.Enabled = True
    cmdStnUp.Enabled = True
    pnlStn.Enabled = True
    If NR_SHIFT > 1 Then
        cmdShiftDn.Enabled = True
        cmdShiftUp.Enabled = True
        pnlShift.Enabled = True
    Else
        cmdShiftDn.Enabled = False
        cmdShiftUp.Enabled = False
        pnlShift.Enabled = False
    End If
    frmCan.Enabled = True
    lblCanWC.Enabled = True
    lblCanWorkingCapacity.Enabled = True
    lblCanWC_Units.Enabled = True
    lblCanVol.Enabled = True
    lblCanVolume.Enabled = True
    lblCanVol_Units.Enabled = True
    frmRcp.Enabled = True
    opt2Gm.Enabled = True
    optWcm.Enabled = True
    cmdLoad.Enabled = True
    cmdStart.Enabled = False
    lblMessage.Caption = lblMessage.Caption & "Press Save to Load the Canister and Recipe to the station" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Change the Canister or Recipe Values" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Change the Station or Shift" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Select a Different Task" & vbCrLf
End Sub

Private Sub Setup_ReadyToStart()
    cmdSelect.Enabled = True
    cmdStnDn.Enabled = True
    cmdStnUp.Enabled = True
    pnlStn.Enabled = True
    If NR_SHIFT > 1 Then
        cmdShiftDn.Enabled = True
        cmdShiftUp.Enabled = True
        pnlShift.Enabled = True
    Else
        cmdShiftDn.Enabled = False
        cmdShiftUp.Enabled = False
        pnlShift.Enabled = False
    End If
    frmCan.Enabled = True
    lblCanWC.Enabled = True
    lblCanWorkingCapacity.Enabled = True
    lblCanWC_Units.Enabled = True
    lblCanVol.Enabled = True
    lblCanVolume.Enabled = True
    lblCanVol_Units.Enabled = True
    frmRcp.Enabled = True
    opt2Gm.Enabled = True
    optWcm.Enabled = True
    cmdLoad.Enabled = True
    cmdStart.Enabled = True
    lblMessage.Caption = vbCrLf & "Requested Station/Shift is Ready to Start" & vbCrLf & vbCrLf & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Press Start to start the test" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Change the Canister or Recipe Values" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Change the Station or Shift" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Select a Different Task" & vbCrLf
End Sub

Private Sub CheckForIdle()
    If StationControl(CurTomTask.ActualStation, CurTomTask.ActualShift).Mode = VBIDLE Then
        lblMessage.Caption = vbCrLf & "Selected Station/Shift is Ready for Canister and Recipe Values" & vbCrLf & vbCrLf & vbCrLf
        Setup_ReadyToLoad
    Else
        lblMessage.Caption = vbCrLf & "Selected Station/Shift is not Idle" & vbCrLf & vbCrLf & vbCrLf
        Setup_ChangeStnShift
    End If
End Sub

Private Sub tmrScreen_Timer()
    If exitFlag Then
        exitCntr = exitCntr + 1
        If (exitCntr > 5) Then
            frmStnDetail.RemoteStnStart CurTomTask.ActualStation, CurTomTask.ActualShift
            frmStnDetail.Show
' Testing
            Unload frmSearchTom
            Set frmSearchTom = Nothing
'            Unload frmTomLoad
'            Set frmTomLoad = Nothing

        End If
    Else
        exitCntr = 0
    End If
End Sub



