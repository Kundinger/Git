VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSearchRemote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Tasks"
   ClientHeight    =   10395
   ClientLeft      =   195
   ClientTop       =   600
   ClientWidth     =   14880
   Icon            =   "frmSearchTOM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10395
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
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
      Left            =   16080
      TabIndex        =   25
      ToolTipText     =   "Using SI Units (mm, meter))"
      Top             =   240
      Width           =   2370
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
      Left            =   16080
      TabIndex        =   24
      ToolTipText     =   "Using English Units (inch, feet)"
      Top             =   585
      Width           =   2370
   End
   Begin MSDataGridLib.DataGrid dgRemoteTasks 
      Bindings        =   "frmSearchTOM.frx":57E2
      Height          =   5935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   10478
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Remote Tasks that are Ready To Run"
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "REM_TaskID"
         Caption         =   "TaskID"
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
         DataField       =   "REM_VIN"
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
         DataField       =   "REM_RequestedStation"
         Caption         =   "Req Station"
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
         DataField       =   "REM_RequestedShift"
         Caption         =   "Req Shift"
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
         DataField       =   "CAN_Description"
         Caption         =   "CAN_Description"
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
         DataField       =   "CAN_WorkCap"
         Caption         =   "CAN_WorkCap"
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
         DataField       =   "REM_TaskStatus"
         Caption         =   "TaskStatus"
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
         DataField       =   "REM_Comment"
         Caption         =   "Comment"
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
         DataField       =   "REM_ActualJobNumber"
         Caption         =   "JobNumber"
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
         DataField       =   "REM_ActualStation"
         Caption         =   "ActualStation"
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
         DataField       =   "REM_ActualShift"
         Caption         =   "ActualShift"
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
         DataField       =   "REM_InhibitChgs"
         Caption         =   "InhibitChgs"
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
         DataField       =   "REM_OrderDate"
         Caption         =   "OrderDate"
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
         DataField       =   "REM_ActualStartDate"
         Caption         =   "StartDate"
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
         DataField       =   "REM_ActualDoneDate"
         Caption         =   "DoneDate"
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
         DataField       =   "REM_PreviousResult"
         Caption         =   "PreviousResult"
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
         DataField       =   "RCP_Number"
         Caption         =   "RCP_Number"
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
         DataField       =   "RCP_Name"
         Caption         =   "RCP_Name"
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
         DataField       =   "PRG_Number"
         Caption         =   "PRG_Number"
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
         DataField       =   "PRG_Name"
         Caption         =   "PRG_Name"
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
         DataField       =   "SEQ_Number"
         Caption         =   "SEQ_Number"
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
         DataField       =   "SEQ_Name"
         Caption         =   "SEQ_Name"
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
         DataField       =   "AVL_FileRoot"
         Caption         =   "AVL_FileRoot"
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
            ColumnWidth     =   2805.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   2115.213
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoRemoteTasks 
      Height          =   375
      Left            =   6480
      Top             =   6000
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=cpsRemote"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "cpsRemote"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM [RemoteTasks] ORDER BY [REM_TaskID] ASC"
      Caption         =   "Remote Task Orders"
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
      Height          =   4305
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   14625
      _Version        =   65536
      _ExtentX        =   25797
      _ExtentY        =   7594
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
      Begin VB.CommandButton cmdDisplayReadyTasks 
         Caption         =   "Ready Tasks"
         DisabledPicture =   "frmSearchTOM.frx":57FF
         DownPicture     =   "frmSearchTOM.frx":6441
         BeginProperty Font 
            Name            =   "Arial"
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
         Picture         =   "frmSearchTOM.frx":7083
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDisplayAllTasks 
         Caption         =   "All Tasks"
         DisabledPicture =   "frmSearchTOM.frx":7CC5
         DownPicture     =   "frmSearchTOM.frx":8907
         BeginProperty Font 
            Name            =   "Arial"
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
         Picture         =   "frmSearchTOM.frx":9549
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3060
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDisplayActiveTasks 
         Caption         =   " Running Tasks "
         DisabledPicture =   "frmSearchTOM.frx":A18B
         DownPicture     =   "frmSearchTOM.frx":ADCD
         BeginProperty Font 
            Name            =   "Arial"
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
         Picture         =   "frmSearchTOM.frx":BA0F
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1590
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Save Canister && Recipe to Station"
         DisabledPicture =   "frmSearchTOM.frx":C651
         DownPicture     =   "frmSearchTOM.frx":D293
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
         Picture         =   "frmSearchTOM.frx":DED5
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   2925
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
         Left            =   0
         TabIndex        =   13
         Top             =   2040
         Width           =   2940
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
            TabIndex        =   19
            Top             =   675
            Width           =   735
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
            TabIndex        =   18
            Top             =   330
            Width           =   735
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
            TabIndex        =   17
            Top             =   705
            Width           =   1605
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
            TabIndex        =   16
            Top             =   705
            Width           =   315
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
            TabIndex        =   15
            Top             =   360
            Width           =   315
         End
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
            TabIndex        =   14
            Top             =   360
            Width           =   1605
         End
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
         Left            =   3000
         TabIndex        =   12
         Top             =   2040
         Width           =   2940
         Begin VB.Label lblRcpDesc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "description"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   245
            Left            =   90
            TabIndex        =   28
            Top             =   700
            Width           =   2745
         End
         Begin VB.Label lblNumber 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Number"
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
            Left            =   600
            TabIndex        =   27
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblRcpNum 
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
            Left            =   1440
            TabIndex        =   26
            Top             =   330
            Width           =   735
         End
      End
      Begin VB.Timer tmrScreen 
         Interval        =   250
         Left            =   5880
         Top             =   -120
      End
      Begin VB.CommandButton cmdStnDn 
         DisabledPicture =   "frmSearchTOM.frx":EB17
         DownPicture     =   "frmSearchTOM.frx":F219
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
         Picture         =   "frmSearchTOM.frx":F91B
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmSearchTOM.frx":1001D
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Next"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.CommandButton cmdShiftDn 
         DisabledPicture =   "frmSearchTOM.frx":1071F
         DownPicture     =   "frmSearchTOM.frx":10E21
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
         Picture         =   "frmSearchTOM.frx":11523
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Previous"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
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
         Picture         =   "frmSearchTOM.frx":11C25
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Next"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Test"
         DisabledPicture =   "frmSearchTOM.frx":12327
         DownPicture     =   "frmSearchTOM.frx":12F69
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
         Picture         =   "frmSearchTOM.frx":13BAB
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3300
         UseMaskColor    =   -1  'True
         Width           =   2925
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
         TabIndex        =   3
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
            TabIndex        =   4
            Top             =   240
            Width           =   7260
         End
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select Task"
         DisabledPicture =   "frmSearchTOM.frx":147ED
         DownPicture     =   "frmSearchTOM.frx":1542F
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
         Picture         =   "frmSearchTOM.frx":16071
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   5925
      End
      Begin Threed.SSPanel pnlStn 
         Height          =   600
         Left            =   600
         TabIndex        =   5
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
         TabIndex        =   6
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
Attribute VB_Name = "frmSearchRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 858'''''''''' Form SearchRemote.frm '''''''''''''''''''
Option Explicit
Dim exitFlag As Boolean
Dim exitCntr As Integer
Dim sPath As String
Dim curStatusCrit As String
Dim rsCrit As String
Dim RowHgt As Single
Dim InitRow As Integer
Dim canFlag As Boolean

Public Sub NewMsg(ByVal sTxt As String)
    lblMessage.Caption = vbCrLf & sTxt
End Sub

Private Sub adoRemoteTasks_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub cmdDisplayActiveTasks_Click()
    curStatusCrit = "InProcess"
    DisplayData 3
    Setup_DisplayTasks
End Sub

Private Sub cmdDisplayAllTasks_Click()
    curStatusCrit = "All"
    DisplayData 1
    Setup_DisplayTasks
End Sub

Private Sub cmdDisplayReadyTasks_Click()
    If USINGREMAVLFILES Then
        AVL_TaskFiles_Check
        adoRemoteTasks.Refresh
        dgRemoteTasks.Refresh
        canFlag = False
    Else
        canFlag = True
    End If
    curStatusCrit = "Ready"
    DisplayData 14
    Setup_SelectTask
End Sub

Private Sub cmdLoad_Click()
Dim flag As Boolean
Dim iStation As Integer
Dim iShift As Integer


    iStation = CurRemoteTask.ActualStation
    iShift = CurRemoteTask.ActualShift
    
    ' save VIN to the station/shift
    JobInfo(iStation, iShift).Vehicle = CurRemoteTask.VIN
'    JobInfo(iStation, iShift).Comment = JobInfo(iStation, iShift).Comment & vbCrLf & "  REM Task " & CurRemoteTask.TaskID
'    JobInfo(iStation, iShift).Engineer = CurRemoteTask.Specialist
    ' save canister to the station/shift
    StationCanister(iStation, iShift).Description = "Canister from Host"
    StationCanister(iStation, iShift).Number = CInt(0)
    StationCanister(iStation, iShift).Validated = True
    StationCanister(iStation, iShift).WorkingVolume = CurRemoteTask.Can.WorkingVolume
    StationCanister(iStation, iShift).WorkingCapacity = CurRemoteTask.Can.WorkingCapacity
    ' save station canister recipes
    Save_StationCanisters
    ' clear REMOTE Data for this Station/Shift
    RemData_Clear StnRemoteTask(iStation, iShift)
    ' save recipe to the station/shift
    flag = ValidRemRecipe(CurRemoteTask.Rcp.Number)
    If flag Then
        ' recipe validation was successful
        lblMessage.Caption = vbCrLf & "Canister and Recipe Validated for Station #" + Format(iStation, "0") & " Shift #" + Format(iShift, "0")
        lblMessage.Caption = lblMessage.Caption & vbCrLf
        CurRemoteTask.Rcp = StationRecipe(DispStn, DispShift)
        lblRcpDesc.Caption = CurRemoteTask.Rcp.Name
        ' copy new Task Data to the Station/Shift
        StnRemoteTask(iStation, iShift) = CurRemoteTask
        ' close recipe screen
        Unload frmRecipe
        ' set Job Sequence to 1 Course of "RunStationRecipe"
        StationSequence(iStation, iShift).Description = "Remote Task " & StnRemoteTask(iStation, iShift).TaskID
        StationSequence(iStation, iShift).NumCourses = 1                    ' one course
        StationSequence(iStation, iShift).CourseData(1).Type = 3            ' run a recipe
        StationSequence(iStation, iShift).CourseData(1).RecipeNumber = 0    ' run current station recipe (as modifed for this Remote Task)
        StationSequence(iStation, iShift).Number = 0
        StationSequence(iStation, iShift).AuxScaleNo = StationRecipe(iStation, iShift).AuxScaleNo
        StationSequence(iStation, iShift).PriScaleNo = StationRecipe(iStation, iShift).PriScaleNo
        StationSequence(iStation, iShift).Validated = True
        ' open station detail screen
        frmStnDetail.Refresh
        ' briefly show SearchRemote screen
        If USINGREMCANLOAD Then
            frmSearchRemote.Show
        End If
        If USINGTOMCANLOAD Then
            frmSearchTom.Show
        End If

        ' setup SearchRemote screen as Ready-To-Start
        Setup_ReadyToStart
    Else
        ' recipe validation failed
        lblMessage.Caption = vbCrLf & "Recipe Validation Failed for Station #" + Format(iStation, "0") & " Shift #" + Format(iShift, "0")
        lblMessage.Caption = lblMessage.Caption & vbCrLf
    End If
End Sub

Private Sub cmdSelect_Click()
Dim errorFlag As Boolean
Dim chgFlag As Boolean
Dim newInhibit As Boolean
Dim newStn As String
Dim newShift As String
Dim newTaskID As String
Dim newRcpNum As String
Dim newVIN As String
Dim newRoot As String
Dim newWC As String
Dim newVol As String
Dim newStatus As String
Dim newPreviousResult As String
    lblMessage.Caption = ""
    errorFlag = False
    newTaskID = dgRemoteTasks.Columns(0).CellValue(dgRemoteTasks.GetBookmark(0))
    newVIN = dgRemoteTasks.Columns(1).CellValue(dgRemoteTasks.GetBookmark(0))
    newStn = dgRemoteTasks.Columns(2).CellValue(dgRemoteTasks.GetBookmark(0))
    newShift = dgRemoteTasks.Columns(3).CellValue(dgRemoteTasks.GetBookmark(0))
    newInhibit = dgRemoteTasks.Columns(12).CellValue(dgRemoteTasks.GetBookmark(0))
    newRcpNum = dgRemoteTasks.Columns(17).CellValue(dgRemoteTasks.GetBookmark(0))
    newVol = dgRemoteTasks.Columns(5).CellValue(dgRemoteTasks.GetBookmark(0))
    newWC = dgRemoteTasks.Columns(6).CellValue(dgRemoteTasks.GetBookmark(0))
    newStatus = dgRemoteTasks.Columns(7).CellValue(dgRemoteTasks.GetBookmark(0))
    newRoot = dgRemoteTasks.Columns(23).CellValue(dgRemoteTasks.GetBookmark(0))
    If (Not IsNull(dgRemoteTasks.Columns(16).CellValue(dgRemoteTasks.GetBookmark(0)))) Then newPreviousResult = dgRemoteTasks.Columns(16).CellValue(dgRemoteTasks.GetBookmark(0))
    newTaskID = Trim(newTaskID)
    newVIN = Trim(newVIN)
    ' check requested recipe number
    If (Not IsNumeric(newRcpNum)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Recipe # is Not a Number; set to 1" & vbCrLf
        newRcpNum = "1"
    End If
    If ((CInt(newRcpNum) < 1) Or (CInt(newRcpNum) > NR_RCP)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Recipe is Out-Of-Range; set to 1" & vbCrLf
        newRcpNum = "1"
    End If
    ' check requested station number
    If (Not IsNumeric(newStn)) Then
        lblMessage.Caption = lblMessage.Caption & "Requested Station is Not a Number; set to 1" & vbCrLf
        newStn = "1"
    End If
    If ((CInt(newStn) < 1) Or (CInt(newStn) > NR_STN)) Then
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
    ' No errors; proceed
    If (Not errorFlag) Then
        If (REMCHGSENABLED And (Not CurRemoteTask.InhibitChanges)) Then
            chgFlag = True
        Else
            chgFlag = False
        End If
        CurRemoteTask.RequestedStation = CInt(newStn)
        CurRemoteTask.RequestedShift = CInt(newShift)
        CurRemoteTask.TaskID = newTaskID
        CurRemoteTask.VIN = newVIN
        CurRemoteTask.AVL_FileRoot = newRoot
        CurRemoteTask.TaskStatus = newStatus
        CurRemoteTask.PreviousResult = newPreviousResult
        CurRemoteTask.Can.WorkingCapacity = CSng(newWC)
        CurRemoteTask.Can.WorkingVolume = CSng(newVol)
        CurRemoteTask.InhibitChanges = newInhibit
        CurRemoteTask.Rcp.Number = newRcpNum
        pnlStn.Caption = "Station " & Format(CurRemoteTask.RequestedStation, "0")
        pnlShift.Caption = "Shift " & Format(CurRemoteTask.RequestedShift, "0")
        lblCanWorkingCapacity.Caption = Format(CurRemoteTask.Can.WorkingCapacity, "##0.0")
        lblCanVolume.Caption = Format(CurRemoteTask.Can.WorkingVolume, "##0.0##")
        lblRcpNum.Caption = Format(CurRemoteTask.Rcp.Number, "##0")
        lblCanWorkingCapacity.Enabled = IIf((canFlag And chgFlag), True, False)
        lblCanVolume.Enabled = IIf((canFlag And chgFlag), True, False)
        lblRcpNum.Enabled = IIf((canFlag And chgFlag), True, False)
        lblRcpDesc.Caption = ""
        ' set the actual station/shift
        CurRemoteTask.ActualStation = CurRemoteTask.RequestedStation
        CurRemoteTask.ActualShift = CurRemoteTask.RequestedShift
        cmdStnDn.Enabled = chgFlag
        cmdStnUp.Enabled = chgFlag
        pnlStn.Enabled = chgFlag
        cmdShiftDn.Enabled = IIf((chgFlag And (NR_SHIFT > 1)), True, False)
        cmdShiftUp.Enabled = IIf((chgFlag And (NR_SHIFT > 1)), True, False)
        pnlShift.Enabled = IIf((chgFlag And (NR_SHIFT > 1)), True, False)
        CheckForIdle
    End If
End Sub

Private Sub cmdStart_Click()
    If (NR_SHIFT > 1) Then
        lblMessage.Caption = vbCrLf & "Starting Test on Station " & Format(CurRemoteTask.ActualStation, "0") & " Shift " & Format(CurRemoteTask.ActualShift, "0") & vbCrLf & vbCrLf
    Else
        lblMessage.Caption = vbCrLf & "Starting Test on Station " & Format(CurRemoteTask.ActualStation, "0") & vbCrLf & vbCrLf
    End If
    exitFlag = True
End Sub

Private Sub cmdShiftDn_Click()
    CurRemoteTask.ActualShift = IIf(CurRemoteTask.ActualShift <= 1, NR_SHIFT, CurRemoteTask.ActualShift - 1)
    pnlShift.Caption = "Shift " & Format(CurRemoteTask.ActualShift, "0")
    CheckForIdle
End Sub

Private Sub cmdShiftUp_Click()
    CurRemoteTask.ActualShift = IIf(CurRemoteTask.ActualShift = NR_SHIFT, 1, CurRemoteTask.ActualShift + 1)
    pnlShift.Caption = "Shift " & Format(CurRemoteTask.ActualShift, "0")
    CheckForIdle
End Sub

Private Sub cmdStnDn_Click()
    CurRemoteTask.ActualStation = IIf(CurRemoteTask.ActualStation <= 1, LAST_STN, CurRemoteTask.ActualStation - 1)
    pnlStn.Caption = "Station " & Format(CurRemoteTask.ActualStation, "0")
    CheckForIdle
End Sub

Private Sub cmdStnUp_Click()
    CurRemoteTask.ActualStation = IIf(CurRemoteTask.ActualStation >= LAST_STN, 1, CurRemoteTask.ActualStation + 1)
    pnlStn.Caption = "Station " & Format(CurRemoteTask.ActualStation, "0")
    CheckForIdle
End Sub

Private Sub dgRemoteTasks_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex + 1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
        If USINGREMCANLOAD Then
            frmSearchRemote.Show
        End If
        If USINGTOMCANLOAD Then
            frmSearchTom.Show
        End If
    End If
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 858, 2
    KeyPreview = True
    exitFlag = False
    curStatusCrit = "Ready"
    
    If USINGREMAVLFILES Then
        AVL_TaskFiles_Check
        adoRemoteTasks.Refresh
        dgRemoteTasks.Refresh
        canFlag = False
    Else
        canFlag = True
    End If
    
    lblMessage.ForeColor = Message_ForeColor
    
'    dgRemoteTasks.AllowAddNew = True
    dgRemoteTasks.AllowRowSizing = False
    
    DisplayData 14
    Setup_SelectTask
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If USINGREMCANLOAD Then
            Unload frmSearchRemote
            Set frmSearchRemote = Nothing
          End If
        If USINGTOMCANLOAD Then
            Unload frmSearchTom
            Set frmSearchTom = Nothing
        End If
End Sub

Private Sub DisplayData(sortCol As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 858, 888

Dim statusDesc As String
   ' get Description for desired status criteria
   ' and setup beginning of criteria string
   '"SELECT * FROM [RemoteTasks] WHERE [REM_TaskStatus] = '" & "Ready" & "' ORDER BY [REM_TaskID] ASC"
     Select Case curStatusCrit
        Case "Ready"
            statusDesc = "Ready-To-Run"
            rsCrit = "SELECT * FROM [RemoteTasks] WHERE [REM_TaskStatus] = '" & curStatusCrit & "'"
        Case "InProcess"
            statusDesc = "InProcess"
            rsCrit = "SELECT * FROM [RemoteTasks] WHERE [REM_TaskStatus] = '" & curStatusCrit & "'"
        Case "Done"
            statusDesc = "Completed"
            rsCrit = "SELECT * FROM [RemoteTasks] WHERE [REM_TaskStatus] = '" & curStatusCrit & "'"
        Case "All"
            statusDesc = " "
            rsCrit = "SELECT * FROM [RemoteTasks]"
        Case Else
            statusDesc = "Ready-To-Run"
            rsCrit = "SELECT * FROM [RemoteTasks] WHERE [REM_TaskStatus] = '" & curStatusCrit & "'"
    End Select
   ' Select & Sort
    Select Case sortCol
        Case 1
            rsCrit = rsCrit & " ORDER BY [REM_TaskID] ASC"
        Case 2
            rsCrit = rsCrit & " ORDER BY [REM_VIN] ASC"
        Case 3
            rsCrit = rsCrit & " ORDER BY [REM_RequestedStation] ASC"
        Case 4
            rsCrit = rsCrit & " ORDER BY [REM_RequestedShift] ASC"
        Case 14
            rsCrit = rsCrit & " ORDER BY [REM_OrderDate] ASC"
        Case Else
            rsCrit = rsCrit & " ORDER BY [REM_TaskID] ASC"
    End Select
    adoRemoteTasks.RecordSource = rsCrit
    adoRemoteTasks.Refresh

    ' Display number of RemoteTasks found
    If Not adoRemoteTasks.Recordset.BOF Then
        adoRemoteTasks.Recordset.GetRows
        Select Case adoRemoteTasks.Recordset.RecordCount
            Case 0
                dgRemoteTasks.Caption = " No " & statusDesc & " RemoteTasks"
            Case 1
                dgRemoteTasks.Caption = Format(adoRemoteTasks.Recordset.RecordCount, "###0") & " Only " & statusDesc & " Remote Task"
            Case Else
                dgRemoteTasks.Caption = Format(adoRemoteTasks.Recordset.RecordCount, "###0") & " " & statusDesc & " Remote Tasks"
        End Select
        
        ' Set column properties
    dgRemoteTasks.Columns(0).Width = 3000   ' taskID
    dgRemoteTasks.Columns(1).Width = 1800   ' vin
    dgRemoteTasks.Columns(2).Width = 1200   ' req stn
    dgRemoteTasks.Columns(3).Width = 1200   ' req shift
    dgRemoteTasks.Columns(4).Width = 1800   ' can desc
    dgRemoteTasks.Columns(5).Width = 1500   ' can vol
    dgRemoteTasks.Columns(6).Width = 1500   ' can wc
    dgRemoteTasks.Columns(7).Width = 1600   ' status
    dgRemoteTasks.Columns(8).Width = 2000   ' comment
    dgRemoteTasks.Columns(9).Width = 1200   ' job#
    dgRemoteTasks.Columns(10).Width = 1400  ' act stn
    dgRemoteTasks.Columns(11).Width = 1400  ' act shift
    dgRemoteTasks.Columns(12).Width = 1200  ' inh chgs
    dgRemoteTasks.Columns(13).Width = 1800  ' order date
    dgRemoteTasks.Columns(14).Width = 1800  ' start date
    dgRemoteTasks.Columns(15).Width = 1800  ' done date
    dgRemoteTasks.Columns(16).Width = 1600  ' prev result
    dgRemoteTasks.Columns(17).Width = 1400  ' rcp#
    dgRemoteTasks.Columns(18).Width = 2400  ' rcp name
    dgRemoteTasks.Columns(19).Width = 1400  ' prg#
    dgRemoteTasks.Columns(20).Width = 2400  ' prg name
    dgRemoteTasks.Columns(21).Width = 1400  ' seq#
    dgRemoteTasks.Columns(22).Width = 2400  ' seq name
    dgRemoteTasks.Columns(23).Width = 5400  ' AVL root
        
        ' move pointer to first row
        adoRemoteTasks.Recordset.MoveFirst
    Else
        dgRemoteTasks.Caption = " No " & statusDesc & " RemoteTasks"
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
    lblNumber.Enabled = False
    lblRcpNum.Enabled = False
    lblRcpDesc.Enabled = False
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
    lblNumber.Enabled = False
    lblRcpNum.Enabled = False
    lblRcpDesc.Enabled = False
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
    lblNumber.Enabled = True
    lblRcpNum.Enabled = True
    lblRcpDesc.Enabled = True
    frmRcp.Enabled = True
    opt2Gm.Enabled = True
    optWcm.Enabled = True
    cmdLoad.Enabled = False
    cmdStart.Enabled = False
    If (REMCHGSENABLED And (Not CurRemoteTask.InhibitChanges)) Then
        lblMessage.Caption = lblMessage.Caption & "Change the Station or Shift" & vbCrLf
        lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    End If
    lblMessage.Caption = lblMessage.Caption & "Select a Different Task" & vbCrLf
End Sub

Private Sub Setup_ReadyToLoad()
Dim chgFlag As Boolean
    cmdSelect.Enabled = True
    If (REMCHGSENABLED And (Not CurRemoteTask.InhibitChanges)) Then
        chgFlag = True
    Else
        chgFlag = False
    End If
    cmdStnDn.Enabled = chgFlag
    cmdStnUp.Enabled = chgFlag
    pnlStn.Enabled = chgFlag
    If (chgFlag And (NR_SHIFT > 1)) Then
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
    lblCanWorkingCapacity.Enabled = IIf(canFlag, True, False)
    lblCanWC_Units.Enabled = True
    lblCanVol.Enabled = True
    lblCanVolume.Enabled = IIf(canFlag, True, False)
    lblCanVol_Units.Enabled = True
    lblNumber.Enabled = True
    lblRcpNum.Enabled = IIf(canFlag, True, False)
    lblRcpDesc.Enabled = True
    frmRcp.Enabled = True
    opt2Gm.Enabled = True
    optWcm.Enabled = True
    cmdLoad.Enabled = True
    cmdStart.Enabled = False
    lblMessage.Caption = lblMessage.Caption & "Press Save to Load the Canister and Recipe to the station" & vbCrLf
    If chgFlag Then
        If canFlag Then
            lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
            lblMessage.Caption = lblMessage.Caption & "Change the Canister or Recipe Values" & vbCrLf
        End If
        lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
        lblMessage.Caption = lblMessage.Caption & "Change the Station or Shift" & vbCrLf
    End If
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Select a Different Task" & vbCrLf
End Sub

Private Sub Setup_ReadyToStart()
Dim chgFlag As Boolean
    cmdSelect.Enabled = True
    If (REMCHGSENABLED And (Not CurRemoteTask.InhibitChanges)) Then
        chgFlag = True
    Else
        chgFlag = False
    End If
    cmdStnDn.Enabled = chgFlag
    cmdStnUp.Enabled = chgFlag
    pnlStn.Enabled = chgFlag
    If (chgFlag And (NR_SHIFT > 1)) Then
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
    lblCanWorkingCapacity.Enabled = IIf(canFlag, True, False)
    lblCanWC_Units.Enabled = True
    lblCanVol.Enabled = True
    lblCanVolume.Enabled = IIf(canFlag, True, False)
    lblCanVol_Units.Enabled = True
    lblNumber.Enabled = True
    lblRcpNum.Enabled = IIf(canFlag, True, False)
    lblRcpDesc.Enabled = True
    frmRcp.Enabled = True
    opt2Gm.Enabled = True
    optWcm.Enabled = True
    cmdLoad.Enabled = True
    cmdStart.Enabled = True
    lblMessage.Caption = vbCrLf & "Requested Station/Shift is Ready to Start" & vbCrLf & vbCrLf & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Press Start to start the test" & vbCrLf
    If chgFlag Then
        If canFlag Then
            lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
            lblMessage.Caption = lblMessage.Caption & "Change the Canister or Recipe Values" & vbCrLf
        End If
        lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
        lblMessage.Caption = lblMessage.Caption & "Change the Station or Shift" & vbCrLf
    End If
    lblMessage.Caption = lblMessage.Caption & "or" & vbCrLf
    lblMessage.Caption = lblMessage.Caption & "Select a Different Task" & vbCrLf
End Sub

Private Sub CheckForIdle()
    If StationControl(CurRemoteTask.ActualStation, CurRemoteTask.ActualShift).Mode = VBIDLE Then
        lblMessage.Caption = vbCrLf & "Selected Station/Shift is Ready for Canister and Recipe Values" & vbCrLf & vbCrLf & vbCrLf
        Setup_ReadyToLoad
    Else
        lblMessage.Caption = vbCrLf & "Selected Station/Shift is not Idle" & vbCrLf & vbCrLf & vbCrLf
        If (REMCHGSENABLED And (Not CurRemoteTask.InhibitChanges)) Then
            Setup_ChangeStnShift
        Else
            Setup_SelectTask
        End If
    End If
End Sub

Private Sub tmrScreen_Timer()
    If exitFlag Then
        exitCntr = exitCntr + 1
        If (exitCntr > 5) Then
            frmStnDetail.RemoteStnStart CurRemoteTask.ActualStation, CurRemoteTask.ActualShift
            frmStnDetail.Show
            
            If USINGREMCANLOAD Then
                Unload frmSearchRemote
                Set frmSearchRemote = Nothing
            End If
            If USINGTOMCANLOAD Then
                Unload frmSearchTom
                Set frmSearchTom = Nothing
            End If
            
        End If
    Else
        exitCntr = 0
    End If
End Sub

