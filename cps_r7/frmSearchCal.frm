VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearchCal 
   Caption         =   "Saved Calibrations"
   ClientHeight    =   11145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   Icon            =   "frmSearchCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   14850
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid dbgCalibrations 
      Bindings        =   "frmSearchCal.frx":57E2
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "frmSearchCal.frx":57FD
      TabIndex        =   0
      Top             =   960
      Width           =   14610
   End
   Begin Threed.SSPanel pbxBottom 
      Height          =   960
      Left            =   0
      TabIndex        =   1
      Top             =   10200
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   1693
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
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmSearchCal.frx":61DC
         DownPicture     =   "frmSearchCal.frx":6E1E
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchCal.frx":7A60
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Delete the Selected Calibration"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin MSAdodcLib.Adodc adoRowData 
         Height          =   330
         Left            =   11520
         Top             =   300
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Connect         =   "DSN=CpsCalibrations"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "CpsCalibrations"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [MfcCalibrationsData] ORDER BY [Row] ASC"
         Caption         =   "RowData"
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
      Begin VB.Data adoSavedCals 
         Caption         =   "SavedCals"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   11520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   -30
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Close"
         DisabledPicture =   "frmSearchCal.frx":86A2
         DownPicture     =   "frmSearchCal.frx":92E4
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   13920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchCal.frx":9F26
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         DisabledPicture =   "frmSearchCal.frx":AB68
         DownPicture     =   "frmSearchCal.frx":B7AA
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchCal.frx":C3EC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Restore the Selected Calibration"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin MSAdodcLib.Adodc adoCalCheckData 
         Height          =   330
         Left            =   11520
         Top             =   630
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Connect         =   "DSN=CpsCalibrations"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "CpsCalibrations"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [MfcCalCheckData] ORDER BY [Row] ASC"
         Caption         =   "CalCheckData"
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
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "message"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   690
         Left            =   3840
         TabIndex        =   5
         Top             =   120
         Width           =   7605
      End
   End
   Begin VB.Frame frmRowData 
      Caption         =   "Flow Data for the Selected Calibration"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   14610
      Begin VB.CommandButton cmdPrev 
         Height          =   405
         Left            =   5395
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchCal.frx":D02E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "previous"
         Top             =   450
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdNext 
         Height          =   405
         Left            =   6875
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchCal.frx":D370
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "next"
         Top             =   450
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin MSDataGridLib.DataGrid dbgRowData 
         Bindings        =   "frmSearchCal.frx":D6B2
         Height          =   3705
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   3850
         _ExtentX        =   6800
         _ExtentY        =   6535
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   0   'False
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "MFC"
            Caption         =   "MFC"
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
            DataField       =   "DTS"
            Caption         =   "DTS"
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
            DataField       =   "Row"
            Caption         =   "Row"
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
            DataField       =   "Baro"
            Caption         =   "Baro"
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
            DataField       =   "NominalFlow"
            Caption         =   "NominalFlow"
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
            DataField       =   "ActualFlow"
            Caption         =   "ActualFlow"
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
            DataField       =   "LfeDiffPress"
            Caption         =   "LfeDiffPress"
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
            DataField       =   "LfeInletTemp"
            Caption         =   "LfeInletTemp"
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
            DataField       =   "LfeInletPress"
            Caption         =   "LfeInletPress"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1544.882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2280.189
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgCalCheckData 
         Bindings        =   "frmSearchCal.frx":D6CB
         Height          =   3705
         Index           =   2
         Left            =   12075
         TabIndex        =   11
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   6535
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   0   'False
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "MFC"
            Caption         =   "MFC"
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
            DataField       =   "DTS"
            Caption         =   "DTS"
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
            DataField       =   "CalCheckDTS"
            Caption         =   "CalCheckDTS"
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
            DataField       =   "Row"
            Caption         =   "Row"
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
            DataField       =   "Baro"
            Caption         =   "Baro"
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
            DataField       =   "FlowSP"
            Caption         =   "FlowSP"
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
            DataField       =   "CalCheckFlow"
            Caption         =   "CalCheckFlow"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2280.189
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgCalCheckData 
         Bindings        =   "frmSearchCal.frx":D6E9
         Height          =   3705
         Index           =   0
         Left            =   5395
         TabIndex        =   14
         Top             =   960
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   6535
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   0   'False
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "MFC"
            Caption         =   "MFC"
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
            DataField       =   "DTS"
            Caption         =   "DTS"
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
            DataField       =   "CalCheckDTS"
            Caption         =   "CalCheckDTS"
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
            DataField       =   "Row"
            Caption         =   "Row"
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
            DataField       =   "Baro"
            Caption         =   "Baro"
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
            DataField       =   "FlowSP"
            Caption         =   "FlowSP"
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
            DataField       =   "CalCheckFlow"
            Caption         =   "CalCheckFlow"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
               ColumnWidth     =   2099.906
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2280.189
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgCalCheckData 
         Bindings        =   "frmSearchCal.frx":D707
         Height          =   3705
         Index           =   1
         Left            =   9780
         TabIndex        =   15
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   6535
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   0   'False
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "MFC"
            Caption         =   "MFC"
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
            DataField       =   "DTS"
            Caption         =   "DTS"
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
            DataField       =   "CalCheckDTS"
            Caption         =   "CalCheckDTS"
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
            DataField       =   "Row"
            Caption         =   "Row"
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
            DataField       =   "Baro"
            Caption         =   "Baro"
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
            DataField       =   "FlowSP"
            Caption         =   "FlowSP"
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
            DataField       =   "CalCheckFlow"
            Caption         =   "CalCheckFlow"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   2429.858
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2280.189
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCalibrationDts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dts"
         DataField       =   "DTS"
         DataSource      =   "adoSavedCals"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   1920
         TabIndex        =   19
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label lblCalCheckDts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dts"
         DataField       =   "CalCheckDTS"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "YYYY-MM-DD  hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoCalCheckData"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Index           =   2
         Left            =   12060
         TabIndex        =   18
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label lblCalCheckDts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dts"
         DataField       =   "CalCheckDTS"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "YYYY-MM-DD  hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoCalCheckData"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Index           =   1
         Left            =   9780
         TabIndex        =   17
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label lblCalCheckDts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "dts"
         DataField       =   "CalCheckDTS"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "YYYY-MM-DD  hh:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "adoCalCheckData"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Index           =   0
         Left            =   7500
         TabIndex        =   16
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label lblCalCheck 
         Alignment       =   2  'Center
         Caption         =   "CalCheck"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   5850
         TabIndex        =   10
         Top             =   510
         Width           =   975
      End
      Begin VB.Label lblDts 
         Alignment       =   1  'Right Justify
         Caption         =   "Calibration "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1245
      End
   End
   Begin VB.Label lblCalibrations 
      Alignment       =   2  'Center
      Caption         =   "Previously Saved MFC Calibrations"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   14610
   End
End
Attribute VB_Name = "frmSearchCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 358'''''''''' Form SearchCal.frm '''''''''''''''''''
Option Explicit
Private daodb36 As DAO.Database
Private rS, rS2 As DAO.Recordset
Dim sPath As String
Dim rsCrit, rsCrit2 As String
Dim curStn, curMfc As Integer
Dim curDts As Date
Dim InitialLoadComplete As Boolean

Private Sub adoSavedCals_Reposition()
    If InitialLoadComplete Then DisplayRowData rS("Station"), rS("MFC"), rS("DTS")
End Sub

Private Sub cmdReturn_Click()
    Unload frmSearchCal
    Set frmSearchCal = Nothing
End Sub

Private Sub cmdSelect_Click()
    CalStation = curStn
    CalMfc = curMfc
    CalDts = curDts
'    frmRecipe.Show
'    frmRecipe.LoadNewRcp CInt(recnum)
    Unload frmSearchCal
    Set frmSearchCal = Nothing
End Sub

Private Sub dbgCalibrations_HeadClick(ByVal ColIndex As Integer)
    DisplayCals (ColIndex + 1), True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSearchCal = Nothing
    End If
End Sub

Private Sub Form_Load()
Dim sDesc As String
    SetErrModule 358, 2
    KeyPreview = True
    Form_Center Me
    
    InitialLoadComplete = False
    curStn = CalStation
    curMfc = CalMfc
    curDts = CalDts
    
    sDesc = "Previously Saved Calibrations for the "
    sDesc = sDesc & "Station #" & Format(curStn, "0")
    sDesc = sDesc & " " & Mfc_Description(curMfc) & " MFC"
    lblCalibrations.Caption = sDesc
    lblMsg.Caption = "Select from previously saved calibrations."
    cmdSelect.Visible = IIf(CheckPass("X", False), True, False)
    cmdDelete.Visible = IIf(CheckPass("6", False), True, False)
    dbgCalibrations.AllowRowSizing = False
    
    sPath = FILEPATH_cal & DATACAL
    Set daodb36 = DBEngine.OpenDatabase(sPath)
    DisplayCals 3, False
    
    InitialLoadComplete = True
    ResetErrModule
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub DisplayCals(ByVal sortCol As Integer, ByVal sortFlag As Boolean)
    ' Select & Sort
    rsCrit = "SELECT * FROM [MfcCalibrations] "
    rsCrit = rsCrit & "WHERE "
    rsCrit = rsCrit & "[Station] = " & curStn & " "
    rsCrit = rsCrit & "and "
    rsCrit = rsCrit & "[MFC] = " & curMfc & " "
    Select Case sortCol
        Case 1
            rsCrit = rsCrit & " ORDER BY [Station] ASC"
        Case 2
            rsCrit = rsCrit & " ORDER BY [Mfc] ASC"
        Case 3
            rsCrit = rsCrit & " ORDER BY [DTS] DESC"
        Case 4
            rsCrit = rsCrit & " ORDER BY [MfcDescription] ASC"
        Case 5
            rsCrit = rsCrit & " ORDER BY [CalibratedBy] ASC"
        Case 6
            rsCrit = rsCrit & " ORDER BY [Comment] ASC"
        Case 7
            rsCrit = rsCrit & " ORDER BY [NumRows] ASC"
        Case 8
            rsCrit = rsCrit & " ORDER BY [MfcMaxEU] ASC"
        Case 9
            rsCrit = rsCrit & " ORDER BY [MfcMinEU] ASC"
        Case 10
            rsCrit = rsCrit & " ORDER BY [CoefficientX] ASC"
        Case 11
            rsCrit = rsCrit & " ORDER BY [CoefficientX2] ASC"
        Case 12
            rsCrit = rsCrit & " ORDER BY [CoefficientX3] ASC"
        Case 13
            rsCrit = rsCrit & " ORDER BY [CoefficientX4] ASC"
        Case 14
            rsCrit = rsCrit & " ORDER BY [CoefficientX5] ASC"
        Case 15
            rsCrit = rsCrit & " ORDER BY [CoefficientX6] ASC"
        Case 16
            rsCrit = rsCrit & " ORDER BY [CoefficientR2] ASC"
        Case 17
            rsCrit = rsCrit & " ORDER BY [Lfe_a] ASC"
        Case 18
            rsCrit = rsCrit & " ORDER BY [Lfe_b] ASC"
        Case 19
            rsCrit = rsCrit & " ORDER BY [Lfe_c] ASC"
        Case 20
            rsCrit = rsCrit & " ORDER BY [Lfe_d] ASC"
        Case 21
            rsCrit = rsCrit & " ORDER BY [Lfe_filename] ASC"
        Case 22
            rsCrit = rsCrit & " ORDER BY [Lfe_serialnum] ASC"
        Case Else
            sortCol = 1
            sortFlag = True
            rsCrit = rsCrit & " ORDER BY [Station] ASC"
    End Select
    Set rS = daodb36.OpenRecordset(rsCrit, dbOpenDynaset)
    Set frmSearchCal.adoSavedCals.Recordset = rS

    ' Display number of calibrations found
    rS.GetRows
    Select Case rS.RecordCount
        Case 0
            dbgCalibrations.Caption = " No Saved Calibrations"
        Case 1
            dbgCalibrations.Caption = Format(rS.RecordCount, "###0") & " Saved Calibration"
        Case Else
            dbgCalibrations.Caption = Format(rS.RecordCount, "###0") & " Saved Calibrations"
    End Select
    
    ' Set column properties
        ' width
    dbgCalibrations.Columns(0).Width = 760
    dbgCalibrations.Columns(1).Width = 640
    dbgCalibrations.Columns(2).Width = 2040
    dbgCalibrations.Columns(3).Width = 2400
    dbgCalibrations.Columns(4).Width = 1600
    dbgCalibrations.Columns(5).Width = 4240
    dbgCalibrations.Columns(6).Width = 1100
    dbgCalibrations.Columns(7).Width = 1200
    dbgCalibrations.Columns(8).Width = 1200
    dbgCalibrations.Columns(9).Width = 1480
    dbgCalibrations.Columns(10).Width = 1480
    dbgCalibrations.Columns(11).Width = 1480
    dbgCalibrations.Columns(12).Width = 1480
    dbgCalibrations.Columns(13).Width = 1480
    dbgCalibrations.Columns(14).Width = 1480
    dbgCalibrations.Columns(15).Width = 1480
    dbgCalibrations.Columns(16).Width = 1360
    dbgCalibrations.Columns(17).Width = 1360
    dbgCalibrations.Columns(18).Width = 1360
    dbgCalibrations.Columns(19).Width = 1360
    dbgCalibrations.Columns(20).Width = 2240
    dbgCalibrations.Columns(21).Width = 2240
        ' alignment
    dbgCalibrations.Columns(0).Alignment = 2
    dbgCalibrations.Columns(1).Alignment = 2
    dbgCalibrations.Columns(2).Alignment = 0
    dbgCalibrations.Columns(3).Alignment = 0
    dbgCalibrations.Columns(20).Alignment = 1
    dbgCalibrations.Columns(21).Alignment = 1
    
    ' move pointer to first row
    adoSavedCals.Recordset.MoveFirst
    
    If sortFlag Then
        ' make the Left-Most column the Sorted-By column
        dbgCalibrations.LeftCol = IIf(sortCol > 22, 22, sortCol - 1)
    End If
    
    ' update current selection indexes
    curStn = rS("Station")
    curMfc = rS("MFC")
    curDts = rS("DTS")
    ' update row data
    DisplayRowData curStn, curMfc, curDts
    
End Sub


Private Sub DisplayRowData(ByVal selStn As Integer, ByVal selMfc As Integer, ByVal selDts As Date)
    ' Select & Sort
    rsCrit2 = "SELECT * FROM [MfcCalibrationsData] "
    rsCrit2 = rsCrit2 & "WHERE "
    rsCrit2 = rsCrit2 & "[Station] = " & selStn & " "
    rsCrit2 = rsCrit2 & "and "
    rsCrit2 = rsCrit2 & "[MFC] = " & selMfc & " "
    rsCrit2 = rsCrit2 & "and "
    rsCrit2 = rsCrit2 & "[DTS] = #" & selDts & "# "
    rsCrit2 = rsCrit2 & " ORDER BY [MfcCalibrationsData].[Row] ASC "
    Set rS2 = daodb36.OpenRecordset(rsCrit2, dbOpenDynaset)
    frmSearchCal.adoRowData.RecordSource = rsCrit2
    ' Refresh
    adoRowData.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSearchCal
    Set frmSearchCal = Nothing
End Sub
