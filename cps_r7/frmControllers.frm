VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmControllers 
   Caption         =   "Controllers"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   Icon            =   "frmControllers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmControllerParameters 
      Caption         =   "PID Controller Parameters"
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
      Height          =   10095
      Left            =   13000
      TabIndex        =   297
      Top             =   120
      Width           =   10635
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10800
         MaskColor       =   &H00FF00FF&
         Picture         =   "frmControllers.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   300
         ToolTipText     =   "Reload Controller Setup & Tuning Parameters"
         Top             =   9480
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dbgControllers 
         Bindings        =   "frmControllers.frx":6424
         Height          =   9675
         Left            =   120
         TabIndex        =   298
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   17066
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Number"
            Caption         =   "Number"
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
            DataField       =   "Pgain"
            Caption         =   "Pgain"
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
            DataField       =   "Igain"
            Caption         =   "Igain"
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
            DataField       =   "Dgain"
            Caption         =   "Dgain"
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
            DataField       =   "ReverseAction"
            Caption         =   "ReverseAction"
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
            DataField       =   "OffDuty"
            Caption         =   "OffDuty"
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
            DataField       =   "OffLimitDelta"
            Caption         =   "OffLimitDelta"
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
            DataField       =   "OnDuty"
            Caption         =   "OnDuty"
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
            DataField       =   "OnLimitDelta"
            Caption         =   "OnLimitDelta"
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
            DataField       =   "OffDutyMult"
            Caption         =   "OffDutyMult"
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
            DataField       =   "OnDutyMult"
            Caption         =   "OnDutyMult"
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
            DataField       =   "CumImax"
            Caption         =   "CumImax"
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
            DataField       =   "CumImin"
            Caption         =   "CumImin"
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
            DataField       =   "Outmax"
            Caption         =   "Outmax"
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
            DataField       =   "Outmin"
            Caption         =   "Outmin"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1124.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1065.26
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoControllers 
         Height          =   375
         Left            =   0
         Top             =   9720
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
         Connect         =   "DSN=cpsSysdef"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "cpsSysdef"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [Controllers] ORDER BY [Number] ASC"
         Caption         =   "Controllers"
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
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   5040
      Top             =   2760
   End
   Begin VB.Frame Frame1 
      Caption         =   "Misc"
      ForeColor       =   &H80000002&
      Height          =   795
      Left            =   120
      TabIndex        =   255
      Top             =   2760
      Width           =   11805
      Begin VB.Label lblNowValue 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "datetime"
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
         Left            =   6330
         TabIndex        =   258
         Top             =   240
         Width           =   5325
      End
      Begin VB.Label lblTimerDescr 
         BackStyle       =   0  'Transparent
         Caption         =   "Timer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   257
         Top             =   240
         Width           =   1025
      End
      Begin VB.Label lblTimerValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "86400.888"
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
         Left            =   1200
         TabIndex        =   256
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame frmChiller 
      Caption         =   "WaterBath Control"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   11835
      Begin VB.CommandButton cmdRun 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1665
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmControllers.frx":6441
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Run Chiller Communications"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   680
      End
      Begin VB.TextBox txtWriteTOvalue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Text            =   "1999"
         Top             =   1190
         Width           =   525
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   110
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Running"
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
         Index           =   0
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Width           =   1025
      End
      Begin VB.Label lblChillWrValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   0
         Left            =   2460
         TabIndex        =   108
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Command"
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
         Index           =   1
         Left            =   120
         TabIndex        =   107
         Top             =   480
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   106
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Complete"
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
         Index           =   9
         Left            =   3120
         TabIndex        =   105
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   9
         Left            =   4245
         TabIndex        =   104
         Top             =   240
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Cmd Chars"
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
         Index           =   2
         Left            =   120
         TabIndex        =   103
         Top             =   720
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   2
         Left            =   1200
         TabIndex        =   102
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "PV"
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
         Index           =   5
         Left            =   120
         TabIndex        =   101
         Top             =   1440
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   5
         Left            =   1200
         TabIndex        =   100
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "SP"
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
         Index           =   6
         Left            =   120
         TabIndex        =   99
         Top             =   1680
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   6
         Left            =   1200
         TabIndex        =   98
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label lblChillWrValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   6
         Left            =   2040
         TabIndex        =   97
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Output"
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
         Index           =   7
         Left            =   120
         TabIndex        =   96
         Top             =   1920
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   7
         Left            =   1200
         TabIndex        =   95
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblChillWrValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   94
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Oper Mode"
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
         Index           =   8
         Left            =   120
         TabIndex        =   93
         Top             =   2160
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   8
         Left            =   1200
         TabIndex        =   92
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label lblChillWrValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   91
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "ToBeSent"
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
         Index           =   10
         Left            =   3120
         TabIndex        =   90
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   10
         Left            =   4245
         TabIndex        =   89
         Top             =   480
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Sent"
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
         Index           =   11
         Left            =   3120
         TabIndex        =   88
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   11
         Left            =   4245
         TabIndex        =   87
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "ToBeAck"
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
         Index           =   12
         Left            =   3120
         TabIndex        =   86
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   12
         Left            =   4245
         TabIndex        =   85
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Ack"
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
         Index           =   13
         Left            =   3120
         TabIndex        =   84
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   13
         Left            =   4245
         TabIndex        =   83
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Error #"
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
         Index           =   14
         Left            =   3120
         TabIndex        =   82
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Index           =   14
         Left            =   4125
         TabIndex        =   81
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
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
         Index           =   15
         Left            =   6720
         TabIndex        =   80
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   15
         Left            =   7260
         TabIndex        =   79
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "I"
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
         Index           =   16
         Left            =   6720
         TabIndex        =   78
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   16
         Left            =   7260
         TabIndex        =   77
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
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
         Index           =   17
         Left            =   6720
         TabIndex        =   76
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   17
         Left            =   7260
         TabIndex        =   75
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Index           =   18
         Left            =   4920
         TabIndex        =   74
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   18
         Left            =   5640
         TabIndex        =   73
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Index           =   19
         Left            =   4920
         TabIndex        =   72
         Top             =   480
         Width           =   765
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   19
         Left            =   5640
         TabIndex        =   71
         Top             =   480
         Width           =   825
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Index           =   20
         Left            =   4920
         TabIndex        =   70
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   20
         Left            =   5640
         TabIndex        =   69
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Stat"
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
         Index           =   21
         Left            =   4920
         TabIndex        =   68
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00000"
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
         Index           =   21
         Left            =   5640
         TabIndex        =   67
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "OverTemp"
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
         Index           =   22
         Left            =   4920
         TabIndex        =   66
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   22
         Left            =   6120
         TabIndex        =   65
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "LowLevel"
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
         Index           =   23
         Left            =   4920
         TabIndex        =   64
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   23
         Left            =   6120
         TabIndex        =   63
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "PumpBlock"
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
         Index           =   24
         Left            =   4920
         TabIndex        =   62
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   24
         Left            =   6120
         TabIndex        =   61
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "IntMc1"
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
         Index           =   25
         Left            =   4920
         TabIndex        =   60
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   25
         Left            =   6120
         TabIndex        =   59
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "IntMc2"
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
         Index           =   26
         Left            =   4920
         TabIndex        =   58
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   26
         Left            =   6120
         TabIndex        =   57
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Rec Chars"
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
         Index           =   3
         Left            =   120
         TabIndex        =   56
         Top             =   960
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   3
         Left            =   1200
         TabIndex        =   55
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout at"
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
         Index           =   4
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Width           =   1025
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "86400.888"
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
         Index           =   4
         Left            =   1200
         TabIndex        =   53
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   0
         Left            =   10080
         TabIndex        =   52
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblResponse 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Response"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   10560
         TabIndex        =   51
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   0
         Left            =   8400
         TabIndex        =   50
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label lblCommand 
         BackStyle       =   0  'Transparent
         Caption         =   "Command"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   8400
         TabIndex        =   49
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   1
         Left            =   10080
         TabIndex        =   48
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   1
         Left            =   8400
         TabIndex        =   47
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   2
         Left            =   10080
         TabIndex        =   46
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   2
         Left            =   8400
         TabIndex        =   45
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   3
         Left            =   10080
         TabIndex        =   44
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   3
         Left            =   8400
         TabIndex        =   43
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   4
         Left            =   10080
         TabIndex        =   42
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   4
         Left            =   8400
         TabIndex        =   41
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   5
         Left            =   10080
         TabIndex        =   40
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   5
         Left            =   8400
         TabIndex        =   39
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   6
         Left            =   10080
         TabIndex        =   38
         Top             =   1920
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   6
         Left            =   8400
         TabIndex        =   37
         Top             =   1920
         Width           =   1605
      End
      Begin VB.Label lblChillResponse 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   7
         Left            =   10080
         TabIndex        =   36
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label lblChillCommand 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   7
         Left            =   8400
         TabIndex        =   35
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label lblHistory 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9660
         TabIndex        =   34
         Top             =   120
         Width           =   765
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   28
         Left            =   4245
         TabIndex        =   33
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Comm OK"
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
         Index           =   28
         Left            =   3120
         TabIndex        =   32
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   27
         Left            =   4125
         TabIndex        =   31
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "ErrorCount"
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
         Index           =   27
         Left            =   3120
         TabIndex        =   30
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Online"
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
         Index           =   29
         Left            =   3120
         TabIndex        =   29
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
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
         Index           =   29
         Left            =   4245
         TabIndex        =   28
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "333.3"
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
         Index           =   30
         Left            =   6720
         TabIndex        =   27
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "at"
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
         Index           =   30
         Left            =   6480
         TabIndex        =   26
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label lblDegUnits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "degC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7440
         TabIndex        =   25
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label lblChillRdValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
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
         Index           =   31
         Left            =   7800
         TabIndex        =   24
         Top             =   2160
         Width           =   210
      End
      Begin VB.Label lblChiller 
         BackStyle       =   0  'Transparent
         Caption         =   "Handshake"
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
         Index           =   31
         Left            =   6720
         TabIndex        =   23
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label lblDegUnits2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "degC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   1470
         Width           =   525
      End
   End
   Begin VB.Frame frmPIDControl 
      Caption         =   "PID Control"
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
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   11835
      Begin VB.CommandButton cmdLowTemp 
         Height          =   525
         Left            =   9360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmControllers.frx":6783
         Style           =   1  'Graphical
         TabIndex        =   304
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdNormalTemp 
         Height          =   525
         Left            =   8520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmControllers.frx":73C5
         Style           =   1  'Graphical
         TabIndex        =   303
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdHighTemp 
         Height          =   525
         Left            =   7680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmControllers.frx":8007
         Style           =   1  'Graphical
         TabIndex        =   302
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdWaterbathOK 
         Height          =   525
         Left            =   6533
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   301
         Top             =   5700
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdToggleView 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10800
         MaskColor       =   &H00FF00FF&
         Picture         =   "frmControllers.frx":8C49
         Style           =   1  'Graphical
         TabIndex        =   299
         ToolTipText     =   "Toggle Display of PID Controller Parameters"
         Top             =   5700
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdLoadControllers 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmControllers.frx":988B
         Style           =   1  'Graphical
         TabIndex        =   296
         ToolTipText     =   "Reload Controller Setup & Tuning Parameters"
         Top             =   5700
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdResetTOs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmControllers.frx":9F8D
         Style           =   1  'Graphical
         TabIndex        =   295
         ToolTipText     =   "Reset Temperature Controller Timeouts"
         Top             =   5700
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblItem20Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   294
         Top             =   5040
         Width           =   1500
      End
      Begin VB.Label lblItem19Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   293
         Top             =   4800
         Width           =   1500
      End
      Begin VB.Label lblItem21Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   292
         Top             =   5280
         Width           =   1500
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   291
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   290
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   289
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   288
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   287
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   286
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   285
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   284
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   283
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   282
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   281
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   280
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   279
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   278
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   277
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   276
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   275
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   274
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   273
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   272
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   271
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblValue21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   270
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label lblValue20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   269
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label lblValue19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   268
         Top             =   4800
         Width           =   1200
      End
      Begin VB.Label lblItem18Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   267
         Top             =   4560
         Width           =   1500
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   266
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   265
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   264
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   263
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   262
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   261
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   260
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblValue18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   259
         Top             =   4560
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   8
         Left            =   9960
         TabIndex        =   254
         Top             =   240
         Width           =   1800
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   253
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   252
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   251
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   250
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   249
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   248
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   247
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   246
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   245
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   244
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   243
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   242
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   241
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   240
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   239
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   238
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   8
         Left            =   9960
         TabIndex        =   237
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   7
         Left            =   8760
         TabIndex        =   236
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   235
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   234
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   233
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   232
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   231
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   230
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   229
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   228
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   227
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   226
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   225
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   224
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   223
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   222
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   221
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   220
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   7
         Left            =   8760
         TabIndex        =   219
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   6
         Left            =   7560
         TabIndex        =   218
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   217
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   216
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   215
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   214
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   213
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   212
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   211
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   210
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   209
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   208
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   207
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   206
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   205
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   204
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   203
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   202
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   201
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   200
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   199
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   198
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   197
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   196
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   195
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   194
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   193
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   192
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   191
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   190
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   189
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   188
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   187
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   186
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   185
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   184
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   5
         Left            =   6360
         TabIndex        =   183
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   182
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   181
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   180
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   179
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   178
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   177
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   176
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   175
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   174
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   173
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   172
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   171
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   170
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   169
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   168
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   167
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   166
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   4
         Left            =   5160
         TabIndex        =   165
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   164
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   163
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   162
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   161
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   160
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   159
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   158
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   157
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   156
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   155
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   154
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   153
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   152
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   151
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   150
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   149
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   148
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   3
         Left            =   3960
         TabIndex        =   147
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   146
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   145
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   144
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   143
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   142
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   141
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   140
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   139
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   138
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   137
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   136
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   135
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   134
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   133
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   132
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   131
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   130
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   129
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblTemperature 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   128
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValue01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   127
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblValue02 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   126
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblValue03 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   125
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblValue04 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   124
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblValue05 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   123
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label lblValue06 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   122
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblValue07 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   121
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label lblValue08 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   120
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblValue09 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   119
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblValue10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   118
         Top             =   2640
         Width           =   1200
      End
      Begin VB.Label lblValue11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   117
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblValue12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   116
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label lblValue13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   115
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblValue14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   114
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblValue15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   113
         Top             =   3840
         Width           =   1200
      End
      Begin VB.Label lblValue16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   112
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lblValue17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   111
         Top             =   4320
         Width           =   1200
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItem01Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblItem02Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblItem03Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lblItem04Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lblItem05Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lblItem06Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lblItem07Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label lblItem08Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label lblItem09Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label lblItem10Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label lblItem11Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label lblItem12Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1500
      End
      Begin VB.Label lblItem13Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label lblItem14Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   1500
      End
      Begin VB.Label lblItem15Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   1500
      End
      Begin VB.Label lblItem16Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   2
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label lblItem17Descr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
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
         Left            =   120
         TabIndex        =   1
         Top             =   4320
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmControllers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Controllers Monitor Form
'
'

Private Sub cmdLoadControllers_Click()
    Load_Controllers
End Sub

Private Sub cmdResetTOs_Click()
    ' nothing defined
End Sub

Private Sub cmdRun_Click()
    If ChillCommOn Then
        ChillerRun
    Else
        LF_Chiller.ChillerRunning = IIf(LF_Chiller.ChillerRunning, False, True)
    End If
End Sub

Private Sub cmdToggleView_Click()
Dim iLeft As Long
    Select Case frmControllerParameters.Left
        Case frmChiller.Left
            iLeft = OutOfSight
        Case Else
            iLeft = frmChiller.Left
    End Select
    frmControllerParameters.Left = iLeft
End Sub

Private Sub Form_Load()

    ' foecolors
    frmPIDControl.ForeColor = Titles_ForeColor
    frmChiller.ForeColor = Titles_ForeColor
    frmControllerParameters.ForeColor = Titles_ForeColor
    Frame1.ForeColor = Titles_ForeColor
    Frame1.FontBold = True

    ' PID Control Column Descriptions
    lblTemperature(pasTEMPERATURE).Caption = "PAS Temp"
    lblTemperature(pasMOISTURE).Caption = "PAS Moisture"
    lblTemperature(3).Caption = "unused"
    lblTemperature(4).Caption = "unused"
    lblTemperature(wbSuperTemp).Caption = "WB Super"
    lblTemperature(6).Caption = "unused"
    lblTemperature(7).Caption = "unused"
    lblTemperature(8).Caption = "unused"
    
    ' PID Control Item Descriptions
    lblItem01Descr.Caption = "SP    (deg C)"
    lblItem02Descr.Caption = "PV    (deg C)"
    lblItem03Descr.Caption = "Error (deg C)"
    lblItem04Descr.Caption = "Cum I"
    lblItem05Descr.Caption = "Output %"
    lblItem06Descr.Caption = "P"
    lblItem07Descr.Caption = "I"
    lblItem08Descr.Caption = "D"
    lblItem09Descr.Caption = "Enable"
    lblItem10Descr.Caption = "Inhibit"
    lblItem11Descr.Caption = "Action"
    lblItem12Descr.Caption = "Duty Cycle"
    lblItem13Descr.Caption = "OnDuty"
    lblItem14Descr.Caption = "OffDuty"
    lblItem15Descr.Caption = "On Timer"
    lblItem16Descr.Caption = "Off Timer"
    lblItem17Descr.Caption = "Output Relay"
    lblItem18Descr.Caption = "TO Max Secs"
    lblItem19Descr.Caption = "TO Timer"
    lblItem20Descr.Caption = "TO Reset PV"
    lblItem21Descr.Caption = "OOT Timer"
    lblItem17Descr.ForeColor = DK3RED


    ' clear captions for Heater Booleans
    lblChillRdValue(0).Caption = " "
    lblChillRdValue(9).Caption = " "
    lblChillRdValue(10).Caption = " "
    lblChillRdValue(11).Caption = " "
    lblChillRdValue(12).Caption = " "
    lblChillRdValue(13).Caption = " "
    lblChillRdValue(22).Caption = " "
    lblChillRdValue(23).Caption = " "
    lblChillRdValue(24).Caption = " "
    lblChillRdValue(25).Caption = " "
    lblChillRdValue(26).Caption = " "
    lblChillRdValue(27).Caption = " "
    lblChillRdValue(28).Caption = " "
    lblChillRdValue(29).Caption = " "
    lblChillWrValue(0).Caption = " "
    ' set caption for temp units
    lblDegUnits.Caption = IIf(USINGC, "degC", "degF")
    lblDegUnits2.Caption = lblDegUnits.Caption
    ' set caption for timeout (in msec)
    txtWriteTOvalue.text = Format(LF_Chiller.TimeoutValue, "###0")

       
    ' Setup Controllers Table
    frmControllers.adoControllers.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & FILEPATH_sysdbf & DATASYSDEF & ";" _
        & "Persist Security Info=False"
    adoControllers.Refresh
    frmControllers.dbgControllers.Refresh
    
    
    ' Start the timer
    tmrUpdate.Enabled = True


End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrUpdate.Enabled = False
    Unload Me
End Sub

Private Sub lblChiller_Click(Index As Integer)
Dim tmpI As Integer
    Select Case Index
        Case 31
            tmpI = frmMainMenu.MSComm(mscommChiller).Handshaking
            frmMainMenu.MSComm(mscommChiller).Handshaking = IIf(tmpI < comRTSXOnXOff, tmpI + 1, comNone)
    End Select
End Sub

Private Sub tmrUpdate_Timer()
    UpdateReadouts
End Sub

Private Sub txtWriteTOvalue_Change()
    If Not IsNumeric(txtWriteTOvalue.text) Then Exit Sub
    LF_Chiller.TimeoutValue = CSng(txtWriteTOvalue.text)
End Sub

Private Sub UpdateReadouts()
Dim Index As Integer
Dim tempTemp As Single
    
    ' Chiller
    lblChillRdValue(0).BackColor = IIf(LF_Chiller.ChillerRunning, DKGREEN, MEDORANGE)
    lblChillRdValue(1).Caption = LF_Chiller.CurCmdDesc
    lblChillRdValue(2).Caption = LF_Chiller.CurCmdChars
    lblChillRdValue(3).Caption = LF_Chiller.CmdRecChars
    lblChillRdValue(4).BackColor = IIf(LF_Chiller.CurCmdTimeout, MEDRED, White)
    lblChillRdValue(4).Caption = Format(LF_Chiller.CmdTimeoutTimer, "####0.000")
    tempTemp = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
    lblChillRdValue(5).Caption = Format(tempTemp, "##0.0")
    tempTemp = IIf(USINGC, LF_Chiller.SpIn, DegCtoF(LF_Chiller.SpIn))
    lblChillRdValue(6).Caption = Format(tempTemp, "##0.0")
    lblChillRdValue(7).Caption = Format(LF_Chiller.OutIn, "0")
    lblChillRdValue(8).Caption = Format(LF_Chiller.OperModeIn, "0")
    lblChillRdValue(9).BackColor = IIf(LF_Chiller.CurCmdComplete, DKGREEN, MEDORANGE)
    lblChillRdValue(10).BackColor = IIf(LF_Chiller.CmdToBeSentFlag, DKGREEN, MEDORANGE)
    lblChillRdValue(11).BackColor = IIf(LF_Chiller.CmdSentFlag, DKGREEN, MEDORANGE)
    lblChillRdValue(12).BackColor = IIf(LF_Chiller.CmdToBeAckFlag, DKGREEN, MEDORANGE)
    lblChillRdValue(13).BackColor = IIf(LF_Chiller.CmdRecAckFlag, DKGREEN, MEDORANGE)
    lblChillRdValue(14).BackColor = IIf(LF_Chiller.CmdRecErrorFlag, MEDRED, White)
    lblChillRdValue(14).Caption = Format(LF_Chiller.CmdRecErrorNumber, "##0")
    lblChillRdValue(15).Caption = Format(LF_Chiller.P_In, "##0.0")
    lblChillRdValue(16).Caption = Format(LF_Chiller.I_In, "#0")
    lblChillRdValue(17).Caption = Format(LF_Chiller.ModeIn, "0")
    lblChillRdValue(18).Caption = CStr(LF_Chiller.Type)
    lblChillRdValue(19).Caption = CStr(LF_Chiller.Version)
    lblChillRdValue(20).Caption = Format(LF_Chiller.StatusIn, "#0")
    lblChillRdValue(21).Caption = Format(LF_Chiller.StatIn, "00000")
    lblChillRdValue(22).BackColor = IIf(LF_Chiller.Overtemp, DKGREEN, MEDORANGE)
    lblChillRdValue(23).BackColor = IIf(LF_Chiller.LowLevel, DKGREEN, MEDORANGE)
    lblChillRdValue(24).BackColor = IIf(LF_Chiller.PumpBlocked, DKGREEN, MEDORANGE)
    lblChillRdValue(25).BackColor = IIf(LF_Chiller.IntFaultMc1, DKGREEN, MEDORANGE)
    lblChillRdValue(26).BackColor = IIf(LF_Chiller.IntFaultMc2, DKGREEN, MEDORANGE)
    lblChillRdValue(27).Caption = Format(LF_Chiller.ErrorCount, "##0")
    lblChillRdValue(27).BackColor = IIf(LF_Chiller.ErrorCount > LF_Chiller.MaxErrorCount, MEDRED, _
                                    IIf(LF_Chiller.ErrorCount > 0, MEDYELLOW, DKGREEN))
    lblChillRdValue(28).BackColor = IIf(LF_Chiller.CommOK, DKGREEN, MEDRED)
    lblChillRdValue(29).BackColor = IIf(LF_Chiller.CommOnline, DKGREEN, MEDRED)
    tempTemp = IIf(USINGC, LF_Chiller.OvrTmpSpIn, DegCtoF(LF_Chiller.OvrTmpSpIn))
    lblChillRdValue(30).Caption = Format(tempTemp, "##0.0")
    lblChillRdValue(30).BackColor = IIf(LF_Chiller.Overtemp, MEDRED, White)
    lblChillRdValue(31).Caption = Format(frmMainMenu.MSComm(mscommChiller).Handshaking, "0")
    lblChillWrValue(0).BackColor = IIf(LF_Chiller.RunChiller, DKGREEN, MEDORANGE)
    tempTemp = IIf(USINGC, LF_Chiller.SpOut, DegCtoF(LF_Chiller.SpOut))
    lblChillWrValue(6).Caption = Format(tempTemp, "##0.0")
    lblChillWrValue(7).Caption = Format(LF_Chiller.OutIn, "0")
    lblChillWrValue(8).Caption = Format(LF_Chiller.OperModeIn, "0")
            
    lblTimerValue.Caption = Format(Timer, "####0.000")
    lblNowValue.Caption = Format(Now, "YYYY MMMM DD   hh:mm:ss")
            
    For Index = 0 To 7
        lblChillResponse(Index).Caption = ChillerResponse(Index)
        lblChillResponse(Index).BackColor = IIf(InStr(ChillerResponse(Index), "ERR"), MEDRED, _
                                            IIf(InStr(ChillerResponse(Index), "timeout"), MEDRED, White))
        lblChillCommand(Index).Caption = ChillerCommands(Index)
        lblChillCommand(Index).BackColor = lblChillResponse(Index).BackColor
    '    lblChillResponse(Index).Caption = Format(TimerTimes(5, Index + 1), "####0.000")
    '    lblChillCommand(Index).Caption = Format(TimerTimes(1, Index + 1), "####0.000")
    Next Index
    
    For Index = 1 To 8
        ' show PID control block values
        lblValue01(Index).Caption = Format(PID_INFO(Index).SP, "####0.00")
        lblValue02(Index).Caption = Format(PID_INFO(Index).PV, "####0.00")
'        lblValue02(Index).ForeColor = IIf(PID_INFO(Index).OOT, MEDRED, Black)
        lblValue03(Index).Caption = Format(PID_INFO(Index).Er, "####0.00")
        lblValue04(Index).Caption = Format(PID_INFO(Index).CumI, "####0.00")
        lblValue05(Index).Caption = Format(PID_INFO(Index).out, "####0.00")
        lblValue06(Index).Caption = Format(PID_INFO(Index).Pgain, "####0.00")
        lblValue07(Index).Caption = Format(PID_INFO(Index).Igain, "####0.00")
        lblValue08(Index).Caption = Format(PID_INFO(Index).Dgain, "####0.00")
        lblValue09(Index).Caption = IIf(PID_INFO(Index).Enable, "ENABLED", "false")
        lblValue09(Index).ForeColor = IIf(PID_INFO(Index).Enable, DKGREEN, DKYELLOW)
        lblValue10(Index).Caption = IIf(PID_INFO(Index).Inhibit, "INHIBITED", "false")
        lblValue10(Index).ForeColor = IIf(PID_INFO(Index).Inhibit, DKGREEN, DKYELLOW)
        lblValue11(Index).Caption = IIf(PID_INFO(Index).Rev, "reverse", "direct")
        lblValue11(Index).ForeColor = IIf(PID_INFO(Index).Rev, DKGREEN, DKYELLOW)
'        lblValue12(Index).Caption = Format(PID_INFO(Index).DutyCycle, "#####0.000")
        lblValue13(Index).Caption = Format(PID_INFO(Index).OnDuty, "#####0.000")
        lblValue14(Index).Caption = Format(PID_INFO(Index).OffDuty, "#####0.000")
        lblValue15(Index).Caption = Format(PID_INFO(Index).OnTimer, "#####0.000")
        lblValue16(Index).Caption = Format(PID_INFO(Index).OffTimer, "#####0.000")
        lblValue17(Index).Caption = IIf(PID_INFO(Index).Output, "HEAT", "off")
        lblValue17(Index).ForeColor = IIf(PID_INFO(Index).Output, DK2RED, DKYELLOW)
'        lblValue18(Index).Caption = Format(PID_INFO(Index).TO_TimerMax, "####0.00")
'        lblValue19(Index).Caption = Format(PID_INFO(Index).TO_Timer, "####0.00")
'        lblValue19(Index).ForeColor = IIf(PID_INFO(Index).timeOut, MEDRED, Black)
'        lblValue20(Index).Caption = Format(PID_INFO(Index).TO_ResetPV, "####0.00")
'        lblValue21(Index).Caption = Format(PID_INFO(Index).OotTimer, "#####0.000")
'        lblValue21(Index).ForeColor = IIf(PID_INFO(Index).OOT, MEDRED, Black)
        lblValue01(Index).Visible = True
        lblValue02(Index).Visible = True
        lblValue03(Index).Visible = True
'        lblValue04(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue05(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue06(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue07(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue08(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue09(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue10(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue11(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue12(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue13(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue14(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue15(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue16(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue17(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue18(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue19(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
'        lblValue20(Index).Visible = IIf(Index = tcbWaterBathTemp, False, True)
        lblValue21(Index).Visible = True
    Next Index
    cmdWaterbathOK.Picture = IIf(LoadControl(3, 1).WaterBathTempOK, cmdNormalTemp.Picture, cmdHighTemp.Picture)
End Sub
