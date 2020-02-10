VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmMassFlowCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Flow Controller Calibration Screen"
   ClientHeight    =   10695
   ClientLeft      =   195
   ClientTop       =   720
   ClientWidth     =   15255
   Icon            =   "FrmMassFlowCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   15255
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   218
      Top             =   3525
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   216
      Top             =   3765
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   2
      Left            =   720
      TabIndex        =   215
      Top             =   4005
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   214
      Top             =   4245
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   4
      Left            =   720
      TabIndex        =   213
      Top             =   4485
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   212
      Top             =   4725
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   6
      Left            =   720
      TabIndex        =   211
      Top             =   4965
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   7
      Left            =   720
      TabIndex        =   210
      Top             =   5205
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   8
      Left            =   720
      TabIndex        =   209
      Top             =   5445
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Caption         =   "Option8"
      Height          =   195
      Index           =   9
      Left            =   720
      TabIndex        =   208
      Top             =   5685
      Width           =   255
   End
   Begin VB.OptionButton optCalibTable 
      Height          =   195
      Index           =   10
      Left            =   720
      TabIndex        =   207
      Top             =   5925
      Width           =   255
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   206
      Text            =   "FrmMassFlowCal.frx":57E2
      Top             =   9960
      Width           =   3855
   End
   Begin VB.CommandButton cmdReadOnly 
      Caption         =   "Toggle Read-Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9420
      TabIndex        =   205
      Top             =   9510
      Width           =   2115
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   136
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   10
      Left            =   4680
      TabIndex        =   140
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   96
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   100
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   104
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   108
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   112
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   116
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   120
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   7
      Left            =   4680
      TabIndex        =   124
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtLFEDiffPress 
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   132
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   0
      Left            =   2670
      TabIndex        =   93
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   1
      Left            =   2670
      TabIndex        =   97
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   2
      Left            =   2670
      TabIndex        =   101
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   3
      Left            =   2670
      TabIndex        =   105
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   4
      Left            =   2670
      TabIndex        =   109
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   5
      Left            =   2670
      TabIndex        =   113
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   6
      Left            =   2670
      TabIndex        =   117
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   7
      Left            =   2670
      TabIndex        =   121
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   8
      Left            =   2670
      TabIndex        =   125
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   9
      Left            =   2670
      TabIndex        =   133
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtInletTemp 
      Height          =   285
      Index           =   10
      Left            =   2670
      TabIndex        =   137
      Top             =   5880
      Width           =   1095
   End
   Begin MSChart20Lib.MSChart chtMFCChart 
      Height          =   5715
      Left            =   7920
      OleObjectBlob   =   "FrmMassFlowCal.frx":580D
      TabIndex        =   33
      Top             =   3720
      Visible         =   0   'False
      Width           =   7305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7800
      TabIndex        =   0
      Top             =   9510
      Width           =   1620
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   10
      Left            =   6600
      TabIndex        =   30
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   9
      Left            =   6600
      TabIndex        =   29
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   8
      Left            =   6600
      TabIndex        =   28
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   7
      Left            =   6600
      TabIndex        =   27
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   6
      Left            =   6600
      TabIndex        =   26
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   5
      Left            =   6600
      TabIndex        =   25
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   4
      Left            =   6600
      TabIndex        =   24
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   22
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtActualFlowSLPM 
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   5640
      TabIndex        =   185
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   5640
      TabIndex        =   186
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   187
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   5640
      TabIndex        =   188
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   189
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   190
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5640
      TabIndex        =   191
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   192
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   193
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   194
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtBaromPress 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   195
      Top             =   3480
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6240
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   10
      Left            =   3750
      TabIndex        =   139
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   9
      Left            =   3750
      TabIndex        =   135
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   8
      Left            =   3750
      TabIndex        =   127
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   7
      Left            =   3750
      TabIndex        =   123
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   6
      Left            =   3750
      TabIndex        =   119
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   5
      Left            =   3750
      TabIndex        =   115
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   4
      Left            =   3750
      TabIndex        =   111
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   3
      Left            =   3750
      TabIndex        =   107
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   2
      Left            =   3750
      TabIndex        =   103
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   1
      Left            =   3750
      TabIndex        =   99
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtLFEInletPress 
      Height          =   285
      Index           =   0
      Left            =   3750
      TabIndex        =   95
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveNewCalib 
      Caption         =   "Save New Calib."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   9960
      Width           =   1500
   End
   Begin VB.CommandButton cmdKeepPrev 
      Caption         =   "Keep Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4620
      TabIndex        =   1
      Top             =   9960
      Width           =   1500
   End
   Begin VB.CommandButton cmdCalCheckFunction 
      Caption         =   "Cal. check Function"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   9960
      Width           =   1500
   End
   Begin VB.CommandButton cmdCalcActualFlow 
      Caption         =   "Calculate Actual Flow"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   9960
      Width           =   1500
   End
   Begin VB.CommandButton cmdCalibrate 
      Caption         =   "Calibrate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1620
      TabIndex        =   4
      Top             =   9960
      Width           =   1500
   End
   Begin Threed.SSCommand cmdNewLFEFile 
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Select File"
   End
   Begin VB.Frame fraCalibInfo 
      Caption         =   "Calibration Information"
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
      Height          =   2415
      Left            =   7800
      TabIndex        =   39
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtComments 
         Height          =   1095
         Left            =   1440
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox txtNumCalPts 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   600
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "11"
         ToolTipText     =   "Number of points used to calibrate the system.  Enter a number between 2 and 11."
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtCalibBy 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   18
         ToolTipText     =   "Maximum length: 20 characters"
         Top             =   840
         Width           =   5775
      End
      Begin Threed.SSCommand cmdCalPtUp 
         Height          =   600
         Left            =   4575
         TabIndex        =   38
         ToolTipText     =   "Up"
         Top             =   210
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         ForeColor       =   8421504
         BevelWidth      =   5
         AutoSize        =   1
         Picture         =   "FrmMassFlowCal.frx":7B63
      End
      Begin Threed.SSCommand cmdCalPtDown 
         Height          =   600
         Left            =   3240
         TabIndex        =   36
         ToolTipText     =   "Down"
         Top             =   210
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1058
         _StockProps     =   78
         ForeColor       =   8421504
         BevelWidth      =   5
         AutoSize        =   1
         Picture         =   "FrmMassFlowCal.frx":7FB5
      End
      Begin VB.Label lblComments 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         BeginProperty Font 
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
         TabIndex        =   42
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblNumCalPts 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Calibration Points"
         BeginProperty Font 
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
         TabIndex        =   41
         Top             =   390
         Width           =   2535
      End
      Begin VB.Label lblCalibBy 
         BackStyle       =   0  'Transparent
         Caption         =   "Calibrated By:"
         BeginProperty Font 
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
         TabIndex        =   40
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame fraCalibMethod 
      Caption         =   "Calibration Method"
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
      Height          =   1815
      Left            =   3030
      TabIndex        =   34
      Top             =   0
      Width           =   4695
      Begin Threed.SSCommand cmdSetToNominal 
         Height          =   375
         Left            =   120
         TabIndex        =   200
         ToolTipText     =   "Set ActualFlow values as ""Ideal"" (i.e. Actual = Nominal; always)"
         Top             =   1320
         Visible         =   0   'False
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Set Actual Flows to Nominal "
         ForeColor       =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   4
      End
      Begin VB.OptionButton optLaminarFlowElement 
         Caption         =   "Laminar Flow Element"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optFlowStandard 
         Caption         =   "Flow Standard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblLFEConfig 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*.LFE"
         Height          =   255
         Left            =   120
         TabIndex        =   184
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label lblLFEConfigFile 
         BackStyle       =   0  'Transparent
         Caption         =   "LFE Configuration File"
         BeginProperty Font 
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
         TabIndex        =   35
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame fraMFCSelection 
      Caption         =   "Mass Flow Controller Selection"
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
      Height          =   2685
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   2895
      Begin VB.OptionButton optORVRPurgeAir 
         Caption         =   "ORVR - PurgeAir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   1980
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton optORVRNit 
         Caption         =   "ORVR - Nitrogen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   1620
         Width           =   1815
      End
      Begin VB.OptionButton optORVRBut 
         Caption         =   "ORVR - Butane"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1260
         Width           =   1815
      End
      Begin VB.OptionButton optLiveFuel 
         Caption         =   "Live Fuel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   2340
         Width           =   1335
      End
      Begin VB.OptionButton optPurgeAir 
         Caption         =   "Purge Air"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton optNitrogen 
         Caption         =   "Nitrogen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton optButane 
         Caption         =   "Butane"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin Threed.SSCommand cmdMFCUp 
      Height          =   600
      Left            =   0
      TabIndex        =   31
      ToolTipText     =   "Up"
      Top             =   3720
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   78
      ForeColor       =   8421504
      BevelWidth      =   5
      AutoSize        =   1
      Picture         =   "FrmMassFlowCal.frx":8407
   End
   Begin Threed.SSCommand cmdMFCDown 
      Height          =   600
      Left            =   0
      TabIndex        =   32
      ToolTipText     =   "Down"
      Top             =   4800
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   78
      ForeColor       =   8421504
      BevelWidth      =   5
      AutoSize        =   1
      Picture         =   "FrmMassFlowCal.frx":8859
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   615
      Left            =   13680
      TabIndex        =   201
      Top             =   9960
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel txtDispStn 
      Height          =   615
      Left            =   12240
      TabIndex        =   202
      ToolTipText     =   "Station Number Displayed"
      Top             =   9960
      Width           =   690
      _Version        =   65536
      _ExtentX        =   1217
      _ExtentY        =   1085
      _StockProps     =   15
      Caption         =   "9"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   615
      Left            =   12960
      TabIndex        =   203
      ToolTipText     =   "Next"
      Top             =   9960
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1085
      _StockProps     =   78
      ForeColor       =   8421504
      BevelWidth      =   5
      AutoSize        =   1
      Picture         =   "FrmMassFlowCal.frx":8CAB
   End
   Begin Threed.SSCommand cmdDown 
      Height          =   615
      Left            =   11640
      TabIndex        =   204
      ToolTipText     =   "Previous"
      Top             =   9960
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1085
      _StockProps     =   78
      ForeColor       =   8421504
      BevelWidth      =   5
      AutoSize        =   1
      Picture         =   "FrmMassFlowCal.frx":90FD
   End
   Begin VB.Frame frmCalFormula 
      Caption         =   "Calibration Formula"
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
      Height          =   975
      Left            =   7800
      TabIndex        =   219
      Top             =   2430
      Width           =   7335
      Begin VB.Label lblCalibRslt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "The formula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   220
         Top             =   240
         Visible         =   0   'False
         Width           =   7095
      End
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   231
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   230
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   229
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   228
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   227
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   226
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   225
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   224
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   223
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   222
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblMFCPercFS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1080
      TabIndex        =   221
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   217
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblLFEDiffPressCol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LFE Diff. Press."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      TabIndex        =   199
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblInletTempCol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inlet Temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2670
      TabIndex        =   198
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblCalGraphTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calibration Graph"
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
      Height          =   255
      Left            =   9720
      TabIndex        =   197
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblCalGraphRow 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Percent Full Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12960
      TabIndex        =   196
      Top             =   9480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBaromPressCol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barometric Pressure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5640
      TabIndex        =   52
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblActiualFlowCol1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual Flow (slpm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6600
      TabIndex        =   53
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lblLFEInletPressCol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LFE Inlet Press."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3765
      TabIndex        =   51
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblCalTableTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Mass Flow Controller Calibration Table"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   173
      Top             =   2700
      Width           =   3375
   End
   Begin VB.Label lblPrevCalTitle 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Calibration"
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
      Height          =   255
      Left            =   5640
      TabIndex        =   172
      Top             =   6420
      Width           =   1935
   End
   Begin VB.Label lblCurrCalTitle 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Calibration"
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
      Height          =   255
      Left            =   1560
      TabIndex        =   171
      Top             =   6420
      Width           =   1695
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   170
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   6480
      TabIndex        =   169
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   6480
      TabIndex        =   168
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   6480
      TabIndex        =   167
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   6480
      TabIndex        =   166
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   6480
      TabIndex        =   165
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   164
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   6480
      TabIndex        =   163
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   162
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   161
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCol2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% Diff Reading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6480
      TabIndex        =   160
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   159
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   5280
      TabIndex        =   158
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   5280
      TabIndex        =   157
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   5280
      TabIndex        =   156
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   5280
      TabIndex        =   155
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   5280
      TabIndex        =   154
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   153
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   152
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   151
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   150
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   149
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblCalFlowCol2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Calibrated Flow (slpm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      TabIndex        =   148
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMPrev 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   147
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   3840
      TabIndex        =   146
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   3840
      TabIndex        =   145
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   3840
      TabIndex        =   144
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   143
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   3840
      TabIndex        =   142
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   141
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   138
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   134
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   131
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   130
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCol1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% Diff Reading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      TabIndex        =   129
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblPercDiffReadCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   128
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   2640
      TabIndex        =   126
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   2640
      TabIndex        =   122
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   2640
      TabIndex        =   118
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   2640
      TabIndex        =   114
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   110
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   106
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   102
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   98
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   94
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   92
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblCalFlowCol1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calibrated Flow (slpm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2640
      TabIndex        =   91
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblCalFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   90
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1560
      TabIndex        =   89
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1560
      TabIndex        =   88
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1560
      TabIndex        =   87
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   86
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   85
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   84
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   83
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   82
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   81
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   80
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label lblActiualFlowCol2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual Flow (slpm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   79
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label lblActualFlowSLPMCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   78
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label lblMFCNomFlowCol2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MFC   Nom. Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   600
      TabIndex        =   77
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   600
      TabIndex        =   76
      Top             =   9600
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   0
      TabIndex        =   75
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   74
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   0
      TabIndex        =   73
      Top             =   9360
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   72
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   0
      TabIndex        =   71
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   70
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   0
      TabIndex        =   69
      Top             =   8880
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   68
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   0
      TabIndex        =   67
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   66
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   0
      TabIndex        =   65
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   64
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   63
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   62
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   61
      Top             =   7920
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   60
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   59
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   58
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   57
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label lblMFCPercFSCol2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MFC %F.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   56
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblMFCNomFlowCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   55
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCurr 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   54
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label lblMicropoise 
      BackStyle       =   0  'Transparent
      Caption         =   "Micropoise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   48
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblInchHg 
      BackStyle       =   0  'Transparent
      Caption         =   "Inches of Hg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   47
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblVisc 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   5160
      TabIndex        =   46
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblBaroPress 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "29.92"
      Height          =   255
      Left            =   5160
      TabIndex        =   45
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblViscTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Viscosity of Gas  ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   44
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblBaroPressTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Barometric Pressure  ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   43
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblStationTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12120
      TabIndex        =   6
      Top             =   9600
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   174
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   175
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   176
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   177
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   178
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   179
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   180
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   181
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   182
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   183
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblMFCNomFlowCol1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MFC   Nom. Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1695
      TabIndex        =   50
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblMFCPercFSCol1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MFC %F.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1080
      TabIndex        =   49
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "frmMassFlowCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''ERROR module 81
''
''frmMassFlowCal
''
Option Explicit

Const MAXTABLELENGTH = 11   ' The largest number of calibration points
Const MINTABLELENGTH = 3    ' The smallest number of calibration points

' Constants to indicate the form's selected calibration method
Const METHODLFE = 2
Const METHODFS = 1

' Results from text field validation
Const VALID = 0
Const TOOHIGH = 1
Const TOOLOW = 2
Const NOTNUMERIC = 3
Const NOTANINTEGER = 4

' Read-Only or Actually Performing a Calibration?  (default is Read-Only)
Private CalReadOnly As Boolean

' The loaded-in LFE coefficients
Private LFE_A(MAXMFC, 1 To MAX_STN) As Single
Private LFE_B(MAXMFC, 1 To MAX_STN) As Single
Private LFE_C(MAXMFC, 1 To MAX_STN) As Single
Private LFE_D(MAXMFC, 1 To MAX_STN) As Single
Private lfe_filename(MAXMFC, 1 To MAX_STN) As String
Private lfe_serialnum(MAXMFC, 1 To MAX_STN) As String

' The LFE coefficients for the currently displayed calibration
Private curr_lfe_a(MAXMFC, 1 To MAX_STN) As Single
Private curr_lfe_b(MAXMFC, 1 To MAX_STN) As Single
Private curr_lfe_c(MAXMFC, 1 To MAX_STN) As Single
Private curr_lfe_d(MAXMFC, 1 To MAX_STN) As Single
Private curr_lfe_filename(MAXMFC, 1 To MAX_STN) As String
Private curr_lfe_serialnum(MAXMFC, 1 To MAX_STN) As String

' The LFE coefficients for the previous calibration
Private prev_lfe_a(MAXMFC, 1 To MAX_STN) As Single
Private prev_lfe_b(MAXMFC, 1 To MAX_STN) As Single
Private prev_lfe_c(MAXMFC, 1 To MAX_STN) As Single
Private prev_lfe_d(MAXMFC, 1 To MAX_STN) As Single
Private prev_lfe_filename(MAXMFC, 1 To MAX_STN) As String
Private prev_lfe_serialnum(MAXMFC, 1 To MAX_STN) As String

' LFE coefficients for the buffered calibration
Private buffer_lfe_a(MAXMFC, 1 To MAX_STN) As Single
Private buffer_lfe_b(MAXMFC, 1 To MAX_STN) As Single
Private buffer_lfe_c(MAXMFC, 1 To MAX_STN) As Single
Private buffer_lfe_d(MAXMFC, 1 To MAX_STN) As Single
Private buffer_lfe_filename(MAXMFC, 1 To MAX_STN) As String
Private buffer_lfe_serialnum(MAXMFC, 1 To MAX_STN) As String


Private aryActualFlow(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' Actual Flow
Private aryBaroPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single   ' Barometric pressure

Private aryBufferActualFlow(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single  ' Actual flow values for the previous calibration
Private aryBufferBaroPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' Buffer for barometric pressure
Private aryBufferCalibDate(MAXMFC, 1 To MAX_STN) As Date   ' buffer Date of the previous calibration
Private aryBufferCalibrByText(MAXMFC, 1 To MAX_STN) As String ' "Calibrated by" text for the previous calibration
Private aryBufferCommentText(MAXMFC, 1 To MAX_STN) As String ' Comment text for the previous calibration
Private aryBufferLFEDiffPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' Buffer for LFE Differential Pressure
Private aryBufferLFEInletPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' buffer for LFE Inlet pressure
Private aryBufferLFEInletTemp(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' buffer for LFE Inlet temperature
Private aryBufferMFCMax(MAXMFC, 1 To MAX_STN) As Single   ' MFC range data for the previous calibration
Private aryBufferTableLength(MAXMFC, 1 To MAX_STN) As Integer   ' Number of points used in the previous calibration

Private aryCalibMethod(MAXMFC, 1 To MAX_STN) As Integer     ' Calibration methods
Private aryCalibrByText(MAXMFC, 1 To MAX_STN) As String     ' "Calibrated by" text on form
Private aryCalPointTable(MAXTABLELENGTH, MAXTABLELENGTH - 1) As Single    ' Hard-coded calibration points
Private aryCommentText(MAXMFC, 1 To MAX_STN) As String  ' Comment text on form

Private aryCurrActualFlow(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single  ' Actual flow for the current calibration
Private aryCurrBaroPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single   ' Barometric pressure for the current calibration
Private aryCurrCalibDate(MAXMFC, 1 To MAX_STN) As Date   ' Date of the current calibration
Private aryCurrCalibrByText(MAXMFC, 1 To MAX_STN) As String     ' "Calibrated by" text on form
Private aryCurrCalMFCMax(MAXMFC, 1 To MAX_STN) As Single    ' MFC range data for the current calibration
Private aryCurrCommentText(MAXMFC, 1 To MAX_STN) As String  ' Comment text on form
Private aryCurrLFEDiffPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' LFE Differential Pressure for the current calibration
Private aryCurrLFEInletPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single   ' LFE Inlet pressure for the current calibration
Private aryCurrLFEInletTemp(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' LFE Inlet Temperature for the current calibration
Private aryCurrTableLength(MAXMFC, 1 To MAX_STN) As Integer    ' Number of calibration points displayed for current and previous calibrations

Private aryLFEDiffPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single    ' Laminar Flow Element Differential Pressure
Private aryLFEInletPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single   ' Laminar Flow Element Inlet Pressure
Private aryLFEInletTemp(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single    ' Laminar Flow Element Inlet Temperature
Private aryMFCTableLength(MAXMFC, 1 To MAX_STN) As Integer  ' Number of calibration points for data entry

Private aryPrevActualFlow(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single  ' Actual flow values for the previous calibration
Private aryPrevBaroPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single   ' Barometric pressure for the previous calibration
Private aryPrevCalibDate(MAXMFC, 1 To MAX_STN) As Date   ' Date of the previous calibration
Private aryPrevCalibrByText(MAXMFC, 1 To MAX_STN) As String ' "Calibrated by" text for the previous calibration
Private aryPrevCommentText(MAXMFC, 1 To MAX_STN) As String ' Comment text for the previous calibration
Private aryPrevLFEDiffPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' LFE Diferential Pressure for the previous calibration
Private aryPrevLFEInletPress(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single   ' LFE Inlet pressure for the previous calibration
Private aryPrevLFEInletTemp(MAXMFC, 1 To MAX_STN, MAXTABLELENGTH - 1) As Single ' LFE Inlet Temperature for the previous calibration
Private aryPrevMFCMax(MAXMFC, 1 To MAX_STN) As Single   ' MFC range data for the previous calibration
Private aryPrevTableLength(MAXMFC, 1 To MAX_STN) As Integer   ' Number of points used in the previous calibration

' Least-Squares calibration coefficients for the buffered calibration
Private buffer_coefX(MAXMFC, 1 To MAX_STN) As Single
Private buffer_coefX2(MAXMFC, 1 To MAX_STN) As Single
Private buffer_coefX3(MAXMFC, 1 To MAX_STN) As Single
Private buffer_coefX4(MAXMFC, 1 To MAX_STN) As Single
Private buffer_coefX5(MAXMFC, 1 To MAX_STN) As Single
Private buffer_coefX6(MAXMFC, 1 To MAX_STN) As Single
Private buffer_coefR2(MAXMFC, 1 To MAX_STN) As Single

' Least-Squares calibration coeficients for the current calibration
Private curr_coefR2(MAXMFC, 1 To MAX_STN) As Single
Private curr_coefX(MAXMFC, 1 To MAX_STN) As Single
Private curr_coefX2(MAXMFC, 1 To MAX_STN) As Single
Private curr_coefX3(MAXMFC, 1 To MAX_STN) As Single
Private curr_coefX4(MAXMFC, 1 To MAX_STN) As Single
Private curr_coefX5(MAXMFC, 1 To MAX_STN) As Single
Private curr_coefX6(MAXMFC, 1 To MAX_STN) As Single

Private blnBufferExists(MAXMFC, 1 To MAX_STN) As Boolean   ' Flag - Whether buffer of previous calibration data is loaded
Private blnCurrCalExists(MAXMFC, 1 To MAX_STN) As Boolean   ' Flag - Whether current calibration data is loaded
Private blnPrevCalExists(MAXMFC, 1 To MAX_STN) As Boolean   ' Flag - Whether previous calibration data is loaded
Private blnUnsavedDisplacement(MAXMFC, 1 To MAX_STN) As Boolean   ' Flag - Whether an unsaved calibration exists

' Least-Squares calibration coefficients for previous calibrations
Private prev_coefX(MAXMFC, 1 To MAX_STN) As Single
Private prev_coefX2(MAXMFC, 1 To MAX_STN) As Single
Private prev_coefX3(MAXMFC, 1 To MAX_STN) As Single
Private prev_coefX4(MAXMFC, 1 To MAX_STN) As Single
Private prev_coefX5(MAXMFC, 1 To MAX_STN) As Single
Private prev_coefX6(MAXMFC, 1 To MAX_STN) As Single
Private prev_coefR2(MAXMFC, 1 To MAX_STN) As Single

Private Result As Integer                       ' Result of text box validation test
Private SelectedMFC(1 To MAX_STN) As Integer    ' The Mass Flow Controller currently selected
Private selectedRow As Integer                  ' The currently selected calibration date entry row
Public SelectedStation As Integer               ' The currently selected station
Public MFCInUse As Integer                      ' The currently selected mfc
Public SelectedMFCMax As Single
Public LFE_A_InUse As Single
Public LFE_B_InUse As Single
Public LFE_C_InUse As Single
Public LFE_D_InUse As Single
Public CurrBaroPress As Single                  ' Current barometric pressure
Public Viscosity As Single                      ' The viscosity


Private Sub InitLFE()
    ' fill the LFE variables with values
    ' Note: This was written as a function so that
    ' it could be replaced by another initializing method
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 81, 1
Dim i As Integer
Dim j As Integer
    For i = 0 To MAXMFC
        For j = 1 To LAST_STN
            LFE_A(i, j) = 0!
            LFE_B(i, j) = 0.011676!
            LFE_C(i, j) = -1.6879E-04!
            LFE_D(i, j) = 0!
            lfe_filename(i, j) = "*.LFE"
            lfe_serialnum(i, j) = "None"
        Next j
    Next i
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

Private Sub FillCalPointTable()
    ' Define every MFC calibration point
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 81, 1
    
    aryCalPointTable(2, 0) = 0
    aryCalPointTable(2, 1) = 100
    
    aryCalPointTable(3, 0) = 0
    aryCalPointTable(3, 1) = 50
    aryCalPointTable(3, 2) = 100
    
    aryCalPointTable(4, 0) = 0
    aryCalPointTable(4, 1) = 30
    aryCalPointTable(4, 2) = 70
    aryCalPointTable(4, 3) = 100
    
    aryCalPointTable(5, 0) = 0
    aryCalPointTable(5, 1) = 25
    aryCalPointTable(5, 2) = 50
    aryCalPointTable(5, 3) = 75
    aryCalPointTable(5, 4) = 100
    
    aryCalPointTable(6, 0) = 0
    aryCalPointTable(6, 1) = 20
    aryCalPointTable(6, 2) = 40
    aryCalPointTable(6, 3) = 60
    aryCalPointTable(6, 4) = 80
    aryCalPointTable(6, 5) = 100
    
    aryCalPointTable(7, 0) = 0
    aryCalPointTable(7, 1) = 16.7
    aryCalPointTable(7, 2) = 33.3
    aryCalPointTable(7, 3) = 50
    aryCalPointTable(7, 4) = 66.7
    aryCalPointTable(7, 5) = 83.4
    aryCalPointTable(7, 6) = 100
    
    aryCalPointTable(8, 0) = 0
    aryCalPointTable(8, 1) = 14.3
    aryCalPointTable(8, 2) = 28.6
    aryCalPointTable(8, 3) = 42.9
    aryCalPointTable(8, 4) = 57.2
    aryCalPointTable(8, 5) = 71.5
    aryCalPointTable(8, 6) = 85.8
    aryCalPointTable(8, 7) = 100
    
    aryCalPointTable(9, 0) = 0
    aryCalPointTable(9, 1) = 12.5
    aryCalPointTable(9, 2) = 25
    aryCalPointTable(9, 3) = 37.5
    aryCalPointTable(9, 4) = 50
    aryCalPointTable(9, 5) = 62.5
    aryCalPointTable(9, 6) = 75
    aryCalPointTable(9, 7) = 87.5
    aryCalPointTable(9, 8) = 100
    
    aryCalPointTable(10, 0) = 0
    aryCalPointTable(10, 1) = 11.1
    aryCalPointTable(10, 2) = 22.2
    aryCalPointTable(10, 3) = 33.3
    aryCalPointTable(10, 4) = 44.4
    aryCalPointTable(10, 5) = 55.5
    aryCalPointTable(10, 6) = 66.6
    aryCalPointTable(10, 7) = 77.7
    aryCalPointTable(10, 8) = 88.8
    aryCalPointTable(10, 9) = 100
    
    aryCalPointTable(11, 0) = 0
    aryCalPointTable(11, 1) = 10
    aryCalPointTable(11, 2) = 20
    aryCalPointTable(11, 3) = 30
    aryCalPointTable(11, 4) = 40
    aryCalPointTable(11, 5) = 50
    aryCalPointTable(11, 6) = 60
    aryCalPointTable(11, 7) = 70
    aryCalPointTable(11, 8) = 80
    aryCalPointTable(11, 9) = 90
    aryCalPointTable(11, 10) = 100
    
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

Private Function NumMFCRows() As Integer
    ' Returns the number of calibration points for the currently selected MFC in the current station
    NumMFCRows = aryMFCTableLength(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetNumMFCRows(NUMROWS As Integer)
    ' Sets the number of calibration points for the currently selected MFC in the current station
    aryMFCTableLength(SelectedMFC(SelectedStation), SelectedStation) = NUMROWS
End Sub

Private Function NumCurrMFCRows() As Integer
    ' Returns the number of current calibration points for the currently selected MFC in the current station
    NumCurrMFCRows = aryCurrTableLength(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetNumCurrMFCRows(NUMROWS As Integer)
    ' Sets the number of current calibration points for the currently selected MFC in the current station
    aryCurrTableLength(SelectedMFC(SelectedStation), SelectedStation) = NUMROWS
End Sub

Private Function NumPrevMFCRows() As Integer
    ' Returns the number of previous calibration points for the currently selected MFC in the current station
    NumPrevMFCRows = aryPrevTableLength(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetNumPrevMFCRows(NUMROWS As Integer)
    ' Sets the number of previous calibration points for the currently selected MFC in the current station
    aryPrevTableLength(SelectedMFC(SelectedStation), SelectedStation) = NUMROWS
End Sub

Private Function NumBufferMFCRows() As Integer
    ' Returns the number of calibration points in the buffer for the currently selected MFC in the current station
    NumBufferMFCRows = aryBufferTableLength(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetNumBufferMFCRows(NUMROWS As Integer)
    ' Sets the number of calibration points in the buffer for the currently selected MFC in the current station
    aryBufferTableLength(SelectedMFC(SelectedStation), SelectedStation) = NUMROWS
End Sub

Private Function CommentText() As String
    ' Returns the comment text for the current station / mfc
    CommentText = aryCommentText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetCommentText(TheText As String)
    ' Sets the text for the comment for the current station / mfc
    aryCommentText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function CurrCommentText() As String
    ' Returns the current comment text for the current station / mfc
    CurrCommentText = aryCurrCommentText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetCurrCommentText(TheText As String)
    ' Sets the text for the current comment for the current station / mfc
    aryCurrCommentText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function PrevCommentText() As String
    ' Returns the previous comment text for the current station / mfc
    PrevCommentText = aryPrevCommentText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetPrevCommentText(TheText As String)
    ' Sets the text for the previous comment for the current station / mfc
    aryPrevCommentText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function BufferCommentText() As String
    ' Returns the buffer comment text for the current station / mfc
    BufferCommentText = aryBufferCommentText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetBufferCommentText(TheText As String)
    ' Sets the text for the buffered comment for the current station / mfc
    aryBufferCommentText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function CalibrByText() As String
    ' Returns the "calibrated by" text for the current station / mfc
    CalibrByText = aryCalibrByText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetCalibrByText(TheText As String)
    ' Sets the text for the "calibrated by" statement for the current station / mfc
    aryCalibrByText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function CurrCalibrByText() As String
    ' Returns the current "calibrated by" text for the current station / mfc
    CurrCalibrByText = aryCurrCalibrByText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetCurrCalibrByText(TheText As String)
    ' Sets the text for the current "calibrated by" statement for the current station / mfc
    aryCurrCalibrByText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function PrevCalibrByText() As String
    ' Returns the previous "calibrated by" text for the current station / mfc
    PrevCalibrByText = aryPrevCalibrByText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetPrevCalibrByText(TheText As String)
    ' Sets the text for the previous "calibrated by" statement for the current station / mfc
    aryPrevCalibrByText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function BufferCalibrByText() As String
    ' Returns the buffered "calibrated by" text for the current station / mfc
    BufferCalibrByText = aryBufferCalibrByText(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetBufferCalibrByText(TheText As String)
    ' Sets the text for the buffered "calibrated by" statement for the current station / mfc
    aryBufferCalibrByText(SelectedMFC(SelectedStation), SelectedStation) = TheText
End Sub

Private Function CalibMethod() As Integer
    ' Returns the calibration method for the current station / mfc
    CalibMethod = aryCalibMethod(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetCalibMethod(Method As Integer)
    ' Sets the text for the "calibrated by" statement for the current station / mfc
    aryCalibMethod(SelectedMFC(SelectedStation), SelectedStation) = Method
End Sub

Private Function CurrCalMFCMax() As Single
    ' Returns the maximum value for the MFC when the current calibration was taken
    CurrCalMFCMax = aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation)
End Function

Private Sub SetCurrCalMFCMax(Value As Single)
    ' Sets the maximum value for the MFC for the current calibration table
    aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Private Function CurrCalExists() As Boolean
    ' Returns true if a current calibration
    ' is noted for this station / mfc
    If blnCurrCalExists(SelectedMFC(SelectedStation), SelectedStation) = True Then
        CurrCalExists = True
    Else
        CurrCalExists = False
    End If
End Function

Private Sub SetCurrCalExists(flag As Boolean)
    ' Note that a current calibration does or does not exists for this station / mfc
    If flag = True Then
        blnCurrCalExists(SelectedMFC(SelectedStation), SelectedStation) = True
'        cmdSaveNewCalib.Enabled = True
        cmdPrint.Enabled = True
    Else
        blnCurrCalExists(SelectedMFC(SelectedStation), SelectedStation) = False
'        cmdSaveNewCalib.Enabled = False
        cmdPrint.Enabled = False
    End If
End Sub

Private Function PrevCalExists() As Boolean
    ' Returns true if a previous calibration
    ' is noted for this station / mfc
    If blnPrevCalExists(SelectedMFC(SelectedStation), SelectedStation) = True Then
        PrevCalExists = True
    Else
        PrevCalExists = False
    End If
End Function

Private Sub SetPrevCalExists(flag As Boolean)
    ' Note that a previous calibration does or does not exists for this station / mfc
    If flag = True Then
        blnPrevCalExists(SelectedMFC(SelectedStation), SelectedStation) = True
    Else
        blnPrevCalExists(SelectedMFC(SelectedStation), SelectedStation) = False
    End If
End Sub

Private Sub cmdCalcActualFlow_Click()
    ' Validates each LFE input box in the MFC Calibration table to
    ' make sure they they all contain numbers.
    ' And boxes that don't pass validation are marked in yellow.
    ' If any boxes don't pass validation, a message will appear.
    ' Otherwise, the actual flow values will be calculated and displayed
    ' on the form
    Dim row As Integer
    Dim ErrorExists As Boolean
    
    ErrorExists = False
    
    For row = 1 To NumMFCRows
        If ValidateBox(txtLFEDiffPress(row - 1)) = False Then ErrorExists = True
        If ValidateBox(txtLFEInletPress(row - 1)) = False Then ErrorExists = True
        If ValidateBox(txtInletTemp(row - 1)) = False Then ErrorExists = True
    Next row
    
    If ErrorExists = True Then
'        Delay_Box "Data entry boxes need numeric characters.", MSGDELAY, msgSHOW
        txtMsg.ForeColor = MEDRED
        txtMsg.text = "Data entry boxes need numeric characters."
    Else
'        Delay_Box "Actual Flow Calculated", MSGDELAY, msgSHOW
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = "Actual Flow Calculated"
        CalcActualFlow
        DisplMFCData
    End If
End Sub

Private Sub cmdCalCheckFunction_Click()
    ' Open the check calibration form
    LFE_A_InUse = LFE_A(SelectedMFC(SelectedStation), SelectedStation)
    LFE_B_InUse = LFE_B(SelectedMFC(SelectedStation), SelectedStation)
    LFE_C_InUse = LFE_C(SelectedMFC(SelectedStation), SelectedStation)
    LFE_D_InUse = LFE_D(SelectedMFC(SelectedStation), SelectedStation)
    
    MFCInUse = SelectedMFC(SelectedStation)
    frmCalCheck.Show

End Sub

Private Sub cmdCalibrate_Click()
   
    Dim row As Integer
    Dim ErrorExists As Boolean
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 81, 3
    
    ' Validate the actual flow boxes
    ErrorExists = False
    For row = 1 To NumMFCRows
        If ValidateBox(txtActualFlowSLPM(row - 1)) = False Then ErrorExists = True
    Next row
    
    If ErrorExists = True Then
'        Delay_Box "All Actual Flow boxes need numeric characters.", MSGDELAY, msgSHOW
        txtMsg.ForeColor = MEDRED
        txtMsg.text = "All Actual Flow boxes need numeric characters."
    Else
   
        ' the calibrate button was clicked
        ' Transfer current calibration values
        If PrevCalExists = True Then
            TransferPrevToBuff
        End If
    
'        If CurrCalExists = True Then
        TransferCurrToPrev
'        End If
    
        ' Store Calibration information in memory
        SetNumCurrMFCRows (NumMFCRows)
        SetCurrCalibrByText CalibrByText
        SetCurrCommentText CommentText
    
        SetCurrCalMFCMax (SelectedMFCMax)
    
        If PrevCalExists = True Then
            ShowPrevCalTable
            HidePrevInactiveTableRows
            DisplPrevCalResults
        End If
    
        ' Store the table values
        StoreCurrValues
    
        ' Copy the LFE coefficients
        curr_lfe_a(SelectedMFC(SelectedStation), SelectedStation) = LFE_A(SelectedMFC(SelectedStation), SelectedStation)
        curr_lfe_b(SelectedMFC(SelectedStation), SelectedStation) = LFE_B(SelectedMFC(SelectedStation), SelectedStation)
        curr_lfe_c(SelectedMFC(SelectedStation), SelectedStation) = LFE_C(SelectedMFC(SelectedStation), SelectedStation)
        curr_lfe_d(SelectedMFC(SelectedStation), SelectedStation) = LFE_D(SelectedMFC(SelectedStation), SelectedStation)
        curr_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) = lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
        curr_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation) = lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
    
  
        ' Store the date and time of the calibration
        aryCurrCalibDate(SelectedMFC(SelectedStation), SelectedStation) = Now
        
        ' Get the calibration coefficients
        Calibrate
    
        ' Display the previous calibration button if one exits
        If PrevCalExists = True Then
            blnUnsavedDisplacement(SelectedMFC(SelectedStation), SelectedStation) = True
            cmdKeepPrev.Enabled = True
        End If
    
        ' Display the graph
        ChartValues
    
        ' display the current calibration table
        ShowCurrCalTable
        DisplCalResults
        UpdateFormDisplay
'        Delay_Box "Calibrate Done", MSGDELAY, msgSHOW
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = "Calibrate Done"
    End If
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

Private Sub Calibrate()
' Calls an Excel funcion to get calibration coefficients
' for the data entered in the main table

' Enter data from the form into the worksheet
' The Excel workbook
' (Used for Least-Squares calculations)
Dim xlApp As Object
Dim xlWB As Object
Dim xlSht As Object

Dim i, RangeOffset As Integer
Dim InputDataRange, FormulaRange As Range
Dim InputRangeText, XRangeText, YRangeText As String
Dim LINESTFormulaText As String

    RangeOffset = NumMFCRows - 1
    
    ' open Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    ' open the empty Microsoft Excel workbook used for the calibration calculations
    Set xlWB = xlApp.Workbooks.Open(FILEPATH_cal & "\calibration.xls")
    DoEvents

    ' set activesheet
    xlWB.Sheets(1).Select
    xlWB.Sheets(1).Activate
    Set xlSht = xlWB.Sheets(1)
    
    For i = 0 To RangeOffset
        xlSht.Range("A" & i + 2) = (aryCalPointTable(NumMFCRows, i) / 100)
        xlSht.Range("B" & i + 2) = (aryCalPointTable(NumMFCRows, i) / 100) ^ 2
        xlSht.Range("C" & i + 2) = (aryCalPointTable(NumMFCRows, i) / 100) ^ 3
        xlSht.Range("D" & i + 2) = (aryCalPointTable(NumMFCRows, i) / 100) ^ 4
        xlSht.Range("E" & i + 2) = (aryCalPointTable(NumMFCRows, i) / 100) ^ 5
        xlSht.Range("F" & i + 2) = (aryCalPointTable(NumMFCRows, i) / 100) ^ 6
        xlSht.Range("G" & i + 2) = CSng(txtActualFlowSLPM(i))
    Next i
    
    ' Select cells for the array formula
    Set FormulaRange = xlSht.Range("A14:F18")
    ' FormulaRange.Select
    
    ' Enter the LINEST formula
    InputRangeText = "A2:G" & RangeOffset + 2
    XRangeText = "R2C1:R" & RangeOffset + 2 & "C6"
    YRangeText = "R2C7:R" & RangeOffset + 2 & "C7"
    Set InputDataRange = xlSht.Range(InputRangeText)
    LINESTFormulaText = "=LINEST(" & YRangeText & "," & XRangeText & ", FALSE, TRUE)"
    FormulaRange.FormulaArray = LINESTFormulaText
    FormulaRange.Calculate
    'MsgBox xlWB.Sheets(1).Range("A14").value
    
    ' Assign the coefficients to variables
    curr_coefX(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("F14").Value
    curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("E14").Value
    curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("D14").Value
    curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("C14").Value
    curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("B14").Value
    curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("A14").Value
    
    curr_coefR2(SelectedMFC(SelectedStation), SelectedStation) = xlSht.Range("A16").Value

    ' Save as a file to examine later
'    xlWB.SaveAs FILEPATH_cal & "debugCal.xls"      ' This is for testing ******************
    
    
    ' Note that a current calibration now exists
    SetCurrCalExists True

    ' make sure xlWB is closed
    xlWB.Saved = True
    xlWB.Close
    Set xlWB = Nothing
    ' Close Excel
    xlApp.Quit
    Set xlApp = Nothing
    
End Sub

Private Function DisplCalFormula()
    ' Displays the calibration formula on the screen
    Dim FormulaText As String
    FormulaText = "Y = " & curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) & "X6"
    FormulaText = FormulaText & IIf(curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)) & "X5"
    FormulaText = FormulaText & IIf(curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)) & "X4"
    FormulaText = FormulaText & IIf(curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)) & "X3"
    FormulaText = FormulaText & IIf(curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)) & "X2"
    FormulaText = FormulaText & IIf(curr_coefX(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX(SelectedMFC(SelectedStation), SelectedStation)) & "X"
    FormulaText = FormulaText & "      R2=" & curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)
    frmCalFormula.Visible = True
    lblCalibRslt.Visible = True
    lblCalibRslt.Caption = FormulaText
End Function

Private Function PrevCalibratedValue(InputVal As Single) As Single
    ' The inputs value should be (percent full scale) / 100
    
    ' Returns the calibrated value in SLPM
    Dim tempdbl As Double
    tempdbl = prev_coefX(SelectedMFC(SelectedStation), SelectedStation) * InputVal
    tempdbl = tempdbl + prev_coefX2(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 2)
    tempdbl = tempdbl + prev_coefX3(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 3)
    tempdbl = tempdbl + prev_coefX4(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 4)
    tempdbl = tempdbl + prev_coefX5(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 5)
    tempdbl = tempdbl + prev_coefX6(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 6)
    
    PrevCalibratedValue = CSng(tempdbl)
End Function

Private Sub StoreCurrValues()
    ' Stores the actual flow data from the form to an array
    Dim row As Integer
    For row = 1 To NumCurrMFCRows
        aryCurrLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
End Sub

Private Sub ChartValues()
    Dim ChartArray(MAXTABLELENGTH, 2)
    Dim i As Integer

    Dim Graph() As Single
         
    ' Note: the number of points displayed on the graph is the number of elements
    ' allocated in the first dimension of the Graph array
    ReDim Graph(NumCurrMFCRows, 1 To 6)
    'Dim i As Integer
    For i = 1 To NumCurrMFCRows
        Graph(i, 1) = aryCalPointTable(NumCurrMFCRows, i - 1)  ' value for X-axis
        Graph(i, 2) = CSng(CurrCalMFCMax * aryCalPointTable(NumCurrMFCRows, i - 1) / 100)
        Graph(i, 3) = aryCalPointTable(NumCurrMFCRows, i - 1)  ' value for X-axis
        Graph(i, 4) = CSng(aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, i - 1))
        Graph(i, 5) = aryCalPointTable(NumCurrMFCRows, i - 1)  ' value for X-axis
        Graph(i, 6) = CalibratedValue(aryCalPointTable(NumCurrMFCRows, i - 1) / 100) ' value for Y-axis
    Next i
         
    chtMfcChart.chartType = VtChChartType2dXY  ' set to X Y Scatter chart
    chtMfcChart = Graph ' populate chart's data grid using Graph array
    ' Leave the following line commented until step 6:
    chtMfcChart.Plot.UniformAxis = False
    chtMfcChart.Column = 1
    chtMfcChart.ColumnLabel = "MFC Nom. Flow"
    chtMfcChart.Column = 3
    chtMfcChart.ColumnLabel = "Actual Flow (slpm)"
    chtMfcChart.Column = 5
    chtMfcChart.ColumnLabel = "Calib. Flow (slpm)"
    chtMfcChart.Visible = True
    ' MsgBox chtMFCChart.Title.Font.Bold
    lblCalGraphTitle.Visible = True
    lblCalGraphRow.Visible = True

    ' Display the calibration formula
    DisplCalFormula
End Sub

Private Sub TransferCurrToPrev()
    ' Transfer data from the current calibration to the previous calibration
    Dim row As Integer
    For row = 1 To NumCurrMFCRows
        ' Copy actual flows
        aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryCurrLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryCurrLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryCurrLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryCurrBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
        
    ' Copy calibration coefficients
    prev_coefX(SelectedMFC(SelectedStation), SelectedStation) = curr_coefX(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX2(SelectedMFC(SelectedStation), SelectedStation) = curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX3(SelectedMFC(SelectedStation), SelectedStation) = curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX4(SelectedMFC(SelectedStation), SelectedStation) = curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX5(SelectedMFC(SelectedStation), SelectedStation) = curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX6(SelectedMFC(SelectedStation), SelectedStation) = curr_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefR2(SelectedMFC(SelectedStation), SelectedStation) = curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)

    ' Copy the LFE coefficients
    prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation) = curr_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation) = curr_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation) = curr_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation) = curr_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) = curr_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation) = curr_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
    
    ' store number of points (rows)
    SetNumPrevMFCRows NumCurrMFCRows
    
    ' store maximum flow value
    aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation) = aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation)
  
    ' transfer the date values
    aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation) = aryCurrCalibDate(SelectedMFC(SelectedStation), SelectedStation)
    
    SetPrevCalibrByText CurrCalibrByText
    SetPrevCommentText CurrCommentText
    
    SetPrevCalExists True
End Sub

Private Sub TransferPrevToCurr()
    ' Transfer data from the current calibration to the previous calibration
    Dim row As Integer
    For row = 1 To NumCurrMFCRows
        ' Copy actual flows
        aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryCurrBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
    
    ' Copy calibration coefficients
    curr_coefX(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX(SelectedMFC(SelectedStation), SelectedStation)
    curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    curr_coefR2(SelectedMFC(SelectedStation), SelectedStation) = prev_coefR2(SelectedMFC(SelectedStation), SelectedStation)

    ' Copy the LFE coefficients
    curr_lfe_a(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    curr_lfe_b(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    curr_lfe_c(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    curr_lfe_d(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    curr_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
    curr_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)

    
    ' store number of points (rows)
    SetNumCurrMFCRows NumPrevMFCRows
    
    ' store maximum flow value
    aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation) = aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation)
  
    ' transfer the date values
    aryCurrCalibDate(SelectedMFC(SelectedStation), SelectedStation) = aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation)
    
    SetCurrCalibrByText PrevCalibrByText
    SetCurrCommentText PrevCommentText

End Sub

Private Sub TransferPrevToBuff()
    ' Transfer data from the previous calibration to the calibration buffer
    Dim row As Integer
    For row = 1 To aryPrevTableLength(SelectedMFC(SelectedStation), SelectedStation)
        ' Copy actual flows
        aryBufferActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryBufferLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryBufferLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryBufferLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryBufferBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
    
    ' Copy calibration coefficients
    buffer_coefX(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX(SelectedMFC(SelectedStation), SelectedStation)
    buffer_coefX2(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    buffer_coefX3(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    buffer_coefX4(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    buffer_coefX5(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    buffer_coefX6(SelectedMFC(SelectedStation), SelectedStation) = prev_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    buffer_coefR2(SelectedMFC(SelectedStation), SelectedStation) = prev_coefR2(SelectedMFC(SelectedStation), SelectedStation)

    ' Copy the LFE coefficients
    buffer_lfe_a(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    buffer_lfe_b(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    buffer_lfe_c(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    buffer_lfe_d(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    buffer_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
    buffer_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation) = prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)

    ' store number of points (rows)
    SetNumBufferMFCRows NumPrevMFCRows
    
    ' store maximum flow value
    aryBufferMFCMax(SelectedMFC(SelectedStation), SelectedStation) = aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation)
  
    ' transfer the date values
    aryBufferCalibDate(SelectedMFC(SelectedStation), SelectedStation) = aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation)
    
    SetBufferCalibrByText PrevCalibrByText
    SetBufferCommentText PrevCommentText
    
    blnBufferExists(SelectedMFC(SelectedStation), SelectedStation) = True
End Sub

Private Sub TransferBuffToPrev()
    ' Transfer data from the current calibration to the previous calibration
    Dim row As Integer
    For row = 1 To aryBufferTableLength(SelectedMFC(SelectedStation), SelectedStation)
        ' Copy actual flows
        aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryBufferLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryBufferLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryBufferLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1) = aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
    
    ' Copy calibration coefficients
    prev_coefX(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefX(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX2(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX3(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX4(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX5(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefX6(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    prev_coefR2(SelectedMFC(SelectedStation), SelectedStation) = buffer_coefR2(SelectedMFC(SelectedStation), SelectedStation)

    ' Copy the LFE coefficients
    prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation) = buffer_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation) = buffer_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation) = buffer_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation) = buffer_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) = buffer_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
    prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation) = buffer_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)

    ' store number of points (rows)
    SetNumPrevMFCRows NumBufferMFCRows
    
    ' store maximum flow value
    aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation) = aryBufferMFCMax(SelectedMFC(SelectedStation), SelectedStation)
  
    ' transfer the date values
    aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation) = aryBufferCalibDate(SelectedMFC(SelectedStation), SelectedStation)
    
    SetPrevCalibrByText BufferCalibrByText
    SetPrevCommentText BufferCommentText
 
    ' The buffer is no longer useful
    blnBufferExists(SelectedMFC(SelectedStation), SelectedStation) = False
    
    SetPrevCalExists True
End Sub

Private Sub cmdCalPtDown_Click()
    ' Decrement the number of calibration points
     SubtractCalPt
End Sub

Private Sub cmdCalPtUp_Click()
    ' Increment the number of calibration points
    AddCalPt
End Sub

Private Sub cmdDown_Click()
    ' This command increments the station number variable,
    ' the station number displayed, and triggers an update for
    ' the values displayed on the form for the current station
    
    ' Turn off valves of old station
    If Not CalReadOnly Then
        Close_Stn_Valves SelectedStation, 1
        PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = False
    End If
     
    SetCurrCalExists False
    If SelectedStation > 1 Then
        SelectedStation = SelectedStation - 1
        txtDispStn.Caption = SelectedStation
    Else
        SelectedStation = LAST_STN
        txtDispStn.Caption = SelectedStation
    End If
    
    ' Turn off valves for new station
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
     
    ResetMFCSelection
    UpdateMFCMode
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub cmdExit_Click()
    Xit
End Sub

Private Sub cmdKeepPrev_Click()
    ' Undoes the effect of the calibration by transferring data from
    ' the previous calibration to th current calibration, transferring data
    ' from the temporary calibration buffer to the previous calibration,
    ' and updating the form
    
    If blnPrevCalExists(SelectedMFC(SelectedStation), SelectedStation) = True Then
        TransferPrevToCurr
    End If
    
    If blnBufferExists(SelectedMFC(SelectedStation), SelectedStation) = True Then
        TransferBuffToPrev
    End If
    blnUnsavedDisplacement(SelectedMFC(SelectedStation), SelectedStation) = False
    SetCommentText CurrCommentText
    SetCalibrByText CurrCalibrByText
    UpdateFormDisplay
'    Delay_Box "Returning to Previously Saved Cal.", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "Returning to Previously Saved Cal."
End Sub

Private Sub cmdMFCDown_Click()
    ' Selects the row below the currently selected row in the
    ' MFC calibration table, for editing
    
    If selectedRow >= NumMFCRows Then Exit Sub
    
    selectedRow = selectedRow + 1
    ' Turn On Valves for selected MFC
    UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub cmdMFCUp_Click()
    ' Selects the row above the currently selected row in the
    ' MFC calibration table, for editing
    
    If selectedRow <= 1 Then Exit Sub
    
    selectedRow = selectedRow - 1
    ' Turn On Valves for selected MFC
    UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub cmdNewLFEFile_Click()
    ' Open the LFE definition screen
    Me.Enabled = False
    frmOpenLFE.Show
    frmOpenLFE.cmdSave.Enabled = False
    frmOpenLFE.cmdSaveAs.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    PrintCalibration
'    Delay_Box "File Released to the Printer", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "File Released to the Printer"
End Sub

Private Sub cmdReadOnly_Click()
Dim count, station, Shift As Integer
    If CalReadOnly Then
        For station = 1 To LAST_STN
            For Shift = 1 To NR_SHIFT
                If StationControl(station, Shift).Mode <> VBIDLE Then
                    count = count + 1
                    ' Report an error
                    txtMsg.ForeColor = MEDRED
                    txtMsg.text = "Station:" & station & "  is still running in Shift " & Shift
                End If
            Next Shift
        Next station
        If count = 0 Then
            CalReadOnly = False
            txtMsg.ForeColor = DKPURPLE
            txtMsg.text = "Calibration is No Longer Read-Only"
            Close_Stn_Valves SelectedStation, 1
            UpdateMfcSelection
            UpdateFormDisplay
        End If
    Else
        CalReadOnly = True
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = "Calibration is Read-Only"
        Close_Stn_Valves SelectedStation, 1
        UpdateFormDisplay
    End If
End Sub

Private Sub cmdSaveNewCalib_Click()
    ' Save the new calibration
    SaveCalibData
    blnUnsavedDisplacement(SelectedMFC(SelectedStation), SelectedStation) = False
'    Delay_Box "Calibration Saved", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "Calibration Saved"
    ' Read it back for regular use
    '     note: Subs ReadCalibData & LoadMfcCalibration
    '           read the same file
    '           but read it into different variables
    '           ReadCalibData is just for Calibration
    '           LoadMfcCalibration is for everyone else
    LoadMfcCalibration SelectedStation
    UpdateFormDisplay
End Sub

Private Sub cmdSetToNominal_Click()
    Dim row As Integer
    Dim slpm As Single
    
    ' for each row in the MFC calibration table
    For row = 1 To NumMFCRows
    
        ' Actual Flow = Nominal Flow
        slpm = CSng(lblMFCNomFlow(row - 1).Caption)

        aryActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = slpm
    Next row
        
    txtCalibBy.text = "default"
    txtComments.text = "Actual Flows set, by default, to the Nominal Values"
    
'    Delay_Box "Actual Flow set to Nominal Flow", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "Actual Flow set to Nominal Flow"
    DisplMFCData
End Sub

Private Sub cmdUp_Click()
    ' This command increments the station number variable,
    ' the station number displayed, and triggers an update for
    ' the values displayed on the form for the current station
    
    ' Turn off valves of old station
    If Not CalReadOnly Then
        Close_Stn_Valves SelectedStation, 1
        PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = False
    End If
     
    SetCurrCalExists False
    If SelectedStation < LAST_STN Then
        SelectedStation = SelectedStation + 1
        txtDispStn.Caption = SelectedStation
    Else
        SelectedStation = 1
        txtDispStn.Caption = SelectedStation
    End If
    
    ' Turn off valves of new station
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
     
    ResetMFCSelection
    UpdateMFCMode
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim k As Integer
    fraMFCSelection.ForeColor = Titles_ForeColor
    fraCalibMethod.ForeColor = Titles_ForeColor
    fraCalibInfo.ForeColor = Titles_ForeColor
    txtNumCalPts.ForeColor = TitlesData_Forecolor
    lblCalTableTitle.ForeColor = Titles_ForeColor
    lblCurrCalTitle.ForeColor = Titles_ForeColor
    lblPrevCalTitle.ForeColor = Titles_ForeColor
    txtDispStn.ForeColor = TitlesData_Forecolor
    lblCalGraphTitle.ForeColor = Titles_ForeColor

    CalReadOnly = True
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "Calibration is Read-Only"
    CalibrateActive = True

    ' Set the current station number  1
    SelectedStation = 1
    txtDispStn.Caption = SelectedStation
    
    ' Display Baro in mBar
    CurrBaroPress = AmbBaro
    lblBaroPress.Caption = Format(CurrBaroPress, "#000")
    ' Initialize calibration points
    FillCalPointTable
    ManualMode
    
    InitLFE
    ' Sets the number of calibration points to 5 for each MFC in each station
    For i = 1 To LAST_STN
        SelectedMFC(i) = MFCPURGEAIR
        For j = 0 To MAXMFC
            aryCalibMethod(j, i) = METHODFS
            aryMFCTableLength(j, i) = 11 ' 5
            aryCurrTableLength(j, i) = 7
            aryCalibrByText(j, i) = ""
            aryCommentText(j, i) = ""
            For k = 0 To MAXTABLELENGTH - 1
                aryBaroPress(j, i, k) = CurrBaroPress
                aryActualFlow(j, i, k) = 0#
            Next k
        Next j
    Next i
    
    ResetMFCSelection
    SetPurgeAirMode
    SetFSMode
    DisplMFCScale
    DisplMFCData
    DisplCurrScale
    UpdateCalPtLabel
    HideInactiveTableRows
    HideCurrInactiveTableRows
    HidePrevInactiveTableRows
    
    ' Turn Off All Valves
    If Not CalReadOnly Then Reset_Valves
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub optButane_Click()
    ' The butane radio button was selected
    SetButaneMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub SetButaneMode()
    ' Sets the form to the Butane mode
    SelectedMFC(SelectedStation) = MFCBUTANE
    SelectedMFCMax = Stn_AIO(SelectedStation, asButaneFlow).EuMax
    Viscosity = 74
    SetCurrCalExists False
    optButane.Value = True
End Sub

Private Sub optCalibTable_Click(Index As Integer)
    ' Selects the row that was clicked on, for editing
    selectedRow = Index + 1
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub optFlowStandard_Click()
    ' The flow standard radio button was selected
    SetFSMode
'    UpdateMFCSelection
    UpdateFormDisplay
End Sub

Private Sub optLaminarFlowElement_Click()
    ' The Laminar Flow Element (LFE) button was selected
    SetLFEMode
'    UpdateMFCSelection
    UpdateFormDisplay
End Sub

Private Sub optLiveFuel_Click()
    ' The Live Fuel radio button was selected
    SetLiveFuelMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub SetLiveFuelMode()
    ' Sets the form to the Live Fuel mode
    SelectedMFC(SelectedStation) = MFCLIVEFUEL
    SelectedMFCMax = Stn_AIO(SelectedStation, asLiveFuelVaporFlow).EuMax
    Viscosity = 181.87
    SetCurrCalExists False
    optLiveFuel.Value = True
End Sub

Private Sub optNitrogen_Click()
    ' the Nitrogen radio button was selected
    SetNitrogenMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub SetNitrogenMode()
    ' Sets the form to the Nitrogen mode
    SelectedMFC(SelectedStation) = MFCNITROGEN
    SelectedMFCMax = Stn_AIO(SelectedStation, asNitrogenFlow).EuMax
    Viscosity = 174
    SetCurrCalExists False
    optNitrogen.Value = True
End Sub

Private Sub UpdateViscLabel()
    ' Sets the viscosity display to the current value
    lblVisc.Caption = Viscosity
End Sub

Private Sub SetORVRButMode()
    ' Sets the form to the ORVR - Butane mode
    SelectedMFC(SelectedStation) = MFCORVRBUT
    SelectedMFCMax = Stn_AIO(SelectedStation, asButaneORVRFlow).EuMax
    Viscosity = 74
    SetCurrCalExists False
    optORVRBut.Value = True
End Sub

Private Sub optORVRBut_Click()
    ' the ORVR - Butane radio button was selected
    SetORVRButMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub optORVRNit_Click()
    ' The ORVR - Nitrogen radio button was selected
    SetORVRNitMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub SetORVRNitMode()
    ' Sets the form to the ORVR - Butane mode
    SelectedMFC(SelectedStation) = MFCORVRNIT
    SelectedMFCMax = Stn_AIO(SelectedStation, asNitrogenORVRFlow).EuMax
    Viscosity = 174
    SetCurrCalExists False
    optORVRNit.Value = True
End Sub

Private Sub optORVRPurgeAir_Click()
    ' Set the mode to ORVR Purge Air
    SetORVRPurgeAirMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub optPurgeAir_Click()
    ' The Purge Air radio button was selected
    SetPurgeAirMode
    ' Turn off valves
    If Not CalReadOnly Then Close_Stn_Valves SelectedStation, 1                        ' One per opto board stn level
    ' Turn On Valves for selected MFC
    If Not CalReadOnly Then UpdateMfcSelection
    ' Update the Display
    UpdateFormDisplay
End Sub

Private Sub SetPurgeAirMode()
    ' Sets the form to the Purge Air mode
    If Not CalReadOnly Then
        PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = True
    End If
    SelectedMFC(SelectedStation) = MFCPURGEAIR
    SelectedMFCMax = Stn_AIO(SelectedStation, asPurgeAirFlow).EuMax
    Viscosity = 181.87
    SetCurrCalExists False
    optPurgeAir.Value = True
End Sub

Private Sub SetORVRPurgeAirMode()
    ' Sets the form to the ORVR Purge Air mode
    If Not CalReadOnly Then
        PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = True
    End If
    SelectedMFC(SelectedStation) = MFCPURGEAIR
    SelectedMFCMax = Stn_AIO(SelectedStation, asPurgeAirFlow).EuMax
    Viscosity = 181.87
    SetCurrCalExists False
    optORVRPurgeAir.Value = True
End Sub

Private Sub UpdateFormDisplay()
    If Not CurrCalExists Then ReadCalibData
    If CurrCalExists Then
        ShowCurrCalTable
        HideCurrInactiveTableRows
        DisplCalFormula
        DisplCalResults
        DisplCurrScale
        DisplMFCNomFlowCurr
        ChartValues
'        cmdSaveNewCalib.Enabled = True
        cmdPrint.Enabled = True
        cmdCalCheckFunction.Enabled = True
    Else
        ' Hide the table
        HideCurrCalTable
        ' Hide the chart, formula, and titles
        chtMfcChart.Visible = False
        lblCalibRslt.Visible = False
        lblCalGraphTitle.Visible = False
        frmCalFormula.Visible = False
        lblCalGraphRow.Visible = False
'        cmdSaveNewCalib.Enabled = False
        cmdPrint.Enabled = False
        cmdCalCheckFunction.Enabled = False
    End If
    HideInactiveTableRows
    If PrevCalExists Then
        ShowPrevCalTable
        HidePrevInactiveTableRows
        DisplPrevCalResults
    Else
        HidePrevCalTable
    End If
    
    DisplMFCScale
    DisplMFCNomFlow
    DisplMFCData
    UpdateCalPtLabel
    UpdateLFELabel
    UpdateViscLabel
        
    txtCalibBy.text = CalibrByText
    txtComments.text = CommentText
    
    If CalReadOnly Then
        ' disable cal buttons
        cmdCalibrate.Enabled = False
        cmdCalCheckFunction.Enabled = False
        cmdKeepPrev.Enabled = False
        cmdSaveNewCalib.Enabled = False
    Else
        cmdCalibrate.Enabled = True
        cmdCalCheckFunction.Enabled = True
        ' If unsaved calibration then show the choices
        If blnUnsavedDisplacement(SelectedMFC(SelectedStation), SelectedStation) = True Then
            cmdSaveNewCalib.Enabled = True
            If PrevCalExists Then
                cmdKeepPrev.Enabled = True
            Else
                cmdKeepPrev.Enabled = False
            End If
        Else
            cmdKeepPrev.Enabled = False
            cmdSaveNewCalib.Enabled = False
        End If
    End If
    ' Update the form's display
    Select Case STN_INFO(SelectedStation).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            optLiveFuel.Visible = False
            optButane.Visible = True
            optNitrogen.Visible = True
            optORVRBut.Visible = False
            optORVRNit.Visible = False
            optPurgeAir.Visible = True
            optFlowStandard.Visible = True
            optLaminarFlowElement.Visible = True
            If CalibMethod = METHODFS Then
                cmdCalcActualFlow.Visible = False
                cmdNewLFEFile.Visible = False
                cmdSetToNominal.Visible = IIf(CalReadOnly, False, True)
                lblLFEConfig.Visible = False
                lblLFEConfigFile.Visible = False
            End If
            If CalibMethod = METHODLFE Then
                cmdCalcActualFlow.Visible = IIf(CalReadOnly, False, True)
                cmdNewLFEFile.Visible = True
                cmdSetToNominal.Visible = False
                lblLFEConfig.Visible = True
                lblLFEConfigFile.Visible = True
           End If
     
        Case STN_ORVR2_TYPE
            optLiveFuel.Visible = False
            optButane.Visible = True
            optNitrogen.Visible = True
            optORVRBut.Visible = True
            optORVRNit.Visible = True
            optPurgeAir.Visible = True
            optFlowStandard.Visible = True
            optLaminarFlowElement.Visible = True
            If CalibMethod = METHODFS Then
                cmdCalcActualFlow.Visible = False
                cmdNewLFEFile.Visible = False
                cmdSetToNominal.Visible = IIf(CalReadOnly, False, True)
                lblLFEConfig.Visible = False
                lblLFEConfigFile.Visible = False
            End If
            If CalibMethod = METHODLFE Then
                cmdCalcActualFlow.Visible = IIf(CalReadOnly, False, True)
                cmdNewLFEFile.Visible = True
                cmdSetToNominal.Visible = False
                lblLFEConfig.Visible = True
                lblLFEConfigFile.Visible = True
            End If
        
        Case STN_LIVEFUEL_TYPE
            optLiveFuel.Visible = True
            optButane.Visible = False
            optNitrogen.Visible = False
            optORVRBut.Visible = False
            optORVRNit.Visible = False
            optPurgeAir.Visible = True
            optFlowStandard.Visible = True
            optLaminarFlowElement.Visible = True
            If CalibMethod = METHODFS Then
                cmdCalcActualFlow.Visible = False
                cmdNewLFEFile.Visible = False
                cmdSetToNominal.Visible = IIf(CalReadOnly, False, True)
                lblLFEConfig.Visible = False
                lblLFEConfigFile.Visible = False
            End If
            If CalibMethod = METHODLFE Then
                cmdCalcActualFlow.Visible = IIf(CalReadOnly, False, True)
                cmdNewLFEFile.Visible = True
                cmdSetToNominal.Visible = False
                lblLFEConfig.Visible = True
                lblLFEConfigFile.Visible = True
            End If
        
        Case STN_LIVEREG_TYPE
            optLiveFuel.Visible = True
            optButane.Visible = True
            optNitrogen.Visible = True
            optORVRBut.Visible = False
            optORVRNit.Visible = False
            optPurgeAir.Visible = True
            optFlowStandard.Visible = True
            optLaminarFlowElement.Visible = True
            If CalibMethod = METHODFS Then
                cmdCalcActualFlow.Visible = False
                cmdNewLFEFile.Visible = False
                cmdSetToNominal.Visible = IIf(CalReadOnly, False, True)
                lblLFEConfig.Visible = False
                lblLFEConfigFile.Visible = False
            End If
            If CalibMethod = METHODLFE Then
                cmdCalcActualFlow.Visible = IIf(CalReadOnly, False, True)
                cmdNewLFEFile.Visible = True
                cmdSetToNominal.Visible = False
                lblLFEConfig.Visible = True
                lblLFEConfigFile.Visible = True
            End If
        
        Case STN_LIVEORVR2_TYPE
            ' future
                    
        Case STN_COMBO3_TYPE
            ' future
                    
        Case Else
            optLiveFuel.Visible = False
            optButane.Visible = False
            optNitrogen.Visible = False
            optORVRBut.Visible = False
            optORVRNit.Visible = False
            optPurgeAir.Visible = False
            optFlowStandard.Visible = False
            optLaminarFlowElement.Visible = False
            lblLFEConfigFile.Visible = False
            lblLFEConfig.Visible = False
            cmdNewLFEFile.Visible = False
            cmdCalcActualFlow.Visible = False
            cmdSetToNominal.Visible = False
    End Select
    
End Sub

Public Sub UpdateLFELabel()
    ' Updates the LFE configuration file label for the current LFE filename
    lblLFEConfig.Caption = lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
End Sub

Private Sub HideCurrCalTable()
    ' Hide the "Current Calibration" table
    Dim i As Integer
    lblMFCPercFSCol2.Visible = False
    lblMFCNomFlowCol2.Visible = False
    lblActiualFlowCol2.Visible = False
    lblCalFlowCol1.Visible = False
    lblPercDiffReadCol1.Visible = False
    For i = 0 To MAXTABLELENGTH - 1
        lblMFCPercFSCurr(i).Visible = False
        lblMFCNomFlowCurr(i).Visible = False
        lblActualFlowSLPMCurr(i).Visible = False
        lblCalFlowSLPMCurr(i).Visible = False
        lblPercDiffReadCurr(i).Visible = False
    Next i
    
    ' Hide the title
    lblCurrCalTitle.Visible = False
End Sub

Private Sub ShowCurrCalTable()
    ' Display the "Current Calibration" table
    Dim i As Integer
    lblMFCPercFSCol2.Visible = True
    lblMFCNomFlowCol2.Visible = True
    lblActiualFlowCol2.Visible = True
    lblCalFlowCol1.Visible = True
    lblPercDiffReadCol1.Visible = True
    For i = 0 To MAXTABLELENGTH - 1
        lblMFCPercFSCurr(i).Visible = True
        lblMFCNomFlowCurr(i).Visible = True
        lblActualFlowSLPMCurr(i).Visible = True
        lblCalFlowSLPMCurr(i).Visible = True
        lblPercDiffReadCurr(i).Visible = True
    Next i
    
    ' Show the title
    lblCurrCalTitle.Visible = True
End Sub

Private Sub HidePrevCalTable()
    ' Hide the "Previous calibration" table
    Dim i As Integer
    lblCalFlowCol2.Visible = False
    lblPercDiffReadCol2.Visible = False
    For i = 0 To MAXTABLELENGTH - 1
        lblCalFlowSLPMPrev(i).Visible = False
        lblPercDiffReadPrev(i).Visible = False
    Next i
    
    ' Hide the title
    lblPrevCalTitle.Visible = False
    ' Hide the associated KeepPrevious pushbutton
    cmdKeepPrev.Visible = False
End Sub

Private Sub ShowPrevCalTable()
    ' Display the "Previous Calibration" table
    Dim i As Integer
    lblCalFlowCol2.Visible = True
    lblPercDiffReadCol2.Visible = True
    For i = 0 To MAXTABLELENGTH - 1
        lblCalFlowSLPMPrev(i).Visible = True
        lblPercDiffReadPrev(i).Visible = True
    Next i
    
    ' Show the title
    lblPrevCalTitle.Visible = True
    ' Show the associated KeepPrevious pushbutton
    cmdKeepPrev.Visible = True
End Sub

Private Sub AddCalPt()
    ' Adds a calibration point (table row) to the selected MFC in the current station
    If NumMFCRows < MAXTABLELENGTH Then
        SetNumMFCRows (NumMFCRows + 1)
        ResetMFCSelection
        UpdateCalPtLabel
        HideInactiveTableRows
        DisplMFCScale
        DisplMFCNomFlow
        DisplMFCData
    End If
End Sub

Private Sub SubtractCalPt()
    ' Removes a calibration point (table row) to the selected MFC in the current station
    If NumMFCRows > MINTABLELENGTH Then
        SetNumMFCRows (NumMFCRows - 1)
        ResetMFCSelection
        UpdateCalPtLabel
        HideInactiveTableRows
        DisplMFCScale
        DisplMFCNomFlow
        DisplMFCData
    End If
End Sub

Private Sub UpdateCalPtLabel()
    ' Updates the calibration point text box (next to the spinner)
    txtNumCalPts.text = NumMFCRows
    txtNumCalPts.BackColor = Entry_BackColor
End Sub

Private Sub Timer1_Timer()
    ' updates the barometric pressure to a variable and to a label
    CurrBaroPress = AmbBaro
    lblBaroPress.Caption = Format(CurrBaroPress, "#000")
    If SelectedMFC(SelectedStation) = MFCPURGEAIR Or SelectedMFC(SelectedStation) = MFCORVRPURGEAIR Then
        If Not CalReadOnly Then PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = True
    End If
End Sub

Private Sub HideInactiveTableRows()
    ' Hide calibration table rows and radio buttons
    ' in excess of the current number of points
    Dim TableRow As Integer
    For TableRow = 1 To MAXTABLELENGTH
        If TableRow <= NumMFCRows Then
            ' Set every cell in the row to visible
            lblMFCPercFS(TableRow - 1).Visible = True
            lblMFCNomFlow(TableRow - 1).Visible = True
            txtLFEDiffPress(TableRow - 1).Visible = True
            txtLFEInletPress(TableRow - 1).Visible = True
            txtInletTemp(TableRow - 1).Visible = True
            txtBaromPress(TableRow - 1).Visible = True
            txtActualFlowSLPM(TableRow - 1).Visible = True
        
            ' Set the radio button for the row to visible
            optCalibTable(TableRow - 1).Visible = True
        Else
            ' Hide every cell in the row
            lblMFCPercFS(TableRow - 1).Visible = False
            lblMFCNomFlow(TableRow - 1).Visible = False
            txtLFEDiffPress(TableRow - 1).Visible = False
            txtLFEInletPress(TableRow - 1).Visible = False
            txtInletTemp(TableRow - 1).Visible = False
            txtBaromPress(TableRow - 1).Visible = False
            txtActualFlowSLPM(TableRow - 1).Visible = False
            
            ' Set the radio button for the row to invisible
            optCalibTable(TableRow - 1).Visible = False
        End If
        
    Next TableRow
End Sub

Private Sub SetLFEMode()
    ' Configures the form for the LFE method
    SetCalibMethod (METHODLFE)
    
    ' Enable LFE screen
    optLaminarFlowElement.Value = True

    Dim TableRow As Integer
    ' Enable the cells used for LFE mode and
    ' Disable the cells used for FS mode
    For TableRow = 1 To MAXTABLELENGTH
        txtLFEDiffPress(TableRow - 1).Enabled = True
        txtLFEDiffPress(TableRow - 1).BackColor = Entry_BackColor
        
        txtLFEInletPress(TableRow - 1).Enabled = True
        txtLFEInletPress(TableRow - 1).BackColor = Entry_BackColor
        
        txtInletTemp(TableRow - 1).Enabled = True
        txtInletTemp(TableRow - 1).BackColor = Entry_BackColor
        
        txtActualFlowSLPM(TableRow - 1).Enabled = False
        txtActualFlowSLPM(TableRow - 1).BackColor = Common_BackColor
    Next TableRow
    
    cmdCalcActualFlow.Enabled = True
    UpdateFormDisplay
End Sub

Private Sub SetFSMode()
    ' Configures the form to the FS method
    SetCalibMethod (METHODFS)
    
    ' Disable LFE screen elements
    optFlowStandard.Value = True

    Dim TableRow As Integer
    ' Disable the cells used for LFE mode and
    ' enable the cells used for FS mode
    For TableRow = 1 To MAXTABLELENGTH
        txtLFEDiffPress(TableRow - 1).Enabled = False
        txtLFEDiffPress(TableRow - 1).BackColor = Common_BackColor
        
        txtLFEInletPress(TableRow - 1).Enabled = False
        txtLFEInletPress(TableRow - 1).BackColor = Common_BackColor
        
        txtInletTemp(TableRow - 1).Enabled = False
        txtInletTemp(TableRow - 1).BackColor = Common_BackColor
 
        txtActualFlowSLPM(TableRow - 1).Enabled = True
        txtActualFlowSLPM(TableRow - 1).BackColor = Common_BackColor
    Next TableRow
    
    cmdCalcActualFlow.Enabled = False
    UpdateFormDisplay
End Sub

Private Sub DisplMFCScale()
    ' Displays the scale values for the MFC based on the number of points
    Dim TableRow As Integer
    
    For TableRow = 1 To NumMFCRows
        lblMFCPercFS(TableRow - 1).Caption = aryCalPointTable(NumMFCRows, TableRow - 1) & "%"
    Next TableRow
End Sub


Private Sub DisplMFCNomFlow()
    '   Display Mass Flow Controller Nominal Flow
    Dim TableRow As Integer
    Dim ScaleMultiplier As Single
    For TableRow = 1 To NumMFCRows
        ScaleMultiplier = aryCalPointTable(NumMFCRows, TableRow - 1) / 100
        lblMFCNomFlow(TableRow - 1).Caption = ScaleMultiplier * SelectedMFCMax
    Next TableRow
    
End Sub

Private Sub DisplMFCNomFlowCurr()
    ' Display Mass Flow Controller Nominal Flow for the Current calibration
    ' Display nominal flow in the current calibration table
    Dim TableRow As Integer
    Dim ScaleMultiplier As Single
    For TableRow = 1 To NumCurrMFCRows
        ScaleMultiplier = aryCalPointTable(NumCurrMFCRows, TableRow - 1) / 100
        lblMFCNomFlowCurr(TableRow - 1).Caption = ScaleMultiplier * CurrCalMFCMax
    Next TableRow
    
End Sub

Private Sub DisplCurrScale()
    ' Display Current Scale
    ' Displays %F.S. labels for the Current Calibration table
    Dim TableRow As Integer
    
    For TableRow = 1 To NumCurrMFCRows
        lblMFCPercFSCurr(TableRow - 1).Caption = aryCalPointTable(NumCurrMFCRows, TableRow - 1) & "%"
    Next TableRow
    DisplMFCNomFlowCurr
End Sub

Private Function CalPtsBoxValid() As Boolean
    ' Determines if the values in the calibration points text box are valid
    Dim TextBoxString As String

    TextBoxString = txtNumCalPts.text
    
    If IsNumeric(TextBoxString) Then
        If IsInteger(TextBoxString) Then
            If CInt(TextBoxString) <= MAXTABLELENGTH Then
                If CInt(TextBoxString) >= MINTABLELENGTH Then
                    ' *** The value in the box is valid ***
                    Result = VALID
                    CalPtsBoxValid = True
                Else
                    ' the value is too low
                    Result = TOOLOW
                    CalPtsBoxValid = False
                End If
            Else
                ' the value is too high
                Result = TOOHIGH
                CalPtsBoxValid = False
            End If
        Else
            ' the value is not an integer
            Result = NOTANINTEGER
            CalPtsBoxValid = False
        End If
    Else
        ' The value is not a numeric value
        Result = NOTNUMERIC
        CalPtsBoxValid = False
    End If
End Function

Private Function IsInteger(ValToTest)
    ' Determines if the number input is an integer type value
    Dim a As Long
    Dim b, Result, tolerance As Double
    
    tolerance = 0.00000001
    a = CLng(ValToTest)
    b = CDbl(a)
    
    Result = CDbl(ValToTest) - b
    
    If Result < tolerance Then
        IsInteger = True
    Else
        IsInteger = False
    End If
    
End Function

Private Sub HideCurrInactiveTableRows()
    ' Hide current calibration table rows
    ' in excess of the current number of points
    Dim TableRow As Integer
    For TableRow = 1 To MAXTABLELENGTH
        If TableRow <= NumCurrMFCRows Then
            ' Set every cell in the row to visible
            lblMFCPercFSCurr(TableRow - 1).Visible = True
            lblMFCNomFlowCurr(TableRow - 1).Visible = True
            lblActualFlowSLPMCurr(TableRow - 1).Visible = True
            lblCalFlowSLPMCurr(TableRow - 1).Visible = True
            lblPercDiffReadCurr(TableRow - 1).Visible = True
        Else
            ' Hide every cell in the row
            lblMFCPercFSCurr(TableRow - 1).Visible = False
            lblMFCNomFlowCurr(TableRow - 1).Visible = False
            lblActualFlowSLPMCurr(TableRow - 1).Visible = False
            lblCalFlowSLPMCurr(TableRow - 1).Visible = False
            lblPercDiffReadCurr(TableRow - 1).Visible = False
        End If
        
    Next TableRow
End Sub

Private Sub HidePrevInactiveTableRows()
    ' Hide previous calibration table rows
    ' in excess of the current number of points
    Dim TableRow As Integer
    For TableRow = 1 To MAXTABLELENGTH
        If TableRow <= NumPrevMFCRows Then
            ' Set every cell in the row to visible
            lblCalFlowSLPMPrev(TableRow - 1).Visible = True
            lblPercDiffReadPrev(TableRow - 1).Visible = True
        Else
            ' Hide every cell in the row
            lblCalFlowSLPMPrev(TableRow - 1).Visible = False
            lblPercDiffReadPrev(TableRow - 1).Visible = False
        End If
        
    Next TableRow
End Sub

Private Sub SelectMFCRow()
    ' Enables the radio button for the current MFC
    optCalibTable(selectedRow - 1).Value = True
    ' Selects every element in this row
    txtLFEDiffPress(selectedRow - 1).Enabled = True
    txtLFEInletPress(selectedRow - 1).Enabled = True
    txtInletTemp(selectedRow - 1).Enabled = True
    txtActualFlowSLPM(selectedRow - 1).Enabled = True
    txtActualFlowSLPM(selectedRow - 1).BackColor = Entry_BackColor
End Sub

Private Sub DisableMFCRow(RowToDisable)
    ' Unselects every element in this row
    txtLFEDiffPress(RowToDisable - 1).Enabled = False
    txtLFEInletPress(RowToDisable - 1).Enabled = False
    txtInletTemp(RowToDisable - 1).Enabled = False
    txtActualFlowSLPM(RowToDisable - 1).Enabled = False
End Sub

Private Sub UpdateMfcSelection()
    ' Updates the MFC calibration table settings based on the
    ' current row selected in SelectedMFC()
    ' Note: this function resets valves
    Dim row_num As Integer
    Dim sngScaleMultiplier As Single
    Dim sngFlowOutput As Single
    Dim sngFlowSpan As Single
       
    If CalReadOnly Then Exit Sub
    
    For row_num = 1 To NumMFCRows
        If row_num <> selectedRow Then DisableMFCRow (row_num)
    Next row_num
    SelectMFCRow
    
    ' Output the nominal flow value for the selected row to the selected MFC
    ' Get the value (in engineering units) of the flow to output
    sngScaleMultiplier = aryCalPointTable(NumMFCRows, selectedRow - 1) / 100
    sngFlowOutput = sngScaleMultiplier
    
    ' Turn on the flow valves and output the desired flow to the current MFC
    Select Case SelectedMFC(SelectedStation)
        Case MFCBUTANE
                If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_ORVR_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then
                    Stn_OutDigital SelectedStation, isButaneSol, cON
                    sngFlowSpan = Stn_AIO(SelectedStation, asButaneFlowSP).EuMax - Stn_AIO(SelectedStation, asButaneFlowSP).EuMin
                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asButaneFlowSP).EuMin
                    Stn_OutAnalog SelectedStation, asButaneFlowSP, sngFlowOutput, outNORMAL
                End If
        Case MFCNITROGEN
                If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_ORVR_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then
                    Stn_OutDigital SelectedStation, isNitrogenSol, cON
                    sngFlowSpan = Stn_AIO(SelectedStation, asNitrogenFlowSP).EuMax - Stn_AIO(SelectedStation, asNitrogenFlowSP).EuMin
                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asNitrogenFlowSP).EuMin
                    Stn_OutAnalog SelectedStation, asNitrogenFlowSP, sngFlowOutput, outNORMAL
                End If
        Case MFCPURGEAIR
                If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_ORVR_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_LIVEFUEL_TYPE Then
                    PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = True
                    Stn_OutDigital SelectedStation, isPurgeSol, cON
                    Stn_OutDigital SelectedStation, isPriDirectionSol, cON
                    sngFlowSpan = Stn_AIO(SelectedStation, asPurgeAirFlowSP).EuMax - Stn_AIO(SelectedStation, asPurgeAirFlowSP).EuMin
                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asPurgeAirFlowSP).EuMin
                    Stn_OutAnalog SelectedStation, asPurgeAirFlowSP, sngFlowOutput, outNORMAL
                End If
        Case MFCORVRBUT
               If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then
                    Stn_OutDigital SelectedStation, isButaneOrvrSol, cON
                    sngFlowSpan = Stn_AIO(SelectedStation, asButaneORVRFlowSP).EuMax - Stn_AIO(SelectedStation, asButaneORVRFlowSP).EuMin
                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asButaneORVRFlowSP).EuMin
                    Stn_OutAnalog SelectedStation, asButaneORVRFlowSP, sngFlowOutput, outNORMAL
               End If
        Case MFCORVRNIT
               If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then
                    Stn_OutDigital SelectedStation, isNitrogenOrvrSol, cON
                    sngFlowSpan = Stn_AIO(SelectedStation, asNitrogenORVRFlowSP).EuMax - Stn_AIO(SelectedStation, asNitrogenORVRFlowSP).EuMin
                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asNitrogenORVRFlowSP).EuMin
                    Stn_OutAnalog SelectedStation, asNitrogenORVRFlowSP, sngFlowOutput, outNORMAL
               End If
        Case MFCORVRPURGEAIR
'               If STN_INFO(SelectedStation).Type = STN_ORVR3_TYPE Then
'                    PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = True
'                    Stn_OutDigital SelectedStation, isPurgeSol, cON
'                    Stn_OutDigital SelectedStation, isPriDirectionSol, cON
'                    sngFlowSpan = Stn_AIO(SelectedStation, asPurgeAirFlowSP).EUMax - Stn_AIO(SelectedStation, asPurgeAirFlowSP).EUMin
'                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asPurgeAirFlowSP).EUMin
'                    Stn_OutAnalog SelectedStation, asPurgeAirFlowSP, sngFlowOutput, outNORMAL
'                End If
        Case MFCLIVEFUEL
               If STN_INFO(SelectedStation).Type = STN_LIVEFUEL_TYPE _
                        Or STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then
                    Stn_OutDigital SelectedStation, isLiveFuelSol, cON
                    sngFlowSpan = Stn_AIO(SelectedStation, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(SelectedStation, asLiveFuelVaporFlowSP).EuMin
                    sngFlowOutput = (sngFlowOutput * sngFlowSpan) + Stn_AIO(SelectedStation, asLiveFuelVaporFlowSP).EuMin
                    Stn_OutAnalog SelectedStation, asLiveFuelVaporFlowSP, sngFlowOutput, outNORMAL
'                    Stn_OutDigital SelectedStation, isFuelTankBypassSol, cON
               End If
    End Select

End Sub

Private Sub ResetMFCSelection()
    ' Points the MFC Calibration table to the first row
    selectedRow = 1
'    UpdateMFCSelection
End Sub

Private Sub UpdateMFCMode()
    ' Updates the form for the current station's Mass Flow Controller selection
    Select Case SelectedMFC(SelectedStation)
        Case MFCBUTANE
            SetButaneMode
        Case MFCNITROGEN
            SetNitrogenMode
        Case MFCPURGEAIR
            SetPurgeAirMode
        Case MFCORVRBUT
            SetORVRButMode
        Case MFCORVRNIT
            SetORVRNitMode
        Case MFCLIVEFUEL
            SetLiveFuelMode
    End Select
End Sub

Private Sub txtActualFlowSLPM_Validate(Index As Integer, Cancel As Boolean)
    If IsNumeric(txtActualFlowSLPM(Index).text) Then
        ' Store the value as Single and display the formatted value
        aryActualFlow(SelectedMFC(SelectedStation), SelectedStation, Index) = CSng(txtActualFlowSLPM(Index).text)
    End If
    txtActualFlowSLPM(Index).text = Format(aryActualFlow(SelectedMFC(SelectedStation), SelectedStation, Index), "###0.0000")
End Sub

Private Sub txtCalibBy_Change()
    SetCalibrByText (txtCalibBy.text)
End Sub

Private Sub txtComments_Change()
   SetCommentText (txtComments.text)
End Sub

Private Sub txtInletTemp_Validate(Index As Integer, Cancel As Boolean)
    If IsNumeric(txtInletTemp(Index).text) Then
        ' Store the value as Single and display the formatted value
        aryLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, Index) = CSng(txtInletTemp(Index).text)
    End If
    txtInletTemp(Index).text = Format(aryLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, Index), "##0.0")
End Sub


Private Sub txtLFEDiffPress_Validate(Index As Integer, Cancel As Boolean)
    If IsNumeric(txtLFEDiffPress(Index).text) Then
        ' Store the value as Single and display the formatted value
        aryLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, Index) = CSng(txtLFEDiffPress(Index).text)
    End If
    txtLFEDiffPress(Index).text = Format(aryLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, Index), "##0.000")
End Sub

Private Sub txtLFEInletPress_Validate(Index As Integer, Cancel As Boolean)
    If IsNumeric(txtLFEInletPress(Index).text) Then
        ' Store the value as Single and display the formatted value
        aryLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, Index) = CSng(txtLFEInletPress(Index).text)
    End If
    txtLFEInletPress(Index).text = Format(aryLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, Index), "00.00")
End Sub

Private Sub txtNumCalPts_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCrLf) Then
        txtNumCalPts_Validate False
        HandleBoxResult
    End If
End Sub

Private Sub txtNumCalPts_LostFocus()
    ' Validation results are handled after the control loses focus
    ' because the form covers the delay box if the results are handled
    ' in CalPtsBoxValid()
    HandleBoxResult
End Sub

Private Sub txtNumCalPts_Validate(Cancel As Boolean)
    Dim intTextBoxVal As Integer
    ' Update the number of calibration points
    frmMassFlowCal.SetFocus
    If CalPtsBoxValid Then
        intTextBoxVal = CInt(txtNumCalPts.text)
        SetNumMFCRows (intTextBoxVal)
        ResetMFCSelection
        HideInactiveTableRows
        DisplMFCScale
        DisplMFCNomFlow
        DisplMFCData
    End If
End Sub

Private Sub HandleBoxResult()
    ' ***************************************
    ' *** Handle valid and invalid values ***
    ' ***************************************
    Select Case Result
        Case VALID
            txtNumCalPts.BackColor = Entry_BackColor
            
        Case TOOLOW
'            Delay_Box "Value too low.  See tooltips", MSGDELAY, msgSHOW
            txtMsg.ForeColor = MEDRED
            txtMsg.text = "Value too low.  See tooltips"
            txtNumCalPts.BackColor = EntryInvalid_BackColor
            
        Case TOOHIGH
'            Delay_Box "Value too high.  See tooltips", MSGDELAY, msgSHOW
            txtMsg.ForeColor = MEDRED
            txtMsg.text = "Value too high.  See tooltips"
            txtNumCalPts.BackColor = EntryInvalid_BackColor
            
        Case NOTANINTEGER
'            Delay_Box "Value is not an integer.  See tooltips", MSGDELAY, msgSHOW
            txtMsg.ForeColor = MEDRED
            txtMsg.text = "Value is not an integer.  See tooltips"
            txtNumCalPts.BackColor = EntryInvalid_BackColor
            
        Case NOTNUMERIC
'            Delay_Box "Value is not an integer.  See tooltips", MSGDELAY, msgSHOW
            txtMsg.ForeColor = MEDRED
            txtMsg.text = "Value is not an integer.  See tooltips"
            txtNumCalPts.BackColor = EntryInvalid_BackColor
            
    End Select
End Sub

Private Function CalcActualFlow()
    ' Calculates the actual flow from the LFE diff. Pressure,
    ' LFE Inlet pressure, Inlet temperature, Barometric pressure,
    ' viscosity, and the LFE coefficients entered in the LFE
    ' definition screen
    
    ' Read in the LFE coefficients
    Dim row As Integer
    Dim a, b, c, D As Single
    Dim barometric_pressure As Single
    Dim LFEDiffPress, X, LFEInletPress As Single
    Dim InletTemp As Single
    Dim CFM, ACFM, Pcf, Tcf, SCFM As Double
    a = LFE_A(SelectedMFC(SelectedStation), SelectedStation)
    b = LFE_B(SelectedMFC(SelectedStation), SelectedStation)
    c = LFE_C(SelectedMFC(SelectedStation), SelectedStation)
    D = LFE_D(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Read in the Barometric pressure
    barometric_pressure = CurrBaroPress
    ' for each row in the MFC calibration table
    For row = 1 To NumMFCRows
        ' LFE Differential Pressure
        LFEDiffPress = CSng(txtLFEDiffPress(row - 1).text)
        X = LFEDiffPress
    
        ' LFE Inlet Pressure
        LFEInletPress = CSng(txtLFEInletPress(row - 1).text)

        InletTemp = CSng(txtInletTemp(row - 1).text)
    
        CFM = a + b * X + c * (X ^ 2) + D * (X ^ 3)
    
        ACFM = CFM * 181.87 / Viscosity
    
        Pcf = LFEInletPress / 29.92

        Tcf = 529.67 / (459.67 + InletTemp)
    
        ' Actual Flow (SCFM)
        SCFM = ACFM * Pcf * Tcf * 28.317

        aryActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1) = CSng(SCFM)
    Next row
End Function

Private Function ValidateBox(thebox As TextBox)
    ' Colors the textbox white and returns true if the value in the textbox is a number
    ' Otherwise, it colors the box yellow and returns false
    If IsNumeric(thebox.text) Then
        thebox.BackColor = Entry_BackColor
        ValidateBox = True
    Else
        thebox.BackColor = EntryInvalid_BackColor
        ValidateBox = False
    End If
End Function

Private Sub DisplCalResults()
    ' Display Calibration Results
    ' Fills the current calibration table with the calibration results
    
    Dim row As Integer
    Dim ActualFlow, CalibratedFlow, PercDiff As Single

    HideCurrInactiveTableRows
    DisplCurrScale
    For row = 1 To NumCurrMFCRows
        ' Copy the actual flow value
        ActualFlow = CSng(aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1))
        lblActualFlowSLPMCurr(row - 1).Caption = Format(ActualFlow, "##0.0000")
        ' Copy the Calibrated flow value
        CalibratedFlow = CSng(CalibratedValue(aryCalPointTable(NumCurrMFCRows, row - 1) / 100))
        lblCalFlowSLPMCurr(row - 1).Caption = Format(CalibratedFlow, "##0.0000")
        ' Display the percent difference
        If ActualFlow > 0! Then
            PercDiff = ((CalibratedFlow - ActualFlow) / ActualFlow) * 100
        Else
            PercDiff = 0!
        End If
        lblPercDiffReadCurr(row - 1).Caption = Format(PercDiff, "###0.0") & "%"
    Next row
End Sub

Private Sub DisplPrevCalResults()
    ' Display previous calibration results
    ' Fills the previous calibration table with the previous calibration results
    Dim row As Integer
    Dim ActualFlow, CalibratedFlow, PercDiff As Single
    
    HidePrevInactiveTableRows
    
    For row = 1 To NumPrevMFCRows
        ' Copy the actual flow value
        ActualFlow = CSng(aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1))
        
        ' Copy the Calibrated flow value
        CalibratedFlow = CSng(PrevCalibratedValue(aryCalPointTable(NumPrevMFCRows, row - 1) / 100))
        lblCalFlowSLPMPrev(row - 1).Caption = Format(CalibratedFlow, "##0.0000")
    
        ' Display the percent difference
        If ActualFlow > 0! Then
            PercDiff = ((CalibratedFlow - ActualFlow) / ActualFlow) * 100
        Else
            PercDiff = 0!
        End If
        lblPercDiffReadPrev(row - 1).Caption = Format(PercDiff, "###0.0") & "%"
    Next row
End Sub

Private Sub SaveCalibData()
    ' Stores the current and previous calibration data
    ' for the selected MFC to a file (ie. FILEPATH_cal & "Butane_3.cal)
    ' Also backs up the previous file (ie. FILEPATH_cal & "Butane_3.bu)
    
    Dim filename, PathFileName, BackUpName As String
    Dim fs, f As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim iFileNumber As Integer
    Dim row As Integer

    ' Build the file name
    filename = Mfc_Calib_Filename(SelectedMFC(SelectedStation))
    PathFileName = FILEPATH_cal & filename & SelectedStation & ".cal"
    BackUpName = FILEPATH_cal & filename & SelectedStation & ".bu"
    ' If the file exists then
    If fs.FileExists(PathFileName) Then
    '     ***TODO*** Read the previous calibration data
    '     Back-up the old file
       If fs.FileExists(BackUpName) Then fs.DeleteFile BackUpName
       Set f = fs.GetFile(PathFileName)
       f.Name = filename & SelectedStation & ".bu"
       Set f = Nothing
    End If
    
    ' Open the calibration file
    iFileNumber = FreeFile
    Open PathFileName For Output As iFileNumber
    
    ' Note the number of calibrations in the file
    If PrevCalExists = True Then
        Write #iFileNumber, 2
    Else
        Write #iFileNumber, 1
    End If
    
    ' Write the data to the file
    ' write a separator (ie. -----)
    Write #iFileNumber, "-----"
    ' write the "calibrated by" text
    Write #iFileNumber, CurrCalibrByText
    ' write the comments for the calibration
    Write #iFileNumber, CurrCommentText
    ' write the number of rows
    Write #iFileNumber, NumCurrMFCRows
    
    ' Write the calibration coefficients
    Write #iFileNumber, curr_coefX(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Write the MFC Max scale for the current calibration
    Write #iFileNumber, aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Write the lfe definition values
    Write #iFileNumber, curr_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
    Write #iFileNumber, curr_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Write the date
    Write #iFileNumber, aryCurrCalibDate(SelectedMFC(SelectedStation), SelectedStation)
    
    For row = 1 To NumCurrMFCRows
    '     Write values from the input table
        Write #iFileNumber, aryCurrLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Write #iFileNumber, aryCurrLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Write #iFileNumber, aryCurrLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Write #iFileNumber, aryCurrBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Write #iFileNumber, aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
    
    If PrevCalExists Then
        ' write a separator (ie. -----)
        Write #iFileNumber, "-----"
    
        ' write the "calibrated by" text for the previous
        Write #iFileNumber, PrevCalibrByText
        ' write the comments for the calibrationn for the previous calibration
        Write #iFileNumber, PrevCommentText
   
        ' Write the number of rows for the previous calibration
        Write #iFileNumber, NumPrevMFCRows
        
        ' Write the calibration coefficients
        Write #iFileNumber, prev_coefX(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_coefX2(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_coefX3(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_coefX4(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_coefX5(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_coefX6(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_coefR2(SelectedMFC(SelectedStation), SelectedStation)
        
        ' Write the MFC Max scale for the previous calibration
        Write #iFileNumber, aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation)
        
        ' Write the lfe definition values
        Write #iFileNumber, prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
        Write #iFileNumber, prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
        
        ' Write the date of the previous calibration
        Write #iFileNumber, aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation)
        
        For row = 1 To NumPrevMFCRows
            ' Write values from the input table
            Write #iFileNumber, aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Write #iFileNumber, aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Write #iFileNumber, aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Write #iFileNumber, aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Write #iFileNumber, aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Next row
    End If
    Close #iFileNumber
End Sub

Private Sub ReadCalibData()
    ' Reads in current and previous calibration data from a file and reports that
    ' current and previous calibrations now exist
    Dim filename, PathFileName As String
    Dim TempString As String
    Dim NumCalibrations, NUMROWS As Integer
    Dim fs, f As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim iFileNumber As Integer
    Dim row As Integer

    ' Build the file name
    filename = Mfc_Calib_Filename(SelectedMFC(SelectedStation))
    PathFileName = FILEPATH_cal & filename & SelectedStation & ".cal"

    ' Continue only if the file exists
    If fs.FileExists(PathFileName) = False Then
        ' Delay_Box "no calibration file", MSGDELAY, msgSHOW
        txtMsg.ForeColor = MEDRED
        txtMsg.text = "no calibration file"
        Exit Sub
    End If
    ' Open the calibration file
    iFileNumber = FreeFile
    Open PathFileName For Input As iFileNumber
    
    ' Read number of calibrations in file
    Input #iFileNumber, NumCalibrations
    
    ' Read the separator
    Input #iFileNumber, TempString
    
    ' Read the "Calibrated By" text
    Input #iFileNumber, TempString
    SetCurrCalibrByText TempString
    SetCalibrByText TempString
    
    ' Read the comment text
    Input #iFileNumber, TempString
    SetCurrCommentText TempString
    SetCommentText TempString
    
    ' Read the number of rows
    Input #iFileNumber, NUMROWS
    SetNumCurrMFCRows NUMROWS
    
    ' Read the calibration coefficient
    Input #iFileNumber, curr_coefX(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Read the MFC Max scale for the current calibration
     Input #iFileNumber, aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Read the lfe definition values
    Input #iFileNumber, curr_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
    Input #iFileNumber, curr_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
    
    ' Read the date
    Input #iFileNumber, aryCurrCalibDate(SelectedMFC(SelectedStation), SelectedStation)

    For row = 1 To NUMROWS
        ' Read values from the input table
        Input #iFileNumber, aryCurrLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Input #iFileNumber, aryCurrLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Input #iFileNumber, aryCurrLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Input #iFileNumber, aryCurrBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Input #iFileNumber, aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
    Next row
    ' Report that a current calibration now exists
    SetCurrCalExists True

    If NumCalibrations > 1 Then

        ' Read the separator
        Input #iFileNumber, TempString
    
        ' Read the "Calibrated By" text
        Input #iFileNumber, TempString
        SetPrevCalibrByText TempString
    
        ' Read the comment text
        Input #iFileNumber, TempString
        SetPrevCommentText TempString
    
        ' Read the number of rows
        Input #iFileNumber, NUMROWS
        SetNumPrevMFCRows NUMROWS
        
        ' Read the calibration coefficients
        Input #iFileNumber, prev_coefX(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_coefX2(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_coefX3(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_coefX4(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_coefX5(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_coefX6(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_coefR2(SelectedMFC(SelectedStation), SelectedStation)
        ' Read the MFC Max scale for the previous calibration
        Input #iFileNumber, aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation)
        ' Read the lfe definition values
        Input #iFileNumber, prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation)
        Input #iFileNumber, prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)

    '     Read the date
        Input #iFileNumber, aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation)
    
        For row = 1 To NUMROWS
            ' Read values from the input table
            Input #iFileNumber, aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Input #iFileNumber, aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Input #iFileNumber, aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Input #iFileNumber, aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            Input #iFileNumber, aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        Next row
        ' Report that a previous calibration now exists
        SetPrevCalExists True
    End If
    ' Close the file
    Close #iFileNumber
End Sub

Private Sub PrintCalibration()
' Print the values found for the current and previous calibrations
Dim TempString As String
Dim strPerc As String
Dim strDiffPress As String
Dim strInletTemp As String
Dim strInletPress As String
Dim strBaroPress As String
Dim strActualFlow As String
Dim strCalibFlow As String
Dim strPercDiff As String
Dim row As Integer
Dim intNumRows As Integer
Dim intNumPrevRows As Integer
Dim sngCalibFlow As Single
Dim sngActualFlow As Single
Dim sngPercDiff As Single
' ************  NEW  ****************
Dim oldFont As New StdFont
' Save current printer font
oldFont = Printer.Font
Printer.Font = FILEFONT
Printer.Font.Size = 8.5  ' FILEFONTSIZE
Printer.Font.Bold = False
Printer.Font.Italic = False

    intNumRows = NumCurrMFCRows
    
    ' Print a title / header
   Print_Center Trim(SysConfig.Heading)
   Print_Center Trim(SysConfig.Heading2)
   Print_Center "CANISTER PRECONDITIONING SYSTEM"
   Print_Line ""
    Print_Center "Mass Flow Controller Calibration"
    Print_Center ("Date: " & Format(Now, "mmm d, yyyy"))
    Print_Line ""
    ' print the the name of the mass flow controller
    Print_Line "Station #" & SelectedStation
    Print_Line "Mass Flow Controller: " & Mfc_Description(SelectedMFC(SelectedStation))
    Print_Line ""
    Print_Line ""
    Print_Line "Current Calibration"
    ' Print "Calibrated By:" then the calibrated by text
    Print_Line "Calibrated By: " & CalibrByText
    
    ' Print the comments
    Print_Line "Comments: " & Mid(CurrCommentText, 1, 100)
    If Len(CurrCommentText) > 100 Then
        Print_Line Mid(CurrCommentText, 101, 110)
        If Len(CurrCommentText) > 210 Then
            Print_Line Mid(CurrCommentText, 211, 110)
        End If
    End If
        
    ' Print "Calibration date" then aryCurrCalibDate
    Print_Line "Calibration date: " & aryCurrCalibDate(SelectedMFC(SelectedStation), SelectedStation)

    ' Print "Mass Flow Controller scale" then the scale value (aryCurrCalMFCMax)
    Print_Line "Mass Flow Controller scale: " & aryCurrCalMFCMax(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "Number of calibration points: " & intNumRows
    
    ' Print a new line
    Print_Line ""
    
    ' Print the LFE definition info
    Print_Line "LFE Information: " & curr_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) & ".LFE , " & curr_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "LFE Flow = A + Bx + Cx2 + Dx3"
    Print_Line "x = Differential Pressure"
    Print_Line "Coefficients:  A:" & curr_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "               B:" & curr_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "               C:" & curr_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "               D:" & curr_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    
    Print_Line ""
    Print_Line "Flow (SLPM) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx"
    Print_Line "      Where x = % Full Scale"
    Print_Line ""
    ' Print "Calibration Formula", centered
    Print_Center "Calibration Coefficients"
    
    ' Print the calbration formula text
    ' Print_Line "Y = " & curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) & "X6" & IIf(curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)) & "X5" & IIf(curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)) & "X4" & IIf(curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)) & "X3" & IIf(curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)) & "X2" & IIf(curr_coefX(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX(SelectedMFC(SelectedStation), SelectedStation)) & "X" & "      R2=" & curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "A:            " & curr_coefX6(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "B:            " & curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "C:            " & curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "D:            " & curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "E:            " & curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line "F:            " & curr_coefX(SelectedMFC(SelectedStation), SelectedStation)
    Print_Line ""
    Print_Line "r2:           " & curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)
    
    Print_Line ""
    ' Print "Calibration points" centered
    Print_Center "Calibration Points"
    ' For each calibration point

    Print_Line "%F.S.   LFE Dif. Pres.  LFE Inl. Pres.  LFE Inl. Temp.  Baro. Pres.  Actual Flow  Calibr. Flow  % Diff."
    For row = 1 To intNumRows
        ' Print data from the input table
        strPerc = Format(aryCalPointTable(intNumRows, row - 1), "000.0") & "%"
        strDiffPress = Format(aryCurrLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "00.000")
        strInletPress = Format(aryCurrLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "00.00")
        strInletTemp = Format(aryCurrLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1), "000.0")
        strBaroPress = Format(aryCurrBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "00.00")
        
        sngActualFlow = aryCurrActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
        strActualFlow = Format(sngActualFlow, "000.0000")
        
        sngCalibFlow = CalibratedValue(aryCalPointTable(intNumRows, row - 1) / 100)
        strCalibFlow = Format(sngCalibFlow, "000.0000")
        
        ' The percent difference
        If sngActualFlow > 0! Then
            sngPercDiff = ((sngCalibFlow - sngActualFlow) / sngActualFlow) * 100
        Else
            sngPercDiff = 0!
        End If
        
        strPercDiff = Format(sngPercDiff, "0000.0") & "%"
        Print_Line strPerc & "  " & strDiffPress & "          " & strInletPress & "           " & strInletTemp & "           " & strBaroPress & "        " & strActualFlow & "     " & strCalibFlow & "      " & strPercDiff

    Next row
    
    If PrevCalExists = True Then
        intNumPrevRows = NumPrevMFCRows
        
        ' print two new lines
        Print_Line ""
        Print_Line ""
        ' Print the previous calibration data (as above)
        Print_Line "Previous Calibration"
        ' Print "Calibrated By:" then the calibrated by text

        Print_Line "Calibrated By: " & PrevCalibrByText
       
        ' Print the comments
        Print_Line "Comments: " & Mid(PrevCommentText, 1, 100)
        If Len(PrevCommentText) > 100 Then
            Print_Line Mid(PrevCommentText, 101, 110)
            If Len(PrevCommentText) > 210 Then
                Print_Line Mid(PrevCommentText, 211, 100)
            End If
        End If
        
        ' Print the previous calibration date
        Print_Line "Calibration date: " & aryPrevCalibDate(SelectedMFC(SelectedStation), SelectedStation)

        ' Print "Mass Flow Controller scale" then the scale value (aryCurrCalMFCMax)
        Print_Line "Mass Flow Controller scale: " & aryPrevMFCMax(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "Number of calibration points: " & intNumPrevRows
        ' Print a new line
        Print_Line ""
        ' Print the LFE definition info
        Print_Line "LFE Information: " & prev_lfe_filename(SelectedMFC(SelectedStation), SelectedStation) & ".LFE , " & prev_lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "LFE Flow = A + Bx + Cx2 + Dx3"
        Print_Line "x = Differential Pressure"
        Print_Line "Coefficients:  A:" & prev_lfe_a(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "               B:" & prev_lfe_b(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "               C:" & prev_lfe_c(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "               D:" & prev_lfe_d(SelectedMFC(SelectedStation), SelectedStation)
    
        Print_Line ""
        Print_Line "Flow (SLPM) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx"
        Print_Line "      Where x = % Full Scale"
        Print_Line ""
        ' Print "Calibration Formula", centered
        Print_Center "Calibration Coefficients"
        ' Print the calbration formula text
        ' Print_Line "Y = " & curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) & "X6" & IIf(curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX5(SelectedMFC(SelectedStation), SelectedStation)) & "X5" & IIf(curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX4(SelectedMFC(SelectedStation), SelectedStation)) & "X4" & IIf(curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX3(SelectedMFC(SelectedStation), SelectedStation)) & "X3" & IIf(curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX2(SelectedMFC(SelectedStation), SelectedStation)) & "X2" & IIf(curr_coefX(SelectedMFC(SelectedStation), SelectedStation) < 0, " - ", " + ") & Abs(curr_coefX(SelectedMFC(SelectedStation), SelectedStation)) & "X" & "      R2=" & curr_coefR2(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "A:            " & prev_coefX6(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "B:            " & prev_coefX5(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "C:            " & prev_coefX4(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "D:            " & prev_coefX3(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "E:            " & prev_coefX2(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line "F:            " & prev_coefX(SelectedMFC(SelectedStation), SelectedStation)
        Print_Line ""
        Print_Line "r2:           " & prev_coefR2(SelectedMFC(SelectedStation), SelectedStation)
        
        Print_Line ""
        
        ' Print "Calibration points" centered
        Print_Center "Calibration Points"
        ' For each calibration point
        Print_Line "%F.S.   LFE Dif. Pres.  LFE Inl. Pres.  LFE Inl. Temp.  Baro. Pres.  Actual Flow  Calibr. Flow  % Diff."

        For row = 1 To intNumPrevRows
            ' Print data from the input table
            strPerc = Format(aryCalPointTable(intNumPrevRows, row - 1), "000.0") & "%"
            strDiffPress = Format(aryPrevLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "00.000")
            strInletPress = Format(aryPrevLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "00.00")
            strInletTemp = Format(aryPrevLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1), "000.0")
            strBaroPress = Format(aryPrevBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "00.00")
        
            sngActualFlow = aryPrevActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1)
            strActualFlow = Format(sngActualFlow, "000.0000")
        
            sngCalibFlow = PrevCalibratedValue(aryCalPointTable(intNumRows, row - 1) / 100)
            strCalibFlow = Format(sngCalibFlow, "000.0000")
        
            ' The percent difference
            If sngActualFlow > 0! Then
                sngPercDiff = ((sngCalibFlow - sngActualFlow) / sngActualFlow) * 100
            Else
                sngPercDiff = 0!
            End If
        
            strPercDiff = Format(sngPercDiff, "0000.0") & "%"
            Print_Line strPerc & "  " & strDiffPress & "          " & strInletPress & "           " & strInletTemp & "           " & strBaroPress & "        " & strActualFlow & "     " & strCalibFlow & "      " & strPercDiff
        Next row
    
    End If

    Print_Footer
    Printer.EndDoc
    Printer.Font = oldFont
End Sub

Private Sub DisplMFCData()
    ' Displays LFE, barometric pressure, and actual flow data to the main MFC table
    Dim row As Integer
    For row = 1 To NumMFCRows
        txtLFEDiffPress(row - 1).text = Format(aryLFEDiffPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "####0.0##")
        txtLFEInletPress(row - 1).text = Format(aryLFEInletPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "####0.0##")
        txtInletTemp(row - 1).text = Format(aryLFEInletTemp(SelectedMFC(SelectedStation), SelectedStation, row - 1), "####0.0##")
        txtBaromPress(row - 1).text = Format(aryBaroPress(SelectedMFC(SelectedStation), SelectedStation, row - 1), "####0.0##")
        txtActualFlowSLPM(row - 1).text = Format(aryActualFlow(SelectedMFC(SelectedStation), SelectedStation, row - 1), "####0.0##")
    Next row
End Sub

Public Sub SetLFE_A(Value As Single)
    ' Assign the constant LFE coefficient for the selected mass flow controller
    LFE_A(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Public Sub SetLFE_B(Value As Single)
    ' Assign the X LFE coefficient for the selected mass flow controller
    LFE_B(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Public Sub SetLFE_C(Value As Single)
    ' Assign the X2 LFE coefficient for the selected mass flow controller
    LFE_C(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Public Sub SetLFE_D(Value As Single)
    ' Assign the X3 LFE coefficient for the selected mass flow controller
    LFE_D(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Public Sub SetLFE_SerialNum(Value As String)
    ' Assign the LFE serial number for the selected mass flow controller
    lfe_serialnum(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Public Sub SetLFE_FileName(Value As String)
    ' Assign the LFE file name for the selected mass flow controller
    lfe_filename(SelectedMFC(SelectedStation), SelectedStation) = Value
End Sub

Public Function SolveCalibFor(DesiredVal As Single) As Single
    ' Uses Newton's method to find the percent full scale value for
    ' a certain flow in SLPM
    Dim calx As Double
    Dim calx0 As Double
    Dim calx1 As Double
    Dim DesVal As Double
    Dim Test1 As Double
    Dim Tolerance_1, Tolerance_2 As Double
    
    ' Set the tolerance values
    Tolerance_1 = 0.0000001
    Tolerance_2 = 0.0000001
    
    DesVal = CDbl(DesiredVal)
    calx = DesVal / CDbl(SelectedMFCMax)
    calx0 = calx
    calx1 = calx0


    If CalibValue_D(calx0, DesVal) <> 0 And DerCalibValue_D(calx0) <> 0 Then
        Do
            calx0 = calx1
            calx1 = calx0 - CalibValue_D(calx0, DesVal) / DerCalibValue_D(calx0)
        Loop Until (Abs(calx0 - calx1) < Tolerance_1) Or (Abs(CalibValue_D(calx1, DesVal)) < Tolerance_2)
    End If
    SolveCalibFor = calx1
End Function

Public Function CalibratedValue(InputVal As Single) As Single
    ' The inputs value should be (percent full scale) / 100
    ' Uses the formula f(x) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx
    ' Returns the calibrated value in SLPM
    Dim tempdbl As Double
    tempdbl = curr_coefX(SelectedMFC(SelectedStation), SelectedStation) * InputVal
    tempdbl = tempdbl + curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 2)
    tempdbl = tempdbl + curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 3)
    tempdbl = tempdbl + curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 4)
    tempdbl = tempdbl + curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 5)
    tempdbl = tempdbl + curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 6)
    
    CalibratedValue = CSng(tempdbl)

End Function


Public Function CalibValue_D(InputVal As Double, DesiredValue As Double) As Double
    ' Used for Newton's method calculations
    
    ' The inputs value should be (percent full scale) / 100
    ' Uses the formula f(x) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx - DesiredValue
    ' Returns the calibrated value in SLPM
    Dim tempdbl As Double
    tempdbl = curr_coefX(SelectedMFC(SelectedStation), SelectedStation) * InputVal
    tempdbl = tempdbl + curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 2)
    tempdbl = tempdbl + curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 3)
    tempdbl = tempdbl + curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 4)
    tempdbl = tempdbl + curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 5)
    tempdbl = tempdbl + curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 6)
    tempdbl = tempdbl - DesiredValue
    CalibValue_D = tempdbl

End Function

Public Function DerCalibValue(InputVal As Single) As Single
    ' The inputs value should be (percent full scale) / 100
    
    ' Returns the derivative of the calibrated value
    ' using the formula f'(x) = 6Ax5 + 5Bx4 + 4Cx3 + 3Dx2 + 2Ex + F
    Dim tempdbl As Double
    tempdbl = curr_coefX(SelectedMFC(SelectedStation), SelectedStation)
    tempdbl = tempdbl + 2 * curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) * (InputVal)
    tempdbl = tempdbl + 3 * curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 2)
    tempdbl = tempdbl + 4 * curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 3)
    tempdbl = tempdbl + 5 * curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 4)
    tempdbl = tempdbl + 6 * curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 5)
    
    DerCalibValue = CSng(tempdbl)
   
End Function

Public Function DerCalibValue_D(InputVal As Double) As Double
    ' This is DerCalibValue_D(), but with the Double type for the input and the output
    ' The inputs value should be (percent full scale) / 100
    
    ' Returns the derivative of the calibrated value
    ' using the formula f'(x) = 6Ax5 + 5Bx4 + 4Cx3 + 3Dx2 + 2Ex + F
    Dim tempdbl As Double
    tempdbl = curr_coefX(SelectedMFC(SelectedStation), SelectedStation)
    tempdbl = tempdbl + 2 * curr_coefX2(SelectedMFC(SelectedStation), SelectedStation) * (InputVal)
    tempdbl = tempdbl + 3 * curr_coefX3(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 2)
    tempdbl = tempdbl + 4 * curr_coefX4(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 3)
    tempdbl = tempdbl + 5 * curr_coefX5(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 4)
    tempdbl = tempdbl + 6 * curr_coefX6(SelectedMFC(SelectedStation), SelectedStation) * (InputVal ^ 5)
    
    DerCalibValue_D = tempdbl
   
End Function

Private Sub ManualMode()
    ' Start the timer
    'Timer1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Xit()
        ' close valves
    If Not CalReadOnly Then
        Close_Stn_Valves SelectedStation, 1
        PRG_INFO(STN_INFO(SelectedStation).AspiratorNum).RequestRun = False
    End If
    ' close screen
    Unload Me
    Set frmMassFlowCal = Nothing
End Sub


