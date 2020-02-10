VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form frmScalesCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scales Calibration Screen"
   ClientHeight    =   11160
   ClientLeft      =   195
   ClientTop       =   720
   ClientWidth     =   15315
   Icon            =   "FrmScalesCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   15315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmScaleInformation 
      Caption         =   "Scale Information"
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
      Height          =   1695
      Left            =   90
      TabIndex        =   168
      Top             =   3480
      Width           =   7065
      Begin VB.TextBox txtSclEuMin 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3750
         TabIndex        =   171
         Text            =   "0"
         ToolTipText     =   "Min Scale Range in grams"
         Top             =   840
         Width           =   960
      End
      Begin VB.TextBox txtSclEuMax 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3750
         TabIndex        =   169
         Text            =   "68888"
         ToolTipText     =   "Max Scale Range in grams"
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblSclEuMin 
         BackStyle       =   0  'Transparent
         Caption         =   "Scale Min Grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   172
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblSclEuMax 
         BackStyle       =   0  'Transparent
         Caption         =   "Scale Max Grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   170
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame frmScaleSelection 
      Caption         =   "Scale Selection"
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
      Height          =   975
      Left            =   90
      TabIndex        =   2
      Top             =   2400
      Width           =   7065
      Begin VB.CommandButton cmdScaleUp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "next scale"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.CommandButton cmdScaleDn 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":5EE4
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "previous scale"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.TextBox txtDispScl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   615
         Left            =   1605
         TabIndex        =   3
         Text            =   "Scale #8"
         Top             =   240
         Width           =   5340
      End
   End
   Begin VB.Frame frmCalControls 
      Caption         =   "Calibration Controls"
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
      Height          =   2175
      Left            =   90
      TabIndex        =   11
      Top             =   120
      Width           =   7065
      Begin VB.CommandButton cmdCalPoints 
         Caption         =   "Edit CalPoints"
         DisabledPicture =   "FrmScalesCal.frx":65E6
         DownPicture     =   "FrmScalesCal.frx":6928
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":6C6A
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdPrintCal 
         Caption         =   "Print"
         DisabledPicture =   "FrmScalesCal.frx":6FAC
         DownPicture     =   "FrmScalesCal.frx":76AE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4957
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":7DB0
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   "Print a Report of the Selected Calibration"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdRestorePrevCal 
         Caption         =   "Restore Previous"
         DisabledPicture =   "FrmScalesCal.frx":84B2
         DownPicture     =   "FrmScalesCal.frx":87F4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3915
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":8B36
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdSaveCurrCal 
         Caption         =   "     Save      Current"
         DisabledPicture =   "FrmScalesCal.frx":8E78
         DownPicture     =   "FrmScalesCal.frx":91BA
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2970
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":94FC
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdCreateNewCalib 
         Caption         =   " Create New"
         DisabledPicture =   "FrmScalesCal.frx":983E
         DownPicture     =   "FrmScalesCal.frx":9B80
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":9EC2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.TextBox txtMsg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "FrmScalesCal.frx":A204
         Top             =   1080
         Width           =   6795
      End
      Begin VB.CommandButton cmdRunCal 
         Caption         =   "Calibration"
         DisabledPicture =   "FrmScalesCal.frx":A22F
         DownPicture     =   "FrmScalesCal.frx":A571
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2025
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":A8B3
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
   End
   Begin VB.Frame frmCalInformation 
      Caption         =   "Calibration Information"
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
      Height          =   4155
      Left            =   90
      TabIndex        =   8
      Top             =   6960
      Width           =   4725
      Begin VB.ComboBox PressUnits 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         ItemData        =   "FrmScalesCal.frx":ABF5
         Left            =   3120
         List            =   "FrmScalesCal.frx":AC05
         Style           =   2  'Dropdown List
         TabIndex        =   186
         ToolTipText     =   "Units for Standard Pressure"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtStandardPress 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   184
         Text            =   "1"
         ToolTipText     =   "Standard Pressure Value"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox TempUnits 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         ItemData        =   "FrmScalesCal.frx":AC20
         Left            =   3120
         List            =   "FrmScalesCal.frx":AC30
         Style           =   2  'Dropdown List
         TabIndex        =   183
         ToolTipText     =   "Units for Standard Temperature"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtStandardTemp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   181
         Text            =   "20"
         ToolTipText     =   "Standard Temperature Value"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txtCalibBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   176
         Text            =   "FrmScalesCal.frx":AC50
         ToolTipText     =   "Maximum length: 32 characters"
         Top             =   2130
         Width           =   2775
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   1215
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   175
         Text            =   "FrmScalesCal.frx":AC5B
         Top             =   2790
         Width           =   2775
      End
      Begin VB.TextBox txtCalibDts 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   174
         Text            =   "YYYY-MMM-DD hh:mm:ss"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtEquipment 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   173
         Text            =   "FrmScalesCal.frx":AC65
         ToolTipText     =   "Maximum length: 32 characters"
         Top             =   2460
         Width           =   2775
      End
      Begin VB.CommandButton cmdApply 
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
         Left            =   3615
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmScalesCal.frx":AC6F
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   840
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2790
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "11"
         ToolTipText     =   "Number of points used to calibrate the Analog Input"
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblStandardPress 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Pressure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   185
         Top             =   1710
         Width           =   1965
      End
      Begin VB.Label lblStandardTemp 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Temperature"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   182
         Top             =   1350
         Width           =   1965
      End
      Begin VB.Label lblEquipment 
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   180
         Top             =   2490
         Width           =   1365
      End
      Begin VB.Label lblCalibDts 
         BackStyle       =   0  'Transparent
         Caption         =   "Calibration DTS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   179
         Top             =   1005
         Width           =   1365
      End
      Begin VB.Label lblComments 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   178
         Top             =   2820
         Width           =   1365
      End
      Begin VB.Label lblCalibBy 
         BackStyle       =   0  'Transparent
         Caption         =   "Calibrated By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   177
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label lblNumCalPts 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Calibration Points"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frmCalPointData 
      Caption         =   "Calibration Point Data"
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
      Height          =   4155
      Left            =   4920
      TabIndex        =   4
      Top             =   6960
      Width           =   10335
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   675
         TabIndex        =   64
         Top             =   3405
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   675
         TabIndex        =   63
         Top             =   3120
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   675
         TabIndex        =   62
         Top             =   2835
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   675
         TabIndex        =   61
         Top             =   2550
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   675
         TabIndex        =   60
         Top             =   2265
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   675
         TabIndex        =   59
         Top             =   1980
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   675
         TabIndex        =   58
         Text            =   "23456.78"
         Top             =   1695
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   675
         TabIndex        =   57
         Top             =   1410
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   675
         TabIndex        =   56
         Top             =   1125
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   675
         TabIndex        =   55
         Top             =   840
         Width           =   780
      End
      Begin VB.TextBox txtNewRawValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   11
         Left            =   675
         TabIndex        =   54
         Top             =   3690
         Width           =   780
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   2115
         TabIndex        =   53
         Top             =   3405
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   2115
         TabIndex        =   52
         Top             =   3120
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   2115
         TabIndex        =   51
         Top             =   2835
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   2115
         TabIndex        =   50
         Top             =   2550
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   2115
         TabIndex        =   49
         Top             =   2265
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   2115
         TabIndex        =   48
         Top             =   1980
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   2115
         TabIndex        =   47
         Top             =   1695
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   2115
         TabIndex        =   46
         Top             =   1410
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   2115
         TabIndex        =   45
         Top             =   1125
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   2115
         TabIndex        =   44
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox txtNewActualValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   11
         Left            =   2115
         TabIndex        =   43
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   1455
         TabIndex        =   165
         Top             =   3690
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   1455
         TabIndex        =   164
         Top             =   3405
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   1455
         TabIndex        =   163
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   1455
         TabIndex        =   162
         Top             =   2835
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   1455
         TabIndex        =   161
         Top             =   2550
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   1455
         TabIndex        =   160
         Top             =   2265
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   1455
         TabIndex        =   159
         Top             =   1980
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   1455
         TabIndex        =   158
         Top             =   1695
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   1455
         TabIndex        =   157
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1455
         TabIndex        =   156
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label lblNewRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1455
         TabIndex        =   155
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblNewRaws2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Raw %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1455
         TabIndex        =   154
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   5730
         TabIndex        =   148
         Top             =   3405
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   5730
         TabIndex        =   147
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   5730
         TabIndex        =   146
         Top             =   2835
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   5730
         TabIndex        =   145
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   5730
         TabIndex        =   144
         Top             =   2265
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   5730
         TabIndex        =   143
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5730
         TabIndex        =   142
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   5730
         TabIndex        =   141
         Top             =   1410
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   5730
         TabIndex        =   140
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   5730
         TabIndex        =   139
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblCurrDiffs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "% Diff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5730
         TabIndex        =   138
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblCurrDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   5730
         TabIndex        =   137
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   4710
         TabIndex        =   136
         Top             =   3405
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   4710
         TabIndex        =   135
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   4710
         TabIndex        =   134
         Top             =   2835
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4710
         TabIndex        =   133
         Top             =   2550
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   4710
         TabIndex        =   132
         Top             =   2265
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   4710
         TabIndex        =   131
         Top             =   1980
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4710
         TabIndex        =   130
         Top             =   1695
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   4710
         TabIndex        =   129
         Top             =   1410
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   4710
         TabIndex        =   128
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   4710
         TabIndex        =   127
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblCurrCals 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Calibrated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4710
         TabIndex        =   126
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblCurrCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   4710
         TabIndex        =   125
         Top             =   3690
         Width           =   1020
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   9285
         TabIndex        =   124
         Top             =   3405
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   9285
         TabIndex        =   123
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   9285
         TabIndex        =   122
         Top             =   2835
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   9285
         TabIndex        =   121
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   9285
         TabIndex        =   120
         Top             =   2265
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   9285
         TabIndex        =   119
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   9285
         TabIndex        =   118
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   9285
         TabIndex        =   117
         Top             =   1410
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   9285
         TabIndex        =   116
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   9285
         TabIndex        =   115
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblPrevDiffs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "% Diff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9285
         TabIndex        =   114
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblPrevDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   9285
         TabIndex        =   113
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   8265
         TabIndex        =   112
         Top             =   3405
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   8265
         TabIndex        =   111
         Top             =   3120
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   8265
         TabIndex        =   110
         Top             =   2835
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   8265
         TabIndex        =   109
         Top             =   2550
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   8265
         TabIndex        =   108
         Top             =   2265
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   8265
         TabIndex        =   107
         Top             =   1980
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   8265
         TabIndex        =   106
         Top             =   1695
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   8265
         TabIndex        =   105
         Top             =   1410
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   8265
         TabIndex        =   104
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   8265
         TabIndex        =   103
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblPrevCals 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Calibrated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8265
         TabIndex        =   102
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblPrevCalValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   8265
         TabIndex        =   101
         Top             =   3690
         Width           =   1020
      End
      Begin VB.Label lblCurrPoints 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Point"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   100
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   99
         Top             =   3690
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   98
         Top             =   3405
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   97
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   96
         Top             =   2835
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   95
         Top             =   2550
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   94
         Top             =   2265
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   93
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   92
         Top             =   1695
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   91
         Top             =   1410
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   90
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label lblPointNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   89
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   7365
         TabIndex        =   88
         Top             =   3405
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   7365
         TabIndex        =   87
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   7365
         TabIndex        =   86
         Top             =   2835
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   7365
         TabIndex        =   85
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   7365
         TabIndex        =   84
         Top             =   2265
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   7365
         TabIndex        =   83
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   7365
         TabIndex        =   82
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   7365
         TabIndex        =   81
         Top             =   1410
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   7365
         TabIndex        =   80
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   7365
         TabIndex        =   79
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7365
         TabIndex        =   78
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblPrevActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   7365
         TabIndex        =   77
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Raw %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6705
         TabIndex        =   76
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   6705
         TabIndex        =   75
         Top             =   3405
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   6705
         TabIndex        =   74
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   6705
         TabIndex        =   73
         Top             =   2835
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   6705
         TabIndex        =   72
         Top             =   2550
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   6705
         TabIndex        =   71
         Top             =   2265
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   6705
         TabIndex        =   70
         Top             =   1980
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   6705
         TabIndex        =   69
         Top             =   1695
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   6705
         TabIndex        =   68
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   6705
         TabIndex        =   67
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   6705
         TabIndex        =   66
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblPrevRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   6705
         TabIndex        =   65
         Top             =   3690
         Width           =   660
      End
      Begin VB.Label lblNewActuals 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2115
         TabIndex        =   42
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblNewRaws 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Raw"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   675
         TabIndex        =   41
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   3810
         TabIndex        =   40
         Top             =   3405
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   3810
         TabIndex        =   39
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   3810
         TabIndex        =   38
         Top             =   2835
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3810
         TabIndex        =   37
         Top             =   2550
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   3810
         TabIndex        =   36
         Top             =   2265
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   3810
         TabIndex        =   35
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   3810
         TabIndex        =   34
         Top             =   1695
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3810
         TabIndex        =   33
         Top             =   1410
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3810
         TabIndex        =   32
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3810
         TabIndex        =   31
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblCurrActuals 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3810
         TabIndex        =   30
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblCurrActualValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   3810
         TabIndex        =   29
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label lblCurrRaws 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Raw %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3150
         TabIndex        =   28
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   10
         Left            =   3150
         TabIndex        =   27
         Top             =   3405
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   9
         Left            =   3150
         TabIndex        =   26
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   8
         Left            =   3150
         TabIndex        =   25
         Top             =   2835
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   3150
         TabIndex        =   24
         Top             =   2550
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   3150
         TabIndex        =   23
         Top             =   2265
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   3150
         TabIndex        =   22
         Top             =   1980
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   3150
         TabIndex        =   21
         Top             =   1695
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   3150
         TabIndex        =   20
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3150
         TabIndex        =   19
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "188.88"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   3150
         TabIndex        =   18
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblCurrRawPerc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   11
         Left            =   3150
         TabIndex        =   17
         Top             =   3690
         Width           =   660
      End
      Begin VB.Label lblPrevious 
         Alignment       =   2  'Center
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   6705
         TabIndex        =   7
         Top             =   285
         Width           =   3480
      End
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   3150
         TabIndex        =   6
         Top             =   285
         Width           =   3465
      End
      Begin VB.Label lblNew 
         Alignment       =   2  'Center
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   675
         TabIndex        =   5
         Top             =   285
         Width           =   2340
      End
   End
   Begin VB.Frame frmCalFormula 
      Caption         =   "Calibration Formula"
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
      Height          =   1575
      Left            =   90
      TabIndex        =   0
      Top             =   5280
      Width           =   7065
      Begin VB.Label lblCalFormula 
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
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   6795
      End
   End
   Begin VB.Frame frmCalGraph 
      Caption         =   "Calibration Graph"
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
      Height          =   6735
      Left            =   7200
      TabIndex        =   150
      Top             =   120
      Width           =   8055
      Begin MSChart20Lib.MSChart chtSclChart 
         Height          =   6435
         Left            =   60
         OleObjectBlob   =   "FrmScalesCal.frx":AFB1
         TabIndex        =   151
         Top             =   180
         Width           =   7920
      End
   End
End
Attribute VB_Name = "frmScalesCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''ERROR module 8118
''
''frmScalesCal
''
Option Explicit

Private blnPrevCalExists As Boolean                       ' Flag - Whether previous calibration data is loaded
' Scales calibration variables
Private SelCalScl As Integer                              ' The Index of the scale currently selected
Private NumCalPoints As Integer                           ' The Number of Data Points for the currently selected Scale
Private SclMaxGrams As Single                             ' The Max Scale Range in grams
Private SclMinGrams As Single                             ' The Min Scale Range in grams
Private Curr_SclCal As SclCalibration
Private New_SclCal As SclCalibration
Private Prev_SclCal As SclCalibration
' Max Scales
Const MAXSCL = 16

' Results from text field validation
Const VALID = 0
Const TOOHIGH = 1
Const TOOLOW = 2
Const NOTNUMERIC = 3
Const NOTANINTEGER = 4
' Result of text box validation test
Private Result As Integer

' Read-Only or Actually Performing a Calibration?  (default is Read-Only)
Private CalReadOnly As Boolean
' various flags
Private bEditNumCalPts As Boolean        ' Flag - Whether the Number of Calibration Points is being edited
Private bNewCalEnabled As Boolean        ' Flag - Whether new calibration point data is loaded
Private bCurrCalEnabled As Boolean       ' Flag - Whether current calibration point data is loaded
Private bPrevCalEnabled As Boolean       ' Flag - Whether previous calibration point data is loaded
Private bUnsavedCal As Boolean           ' Flag - Whether an unsaved calibration exists

Private Sub FillNewCalPointTable()
    ' Define every Scale calibration point
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8118, 1
Dim iPoint As Integer
Dim idx As Integer
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
    
    Select Case NumCalPoints
        Case 3
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(50)
            New_SclCal.PointData(3).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(50)
            New_SclCal.PointData(3).ActualPercent = CSng(100)
        Case 4
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(30)
            New_SclCal.PointData(3).RawPercent = CSng(70)
            New_SclCal.PointData(4).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(30)
            New_SclCal.PointData(3).ActualPercent = CSng(70)
            New_SclCal.PointData(4).ActualPercent = CSng(100)
        Case 5
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(25)
            New_SclCal.PointData(3).RawPercent = CSng(50)
            New_SclCal.PointData(4).RawPercent = CSng(75)
            New_SclCal.PointData(5).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(25)
            New_SclCal.PointData(3).ActualPercent = CSng(50)
            New_SclCal.PointData(4).ActualPercent = CSng(75)
            New_SclCal.PointData(5).ActualPercent = CSng(100)
        Case 6
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(20)
            New_SclCal.PointData(3).RawPercent = CSng(40)
            New_SclCal.PointData(4).RawPercent = CSng(60)
            New_SclCal.PointData(5).RawPercent = CSng(80)
            New_SclCal.PointData(6).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(20)
            New_SclCal.PointData(3).ActualPercent = CSng(40)
            New_SclCal.PointData(4).ActualPercent = CSng(60)
            New_SclCal.PointData(5).ActualPercent = CSng(80)
            New_SclCal.PointData(6).ActualPercent = CSng(100)
        Case 7
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(16.7)
            New_SclCal.PointData(3).RawPercent = CSng(33.3)
            New_SclCal.PointData(4).RawPercent = CSng(50)
            New_SclCal.PointData(5).RawPercent = CSng(66.7)
            New_SclCal.PointData(6).RawPercent = CSng(83.3)
            New_SclCal.PointData(7).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(16.7)
            New_SclCal.PointData(3).ActualPercent = CSng(33.3)
            New_SclCal.PointData(4).ActualPercent = CSng(50)
            New_SclCal.PointData(5).ActualPercent = CSng(66.7)
            New_SclCal.PointData(6).ActualPercent = CSng(83.3)
            New_SclCal.PointData(7).ActualPercent = CSng(100)
        Case 8
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(14.3)
            New_SclCal.PointData(3).RawPercent = CSng(28.6)
            New_SclCal.PointData(4).RawPercent = CSng(42.9)
            New_SclCal.PointData(5).RawPercent = CSng(57.1)
            New_SclCal.PointData(6).RawPercent = CSng(71.4)
            New_SclCal.PointData(7).RawPercent = CSng(85.7)
            New_SclCal.PointData(8).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(14.3)
            New_SclCal.PointData(3).ActualPercent = CSng(28.6)
            New_SclCal.PointData(4).ActualPercent = CSng(42.9)
            New_SclCal.PointData(5).ActualPercent = CSng(57.1)
            New_SclCal.PointData(6).ActualPercent = CSng(71.4)
            New_SclCal.PointData(7).ActualPercent = CSng(85.7)
            New_SclCal.PointData(8).ActualPercent = CSng(100)
        Case 9
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(12.5)
            New_SclCal.PointData(3).RawPercent = CSng(25)
            New_SclCal.PointData(4).RawPercent = CSng(37.5)
            New_SclCal.PointData(5).RawPercent = CSng(50)
            New_SclCal.PointData(6).RawPercent = CSng(62.5)
            New_SclCal.PointData(7).RawPercent = CSng(75)
            New_SclCal.PointData(8).RawPercent = CSng(87.5)
            New_SclCal.PointData(9).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(12.5)
            New_SclCal.PointData(3).ActualPercent = CSng(25)
            New_SclCal.PointData(4).ActualPercent = CSng(37.5)
            New_SclCal.PointData(5).ActualPercent = CSng(50)
            New_SclCal.PointData(6).ActualPercent = CSng(62.5)
            New_SclCal.PointData(7).ActualPercent = CSng(75)
            New_SclCal.PointData(8).ActualPercent = CSng(87.5)
            New_SclCal.PointData(9).ActualPercent = CSng(100)
        Case 10
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(11.1)
            New_SclCal.PointData(3).RawPercent = CSng(22.2)
            New_SclCal.PointData(4).RawPercent = CSng(33.3)
            New_SclCal.PointData(5).RawPercent = CSng(44.4)
            New_SclCal.PointData(6).RawPercent = CSng(55.5)
            New_SclCal.PointData(7).RawPercent = CSng(66.6)
            New_SclCal.PointData(8).RawPercent = CSng(77.7)
            New_SclCal.PointData(9).RawPercent = CSng(88.8)
            New_SclCal.PointData(10).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(11.1)
            New_SclCal.PointData(3).ActualPercent = CSng(22.2)
            New_SclCal.PointData(4).ActualPercent = CSng(33.3)
            New_SclCal.PointData(5).ActualPercent = CSng(44.4)
            New_SclCal.PointData(6).ActualPercent = CSng(55.5)
            New_SclCal.PointData(7).ActualPercent = CSng(66.6)
            New_SclCal.PointData(8).ActualPercent = CSng(77.7)
            New_SclCal.PointData(9).ActualPercent = CSng(88.8)
            New_SclCal.PointData(10).ActualPercent = CSng(100)
        Case 11
            New_SclCal.PointData(1).RawPercent = CSng(0)
            New_SclCal.PointData(2).RawPercent = CSng(10)
            New_SclCal.PointData(3).RawPercent = CSng(20)
            New_SclCal.PointData(4).RawPercent = CSng(30)
            New_SclCal.PointData(5).RawPercent = CSng(40)
            New_SclCal.PointData(6).RawPercent = CSng(50)
            New_SclCal.PointData(7).RawPercent = CSng(60)
            New_SclCal.PointData(8).RawPercent = CSng(70)
            New_SclCal.PointData(9).RawPercent = CSng(80)
            New_SclCal.PointData(10).RawPercent = CSng(90)
            New_SclCal.PointData(11).RawPercent = CSng(100)
            New_SclCal.PointData(1).ActualPercent = CSng(0)
            New_SclCal.PointData(2).ActualPercent = CSng(10)
            New_SclCal.PointData(3).ActualPercent = CSng(20)
            New_SclCal.PointData(4).ActualPercent = CSng(30)
            New_SclCal.PointData(5).ActualPercent = CSng(40)
            New_SclCal.PointData(6).ActualPercent = CSng(50)
            New_SclCal.PointData(7).ActualPercent = CSng(60)
            New_SclCal.PointData(8).ActualPercent = CSng(70)
            New_SclCal.PointData(9).ActualPercent = CSng(80)
            New_SclCal.PointData(10).ActualPercent = CSng(90)
            New_SclCal.PointData(11).ActualPercent = CSng(100)
    End Select
    
    ' get min/max EU & Raw  for appropriate input
    ' Common Scale Calibration Parameters
    sEuMax = SclMaxGrams
    sEuMin = SclMinGrams
    ' calc EU & Vdc spans
    sEuSpan = sEuMax - sEuMin
    
    For iPoint = 1 To MAXLSQCALPOINTS
        New_SclCal.PointData(iPoint).RawValue = sEuMin + (sEuSpan * (New_SclCal.PointData(iPoint).RawPercent / CSng(100)))
        New_SclCal.PointData(iPoint).ActualValue = sEuMin + (sEuSpan * (New_SclCal.PointData(iPoint).ActualPercent / CSng(100)))
    Next iPoint
    
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

Private Sub ClearSclCal(ByRef tmpCal As SclCalibration)
Dim iPoint As Integer

    ' Set Scale Calibration to Linear (default)
    ' set calibration parameters
    tmpCal.Dts = Now
    tmpCal.CalibratedBy = "na"
    tmpCal.Comment = "cleared"
    tmpCal.CalRangeMax = CSng(DefScaleMax)
    tmpCal.CalRangeMin = CSng(0)
    tmpCal.NumPoints = MINLSQCALPOINTS
    '            tmpCal.CalData.X = sEuMax
    tmpCal.CalData.X = CSng(0)
    tmpCal.CalData.X2 = CSng(0)
    tmpCal.CalData.X3 = CSng(0)
    tmpCal.CalData.X4 = CSng(0)
    tmpCal.CalData.X5 = CSng(0)
    tmpCal.CalData.X6 = CSng(0)
    tmpCal.CalData.R2 = CSng(0)
    
    ' set Scale Calibration Point Data
    For iPoint = 1 To MAXLSQCALPOINTS
        tmpCal.PointData(iPoint).ActualPercent = CSng(0)
        tmpCal.PointData(iPoint).RawPercent = CSng(0)
        tmpCal.PointData(iPoint).ActualValue = CSng(0)
        tmpCal.PointData(iPoint).RawValue = CSng(0)
    Next iPoint
            
End Sub

Private Sub cmdApply_Click()
Dim idx As Integer
Dim iPoint As Integer
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim tmpVal As Single
Dim tmpval2 As Single
Dim flag As Boolean

    ' valid value in NumCalPts text box ??
    flag = RangeCheck(MAXLSQCALPOINTS, MINLSQCALPOINTS, txtNumCalPts, "Number of Calibration Points")
    If flag Then
        
        ' setup screen for different NumberOfCalibrationPoints
        If (CInt(ValueFromText(txtNumCalPts.text)) <> NumCalPoints) Then
            ' Number of Calibration Points for this Scale has been changed
            NumCalPoints = CInt(ValueFromText(txtNumCalPts.text))
            ClearSclCal New_SclCal
            ClearSclCal Curr_SclCal
            ClearSclCal Prev_SclCal
            EnableNewCal False
            EnablePrevCal False
            ' Set Scale Calibration to Linear (default)
            ' get min/max EU
            sEuMax = SclMaxGrams
            sEuMin = SclMinGrams
            
            ' calc EU & Vdc spans
            sEuSpan = sEuMax - sEuMin
            ' set calibration parameters
            Curr_SclCal.Dts = Now()
            Curr_SclCal.CalibratedBy = "default"
            Curr_SclCal.Equipment = "equipment"
            Curr_SclCal.Comment = "linear"
            Curr_SclCal.NumPoints = NumCalPoints
            Curr_SclCal.CalRangeMax = SclMaxGrams
            Curr_SclCal.CalRangeMin = SclMinGrams
    '            New_SclCal.CalData.X = sEuMax
            Curr_SclCal.CalData.X = CSng(1)
            Curr_SclCal.CalData.X2 = CSng(0)
            Curr_SclCal.CalData.X3 = CSng(0)
            Curr_SclCal.CalData.X4 = CSng(0)
            Curr_SclCal.CalData.X5 = CSng(0)
            Curr_SclCal.CalData.X6 = CSng(0)
            Curr_SclCal.CalData.R2 = CSng(0)
                
            ' set Scale Calibration Point Data
            For iPoint = 1 To MAXLSQCALPOINTS
                idx = iPoint - 1
                tmpVal = CSng(idx) * CSng(10)
                tmpval2 = tmpVal / CSng(100)
                Curr_SclCal.PointData(iPoint).ActualPercent = tmpVal
                Curr_SclCal.PointData(iPoint).RawPercent = tmpVal
                Curr_SclCal.PointData(iPoint).ActualValue = sEuMin + (tmpval2 * sEuSpan)
                Curr_SclCal.PointData(iPoint).RawValue = sEuMin + (tmpval2 * sEuSpan)
            Next iPoint
            New_SclCal = Curr_SclCal
        End If
        bEditNumCalPts = False
        DisplaySclAll
        
    End If
    
End Sub

Private Sub cmdCalPoints_Click()
'
Dim flag As Boolean

    Select Case bEditNumCalPts
        Case True
            ' cancel the edit
            flag = False
            ' reset the contents of the text box
            txtNumCalPts.text = Format(NumCalPoints, "#0")
        Case False
            ' start an edit
            flag = True
    End Select
    bEditNumCalPts = flag
    UpdateCmdButtons
End Sub

Private Sub cmdCreateNewCalib_Click()
    ClearSclCal New_SclCal
    SclMaxGrams = CInt(ValueFromText(txtSclEuMax.text))
    SclMinGrams = CInt(ValueFromText(txtSclEuMin.text))
    New_SclCal.NumPoints = NumCalPoints
    New_SclCal.CalRangeMax = SclMaxGrams
    New_SclCal.CalRangeMin = SclMinGrams
    New_SclCal.Dts = Now()
    New_SclCal.StandardTempValue = 20
    New_SclCal.StandardTempUnits = "deg C"
    New_SclCal.StandardPressValue = 1
    New_SclCal.StandardPressUnits = "atm"
    New_SclCal.CalibratedBy = "default"
    New_SclCal.Equipment = "equipment"
    New_SclCal.Comment = "linear"
    EnableNewCal True
    FillNewCalPointTable
    DisplayNewCalInformation
    DisplayCalPointData
    UpdateCmdButtons
End Sub

Private Sub cmdScaleDn_Click()
    SelCalScl = SelCalScl - 1
    If SelCalScl < 1 Then SelCalScl = NR_SCALES
    bUnsavedCal = False
    ' Update the Display
    UpdateSclSelection
    DisplaySclAll
End Sub

Private Sub cmdScaleUp_Click()
    SelCalScl = SelCalScl + 1
    If SelCalScl > NR_SCALES Then SelCalScl = 1
    bUnsavedCal = False
    ' Update the Display
    UpdateSclSelection
    DisplaySclAll
End Sub

Private Sub cmdRunCal_Click()
   
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8118, 3
Dim row As Integer
Dim iPoint As Integer
Dim tmpPercent As Single
Dim tmpPercent2 As Single
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim ErrorExists As Boolean

    ' read Scale information from screen
    sEuMax = SclMaxGrams
    sEuMin = SclMinGrams
    sEuSpan = sEuMax - sEuMin
    
    ' Validate the new actual value boxes
    ErrorExists = False
    ' check all boxes
    For row = 1 To NumCalPoints
        If (Not RangeCheck(sEuMax, sEuMin, txtNewActualValue(row), "Actual Value Entry")) Then ErrorExists = True
    Next row
    ' all boxes valid ??
    If Not ErrorExists Then
        ' all entries are valid; calibrate
        ' update new calibration DTS
        txtCalibDts.text = Format(Now, "YYYY-MMM-DD  hh:mm:ss")
        ' update new calibration from screen
        New_SclCal.NumPoints = CInt(txtNumCalPts.text)
        New_SclCal.CalRangeMax = ValueFromText(txtSclEuMax.text)
        New_SclCal.CalRangeMin = ValueFromText(txtSclEuMin.text)
        New_SclCal.Dts = CDate(txtCalibDts.text)
        New_SclCal.StandardTempValue = ValueFromText(txtStandardTemp.text)
        New_SclCal.StandardTempUnits = TempUnits.List(TempUnits.ListIndex)
        New_SclCal.StandardPressValue = ValueFromText(txtStandardPress.text)
        New_SclCal.StandardPressUnits = TempUnits.List(PressUnits.ListIndex)
        New_SclCal.CalibratedBy = txtCalibBy.text
        New_SclCal.Comment = txtComment.text
        ' update new calibration point data from screen
        For iPoint = 1 To NumCalPoints
            New_SclCal.PointData(iPoint).RawValue = ValueFromText(txtNewRawValue(iPoint).text)
            New_SclCal.PointData(iPoint).ActualValue = ValueFromText(txtNewActualValue(iPoint).text)
            tmpPercent = CSng(100) * ((New_SclCal.PointData(iPoint).RawValue - sEuMin) / sEuSpan)
            New_SclCal.PointData(iPoint).RawPercent = tmpPercent
            tmpPercent2 = CSng(100) * ((New_SclCal.PointData(iPoint).ActualValue - sEuMin) / sEuSpan)
            New_SclCal.PointData(iPoint).ActualPercent = tmpPercent2
        Next iPoint
    
        ' Copy Current to Prev
        If (Not bPrevCalEnabled) Then Prev_SclCal = Curr_SclCal
        ' Calculate Calibration Coefficients
        Calibrate
        ' Copy New to Current
        Curr_SclCal = New_SclCal
            
    '    ClearSclCal New_SclCal
        EnableNewCal True
        EnablePrevCal True
        bUnsavedCal = True
        blnPrevCalExists = True
        DisplaySclAll
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "Calibration Done"
        
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

' copy data from the form to the worksheet
' The Excel workbook
' (Used for Least-Squares calculations)
Dim xlApp As Object
Dim xlWB As Object
Dim xlSht As Object

Dim i, RangeOffset As Integer
Dim InputDataRange, FormulaRange As Range
Dim InputRangeText, XRangeText, YRangeText As String
Dim LINESTFormulaText As String

    RangeOffset = NumCalPoints + 1
    
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
    
    For i = 1 To NumCalPoints
        xlSht.Range("A" & i + 1) = (New_SclCal.PointData(i).RawPercent / CSng(100))
        xlSht.Range("B" & i + 1) = (New_SclCal.PointData(i).RawPercent / CSng(100)) ^ 2
        xlSht.Range("C" & i + 1) = (New_SclCal.PointData(i).RawPercent / CSng(100)) ^ 3
        xlSht.Range("D" & i + 1) = (New_SclCal.PointData(i).RawPercent / CSng(100)) ^ 4
        xlSht.Range("E" & i + 1) = (New_SclCal.PointData(i).RawPercent / CSng(100)) ^ 5
        xlSht.Range("F" & i + 1) = (New_SclCal.PointData(i).RawPercent / CSng(100)) ^ 6
        xlSht.Range("G" & i + 1) = New_SclCal.PointData(i).ActualPercent / CSng(100)
    Next i
    
    ' Select cells for the array formula
    Set FormulaRange = xlSht.Range("A14:F18")
    ' FormulaRange.Select
    
    ' Enter the LINEST formula
    InputRangeText = "A2:G" & RangeOffset
    XRangeText = "R2C1:R" & RangeOffset & "C6"
    YRangeText = "R2C7:R" & RangeOffset & "C7"
    Set InputDataRange = xlSht.Range(InputRangeText)
    LINESTFormulaText = "=LINEST(" & YRangeText & "," & XRangeText & ", FALSE, TRUE)"
    FormulaRange.FormulaArray = LINESTFormulaText
    FormulaRange.Calculate
    'MsgBox xlWB.Sheets(1).Range("A14").value
'Exit Sub
    ' Assign the coefficients to variables
    New_SclCal.CalData.X = xlSht.Range("F14").Value
    New_SclCal.CalData.X2 = xlSht.Range("E14").Value
    New_SclCal.CalData.X3 = xlSht.Range("D14").Value
    New_SclCal.CalData.X4 = xlSht.Range("C14").Value
    New_SclCal.CalData.X5 = xlSht.Range("B14").Value
    New_SclCal.CalData.X6 = xlSht.Range("A14").Value
    
    New_SclCal.CalData.R2 = xlSht.Range("A16").Value

    ' make sure xlWB is closed
    xlWB.Saved = True
    xlWB.Close
    Set xlWB = Nothing
    ' Close Excel
    xlApp.Quit
    Set xlApp = Nothing
    
End Sub

Private Sub DisplayCalFormula()
    ' Displays the calibration formula on the screen
    Dim FormulaText As String
    FormulaText = vbCrLf
    FormulaText = FormulaText & "Y = " & Curr_SclCal.CalData.X6 & "X6"
    FormulaText = FormulaText & IIf(Curr_SclCal.CalData.X5 < 0, " - ", " + ") & Abs(Curr_SclCal.CalData.X5) & "X5"
    FormulaText = FormulaText & IIf(Curr_SclCal.CalData.X4 < 0, " - ", " + ") & Abs(Curr_SclCal.CalData.X4) & "X4"
    FormulaText = FormulaText & IIf(Curr_SclCal.CalData.X3 < 0, " - ", " + ") & Abs(Curr_SclCal.CalData.X3) & "X3"
    FormulaText = FormulaText & IIf(Curr_SclCal.CalData.X2 < 0, " - ", " + ") & Abs(Curr_SclCal.CalData.X2) & "X2"
    FormulaText = FormulaText & IIf(Curr_SclCal.CalData.X < 0, " - ", " + ") & Abs(Curr_SclCal.CalData.X) & "X"
    FormulaText = FormulaText & vbCrLf
    FormulaText = FormulaText & "      R2=" & Curr_SclCal.CalData.R2
    lblCalFormula.Visible = True
    lblCalFormula.ForeColor = White
    lblCalFormula.Caption = FormulaText
End Sub

Private Sub DisplayCalGraph()
Dim ChartArray(MAXLSQCALPOINTS, 2)
Dim idx As Integer
Dim iPoint As Integer
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim tmpVal As Single
Dim Graph() As Single
         
    ReDim Graph(NumCalPoints, 1 To 6)
    
    ' get min/max EU & Raw
    sEuMax = SclMaxGrams
    sEuMin = SclMinGrams
    sEuSpan = sEuMax - sEuMin
    
    For iPoint = 1 To NumCalPoints
        Graph(iPoint, 1) = Curr_SclCal.PointData(iPoint).RawPercent      ' value for X-axis
        Graph(iPoint, 2) = CSng(Curr_SclCal.PointData(iPoint).RawPercent)
        Graph(iPoint, 3) = Curr_SclCal.PointData(iPoint).RawPercent      ' value for X-axis
        Graph(iPoint, 4) = CSng(Curr_SclCal.PointData(iPoint).ActualPercent)
        Graph(iPoint, 5) = Curr_SclCal.PointData(iPoint).RawPercent      ' value for X-axis
        tmpVal = Cal_Scale(Curr_SclCal.PointData(iPoint).ActualPercent / 100, SelCalScl, Curr_SclCal)
        Graph(iPoint, 6) = CSng(100) * ((tmpVal - sEuMin) / sEuSpan)    ' value for Y-axis
    Next iPoint
         
    chtSclChart.chartType = VtChChartType2dXY  ' set to X Y Scatter chart
    chtSclChart = Graph ' populate chart's data grid using Graph array
    chtSclChart.Plot.UniformAxis = False
    chtSclChart.Column = 1
    chtSclChart.ColumnLabel = "Raw Value"
    chtSclChart.Column = 3
    chtSclChart.ColumnLabel = "Actual Value"
    chtSclChart.Column = 5
    chtSclChart.ColumnLabel = "Calib. Value"
    chtSclChart.Visible = True
End Sub

Private Sub cmdRestorePrevCal_Click()
    ' Undoes the effect of the calibration by transferring data from
    ' the previous calibration to the current calibration, transferring data
    ' from the temporary calibration buffer to the previous calibration,
    ' and updating the form
    
    ' shift cal point data from prev to current
'    ClearSclCal New_SclCal
    New_SclCal = Curr_SclCal
'    ClearSclCal Curr_SclCal
    Curr_SclCal = Prev_SclCal
    ClearSclCal Prev_SclCal
    EnablePrevCal False
    bUnsavedCal = False
    ' update display
    DisplaySclCalibration
    UpdateCmdButtons
'    Delay_Box "Returning to Previously Saved Cal.", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "Previous Calibration Restored"
    
End Sub

Private Sub cmdPrintCal_Click()
    PrintCalibration
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = "File Released to the Printer"
End Sub

Private Sub UpdateCmdButtons()
' update the calibration command buttons
'
' bEditNumCalPoints - Whether the Number of Calibration Points is being edited
' bNewCalEnabled - Whether new calibration point data is loaded
' bCurrCalEnabled - Whether current calibration point data is loaded
' bPrevCalEnabled - Whether previous calibration point data is loaded
' bUnsavedCal - Whether an unsaved calibration exists
    
    If bEditNumCalPts Then
        cmdScaleDn.Enabled = True
        cmdScaleUp.Enabled = True
        cmdApply.Visible = True
        txtNumCalPts.ForeColor = TitlesData_Forecolor
        txtNumCalPts.BackColor = White
        cmdCalPoints.Caption = "Cancel Edit"
        cmdCalPoints.Enabled = True
        cmdCreateNewCalib.Enabled = False
        cmdRunCal.Enabled = False
        cmdSaveCurrCal.Enabled = False
        cmdRestorePrevCal.Enabled = False
        cmdPrintCal.Enabled = False
    Else
        cmdScaleDn.Enabled = True
        cmdScaleUp.Enabled = True
        cmdApply.Visible = False
        txtNumCalPts.ForeColor = Black
        txtNumCalPts.BackColor = frmCalInformation.BackColor
        cmdCalPoints.Caption = "Edit CalPoints"
        cmdCalPoints.Enabled = Not CalReadOnly
        cmdCreateNewCalib.Enabled = Not CalReadOnly
        cmdRunCal.Enabled = bNewCalEnabled
        cmdSaveCurrCal.Enabled = bUnsavedCal
        cmdRestorePrevCal.Enabled = bPrevCalEnabled
        cmdPrintCal.Enabled = IIf(PRINTERAVAILABLE, True, False)
    End If
End Sub

Private Sub UpdateSclSelection()
Dim idx As Integer
    ClearSclCal New_SclCal
    ClearSclCal Curr_SclCal
    ClearSclCal Prev_SclCal
    Curr_SclCal = Scale_Cal(SelCalScl)
    NumCalPoints = Curr_SclCal.NumPoints
    ' init cal point columns
    EnableNewCal False
    EnablePrevCal False
End Sub

Private Sub DisplaySclAll()
    ' update the screen
    DisplaySclSelection
    HideInactiveTableRows
    DisplaySclCalibration
    UpdateCmdButtons
End Sub

Private Sub DisplaySclCalibration()
'
    ' Cal is ReadOnly unless all stations are idle
    If AllStationsIdle And CalReadOnly Then
        CalReadOnly = False
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "Scale Calibration is Enabled"
    ElseIf Not AllStationsIdle And Not CalReadOnly Then
        CalReadOnly = True
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "Scale Calibration is Read-Only"
    End If
    ' update the screen
    DisplayCurrCalInformation
    DisplayCalPointData
    DisplayCalFormula
    DisplayCalGraph
    
End Sub

Private Sub cmdSaveCurrCal_Click()
    ' Save the current calibration
    SaveCurrCalibration SelCalScl
    bUnsavedCal = False
    ' clear new & prev calibrations
    ClearSclCal New_SclCal
    ClearSclCal Prev_SclCal
    EnableNewCal False
    EnablePrevCal False
    blnPrevCalExists = False
    ' update screen
    DisplaySclCalibration
    ' update command buttons
    UpdateCmdButtons
'    Delay_Box "Calibration Saved", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = vbCrLf & "Calibration Saved"
End Sub

Private Sub Form_Load()
'
Dim idx As Integer
    ' set display colors
    frmCalControls.ForeColor = Titles_ForeColor
    frmScaleSelection.ForeColor = Titles_ForeColor
    frmScaleInformation.ForeColor = Titles_ForeColor
    frmCalFormula.ForeColor = Titles_ForeColor
    frmCalInformation.ForeColor = Titles_ForeColor
    frmCalPointData.ForeColor = Titles_ForeColor
    frmCalGraph.ForeColor = Titles_ForeColor
    txtNumCalPts.ForeColor = TitlesData_Forecolor
    txtDispScl.ForeColor = TitlesData_Forecolor
    txtCalibBy.ForeColor = TitlesData_Forecolor
    txtEquipment.ForeColor = TitlesData_Forecolor
    txtComment.ForeColor = TitlesData_Forecolor
    For idx = 1 To MAXLSQCALPOINTS
        txtNewRawValue(idx).ForeColor = TitlesData_Forecolor
        txtNewRawValue(idx).BackColor = Entry_BackColor
        txtNewActualValue(idx).ForeColor = TitlesData_Forecolor
        txtNewActualValue(idx).BackColor = Entry_BackColor
    Next idx

    ' Cal is ReadOnly unless all stations are idle
    If AllStationsIdle Then
        CalReadOnly = False
        txtMsg.ForeColor = Message_ForeColor
        txtMsg.text = vbCrLf & "Calibration is Enabled"
    Else
        CalReadOnly = True
        txtMsg.ForeColor = Message_ForeColor
        txtMsg.text = vbCrLf & "Calibration is Read-Only"
    End If

    ' Set the current group and input
    SelCalScl = 1
    ' init cal point columns
    EnableNewCal False
    EnablePrevCal False
    ' init Unsaved Calibration flag
    bUnsavedCal = False
    blnPrevCalExists = False
    ' update the screen
    UpdateSclSelection
    DisplaySclAll
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub HideInactiveTableRows()
' Hide calibration table rows
' in excess of the current number of points
Dim flag As Boolean
Dim iRow As Integer
    For iRow = 1 To MAXLSQCALPOINTS
        ' Visible or NOT ??
        flag = IIf((iRow <= NumCalPoints), True, False)
        ' Set every cell in the row
        lblPointNum(iRow).Visible = flag
        txtNewRawValue(iRow).Visible = flag
        lblNewRawPerc(iRow).Visible = flag
        txtNewActualValue(iRow).Visible = flag
        lblCurrRawPerc(iRow).Visible = flag
        lblCurrActualValue(iRow).Visible = flag
        lblCurrCalValue(iRow).Visible = flag
        lblCurrDiff(iRow).Visible = flag
        lblPrevRawPerc(iRow).Visible = flag
        lblPrevActualValue(iRow).Visible = flag
        lblPrevCalValue(iRow).Visible = flag
        lblPrevDiff(iRow).Visible = flag
    Next iRow
End Sub

Private Sub EnableNewCal(ByVal flagEnable As Boolean)
' Enable (or Not) new calibration table columns
'
Dim iPoint As Integer
    For iPoint = 1 To MAXLSQCALPOINTS
        ' Enabled or NOT ??
        txtNewRawValue(iPoint).Enabled = flagEnable
        lblNewRawPerc(iPoint).Enabled = flagEnable
        txtNewActualValue(iPoint).Enabled = flagEnable
    Next iPoint
    bNewCalEnabled = flagEnable
End Sub

Private Sub EnablePrevCal(ByVal flagEnable As Boolean)
' Enable (or Not) prev calibration table columns
'
Dim iPoint As Integer
    For iPoint = 1 To MAXLSQCALPOINTS
        ' Enabled or NOT ??
        lblPrevRawPerc(iPoint).Enabled = flagEnable
        lblPrevActualValue(iPoint).Enabled = flagEnable
        lblPrevCalValue(iPoint).Enabled = flagEnable
        lblPrevDiff(iPoint).Enabled = flagEnable
    Next iPoint
    bPrevCalEnabled = flagEnable
End Sub

Private Sub DisplaySclSelection()
' DisplayAISelection
' Displays Information on the Currently Selected Scale
    
    txtDispScl.text = "Scale #" & Format(SelCalScl, "#0")
    
End Sub

Private Sub txtCalibBy_Change()
'    bUnsavedCal = True
End Sub

Private Sub txtComment_Change()
'    bUnsavedCal = True
End Sub

Private Sub DisplayCurrCalValues()
' Display Calibration Results
' Fills the current calibration table with the calibration results
Dim iPoint As Integer
Dim ActualValue As Single
Dim CalibValue As Single
Dim percDiff As Single

    For iPoint = 1 To NumCalPoints
        ' Actual Value
        ActualValue = Curr_SclCal.PointData(iPoint).ActualValue
        ' Calibrated Value
        CalibValue = Cal_Scale((Curr_SclCal.PointData(iPoint).ActualPercent / CSng(100)), SelCalScl, Curr_SclCal)
        lblCurrCalValue(iPoint).Caption = IIf(lblCurrCalValue(iPoint).Enabled, Format(CalibValue, "####0.0##"), "")
        ' Percent Difference
        If ActualValue > 0! Then
            percDiff = ((CalibValue - ActualValue) / ActualValue) * 100
        Else
            percDiff = 0!
        End If
        lblCurrDiff(iPoint).Caption = IIf(lblCurrDiff(iPoint).Enabled, Format(percDiff, "####0.0##"), "")
'        lblCurrDiff(iPoint).Caption = Format(PercDiff, "###0.0") & "%"
    Next iPoint
End Sub

Private Sub DisplayPrevCalValues()
' Display Calibration Results
' Fills the prev calibration table with the calibration results
Dim iPoint As Integer
Dim ActualValue As Single
Dim CalibValue As Single
Dim percDiff As Single

    For iPoint = 1 To NumCalPoints
        ' Actual Value
        ActualValue = Prev_SclCal.PointData(iPoint).ActualValue
        ' Calibrated Value
        CalibValue = Cal_Scale((Prev_SclCal.PointData(iPoint).ActualPercent / CSng(100)), SelCalScl, Prev_SclCal)
        lblPrevCalValue(iPoint).Caption = IIf(lblPrevCalValue(iPoint).Enabled, Format(CalibValue, "####0.0##"), "")
        ' Percent Difference
        If ActualValue > 0! Then
            percDiff = ((CalibValue - ActualValue) / ActualValue) * 100
        Else
            percDiff = 0!
        End If
        lblPrevDiff(iPoint).Caption = IIf(lblPrevDiff(iPoint).Enabled, Format(percDiff, "####0.0##"), "")
'        lblPrevDiff(iPoint).Caption = Format(PercDiff, "###0.0") & "%"
    Next iPoint
End Sub

Public Sub SaveCurrCalibration(ByVal iScale As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date
Dim iPoint As Integer

    ' delete existing calibration
    ClearSclCalRecords iScale

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Save Scale Calibration Parameters
    Criteria = "SELECT * FROM [ScaleCalibrations] WHERE [Scale] = " & iScale & "  ORDER BY [Dts] DESC"
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' any records found ??
    If rsRecord.BOF Then
        ' no records; add a new one
        rsRecord.AddNew
    Else
        ' record exists; edit it
        rsRecord.Edit
    End If
    ' update the calibration information
    rsRecord("Scale") = iScale
    rsRecord("Dts") = Curr_SclCal.Dts
    rsRecord("CalibratedBy") = Curr_SclCal.CalibratedBy
    rsRecord("Equipment") = Curr_SclCal.Equipment
    rsRecord("Comment") = Curr_SclCal.Comment
    rsRecord("NumPoints") = Curr_SclCal.NumPoints
    rsRecord("CalRangeMax") = Curr_SclCal.CalRangeMax
    rsRecord("CalRangeMin") = Curr_SclCal.CalRangeMin
    rsRecord("CoefficientX") = Curr_SclCal.CalData.X
    rsRecord("CoefficientX2") = Curr_SclCal.CalData.X2
    rsRecord("CoefficientX3") = Curr_SclCal.CalData.X3
    rsRecord("CoefficientX4") = Curr_SclCal.CalData.X4
    rsRecord("CoefficientX5") = Curr_SclCal.CalData.X5
    rsRecord("CoefficientX6") = Curr_SclCal.CalData.X6
    rsRecord("CoefficientR2") = Curr_SclCal.CalData.R2
    rsRecord("StandardTempValue") = Curr_SclCal.StandardTempValue
    rsRecord("StandardTempUnits") = Curr_SclCal.StandardTempUnits
    rsRecord("StandardPressValue") = Curr_SclCal.StandardPressValue
    rsRecord("StandardPressUnits") = Curr_SclCal.StandardPressUnits
    rsRecord.Update
    rsRecord.Close

            
    ' Save Scale Calibration Point Data
    dDts = Curr_SclCal.Dts
    CriteriaPts = "SELECT * FROM [ScaleCalibrationsData] WHERE [Scale] = " & iScale & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
    Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
            
    ' update the calibration information
    For iPoint = 1 To NumCalPoints
        rsRecordPts.AddNew
        rsRecordPts("Scale") = iScale
        rsRecordPts("Point") = iPoint
        rsRecordPts("Dts") = dDts
        rsRecordPts("ActualPercent") = Curr_SclCal.PointData(iPoint).ActualPercent
        rsRecordPts("RawPercent") = Curr_SclCal.PointData(iPoint).RawPercent
        rsRecordPts("ActualValue") = Curr_SclCal.PointData(iPoint).ActualValue
        rsRecordPts("RawValue") = Curr_SclCal.PointData(iPoint).RawValue
        rsRecordPts.Update
    Next iPoint
    ' done with points
    rsRecordPts.Close
                        
    ' copy calibration to appropriate scale
    PrevScale_Cal(iScale) = Scale_Cal(iScale)
    Scale_Cal(iScale) = Curr_SclCal
                
End Sub

Private Sub PrintCalibration()
' Print the values found for the current and (optionally) the previous calibrations
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8118, 22
Dim tmpString(0 To 49) As String
Dim descCol As Integer
Dim dataLeft As Integer
Dim dataRight As Integer
Dim dataWidth As Integer
Dim iLine As Integer
Dim commentWidth As Integer
Dim commentLength As Integer
Dim currComment As String
Dim row As Integer
Dim intNumRows As Integer
Dim intNumCurrRows As Integer
Dim intNumPrevRows As Integer
Dim sngActualValue As Single
Dim sngCalibValue As Single
Dim sngPercDiff As Single
Dim CurrCal As SclCalibration
Dim PrevCal As SclCalibration
Dim bPrintPrevCal As Boolean
Dim oldFont As New StdFont


    ' current calibration
    CurrCal = Scale_Cal(SelCalScl)
    PrevCal = PrevScale_Cal(SelCalScl)
    
    ' previous cal empty??
    bPrintPrevCal = IIf(((PrevCal.CalibratedBy = "na") And (PrevCal.Comment = "cleared")), False, True)

    ' Save current printer font
    oldFont = Printer.Font
    Printer.Font = FILEFONT
    Printer.Font.Size = 8.5  ' FILEFONTSIZE
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    
    ' setup columns
    descCol = IIf(bPrintPrevCal, 5, 20)
    dataLeft = IIf(bPrintPrevCal, 29, 50)
    dataRight = IIf(bPrintPrevCal, 73, 99)
    dataWidth = IIf(bPrintPrevCal, 12, 14)
    commentWidth = IIf(bPrintPrevCal, 40, 50)


    intNumCurrRows = CurrCal.NumPoints
    intNumPrevRows = PrevCal.NumPoints
    intNumRows = IIf((intNumCurrRows > intNumPrevRows), intNumCurrRows, intNumPrevRows)
    
    ' Print titles / header
    Print_Center Trim(SysConfig.Heading)
    Print_Center Trim(SysConfig.Heading2)
    Print_Center "CANISTER PRECONDITIONING SYSTEM"
    Print_Line ""
    Print_Center "Calibration Report for Scale # " & Format(SelCalScl, "#0")
    Print_Center ("Date: " & Format(Now, "mmm d yyyy"))
    Print_Line ""
    Print_Line ""
    Print_Line ""
    
    ' left pad tmpString's
    For row = 0 To 6
        tmpString(row) = Space(descCol)
    Next row

    ' ROW DESCRIPTIONS
    tmpString(0) = tmpString(0) & "                      "
    ' Calibration DateTime
    tmpString(1) = tmpString(1) & "Calibration Date     :"
    ' Calibrated By
    tmpString(2) = tmpString(2) & "Calibrated By        :"
    ' Calibrated By
    tmpString(3) = tmpString(3) & "Equipment            :"
    ' Standard Pressure
    tmpString(4) = tmpString(4) & "Standard Pressure    :"
    ' Standard Temperature
    tmpString(5) = tmpString(5) & "Standard Temperature :"
    ' Comments
    tmpString(6) = tmpString(6) & "Comments             :"
    ' right pad tmpString's
    For row = 0 To 6
        tmpString(row) = tmpString(row) & Space(dataLeft - Len(tmpString(row)))
    Next row
   
    ' CURRENT CALIBRATION
    tmpString(0) = tmpString(0) & "CURRENT CALIBRATION"
    ' Calibration DateTime
    tmpString(1) = tmpString(1) & Format(CurrCal.Dts, "YYYY-MMM-DD")
    ' Calibrated By
    tmpString(2) = tmpString(2) & CurrCal.CalibratedBy
    ' Calibrated By
    tmpString(3) = tmpString(3) & CurrCal.Equipment
    ' Standard Pressure
    tmpString(4) = tmpString(4) & Format(CurrCal.StandardPressValue, "###0.0#") & " " & CurrCal.StandardPressUnits
    ' Standard Temperature
    tmpString(5) = tmpString(5) & Format(CurrCal.StandardTempValue, "##0.0#") & " " & CurrCal.StandardTempUnits

    If bPrintPrevCal Then
        ' right pad tmpString's
        For row = 0 To 6
            tmpString(row) = tmpString(row) & Space(dataRight - Len(tmpString(row)))
        Next row
        ' PREVIOUS CALIBRATION
        tmpString(0) = tmpString(0) & "PREVIOUS CALIBRATION"
        ' Calibration DateTime
        tmpString(1) = tmpString(1) & Format(PrevCal.Dts, "YYYY-MMM-DD")
        ' Calibrated By
        tmpString(2) = tmpString(2) & PrevCal.CalibratedBy
        ' Calibrated By
        tmpString(3) = tmpString(3) & PrevCal.Equipment
        ' Standard Pressure
        tmpString(4) = tmpString(4) & Format(PrevCal.StandardPressValue, "###0.0#") & " " & PrevCal.StandardPressUnits
        ' Standard Temperature
        tmpString(5) = tmpString(5) & Format(PrevCal.StandardTempValue, "##0.0#") & " " & PrevCal.StandardTempUnits
    End If

    ' COMMENTS
    ' left pad tmpString's
    For row = 7 To 16
        tmpString(row) = Space(dataLeft)
    Next row
    ' CURRENT COMMENTS
CurrCal.Comment = "12345678901234567890123456789012345678901234567890"
CurrCal.Comment = CurrCal.Comment & "12345678901234567890123456789012345678901234567890"
CurrCal.Comment = CurrCal.Comment & "1234567"
PrevCal.Comment = CurrCal.Comment & "89"
    currComment = CurrCal.Comment
    commentLength = Len(currComment)
    iLine = 7
    Do While commentLength > commentWidth
        tmpString(iLine) = tmpString(iLine) & Mid(currComment, 1, commentWidth)
        currComment = Mid(currComment, (commentWidth + 1), (commentLength - commentWidth))
        commentLength = Len(currComment)
        iLine = iLine + 1
    Loop
    tmpString(iLine) = tmpString(iLine) & currComment
    
    If bPrintPrevCal Then
        ' right pad tmpString's
        For row = 7 To 16
            tmpString(row) = tmpString(row) & Space(dataRight - Len(tmpString(row)))
        Next row
        ' PREVIOUS COMMENTS
        currComment = PrevCal.Comment
        commentLength = Len(currComment)
        iLine = 7
        Do While commentLength > commentWidth
            tmpString(iLine) = tmpString(iLine) & Mid(currComment, 1, commentWidth)
            currComment = Mid(currComment, (commentWidth + 1), (commentLength - commentWidth))
            commentLength = Len(currComment)
            iLine = iLine + 1
        Loop
        tmpString(iLine) = tmpString(iLine) & currComment
    End If
        
    ' PRINT tmpString's
    ' PRINT tmpString's
    ' PRINT tmpString's
    Print_Line tmpString(0)
    Print_Line ""
    Print_Line ""
    For row = 1 To 6
        Print_Line tmpString(row)
    Next row
    For row = 7 To 16
        If Len(Trim(tmpString(row))) > 1 Then
            Print_Line tmpString(row)
        End If
    Next row
    
    
    Print_Line ""
    Print_Line ""
    Print_Line ""
    ' Print "Calibration Formula", centered
    Print_Center "Calibration Formula Coefficients"
    Print_Center "EU Value = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx"
    Print_Center "Where x = % Full Scale"
    Print_Line ""
    
    ' left pad tmpString's
    For row = 0 To 7
        tmpString(row) = Space(descCol)
    Next row

    ' ROW DESCRIPTIONS
    tmpString(0) = tmpString(0) & "coefficient  A       :"
    tmpString(1) = tmpString(1) & "coefficient  B       :"
    tmpString(2) = tmpString(2) & "coefficient  C       :"
    tmpString(3) = tmpString(3) & "coefficient  D       :"
    tmpString(4) = tmpString(4) & "coefficient  E       :"
    tmpString(5) = tmpString(5) & "coefficient  F       :"
    tmpString(6) = tmpString(6) & "                      "
    tmpString(7) = tmpString(7) & "coefficient r2       :"
    ' right pad tmpString's
    For row = 0 To 7
        tmpString(row) = tmpString(row) & Space(dataLeft - Len(tmpString(row)))
    Next row
    
    ' current calibration
    tmpString(0) = tmpString(0) & CurrCal.CalData.X6
    tmpString(1) = tmpString(1) & CurrCal.CalData.X5
    tmpString(2) = tmpString(2) & CurrCal.CalData.X4
    tmpString(3) = tmpString(3) & CurrCal.CalData.X3
    tmpString(4) = tmpString(4) & CurrCal.CalData.X2
    tmpString(5) = tmpString(5) & CurrCal.CalData.X
    tmpString(6) = tmpString(6) & ""
    tmpString(7) = tmpString(7) & CurrCal.CalData.R2
    
    If bPrintPrevCal Then
        ' right pad tmpString's
        For row = 0 To 7
            tmpString(row) = tmpString(row) & Space(dataRight - Len(tmpString(row)))
        Next row
        ' previous calibration
        tmpString(0) = tmpString(0) & PrevCal.CalData.X6
        tmpString(1) = tmpString(1) & PrevCal.CalData.X5
        tmpString(2) = tmpString(2) & PrevCal.CalData.X4
        tmpString(3) = tmpString(3) & PrevCal.CalData.X3
        tmpString(4) = tmpString(4) & PrevCal.CalData.X2
        tmpString(5) = tmpString(5) & PrevCal.CalData.X
        tmpString(6) = tmpString(6) & ""
        tmpString(7) = tmpString(7) & PrevCal.CalData.R2
    End If
    
    ' PRINT tmpString's
    ' PRINT tmpString's
    ' PRINT tmpString's
    For row = 0 To 7
        Print_Line tmpString(row)
    Next row
    
        
    Print_Line ""
    Print_Line ""
    Print_Line ""
    ' Print "Calibration points" centered
    Print_Center "Calibration Points"
    Print_Line ""
    
    ' left pad tmpString's
    For row = 0 To intNumRows
        tmpString(row) = Space(descCol)
    Next row

    ' COL DESCRIPTIONS
    tmpString(0) = tmpString(0) & "% FullScale "
    tmpString(0) = tmpString(0) & Space(dataLeft - Len(tmpString(0)))
    tmpString(0) = tmpString(0) & "Actual"
    tmpString(0) = tmpString(0) & Space((dataLeft + dataWidth) - Len(tmpString(0)))
    tmpString(0) = tmpString(0) & "Calibrated"
    tmpString(0) = tmpString(0) & Space((dataLeft + dataWidth + dataWidth) - Len(tmpString(0)))
    tmpString(0) = tmpString(0) & "% Difference"
    If bPrintPrevCal Then
        tmpString(0) = tmpString(0) & Space(dataRight - Len(tmpString(0)))
        tmpString(0) = tmpString(0) & "Actual"
        tmpString(0) = tmpString(0) & Space((dataRight + dataWidth) - Len(tmpString(0)))
        tmpString(0) = tmpString(0) & "Calibrated"
        tmpString(0) = tmpString(0) & Space((dataRight + dataWidth + dataWidth) - Len(tmpString(0)))
        tmpString(0) = tmpString(0) & "% Difference"
    End If
    
    ' COL DATA
    For row = 1 To intNumRows
        ' current calibration
        If row <= intNumCurrRows Then
            tmpString(row) = tmpString(row) & CurrCal.PointData(row).RawPercent
            sngActualValue = CurrCal.PointData(row).ActualValue
            sngCalibValue = Cal_Scale((CurrCal.PointData(row).ActualPercent / 100), SelCalScl, CurrCal)
            If (sngActualValue > CSng(0)) Then
                sngPercDiff = ((sngCalibValue - sngActualValue) / sngActualValue) * 100
            Else
                sngPercDiff = CSng(0)
            End If
            tmpString(row) = tmpString(row) & Space(dataLeft - Len(tmpString(row)))
            tmpString(row) = tmpString(row) & Format(sngActualValue, "#,###,##0.0#")
            tmpString(row) = tmpString(row) & Space((dataLeft + dataWidth) - Len(tmpString(row)))
            tmpString(row) = tmpString(row) & Format(sngCalibValue, "#,###,##0.0#")
            tmpString(row) = tmpString(row) & Space((dataLeft + dataWidth + dataWidth) - Len(tmpString(row)))
            tmpString(row) = tmpString(row) & Format(sngPercDiff, "0000.0")
        End If
        ' previous calibration
        If bPrintPrevCal Then
            If row <= intNumPrevRows Then
                sngActualValue = PrevCal.PointData(row).ActualValue
                sngCalibValue = Cal_Scale((PrevCal.PointData(row).ActualPercent / 100), SelCalScl, PrevCal)
                If (sngActualValue > CSng(0)) Then
                    sngPercDiff = ((sngCalibValue - sngActualValue) / sngActualValue) * 100
                Else
                    sngPercDiff = CSng(0)
                End If
                tmpString(row) = tmpString(row) & Space(dataRight - Len(tmpString(row)))
                tmpString(row) = tmpString(row) & Format(sngActualValue, "#,###,##0.0#")
                tmpString(row) = tmpString(row) & Space((dataRight + dataWidth) - Len(tmpString(row)))
                tmpString(row) = tmpString(row) & Format(sngCalibValue, "#,###,##0.0#")
                tmpString(row) = tmpString(row) & Space((dataRight + dataWidth + dataWidth) - Len(tmpString(row)))
                tmpString(row) = tmpString(row) & Format(sngPercDiff, "0000.0")
            End If
        End If
    Next row
    
    
    ' PRINT tmpString's
    ' PRINT tmpString's
    ' PRINT tmpString's
    For row = 0 To intNumRows
        Print_Line tmpString(row)
    Next row
      
    Print_Footer
    Printer.EndDoc
    Printer.Font = oldFont
    
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

Private Sub DisplayCurrCalInformation()
' Displays current calibration information
    NumCalPoints = Curr_SclCal.NumPoints
    SclMaxGrams = Curr_SclCal.CalRangeMax
    SclMinGrams = Curr_SclCal.CalRangeMin
    txtNumCalPts.text = Format(NumCalPoints, "#0")
    txtNumCalPts.ForeColor = Black
    txtNumCalPts.BackColor = frmCalInformation.BackColor
    txtSclEuMax.text = Format(SclMaxGrams, "###,##0")
    txtSclEuMax.ForeColor = Black
    txtSclEuMax.BackColor = frmCalInformation.BackColor
    txtSclEuMin.text = Format(SclMinGrams, "###,##0")
    txtSclEuMin.ForeColor = Black
    txtSclEuMin.BackColor = frmCalInformation.BackColor
    txtCalibDts.text = Format(Curr_SclCal.Dts, "YYYY-MMM-DD hh:mm:ss")
    txtStandardTemp.text = Format(Curr_SclCal.StandardTempValue, "###0.0")
    TempUnits.ListIndex = TempUnitsIndex(Curr_SclCal.StandardTempUnits)
    txtStandardPress.text = Format(Curr_SclCal.StandardPressValue, "###0.0##")
    PressUnits.ListIndex = PressUnitsIndex(Curr_SclCal.StandardPressUnits)
    txtCalibBy.text = Curr_SclCal.CalibratedBy
    txtEquipment.text = "equipment"
    txtComment.text = Curr_SclCal.Comment
    TempUnits.Refresh
    PressUnits.Refresh
End Sub

Private Sub DisplayNewCalInformation()
' Displays new calibration information
    NumCalPoints = New_SclCal.NumPoints
    SclMaxGrams = New_SclCal.CalRangeMax
    SclMinGrams = New_SclCal.CalRangeMin
    txtNumCalPts.text = Format(NumCalPoints, "#0")
    txtNumCalPts.ForeColor = Black
    txtNumCalPts.BackColor = frmCalInformation.BackColor
    txtSclEuMax.text = Format(SclMaxGrams, "###,##0")
    txtSclEuMax.ForeColor = TitlesData_Forecolor
    txtSclEuMax.BackColor = White
    txtSclEuMin.text = Format(SclMinGrams, "###,##0")
    txtSclEuMin.ForeColor = TitlesData_Forecolor
    txtSclEuMin.BackColor = White
    txtCalibDts.text = Format(New_SclCal.Dts, "YYYY-MMM-DD hh:mm:ss")
    txtStandardTemp.text = Format(New_SclCal.StandardTempValue, "###0.0")
    TempUnits.ListIndex = TempUnitsIndex(New_SclCal.StandardTempUnits)
    txtStandardPress.text = Format(New_SclCal.StandardPressValue, "###0.0##")
    PressUnits.ListIndex = PressUnitsIndex(New_SclCal.StandardPressUnits)
    txtCalibBy.text = New_SclCal.CalibratedBy
    txtEquipment.text = "equipment"
    txtComment.text = New_SclCal.Comment
    TempUnits.Refresh
    PressUnits.Refresh
End Sub

Private Sub DisplayCalPointData()
' Displays new, current & previous calibration point data
    Dim iPoint As Integer
    HideInactiveTableRows
    For iPoint = 1 To NumCalPoints
        lblPointNum(iPoint).Caption = Format(iPoint, "#0")
        txtNewRawValue(iPoint).text = IIf(txtNewRawValue(iPoint).Enabled, Format(New_SclCal.PointData(iPoint).RawValue, "####0.0##"), "")
        lblNewRawPerc(iPoint).Caption = IIf(lblNewRawPerc(iPoint).Enabled, Format(New_SclCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        txtNewActualValue(iPoint).text = IIf(txtNewActualValue(iPoint).Enabled, Format(New_SclCal.PointData(iPoint).ActualValue, "####0.0##"), "")
        lblCurrRawPerc(iPoint).Caption = IIf(lblCurrRawPerc(iPoint).Enabled, Format(Curr_SclCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        lblCurrActualValue(iPoint).Caption = IIf(lblCurrActualValue(iPoint).Enabled, Format(Curr_SclCal.PointData(iPoint).ActualValue, "####0.0##"), "")
        lblPrevRawPerc(iPoint).Caption = IIf(lblPrevRawPerc(iPoint).Enabled, Format(Prev_SclCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        lblPrevActualValue(iPoint).Caption = IIf(lblPrevActualValue(iPoint).Enabled, Format(Prev_SclCal.PointData(iPoint).ActualValue, "####0.0##"), "")
    Next iPoint
    DoEvents
    DisplayCurrCalValues
    DoEvents
    DisplayPrevCalValues
    DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Xit()
    ' make sure xlWB is closed
'    Set xlWB = Nothing
    ' Close Excel
'    Set xlApp = Nothing
    ' close screen
    Unload Me
    Set frmScalesCal = Nothing
End Sub

Private Sub txtNewActualValue_Change(Index As Integer)
        txtNewActualValue(Index).ForeColor = TitlesData_Forecolor
        txtNewActualValue(Index).BackColor = Entry_BackColor
End Sub

Private Sub txtNewRawValue_Change(Index As Integer)
        txtNewRawValue(Index).ForeColor = TitlesData_Forecolor
        txtNewRawValue(Index).BackColor = Entry_BackColor
End Sub

Private Sub txtNumCalPts_Change()
'
    If (Not bEditNumCalPts) Then
        txtNumCalPts.text = Format(NumCalPoints, "#0")
    End If
End Sub

Private Function RangeCheck(hiVal As Single, loVal As Single, ByRef tbox As TextBox, sMessage As String) As Boolean
'
' Module Name:  RangeCheck
' Description:  This routine checks the value of the text box supplied and
'               compares to the high and low limits.  If the values is
'               within or equal to the limits, the routine returns a value
'               of TRUE, otherwise a FALSE value.
'
'               When the return value is false, a message supplying the
'               allowable limits is displayed, and the contents of the
'               text box are highlighted.
'
SetErrModule 8118, 7
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' convert empty or non-numeric entry to "0"
    If (tbox.text = "") Then tbox.text = "0"
    If (Not IsNumeric(tbox.text)) Then tbox.text = "0"
    ' check entry for valid range
    If CStr(tbox) < loVal Or CStr(tbox) > hiVal Then
        ' entry is outside limits
        RangeCheck = False
        tbox.BackColor = EntryInvalid_BackColor
        tbox.ForeColor = Alarm_ForeColor
        txtMsg.ForeColor = Alarm_ForeColor
        txtMsg.text = sMessage & " Range Error!" & vbCrLf _
                    & "Allowable Range = " & loVal & " - " & hiVal
        MyFocus tbox
    Else
        ' entry is valid
        RangeCheck = True
    End If

ResetErrModule
Exit Function

localhandler:
RangeCheck = False
tbox.BackColor = EntryInvalid_BackColor
tbox.ForeColor = Alarm_ForeColor
txtMsg.ForeColor = Alarm_ForeColor
txtMsg.text = sMessage & " RangeCheck Error!" & vbCrLf _
            & "Allowable Range = " & loVal & " - " & hiVal
MyFocus tbox
End Function

Private Function PressUnitsIndex(ByVal unitsText As String) As Integer
' returns ListIndex of PressUnits that matches unitsText
' defaults to 0
Dim idx As Integer
Dim idxList As Integer

    idxList = 0
    For idx = 0 To (PressUnits.ListCount - 1)
        If (Trim(unitsText) = Trim(PressUnits.List(idx))) Then idxList = idx
    Next idx
    PressUnitsIndex = idxList
End Function

Private Function TempUnitsIndex(ByVal unitsText As String) As Integer
' returns ListIndex of TempUnits that matches unitsText
' defaults to 0
Dim idx As Integer
Dim idxList As Integer

    idxList = 0
    For idx = 0 To (TempUnits.ListCount - 1)
        If (Trim(unitsText) = Trim(TempUnits.List(idx))) Then idxList = idx
    Next idx
    TempUnitsIndex = idxList
End Function



