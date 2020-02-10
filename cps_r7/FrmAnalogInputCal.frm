VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form frmAnalogInputCal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AnalogInput Calibration Screen"
   ClientHeight    =   11160
   ClientLeft      =   195
   ClientTop       =   720
   ClientWidth     =   15345
   Icon            =   "FrmAnalogInputCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   15345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmGroupSelection 
      Caption         =   "Analog Group Selection"
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
      TabIndex        =   12
      Top             =   2400
      Width           =   7065
      Begin VB.CommandButton cmdGroupUp 
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
         Picture         =   "FrmAnalogInputCal.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "next group of analogs"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.CommandButton cmdGroupDn 
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
         Picture         =   "FrmAnalogInputCal.frx":5EE4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "previous group of analogs"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.TextBox txtDispGrp 
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
         TabIndex        =   13
         Text            =   "Station 8"
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
      TabIndex        =   27
      Top             =   120
      Width           =   7065
      Begin VB.CommandButton cmdCalPoints 
         Caption         =   "Edit CalPoints"
         DisabledPicture =   "FrmAnalogInputCal.frx":65E6
         DownPicture     =   "FrmAnalogInputCal.frx":6928
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
         Picture         =   "FrmAnalogInputCal.frx":6C6A
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdPrintCal 
         Caption         =   "Print"
         DisabledPicture =   "FrmAnalogInputCal.frx":6FAC
         DownPicture     =   "FrmAnalogInputCal.frx":76AE
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
         Picture         =   "FrmAnalogInputCal.frx":7DB0
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Print a Report of the Selected Calibration"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdRestorePrevCal 
         Caption         =   "Restore Previous"
         DisabledPicture =   "FrmAnalogInputCal.frx":84B2
         DownPicture     =   "FrmAnalogInputCal.frx":87F4
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
         Picture         =   "FrmAnalogInputCal.frx":8B36
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdSaveCurrCal 
         Caption         =   "     Save      Current"
         DisabledPicture =   "FrmAnalogInputCal.frx":8E78
         DownPicture     =   "FrmAnalogInputCal.frx":91BA
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
         Picture         =   "FrmAnalogInputCal.frx":94FC
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdCreateNewCalib 
         Caption         =   " Create New"
         DisabledPicture =   "FrmAnalogInputCal.frx":983E
         DownPicture     =   "FrmAnalogInputCal.frx":9B80
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
         Picture         =   "FrmAnalogInputCal.frx":9EC2
         Style           =   1  'Graphical
         TabIndex        =   30
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
         TabIndex        =   29
         Text            =   "FrmAnalogInputCal.frx":A204
         Top             =   1080
         Width           =   6795
      End
      Begin VB.CommandButton cmdRunCal 
         Caption         =   "Calibration"
         DisabledPicture =   "FrmAnalogInputCal.frx":A22F
         DownPicture     =   "FrmAnalogInputCal.frx":A571
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
         Picture         =   "FrmAnalogInputCal.frx":A8B3
         Style           =   1  'Graphical
         TabIndex        =   28
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
      TabIndex        =   20
      Top             =   6960
      Width           =   4725
      Begin VB.CommandButton cmdHelp 
         Height          =   615
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmAnalogInputCal.frx":ABF5
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   3360
         UseMaskColor    =   -1  'True
         Width           =   615
      End
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
         ItemData        =   "FrmAnalogInputCal.frx":B837
         Left            =   3120
         List            =   "FrmAnalogInputCal.frx":B84A
         Style           =   2  'Dropdown List
         TabIndex        =   194
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
         TabIndex        =   192
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
         ItemData        =   "FrmAnalogInputCal.frx":B86A
         Left            =   3120
         List            =   "FrmAnalogInputCal.frx":B87A
         Style           =   2  'Dropdown List
         TabIndex        =   191
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
         TabIndex        =   189
         Text            =   "20"
         ToolTipText     =   "Standard Temperature Value"
         Top             =   1320
         Width           =   735
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
         TabIndex        =   188
         Text            =   "FrmAnalogInputCal.frx":B89A
         ToolTipText     =   "Maximum length: 32 characters"
         Top             =   2460
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
         TabIndex        =   187
         Text            =   "YYYY-MMM-DD hh:mm:ss"
         Top             =   960
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
         TabIndex        =   186
         Text            =   "FrmAnalogInputCal.frx":B8A4
         Top             =   2790
         Width           =   2775
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
         TabIndex        =   185
         Text            =   "FrmAnalogInputCal.frx":B8AE
         ToolTipText     =   "Maximum length: 32 characters"
         Top             =   2130
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
         Picture         =   "FrmAnalogInputCal.frx":B8B9
         Style           =   1  'Graphical
         TabIndex        =   169
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
         TabIndex        =   21
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
         TabIndex        =   193
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
         TabIndex        =   190
         Top             =   1350
         Width           =   1965
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   360
         Width           =   2535
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
         TabIndex        =   24
         Top             =   2820
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
         TabIndex        =   23
         Top             =   1005
         Width           =   1365
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
         TabIndex        =   22
         Top             =   2490
         Width           =   1365
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
      TabIndex        =   16
      Top             =   6960
      Width           =   10335
      Begin VB.CommandButton cmdAcquireActual 
         DisabledPicture =   "FrmAnalogInputCal.frx":BBFB
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
         Left            =   2115
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmAnalogInputCal.frx":BF3D
         Style           =   1  'Graphical
         TabIndex        =   198
         ToolTipText     =   "Copy Current I/O Value to Selected Actual Entry"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdAcquireRaw 
         DisabledPicture =   "FrmAnalogInputCal.frx":C27F
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
         Left            =   675
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmAnalogInputCal.frx":C5C1
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Copy Current I/O Value to Selected Raw Entry"
         Top             =   240
         UseMaskColor    =   -1  'True
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
         Index           =   10
         Left            =   675
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   184
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
         TabIndex        =   183
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
         TabIndex        =   182
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
         TabIndex        =   181
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
         TabIndex        =   180
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
         TabIndex        =   179
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
         TabIndex        =   178
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
         TabIndex        =   177
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
         TabIndex        =   176
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
         TabIndex        =   175
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
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   161
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
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
         TabIndex        =   154
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   141
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   1125
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
         Index           =   1
         Left            =   3150
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   285
         Width           =   2340
      End
   End
   Begin VB.Frame frmInputSelection 
      Caption         =   "Analog Input Selection"
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
      TabIndex        =   2
      Top             =   3480
      Width           =   7065
      Begin VB.ComboBox CalEntries 
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
         Height          =   360
         ItemData        =   "FrmAnalogInputCal.frx":C903
         Left            =   3270
         List            =   "FrmAnalogInputCal.frx":C910
         Style           =   2  'Dropdown List
         TabIndex        =   200
         ToolTipText     =   "Which Column(s) of Data to be entered for a New Calibration"
         Top             =   1260
         Width           =   1335
      End
      Begin VB.ComboBox RawValues 
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
         Height          =   360
         ItemData        =   "FrmAnalogInputCal.frx":C939
         Left            =   3270
         List            =   "FrmAnalogInputCal.frx":C949
         Style           =   2  'Dropdown List
         TabIndex        =   195
         ToolTipText     =   "Units for Standard Pressure"
         Top             =   900
         Width           =   1335
      End
      Begin VB.CommandButton cmdInputDn 
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
         Picture         =   "FrmAnalogInputCal.frx":C967
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "previous input"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.CommandButton cmdInputUp 
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
         Picture         =   "FrmAnalogInputCal.frx":D069
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "next input"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.TextBox txtaEUMin 
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
         Height          =   285
         Left            =   5805
         MaxLength       =   6
         TabIndex        =   170
         Text            =   "01234"
         ToolTipText     =   "Min Value in Engineering Units"
         Top             =   540
         Width           =   1005
      End
      Begin VB.TextBox txtaFuncDesc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   48
         TabIndex        =   6
         Text            =   "Function Description123456789012345678901234567890"
         ToolTipText     =   "Function Description"
         Top             =   540
         Width           =   4485
      End
      Begin VB.TextBox txtaVDCMin 
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
         Height          =   285
         Left            =   5805
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Min Value in Volts"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtaVDCMax 
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
         Height          =   285
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "12345678"
         ToolTipText     =   "Max Value in Volts"
         Top             =   1260
         Width           =   1005
      End
      Begin VB.TextBox txtaEUMax 
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
         Height          =   285
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "12345"
         ToolTipText     =   "Max Value in Engineering Units"
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblCalMethod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cal Entries "
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
         Left            =   1800
         TabIndex        =   201
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label lblRawVal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Raw Values as "
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
         Left            =   1800
         TabIndex        =   196
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lblaFuncDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   3885
      End
      Begin VB.Label lblaVdcMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vdc Min"
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
         Left            =   5805
         TabIndex        =   10
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label lblaVdcMax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Vdc Max"
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
         Left            =   4800
         TabIndex        =   9
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label lblaEUMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EU Min"
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
         Left            =   5805
         TabIndex        =   8
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblaEUMax 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EU Max"
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
         Left            =   4800
         TabIndex        =   7
         Top             =   300
         Width           =   1005
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
      TabIndex        =   166
      Top             =   120
      Width           =   8055
      Begin MSChart20Lib.MSChart chtAIChart 
         Height          =   6435
         Left            =   60
         OleObjectBlob   =   "FrmAnalogInputCal.frx":D76B
         TabIndex        =   167
         Top             =   180
         Width           =   7920
      End
   End
End
Attribute VB_Name = "frmAnalogInputCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''ERROR module 1181
''
''frmAnalogInputCal
''
Option Explicit

Private blnPrevCalExists As Boolean                       ' Flag - Whether previous calibration data is loaded
' Analog Input calibration variables
Private SelCalGrp As Integer                              ' Group of Analogs; 0=Common, 1-9=Station#, 10=FID, 11-19=Purge#
Private SelCalInp As Integer                              ' The Index of the Analog Input currently selected
Private SelCalPoint As Integer                            ' The Index of the currently selected Calibration Point(row)
Private NumCalPoints As Integer                           ' The Number of Data Points for the currently selected Analog Input
Private AiOptoType As Integer                             ' The Type of Opto22 Hardware Module for the currently selected Analog Input
Private sRawMax As Single
Private sEuMax As Single
Private sRawMin As Single
Private sEuMin As Single
Private sRawSpan As Single
Private sEuSpan As Single
Private Curr_AiCal As AICalibration
Private New_AiCal As AICalibration
Private Prev_AiCal As AICalibration
' Max Group Index
Const MAXGRP = 19
' Min Group Index
Const MINGRP = 0
' Max Inputs per Group
Const MAXINP = 29
' AI Selecton Update Options
Const useExistRawInputType = 0
Const useNewRawInputType = 1

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
Private bFormLoaded As Boolean           ' Flag - Whether the form has loaded(can't SetFocus on a TextBox until form is Loaded

Private Sub FillNewCalPointTable()
    ' Define every AI calibration point
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 1181, 1
Dim iPoint As Integer
Dim idx As Integer
    
    Select Case NumCalPoints
        Case 3
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(50)
            New_AiCal.PointData(3).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(50)
            New_AiCal.PointData(3).ActualPercent = CSng(100)
        Case 4
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(30)
            New_AiCal.PointData(3).RawPercent = CSng(70)
            New_AiCal.PointData(4).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(30)
            New_AiCal.PointData(3).ActualPercent = CSng(70)
            New_AiCal.PointData(4).ActualPercent = CSng(100)
        Case 5
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(25)
            New_AiCal.PointData(3).RawPercent = CSng(50)
            New_AiCal.PointData(4).RawPercent = CSng(75)
            New_AiCal.PointData(5).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(25)
            New_AiCal.PointData(3).ActualPercent = CSng(50)
            New_AiCal.PointData(4).ActualPercent = CSng(75)
            New_AiCal.PointData(5).ActualPercent = CSng(100)
        Case 6
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(20)
            New_AiCal.PointData(3).RawPercent = CSng(40)
            New_AiCal.PointData(4).RawPercent = CSng(60)
            New_AiCal.PointData(5).RawPercent = CSng(80)
            New_AiCal.PointData(6).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(20)
            New_AiCal.PointData(3).ActualPercent = CSng(40)
            New_AiCal.PointData(4).ActualPercent = CSng(60)
            New_AiCal.PointData(5).ActualPercent = CSng(80)
            New_AiCal.PointData(6).ActualPercent = CSng(100)
        Case 7
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(16.7)
            New_AiCal.PointData(3).RawPercent = CSng(33.3)
            New_AiCal.PointData(4).RawPercent = CSng(50)
            New_AiCal.PointData(5).RawPercent = CSng(66.7)
            New_AiCal.PointData(6).RawPercent = CSng(83.3)
            New_AiCal.PointData(7).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(16.7)
            New_AiCal.PointData(3).ActualPercent = CSng(33.3)
            New_AiCal.PointData(4).ActualPercent = CSng(50)
            New_AiCal.PointData(5).ActualPercent = CSng(66.7)
            New_AiCal.PointData(6).ActualPercent = CSng(83.3)
            New_AiCal.PointData(7).ActualPercent = CSng(100)
        Case 8
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(14.3)
            New_AiCal.PointData(3).RawPercent = CSng(28.6)
            New_AiCal.PointData(4).RawPercent = CSng(42.9)
            New_AiCal.PointData(5).RawPercent = CSng(57.1)
            New_AiCal.PointData(6).RawPercent = CSng(71.4)
            New_AiCal.PointData(7).RawPercent = CSng(85.7)
            New_AiCal.PointData(8).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(14.3)
            New_AiCal.PointData(3).ActualPercent = CSng(28.6)
            New_AiCal.PointData(4).ActualPercent = CSng(42.9)
            New_AiCal.PointData(5).ActualPercent = CSng(57.1)
            New_AiCal.PointData(6).ActualPercent = CSng(71.4)
            New_AiCal.PointData(7).ActualPercent = CSng(85.7)
            New_AiCal.PointData(8).ActualPercent = CSng(100)
        Case 9
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(12.5)
            New_AiCal.PointData(3).RawPercent = CSng(25)
            New_AiCal.PointData(4).RawPercent = CSng(37.5)
            New_AiCal.PointData(5).RawPercent = CSng(50)
            New_AiCal.PointData(6).RawPercent = CSng(62.5)
            New_AiCal.PointData(7).RawPercent = CSng(75)
            New_AiCal.PointData(8).RawPercent = CSng(87.5)
            New_AiCal.PointData(9).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(12.5)
            New_AiCal.PointData(3).ActualPercent = CSng(25)
            New_AiCal.PointData(4).ActualPercent = CSng(37.5)
            New_AiCal.PointData(5).ActualPercent = CSng(50)
            New_AiCal.PointData(6).ActualPercent = CSng(62.5)
            New_AiCal.PointData(7).ActualPercent = CSng(75)
            New_AiCal.PointData(8).ActualPercent = CSng(87.5)
            New_AiCal.PointData(9).ActualPercent = CSng(100)
        Case 10
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(11.1)
            New_AiCal.PointData(3).RawPercent = CSng(22.2)
            New_AiCal.PointData(4).RawPercent = CSng(33.3)
            New_AiCal.PointData(5).RawPercent = CSng(44.4)
            New_AiCal.PointData(6).RawPercent = CSng(55.5)
            New_AiCal.PointData(7).RawPercent = CSng(66.6)
            New_AiCal.PointData(8).RawPercent = CSng(77.7)
            New_AiCal.PointData(9).RawPercent = CSng(88.8)
            New_AiCal.PointData(10).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(11.1)
            New_AiCal.PointData(3).ActualPercent = CSng(22.2)
            New_AiCal.PointData(4).ActualPercent = CSng(33.3)
            New_AiCal.PointData(5).ActualPercent = CSng(44.4)
            New_AiCal.PointData(6).ActualPercent = CSng(55.5)
            New_AiCal.PointData(7).ActualPercent = CSng(66.6)
            New_AiCal.PointData(8).ActualPercent = CSng(77.7)
            New_AiCal.PointData(9).ActualPercent = CSng(88.8)
            New_AiCal.PointData(10).ActualPercent = CSng(100)
        Case 11
            New_AiCal.PointData(1).RawPercent = CSng(0)
            New_AiCal.PointData(2).RawPercent = CSng(10)
            New_AiCal.PointData(3).RawPercent = CSng(20)
            New_AiCal.PointData(4).RawPercent = CSng(30)
            New_AiCal.PointData(5).RawPercent = CSng(40)
            New_AiCal.PointData(6).RawPercent = CSng(50)
            New_AiCal.PointData(7).RawPercent = CSng(60)
            New_AiCal.PointData(8).RawPercent = CSng(70)
            New_AiCal.PointData(9).RawPercent = CSng(80)
            New_AiCal.PointData(10).RawPercent = CSng(90)
            New_AiCal.PointData(11).RawPercent = CSng(100)
            New_AiCal.PointData(1).ActualPercent = CSng(0)
            New_AiCal.PointData(2).ActualPercent = CSng(10)
            New_AiCal.PointData(3).ActualPercent = CSng(20)
            New_AiCal.PointData(4).ActualPercent = CSng(30)
            New_AiCal.PointData(5).ActualPercent = CSng(40)
            New_AiCal.PointData(6).ActualPercent = CSng(50)
            New_AiCal.PointData(7).ActualPercent = CSng(60)
            New_AiCal.PointData(8).ActualPercent = CSng(70)
            New_AiCal.PointData(9).ActualPercent = CSng(80)
            New_AiCal.PointData(10).ActualPercent = CSng(90)
            New_AiCal.PointData(11).ActualPercent = CSng(100)
    End Select
    
    ' get min/max EU & Raw  for appropriate input
    ' calc EU & Raw spans
    CalcSpans New_AiCal.RawInputType
    
    For iPoint = 1 To MAXLSQCALPOINTS
        New_AiCal.PointData(iPoint).RawValue = sRawMin + (sRawSpan * (New_AiCal.PointData(iPoint).RawPercent / CSng(100)))
        New_AiCal.PointData(iPoint).ActualValue = sEuMin + (sEuSpan * (New_AiCal.PointData(iPoint).ActualPercent / CSng(100)))
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

Private Sub ClearAiCal(ByRef tmpCal As AICalibration)
Dim iPoint As Integer

    ' Set Analog Input Calibration to Linear (default)
    ' set calibration parameters
    tmpCal.Dts = Now
    tmpCal.CalibratedBy = "na"
    tmpCal.Comment = "cleared"
    tmpCal.NumPoints = MINLSQCALPOINTS
    tmpCal.RawInputType = CalRawAsVolts
    '            tmpCal.CalData.X = sEuMax
    tmpCal.CalData.X = CSng(0)
    tmpCal.CalData.X2 = CSng(0)
    tmpCal.CalData.X3 = CSng(0)
    tmpCal.CalData.X4 = CSng(0)
    tmpCal.CalData.X5 = CSng(0)
    tmpCal.CalData.X6 = CSng(0)
    tmpCal.CalData.R2 = CSng(0)
    
    ' set Analog Input Calibration Point Data
    For iPoint = 1 To MAXLSQCALPOINTS
        tmpCal.PointData(iPoint).ActualPercent = CSng(0)
        tmpCal.PointData(iPoint).RawPercent = CSng(0)
        tmpCal.PointData(iPoint).ActualValue = CSng(0)
        tmpCal.PointData(iPoint).RawValue = CSng(0)
    Next iPoint
            
End Sub

Private Sub cmdAcquireActual_Click()
    If bNewCalEnabled Then
        If ((Com_AIO(acCustCalDevice).addr <> 0) Or (Com_AIO(acCustCalDevice).chan <> 0)) Then
            txtNewActualValue(SelCalPoint).text = Format(Com_AIO(acCustCalDevice).EUValue, "####0.0##")
        End If
    End If
End Sub

Private Sub cmdAcquireRaw_Click()
Dim Sel_AIO As FuncAnalogIO                                        ' Analog IO Information
Dim idx As Integer
Dim address As Integer
Dim channel As Integer
Dim sRawMlt As Single
Dim sRawVal As Single
Dim sRawVal2 As Single
    
    If bNewCalEnabled Then
        Select Case SelCalGrp
            Case calgrpComm
                Sel_AIO = Com_AIO(SelCalInp)
            Case calgrpStn1 To calgrpStn9
                idx = SelCalGrp
                Sel_AIO = Stn_AIO(idx, SelCalInp)
            Case calgrpFid
'                Sel_AIO = Fid_AIO(SelCalInp)
            Case calgrpPrg1 To calgrpPrg9
                idx = SelCalGrp - 10
                Sel_AIO = Prg_AIO(idx, SelCalInp)
            Case Else
                Exit Sub
        End Select
        address = Sel_AIO.addr
        channel = Sel_AIO.chan
        Select Case New_AiCal.RawInputType
            Case CalRawAsVolts  ' 0-5 volts = 0-327670 counts
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawMlt = sRawVal / (FULLSCALE / CSng(2))
                sRawVal2 = sRawMlt * CSng(5)
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
            Case CalRawAsMa     ' 0-20ma (converted from Vdc) = 0-327670 counts
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawMlt = sRawVal / (FULLSCALE / CSng(2))
                sRawVal2 = sRawMlt * CSng(20)
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
            Case CalRawAsDegC   ' Opto TC & RTD modules return temp as ***.0 degC
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawVal2 = sRawVal * 0.1
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
            Case CalRawAsEU     ' Raw range = EU range = 0-327670 counts
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawMlt = (sRawVal / (FULLSCALE / CSng(2)))
                sRawVal2 = sEuMin + (sRawMlt * sEuSpan)
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
        End Select
    End If
End Sub

Private Sub cmdApply_Click()
Dim iRawInputType As Integer
Dim idx As Integer
Dim iPoint As Integer
Dim tmpVal As Single
Dim tmpval2 As Single
Dim flag As Boolean

    ' valid value in NumCalPts text box ??
    flag = RangeCheck(MAXLSQCALPOINTS, MINLSQCALPOINTS, txtNumCalPts, "Number of Calibration Points")
    If flag Then
        
        ' setup screen for different NumberOfCalibrationPoints
        If (CInt(ValueFromText(txtNumCalPts.text)) <> NumCalPoints) Then
            ' Number of Calibration Points for this Analog Input has been changed
            NumCalPoints = CInt(ValueFromText(txtNumCalPts.text))
            iRawInputType = Curr_AiCal.RawInputType
            ClearAiCal New_AiCal
            ClearAiCal Curr_AiCal
            ClearAiCal Prev_AiCal
            EnableNewCal False
            EnablePrevCal False
            ' Set Analog Input Calibration to Linear (default)
            ' get min/max EU & Raw  for appropriate input
            ' calc EU & Vdc spans
            CalcSpans iRawInputType
            ' set calibration parameters
            Curr_AiCal.Dts = Now()
            Curr_AiCal.CalibratedBy = "default"
            Curr_AiCal.Equipment = "equipment"
            Curr_AiCal.Comment = "linear"
            Curr_AiCal.NumPoints = NumCalPoints
            Curr_AiCal.RawInputType = iRawInputType
    '            New_AiCal.CalData.X = sEuMax
            Curr_AiCal.CalData.X = CSng(1)
            Curr_AiCal.CalData.X2 = CSng(0)
            Curr_AiCal.CalData.X3 = CSng(0)
            Curr_AiCal.CalData.X4 = CSng(0)
            Curr_AiCal.CalData.X5 = CSng(0)
            Curr_AiCal.CalData.X6 = CSng(0)
            Curr_AiCal.CalData.R2 = CSng(0)
                
            ' set Analog Input Calibration Point Data
            For iPoint = 1 To MAXLSQCALPOINTS
                idx = iPoint - 1
                tmpVal = CSng(idx) * CSng(10)
                tmpval2 = tmpVal / CSng(100)
                Curr_AiCal.PointData(iPoint).ActualPercent = tmpVal
                Curr_AiCal.PointData(iPoint).RawPercent = tmpVal
                Curr_AiCal.PointData(iPoint).ActualValue = sEuMin + (tmpval2 * sEuSpan)
                Curr_AiCal.PointData(iPoint).RawValue = sRawMin + (tmpval2 * sRawSpan)
            Next iPoint
            New_AiCal = Curr_AiCal
        End If
        bEditNumCalPts = False
        DisplayAiAll
        
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
    ClearAiCal New_AiCal
    New_AiCal.NumPoints = NumCalPoints
    New_AiCal.Dts = Now()
    New_AiCal.StandardTempValue = 20
    New_AiCal.StandardTempUnits = "deg C"
    New_AiCal.StandardPressValue = 1
    New_AiCal.StandardPressUnits = "atm"
    New_AiCal.CalibratedBy = "default"
    New_AiCal.Equipment = "equipment"
    New_AiCal.Comment = "linear"
    New_AiCal.RawInputType = Curr_AiCal.RawInputType
    EnableNewCal True
    FillNewCalPointTable
    DisplayNewCalInformation
    DisplayCalPointData
    SelectCalPoint 1
    UpdateCmdButtons
End Sub

Private Sub cmdGroupDn_Click()
' This command decrements the analog group number,
' the displayed name of the group, and triggers an update for
' the values displayed on the form for the current input
Dim idx As Integer
Dim flag As Boolean
    flag = False
    Do While Not flag
        SelCalGrp = SelCalGrp - 1
        If SelCalGrp = calgrpFid Then SelCalGrp = SelCalGrp - 1
        If SelCalGrp < MINGRP Then SelCalGrp = MAXGRP
        Select Case SelCalGrp
            Case calgrpStn1 To calgrpStn9
                idx = SelCalGrp
                If idx <= LAST_STN Then flag = True
            Case calgrpFid
                If USINGFIDANALYZER Then flag = True
            Case calgrpPrg1 To calgrpPrg9
                idx = SelCalGrp - 10
                If (USINGPASLOCALCONTROL And (idx <= NR_PRGAIR)) Then flag = True
            Case Else
                flag = True
        End Select
    Loop
    ' check for valid input in new group
    SelCalInp = SelCalInp - 1
    NextInput
    bUnsavedCal = False
    ' Update the Display
    UpdateAiSelection
    DisplayAiAll
End Sub

Private Sub cmdGroupUp_Click()
' This command increments the analog group number,
' the displayed name of the group, and triggers an update for
' the values displayed on the form for the current input
Dim idx As Integer
Dim flag As Boolean
    flag = False
    Do While Not flag
        SelCalGrp = SelCalGrp + 1
        If SelCalGrp = calgrpFid Then SelCalGrp = SelCalGrp + 1
        If SelCalGrp > MAXGRP Then SelCalGrp = 0
        Select Case SelCalGrp
            Case calgrpStn1 To calgrpStn9
                idx = SelCalGrp
                If idx <= LAST_STN Then flag = True
            Case calgrpFid
                If USINGFIDANALYZER Then flag = True
            Case calgrpPrg1 To calgrpPrg9
                idx = SelCalGrp - 10
                If (USINGPASLOCALCONTROL And (idx <= NR_PRGAIR)) Then flag = True
            Case Else
                flag = True
        End Select
    Loop
    ' check for valid input in new group
    SelCalInp = SelCalInp - 1
    NextInput
    bUnsavedCal = False
    ' Update the Display
    UpdateAiSelection
    DisplayAiAll
End Sub

Private Sub cmdHelp_Click()
    frmCalHelp.Show
End Sub

Private Sub cmdInputDn_Click()
    PrevInput
End Sub

Private Sub cmdInputUp_Click()
    NextInput
End Sub

Private Sub PrevInput()
' This command decrements the analog input number,
' the displayed name of the group, and triggers an update for
' the values displayed on the form for the current input
Dim idx As Integer
Dim iAddr As Integer
Dim iChan As Integer
Dim iCntr As Integer
Dim iMax As Integer
Dim sDesc As String
Dim flag As Boolean

    ' get max input index for appropriate array
    Select Case SelCalGrp
        Case calgrpComm
            ' Common Analog Input Calibration Parameters
            iMax = MAX_ANA_COM
        Case calgrpStn1 To calgrpStn9
            ' Station Analog Input Calibration Parameters
            iMax = MAX_ANA_STN
        Case calgrpFid
            ' FID Analog Input Calibration Parameters
            iMax = MAX_ANA_FID
        Case calgrpPrg1 To calgrpPrg9
            ' Purge Analog Input Calibration Parameters
            iMax = MAX_ANA_PRG
    End Select
            
    iCntr = 0
    flag = False
    Do While Not flag
        SelCalInp = SelCalInp - 1
        If SelCalInp < 0 Then SelCalInp = iMax
        Select Case SelCalGrp
            Case calgrpComm
                iAddr = Com_AIO(SelCalInp).addr
                iChan = Com_AIO(SelCalInp).chan
                sDesc = Com_AnaDef(SelCalInp).desc
                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case calgrpStn1 To calgrpStn9
                idx = SelCalGrp
                iAddr = Stn_AIO(idx, SelCalInp).addr
                iChan = Stn_AIO(idx, SelCalInp).chan
                sDesc = Stn_AnaDef(SelCalInp).desc
                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case calgrpFid
'                iAddr = Fid_AIO(SelCalInp).addr
'                iChan = Fid_AIO(SelCalInp).chan
'                sDesc = Fid_AnaDef(SelCalInp).desc
'                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case calgrpPrg1 To calgrpPrg9
                idx = SelCalGrp - 10
                iAddr = Prg_AIO(idx, SelCalInp).addr
                iChan = Prg_AIO(idx, SelCalInp).chan
                sDesc = Prg_AnaDef(SelCalInp).desc
                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case Else
                iAddr = 0
                iChan = 0
                sDesc = "undefined"
                AiOptoType = 0
        End Select
        ' Addr=0 AND Chan=0 is an undefined input
        ' MFC inputs and outputs are not done here
        If ((iAddr <> 0) Or (iChan <> 0)) Then
            If (InStr(1, sDesc, "MFC") = 0) Then
                ' analog input is defined and Not related to a MFC
                flag = True
            End If
        End If
        iCntr = iCntr + 1
        If iCntr > iMax Then
            flag = True
        End If
    Loop
    bUnsavedCal = False
    ' Update the Display
    UpdateAiSelection
    DisplayAiAll
End Sub

Private Sub NextInput()
' This command increments the analog input number,
' the displayed name of the group, and triggers an update for
' the values displayed on the form for the current input
Dim idx As Integer
Dim iAddr As Integer
Dim iChan As Integer
Dim iCntr As Integer
Dim iMax As Integer
Dim sDesc As String
Dim flag As Boolean

    ' get max input index for appropriate array
    Select Case SelCalGrp
        Case calgrpComm
            ' Common Analog Input Calibration Parameters
            iMax = MAX_ANA_COM
        Case calgrpStn1 To calgrpStn9
            ' Station Analog Input Calibration Parameters
            iMax = MAX_ANA_STN
        Case calgrpFid
            ' FID Analog Input Calibration Parameters
            iMax = MAX_ANA_FID
        Case calgrpPrg1 To calgrpPrg9
            ' Purge Analog Input Calibration Parameters
            iMax = MAX_ANA_PRG
    End Select
            
    iCntr = 0
    flag = False
    Do While Not flag
        SelCalInp = SelCalInp + 1
        If SelCalInp > iMax Then SelCalInp = 0
        Select Case SelCalGrp
            Case calgrpComm
                iAddr = Com_AIO(SelCalInp).addr
                iChan = Com_AIO(SelCalInp).chan
                sDesc = Com_AnaDef(SelCalInp).desc
                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case calgrpStn1 To calgrpStn9
                idx = SelCalGrp
                iAddr = Stn_AIO(idx, SelCalInp).addr
                iChan = Stn_AIO(idx, SelCalInp).chan
                sDesc = Stn_AnaDef(SelCalInp).desc
                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case calgrpFid
'                iAddr = Fid_AIO(SelCalInp).addr
'                iChan = Fid_AIO(SelCalInp).chan
'                sDesc = Fid_AnaDef(SelCalInp).desc
'                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case calgrpPrg1 To calgrpPrg9
                idx = SelCalGrp - 10
                iAddr = Prg_AIO(idx, SelCalInp).addr
                iChan = Prg_AIO(idx, SelCalInp).chan
                sDesc = Prg_AnaDef(SelCalInp).desc
                AiOptoType = OptoAIO(iAddr, iChan).Type
            Case Else
                iAddr = 0
                iChan = 0
                sDesc = "undefined"
                AiOptoType = 0
        End Select
        ' Addr=0 AND Chan=0 is an undefined input
        ' MFC inputs and outputs are not done here
        If ((iAddr <> 0) Or (iChan <> 0)) Then
            If (InStr(1, sDesc, "MFC") = 0) Then
                ' analog input is defined and Not related to an MFC
                flag = True
            End If
        End If
        iCntr = iCntr + 1
        If iCntr > iMax Then
            flag = True
        End If
    Loop
    bUnsavedCal = False
    ' Update the Display
    UpdateAiSelection
    DisplayAiAll
End Sub

Private Sub cmdRunCal_Click()
   
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 1181, 3
Dim row As Integer
Dim iPoint As Integer
Dim tmpPercent As Single
Dim tmpPercent2 As Single
Dim ErrorExists As Boolean

    ' Validate the new actual value boxes
    ErrorExists = False
    ' check all boxes
    For row = 1 To NumCalPoints
        If (Not RangeCheck((1.25 * sEuMax), sEuMin, txtNewActualValue(row), "Actual Value Entry")) Then ErrorExists = True
    Next row
    ' all boxes valid ??
    If Not ErrorExists Then
        ' all entries are valid; calibrate
        ' update new calibration DTS
        txtCalibDts.text = Format(Now, "YYYY-MMM-DD  hh:mm:ss")
        ' update new calibration from screen
        New_AiCal.NumPoints = CInt(txtNumCalPts.text)
        New_AiCal.Dts = CDate(txtCalibDts.text)
        New_AiCal.StandardTempValue = ValueFromText(txtStandardTemp.text)
        New_AiCal.StandardTempUnits = TempUnits.List(TempUnits.ListIndex)
        New_AiCal.StandardPressValue = ValueFromText(txtStandardPress.text)
        New_AiCal.StandardPressUnits = TempUnits.List(PressUnits.ListIndex)
        New_AiCal.CalibratedBy = txtCalibBy.text
        New_AiCal.Comment = txtComment.text
        ' update new calibration point data from screen
        For iPoint = 1 To NumCalPoints
            New_AiCal.PointData(iPoint).RawValue = ValueFromText(txtNewRawValue(iPoint).text)
            New_AiCal.PointData(iPoint).ActualValue = ValueFromText(txtNewActualValue(iPoint).text)
            tmpPercent = CSng(100) * ((New_AiCal.PointData(iPoint).RawValue - sRawMin) / sRawSpan)
            New_AiCal.PointData(iPoint).RawPercent = tmpPercent
            tmpPercent2 = CSng(100) * ((New_AiCal.PointData(iPoint).ActualValue - sEuMin) / sEuSpan)
            New_AiCal.PointData(iPoint).ActualPercent = tmpPercent2
        Next iPoint
    
        ' Copy Current to Prev
        If (Not bPrevCalEnabled) Then Prev_AiCal = Curr_AiCal
        ' Calculate Calibration Coefficients
        Calibrate
        ' Copy New to Current
        Curr_AiCal = New_AiCal
            
    '    ClearAiCal New_AiCal
        EnableNewCal True
        EnablePrevCal True
        bUnsavedCal = True
        blnPrevCalExists = True
        DisplayAiAll
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
        xlSht.Range("A" & i + 1) = (New_AiCal.PointData(i).RawPercent / CSng(100))
        xlSht.Range("B" & i + 1) = (New_AiCal.PointData(i).RawPercent / CSng(100)) ^ 2
        xlSht.Range("C" & i + 1) = (New_AiCal.PointData(i).RawPercent / CSng(100)) ^ 3
        xlSht.Range("D" & i + 1) = (New_AiCal.PointData(i).RawPercent / CSng(100)) ^ 4
        xlSht.Range("E" & i + 1) = (New_AiCal.PointData(i).RawPercent / CSng(100)) ^ 5
        xlSht.Range("F" & i + 1) = (New_AiCal.PointData(i).RawPercent / CSng(100)) ^ 6
        xlSht.Range("G" & i + 1) = New_AiCal.PointData(i).ActualPercent / CSng(100)
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
    New_AiCal.CalData.X = xlSht.Range("F14").Value
    New_AiCal.CalData.X2 = xlSht.Range("E14").Value
    New_AiCal.CalData.X3 = xlSht.Range("D14").Value
    New_AiCal.CalData.X4 = xlSht.Range("C14").Value
    New_AiCal.CalData.X5 = xlSht.Range("B14").Value
    New_AiCal.CalData.X6 = xlSht.Range("A14").Value
    
    New_AiCal.CalData.R2 = xlSht.Range("A16").Value

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
    FormulaText = FormulaText & "Y = " & Curr_AiCal.CalData.X6 & "X6"
    FormulaText = FormulaText & IIf(Curr_AiCal.CalData.X5 < 0, " - ", " + ") & Abs(Curr_AiCal.CalData.X5) & "X5"
    FormulaText = FormulaText & IIf(Curr_AiCal.CalData.X4 < 0, " - ", " + ") & Abs(Curr_AiCal.CalData.X4) & "X4"
    FormulaText = FormulaText & IIf(Curr_AiCal.CalData.X3 < 0, " - ", " + ") & Abs(Curr_AiCal.CalData.X3) & "X3"
    FormulaText = FormulaText & IIf(Curr_AiCal.CalData.X2 < 0, " - ", " + ") & Abs(Curr_AiCal.CalData.X2) & "X2"
    FormulaText = FormulaText & IIf(Curr_AiCal.CalData.X < 0, " - ", " + ") & Abs(Curr_AiCal.CalData.X) & "X"
    FormulaText = FormulaText & vbCrLf
    FormulaText = FormulaText & "      R2=" & Curr_AiCal.CalData.R2
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
    
    ' get min/max EU & Raw  for appropriate input
    Select Case SelCalGrp
        Case calgrpComm
            ' Common Analog Input Calibration Parameters
            sEuMax = Com_AIO(SelCalInp).EuMax
            sEuMin = Com_AIO(SelCalInp).EuMin
        Case calgrpStn1 To calgrpStn9
            ' Station Analog Input Calibration Parameters
            idx = SelCalGrp
            sEuMax = Stn_AIO(idx, SelCalInp).EuMax
            sEuMin = Stn_AIO(idx, SelCalInp).EuMin
        Case calgrpFid
            ' FID Analog Input Calibration Parameters
'            sEuMax = Fid_AIO(SelCalInp).EuMax
'            sEuMin = Fid_AIO(SelCalInp).EuMin
        Case calgrpPrg1 To calgrpPrg9
            ' Purge Analog Input Calibration Parameters
            idx = SelCalGrp - 10
            sEuMax = Prg_AIO(idx, SelCalInp).EuMax
            sEuMin = Prg_AIO(idx, SelCalInp).EuMin
    End Select
    sEuSpan = sEuMax - sEuMin
    
    For iPoint = 1 To NumCalPoints
        Graph(iPoint, 1) = Curr_AiCal.PointData(iPoint).RawPercent                              ' value for X-axis - Pen #1
        Graph(iPoint, 2) = CSng(Curr_AiCal.PointData(iPoint).RawPercent)                        ' value for Y-axis - Pen #1
        Graph(iPoint, 3) = Curr_AiCal.PointData(iPoint).RawPercent                              ' value for X-axis - Pen #2
        Graph(iPoint, 4) = CSng(Curr_AiCal.PointData(iPoint).ActualPercent)                     ' value for Y-axis - Pen #2
        Graph(iPoint, 5) = Curr_AiCal.PointData(iPoint).RawPercent                              ' value for X-axis - Pen #3
        tmpVal = Cal_AnalogInput(Curr_AiCal.PointData(iPoint).RawPercent / 100, SelCalGrp, SelCalInp, Curr_AiCal)
        If (sEuSpan <> 0) Then Graph(iPoint, 6) = CSng(100) * ((tmpVal - sEuMin) / sEuSpan)     ' value for Y-axis - Pen #3
    Next iPoint
         
    chtAIChart.chartType = VtChChartType2dXY  ' set to X Y Scatter chart
    chtAIChart = Graph ' populate chart's data grid using Graph array
    chtAIChart.Plot.UniformAxis = False
    chtAIChart.Column = 1
    chtAIChart.ColumnLabel = "Raw Value"
    chtAIChart.Column = 3
    chtAIChart.ColumnLabel = "Actual Value"
    chtAIChart.Column = 5
    chtAIChart.ColumnLabel = "Calib. Value"
    chtAIChart.Visible = True
End Sub

Private Sub cmdRestorePrevCal_Click()
    ' Undoes the effect of the calibration by transferring data from
    ' the previous calibration to the current calibration, transferring data
    ' from the temporary calibration buffer to the previous calibration,
    ' and updating the form
    
    ' shift cal point data from prev to current
'    ClearAiCal New_AiCal
    New_AiCal = Curr_AiCal
'    ClearAiCal Curr_AiCal
    Curr_AiCal = Prev_AiCal
    ClearAiCal Prev_AiCal
    EnablePrevCal False
    bUnsavedCal = False
    ' update display
    DisplayAiCalibration
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
        cmdGroupDn.Enabled = True
        cmdGroupUp.Enabled = True
        cmdInputDn.Enabled = True
        cmdInputUp.Enabled = True
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
        cmdAcquireRaw.Visible = False
        cmdAcquireActual.Visible = False
    Else
        cmdGroupDn.Enabled = True
        cmdGroupUp.Enabled = True
        cmdInputDn.Enabled = True
        cmdInputUp.Enabled = True
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
        cmdAcquireRaw.Visible = IIf(bNewCalEnabled And ((New_AiCal.Method = calmetRawOnly) Or (New_AiCal.Method = calmetRawAndActual)), True, False)
        cmdAcquireActual.Visible = IIf(bNewCalEnabled And ((New_AiCal.Method = calmetActualOnly) Or (New_AiCal.Method = calmetRawAndActual)), True, False)
    End If
End Sub

Private Sub UpdateAiSelection()
Dim idx As Integer
    ClearAiCal New_AiCal
    ClearAiCal Curr_AiCal
    ClearAiCal Prev_AiCal
    Select Case SelCalGrp
        Case calgrpComm
            Curr_AiCal = Com_AiCal(SelCalInp)
        Case calgrpStn1 To calgrpStn9
            idx = SelCalGrp
            Curr_AiCal = Stn_AiCal(idx, SelCalInp)
        Case calgrpFid
'            Curr_AiCal = Fid_AiCal(SelCalInp)
        Case calgrpPrg1 To calgrpPrg9
            idx = SelCalGrp - 10
            Curr_AiCal = Prg_AiCal(idx, SelCalInp)
    End Select
    NumCalPoints = Curr_AiCal.NumPoints
    If (Curr_AiCal.RawInputType = CalRawUndefined) Then Curr_AiCal.RawInputType = CalRawAsVolts
    CalcSpans Curr_AiCal.RawInputType
    ' init cal point columns
    EnableNewCal False
    EnablePrevCal False
End Sub

Private Sub DisplayAiAll()
    ' update the screen
    DisplayAiSelection useExistRawInputType
    HideInactiveTableRows
    DisplayAiCalibration
    UpdateCmdButtons
End Sub

Private Sub DisplayAiCalibration()
'
    ' Cal is ReadOnly unless all stations are idle
    If AllStationsIdle And CalReadOnly Then
        CalReadOnly = False
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "AI Calibration is Enabled"
    ElseIf Not AllStationsIdle And Not CalReadOnly Then
        CalReadOnly = True
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "AI Calibration is Read-Only"
    End If
    ' update the screen
    DisplayCurrCalInformation
    DisplayCalPointData
    DisplayCalFormula
    DisplayCalGraph
    
End Sub

Private Sub cmdSaveCurrCal_Click()
    ' Save the current calibration
    SaveCurrCalibration SelCalGrp, SelCalInp
    bUnsavedCal = False
    ' clear new & prev calibrations
    ClearAiCal New_AiCal
    ClearAiCal Prev_AiCal
    EnableNewCal False
    EnablePrevCal False
    blnPrevCalExists = False
    ' update screen
    DisplayAiCalibration
    ' update command buttons
    UpdateCmdButtons
'    Delay_Box "Calibration Saved", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = vbCrLf & "Calibration Saved"
End Sub

Private Sub Form_Load()
'
Dim idx As Integer
    bFormLoaded = False
    ' set display colors
    frmCalControls.ForeColor = Titles_ForeColor
    frmGroupSelection.ForeColor = Titles_ForeColor
    frmInputSelection.ForeColor = Titles_ForeColor
    frmCalFormula.ForeColor = Titles_ForeColor
    frmCalInformation.ForeColor = Titles_ForeColor
    frmCalPointData.ForeColor = Titles_ForeColor
    frmCalGraph.ForeColor = Titles_ForeColor
    txtNumCalPts.ForeColor = TitlesData_Forecolor
    txtDispGrp.ForeColor = TitlesData_Forecolor
    txtCalibBy.ForeColor = TitlesData_Forecolor
    txtEquipment.ForeColor = TitlesData_Forecolor
    txtComment.ForeColor = TitlesData_Forecolor
    For idx = 1 To MAXLSQCALPOINTS
        txtNewRawValue(idx).ForeColor = TitlesData_Forecolor
        txtNewActualValue(idx).ForeColor = TitlesData_Forecolor
    Next idx

    ' Cal is ReadOnly unless all stations are idle
    If AllStationsIdle Then
        CalReadOnly = False
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "Calibration is Enabled"
    Else
        CalReadOnly = True
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "Calibration is Read-Only"
    End If

    ' Set the current group and input
    SelCalGrp = 0
    SelCalInp = 0
    ' find the first valid analog input
    NextInput
    ' init cal point columns
    EnableNewCal False
    EnablePrevCal False
    ' init Unsaved Calibration flag
    bUnsavedCal = False
    blnPrevCalExists = False
    ' update the screen
    DisplayAiAll
    
    bFormLoaded = True
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

Private Sub DisplayAiSelection(ByVal updateRawInputType As Integer)
' DisplayAISelection
' Displays Information on the Currently Selected Analog Input
Dim idx As Integer
Dim sGrpDesc As String
Dim sAiDesc As String
Dim iRawInputType As Integer
    
    Select Case SelCalGrp
        Case calgrpComm
            sGrpDesc = "Common"
            sAiDesc = Com_AnaDef(SelCalInp).desc
            iRawInputType = Com_AiCal(SelCalInp).RawInputType
        Case calgrpStn1 To calgrpStn9
            idx = SelCalGrp
            sGrpDesc = "Station #" & Format(idx, "#0")
            sAiDesc = Stn_AnaDef(SelCalInp).desc
            iRawInputType = Stn_AiCal(idx, SelCalInp).RawInputType
        Case calgrpFid
            sGrpDesc = "FID"
'            sAiDesc = Fid_AnaDef(SelCalInp).desc
'            iRawInputType = Fid_AiCal(SelCalInp).RawInputType
        Case calgrpPrg1 To calgrpPrg9
            idx = SelCalGrp - 10
            sGrpDesc = "Purge #" & Format(idx, "#0")
            sAiDesc = Prg_AnaDef(SelCalInp).desc
            iRawInputType = Prg_AiCal(idx, SelCalInp).RawInputType
    End Select
    If (updateRawInputType = useNewRawInputType) Then iRawInputType = New_AiCal.RawInputType
    RawValues.ListIndex = iRawInputType - 1
    CalcSpans iRawInputType
    txtDispGrp.text = sGrpDesc
    txtaFuncDesc.text = sAiDesc
    txtaEUMax.text = Format(sEuMax, "####0.0##")
    txtaEUMin.text = Format(sEuMin, "####0.0##")
    Select Case iRawInputType
        Case CalRawAsMa
            lblaVdcMax.Caption = "ma Max"
            lblaVdcMin.Caption = "ma Min"
            txtaVDCMax.text = Format(sRawMax, "####0.0##")
            txtaVDCMin.text = Format(sRawMin, "####0.0##")
        Case CalRawAsVolts
            lblaVdcMax.Caption = "Vdc Max"
            lblaVdcMin.Caption = "Vdc Min"
            txtaVDCMax.text = Format(sRawMax, "####0.0##")
            txtaVDCMin.text = Format(sRawMin, "####0.0##")
        Case CalRawAsDegC
            lblaVdcMax.Caption = "degC Max"
            lblaVdcMin.Caption = "degC Min"
            txtaVDCMax.text = Format(sRawMax, "####0.0#")
            txtaVDCMin.text = Format(sRawMin, "####0.0#")
        Case CalRawAsEU
            lblaVdcMax.Caption = "Raw Max"
            lblaVdcMin.Caption = "Raw Min"
            txtaVDCMax.text = Format(sRawMax, "####0.0##")
            txtaVDCMin.text = Format(sRawMin, "####0.0##")
    End Select
    
End Sub

Private Sub lblPointNum_Click(Index As Integer)
' Selects the row that was clicked on, for editing (but don't repeat up/dn cmd)
    If SelCalPoint <> Index Then
        SelectCalPoint Index
'        txtNewActualValue(SelCalPoint).Enabled = True
'        If bFormLoaded Then txtNewActualValue(SelCalPoint).SetFocus
    End If
End Sub

Private Sub SelectRow(ByVal RowToSelect As Integer)
    ' Enables the radio button for the current AI
    Select Case New_AiCal.Method
        Case calmetRawOnly
            txtNewRawValue(RowToSelect).Enabled = True
'            txtNewActualValue(RowToSelect).Enabled = True
        Case calmetActualOnly
'            txtNewRawValue(RowToSelect).Enabled = True
            txtNewActualValue(RowToSelect).Enabled = True
        Case calmetRawAndActual
            txtNewRawValue(RowToSelect).Enabled = True
            txtNewActualValue(RowToSelect).Enabled = True
    End Select
    lblPointNum(RowToSelect).Appearance = 1    ' 3D
    lblPointNum(RowToSelect).FontBold = True
    lblPointNum(RowToSelect).BackColor = PALEBLUE
End Sub

Private Sub DisableRow(ByVal RowToDisable As Integer)
    ' Unselects every element in this row
    txtNewRawValue(RowToDisable).Enabled = False
    txtNewActualValue(RowToDisable).Enabled = False
    lblPointNum(RowToDisable).Appearance = 0    ' Flat
    lblPointNum(RowToDisable).FontBold = False
    lblPointNum(RowToDisable).BackColor = Common_BackColor
End Sub

Private Sub SelectCalPoint(ByVal SelectedRow As Integer)
' Updates the MFC calibration table settings
' based on the current point(row) selected
Dim iRow As Integer
    
    SelCalPoint = SelectedRow
    For iRow = 1 To NumCalPoints
        If iRow = SelCalPoint Then
            SelectRow (iRow)
        Else
            DisableRow (iRow)
        End If
    Next iRow
End Sub

Private Sub RawValues_Click()
Dim idx As Integer
    idx = RawValues.ItemData(RawValues.ListIndex)
    New_AiCal.RawInputType = idx
    Curr_AiCal.RawInputType = idx
    Prev_AiCal.RawInputType = idx
    DisplayAiSelection useNewRawInputType
    DisplayAiCalibration
End Sub

Private Sub CalEntries_Click()
Dim idx As Integer
    idx = CalEntries.ItemData(CalEntries.ListIndex)
    New_AiCal.Method = idx
    Curr_AiCal.Method = idx
    Prev_AiCal.Method = idx
    UpdateCmdButtons
End Sub

Private Sub txtCalibBy_Change()
'    bUnsavedCal = True
End Sub

Private Sub txtComment_Change()
'    bUnsavedCal = True
End Sub

Private Sub CalcSpans(ByVal iRawType As Integer)
'
' CalRawUndefined = 0
' CalRawAsVolts = 1              ' Analog Input Entered Raw Values are in Volts
' CalRawAsMa = 2                 ' Analog Input Entered Raw Values are in MilliAmperes
' CalRawAsDegC = 3               ' Analog Input Entered Raw Values are in Degrees C
' CalRawAsEU = 4                 ' Analog Input Entered Raw Values are in Engr Units
'
Dim idx As Integer
Dim Sel_AIO As FuncAnalogIO                                        ' Analog IO Information
    
    Select Case SelCalGrp
        Case calgrpComm
            Sel_AIO = Com_AIO(SelCalInp)
        Case calgrpStn1 To calgrpStn9
            idx = SelCalGrp
            Sel_AIO = Stn_AIO(idx, SelCalInp)
        Case calgrpFid
'            Sel_AIO = Fid_AIO(SelCalInp)
        Case calgrpPrg1 To calgrpPrg9
            idx = SelCalGrp - 10
            Sel_AIO = Prg_AIO(idx, SelCalInp)
        Case Else
            Exit Sub
    End Select
    Select Case iRawType
        Case CalRawAsVolts  ' 0-5 volts
            sRawMax = Sel_AIO.VdcMax
            sEuMax = Sel_AIO.EuMax
            sRawMin = Sel_AIO.VdcMin
            sEuMin = Sel_AIO.EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
        Case CalRawAsMa     ' 0-20ma (converted from Vdc)
            sRawMax = CSng(4) * Sel_AIO.VdcMax
            sEuMax = Sel_AIO.EuMax
            sRawMin = CSng(4) * Sel_AIO.VdcMin
            sEuMin = Sel_AIO.EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
        Case CalRawAsDegC   ' Opto TC & RTD modules return temp as ***.0 degC
            sRawMax = Sel_AIO.EuMax
            sEuMax = Sel_AIO.EuMax
            sRawMin = Sel_AIO.EuMin
            sEuMin = Sel_AIO.EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
        Case CalRawAsEU     ' Raw range = EU range
            sRawMax = Sel_AIO.EuMax
            sEuMax = Sel_AIO.EuMax
            sRawMin = Sel_AIO.EuMin
            sEuMin = Sel_AIO.EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
    End Select
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
        ActualValue = Curr_AiCal.PointData(iPoint).ActualValue
        ' Calibrated Value
        CalibValue = Cal_AnalogInput((Curr_AiCal.PointData(iPoint).ActualPercent / CSng(100)), SelCalGrp, SelCalInp, Curr_AiCal)
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
        ActualValue = Prev_AiCal.PointData(iPoint).ActualValue
        ' Calibrated Value
        CalibValue = Cal_AnalogInput((Prev_AiCal.PointData(iPoint).ActualPercent / CSng(100)), SelCalGrp, SelCalInp, Prev_AiCal)
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

Public Sub SaveCurrCalibration(ByVal iGroup As Integer, ByVal iInput As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date
Dim iPoint As Integer
Dim idx As Integer

    ' delete existing calibration
    ClearAiCalRecords iGroup, iInput

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Save Analog Input Calibration Parameters
    Criteria = "SELECT * FROM [AiCalibrations] WHERE [Group] = " & iGroup & " AND [Input] = " & iInput & "  ORDER BY [Dts] DESC"
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
    rsRecord("Group") = iGroup
    rsRecord("Input") = iInput
    rsRecord("Dts") = Curr_AiCal.Dts
    rsRecord("CalibratedBy") = Curr_AiCal.CalibratedBy
    rsRecord("Equipment") = Curr_AiCal.Equipment
    rsRecord("Comment") = Curr_AiCal.Comment
    rsRecord("NumPoints") = Curr_AiCal.NumPoints
    rsRecord("RawInputType") = Curr_AiCal.RawInputType
    rsRecord("CoefficientX") = Curr_AiCal.CalData.X
    rsRecord("CoefficientX2") = Curr_AiCal.CalData.X2
    rsRecord("CoefficientX3") = Curr_AiCal.CalData.X3
    rsRecord("CoefficientX4") = Curr_AiCal.CalData.X4
    rsRecord("CoefficientX5") = Curr_AiCal.CalData.X5
    rsRecord("CoefficientX6") = Curr_AiCal.CalData.X6
    rsRecord("CoefficientR2") = Curr_AiCal.CalData.R2
    rsRecord("StandardTempValue") = Curr_AiCal.StandardTempValue
    rsRecord("StandardTempUnits") = Curr_AiCal.StandardTempUnits
    rsRecord("StandardPressValue") = Curr_AiCal.StandardPressValue
    rsRecord("StandardPressUnits") = Curr_AiCal.StandardPressUnits
    rsRecord.Update
    rsRecord.Close

            
    ' Save Analog Input Calibration Point Data
    dDts = Curr_AiCal.Dts
    CriteriaPts = "SELECT * FROM [AiCalibrationsData] WHERE [Group] = " & iGroup & " AND [Input] = " & iInput & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
    Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
            
    ' update the calibration information
    For iPoint = 1 To NumCalPoints
        rsRecordPts.AddNew
        rsRecordPts("Group") = iGroup
        rsRecordPts("Input") = iInput
        rsRecordPts("Point") = iPoint
        rsRecordPts("Dts") = dDts
        rsRecordPts("ActualPercent") = Curr_AiCal.PointData(iPoint).ActualPercent
        rsRecordPts("RawPercent") = Curr_AiCal.PointData(iPoint).RawPercent
        rsRecordPts("ActualValue") = Curr_AiCal.PointData(iPoint).ActualValue
        rsRecordPts("RawValue") = Curr_AiCal.PointData(iPoint).RawValue
        rsRecordPts.Update
    Next iPoint
    ' done with points
    rsRecordPts.Close
                        
    ' copy calibration to appropriate group array
    Select Case iGroup
        Case calgrpComm
            ' Common Analog Input Calibration Parameters
            PrevCom_AiCal(iInput) = Com_AiCal(iInput)
            Com_AiCal(iInput) = Curr_AiCal
        Case calgrpStn1 To calgrpStn9
            ' Station Analog Input Calibration Parameters
            PrevStn_AiCal(iGroup, iInput) = Stn_AiCal(iGroup, iInput)
            Stn_AiCal(iGroup, iInput) = Curr_AiCal
        Case calgrpFid
            ' FID Analog Input Calibration Parameters
'            PrevFid_AiCal(iInput) = Fid_AiCal(iInput)
'            Fid_AiCal(iInput) = Curr_AiCal
        Case calgrpPrg1 To calgrpPrg9
            ' Purge Analog Input Calibration Parameters
            idx = iGroup - 10
            PrevPrg_AiCal(idx, iInput) = Prg_AiCal(idx, iInput)
            Prg_AiCal(idx, iInput) = Curr_AiCal
    End Select
                
End Sub

Private Sub PrintCalibration()
' Print the values found for the current and (optionally) the previous calibrations
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 1881, 22
Dim tmpString(0 To 49) As String
Dim descCol As Integer
Dim dataLeft As Integer
Dim dataRight As Integer
Dim dataWidth As Integer
Dim iLine As Integer
Dim idx As Integer
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
Dim CurrCal As AICalibration
Dim PrevCal As AICalibration
Dim bPrintPrevCal As Boolean
Dim oldFont As New StdFont

    ' current & previous calibration
    Select Case SelCalGrp
        Case calgrpComm
            CurrCal = Com_AiCal(SelCalInp)
            PrevCal = PrevCom_AiCal(SelCalInp)
        Case calgrpStn1 To calgrpStn9
            idx = SelCalGrp
            CurrCal = Stn_AiCal(idx, SelCalInp)
            PrevCal = PrevStn_AiCal(idx, SelCalInp)
        Case calgrpFid
'            CurrCal = Fid_AiCal(SelCalInp)
'            PrevCal = PrevFid_AiCal(SelCalInp)
        Case calgrpPrg1 To calgrpPrg9
            idx = SelCalGrp - 10
            CurrCal = Prg_AiCal(idx, SelCalInp)
            PrevCal = PrevPrg_AiCal(idx, SelCalInp)
    End Select
    
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
    Select Case SelCalGrp
        Case calgrpComm
            idx = SelCalGrp
            Print_Center "Calibration Report for Common " & Com_AnaDef(SelCalInp).desc
        Case calgrpStn1 To calgrpStn9
            idx = SelCalGrp
            Print_Center "Calibration Report for Station # " & Format(idx, "#0") & ", " & Stn_AnaDef(SelCalInp).desc
        Case calgrpFid
            idx = SelCalGrp
'            Print_Center "Calibration Report for FID " & Fid_AnaDef(SelCalInp).desc
        Case calgrpPrg1 To calgrpPrg9
            idx = SelCalGrp - 10
            Print_Center "Calibration Report for Purge # " & Format(idx, "#0") & ", " & Prg_AnaDef(SelCalInp).desc
        Case Else
            Print_Center "Calibration Report"
    End Select
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
'CurrCal.Comment = "12345678901234567890123456789012345678901234567890"
'CurrCal.Comment = CurrCal.Comment & "12345678901234567890123456789012345678901234567890"
'CurrCal.Comment = CurrCal.Comment & "1234567"
'PrevCal.Comment = CurrCal.Comment & "89"
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
            sngCalibValue = Cal_AnalogInput((CurrCal.PointData(row).ActualPercent / 100), SelCalGrp, SelCalInp, CurrCal)
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
                sngCalibValue = Cal_AnalogInput((PrevCal.PointData(row).ActualPercent / 100), SelCalGrp, SelCalInp, PrevCal)
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
    NumCalPoints = Curr_AiCal.NumPoints
    txtNumCalPts.text = Format(NumCalPoints, "#0")
    txtNumCalPts.ForeColor = Black
    txtNumCalPts.BackColor = frmCalInformation.BackColor
    txtCalibDts.text = Format(Curr_AiCal.Dts, "YYYY-MMM-DD hh:mm:ss")
    txtStandardTemp.text = Format(Curr_AiCal.StandardTempValue, "###0.0")
    TempUnits.ListIndex = TempUnitsIndex(Curr_AiCal.StandardTempUnits)
    txtStandardPress.text = Format(Curr_AiCal.StandardPressValue, "###0.0##")
    PressUnits.ListIndex = PressUnitsIndex(Curr_AiCal.StandardPressUnits)
    txtCalibBy.text = Curr_AiCal.CalibratedBy
    txtEquipment.text = "equipment"
    txtComment.text = Curr_AiCal.Comment
    TempUnits.Refresh
    PressUnits.Refresh
End Sub

Private Sub DisplayNewCalInformation()
' Displays new calibration information
    NumCalPoints = New_AiCal.NumPoints
    txtNumCalPts.text = Format(NumCalPoints, "#0")
    txtNumCalPts.ForeColor = Black
    txtNumCalPts.BackColor = frmCalInformation.BackColor
    txtCalibDts.text = Format(New_AiCal.Dts, "YYYY-MMM-DD hh:mm:ss")
    txtStandardTemp.text = Format(New_AiCal.StandardTempValue, "###0.0")
    TempUnits.ListIndex = TempUnitsIndex(New_AiCal.StandardTempUnits)
    txtStandardPress.text = Format(New_AiCal.StandardPressValue, "###0.0##")
    PressUnits.ListIndex = PressUnitsIndex(New_AiCal.StandardPressUnits)
    txtCalibBy.text = New_AiCal.CalibratedBy
    txtEquipment.text = "equipment"
    txtComment.text = New_AiCal.Comment
    TempUnits.Refresh
    PressUnits.Refresh
End Sub

Private Sub DisplayCalPointData()
' Displays new, current & previous calibration point data
    Dim iPoint As Integer
    HideInactiveTableRows
    For iPoint = 1 To NumCalPoints
        lblPointNum(iPoint).Caption = Format(iPoint, "#0")
        txtNewRawValue(iPoint).text = IIf(txtNewRawValue(iPoint).Enabled, Format(New_AiCal.PointData(iPoint).RawValue, "####0.0##"), "")
        lblNewRawPerc(iPoint).Caption = IIf(lblNewRawPerc(iPoint).Enabled, Format(New_AiCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        txtNewActualValue(iPoint).text = IIf(txtNewActualValue(iPoint).Enabled, Format(New_AiCal.PointData(iPoint).ActualValue, "####0.0##"), "")
        lblCurrRawPerc(iPoint).Caption = IIf(lblCurrRawPerc(iPoint).Enabled, Format(Curr_AiCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        lblCurrActualValue(iPoint).Caption = IIf(lblCurrActualValue(iPoint).Enabled, Format(Curr_AiCal.PointData(iPoint).ActualValue, "####0.0##"), "")
        lblPrevRawPerc(iPoint).Caption = IIf(lblPrevRawPerc(iPoint).Enabled, Format(Prev_AiCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        lblPrevActualValue(iPoint).Caption = IIf(lblPrevActualValue(iPoint).Enabled, Format(Prev_AiCal.PointData(iPoint).ActualValue, "####0.0##"), "")
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
    ' close screen
    Unload Me
    Set frmAnalogInputCal = Nothing
End Sub

Private Sub txtNewActualValue_Change(Index As Integer)
    txtNewActualValue(Index).ForeColor = TitlesData_Forecolor
    txtNewActualValue(Index).BackColor = Entry_BackColor
End Sub

Private Sub txtNewRawValue_Change(Index As Integer)
Dim mult As Single
    txtNewRawValue(Index).ForeColor = TitlesData_Forecolor
    txtNewRawValue(Index).BackColor = Entry_BackColor
    mult = (ValueFromText(txtNewRawValue(Index).text) - sRawMin) / sRawSpan
    lblNewRawPerc(Index).Caption = Format((CSng(100) * mult), "###0.0##")
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
SetErrModule 1181, 7
If UseLocalErrorHandler Then On Error GoTo localhandler

    If (tbox.text = "") Then tbox.text = "0"
    If (Not IsNumeric(tbox.text)) Then tbox.text = "0"
    If CStr(tbox) < loVal Or CStr(tbox) > hiVal Then
        RangeCheck = False
        tbox.BackColor = EntryInvalid_BackColor
        tbox.ForeColor = Alarm_ForeColor
        txtMsg.ForeColor = Alarm_ForeColor
        txtMsg.text = sMessage & " Range Error!" & vbCrLf _
                    & "Allowable Range = " & loVal & " - " & hiVal
        MyFocus tbox
    Else
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

Private Function RawValuesIndex(ByVal unitsText As String) As Integer
' returns ListIndex of RawValues that matches unitsText
' defaults to 0
Dim idx As Integer
Dim idxList As Integer

    idxList = 0
    For idx = 0 To (RawValues.ListCount - 1)
        If (Trim(unitsText) = Trim(RawValues.List(idx))) Then idxList = idx
    Next idx
    RawValuesIndex = idxList
End Function


