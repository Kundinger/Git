VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form frmMfcCal 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MFC Calibration Screen"
   ClientHeight    =   11160
   ClientLeft      =   195
   ClientTop       =   720
   ClientWidth     =   15330
   Icon            =   "FrmMfcCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   15330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   16560
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   220
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   4200
      Top             =   6840
   End
   Begin VB.Frame frmQuestion 
      Caption         =   "Copy Calibration ???"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   6840
      TabIndex        =   209
      Top             =   2400
      Width           =   6615
      Begin VB.CommandButton cmdOK 
         Caption         =   "Copy Calibration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmMfcCal.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Do Not Copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmMfcCal.frx":5B24
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.Label lblQuestion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Live Fuel ORVR ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   214
         Top             =   990
         Width           =   5775
      End
      Begin VB.Label lblQuestion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   213
         Top             =   735
         Width           =   5775
      End
      Begin VB.Label lblQuestion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Nitrogen ORVR MFC Calibration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   212
         Top             =   480
         Width           =   5775
      End
   End
   Begin VB.CommandButton cmdPointNum 
      Caption         =   "1"
      Height          =   285
      Index           =   0
      Left            =   16320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   4200
      Width           =   495
   End
   Begin VB.Frame frmGroupSelection 
      Caption         =   "Station Selection"
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
         Picture         =   "FrmMfcCal.frx":5E66
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "next station"
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
         Picture         =   "FrmMfcCal.frx":6568
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "previous station"
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
      Begin VB.CommandButton cmdCalCheck 
         Caption         =   "Calibration Check"
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
         Left            =   6000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmMfcCal.frx":6C6A
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdCalPoints 
         Caption         =   "Edit CalPoints"
         DisabledPicture =   "FrmMfcCal.frx":6FAC
         DownPicture     =   "FrmMfcCal.frx":72EE
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
         Picture         =   "FrmMfcCal.frx":7630
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdPrintCal 
         Caption         =   "Print"
         DisabledPicture =   "FrmMfcCal.frx":7972
         DownPicture     =   "FrmMfcCal.frx":8074
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
         Picture         =   "FrmMfcCal.frx":8776
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "Print a Report of the Selected Calibration"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdRestorePrevCal 
         Caption         =   "Restore Previous"
         DisabledPicture =   "FrmMfcCal.frx":8E78
         DownPicture     =   "FrmMfcCal.frx":91BA
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
         Picture         =   "FrmMfcCal.frx":94FC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdSaveCurrCal 
         Caption         =   "     Save      Current"
         DisabledPicture =   "FrmMfcCal.frx":983E
         DownPicture     =   "FrmMfcCal.frx":9B80
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
         Picture         =   "FrmMfcCal.frx":9EC2
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   950
      End
      Begin VB.CommandButton cmdCreateNewCalib 
         Caption         =   " Create New"
         DisabledPicture =   "FrmMfcCal.frx":A204
         DownPicture     =   "FrmMfcCal.frx":A546
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
         Picture         =   "FrmMfcCal.frx":A888
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
         Text            =   "FrmMfcCal.frx":ABCA
         Top             =   1080
         Width           =   6795
      End
      Begin VB.CommandButton cmdRunCal 
         Caption         =   "Calibration"
         DisabledPicture =   "FrmMfcCal.frx":ABF5
         DownPicture     =   "FrmMfcCal.frx":AF37
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
         Picture         =   "FrmMfcCal.frx":B279
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
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmMfcCal.frx":B5BB
         Style           =   1  'Graphical
         TabIndex        =   217
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
         ItemData        =   "FrmMfcCal.frx":C1FD
         Left            =   3120
         List            =   "FrmMfcCal.frx":C20D
         Style           =   2  'Dropdown List
         TabIndex        =   183
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
         TabIndex        =   181
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
         ItemData        =   "FrmMfcCal.frx":C228
         Left            =   3120
         List            =   "FrmMfcCal.frx":C238
         Style           =   2  'Dropdown List
         TabIndex        =   180
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
         TabIndex        =   178
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
         TabIndex        =   177
         Text            =   "FrmMfcCal.frx":C258
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
         TabIndex        =   176
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
         TabIndex        =   175
         Text            =   "FrmMfcCal.frx":C262
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
         TabIndex        =   174
         Text            =   "FrmMfcCal.frx":C26C
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
         Picture         =   "FrmMfcCal.frx":C277
         Style           =   1  'Graphical
         TabIndex        =   158
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
         TabIndex        =   182
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
         TabIndex        =   179
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
         DisabledPicture =   "FrmMfcCal.frx":C5B9
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
         Picture         =   "FrmMfcCal.frx":C8FB
         Style           =   1  'Graphical
         TabIndex        =   219
         ToolTipText     =   "Copy Current I/O Value to Selected Actual Entry"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   900
      End
      Begin VB.CommandButton cmdAcquireRaw 
         DisabledPicture =   "FrmMfcCal.frx":CC3D
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
         Picture         =   "FrmMfcCal.frx":CF7F
         Style           =   1  'Graphical
         TabIndex        =   218
         ToolTipText     =   "Copy Current I/O Value to Selected Raw Entry"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "11"
         Height          =   285
         Index           =   11
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   3690
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "10"
         Height          =   285
         Index           =   10
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   3405
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "9"
         Height          =   285
         Index           =   9
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   3120
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "8"
         Height          =   285
         Index           =   8
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   2835
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "7"
         Height          =   285
         Index           =   7
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   204
         Top             =   2550
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "6"
         Height          =   285
         Index           =   6
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   2265
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "5"
         Height          =   285
         Index           =   5
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   1980
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "4"
         Height          =   285
         Index           =   4
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   1695
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "3"
         Height          =   285
         Index           =   3
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   1410
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         Caption         =   "2"
         Height          =   285
         Index           =   2
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   1125
         Width           =   495
      End
      Begin VB.CommandButton cmdPointNum 
         BackColor       =   &H00FFFFC0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   198
         Top             =   840
         Width           =   495
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   170
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
         TabIndex        =   169
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
         TabIndex        =   168
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
         TabIndex        =   167
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
         TabIndex        =   166
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
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   162
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
         TabIndex        =   153
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         TabIndex        =   146
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
         TabIndex        =   145
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
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   141
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   135
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
         TabIndex        =   134
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         TabIndex        =   117
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         TabIndex        =   114
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
         TabIndex        =   113
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
         TabIndex        =   112
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
         TabIndex        =   111
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
         TabIndex        =   110
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
         TabIndex        =   109
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
         TabIndex        =   108
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
         TabIndex        =   107
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
         TabIndex        =   106
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
         TabIndex        =   105
         Top             =   600
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
      Caption         =   "MFC Selection"
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
         ItemData        =   "FrmMfcCal.frx":D2C1
         Left            =   3195
         List            =   "FrmMfcCal.frx":D2CE
         Style           =   2  'Dropdown List
         TabIndex        =   221
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
         ItemData        =   "FrmMfcCal.frx":D2F7
         Left            =   3195
         List            =   "FrmMfcCal.frx":D307
         Style           =   2  'Dropdown List
         TabIndex        =   215
         ToolTipText     =   "Units for Raw Entries"
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
         Picture         =   "FrmMfcCal.frx":D325
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "previous mfc"
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
         Picture         =   "FrmMfcCal.frx":DA27
         Style           =   1  'Graphical
         TabIndex        =   160
         ToolTipText     =   "next mfc"
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
         TabIndex        =   159
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
         Left            =   1680
         TabIndex        =   222
         Top             =   1320
         Width           =   1520
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
         Left            =   1680
         TabIndex        =   216
         Top             =   960
         Width           =   1520
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
      TabIndex        =   155
      Top             =   120
      Width           =   8055
      Begin MSChart20Lib.MSChart chtMfcChart 
         Height          =   6435
         Left            =   60
         OleObjectBlob   =   "FrmMfcCal.frx":E129
         TabIndex        =   156
         Top             =   180
         Width           =   7920
      End
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   17190
      TabIndex        =   197
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   17190
      TabIndex        =   196
      Top             =   885
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   3
      Left            =   17190
      TabIndex        =   195
      Top             =   1170
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   4
      Left            =   17190
      TabIndex        =   194
      Top             =   1455
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   5
      Left            =   17190
      TabIndex        =   193
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   6
      Left            =   17190
      TabIndex        =   192
      Top             =   2025
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   7
      Left            =   17190
      TabIndex        =   191
      Top             =   2310
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   8
      Left            =   17190
      TabIndex        =   190
      Top             =   2595
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   9
      Left            =   17190
      TabIndex        =   189
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   10
      Left            =   17190
      TabIndex        =   188
      Top             =   3165
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   11
      Left            =   17190
      TabIndex        =   187
      Top             =   3450
      Width           =   495
   End
   Begin VB.Label lblPointNum 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   0
      Left            =   16320
      TabIndex        =   186
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "frmMfcCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''ERROR module 7781
''
''frmMfcCal
''
Option Explicit

Private blnPrevCalExists As Boolean                       ' Flag - Whether previous calibration data is loaded
' MFC calibration variables
Private SelCalStn As Integer                              ' The Index of the Station currently selected; 1-9=Station#
Private SelCalMfc As Integer                              ' The Index of the MFC currently selected; 1-6
Private SelCalFunc As Integer                             ' The AIO Index of the MFC currently selected
Private SelCalPoint As Integer                            ' The Index of the currently selected Calibration Point(row)
Private NumCalPoints As Integer                           ' The Number of Data Points for the currently selected MFC
Private MfcOptoType As Integer                            ' The Type of Opto22 Hardware Module for the currently selected MFC
Private srcMfc As Integer
Private dstMfc As Integer
Private srcFunc As Integer
Private dstFunc As Integer
Private sRawMax As Single
Private sEuMax As Single
Private sRawMin As Single
Private sEuMin As Single
Private sRawSpan As Single
Private sEuSpan As Single
Private Curr_MfcCal As MfcCalibration
Private New_MfcCal As MfcCalibration
Private Prev_MfcCal As MfcCalibration
' Max Station Index
Const MAXSTN = 9
' Min Station Index
Const MINSTN = 1
' Max MFCs per Station
Const MAXINP = MAXMFC
' MFC Selecton Update Options
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
'
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
    ' Define every MFC calibration point
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 7781, 1
Dim iPoint As Integer
Dim Idx As Integer
    
    Select Case NumCalPoints
        Case 3
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(50)
            New_MfcCal.PointData(3).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(50)
            New_MfcCal.PointData(3).ActualPercent = CSng(100)
        Case 4
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(30)
            New_MfcCal.PointData(3).RawPercent = CSng(70)
            New_MfcCal.PointData(4).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(30)
            New_MfcCal.PointData(3).ActualPercent = CSng(70)
            New_MfcCal.PointData(4).ActualPercent = CSng(100)
        Case 5
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(25)
            New_MfcCal.PointData(3).RawPercent = CSng(50)
            New_MfcCal.PointData(4).RawPercent = CSng(75)
            New_MfcCal.PointData(5).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(25)
            New_MfcCal.PointData(3).ActualPercent = CSng(50)
            New_MfcCal.PointData(4).ActualPercent = CSng(75)
            New_MfcCal.PointData(5).ActualPercent = CSng(100)
        Case 6
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(20)
            New_MfcCal.PointData(3).RawPercent = CSng(40)
            New_MfcCal.PointData(4).RawPercent = CSng(60)
            New_MfcCal.PointData(5).RawPercent = CSng(80)
            New_MfcCal.PointData(6).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(20)
            New_MfcCal.PointData(3).ActualPercent = CSng(40)
            New_MfcCal.PointData(4).ActualPercent = CSng(60)
            New_MfcCal.PointData(5).ActualPercent = CSng(80)
            New_MfcCal.PointData(6).ActualPercent = CSng(100)
        Case 7
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(16.7)
            New_MfcCal.PointData(3).RawPercent = CSng(33.3)
            New_MfcCal.PointData(4).RawPercent = CSng(50)
            New_MfcCal.PointData(5).RawPercent = CSng(66.7)
            New_MfcCal.PointData(6).RawPercent = CSng(83.3)
            New_MfcCal.PointData(7).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(16.7)
            New_MfcCal.PointData(3).ActualPercent = CSng(33.3)
            New_MfcCal.PointData(4).ActualPercent = CSng(50)
            New_MfcCal.PointData(5).ActualPercent = CSng(66.7)
            New_MfcCal.PointData(6).ActualPercent = CSng(83.3)
            New_MfcCal.PointData(7).ActualPercent = CSng(100)
        Case 8
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(14.3)
            New_MfcCal.PointData(3).RawPercent = CSng(28.6)
            New_MfcCal.PointData(4).RawPercent = CSng(42.9)
            New_MfcCal.PointData(5).RawPercent = CSng(57.1)
            New_MfcCal.PointData(6).RawPercent = CSng(71.4)
            New_MfcCal.PointData(7).RawPercent = CSng(85.7)
            New_MfcCal.PointData(8).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(14.3)
            New_MfcCal.PointData(3).ActualPercent = CSng(28.6)
            New_MfcCal.PointData(4).ActualPercent = CSng(42.9)
            New_MfcCal.PointData(5).ActualPercent = CSng(57.1)
            New_MfcCal.PointData(6).ActualPercent = CSng(71.4)
            New_MfcCal.PointData(7).ActualPercent = CSng(85.7)
            New_MfcCal.PointData(8).ActualPercent = CSng(100)
        Case 9
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(12.5)
            New_MfcCal.PointData(3).RawPercent = CSng(25)
            New_MfcCal.PointData(4).RawPercent = CSng(37.5)
            New_MfcCal.PointData(5).RawPercent = CSng(50)
            New_MfcCal.PointData(6).RawPercent = CSng(62.5)
            New_MfcCal.PointData(7).RawPercent = CSng(75)
            New_MfcCal.PointData(8).RawPercent = CSng(87.5)
            New_MfcCal.PointData(9).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(12.5)
            New_MfcCal.PointData(3).ActualPercent = CSng(25)
            New_MfcCal.PointData(4).ActualPercent = CSng(37.5)
            New_MfcCal.PointData(5).ActualPercent = CSng(50)
            New_MfcCal.PointData(6).ActualPercent = CSng(62.5)
            New_MfcCal.PointData(7).ActualPercent = CSng(75)
            New_MfcCal.PointData(8).ActualPercent = CSng(87.5)
            New_MfcCal.PointData(9).ActualPercent = CSng(100)
        Case 10
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(11.1)
            New_MfcCal.PointData(3).RawPercent = CSng(22.2)
            New_MfcCal.PointData(4).RawPercent = CSng(33.3)
            New_MfcCal.PointData(5).RawPercent = CSng(44.4)
            New_MfcCal.PointData(6).RawPercent = CSng(55.5)
            New_MfcCal.PointData(7).RawPercent = CSng(66.6)
            New_MfcCal.PointData(8).RawPercent = CSng(77.7)
            New_MfcCal.PointData(9).RawPercent = CSng(88.8)
            New_MfcCal.PointData(10).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(11.1)
            New_MfcCal.PointData(3).ActualPercent = CSng(22.2)
            New_MfcCal.PointData(4).ActualPercent = CSng(33.3)
            New_MfcCal.PointData(5).ActualPercent = CSng(44.4)
            New_MfcCal.PointData(6).ActualPercent = CSng(55.5)
            New_MfcCal.PointData(7).ActualPercent = CSng(66.6)
            New_MfcCal.PointData(8).ActualPercent = CSng(77.7)
            New_MfcCal.PointData(9).ActualPercent = CSng(88.8)
            New_MfcCal.PointData(10).ActualPercent = CSng(100)
        Case 11
            New_MfcCal.PointData(1).RawPercent = CSng(0)
            New_MfcCal.PointData(2).RawPercent = CSng(10)
            New_MfcCal.PointData(3).RawPercent = CSng(20)
            New_MfcCal.PointData(4).RawPercent = CSng(30)
            New_MfcCal.PointData(5).RawPercent = CSng(40)
            New_MfcCal.PointData(6).RawPercent = CSng(50)
            New_MfcCal.PointData(7).RawPercent = CSng(60)
            New_MfcCal.PointData(8).RawPercent = CSng(70)
            New_MfcCal.PointData(9).RawPercent = CSng(80)
            New_MfcCal.PointData(10).RawPercent = CSng(90)
            New_MfcCal.PointData(11).RawPercent = CSng(100)
            New_MfcCal.PointData(1).ActualPercent = CSng(0)
            New_MfcCal.PointData(2).ActualPercent = CSng(10)
            New_MfcCal.PointData(3).ActualPercent = CSng(20)
            New_MfcCal.PointData(4).ActualPercent = CSng(30)
            New_MfcCal.PointData(5).ActualPercent = CSng(40)
            New_MfcCal.PointData(6).ActualPercent = CSng(50)
            New_MfcCal.PointData(7).ActualPercent = CSng(60)
            New_MfcCal.PointData(8).ActualPercent = CSng(70)
            New_MfcCal.PointData(9).ActualPercent = CSng(80)
            New_MfcCal.PointData(10).ActualPercent = CSng(90)
            New_MfcCal.PointData(11).ActualPercent = CSng(100)
    End Select
    
    ' get min/max EU & Raw  for appropriate mfc
    ' Station MFC Calibration Parameters
    Idx = SelCalStn
    ' calc EU & Raw spans
    CalcSpans New_MfcCal.RawInputType
    
    For iPoint = 1 To MAXLSQCALPOINTS
        New_MfcCal.PointData(iPoint).RawValue = sRawMin + (sRawSpan * (New_MfcCal.PointData(iPoint).RawPercent / CSng(100)))
        New_MfcCal.PointData(iPoint).ActualValue = sEuMin + (sEuSpan * (New_MfcCal.PointData(iPoint).ActualPercent / CSng(100)))
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

Private Sub ClearMfcCal(ByRef tmpCal As MfcCalibration)
Dim iPoint As Integer

    ' Set MFC Calibration to Linear (default)
    ' set calibration parameters
    tmpCal.dts = Now
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
    
    ' set MFC Calibration Point Data
    For iPoint = 1 To MAXLSQCALPOINTS
        tmpCal.PointData(iPoint).ActualPercent = CSng(0)
        tmpCal.PointData(iPoint).RawPercent = CSng(0)
        tmpCal.PointData(iPoint).ActualValue = CSng(0)
        tmpCal.PointData(iPoint).RawValue = CSng(0)
    Next iPoint
            
End Sub

Private Sub CalEntries_Click()
Dim Idx As Integer
    Idx = CalEntries.ItemData(CalEntries.ListIndex)
    New_MfcCal.Method = Idx
    Curr_MfcCal.Method = Idx
    Prev_MfcCal.Method = Idx
    UpdateCmdButtons
End Sub

Private Sub cmdAcquireActual_Click()
    If bNewCalEnabled Then
        If ((Com_AIO(acCustCalDevice).addr <> 0) Or (Com_AIO(acCustCalDevice).chan <> 0)) Then
            txtNewActualValue(SelCalPoint).text = Format(Com_AIO(acCustCalDevice).EUValue, "####0.0##")
        End If
    End If
End Sub

Private Sub cmdAcquireRaw_Click()
Dim address As Integer
Dim channel As Integer
Dim sRawMlt As Single
Dim sRawVal As Single
Dim sRawVal2 As Single
    If bNewCalEnabled Then
        Select Case New_MfcCal.RawInputType
            Case CalRawAsVolts  ' 0-5 volts
                address = Stn_AIO(SelCalStn, SelCalFunc).addr
                channel = Stn_AIO(SelCalStn, SelCalFunc).chan
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawMlt = (sRawVal / (FULLSCALE / CSng(2)))
                sRawVal2 = sRawMin + (sRawMlt * sRawSpan)
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
            Case CalRawAsMa     ' 0-20ma (converted from Vdc)
                address = Stn_AIO(SelCalStn, SelCalFunc).addr
                channel = Stn_AIO(SelCalStn, SelCalFunc).chan
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawMlt = (sRawVal / (FULLSCALE / CSng(2)))
                sRawVal2 = sRawMin + (sRawMlt * sRawSpan)
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
            Case CalRawAsDegC   ' Opto TC & RTD modules return temp as ***.0 degC
                address = Stn_AIO(SelCalStn, SelCalFunc).addr
                channel = Stn_AIO(SelCalStn, SelCalFunc).chan
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawVal2 = sRawVal * 0.1
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
            Case CalRawAsEU     ' Raw range = EU range
                address = Stn_AIO(SelCalStn, SelCalFunc).addr
                channel = Stn_AIO(SelCalStn, SelCalFunc).chan
                sRawVal = CSng(Map_AIO(address, channel).RawValue)
                sRawMlt = (sRawVal / (FULLSCALE / CSng(2)))
                sRawVal2 = sEuMin + (sRawMlt * sEuSpan)
                txtNewRawValue(SelCalPoint).text = Format(sRawVal2, "####0.0##")
        End Select
    End If
End Sub

Private Sub cmdApply_Click()
Dim iRawInputType As Integer
Dim Idx As Integer
Dim iPoint As Integer
Dim tmpVal As Single
Dim tmpval2 As Single
Dim flag As Boolean

    ' valid value in NumCalPts text box ??
    flag = RangeCheck(MAXLSQCALPOINTS, MINLSQCALPOINTS, txtNumCalPts, "Number of Calibration Points")
    If flag Then
        
        ' setup screen for different NumberOfCalibrationPoints
        If (CInt(ValueFromText(txtNumCalPts.text)) <> NumCalPoints) Then
            ' Number of Calibration Points for this MFC has been changed
            NumCalPoints = CInt(ValueFromText(txtNumCalPts.text))
            iRawInputType = Curr_MfcCal.RawInputType
            ClearMfcCal New_MfcCal
            ClearMfcCal Curr_MfcCal
            ClearMfcCal Prev_MfcCal
            EnableNewCal False
            EnablePrevCal False
            ' Set MFC Calibration to Linear (default)
            ' get min/max EU & Raw  for appropriate mfc
            ' calc EU & Vdc spans
            CalcSpans iRawInputType
            
            ' set calibration parameters
            Curr_MfcCal.dts = Now()
            Curr_MfcCal.CalibratedBy = "default"
            Curr_MfcCal.Equipment = "equipment"
            Curr_MfcCal.Comment = "linear"
            Curr_MfcCal.NumPoints = NumCalPoints
            Curr_MfcCal.RawInputType = iRawInputType
    '            New_MfcCal.CalData.X = sEuMax
            Curr_MfcCal.CalData.X = CSng(1)
            Curr_MfcCal.CalData.X2 = CSng(0)
            Curr_MfcCal.CalData.X3 = CSng(0)
            Curr_MfcCal.CalData.X4 = CSng(0)
            Curr_MfcCal.CalData.X5 = CSng(0)
            Curr_MfcCal.CalData.X6 = CSng(0)
            Curr_MfcCal.CalData.R2 = CSng(0)
                
            ' set MFC Calibration Point Data
            For iPoint = 1 To MAXLSQCALPOINTS
                Idx = iPoint - 1
                tmpVal = CSng(Idx) * CSng(10)
                tmpval2 = tmpVal / CSng(100)
                Curr_MfcCal.PointData(iPoint).ActualPercent = tmpVal
                Curr_MfcCal.PointData(iPoint).RawPercent = tmpVal
                Curr_MfcCal.PointData(iPoint).ActualValue = sEuMin + (tmpval2 * sEuSpan)
                Curr_MfcCal.PointData(iPoint).RawValue = sRawMin + (tmpval2 * sRawSpan)
            Next iPoint
            New_MfcCal = Curr_MfcCal
        End If
        bEditNumCalPts = False
        DisplayMfcAll
        
    End If
    
End Sub

Private Sub cmdCalCheck_Click()
    ' Open the check calibration form
    frmCalCheck.Show
    frmCalCheck.SetupCalCheck SelCalStn, SelCalMfc
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

Private Sub cmdCancel_Click()
    frmQuestion.Top = OutOfSight
End Sub

Private Sub cmdCreateNewCalib_Click()
    ClearMfcCal New_MfcCal
    New_MfcCal.NumPoints = NumCalPoints
    New_MfcCal.dts = Now()
    New_MfcCal.StandardTempValue = 20
    New_MfcCal.StandardTempUnits = "deg C"
    New_MfcCal.StandardPressValue = 1
    New_MfcCal.StandardPressUnits = "atm"
    New_MfcCal.CalibratedBy = "default"
    New_MfcCal.Equipment = "equipment"
    New_MfcCal.Comment = "linear"
    New_MfcCal.RawInputType = Curr_MfcCal.RawInputType
    EnableNewCal True
    FillNewCalPointTable
    DisplayNewCalInformation
    DisplayCalPointData
    SelectCalPoint 1
    UpdateCmdButtons
End Sub

Private Sub cmdGroupDn_Click()
' This command decrements the station number,
' the displayed name of the station, and triggers an update for
' the values displayed on the form for the current mfc
Dim flag As Boolean
    flag = False
    If (Not CalReadOnly) Then CalValves False
    Do While Not flag
        SelCalStn = SelCalStn - 1
        If SelCalStn < MINSTN Then SelCalStn = MAXSTN
        If SelCalStn <= LAST_STN Then flag = True
    Loop
    ' check for valid mfc in new station
    SelCalMfc = SelCalMfc - 1
    NextInput
    If (Not CalReadOnly) Then CalValves True
    bUnsavedCal = False
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcAll
'    CalValves True
End Sub

Private Sub cmdGroupUp_Click()
' This command increments the station number,
' the displayed name of the station, and triggers an update for
' the values displayed on the form for the current mfc
Dim flag As Boolean
    flag = False
    If (Not CalReadOnly) Then CalValves False
    Do While Not flag
        SelCalStn = SelCalStn + 1
        If SelCalStn > MAXSTN Then SelCalStn = MINSTN
        If SelCalStn <= LAST_STN Then flag = True
    Loop
    ' check for valid mfc in new group
    SelCalMfc = SelCalMfc - 1
    NextInput
    If (Not CalReadOnly) Then CalValves True
    bUnsavedCal = False
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcAll
End Sub

Private Sub cmdHelp_Click()
    frmCalHelp.Show
End Sub

Private Sub cmdInputDn_Click()
    If (Not CalReadOnly) Then CalValves False
    PrevInput
    If (Not CalReadOnly) Then CalValves True
End Sub

Private Sub cmdInputUp_Click()
    If (Not CalReadOnly) Then CalValves False
    NextInput
    If (Not CalReadOnly) Then CalValves True
End Sub

Private Sub PrevInput()
' This command decrements the mfc number,
' the displayed name of the station, and triggers an update for
' the values displayed on the form for the current mfc
Dim iAddr As Integer
Dim iChan As Integer
Dim iCntr As Integer
Dim iMax As Integer
Dim calneeded As Boolean
Dim iFunc As Integer
    

    ' get max mfc index
    ' Station MFC Input Calibration Parameters
    iMax = MAXMFC
            
    iCntr = 0
    calneeded = False
    Do While Not calneeded
        SelCalMfc = SelCalMfc - 1
        If SelCalMfc < 0 Then SelCalMfc = iMax
        ' get the station analog function index for the selected MFC
        Select Case SelCalMfc
            Case MFCBUTANE
                iFunc = asButaneFlow
            Case MFCNITROGEN
                iFunc = asNitrogenFlow
            Case MFCPURGEAIR
                iFunc = asPurgeAirFlow
            Case MFCORVRBUT
                 iFunc = asButaneORVRFlow
            Case MFCORVRNIT
                iFunc = asNitrogenORVRFlow
            Case MFCORVRPRG
                iFunc = asPurgeAirFlow
            Case MFCLIVEFUEL
                iFunc = asLiveFuelVaporFlow
            Case MFCORVRLIVE
                iFunc = asLiveFuelVaporORVRFlow
        End Select
        iAddr = Stn_AIO(SelCalStn, iFunc).addr
        iChan = Stn_AIO(SelCalStn, iFunc).chan
        ' Addr=0 AND Chan=0 is an undefined input
        ' MFC inputs and outputs are not done here
        If ((iAddr <> 0) Or (iChan <> 0)) Then
            Select Case SelCalMfc
                Case MFCBUTANE
                    If STN_INFO(SelCalStn).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCNITROGEN
                    If STN_INFO(SelCalStn).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCPURGEAIR
                    If STN_INFO(SelCalStn).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRBUT
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRNIT
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRPRG
                    ' not used
                Case MFCLIVEFUEL
                    If STN_INFO(SelCalStn).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRLIVE
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
            End Select
        End If
        iCntr = iCntr + 1
        If iCntr > iMax Then calneeded = True
    Loop
    bUnsavedCal = False
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcAll
End Sub

Private Sub NextInput()
' This command increments the mfc number,
' the displayed name of the station, and triggers an update for
' the values displayed on the form for the current mfc
Dim iAddr As Integer
Dim iChan As Integer
Dim iCntr As Integer
Dim iMax As Integer
Dim sDesc As String
Dim iFunc As Integer
Dim calneeded As Boolean

    ' get max mfc index
    ' Station MFC Input Calibration Parameters
    iMax = MAXMFC
            
    iCntr = 0
    calneeded = False
    Do While Not calneeded
        SelCalMfc = SelCalMfc + 1
        If SelCalMfc > iMax Then SelCalMfc = 0
        ' get the station analog function index for the selected MFC
        Select Case SelCalMfc
            Case MFCBUTANE
                iFunc = asButaneFlow
            Case MFCNITROGEN
                iFunc = asNitrogenFlow
            Case MFCPURGEAIR
                iFunc = asPurgeAirFlow
            Case MFCORVRBUT
                 iFunc = asButaneORVRFlow
            Case MFCORVRNIT
                iFunc = asNitrogenORVRFlow
            Case MFCORVRPRG
                iFunc = asPurgeAirFlow
            Case MFCLIVEFUEL
                iFunc = asLiveFuelVaporFlow
            Case MFCORVRLIVE
                iFunc = asLiveFuelVaporORVRFlow
        End Select
        iAddr = Stn_AIO(SelCalStn, iFunc).addr
        iChan = Stn_AIO(SelCalStn, iFunc).chan
        ' Addr=0 AND Chan=0 is an undefined input
        ' MFC inputs and outputs are not done here
        If ((iAddr <> 0) Or (iChan <> 0)) Then
            Select Case SelCalMfc
                Case MFCBUTANE
                    If STN_INFO(SelCalStn).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCNITROGEN
                    If STN_INFO(SelCalStn).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCPURGEAIR
                    If STN_INFO(SelCalStn).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRBUT
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRNIT
                    If STN_INFO(SelCalStn).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRPRG
                    ' not used
                Case MFCLIVEFUEL
                    If STN_INFO(SelCalStn).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRLIVE
                    If STN_INFO(SelCalStn).Type = STN_LIVEORVR2_TYPE Then calneeded = True
            End Select
        End If
        iCntr = iCntr + 1
        If iCntr > iMax Then calneeded = True
    Loop
    bUnsavedCal = False
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcAll
End Sub

Private Sub cmdOK_Click()
    ' copy the current calibration
    SaveCurrCalibration SelCalStn, dstMfc
    frmQuestion.Top = OutOfSight
End Sub

Private Sub cmdPointNum_Click(Index As Integer)
' Selects the row that was clicked on, for editing (but don't repeat up/dn cmd)
    If SelCalPoint <> Index Then
        SelectCalPoint Index
        DesFlowSLPM_Validate SelCalPoint
'        DisplayMfcAll
'        txtNewActualValue(SelCalPoint).Enabled = True
'        If bFormLoaded Then txtNewActualValue(SelCalPoint).SetFocus
    End If
End Sub

Private Sub SelectRow(ByVal RowToSelect As Integer)
    ' Enables the radio button for the current MFC
    Select Case New_MfcCal.Method
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
    cmdPointNum(RowToSelect).Appearance = 1    ' 3D
    cmdPointNum(RowToSelect).FontBold = True
    cmdPointNum(RowToSelect).BackColor = PALEBLUE
End Sub

Private Sub DisableRow(ByVal RowToDisable As Integer)
    ' Unselects every element in this row
    txtNewRawValue(RowToDisable).Enabled = False
    txtNewActualValue(RowToDisable).Enabled = False
    cmdPointNum(RowToDisable).Appearance = 0    ' Flat
    cmdPointNum(RowToDisable).FontBold = False
    cmdPointNum(RowToDisable).BackColor = Common_BackColor
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

Private Sub DesFlowSLPM_Validate(ByVal Index As Integer)
'
Dim sRawMlt As Single
Dim sEuSP As Single
    If Not IsNumeric(txtNewRawValue(Index).text) Then txtNewRawValue(Index).text = "0.0"
    If (CSng(txtNewRawValue(Index).text) > (CSng(1.25) * sRawMax)) Then txtNewRawValue(Index).text = CStr(CSng(1.25) * sRawMax)
    ' Output the value to the Mass Flow Controller
    sRawMlt = (CSng(txtNewRawValue(Index).text) - sRawMin) / sRawSpan
    sEuSP = sEuMin + (sEuSpan * sRawMlt)
    SendFlow sEuSP
End Sub

Private Sub CalValves(ByVal outFlag As Boolean)
' Outputs the valves for Calibration of
' the selected Mass Flow Controller
Dim outRequest As Integer

    outRequest = IIf(outFlag, cON, cOFF)
    PRG_INFO(STN_INFO(SelCalStn).AspiratorNum).RequestRun = outFlag
    
    Select Case SelCalMfc
        Case MFCBUTANE
            Stn_OutDigital SelCalStn, isButaneSol, outRequest
            Stn_OutDigital SelCalStn, isAuxDirectionSol, outRequest
            
        Case MFCNITROGEN
            Stn_OutDigital SelCalStn, isNitrogenSol, outRequest
            Stn_OutDigital SelCalStn, isAuxDirectionSol, outRequest
         
         Case MFCPURGEAIR
            PRG_INFO(STN_INFO(SelCalStn).AspiratorNum).RequestRdy = outFlag
            Stn_OutDigital SelCalStn, isPurgeSol, outRequest
            Stn_OutDigital SelCalStn, isPriDirectionSol, outRequest
             
        Case MFCORVRBUT
            Stn_OutDigital SelCalStn, isButaneOrvrSol, outRequest
            Stn_OutDigital SelCalStn, isAuxDirectionSol, outRequest
         
         Case MFCORVRNIT
            Stn_OutDigital SelCalStn, isNitrogenOrvrSol, outRequest
            Stn_OutDigital SelCalStn, isAuxDirectionSol, outRequest
             
        Case MFCLIVEFUEL
            Stn_OutDigital SelCalStn, isLiveFuelSol, outRequest
            Stn_OutDigital SelCalStn, isAuxDirectionSol, outRequest
                 
        Case MFCORVRLIVE
            Stn_OutDigital SelCalStn, isLiveFuelOrvrSol, outRequest
            Stn_OutDigital SelCalStn, isAuxDirectionSol, outRequest
                 
    End Select

End Sub

Private Sub SendFlow(ByVal OutputSLPM As Single)
' Outputs the value specified in OutputSLPM to
' the selected Mass Flow Controller

Dim OutputEng, span As Single
    
'    aryOutputFS(SelectedRow) = frmMassFlowCal.SolveCalibFor(OutputSLPM)

    Select Case SelCalMfc
        Case MFCBUTANE
            span = Stn_AIO(SelCalStn, asButaneFlowSP).EuMax - Stn_AIO(SelCalStn, asButaneFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asButaneFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCBUTANE, Stn_MfcCal(SelCalStn, MFCBUTANE)))
'            Stn_OutAnalog SelCalStn, asButaneFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asButaneFlowSP, OutputSLPM, outNORMAL
        
        Case MFCNITROGEN
            span = Stn_AIO(SelCalStn, asNitrogenFlowSP).EuMax - Stn_AIO(SelCalStn, asNitrogenFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCNITROGEN, Stn_MfcCal(SelCalStn, MFCNITROGEN)))
'            Stn_OutAnalog SelCalStn, asNitrogenFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asNitrogenFlowSP, OutputSLPM, outNORMAL
            
        Case MFCPURGEAIR
            span = Stn_AIO(SelCalStn, asPurgeAirFlowSP).EuMax - Stn_AIO(SelCalStn, asPurgeAirFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asPurgeAirFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCPURGEAIR, Stn_MfcCal(SelCalStn, MFCPURGEAIR)))
'            Stn_OutAnalog SelCalStn, asPurgeAirFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asPurgeAirFlowSP, OutputSLPM, outNORMAL
            
        Case MFCORVRBUT
            span = Stn_AIO(SelCalStn, asButaneORVRFlowSP).EuMax - Stn_AIO(SelCalStn, asButaneORVRFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asButaneORVRFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCORVRBUT, Stn_MfcCal(SelCalStn, MFCORVRBUT)))
'            Stn_OutAnalog SelCalStn, asButaneORVRFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asButaneORVRFlowSP, OutputSLPM, outNORMAL
    
        Case MFCORVRNIT
            span = Stn_AIO(SelCalStn, asNitrogenORVRFlowSP).EuMax - Stn_AIO(SelCalStn, asNitrogenORVRFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asNitrogenORVRFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCORVRNIT, Stn_MfcCal(SelCalStn, MFCORVRNIT)))
'            Stn_OutAnalog SelCalStn, asNitrogenORVRFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asNitrogenORVRFlowSP, OutputSLPM, outNORMAL
            
        Case MFCLIVEFUEL
            span = Stn_AIO(SelCalStn, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(SelCalStn, asLiveFuelVaporFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCLIVEFUEL, Stn_MfcCal(SelCalStn, MFCLIVEFUEL)))
'            Stn_OutAnalog SelCalStn, asLiveFuelVaporFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asLiveFuelVaporFlowSP, OutputSLPM, outNORMAL
                
        Case MFCORVRLIVE
            span = Stn_AIO(SelCalStn, asLiveFuelVaporORVRFlowSP).EuMax - Stn_AIO(SelCalStn, asLiveFuelVaporORVRFlowSP).EuMin
            OutputEng = Stn_AIO(SelCalStn, asLiveFuelVaporORVRFlowSP).EuMin + (span * Cal_MfcOutput(OutputSLPM, SelCalStn, MFCORVRLIVE, Stn_MfcCal(SelCalStn, MFCORVRLIVE)))
'            Stn_OutAnalog SelCalStn, asLiveFuelVaporORVRFlowSP, CSng(OutputEng), outNORMAL
            Stn_OutAnalog SelCalStn, asLiveFuelVaporORVRFlowSP, OutputSLPM, outNORMAL
                
    End Select

End Sub

Private Sub cmdRunCal_Click()
   
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 7781, 3
Dim iPoint As Integer
Dim tmpPercent As Single
Dim tmpPercent2 As Single
Dim ErrorExists As Boolean

    ' Validate the new actual value boxes
    ErrorExists = False
    ' check all boxes
    For iPoint = 1 To NumCalPoints
        If (Not RangeCheck((1.25 * sRawMax), sRawMin, txtNewRawValue(iPoint), "Raw Value Entry")) Then ErrorExists = True
        If (Not RangeCheck((1.25 * sEuMax), sEuMin, txtNewActualValue(iPoint), "Actual Value Entry")) Then ErrorExists = True
    Next iPoint
    ' all boxes valid ??
    If Not ErrorExists Then
        ' all entries are valid; calibrate
        ' update new calibration DTS
        txtCalibDts.text = Format(Now, "YYYY-MMM-DD  hh:mm:ss")
        ' update new calibration from screen
        New_MfcCal.NumPoints = CInt(txtNumCalPts.text)
        New_MfcCal.dts = CDate(txtCalibDts.text)
        New_MfcCal.StandardTempValue = ValueFromText(txtStandardTemp.text)
        New_MfcCal.StandardTempUnits = TempUnits.List(TempUnits.ListIndex)
        New_MfcCal.StandardPressValue = ValueFromText(txtStandardPress.text)
        New_MfcCal.StandardPressUnits = PressUnits.List(PressUnits.ListIndex)
        New_MfcCal.CalibratedBy = txtCalibBy.text
        New_MfcCal.Equipment = txtEquipment.text
        New_MfcCal.Comment = txtComment.text
        ' update new calibration point data from screen
        For iPoint = 1 To NumCalPoints
            New_MfcCal.PointData(iPoint).RawValue = ValueFromText(txtNewRawValue(iPoint).text)
            New_MfcCal.PointData(iPoint).ActualValue = ValueFromText(txtNewActualValue(iPoint).text)
            tmpPercent = CSng(100) * ((New_MfcCal.PointData(iPoint).RawValue - sRawMin) / sRawSpan)
            New_MfcCal.PointData(iPoint).RawPercent = tmpPercent
            tmpPercent2 = CSng(100) * ((New_MfcCal.PointData(iPoint).ActualValue - sEuMin) / sEuSpan)
            New_MfcCal.PointData(iPoint).ActualPercent = tmpPercent2
        Next iPoint
    
        ' Copy Current to Prev
        If (Not bPrevCalEnabled) Then Prev_MfcCal = Curr_MfcCal
        ' Calculate New Calibration Coefficients
        Calibrate
        ' Copy New to Current
        Curr_MfcCal = New_MfcCal
            
'       ClearMfcCal New_MfcCal
        ' update the controls
        EnableNewCal True
        EnablePrevCal True
        bUnsavedCal = True
        blnPrevCalExists = True
        DisplayMfcAll
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "New Calibration Done"
        
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
        xlSht.Range("A" & i + 1) = (New_MfcCal.PointData(i).RawPercent / CSng(100))
        xlSht.Range("B" & i + 1) = (New_MfcCal.PointData(i).RawPercent / CSng(100)) ^ 2
        xlSht.Range("C" & i + 1) = (New_MfcCal.PointData(i).RawPercent / CSng(100)) ^ 3
        xlSht.Range("D" & i + 1) = (New_MfcCal.PointData(i).RawPercent / CSng(100)) ^ 4
        xlSht.Range("E" & i + 1) = (New_MfcCal.PointData(i).RawPercent / CSng(100)) ^ 5
        xlSht.Range("F" & i + 1) = (New_MfcCal.PointData(i).RawPercent / CSng(100)) ^ 6
        xlSht.Range("G" & i + 1) = New_MfcCal.PointData(i).ActualValue
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
    New_MfcCal.CalData.X = xlSht.Range("F14").Value
    New_MfcCal.CalData.X2 = xlSht.Range("E14").Value
    New_MfcCal.CalData.X3 = xlSht.Range("D14").Value
    New_MfcCal.CalData.X4 = xlSht.Range("C14").Value
    New_MfcCal.CalData.X5 = xlSht.Range("B14").Value
    New_MfcCal.CalData.X6 = xlSht.Range("A14").Value
    
    New_MfcCal.CalData.R2 = xlSht.Range("A16").Value

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
    FormulaText = FormulaText & "Y = " & Curr_MfcCal.CalData.X6 & "X6"
    FormulaText = FormulaText & IIf(Curr_MfcCal.CalData.X5 < 0, " - ", " + ") & Abs(Curr_MfcCal.CalData.X5) & "X5"
    FormulaText = FormulaText & IIf(Curr_MfcCal.CalData.X4 < 0, " - ", " + ") & Abs(Curr_MfcCal.CalData.X4) & "X4"
    FormulaText = FormulaText & IIf(Curr_MfcCal.CalData.X3 < 0, " - ", " + ") & Abs(Curr_MfcCal.CalData.X3) & "X3"
    FormulaText = FormulaText & IIf(Curr_MfcCal.CalData.X2 < 0, " - ", " + ") & Abs(Curr_MfcCal.CalData.X2) & "X2"
    FormulaText = FormulaText & IIf(Curr_MfcCal.CalData.X < 0, " - ", " + ") & Abs(Curr_MfcCal.CalData.X) & "X"
    FormulaText = FormulaText & vbCrLf
    FormulaText = FormulaText & "      R2=" & Curr_MfcCal.CalData.R2
    lblCalFormula.Visible = True
    lblCalFormula.ForeColor = White
    lblCalFormula.Caption = FormulaText
End Sub

Private Sub DisplayCalGraph()
Dim ChartArray(MAXLSQCALPOINTS, 2)
Dim Idx As Integer
Dim iPoint As Integer
Dim tmpVal As Single
Dim Graph() As Single
         
    ReDim Graph(NumCalPoints, 1 To 6)
    
    ' Station Analog Input Calibration Parameters
    For iPoint = 1 To NumCalPoints
        Graph(iPoint, 1) = Curr_MfcCal.PointData(iPoint).RawPercent                             ' value for X-axis - Pen #1
        Graph(iPoint, 2) = CSng(Curr_MfcCal.PointData(iPoint).RawPercent)                       ' value for Y-axis - Pen #1
        Graph(iPoint, 3) = Curr_MfcCal.PointData(iPoint).RawPercent                             ' value for X-axis - Pen #2
        Graph(iPoint, 4) = CSng(Curr_MfcCal.PointData(iPoint).ActualPercent)                    ' value for Y-axis - Pen #2
        Graph(iPoint, 5) = Curr_MfcCal.PointData(iPoint).RawPercent                             ' value for X-axis - Pen #3
        tmpVal = Cal_MfcInput(Curr_MfcCal.PointData(iPoint).RawPercent / 100, SelCalStn, SelCalMfc, Curr_MfcCal)
        If (sEuSpan <> 0) Then Graph(iPoint, 6) = CSng(100) * ((tmpVal - sEuMin) / sEuSpan)     ' value for Y-axis - Pen #3
    Next iPoint
         
    chtMfcChart.chartType = VtChChartType2dXY  ' set to X Y Scatter chart
    chtMfcChart = Graph ' populate chart's data grid using Graph array
    chtMfcChart.Plot.UniformAxis = False
    chtMfcChart.Column = 1
    chtMfcChart.ColumnLabel = "Raw Value"
    chtMfcChart.Column = 3
    chtMfcChart.ColumnLabel = "Actual Value"
    chtMfcChart.Column = 5
    chtMfcChart.ColumnLabel = "Calib. Value"
    chtMfcChart.Visible = True
End Sub

Private Sub cmdRestorePrevCal_Click()
' Undoes the effect of the calibration by transferring data from
' the previous calibration to the current calibration, transferring data
' from the temporary calibration buffer to the previous calibration,
' and updating the form
    
    ' shift cal point data from prev to current
'    ClearMfcCal New_MfcCal
    New_MfcCal = Curr_MfcCal
'    ClearMfcCal Curr_MfcCal
    Curr_MfcCal = Prev_MfcCal
    ClearMfcCal Prev_MfcCal
    EnablePrevCal False
    bUnsavedCal = False
    ' update display
    DisplayMfcCalibration
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
        cmdAcquireRaw.Visible = IIf(bNewCalEnabled And ((New_MfcCal.Method = calmetRawOnly) Or (New_MfcCal.Method = calmetRawAndActual)), True, False)
        cmdAcquireActual.Visible = IIf(bNewCalEnabled And ((New_MfcCal.Method = calmetActualOnly) Or (New_MfcCal.Method = calmetRawAndActual)), True, False)
    End If
End Sub

Private Sub UpdateMfcSelection()
'
    Select Case SelCalMfc
        Case MFCBUTANE
            SelCalFunc = asButaneFlow
        Case MFCNITROGEN
            SelCalFunc = asNitrogenFlow
        Case MFCPURGEAIR
            SelCalFunc = asPurgeAirFlow
        Case MFCORVRBUT
            SelCalFunc = asButaneORVRFlow
        Case MFCORVRNIT
            SelCalFunc = asNitrogenORVRFlow
        Case MFCORVRPRG
            SelCalFunc = asPurgeAirFlow
        Case MFCLIVEFUEL
            SelCalFunc = asLiveFuelVaporFlow
        Case MFCORVRLIVE
            SelCalFunc = asLiveFuelVaporORVRFlow
    End Select
    ClearMfcCal New_MfcCal
    ClearMfcCal Curr_MfcCal
    ClearMfcCal Prev_MfcCal
    Curr_MfcCal = Stn_MfcCal(SelCalStn, SelCalMfc)
    NumCalPoints = Curr_MfcCal.NumPoints
    If (Curr_MfcCal.RawInputType = CalRawUndefined) Then Curr_MfcCal.RawInputType = CalRawAsVolts
    CalcSpans Curr_MfcCal.RawInputType
    ' init cal point columns
    EnableNewCal False
    EnablePrevCal False
End Sub

Private Sub DisplayMfcAll()
    ' update the screen
    DisplayMfcSelection useExistRawInputType
    HideInactiveTableRows
    DisplayMfcCalibration
    UpdateCmdButtons
End Sub

Private Sub DisplayMfcCalibration()
'
' Cal is ReadOnly unless all stations are idle
'
    If AllStationsIdle And CalReadOnly Then
        CalReadOnly = False
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "MFC Calibration is Enabled"
    ElseIf Not AllStationsIdle And Not CalReadOnly Then
        CalReadOnly = True
        txtMsg.ForeColor = DKPURPLE
        txtMsg.text = vbCrLf & "MFC Calibration is Read-Only"
    End If
    ' update the screen
    DisplayCurrCalInformation
    DisplayCalPointData
    DisplayCalFormula
    DisplayCalGraph
    
End Sub

Private Sub cmdSaveCurrCal_Click()
    ' Save the current calibration
    SaveCurrCalibration SelCalStn, SelCalMfc
    CopyCal SelCalStn, SelCalMfc
    bUnsavedCal = False
    ' clear new & prev calibrations
    ClearMfcCal New_MfcCal
    ClearMfcCal Prev_MfcCal
    EnableNewCal False
    EnablePrevCal False
    blnPrevCalExists = False
    ' update screen
    DisplayMfcCalibration
    ' update command buttons
    UpdateCmdButtons
'    Delay_Box "Calibration Saved", MSGDELAY, msgSHOW
    txtMsg.ForeColor = DKPURPLE
    txtMsg.text = vbCrLf & "Calibration Saved"
End Sub

Private Sub Form_Load()
'
Dim Idx As Integer
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
    For Idx = 1 To MAXLSQCALPOINTS
        txtNewRawValue(Idx).ForeColor = TitlesData_Forecolor
        txtNewActualValue(Idx).ForeColor = TitlesData_Forecolor
    Next Idx

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

    ' Set the current station, mfc and point(row)
    SelCalStn = 1
    SelCalMfc = MAXMFC
    SelCalPoint = 1
    ' find the first valid mfc
    NextInput
    ' init cal point columns
    EnableNewCal False
    EnablePrevCal False
    ' init Unsaved Calibration flag
    bUnsavedCal = False
    blnPrevCalExists = False
    ' hide the "Question" frame
    frmQuestion.Top = OutOfSight
    ' update the screen
    DisplayMfcAll
    
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
        cmdPointNum(iRow).Visible = flag
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

Private Sub DisplayMfcSelection(ByVal updateRawInputType As Integer)
' DisplayMfcSelection
' Displays Information on the Currently Selected Mass Flow Controller
Dim sGrpDesc As String
Dim sAiDesc As String
Dim iRawInputType As Integer
    
    sGrpDesc = "Station #" & Format(SelCalStn, "#0")
    sAiDesc = Stn_AnaDef(SelCalFunc).desc
    iRawInputType = Stn_MfcCal(SelCalStn, SelCalMfc).RawInputType
    If (updateRawInputType = useNewRawInputType) Then iRawInputType = New_MfcCal.RawInputType
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

Private Sub RawValues_Click()
Dim Idx As Integer
    Idx = RawValues.ItemData(RawValues.ListIndex)
    New_MfcCal.RawInputType = Idx
    Curr_MfcCal.RawInputType = Idx
    Prev_MfcCal.RawInputType = Idx
    CalcSpans Idx
    DisplayMfcSelection useNewRawInputType
    DisplayMfcCalibration
End Sub

Private Sub tmrUpdate_Timer()
    If (Not CalReadOnly) Then
        ' Purge Air Piab
        If ((SelCalMfc = MFCPURGEAIR) Or (SelCalMfc = MFCORVRPRG)) Then
            PRG_INFO(STN_INFO(SelCalStn).AspiratorNum).RequestRun = True
        End If
    End If
    If USINGSIMULATION Then
        Dim sVal As Single
        Dim sVal2 As Single
        Dim sValMax As Single
        Dim sValMin As Single
        Dim sValSpan As Single
        
        sVal = (Stn_AIO(SelCalStn, (SelCalFunc - 10)).EUValue - sEuMin) / sEuSpan
        sVal = IIf(USINGSIMNOISE, (sVal - (0.02 * (0.5 + Rnd))), sVal)
        
        If (Map_AIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).VdcMax = Map_AIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).VdcMin) Then
            sVal2 = 0
        Else
            sValMax = (Map_AIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).VdcMax / CSng(5)) * CSng(FULLSCALE / CLng(2))
            sValMin = (Map_AIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).VdcMin / CSng(5)) * CSng(FULLSCALE / CLng(2))
            sValSpan = sValMax - sValMin
            sVal2 = sValMin + (sValSpan * sVal)
        End If
        Map_AIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).RawValue = sVal2
        OptoAIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).RawValue = sVal2
        Map_AIO(Stn_AIO(SelCalStn, SelCalFunc).addr, Stn_AIO(SelCalStn, SelCalFunc).chan).EUValue = Stn_AIO(SelCalStn, (SelCalFunc - 10)).EUValue
'        Stn_AIO(SelCalStn, SelCalFunc).EUValue = Stn_AIO(SelCalStn, (SelCalFunc - 10)).EUValue
        
        If (Map_AIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).VdcMax = Map_AIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).VdcMin) Then
            sVal2 = 0
        Else
            sValMax = (Map_AIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).VdcMax / CSng(5)) * CSng(FULLSCALE / CLng(2))
            sValMin = (Map_AIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).VdcMin / CSng(5)) * CSng(FULLSCALE / CLng(2))
            sValSpan = sValMax - sValMin
            sVal2 = sValMin + (sValSpan * sVal)
        End If
        Map_AIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).RawValue = sVal2
        OptoAIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).RawValue = sVal2
        Map_AIO(Com_AIO(acCustCalDevice).addr, Com_AIO(acCustCalDevice).chan).EUValue = Stn_AIO(SelCalStn, (SelCalFunc - 10)).EUValue
        Com_AIO(acCustCalDevice).EUValue = Stn_AIO(SelCalStn, (SelCalFunc - 10)).EUValue
        
    End If
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
    Select Case iRawType
        Case CalRawAsVolts  ' 0-5 volts
            sRawMax = Stn_AIO(SelCalStn, SelCalFunc).VdcMax
            sEuMax = Stn_AIO(SelCalStn, SelCalFunc).EuMax
            sRawMin = Stn_AIO(SelCalStn, SelCalFunc).VdcMin
            sEuMin = Stn_AIO(SelCalStn, SelCalFunc).EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
        Case CalRawAsMa     ' 0-20ma (converted from Vdc)
            sRawMax = CSng(4) * Stn_AIO(SelCalStn, SelCalFunc).VdcMax
            sEuMax = Stn_AIO(SelCalStn, SelCalFunc).EuMax
            sRawMin = CSng(4) * Stn_AIO(SelCalStn, SelCalFunc).VdcMin
            sEuMin = Stn_AIO(SelCalStn, SelCalFunc).EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
        Case CalRawAsDegC   ' Opto TC & RTD modules return temp as ***.0 degC
            sRawMax = Stn_AIO(SelCalStn, SelCalFunc).EuMax
            sEuMax = Stn_AIO(SelCalStn, SelCalFunc).EuMax
            sRawMin = Stn_AIO(SelCalStn, SelCalFunc).EuMin
            sEuMin = Stn_AIO(SelCalStn, SelCalFunc).EuMin
            sRawSpan = sRawMax - sRawMin
            sEuSpan = sEuMax - sEuMin
        Case CalRawAsEU     ' Raw range = EU range
            sRawMax = Stn_AIO(SelCalStn, SelCalFunc).EuMax
            sEuMax = Stn_AIO(SelCalStn, SelCalFunc).EuMax
            sRawMin = Stn_AIO(SelCalStn, SelCalFunc).EuMin
            sEuMin = Stn_AIO(SelCalStn, SelCalFunc).EuMin
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
        ActualValue = Curr_MfcCal.PointData(iPoint).ActualValue
        ' Calibrated Value
        CalibValue = Cal_MfcInput((Curr_MfcCal.PointData(iPoint).ActualPercent / CSng(100)), SelCalStn, SelCalMfc, Curr_MfcCal)
        lblCurrCalValue(iPoint).Caption = IIf(lblCurrCalValue(iPoint).Enabled, Format(CalibValue, "####0.0##"), "")
        ' Percent Difference
        If ActualValue > 0! Then
            percDiff = ((CalibValue - ActualValue) / ActualValue) * 100
        Else
            percDiff = 0!
        End If
        lblCurrDiff(iPoint).Caption = IIf(lblCurrDiff(iPoint).Enabled, Format(percDiff, "####0.0##"), "")
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
        ActualValue = Prev_MfcCal.PointData(iPoint).ActualValue
        ' Calibrated Value
        CalibValue = Cal_MfcInput((Prev_MfcCal.PointData(iPoint).ActualPercent / CSng(100)), SelCalStn, SelCalMfc, Prev_MfcCal)
        lblPrevCalValue(iPoint).Caption = IIf(lblPrevCalValue(iPoint).Enabled, Format(CalibValue, "####0.0##"), "")
        ' Percent Difference
        If ActualValue > 0! Then
            percDiff = ((CalibValue - ActualValue) / ActualValue) * 100
        Else
            percDiff = 0!
        End If
        lblPrevDiff(iPoint).Caption = IIf(lblPrevDiff(iPoint).Enabled, Format(percDiff, "####0.0##"), "")
    Next iPoint
End Sub

Public Sub SaveCurrCalibration(ByVal iStation As Integer, ByVal iMfc As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date
Dim iPoint As Integer
Dim Idx As Integer

    ' delete existing calibration
    ClearMfcCalRecords iStation, iMfc

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Save MFC Calibration Parameters
    Criteria = "SELECT * FROM [MfcCalibrations] WHERE [Station] = " & iStation & " AND [Mfc] = " & iMfc & "  ORDER BY [Dts] DESC"
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
    rsRecord("Station") = iStation
    rsRecord("Mfc") = iMfc
    rsRecord("Dts") = Curr_MfcCal.dts
    rsRecord("CalibratedBy") = Curr_MfcCal.CalibratedBy
    rsRecord("Equipment") = Curr_MfcCal.Equipment
    rsRecord("Comment") = Curr_MfcCal.Comment
    rsRecord("NumPoints") = Curr_MfcCal.NumPoints
    rsRecord("RawInputType") = Curr_MfcCal.RawInputType
    rsRecord("CoefficientX") = Curr_MfcCal.CalData.X
    rsRecord("CoefficientX2") = Curr_MfcCal.CalData.X2
    rsRecord("CoefficientX3") = Curr_MfcCal.CalData.X3
    rsRecord("CoefficientX4") = Curr_MfcCal.CalData.X4
    rsRecord("CoefficientX5") = Curr_MfcCal.CalData.X5
    rsRecord("CoefficientX6") = Curr_MfcCal.CalData.X6
    rsRecord("CoefficientR2") = Curr_MfcCal.CalData.R2
    rsRecord("StandardTempValue") = Curr_MfcCal.StandardTempValue
    rsRecord("StandardTempUnits") = Curr_MfcCal.StandardTempUnits
    rsRecord("StandardPressValue") = Curr_MfcCal.StandardPressValue
    rsRecord("StandardPressUnits") = Curr_MfcCal.StandardPressUnits
    rsRecord.Update
    rsRecord.Close

            
    ' Save MFC Calibration Point Data
    dDts = Curr_MfcCal.dts
    CriteriaPts = "SELECT * FROM [MfcCalibrationsData] WHERE [Station] = " & iStation & " AND [Mfc] = " & iMfc & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
    Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
            
    ' update the calibration information
    For iPoint = 1 To NumCalPoints
        rsRecordPts.AddNew
        rsRecordPts("Station") = iStation
        rsRecordPts("Mfc") = iMfc
        rsRecordPts("Point") = iPoint
        rsRecordPts("Dts") = dDts
        rsRecordPts("ActualPercent") = Curr_MfcCal.PointData(iPoint).ActualPercent
        rsRecordPts("RawPercent") = Curr_MfcCal.PointData(iPoint).RawPercent
        rsRecordPts("ActualValue") = Curr_MfcCal.PointData(iPoint).ActualValue
        rsRecordPts("RawValue") = Curr_MfcCal.PointData(iPoint).RawValue
        rsRecordPts.Update
    Next iPoint
    ' done with points
    rsRecordPts.Close
                        
    ' copy calibration to appropriate station array
    ' Station Analog Input Calibration Parameters
    PrevStn_MfcCal(iStation, iMfc) = Stn_MfcCal(iStation, iMfc)
    Stn_MfcCal(iStation, iMfc) = Curr_MfcCal
                
                
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
Dim Idx As Integer
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
Dim CurrCal As MfcCalibration
Dim PrevCal As MfcCalibration
Dim bPrintPrevCal As Boolean
Dim oldFont As New StdFont

    ' current & previous calibration
    Idx = SelCalStn
    CurrCal = Stn_MfcCal(Idx, SelCalMfc)
    PrevCal = PrevStn_MfcCal(Idx, SelCalMfc)
    
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
    Idx = SelCalStn
    Print_Center "Calibration Report for Station # " & Format(Idx, "#0") & ", " & Mfc_FunDef(SelCalMfc).desc
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
    tmpString(1) = tmpString(1) & Format(CurrCal.dts, "YYYY-MMM-DD")
    ' Calibrated By
    tmpString(2) = tmpString(2) & CurrCal.CalibratedBy
    ' Calibrated By
    tmpString(3) = tmpString(3) & CurrCal.Equipment
    ' Standard Pressure
    tmpString(4) = tmpString(4) & Format(CurrCal.StandardPressValue, "###0.0##") & " " & CurrCal.StandardPressUnits
    ' Standard Temperature
    tmpString(5) = tmpString(5) & Format(CurrCal.StandardTempValue, "##0.0##") & " " & CurrCal.StandardTempUnits

    If bPrintPrevCal Then
        ' right pad tmpString's
        For row = 0 To 6
            tmpString(row) = tmpString(row) & Space(dataRight - Len(tmpString(row)))
        Next row
        ' PREVIOUS CALIBRATION
        tmpString(0) = tmpString(0) & "PREVIOUS CALIBRATION"
        ' Calibration DateTime
        tmpString(1) = tmpString(1) & Format(PrevCal.dts, "YYYY-MMM-DD")
        ' Calibrated By
        tmpString(2) = tmpString(2) & PrevCal.CalibratedBy
        ' Calibrated By
        tmpString(3) = tmpString(3) & PrevCal.Equipment
        ' Standard Pressure
        tmpString(4) = tmpString(4) & Format(PrevCal.StandardPressValue, "###0.0##") & " " & PrevCal.StandardPressUnits
        ' Standard Temperature
        tmpString(5) = tmpString(5) & Format(PrevCal.StandardTempValue, "##0.0##") & " " & PrevCal.StandardTempUnits
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
            sngCalibValue = Cal_MfcInput((CurrCal.PointData(row).ActualPercent / 100), SelCalStn, SelCalMfc, CurrCal)
            If (sngActualValue > CSng(0)) Then
                sngPercDiff = ((sngCalibValue - sngActualValue) / sngActualValue) * 100
            Else
                sngPercDiff = CSng(0)
            End If
            tmpString(row) = tmpString(row) & Space(dataLeft - Len(tmpString(row)))
            tmpString(row) = tmpString(row) & Format(sngActualValue, "#,###,##0.0##")
            tmpString(row) = tmpString(row) & Space((dataLeft + dataWidth) - Len(tmpString(row)))
            tmpString(row) = tmpString(row) & Format(sngCalibValue, "#,###,##0.0##")
            tmpString(row) = tmpString(row) & Space((dataLeft + dataWidth + dataWidth) - Len(tmpString(row)))
            tmpString(row) = tmpString(row) & Format(sngPercDiff, "0000.0")
        End If
        ' previous calibration
        If bPrintPrevCal Then
            If row <= intNumPrevRows Then
                sngActualValue = PrevCal.PointData(row).ActualValue
                sngCalibValue = Cal_MfcInput((PrevCal.PointData(row).ActualPercent / 100), SelCalStn, SelCalMfc, PrevCal)
                If (sngActualValue > CSng(0)) Then
                    sngPercDiff = ((sngCalibValue - sngActualValue) / sngActualValue) * 100
                Else
                    sngPercDiff = CSng(0)
                End If
                tmpString(row) = tmpString(row) & Space(dataRight - Len(tmpString(row)))
                tmpString(row) = tmpString(row) & Format(sngActualValue, "#,###,##0.0##")
                tmpString(row) = tmpString(row) & Space((dataRight + dataWidth) - Len(tmpString(row)))
                tmpString(row) = tmpString(row) & Format(sngCalibValue, "#,###,##0.0##")
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
    NumCalPoints = Curr_MfcCal.NumPoints
    txtNumCalPts.text = Format(NumCalPoints, "#0")
    txtNumCalPts.ForeColor = Black
    txtNumCalPts.BackColor = frmCalInformation.BackColor
    txtCalibDts.text = Format(Curr_MfcCal.dts, "YYYY-MMM-DD hh:mm:ss")
    txtStandardTemp.text = Format(Curr_MfcCal.StandardTempValue, "###0.0")
    TempUnits.ListIndex = TempUnitsIndex(Curr_MfcCal.StandardTempUnits)
    txtStandardPress.text = Format(Curr_MfcCal.StandardPressValue, "###0.0##")
    PressUnits.ListIndex = PressUnitsIndex(Curr_MfcCal.StandardPressUnits)
    txtCalibBy.text = Curr_MfcCal.CalibratedBy
'    txtEquipment.text = "equipment" & Curr_MfcCal.Equipment
    txtEquipment.text = Curr_MfcCal.Equipment
    txtComment.text = Curr_MfcCal.Comment
    TempUnits.Refresh
    PressUnits.Refresh
End Sub

Private Sub DisplayNewCalInformation()
' Displays new calibration information
    NumCalPoints = New_MfcCal.NumPoints
    txtNumCalPts.text = Format(NumCalPoints, "#0")
    txtNumCalPts.ForeColor = Black
    txtNumCalPts.BackColor = frmCalInformation.BackColor
    txtCalibDts.text = Format(New_MfcCal.dts, "YYYY-MMM-DD hh:mm:ss")
    txtStandardTemp.text = Format(New_MfcCal.StandardTempValue, "###0.0")
    TempUnits.ListIndex = TempUnitsIndex(New_MfcCal.StandardTempUnits)
    txtStandardPress.text = Format(New_MfcCal.StandardPressValue, "###0.0##")
    PressUnits.ListIndex = PressUnitsIndex(New_MfcCal.StandardPressUnits)
    txtCalibBy.text = New_MfcCal.CalibratedBy
'    txtEquipment.text = "equipment"
    txtEquipment.text = New_MfcCal.Equipment
    txtComment.text = New_MfcCal.Comment
    TempUnits.Refresh
    PressUnits.Refresh
End Sub

Private Sub DisplayCalPointData()
' Displays new, current & previous calibration point data
    Dim iPoint As Integer
    HideInactiveTableRows
    For iPoint = 1 To NumCalPoints
        cmdPointNum(iPoint).Caption = Format(iPoint, "#0")
        txtNewRawValue(iPoint).text = IIf(txtNewRawValue(iPoint).Enabled, Format(New_MfcCal.PointData(iPoint).RawValue, "####0.0##"), "")
        lblNewRawPerc(iPoint).Caption = IIf(lblNewRawPerc(iPoint).Enabled, Format(New_MfcCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        txtNewActualValue(iPoint).text = IIf(txtNewActualValue(iPoint).Enabled, Format(New_MfcCal.PointData(iPoint).ActualValue, "####0.0##"), "")
        lblCurrRawPerc(iPoint).Caption = IIf(lblCurrRawPerc(iPoint).Enabled, Format(Curr_MfcCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        lblCurrActualValue(iPoint).Caption = IIf(lblCurrActualValue(iPoint).Enabled, Format(Curr_MfcCal.PointData(iPoint).ActualValue, "####0.0##"), "")
        lblPrevRawPerc(iPoint).Caption = IIf(lblPrevRawPerc(iPoint).Enabled, Format(Prev_MfcCal.PointData(iPoint).RawPercent, "####0.0##"), "")
        lblPrevActualValue(iPoint).Caption = IIf(lblPrevActualValue(iPoint).Enabled, Format(Prev_MfcCal.PointData(iPoint).ActualValue, "####0.0##"), "")
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
    If (Not CalReadOnly) Then PRG_INFO(STN_INFO(SelCalStn).AspiratorNum).RequestRdy = False
    If (Not CalReadOnly) Then CalValves False
    Unload Me
    Set frmMfcCal = Nothing
End Sub

Private Sub txtEquipment_Change()
'    bUnsavedCal = True
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
    ' update the Mfc SP
    If SelCalPoint = Index Then
        DesFlowSLPM_Validate SelCalPoint
'        txtNewActualValue(SelCalPoint).Enabled = True
'        If bFormLoaded Then txtNewRawValue(SelCalPoint).SetFocus
    End If
End Sub

Private Sub txtNewRawValue_Click(Index As Integer)
    ' Select the row that was clicked on
    If SelCalPoint <> Index Then
        SelectCalPoint Index
        DesFlowSLPM_Validate SelCalPoint
'        DisplayMfcAll
'        txtNewActualValue(SelCalPoint).Enabled = True
'        If bFormLoaded Then txtNewRawValue(SelCalPoint).SetFocus
    End If
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
SetErrModule 7781, 7
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
Dim Idx As Integer
Dim idxList As Integer

    idxList = 0
    For Idx = 0 To (PressUnits.ListCount - 1)
        If (Trim(unitsText) = Trim(PressUnits.List(Idx))) Then idxList = Idx
    Next Idx
    PressUnitsIndex = idxList
End Function

Private Function TempUnitsIndex(ByVal unitsText As String) As Integer
' returns ListIndex of TempUnits that matches unitsText
' defaults to 0
Dim Idx As Integer
Dim idxList As Integer

    idxList = 0
    For Idx = 0 To (TempUnits.ListCount - 1)
        If (Trim(unitsText) = Trim(TempUnits.List(Idx))) Then idxList = Idx
    Next Idx
    TempUnitsIndex = idxList
End Function

Private Sub CopyCal(ByVal station As Integer, ByVal iMfc As Integer)
    ' copy calibration to other functional user of a "shared" physical mfc
    Dim flag As Boolean
    
    flag = False
    Select Case STN_INFO(station).Type
        Case STN_LIVEREG_TYPE
            ' LF and N2 may share an mfc
            Select Case srcMfc
                Case MFCNITROGEN
                    srcMfc = MFCNITROGEN
                    dstMfc = MFCLIVEFUEL
                    srcFunc = asNitrogenFlowSP
                    dstFunc = asLiveFuelVaporFlowSP
                    ' do mfc's share same nonzero addr and chan ???
                    If ((Stn_AIO(station, srcFunc).addr = Stn_AIO(station, dstFunc).addr) _
                            And _
                        (Stn_AIO(station, srcFunc).chan = Stn_AIO(station, dstFunc).chan) _
                            And _
                        (Stn_AIO(station, srcFunc).addr <> 0) _
                            And _
                        (Stn_AIO(station, srcFunc).chan <> 0)) _
                        Then
                            flag = True
                    End If
                Case MFCLIVEFUEL
                    srcMfc = MFCLIVEFUEL
                    dstMfc = MFCNITROGEN
                    srcFunc = asLiveFuelVaporFlowSP
                    dstFunc = asNitrogenFlowSP
                    ' do mfc's share same nonzero addr and chan ???
                    If ((Stn_AIO(station, srcFunc).addr = Stn_AIO(station, dstFunc).addr) _
                            And _
                        (Stn_AIO(station, srcFunc).chan = Stn_AIO(station, dstFunc).chan) _
                            And _
                        (Stn_AIO(station, srcFunc).addr <> 0) _
                            And _
                        (Stn_AIO(station, srcFunc).chan <> 0)) _
                        Then
                            flag = True
                    End If
            End Select
        Case STN_LIVEORVR2_TYPE
            ' LF and N2 may share an mfc
            ' ORVRLF and ORVRN2 may share an mfc
            Select Case iMfc
                Case MFCNITROGEN
                    srcMfc = MFCNITROGEN
                    dstMfc = MFCLIVEFUEL
                    srcFunc = asNitrogenFlowSP
                    dstFunc = asLiveFuelVaporFlowSP
                    ' do mfc's share same nonzero addr and chan ???
                    If ((Stn_AIO(station, srcFunc).addr = Stn_AIO(station, dstFunc).addr) _
                            And _
                        (Stn_AIO(station, srcFunc).chan = Stn_AIO(station, dstFunc).chan) _
                            And _
                        (Stn_AIO(station, srcFunc).addr <> 0) _
                            And _
                        (Stn_AIO(station, srcFunc).chan <> 0)) _
                        Then
                            flag = True
                    End If
                Case MFCLIVEFUEL
                    srcMfc = MFCLIVEFUEL
                    dstMfc = MFCNITROGEN
                    srcFunc = asLiveFuelVaporFlowSP
                    dstFunc = asNitrogenFlowSP
                    ' do mfc's share same nonzero addr and chan ???
                    If ((Stn_AIO(station, srcFunc).addr = Stn_AIO(station, dstFunc).addr) _
                            And _
                        (Stn_AIO(station, srcFunc).chan = Stn_AIO(station, dstFunc).chan) _
                            And _
                        (Stn_AIO(station, srcFunc).addr <> 0) _
                            And _
                        (Stn_AIO(station, srcFunc).chan <> 0)) _
                        Then
                            flag = True
                    End If
                Case MFCORVRNIT
                    srcMfc = MFCORVRNIT
                    dstMfc = MFCORVRLIVE
                    srcFunc = asNitrogenORVRFlowSP
                    dstFunc = asLiveFuelVaporORVRFlowSP
                    ' do mfc's share same nonzero addr and chan ???
                    If ((Stn_AIO(station, srcFunc).addr = Stn_AIO(station, dstFunc).addr) _
                            And _
                        (Stn_AIO(station, srcFunc).chan = Stn_AIO(station, dstFunc).chan) _
                            And _
                        (Stn_AIO(station, srcFunc).addr <> 0) _
                            And _
                        (Stn_AIO(station, srcFunc).chan <> 0)) _
                        Then
                            flag = True
                    End If
                Case MFCORVRLIVE
                    srcMfc = MFCORVRLIVE
                    dstMfc = MFCORVRNIT
                    srcFunc = asLiveFuelVaporORVRFlowSP
                    dstFunc = asNitrogenORVRFlowSP
                    ' do mfc's share same nonzero addr and chan ???
                    If ((Stn_AIO(station, srcFunc).addr = Stn_AIO(station, dstFunc).addr) _
                            And _
                        (Stn_AIO(station, srcFunc).chan = Stn_AIO(station, dstFunc).chan) _
                            And _
                        (Stn_AIO(station, srcFunc).addr <> 0) _
                            And _
                        (Stn_AIO(station, srcFunc).chan <> 0)) _
                        Then
                            flag = True
                    End If
            End Select
        Case Else
            ' nothing to do; no "shared" mfc's
    End Select
    If flag Then
        frmQuestion.Top = 2400
        lblQuestion(1).Caption = "Copy " & Mfc_Description(srcMfc) & " Calibration"
        lblQuestion(2).Caption = "to"
        lblQuestion(3).Caption = Mfc_Description(dstMfc) & " Calibration ?"
    End If
End Sub

Private Function RawValuesIndex(ByVal unitsText As String) As Integer
' returns ListIndex of RawValues that matches unitsText
' defaults to 0
Dim Idx As Integer
Dim idxList As Integer

    idxList = 0
    For Idx = 0 To (RawValues.ListCount - 1)
        If (Trim(unitsText) = Trim(RawValues.List(Idx))) Then idxList = Idx
    Next Idx
    RawValuesIndex = idxList
End Function


