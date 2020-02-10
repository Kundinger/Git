VERSION 5.00
Begin VB.Form frmSummary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Summary Statistics"
   ClientHeight    =   4725
   ClientLeft      =   585
   ClientTop       =   1275
   ClientWidth     =   7230
   Icon            =   "frmsumma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Exit"
      DisabledPicture =   "frmsumma.frx":57E2
      DownPicture     =   "frmsumma.frx":6424
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
      Left            =   6240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmsumma.frx":7066
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   3840
   End
   Begin VB.Frame fraLoad 
      Caption         =   "Load Statistics"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6975
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "% Butane"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "slpm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   36
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "slpm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMix_Max 
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
         Left            =   4800
         TabIndex        =   34
         ToolTipText     =   "Mix Ratio Maximum Value"
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblMix_Avg 
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
         Left            =   3720
         TabIndex        =   33
         ToolTipText     =   "Mix Ratio Average Value"
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblMix_Min 
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
         Left            =   2640
         TabIndex        =   32
         ToolTipText     =   "Mix Ratio  Minimum Value"
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblNit_Max 
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
         Left            =   4800
         TabIndex        =   31
         ToolTipText     =   "Nitrogen Maximum Value"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblNit_Avg 
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
         Left            =   3720
         TabIndex        =   30
         ToolTipText     =   "Nitrogen Average Value"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblNit_Min 
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
         Left            =   2640
         TabIndex        =   29
         ToolTipText     =   "Nitrogen Minimum Value"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblBtn_Max 
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
         Left            =   4800
         TabIndex        =   28
         ToolTipText     =   "Butane Maximum Value"
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblBtn_Avg 
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
         Left            =   3720
         TabIndex        =   27
         ToolTipText     =   "Butane Average Value"
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblBtn_Min 
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
         Left            =   2640
         TabIndex        =   26
         ToolTipText     =   "Butane Minimum Value"
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   24
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Mix Ratio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "TankAir Inlet Flow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane Flow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame fraPurge 
      Caption         =   "Purge Statistics"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Grains/Lb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblPurgeAirTemp 
         BackStyle       =   0  'Transparent
         Caption         =   "Degrees C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "slpm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblHum_Max 
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
         Left            =   4800
         TabIndex        =   16
         ToolTipText     =   "Purge Humidity Maximum Value"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblHum_Avg 
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
         Left            =   3720
         TabIndex        =   15
         ToolTipText     =   "Purge Humidity Average Value"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblHum_Min 
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
         Left            =   2640
         TabIndex        =   14
         ToolTipText     =   "Purge Humidity Minimum Value"
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblTemp_Max 
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
         Left            =   4800
         TabIndex        =   13
         ToolTipText     =   "Purge Temperature Maximum Value"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblTemp_Avg 
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
         Left            =   3720
         TabIndex        =   12
         ToolTipText     =   "Purge Temperature Average Value"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblTemp_Min 
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
         Left            =   2640
         TabIndex        =   11
         ToolTipText     =   "Purge Temperature Minimum Value"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lblPAFlow_Max 
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
         Left            =   4800
         TabIndex        =   10
         ToolTipText     =   "Purge Flow Maximum Value"
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblPAFlow_Avg 
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
         Left            =   3720
         TabIndex        =   9
         ToolTipText     =   "Purge Flow Average Value"
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblPAFlow_Min 
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
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Purge Flow Minimum Value"
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Average"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Air Humidity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Air Temperature:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Air Flow:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label lblSettle2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Mass Flow Controller Settling Time Incomplete"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   120
      TabIndex        =   39
      Top             =   3900
      Width           =   2805
   End
   Begin VB.Label lblSettle 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   840
      Left            =   120
      TabIndex        =   38
      Top             =   3720
      Width           =   2895
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 85 ''''''Form SUMMARY.frm ''''''''''''''''''''
Option Explicit

Private Sub Refresh_Stats()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 85, 1

    lblPAFlow_Min = Format(StationStatistics(DispStn, DispShift).Pur.sMin, "#0.00")
    lblPAFlow_Max = Format(StationStatistics(DispStn, DispShift).Pur.sMax, "#0.00")
    lblPAFlow_Avg = Format(StationStatistics(DispStn, DispShift).Pur.sAvg, "#0.00")
    
    lblHum_Min = Format(StationStatistics(DispStn, DispShift).AirMoist.sMin, "#0.00")
    lblHum_Max = Format(StationStatistics(DispStn, DispShift).AirMoist.sMax, "#0.00")
    lblHum_Avg = Format(StationStatistics(DispStn, DispShift).AirMoist.sAvg, "#0.00")
    
    lblTemp_Min = Format(StationStatistics(DispStn, DispShift).AirTemp.sMin, "#0.00")
    lblTemp_Max = Format(StationStatistics(DispStn, DispShift).AirTemp.sMax, "#0.00")
    lblTemp_Avg = Format(StationStatistics(DispStn, DispShift).AirTemp.sAvg, "#0.00")
    
    lblBtn_Min = Format(StationStatistics(DispStn, DispShift).Btn.sMin, "#0.000")
    lblBtn_Max = Format(StationStatistics(DispStn, DispShift).Btn.sMax, "#0.000")
    lblBtn_Avg = Format(StationStatistics(DispStn, DispShift).Btn.sAvg, "#0.000")
    
    lblNit_Min = Format(StationStatistics(DispStn, DispShift).Nit.sMin, "#0.000")
    lblNit_Max = Format(StationStatistics(DispStn, DispShift).Nit.sMax, "#0.000")
    lblNit_Avg = Format(StationStatistics(DispStn, DispShift).Nit.sAvg, "#0.000")
    
    lblMix_Min = Format(StationStatistics(DispStn, DispShift).Mix.sMin, "#0.00")
    lblMix_Max = Format(StationStatistics(DispStn, DispShift).Mix.sMax, "#0.00")
    lblMix_Avg = Format(StationStatistics(DispStn, DispShift).Mix.sAvg, "#0.00")
    
    
    ' If Not NitStat(DispStn, DispShift).FirstTime And NitStat(DispStn, DispShift).sCnt > 1 And StationControl(DispStn, DispShift).Mode = VBLOAD Then
    If StationControl(DispStn, DispShift).Mode = VBLOAD And StationStatistics(DispStn, DispShift).Nit.sCnt > 0 Then
        ' Load Statistics have at least one reading
        lblSettle.Visible = False
        lblSettle2.Visible = False
    ElseIf StationControl(DispStn, DispShift).Mode = VBPURGE And StationStatistics(DispStn, DispShift).Pur.sCnt > 0 Then
        ' Purge Statistics have at least one reading
        lblSettle.Visible = False
        lblSettle2.Visible = False
    Else
        ' no valid statistics at this time for this station,shift
        lblSettle.Visible = True
        lblSettle2.Visible = True
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub cmdReturn_Click()
    Unload Me
    Set frmSummary = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSummary = Nothing
    End If
End Sub

Private Sub Form_Load()

     ' Set Title Foreground color
    fraPurge.ForeColor = Titles_ForeColor
    fraLoad.ForeColor = Titles_ForeColor
    
    KeyPreview = True
    Form_Center Me

    If ((STN_INFO(DispStn).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(DispStn).Type = STN_LIVEREG_TYPE) And StationRecipe(DispStn, DispShift).LiveFuel) Or ((STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE) And StationRecipe(DispStn, DispShift).LiveFuel)) Then
    
        Label20.Caption = "Vapor Carrier Flow"
 
        Label19.Visible = False
        lblBtn_Min.Visible = False
        lblBtn_Avg.Visible = False
        lblBtn_Max.Visible = False
        Label34.Visible = False
            
        Label21.Visible = False
        lblMix_Min.Visible = False
        lblMix_Avg.Visible = False
        lblMix_Max.Visible = False
        Label36.Visible = False
            
    Else
        Label20.Caption = "Nitrogen Flow"
            
        Label19.Visible = True
        lblBtn_Min.Visible = True
        lblBtn_Avg.Visible = True
        lblBtn_Max.Visible = True
        Label34.Visible = True
            
        Label21.Visible = True
        lblMix_Min.Visible = True
        lblMix_Avg.Visible = True
        lblMix_Max.Visible = True
        Label36.Visible = True
            
    End If

    If USINGC Then lblPurgeAirTemp.Caption = "Degrees C"
    If USINGF Then lblPurgeAirTemp.Caption = "Degrees F"
    
    Refresh_Stats

End Sub
Private Sub Timer1_Timer()

  If frmSummary.Visible = True Then
     Refresh_Stats    ' local sub see above
  End If
  
End Sub
