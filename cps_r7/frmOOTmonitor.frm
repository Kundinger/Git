VERSION 5.00
Begin VB.Form frmOOTmonitor 
   Caption         =   "OOTs"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "frmOOTmonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmStnSelection 
      Caption         =   "Station"
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
      Left            =   1440
      TabIndex        =   63
      Top             =   5760
      Width           =   2325
      Begin VB.CommandButton cmdStnUp 
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
         Left            =   1485
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOOTmonitor.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "next live fuel station"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.CommandButton cmdStnDn 
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
         Picture         =   "frmOOTmonitor.frx":5EE4
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "previous livefuel station"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.TextBox txtDispStn 
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
         Left            =   765
         MaxLength       =   2
         TabIndex        =   64
         Text            =   "9"
         Top             =   253
         Width           =   720
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   19
      Left            =   480
      TabIndex        =   62
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   19
      Left            =   4110
      TabIndex        =   61
      Top             =   5280
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   19
      Left            =   2850
      TabIndex        =   60
      Top             =   5280
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   18
      Left            =   480
      TabIndex        =   59
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   18
      Left            =   4110
      TabIndex        =   58
      Top             =   5040
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   18
      Left            =   2850
      TabIndex        =   57
      Top             =   5040
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   17
      Left            =   480
      TabIndex        =   56
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   17
      Left            =   4110
      TabIndex        =   55
      Top             =   4800
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   17
      Left            =   2850
      TabIndex        =   54
      Top             =   4800
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   16
      Left            =   480
      TabIndex        =   53
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   16
      Left            =   4110
      TabIndex        =   52
      Top             =   4560
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   16
      Left            =   2850
      TabIndex        =   51
      Top             =   4560
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   50
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   15
      Left            =   4110
      TabIndex        =   49
      Top             =   4320
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   15
      Left            =   2850
      TabIndex        =   48
      Top             =   4320
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   14
      Left            =   480
      TabIndex        =   47
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   14
      Left            =   4110
      TabIndex        =   46
      Top             =   4080
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   14
      Left            =   2850
      TabIndex        =   45
      Top             =   4080
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   13
      Left            =   480
      TabIndex        =   44
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   13
      Left            =   4110
      TabIndex        =   43
      Top             =   3840
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   13
      Left            =   2850
      TabIndex        =   42
      Top             =   3840
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   12
      Left            =   480
      TabIndex        =   41
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   12
      Left            =   4110
      TabIndex        =   40
      Top             =   3600
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   12
      Left            =   2850
      TabIndex        =   39
      Top             =   3600
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   11
      Left            =   480
      TabIndex        =   38
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   11
      Left            =   4110
      TabIndex        =   37
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   11
      Left            =   2850
      TabIndex        =   36
      Top             =   3360
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   10
      Left            =   480
      TabIndex        =   35
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   10
      Left            =   4110
      TabIndex        =   34
      Top             =   3120
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   10
      Left            =   2850
      TabIndex        =   33
      Top             =   3120
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   9
      Left            =   480
      TabIndex        =   32
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   9
      Left            =   4110
      TabIndex        =   31
      Top             =   2880
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   9
      Left            =   2850
      TabIndex        =   30
      Top             =   2880
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   8
      Left            =   480
      TabIndex        =   29
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   8
      Left            =   4110
      TabIndex        =   28
      Top             =   2640
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   8
      Left            =   2850
      TabIndex        =   27
      Top             =   2640
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   7
      Left            =   480
      TabIndex        =   26
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   7
      Left            =   4110
      TabIndex        =   25
      Top             =   2400
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   7
      Left            =   2850
      TabIndex        =   24
      Top             =   2400
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   6
      Left            =   480
      TabIndex        =   23
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   6
      Left            =   4110
      TabIndex        =   22
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   6
      Left            =   2850
      TabIndex        =   21
      Top             =   2160
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   20
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   5
      Left            =   4110
      TabIndex        =   19
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   5
      Left            =   2850
      TabIndex        =   18
      Top             =   1920
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   4
      Left            =   4110
      TabIndex        =   16
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   4
      Left            =   2850
      TabIndex        =   15
      Top             =   1680
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   14
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   3
      Left            =   4110
      TabIndex        =   13
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   3
      Left            =   2850
      TabIndex        =   12
      Top             =   1440
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   2
      Left            =   4110
      TabIndex        =   10
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   2
      Left            =   2850
      TabIndex        =   9
      Top             =   1200
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   1
      Left            =   4110
      TabIndex        =   7
      Top             =   960
      Width           =   420
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   1
      Left            =   2850
      TabIndex        =   6
      Top             =   960
      Width           =   300
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "an OotOfTolerance"
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblDescCol 
      Alignment       =   2  'Center
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblOotTCol 
      Alignment       =   2  'Center
      Caption         =   "OOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblTol 
      Alignment       =   2  'Center
      Caption         =   "false"
      Height          =   210
      Index           =   0
      Left            =   4117
      TabIndex        =   2
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lblCountCol 
      Alignment       =   2  'Center
      Caption         =   "Dwell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  'Center
      Caption         =   "888"
      Height          =   210
      Index           =   0
      Left            =   2857
      TabIndex        =   0
      Top             =   720
      Width           =   300
   End
End
Attribute VB_Name = "frmOOTmonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private iStn As Integer
Private iShift As Integer
Private iRow As Integer
'

Private Sub Form_Load()
    iStn = 1
    iShift = Stn_ActiveShift(iStn)
    txtDispStn.text = Format(iStn, "#0")
    For iRow = 12 To lblDesc.UBound
        lblDesc(iRow).Visible = False
        lblCnt(iRow).Visible = False
        lblTol(iRow).Visible = False
    Next iRow
End Sub

Private Sub cmdStnDn_Click()
'
    iStn = iStn - 1
    If iStn < 1 Then iStn = NR_STN
    iShift = Stn_ActiveShift(iStn)
    txtDispStn.text = Format(iStn, "#0")
End Sub

Private Sub cmdStnUp_Click()
    iStn = iStn + 1
    If iStn > NR_STN Then iStn = 1
    iShift = Stn_ActiveShift(iStn)
    txtDispStn.text = Format(iStn, "#0")
End Sub

Private Sub tmrUpdate_Timer()
    ' Set all counts to zero (if in Alarm)
    iRow = 0
    lblDesc(iRow).Caption = "Purge Flow"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).PurFlowOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).PurFlowOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).PurFlowOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Butane Flow"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).BtnFlowOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).BtnFlowOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).BtnFlowOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Nitrogen Flow"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).NitFlowOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).NitFlowOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).NitFlowOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Fuel Temp"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).FuelTempOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).FuelTempOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).FuelTempOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Air Temp"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).AirTempOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).AirTempOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).AirTempOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Air Moisture"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).AirMoistOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).AirMoistOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).AirMoistOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Load Rate"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).LoadRateOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).LoadRateOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).LoadRateOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Purge DP"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).PurgeDpOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).PurgeDpOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).PurgeDpOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Purge Oven"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).PurgeOvenOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).PurgeOvenOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).PurgeOvenOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "WaterBath"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).WaterBathOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).WaterBathOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).WaterBathOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Fuel Level"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).FuelLevelOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).FuelLevelOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).FuelLevelOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Storage Level"
    lblCnt(iRow).Caption = Format(OOTs(iStn, iShift).StorageLevelOOTCnt, "###0")
    lblTol(iRow).Caption = IIf(OOTs(iStn, iShift).StorageLevelOOT, "OOT", "ok")
    lblTol(iRow).ForeColor = IIf(OOTs(iStn, iShift).StorageLevelOOT, Alarm_ForeColor, MEDBLUE)
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Fuel is Dead"
    lblCnt(iRow).Caption = Format(AdfControl(iStn).LiveFuelDensityDeadCnt, "###0")
    lblTol(iRow).Caption = ""
    lblTol(iRow).ForeColor = MEDBLUE
    iRow = iRow + 1
    lblDesc(iRow).Caption = "Fuel is Weak"
    lblCnt(iRow).Caption = Format(AdfControl(iStn).LiveFuelDensityWeakCnt, "###0")
    lblTol(iRow).Caption = ""
    lblTol(iRow).ForeColor = MEDBLUE
End Sub


