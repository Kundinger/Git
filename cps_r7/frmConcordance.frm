VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmConcordance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Concordance"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmConcordance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScreen 
      Interval        =   150
      Left            =   4920
      Top             =   0
   End
   Begin Threed.SSPanel pnlConcordance 
      Height          =   3090
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5200
      _Version        =   65536
      _ExtentX        =   9172
      _ExtentY        =   5450
      _StockProps     =   15
      Caption         =   "Concordance"
      ForeColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Alignment       =   8
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         DisabledPicture =   "frmConcordance.frx":57E2
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
         Picture         =   "frmConcordance.frx":5B24
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Reset Values & Timer"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin VB.CommandButton cmdCalcCode 
         Caption         =   "Calc"
         DisabledPicture =   "frmConcordance.frx":5E66
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
         Left            =   4440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConcordance.frx":61A8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cycle through MFC Calculations"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   640
      End
      Begin Threed.SSPanel pnlNetGrams 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "net grams loaded"
         Top             =   555
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlNetGrams 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "net grams loaded"
         Top             =   915
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlNetGrams 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "net grams loaded"
         Top             =   1290
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlNetGrams 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Top             =   1665
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlGramsPerHr 
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   7
         ToolTipText     =   "net grams/hour"
         Top             =   555
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlGramsPerHr 
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   8
         ToolTipText     =   "net grams/hour"
         Top             =   915
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlGramsPerHr 
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   9
         ToolTipText     =   "net grams/hour"
         Top             =   1290
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlGramsPerHr 
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   10
         Top             =   1665
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "012345"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin VB.Label lblGramsPerHr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblTotalScale 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Scale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   23
         Top             =   1335
         Width           =   1110
      End
      Begin VB.Label lblButaneMfc 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane MFC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   1710
         Width           =   1110
      End
      Begin VB.Label lblPri_Scale 
         BackStyle       =   0  'Transparent
         Caption         =   "Pri Scale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblAux_Scale 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Scale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label lblNetGrams 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "net"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblNetUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   2250
         TabIndex        =   18
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblNetUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   2250
         TabIndex        =   17
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblNetUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   2250
         TabIndex        =   16
         Top             =   1335
         Width           =   585
      End
      Begin VB.Label lblNetUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   2250
         TabIndex        =   15
         Top             =   1710
         Width           =   585
      End
      Begin VB.Label lblRateUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams/hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   4050
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblRateUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams/hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   4050
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblRateUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams/hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   4050
         TabIndex        =   12
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label lblRateUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams/hour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   4050
         TabIndex        =   11
         Top             =   1710
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmConcordance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''' form Concordance
' error module 2457
Option Explicit
'
Private sVal As Double
Private sSec As Double
Private sRate As Single
Private LoadEql_CalcCode As Integer
Const CalcCode_GmPerHr = 0
Const CalcCode_LiterPerMin = 1
Const CalcCode_GmPerLiter = 2

Sub ClearScreenOpenFlags()
Dim iStn As Integer
Dim iShft As Integer

    For iStn = 1 To NR_STN
        For iShft = 1 To NR_SHIFT
            LoadControl(iStn, iShft).ConcordanceIsOpen = False
        Next iShft
    Next iStn

End Sub

Sub UnloadScreen()
    ClearScreenOpenFlags
    Unload Me
    Set frmConcordance = Nothing
End Sub

Private Sub cmdCalcCode_Click()
    ' increment to next CalcCode
    ' 0 = MFC Gram/Hour
    ' 1 = MFC Liter/Minute
    ' 2 = MFC Gram/Liter
    LoadEql_CalcCode = IIf(LoadEql_CalcCode < CalcCode_GmPerLiter, LoadEql_CalcCode + 1, 0)
End Sub

Private Sub cmdReset_Click()
    Stn_LoadEql_StartTimer(DispStn, DispShift) = StationControl(DispStn, DispShift).TestTimer
    Stn_LoadEql_StartAuxWt(DispStn, DispShift) = StationControl(DispStn, DispShift).AuxScaleWt
    Stn_LoadEql_StartPriWt(DispStn, DispShift) = StationControl(DispStn, DispShift).PriScaleWt
    Stn_LoadEql_StartLoadTotal(DispStn, DispShift) = LoadControl(DispStn, DispShift).loadTotalGrams
End Sub

Sub UpdateScreen()

    ' **************************************************************************************
    ' LOAD CONCORDANCE DISPLAY
    '     only display if user = APS
    SetErrModule 2457, 13
    
                    
    If ((STN_INFO(DispStn).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(DispStn).Type = STN_LIVEREG_TYPE) And (StationRecipe(DispStn, DispShift).LiveFuel)) Or ((STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(DispStn, DispShift).LiveFuel)) Or (Not CheckPass("G", False))) Then
        ' close screen
        UnloadScreen
    Else
        If StationControl(DispStn, DispShift).Mode = VBLOAD _
          And LoadControl(DispStn, DispShift).Phase = LoadLoading _
          And StationRecipe(DispStn, DispShift).UsePriScale = True Then
            If StationControl(DispStn, DispShift).TestTimer > Stn_LoadEql_StartTimer(DispStn, DispShift) Then
                    
                    sSec = CDbl(StationControl(DispStn, DispShift).TestTimer - Stn_LoadEql_StartTimer(DispStn, DispShift))
                    ' Primary Scale
                    sVal = CDbl(StationControl(DispStn, DispShift).PriScaleWt) - Stn_LoadEql_StartPriWt(DispStn, DispShift)
                    If (Format(sVal, "####0.000") <> pnlNetGrams(0).Caption) Then
                        pnlNetGrams(0).Caption = Format(sVal, "####0.000")
                        If sSec > 0 Then pnlGramsPerHr(0).Caption = Format(3600 * CDbl(sVal / sSec), "###0.000")
                    End If
                    ' Aux Scale
                    If StationRecipe(DispStn, DispShift).UseAuxScale = True Then
                        sVal = CDbl(StationControl(DispStn, DispShift).AuxScaleWt) - Stn_LoadEql_StartAuxWt(DispStn, DispShift)
                        If (Format(sVal, "####0.000") <> pnlNetGrams(1).Caption) Then
                            pnlNetGrams(1).Caption = Format(sVal, "####0.000")
                            If sSec > 0 Then pnlGramsPerHr(1).Caption = Format(3600 * CDbl(sVal / sSec), "###0.000")
                        End If
                        lblAux_Scale.Left = 120
                        pnlNetGrams(1).Left = 1200
                        lblNetUnits(1).Left = 2250
                        pnlGramsPerHr(1).Left = 2940
                        lblRateUnits(1).Left = 3990
                    Else
                        lblAux_Scale.Left = OutOfSight
                        pnlNetGrams(1).Left = OutOfSight
                        lblNetUnits(1).Left = OutOfSight
                        pnlGramsPerHr(1).Left = OutOfSight
                        lblRateUnits(1).Left = OutOfSight
                    End If
                    ' Total Scale
                    sVal = CDbl(StationControl(DispStn, DispShift).PriScaleWt) + StationControl(DispStn, DispShift).AuxScaleWt
                    sVal = sVal - CDbl(Stn_LoadEql_StartPriWt(DispStn, DispShift) + Stn_LoadEql_StartAuxWt(DispStn, DispShift))
                    If (Format(sVal, "####0.000") <> pnlNetGrams(2).Caption) Then
                        pnlNetGrams(2).Caption = Format(sVal, "####0.000")
                        If sSec > 0 Then pnlGramsPerHr(2).Caption = Format(3600 * CDbl(sVal / sSec), "###0.000")
                    End If
                    ' Butane MFC
                    cmdCalcCode.Top = IIf(CheckPass("H", False), cmdReset.Top, OutOfSight)
                    LoadEql_CalcCode = IIf(CheckPass("H", False), LoadEql_CalcCode, CalcCode_GmPerHr)
                    Select Case LoadEql_CalcCode
                        Case CalcCode_GmPerHr
                            ' Calculate Rate of Weight Change in Grams/Hr
                            lblNetUnits(3).Caption = "grams"
                            lblRateUnits(3).Caption = "grams/hour"
                            lblNetUnits(3).ToolTipText = "net grams loaded"
                            lblRateUnits(3).ToolTipText = "net grams/hour"
                            sVal = CDbl(LoadControl(DispStn, DispShift).loadTotalGrams - Stn_LoadEql_StartLoadTotal(DispStn, DispShift))
                            If (Format(sVal, "####0.000") <> pnlNetGrams(3).Caption) Then
                                pnlNetGrams(3).Caption = Format(sVal, "####0.000")
                                If sSec > 0 Then pnlGramsPerHr(3).Caption = Format(3600 * CDbl(sVal / sSec), "###0.000")
                            End If
                        Case CalcCode_LiterPerMin
                            ' Calculate Rate of Flow in SLPM
                            lblNetUnits(3).Caption = "liters"
                            lblRateUnits(3).Caption = "liters/min"
                            lblNetUnits(3).ToolTipText = "net liters loaded"
                            lblRateUnits(3).ToolTipText = "net liters/min (SLPM)"
                            sVal = CDbl(LoadControl(DispStn, DispShift).loadTotalGrams - Stn_LoadEql_StartLoadTotal(DispStn, DispShift)) / CDbl(GramsPerLiter)
                            If (Format(sVal, "####0.000") <> pnlNetGrams(3).Caption) Then
                                pnlNetGrams(3).Caption = Format(sVal, "####0.000")
                                If sSec > 0 Then pnlGramsPerHr(3).Caption = Format(60 * CDbl(sVal / sSec), "###0.0000")
                            End If
                        Case CalcCode_GmPerLiter
                            ' Calculate Actual Butane Density in Grams/Liter (assumes Valid SLPM)
                            lblNetUnits(3).Caption = "liters"
                            lblRateUnits(3).Caption = "grams/liter"
                            lblNetUnits(3).ToolTipText = "net liters loaded"
                            lblRateUnits(3).ToolTipText = "net grams/liter"
                            sVal = CDbl(LoadControl(DispStn, DispShift).loadTotalGrams - Stn_LoadEql_StartLoadTotal(DispStn, DispShift)) / CDbl(GramsPerLiter)
                            If (Format(sVal, "####0.000") <> pnlNetGrams(3).Caption) Then
                                pnlNetGrams(3).Caption = Format(sVal, "####0.000")
                                If sVal > 0 Then pnlGramsPerHr(3).Caption = Format(CDbl(pnlNetGrams(2).Caption) / sVal, "#0.000000")
                            End If
                    End Select
                    
            Else
                ' close screen
                UnloadScreen
            End If
        Else
            ' close screen
            UnloadScreen
        End If
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

Private Sub Form_Activate()
    KeyPreview = True
    LoadControl(DispStn, DispShift).ConcordanceIsOpen = True
End Sub

Private Sub Form_Deactivate()
'    ClearScreenOpenFlags
End Sub

Private Sub Form_Load()
'
Dim idx As Integer

    KeyPreview = True
    LoadControl(DispStn, DispShift).ConcordanceIsOpen = True
    
    ' Set Foreground colors
    pnlConcordance.ForeColor = TitlesLabel_ForeColor
    lblNetGrams.ForeColor = TitlesLabel_ForeColor
    lblGramsPerHr.ForeColor = TitlesLabel_ForeColor
    For idx = 0 To 3
        pnlNetGrams(idx).ForeColor = TitlesData_Forecolor
        pnlGramsPerHr(idx).ForeColor = TitlesData_Forecolor
    Next idx

    LoadEql_CalcCode = 0
    pnlConcordance.Top = 30
    pnlConcordance.Left = 30
    lblPri_Scale.Left = 120
    lblAux_Scale.Left = 120
    lblTotalScale.Left = 120
    lblButaneMfc.Left = 120
    lblNetGrams.Left = 1200
    pnlNetGrams(0).Left = 1200
    pnlNetGrams(1).Left = 1200
    pnlNetGrams(2).Left = 1200
    pnlNetGrams(3).Left = 1200
    lblNetUnits(0).Left = 2250
    lblNetUnits(1).Left = 2250
    lblNetUnits(2).Left = 2250
    lblNetUnits(3).Left = 2250
    lblGramsPerHr.Left = 2940
    pnlGramsPerHr(0).Left = 2940
    pnlGramsPerHr(1).Left = 2940
    pnlGramsPerHr(2).Left = 2940
    pnlGramsPerHr(3).Left = 2940
    lblRateUnits(0).Left = 3990
    lblRateUnits(1).Left = 3990
    lblRateUnits(2).Left = 3990
    lblRateUnits(3).Left = 3990
    cmdCalcCode.Top = OutOfSight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearScreenOpenFlags
End Sub

Private Sub tmrScreen_Timer()
    UpdateScreen
End Sub
