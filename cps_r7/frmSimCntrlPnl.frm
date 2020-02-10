VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimCntrlPnl 
   Caption         =   "Simulation Control Panel"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   Icon            =   "frmSimCntrlPnl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmOOT 
      Caption         =   "OOT Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   4935
      Begin VB.Frame frmPasOot 
         Caption         =   "PAS Error from SP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   3960
         Width           =   4695
         Begin VB.TextBox txtPasError 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Index           =   2
            Left            =   2040
            TabIndex        =   37
            Text            =   "0.0 %"
            ToolTipText     =   "PurgeAir RH Error in Percent"
            Top             =   720
            Width           =   1000
         End
         Begin VB.TextBox txtPasError 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   34
            Text            =   "0.0 deg"
            ToolTipText     =   "PurgeAir Temp Error in Degrees"
            Top             =   360
            Width           =   1000
         End
         Begin VB.Label lblPasErrUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "percent"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   39
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label lblPasErr 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Relative Humidity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   750
            Width           =   1740
         End
         Begin VB.Label lblPasErrUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "degrees"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   36
            Top             =   390
            Width           =   1005
         End
         Begin VB.Label lblPasErr 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Temperature"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   390
            Width           =   1740
         End
      End
      Begin VB.Frame frmMfcOffsets 
         Caption         =   "Percent MFC Error from SP"
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
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtNitMfcError 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Index           =   1
            Left            =   525
            TabIndex        =   5
            Text            =   "0.0%"
            Top             =   600
            Width           =   1000
         End
         Begin VB.TextBox txtButMfcError 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Index           =   1
            Left            =   1515
            TabIndex        =   4
            Text            =   "0.0%"
            Top             =   600
            Width           =   1000
         End
         Begin VB.TextBox txtPrgMfcError 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Index           =   1
            Left            =   2520
            TabIndex        =   3
            Text            =   "0.0%"
            Top             =   600
            Width           =   1000
         End
         Begin VB.TextBox txtLfvMfcError 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Index           =   1
            Left            =   3525
            TabIndex        =   2
            Text            =   "0.0%"
            Top             =   600
            Width           =   1000
         End
         Begin VB.Label lblSta 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sta"
            BeginProperty Font 
               Name            =   "Arial"
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
            Top             =   360
            Width           =   300
         End
         Begin VB.Label lblNitMfc 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Nitrogen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   525
            TabIndex        =   10
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblButMfc 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Butane"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1515
            TabIndex        =   9
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblPrgMfc 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Purge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   8
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblLfvMfc 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "LiveFuel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3525
            TabIndex        =   7
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label lblStnNum 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   6
            Top             =   630
            Width           =   180
         End
      End
   End
   Begin VB.Frame frmSimulation 
      Caption         =   "Simulation Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   8340
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   12615
      Begin VB.Frame frmLiveFuel 
         Caption         =   "LiveFuel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   5160
         TabIndex        =   55
         Top             =   7320
         Width           =   4695
         Begin VB.TextBox txtLfDensity 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   2040
            TabIndex        =   56
            Text            =   "0.0 gm/l"
            ToolTipText     =   "LiveFuel Vapor Density in grams per liter"
            Top             =   360
            Width           =   1000
         End
         Begin VB.Label lblLfDensity 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Density"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label lblLfDensityUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "grams per liter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   57
            Top             =   390
            Width           =   1365
         End
      End
      Begin VB.Frame frmScaleSim 
         Caption         =   "Scale Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   3975
         Left            =   9480
         TabIndex        =   47
         Top             =   3000
         Width           =   3015
         Begin VB.Frame frmStartWeight 
            Caption         =   "Job Start Weight"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3615
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   2775
            Begin VB.TextBox txtAuxLoaded 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   1
               Left            =   1560
               TabIndex        =   54
               Text            =   "0.0"
               ToolTipText     =   "Aux Scale Job Start Weight in % of Working Capacity (0-115%)"
               Top             =   600
               Width           =   1000
            End
            Begin VB.TextBox txtPriLoaded 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   1
               Left            =   525
               TabIndex        =   49
               Text            =   "0.0"
               ToolTipText     =   "Primary Scale Job Start Weight in % of Working Capacity (0-115%)"
               Top             =   600
               Width           =   1000
            End
            Begin VB.Label lblAuxLoaded 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Aux"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   53
               Top             =   360
               Width           =   1000
            End
            Begin VB.Label lblSta4 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Width           =   300
            End
            Begin VB.Label lblPriLoaded 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Primary"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   525
               TabIndex        =   51
               Top             =   360
               Width           =   1000
            End
            Begin VB.Label lblStnNum4 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   50
               Top             =   630
               Width           =   180
            End
         End
      End
      Begin VB.CheckBox chkSimulationNoise 
         Alignment       =   1  'Right Justify
         Caption         =   "Add noise to simulated values"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   9180
         TabIndex        =   46
         ToolTipText     =   "Run Simulation when IOScan is OFF"
         Top             =   600
         Width           =   3200
      End
      Begin VB.Frame frmLoadPressure 
         Caption         =   "Load Pressure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   3975
         Left            =   7320
         TabIndex        =   40
         Top             =   3000
         Width           =   2055
         Begin VB.Frame frmLoadPressIn 
            Caption         =   "Load Pressure In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3615
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1815
            Begin VB.TextBox txtLoadPressure 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   1
               Left            =   525
               TabIndex        =   42
               Text            =   "0.0"
               Top             =   600
               Width           =   1000
            End
            Begin VB.Label lblStnNum3 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   45
               Top             =   630
               Width           =   180
            End
            Begin VB.Label lblLoadPressure 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "psi"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   525
               TabIndex        =   44
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblSta3 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   300
            End
         End
      End
      Begin MSComctlLib.ImageList imgSimConPnlDisabled 
         Left            =   1920
         Top             =   2400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSimCntrlPnl.frx":57E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSimCntrlPnl.frx":6434
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSimConPnlHot 
         Left            =   1200
         Top             =   2400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSimCntrlPnl.frx":7086
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSimCntrlPnl.frx":7CD8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSimConPnlNormal 
         Left            =   480
         Top             =   2400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSimCntrlPnl.frx":892A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSimCntrlPnl.frx":957C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame frmOtherInputs 
         Caption         =   "Other Inputs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1935
         Left            =   5160
         TabIndex        =   27
         Top             =   960
         Width           =   2055
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "PurgeAir Ready"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   5
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "UPS"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   6
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "Ext. Alarm"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "RunLocalPAS"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frmMsgs 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1935
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2775
         Begin VB.Label lblMessages 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "messages"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1800
            Left            =   0
            TabIndex        =   26
            Top             =   120
            Width           =   2775
         End
      End
      Begin VB.Timer tmrUpdate 
         Interval        =   1000
         Left            =   1200
         Top             =   360
      End
      Begin VB.Frame frmAlarmIO 
         Caption         =   "Alarm Inputs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1935
         Left            =   3000
         TabIndex        =   20
         Top             =   960
         Width           =   2055
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "EStop"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "Flow"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1080
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "Doors"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnlAlmIn 
            Height          =   360
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   635
            _StockProps     =   15
            Caption         =   "LEL"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frmLeakCheck 
         Caption         =   "Leak Check Controls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   3975
         Left            =   5160
         TabIndex        =   14
         Top             =   3000
         Width           =   2055
         Begin VB.Frame frmLeakError 
            Caption         =   "Leak Error Rate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3615
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1815
            Begin VB.TextBox txtLeakError 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Index           =   1
               Left            =   525
               TabIndex        =   16
               Text            =   "0.0%"
               Top             =   600
               Width           =   1000
            End
            Begin VB.Label lblSta2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Sta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   300
            End
            Begin VB.Label lblLeakErr 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "% per min."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   525
               TabIndex        =   18
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label lblStnNum2 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   17
               Top             =   630
               Width           =   180
            End
         End
      End
      Begin VB.CheckBox chkSimulation 
         Alignment       =   1  'Right Justify
         Caption         =   "Simulation ON when IOScan is Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   9180
         TabIndex        =   13
         ToolTipText     =   "Run Simulation when IOScan is OFF"
         Top             =   240
         Width           =   3200
      End
   End
   Begin MSComctlLib.Toolbar tbrSimConPnl 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "imgSimConPnlNormal"
      DisabledImageList=   "imgSimConPnlDisabled"
      HotImageList    =   "imgSimConPnlHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Description     =   "Save Settings"
            Object.ToolTipText     =   "Save Settings"
            Object.Tag             =   "save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   11690
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            Description     =   "Close Screen"
            Object.ToolTipText     =   "Close Screen"
            Object.Tag             =   "close"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSimCntrlPnl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Idx, iStn As Integer
Private HdrRowTop As Integer
Private Stn1stRowTop As Integer
Private StnNumLeft As Integer
Private StnNumColWidth As Integer
Private StnRowHeight As Integer
Private StnColWidth As Integer

Private Sub chkSimulation_Click()
    lblMessages.Caption = ""
End Sub

Private Sub pnlAlmIn_Click(Index As Integer)
    ' Toggle Input Value
    Select Case Index
        Case 0
            OptoDIO(Com_DIO(icEStopSw).addr, Com_DIO(icEStopSw).chan).RawValue = _
                IIf(OptoDIO(Com_DIO(icEStopSw).addr, Com_DIO(icEStopSw).chan).RawValue, False, True)
        Case 1
            OptoDIO(Com_DIO(icExhaustFlowFS).addr, Com_DIO(icExhaustFlowFS).chan).RawValue = _
                IIf(OptoDIO(Com_DIO(icExhaustFlowFS).addr, Com_DIO(icExhaustFlowFS).chan).RawValue, False, True)
        Case 2
            OptoDIO(Com_DIO(icDoorSw).addr, Com_DIO(icDoorSw).chan).RawValue = _
                IIf(OptoDIO(Com_DIO(icDoorSw).addr, Com_DIO(icDoorSw).chan).RawValue, False, True)
        Case 3
            OptoDIO(Com_DIO(ic20LelGasSw).addr, Com_DIO(ic20LelGasSw).chan).RawValue = _
                IIf(OptoDIO(Com_DIO(ic20LelGasSw).addr, Com_DIO(ic20LelGasSw).chan).RawValue, False, True)
        Case 4
            OptoDIO(Com_DIO(icPASReadyOut).addr, Com_DIO(icPASReadyOut).chan).RawValue = _
                IIf(OptoDIO(Com_DIO(icPASReadyOut).addr, Com_DIO(icPASReadyOut).chan).RawValue, False, True)
        Case 5
            If USINGUPS Then
                If (Com_DIO(icUpsActiveSw).addr <> 0) Or (Com_DIO(icUpsActiveSw).chan <> 0) Then
                    OptoDIO(Com_DIO(icUpsActiveSw).addr, Com_DIO(icUpsActiveSw).chan).RawValue = _
                        IIf(OptoDIO(Com_DIO(icUpsActiveSw).addr, Com_DIO(icUpsActiveSw).chan).RawValue, False, True)
                    OptoDIO(Com_DIO(icUpsFaultSw).addr, Com_DIO(icUpsFaultSw).chan).RawValue = _
                        OptoDIO(Com_DIO(icUpsActiveSw).addr, Com_DIO(icUpsActiveSw).chan).RawValue
                End If
            End If
        Case 6
            If USING_EXT_CONTACTS Then
                If (Com_DIO(icExtAlmContactSw).addr <> 0) Or (Com_DIO(icExtAlmContactSw).chan <> 0) Then
                    OptoDIO(Com_DIO(icExtAlmContactSw).addr, Com_DIO(icExtAlmContactSw).chan).RawValue = _
                        IIf(OptoDIO(Com_DIO(icExtAlmContactSw).addr, Com_DIO(icExtAlmContactSw).chan).RawValue, False, True)
                End If
            End If
        Case 7
            If USINGPASLOCALCONTROL Then
                If (Com_DIO(icPASisRunningIn).addr <> 0) Or (Com_DIO(icPASisRunningIn).chan <> 0) Then
                    OptoDIO(Com_DIO(icPASisRunningIn).addr, Com_DIO(icPASisRunningIn).chan).RawValue = _
                        IIf(OptoDIO(Com_DIO(icPASisRunningIn).addr, Com_DIO(icPASisRunningIn).chan).RawValue, False, True)
                End If
            End If
    End Select
    lblMessages.Caption = ""
    ' Update Buttons Display
    UpdateAlarmButtons
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSimCntrlPnl = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key codes
End Sub

Private Sub Form_Load()

    ' Set Title Foreground color
    frmMfcOffsets.ForeColor = Titles_ForeColor
    frmLeakError.ForeColor = Titles_ForeColor
    frmPasOot.ForeColor = Titles_ForeColor
    frmLoadPressIn.ForeColor = Titles_ForeColor
    frmStartWeight.ForeColor = Titles_ForeColor
    
    lblMessages.Caption = ""
    HdrRowTop = 360
    Stn1stRowTop = 600
    StnNumLeft = 120
    StnNumColWidth = 400
    StnRowHeight = 315
    StnColWidth = 1000
    
    lblSta.Top = HdrRowTop
    lblSta.Left = StnNumLeft
    
    lblSta2.Top = HdrRowTop
    lblSta2.Left = StnNumLeft
    
    lblSta3.Top = HdrRowTop
    lblSta3.Left = StnNumLeft
    
    lblSta4.Top = HdrRowTop
    lblSta4.Left = StnNumLeft
    
    lblNitMfc.Top = HdrRowTop
    lblNitMfc.Left = StnNumLeft + StnNumColWidth
    lblButMfc.Top = HdrRowTop
    lblButMfc.Left = StnNumLeft + StnNumColWidth + (1 * StnColWidth)
    lblPrgMfc.Top = HdrRowTop
    lblPrgMfc.Left = StnNumLeft + StnNumColWidth + (2 * StnColWidth)
    lblLfvMfc.Top = HdrRowTop
    lblLfvMfc.Left = StnNumLeft + StnNumColWidth + (3 * StnColWidth)
    
    lblLeakErr.Top = HdrRowTop
    lblLeakErr.Left = StnNumLeft + StnNumColWidth
    
    lblLoadPressure.Top = HdrRowTop
    lblLoadPressure.Left = StnNumLeft + StnNumColWidth
    
    lblPriLoaded.Top = HdrRowTop
    lblPriLoaded.Left = StnNumLeft + StnNumColWidth
    
    lblAuxLoaded.Top = HdrRowTop
    lblAuxLoaded.Left = StnNumLeft + StnNumColWidth + StnColWidth + 90
    
    iStn = 1
    lblStnNum(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    lblStnNum(iStn).Left = StnNumLeft
    lblStnNum(iStn).Width = 180
    lblStnNum(iStn).Caption = Format(iStn, "0")
    lblStnNum(iStn).ToolTipText = "Station"
    lblStnNum(iStn).Visible = True
    
    txtNitMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtNitMfcError(iStn).Left = StnNumLeft + StnNumColWidth
    txtNitMfcError(iStn).Width = StnColWidth
    txtNitMfcError(iStn).text = ""
    txtNitMfcError(iStn).ToolTipText = "Nitrogen MFC Error"
    
    txtButMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtButMfcError(iStn).Left = StnNumLeft + StnNumColWidth + (1 * StnColWidth)
    txtButMfcError(iStn).Width = StnColWidth
    txtButMfcError(iStn).text = ""
    txtButMfcError(iStn).ToolTipText = "Butane MFC Error"
    
    txtPrgMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtPrgMfcError(iStn).Left = StnNumLeft + StnNumColWidth + (2 * StnColWidth)
    txtPrgMfcError(iStn).Width = StnColWidth
    txtPrgMfcError(iStn).text = ""
    txtPrgMfcError(iStn).ToolTipText = "Purge MFC Error"
    
    txtLfvMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtLfvMfcError(iStn).Left = StnNumLeft + StnNumColWidth + (3 * StnColWidth)
    txtLfvMfcError(iStn).Width = StnColWidth
    txtLfvMfcError(iStn).text = ""
    txtLfvMfcError(iStn).ToolTipText = "LiveFuel MFC Error"
    
    lblStnNum2(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    lblStnNum2(iStn).Left = StnNumLeft
    lblStnNum2(iStn).Width = 180
    lblStnNum2(iStn).Caption = Format(iStn, "0")
    lblStnNum2(iStn).ToolTipText = "Station"
    lblStnNum2(iStn).Visible = True
    
    txtLeakError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtLeakError(iStn).Left = StnNumLeft + StnNumColWidth
    txtLeakError(iStn).Width = StnColWidth
    txtLeakError(iStn).text = ""
    txtLeakError(iStn).ToolTipText = "Leak Check Leak Rate"
    
    
    lblStnNum3(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    lblStnNum3(iStn).Left = StnNumLeft
    lblStnNum3(iStn).Width = 180
    lblStnNum3(iStn).Caption = Format(iStn, "0")
    lblStnNum3(iStn).ToolTipText = "Station"
    lblStnNum3(iStn).Visible = True
    
    txtLoadPressure(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtLoadPressure(iStn).Left = StnNumLeft + StnNumColWidth
    txtLoadPressure(iStn).Width = StnColWidth
    txtLoadPressure(iStn).text = ""
    txtLoadPressure(iStn).ToolTipText = "Load Pressure"
       
       
    lblStnNum4(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    lblStnNum4(iStn).Left = StnNumLeft
    lblStnNum4(iStn).Width = 180
    lblStnNum4(iStn).Caption = Format(iStn, "0")
    lblStnNum4(iStn).ToolTipText = "Station"
    lblStnNum4(iStn).Visible = True

    txtPriLoaded(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtPriLoaded(iStn).Left = StnNumLeft + StnNumColWidth
    txtPriLoaded(iStn).Width = StnColWidth
    txtPriLoaded(iStn).text = ""
    txtPriLoaded(iStn).ToolTipText = "Primary Scale Job Start Weight in % of Working Capacity (0-115%)"
    txtPriLoaded(iStn).Visible = True
    
    txtAuxLoaded(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
    txtAuxLoaded(iStn).Left = StnNumLeft + StnNumColWidth + StnColWidth + 90
    txtAuxLoaded(iStn).Width = StnColWidth
    txtAuxLoaded(iStn).text = ""
    txtAuxLoaded(iStn).ToolTipText = "Aux Scale Job Start Weight in % of Working Capacity (0-115%)"
    txtAuxLoaded(iStn).Visible = True
            
    
    If LAST_STN > 1 Then
        For iStn = 2 To LAST_STN
        
            Load lblStnNum(iStn)
            lblStnNum(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            lblStnNum(iStn).Left = StnNumLeft
            lblStnNum(iStn).Width = 180
            lblStnNum(iStn).Caption = Format(iStn, "0")
            lblStnNum(iStn).ToolTipText = "Station"
            lblStnNum(iStn).Visible = True
    
            Load txtNitMfcError(iStn)
            txtNitMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtNitMfcError(iStn).Left = StnNumLeft + StnNumColWidth
            txtNitMfcError(iStn).Width = StnColWidth
            txtNitMfcError(iStn).text = ""
            txtNitMfcError(iStn).ToolTipText = "Nitrogen MFC Error"
            
            Load txtButMfcError(iStn)
            txtButMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtButMfcError(iStn).Left = StnNumLeft + StnNumColWidth + (1 * StnColWidth)
            txtButMfcError(iStn).Width = StnColWidth
            txtButMfcError(iStn).text = ""
            txtButMfcError(iStn).ToolTipText = "Butane MFC Error"
            
            Load txtPrgMfcError(iStn)
            txtPrgMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtPrgMfcError(iStn).Left = StnNumLeft + StnNumColWidth + (2 * StnColWidth)
            txtPrgMfcError(iStn).Width = StnColWidth
            txtPrgMfcError(iStn).text = ""
            txtPrgMfcError(iStn).ToolTipText = "Purge MFC Error"
            
            Load txtLfvMfcError(iStn)
            txtLfvMfcError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtLfvMfcError(iStn).Left = StnNumLeft + StnNumColWidth + (3 * StnColWidth)
            txtLfvMfcError(iStn).Width = StnColWidth
            txtLfvMfcError(iStn).text = ""
            txtLfvMfcError(iStn).ToolTipText = "LiveFuel MFC Error"
        
            Load lblStnNum2(iStn)
            lblStnNum2(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            lblStnNum2(iStn).Left = StnNumLeft
            lblStnNum2(iStn).Width = 180
            lblStnNum2(iStn).Caption = Format(iStn, "0")
            lblStnNum2(iStn).ToolTipText = "Station"
            lblStnNum2(iStn).Visible = True
    
            Load txtLeakError(iStn)
            txtLeakError(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtLeakError(iStn).Left = StnNumLeft + StnNumColWidth
            txtLeakError(iStn).Width = StnColWidth
            txtLeakError(iStn).text = ""
            txtLeakError(iStn).ToolTipText = "Leak Check Leak Rate"
            
            Load lblStnNum3(iStn)
            lblStnNum3(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            lblStnNum3(iStn).Left = StnNumLeft
            lblStnNum3(iStn).Width = 180
            lblStnNum3(iStn).Caption = Format(iStn, "0")
            lblStnNum3(iStn).ToolTipText = "Station"
            lblStnNum3(iStn).Visible = True
    
            Load txtLoadPressure(iStn)
            txtLoadPressure(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtLoadPressure(iStn).Left = StnNumLeft + StnNumColWidth
            txtLoadPressure(iStn).Width = StnColWidth
            txtLoadPressure(iStn).text = ""
            txtLoadPressure(iStn).ToolTipText = "Load Pressure"
            txtLoadPressure(iStn).Visible = True
            
            Load lblStnNum4(iStn)
            lblStnNum4(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            lblStnNum4(iStn).Left = StnNumLeft
            lblStnNum4(iStn).Width = 180
            lblStnNum4(iStn).Caption = Format(iStn, "0")
            lblStnNum4(iStn).ToolTipText = "Station"
            lblStnNum4(iStn).Visible = True
    
            Load txtPriLoaded(iStn)
            txtPriLoaded(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtPriLoaded(iStn).Left = StnNumLeft + StnNumColWidth
            txtPriLoaded(iStn).Width = StnColWidth
            txtPriLoaded(iStn).text = ""
            txtPriLoaded(iStn).ToolTipText = "Primary Scale Job Start Weight in % of Working Capacity (0-115%)"
            txtPriLoaded(iStn).Visible = True
            
            Load txtAuxLoaded(iStn)
            txtAuxLoaded(iStn).Top = Stn1stRowTop + (StnRowHeight * (iStn - 1))
            txtAuxLoaded(iStn).Left = StnNumLeft + StnNumColWidth + StnColWidth + 90
            txtAuxLoaded(iStn).Width = StnColWidth
            txtAuxLoaded(iStn).text = ""
            txtAuxLoaded(iStn).ToolTipText = "Aux Scale Job Start Weight in % of Working Capacity (0-115%)"
            txtAuxLoaded(iStn).Visible = True
            
        Next iStn
    End If
    
    For iStn = 1 To LAST_STN
        txtLeakError(iStn).Visible = True
        Select Case STN_INFO(iStn).Type
            Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                txtNitMfcError(iStn).Visible = True
                txtButMfcError(iStn).Visible = True
                txtPrgMfcError(iStn).Visible = True
                txtLfvMfcError(iStn).Visible = False
            Case STN_ORVR2_TYPE
                txtNitMfcError(iStn).Visible = True
                txtButMfcError(iStn).Visible = True
                txtPrgMfcError(iStn).Visible = True
                txtLfvMfcError(iStn).Visible = False
            Case STN_LIVEFUEL_TYPE
                txtNitMfcError(iStn).Visible = False
                txtButMfcError(iStn).Visible = False
                txtPrgMfcError(iStn).Visible = True
                txtLfvMfcError(iStn).Visible = True
            Case STN_LIVEREG_TYPE
                txtNitMfcError(iStn).Visible = True
                txtButMfcError(iStn).Visible = True
                txtPrgMfcError(iStn).Visible = True
                txtLfvMfcError(iStn).Visible = True
            Case STN_LIVEORVR2_TYPE
                txtNitMfcError(iStn).Visible = True
                txtButMfcError(iStn).Visible = True
                txtPrgMfcError(iStn).Visible = True
                txtLfvMfcError(iStn).Visible = True
            Case STN_COMBO3_TYPE
                txtNitMfcError(iStn).Visible = False
                txtButMfcError(iStn).Visible = False
                txtPrgMfcError(iStn).Visible = False
                txtLfvMfcError(iStn).Visible = False
            Case Else
                txtNitMfcError(iStn).Visible = False
                txtButMfcError(iStn).Visible = False
                txtPrgMfcError(iStn).Visible = False
                txtLfvMfcError(iStn).Visible = False
        End Select
    Next iStn

    tmrUpdate.Interval = 250
    tmrUpdate.Enabled = True
    UpdateScreen

End Sub

Private Sub UpdateScreen()
    chkSimulation.Value = IIf(USINGSIMULATION, 1, 0)
    chkSimulationNoise.Value = IIf(USINGSIMNOISE, 1, 0)
    txtPasError(pasTEMPERATURE).text = Format(Sim_PasError(pasTEMPERATURE), "##0.00")
    txtPasError(pasMOISTURE).text = Format(Sim_PasError(pasMOISTURE), "##0.00")
    For iStn = 1 To LAST_STN
        txtLoadPressure(iStn).text = Format(Stn_AIO(iStn, asLoadPressure).EUValue, "##0.00")
        txtLeakError(iStn).text = Format(Sim_LeakError(iStn), "##0.00")
        txtNitMfcError(iStn).text = Format(Sim_MfcError(iStn, MFCNITROGEN), "##0.00")
        txtButMfcError(iStn).text = Format(Sim_MfcError(iStn, MFCBUTANE), "##0.00")
        txtPrgMfcError(iStn).text = Format(Sim_MfcError(iStn, MFCPURGEAIR), "##0.00")
        txtLfvMfcError(iStn).text = Format(Sim_MfcError(iStn, MFCLIVEFUEL), "##0.00")
        txtAuxLoaded(iStn).text = Format(Sim_AuxCan_JobStartPercentLoaded(iStn), "###0.0")
        txtPriLoaded(iStn).text = Format(Sim_PriCan_JobStartPercentLoaded(iStn), "###0.0")
    Next iStn
    txtLfDensity.text = Format(Sim_LiveFuelDensity, "#0.000#")
    UpdateAlarmButtons
End Sub

Private Sub UpdateAlarmButtons()
    pnlAlmIn(0).BackColor = IIf(Com_DIO(icEStopSw).Value, DKGREEN, MEDRED)
    pnlAlmIn(0).ToolTipText = IIf(Com_DIO(icEStopSw).Value, "ESTOP OK", "ESTOP Pressed")
    pnlAlmIn(1).BackColor = IIf(Com_DIO(icExhaustFlowFS).Value, DKGREEN, MEDRED)
    pnlAlmIn(1).ToolTipText = IIf(Com_DIO(icExhaustFlowFS).Value, "Exhaust Flow OK", "Loss of Exhaust Flow")
    If Alm_Doors Then
        pnlAlmIn(2).BackColor = MEDRED
        pnlAlmIn(2).ToolTipText = "Door Open Alarm"
    ElseIf Not Com_DIO(icDoorSw).Value And USINGDOOROPEN Then
        pnlAlmIn(2).BackColor = MEDYELLOW
        pnlAlmIn(2).ToolTipText = "One or more Doors are Open"
    Else
        pnlAlmIn(2).BackColor = DKGREEN
        pnlAlmIn(2).ToolTipText = "Doors Closed"
    End If
    pnlAlmIn(3).BackColor = IIf(Com_DIO(ic20LelGasSw).Value, DKGREEN, MEDRED)
    pnlAlmIn(3).ToolTipText = IIf(Com_DIO(ic20LelGasSw).Value, "LEL20 OK", "LEL20 Alarm")
    pnlAlmIn(4).BackColor = IIf(Com_DIO(icPurgeReadyIn).Value, DKGREEN, MEDRED)
    pnlAlmIn(4).ToolTipText = IIf(Com_DIO(icPurgeReadyIn).Value, "Ready", "Not Ready")
    If USINGUPS Then
        pnlAlmIn(5).BackColor = IIf(Com_DIO(icUpsActiveSw).Value, MEDRED, DKGREEN)
        pnlAlmIn(5).ToolTipText = IIf(Com_DIO(icUpsActiveSw).Value, "UPS Fault", "OK")
    Else
        pnlAlmIn(5).BackColor = LTGRAY
        pnlAlmIn(5).ToolTipText = "UPS not being used"
    End If
    If USING_EXT_CONTACTS Then
        pnlAlmIn(6).BackColor = IIf(Com_DIO(icExtAlmContactSw).Value, MEDRED, DKGREEN)
        pnlAlmIn(6).ToolTipText = IIf(Com_DIO(icExtAlmContactSw).Value, "Alarm", "OK")
    Else
        pnlAlmIn(6).BackColor = LTGRAY
        pnlAlmIn(6).ToolTipText = "External Alarm not in use"
    End If
    If USINGPASLOCALCONTROL Then
        pnlAlmIn(7).BackColor = IIf(Com_DIO(icPASisRunningIn).Value, DKGREEN, MEDRED)
        pnlAlmIn(7).ToolTipText = IIf(Com_DIO(icPASisRunningIn).Value, "Run Local PAS in Auto", "Dont Run Local PAS")
    Else
        pnlAlmIn(7).BackColor = LTGRAY
        pnlAlmIn(7).ToolTipText = "Local PAS Control not in use"
    End If
End Sub

Private Sub tbrSimConPnl_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        ' Save Settings
        USINGSIMULATION = IIf(chkSimulation.Value = cYES, True, False)
        USINGSIMNOISE = IIf(chkSimulationNoise.Value = cYES, True, False)
        Sim_PasError(pasTEMPERATURE) = CSng(txtPasError(pasTEMPERATURE).text)
        Sim_PasError(pasMOISTURE) = CSng(txtPasError(pasMOISTURE).text)
        For iStn = 1 To LAST_STN
            Stn_AIO(iStn, asLoadPressure).EUValue = CSng(txtLoadPressure(iStn).text)
            Sim_LeakError(iStn) = CSng(txtLeakError(iStn).text)
            Sim_MfcError(iStn, MFCNITROGEN) = CSng(txtNitMfcError(iStn).text)
            Sim_MfcError(iStn, MFCBUTANE) = CSng(txtButMfcError(iStn).text)
            Sim_MfcError(iStn, MFCPURGEAIR) = CSng(txtPrgMfcError(iStn).text)
            Sim_MfcError(iStn, MFCLIVEFUEL) = CSng(txtLfvMfcError(iStn).text)
            Sim_AuxCan_JobStartPercentLoaded(iStn) = CSng(txtAuxLoaded(iStn).text)
            Sim_PriCan_JobStartPercentLoaded(iStn) = CSng(txtPriLoaded(iStn).text)
        Next iStn
        Sim_LiveFuelDensity = CSng(txtLfDensity.text)
        Save_Simulation
        Save_SysDef
        lblMessages.ForeColor = Message_ForeColor
        lblMessages.Caption = "Simulation Setup Saved"
        UpdateScreen
    Case 3
        ' Close Screen
        Unload Me
        Set frmSimCntrlPnl = Nothing
End Select
End Sub

Private Sub tmrUpdate_Timer()
    UpdateAlarmButtons
End Sub

Private Sub txtButMfcError_Change(Index As Integer)
    lblMessages.Caption = ""
End Sub

Private Sub txtLeakError_Change(Index As Integer)
    lblMessages.Caption = ""
    If IsNumeric(txtLeakError(Index).text) Then
        If CSng(txtLeakError(Index).text) < 0 Then
            txtLeakError(Index).text = Format(Abs(CSng(txtLeakError(Index).text)), "##0.0")
        End If
    End If
End Sub

Private Sub txtLfvMfcError_Change(Index As Integer)
    lblMessages.Caption = ""
End Sub

Private Sub txtNitMfcError_Change(Index As Integer)
    lblMessages.Caption = ""
End Sub

Private Sub txtPrgMfcError_Change(Index As Integer)
    lblMessages.Caption = ""
End Sub

Private Sub txtAuxLoaded_Change(Index As Integer)
    lblMessages.Caption = ""
End Sub

Private Sub txtPriLoaded_Change(Index As Integer)
    lblMessages.Caption = ""
End Sub


