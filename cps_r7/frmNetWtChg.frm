VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmNetWtChg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NetWeightChange"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmNetWtChg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScreen 
      Interval        =   150
      Left            =   4920
      Top             =   840
   End
   Begin Threed.SSPanel pnlNetWtChg 
      Height          =   3090
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Current Time Remaining"
      Top             =   30
      Width           =   5230
      _Version        =   65536
      _ExtentX        =   9225
      _ExtentY        =   5450
      _StockProps     =   15
      Caption         =   "Net Load Weight"
      ForeColor       =   16512
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   8
      Begin VB.CommandButton cmdDebug 
         Height          =   315
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmNetWtChg.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Open WtChg debug screen"
         Top             =   2700
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblPrevAvg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prev Avg"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2480
         TabIndex        =   40
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   0
         Left            =   2480
         TabIndex        =   39
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   2480
         TabIndex        =   38
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   2
         Left            =   2480
         TabIndex        =   37
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   3
         Left            =   2480
         TabIndex        =   36
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   4
         Left            =   2480
         TabIndex        =   35
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   5
         Left            =   2480
         TabIndex        =   34
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   6
         Left            =   2480
         TabIndex        =   33
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   7
         Left            =   2480
         TabIndex        =   32
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label PrevAvg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   8
         Left            =   2480
         TabIndex        =   31
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   8
         Left            =   3720
         TabIndex        =   30
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   8
         Left            =   1240
         TabIndex        =   28
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   7
         Left            =   3720
         TabIndex        =   27
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   7
         Left            =   1240
         TabIndex        =   25
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   6
         Left            =   3720
         TabIndex        =   24
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   5
         Left            =   3720
         TabIndex        =   23
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   4
         Left            =   3720
         TabIndex        =   22
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   3
         Left            =   3720
         TabIndex        =   21
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   2
         Left            =   3720
         TabIndex        =   20
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   3720
         TabIndex        =   19
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   0
         Left            =   3720
         TabIndex        =   18
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "% Variation"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3720
         TabIndex        =   17
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   6
         Left            =   1240
         TabIndex        =   16
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   5
         Left            =   1240
         TabIndex        =   14
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   4
         Left            =   1240
         TabIndex        =   12
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   3
         Left            =   1240
         TabIndex        =   10
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   2
         Left            =   1240
         TabIndex        =   8
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   1240
         TabIndex        =   6
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   960
      End
      Begin VB.Label WtChg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##.##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   0
         Left            =   1240
         TabIndex        =   4
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label CycleNum 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblWtChg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Grams"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1240
         TabIndex        =   2
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label lblCycleNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmNetWtChg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''' form NetWtChg
' error module 2467
Option Explicit
'
Private idx As Integer
Private iCyc As Integer
Private idx2 As Integer
Private iCyc2 As Integer
Private ActTol As Single
Private AvgChg As Single
Private CurChg As Single
Private SumChg As Single

Sub ClearScreenOpenFlags()
Dim iStn As Integer
Dim iShft As Integer

    For iStn = 1 To NR_STN
        For iShft = 1 To NR_SHIFT
            LoadControl(iStn, iShft).NetWtChgIsOpen = False
        Next iShft
    Next iStn

End Sub

Sub UnloadScreen()
    ClearScreenOpenFlags
    Unload Me
    Set frmNetWtChg = Nothing
End Sub

Sub UpdateScreen()

    ' **************************************************************************************
    ' Net Weight Change Display
    ' **************************************************************************************
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2467, 13

    If (StationRecipe(DispStn, DispShift).EndMethod = ENDWEIGHTCHG) Then
        If StationControl(DispStn, DispShift).Mode <> VBIDLE Then
            pnlNetWtChg.Caption = "Net Primary Load Weight"
            cmdDebug.Visible = IIf(((Not NotDebugSCALES) Or (Not NotDebugPURGE) Or (Not NotDebugMMW)), True, False)
            For idx = CycleNum.LBound To CycleNum.UBound
                iCyc = StationControl(DispStn, DispShift).CurrCycle - idx
                Select Case iCyc
                    Case Is < 1
                        CycleNum(idx).Visible = False
                        WtChg(idx).Visible = False
                        PrevAvg(idx).Visible = False
                        Percent(idx).Visible = False
                    Case Is <= StationRecipe(DispStn, DispShift).EndConsecutiveCycles
                        CycleNum(idx).Visible = True
                        CycleNum(idx).Caption = Format(iCyc, "##0")
                        If idx = 0 Then
                            ' current cycle
                            Select Case StationControl(DispStn, DispShift).Mode
                                Case VBLOAD
                                    CurChg = StationControl(DispStn, DispShift).PriScaleWt - LoadControl(DispStn, DispShift).PriWt_Start
                                Case Else
                                    CurChg = StationCycleWeightData(DispStn, DispShift, iCyc).Load_EndWeight_Pri - StationCycleWeightData(DispStn, DispShift, iCyc).Load_StartWeight_Pri
                            End Select
                            WtChg(idx).ForeColor = DataHiLite_ForeColor
                            CycleNum(idx).ForeColor = DataHiLite_ForeColor
                        Else
                            CurChg = StationCycleWeightData(DispStn, DispShift, iCyc).Load_EndWeight_Pri - StationCycleWeightData(DispStn, DispShift, iCyc).Load_StartWeight_Pri
                            CycleNum(idx).ForeColor = BarActual_ForeColor
                            WtChg(idx).ForeColor = BarActual_ForeColor
                        End If
                        WtChg(idx).Visible = True
                        WtChg(idx).Caption = Format(CurChg, "##0.000")
                        PrevAvg(idx).Visible = False
                        Percent(idx).Visible = False
                    Case Else
                        CycleNum(idx).Visible = True
                        CycleNum(idx).Caption = Format(iCyc, "##0")
                        If idx = 0 Then
                            ' current cycle
                            Select Case StationControl(DispStn, DispShift).Mode
                                Case VBLOAD
                                    CurChg = StationControl(DispStn, DispShift).PriScaleWt - LoadControl(DispStn, DispShift).PriWt_Start
                                Case Else
                                    CurChg = StationCycleWeightData(DispStn, DispShift, iCyc).Load_EndWeight_Pri - StationCycleWeightData(DispStn, DispShift, iCyc).Load_StartWeight_Pri
                            End Select
                            WtChg(idx).ForeColor = DataHiLite_ForeColor
                            CycleNum(idx).ForeColor = DataHiLite_ForeColor
                        Else
                            CurChg = StationCycleWeightData(DispStn, DispShift, iCyc).Load_EndWeight_Pri - StationCycleWeightData(DispStn, DispShift, iCyc).Load_StartWeight_Pri
                            CycleNum(idx).ForeColor = BarActual_ForeColor
                            WtChg(idx).ForeColor = BarActual_ForeColor
                        End If
                        WtChg(idx).Visible = True
                        WtChg(idx).Caption = Format(CurChg, "##0.000")
                        SumChg = CSng(0)
                        For idx2 = 1 To StationRecipe(DispStn, DispShift).EndConsecutiveCycles
                            iCyc2 = iCyc - idx2
                            SumChg = SumChg + (StationCycleWeightData(DispStn, DispShift, iCyc2).Load_EndWeight_Pri - StationCycleWeightData(DispStn, DispShift, iCyc2).Load_StartWeight_Pri)
                        Next idx2
                        AvgChg = SumChg / StationRecipe(DispStn, DispShift).EndConsecutiveCycles
                        PrevAvg(idx).Visible = True
                        PrevAvg(idx).Caption = Format(AvgChg, "####0.000")
                        PrevAvg(idx).ForeColor = BarActual_ForeColor
                        ' do not divide by zero
                        If (CurChg > CSng(0)) Then
                            ActTol = CSng(100) * Abs((CurChg - AvgChg) / CurChg)
                            Percent(idx).Visible = True
                            Percent(idx).Caption = Format(ActTol, "####0.00")
                            If (ActTol <= Abs(StationRecipe(DispStn, DispShift).EndWeightTolerance)) Then
                                Percent(idx).ForeColor = Good_ForeColor
                            Else
                                Percent(idx).ForeColor = Warning_ForeColor
                            End If
                        Else
                            Percent(idx).Visible = False
                        End If
                End Select
            Next idx
        Else
            ' close screen
            UnloadScreen
        End If
    Else
        ' close screen
        UnloadScreen
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

Private Sub cmdDebug_Click()
    frmDebugWtChg.Show
End Sub

Private Sub Form_Activate()
    KeyPreview = True
    LoadControl(DispStn, DispShift).NetWtChgIsOpen = True
End Sub

Private Sub Form_Deactivate()
'    ClearScreenOpenFlags
End Sub

Private Sub Form_Load()
'
    KeyPreview = True
    LoadControl(DispStn, DispShift).NetWtChgIsOpen = True
    
    ' Set Foreground colors
    pnlNetWtChg.ForeColor = TitlesLabel_ForeColor
    lblCycleNum.ForeColor = TitlesLabel_ForeColor
    lblWtChg.ForeColor = TitlesLabel_ForeColor
    lblPrevAvg.ForeColor = TitlesLabel_ForeColor
    lblPercent.ForeColor = TitlesLabel_ForeColor
    For idx = CycleNum.LBound To CycleNum.UBound
        CycleNum(idx).ForeColor = Data_ForeColor
        WtChg(idx).ForeColor = Data_ForeColor
        PrevAvg(idx).ForeColor = Data_ForeColor
        Percent(idx).ForeColor = Data_ForeColor
    Next idx
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClearScreenOpenFlags
End Sub

Private Sub tmrScreen_Timer()
    UpdateScreen
End Sub
