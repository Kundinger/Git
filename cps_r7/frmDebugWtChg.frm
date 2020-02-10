VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmDebugWtChg 
   Caption         =   "DebugWtChg"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   Icon            =   "frmDebugWtChg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScreen 
      Interval        =   150
      Left            =   5040
      Top             =   0
   End
   Begin Threed.SSPanel pnlNetWtChg 
      Height          =   3570
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "Current Time Remaining"
      Top             =   30
      Width           =   7125
      _Version        =   65536
      _ExtentX        =   12568
      _ExtentY        =   6297
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
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   44
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   43
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   42
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   41
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   40
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   39
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   38
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   37
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label AvgChg 
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
         Left            =   3780
         TabIndex        =   36
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblAvgChg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Avg Chg"
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
         Left            =   3600
         TabIndex        =   35
         Top             =   540
         Width           =   1560
      End
      Begin VB.Label CurAvgChg 
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
         Left            =   3810
         TabIndex        =   34
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label CurCycleNum 
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
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label CurWtChg 
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
         Left            =   2025
         TabIndex        =   32
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label CurPercent 
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
         Left            =   5430
         TabIndex        =   31
         Top             =   240
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
         Left            =   150
         TabIndex        =   30
         Top             =   540
         Width           =   1560
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
         Left            =   1815
         TabIndex        =   29
         Top             =   540
         Width           =   1560
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
         Left            =   330
         TabIndex        =   28
         Top             =   840
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
         Index           =   0
         Left            =   1995
         TabIndex        =   27
         Top             =   840
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
         Left            =   330
         TabIndex        =   26
         Top             =   1080
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
         Index           =   1
         Left            =   1995
         TabIndex        =   25
         Top             =   1080
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
         Left            =   330
         TabIndex        =   24
         Top             =   1320
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
         Index           =   2
         Left            =   1995
         TabIndex        =   23
         Top             =   1320
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
         Left            =   330
         TabIndex        =   22
         Top             =   1560
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
         Index           =   3
         Left            =   1995
         TabIndex        =   21
         Top             =   1560
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
         Left            =   330
         TabIndex        =   20
         Top             =   1800
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
         Index           =   4
         Left            =   1995
         TabIndex        =   19
         Top             =   1800
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
         Left            =   330
         TabIndex        =   18
         Top             =   2040
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
         Index           =   5
         Left            =   1995
         TabIndex        =   17
         Top             =   2040
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
         Left            =   330
         TabIndex        =   16
         Top             =   2280
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
         Left            =   1995
         TabIndex        =   15
         Top             =   2280
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
         Left            =   5220
         TabIndex        =   14
         Top             =   540
         Width           =   1560
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
         Left            =   5400
         TabIndex        =   13
         Top             =   840
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
         Left            =   5400
         TabIndex        =   12
         Top             =   1080
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
         Left            =   5400
         TabIndex        =   11
         Top             =   1320
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
         Left            =   5400
         TabIndex        =   10
         Top             =   1560
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
         Left            =   5400
         TabIndex        =   9
         Top             =   1800
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
         Left            =   5400
         TabIndex        =   8
         Top             =   2040
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
         Left            =   5400
         TabIndex        =   7
         Top             =   2280
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
         Index           =   7
         Left            =   1995
         TabIndex        =   6
         Top             =   2520
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
         Left            =   330
         TabIndex        =   5
         Top             =   2520
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
         Left            =   5400
         TabIndex        =   4
         Top             =   2520
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
         Index           =   8
         Left            =   1995
         TabIndex        =   3
         Top             =   2760
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
         Left            =   330
         TabIndex        =   2
         Top             =   2760
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
         Left            =   5400
         TabIndex        =   1
         Top             =   2760
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmDebugWtChg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''' form DebugWtChg
' error module 7467
Option Explicit
'
Private idx As Integer
Private iCyc As Integer
Private idx2 As Integer
Private iCyc2 As Integer
Private ActTol As Single
Private AvChg As Single
Private CurChg As Single
Private SumChg As Single

Sub UnloadScreen()
    Unload Me
    Set frmDebugWtChg = Nothing
End Sub

Sub UpdateScreen()
    Dim iStn As Integer
    Dim iShift As Integer
    iStn = DispStn
    iShift = DispShift
    iCyc = IIf((StationControl(iStn, iShift).Mode = VBLOAD), (StationControl(iStn, iShift).CurrCycle), (1 + StationControl(iStn, iShift).CurrCycle))
    CurCycleNum.Caption = Format(iCyc, "###0")
    CurChg = LoadControl(iStn, iShift).loadTotalGrams
    CurWtChg.Caption = Format(CurChg, "##,##0.0000")
    If iCyc > 3 Then
        SumChg = StationCycleWeightData(iStn, iShift, iCyc - 1).Load_TotalGrams
        SumChg = SumChg + StationCycleWeightData(iStn, iShift, iCyc - 2).Load_TotalGrams
        SumChg = SumChg + StationCycleWeightData(iStn, iShift, iCyc - 3).Load_TotalGrams
        AvChg = SumChg / 3
        CurAvgChg.Caption = Format(AvChg, "##,##0.0000")
        If CurChg <> 0 Then ActTol = 100 * ((CurChg - AvChg) / CurChg)
        CurPercent.Caption = IIf((CurChg <> 0), Format(ActTol, "###0.000"), "---")
    Else
        CurAvgChg.Caption = ""
        CurPercent.Caption = ""
    End If

    For idx = 1 To 9
        idx2 = idx - 1
        iCyc = IIf((StationControl(iStn, iShift).Mode = VBLOAD), (StationControl(iStn, iShift).CompletedCycles - idx2), (StationControl(iStn, iShift).CompletedCycles - idx + 2))
        If iCyc > 0 Then
            CycleNum(idx2).Left = CurCycleNum.Left
            WtChg(idx2).Left = CurWtChg.Left
            AvgChg(idx2).Left = CurAvgChg.Left
            Percent(idx2).Left = CurPercent.Left
            CycleNum(idx2).Caption = Format(iCyc, "###0")
            CurChg = StationCycleWeightData(iStn, iShift, iCyc).Load_TotalGrams
            WtChg(idx2).Caption = Format(CurChg, "##,##0.0000")
            If iCyc > 3 Then
                SumChg = StationCycleWeightData(iStn, iShift, iCyc - 1).Load_TotalGrams
                SumChg = SumChg + StationCycleWeightData(iStn, iShift, iCyc - 2).Load_TotalGrams
                SumChg = SumChg + StationCycleWeightData(iStn, iShift, iCyc - 3).Load_TotalGrams
                AvChg = SumChg / 3
                AvgChg(idx2).Caption = Format(AvChg, "##,##0.0000")
                If (CurChg <> 0) Then ActTol = 100 * ((CurChg - AvChg) / CurChg)
                Percent(idx2).Caption = IIf((CurChg <> 0), Format(ActTol, "###0.000"), "---")
            Else
                AvgChg(idx2).Caption = ""
                Percent(idx2).Caption = ""
            End If
        Else
            CycleNum(idx2).Left = OutOfSight
            WtChg(idx2).Left = OutOfSight
            AvgChg(idx2).Left = OutOfSight
            Percent(idx2).Left = OutOfSight
        End If
    Next idx

End Sub

Private Sub Form_Load()
'
    KeyPreview = True
    
    ' Set Foreground colors
    pnlNetWtChg.ForeColor = TitlesLabel_ForeColor
    lblCycleNum.ForeColor = TitlesLabel_ForeColor
    lblWtChg.ForeColor = TitlesLabel_ForeColor
    lblAvgChg.ForeColor = TitlesLabel_ForeColor
    lblPercent.ForeColor = TitlesLabel_ForeColor
    For idx = CycleNum.LBound To CycleNum.UBound
        CycleNum(idx).ForeColor = Data_ForeColor
        WtChg(idx).ForeColor = Data_ForeColor
        AvgChg(idx).ForeColor = Data_ForeColor
        Percent(idx).ForeColor = Data_ForeColor
    Next idx
    tmrScreen.Interval = 250
End Sub

Private Sub tmrScreen_Timer()
    UpdateScreen
End Sub

