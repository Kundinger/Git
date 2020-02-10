VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmCalCheckHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CalCheck History"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7530
   Icon            =   "frmMfcCalCheckHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   8760
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Frame frmCalChkHst 
      Height          =   10140
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   7320
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Close"
         DisabledPicture =   "frmMfcCalCheckHistory.frx":57E2
         DownPicture     =   "frmMfcCalCheckHistory.frx":5EE4
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMfcCalCheckHistory.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Return to Previous Screen"
         Top             =   9150
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmMfcCalCheckHistory.frx":6CE8
         DownPicture     =   "frmMfcCalCheckHistory.frx":792A
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMfcCalCheckHistory.frx":856C
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Print the displayed retention Test"
         Top             =   9150
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Frame frmCalCheckSelection 
         Caption         =   "CalCheck Selection"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   78
         Top             =   3780
         Width           =   7095
         Begin VB.CommandButton cmdCalCheckUp 
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
            Left            =   1765
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMfcCalCheckHistory.frx":91AE
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "next mfc"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   640
         End
         Begin VB.CommandButton cmdCalCheckDn 
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
            Picture         =   "frmMfcCalCheckHistory.frx":98B0
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "previous mfc"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   640
         End
         Begin Threed.SSPanel pnlCalCheck 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   645
            Left            =   765
            TabIndex        =   81
            ToolTipText     =   "Click for list of Defined Recipes"
            Top             =   360
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   1138
            _StockProps     =   15
            Caption         =   "49"
            ForeColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelInner      =   1
            Font3D          =   3
         End
         Begin VB.Label lblCurrent 
            Alignment       =   2  'Center
            Caption         =   "CalCheck DTS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   345
            Left            =   2520
            TabIndex        =   83
            Top             =   510
            Width           =   4365
         End
         Begin VB.Label lblCalCheckDtsMsg 
            Alignment       =   2  'Center
            Caption         =   "CalCheck 8 of 88"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   240
            Left            =   2520
            TabIndex        =   82
            Top             =   915
            Visible         =   0   'False
            Width           =   4365
         End
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
         Left            =   120
         TabIndex        =   74
         Top             =   120
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
            Left            =   870
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMfcCalCheckHistory.frx":9FB2
            Style           =   1  'Graphical
            TabIndex        =   77
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
            Picture         =   "frmMfcCalCheckHistory.frx":A6B4
            Style           =   1  'Graphical
            TabIndex        =   76
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
            ForeColor       =   &H00000040&
            Height          =   615
            Left            =   1605
            TabIndex        =   75
            Text            =   "Station 8"
            Top             =   300
            Width           =   5340
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
         Left            =   120
         TabIndex        =   64
         Top             =   1140
         Width           =   7065
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
            Picture         =   "frmMfcCalCheckHistory.frx":ADB6
            Style           =   1  'Graphical
            TabIndex        =   69
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
            Left            =   870
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMfcCalCheckHistory.frx":B4B8
            Style           =   1  'Graphical
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
            Text            =   "Function Description123456789012345678901234567890"
            ToolTipText     =   "Function Description"
            Top             =   540
            Width           =   4485
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
            TabIndex        =   65
            Text            =   "12345"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   540
            Width           =   1005
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
            TabIndex        =   73
            Top             =   300
            Width           =   3885
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
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblCurCalDts 
            Alignment       =   2  'Center
            Caption         =   "Current Calibration DTS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   345
            Left            =   1575
            TabIndex        =   70
            Top             =   1110
            Width           =   5235
         End
      End
      Begin VB.Frame frmCalCheckData 
         Caption         =   "CalCheck Data"
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
         Height          =   3975
         Left            =   1320
         TabIndex        =   3
         Top             =   5075
         Width           =   4695
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   2640
            TabIndex        =   25
            Top             =   3450
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   2640
            TabIndex        =   24
            Top             =   3165
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   2640
            TabIndex        =   23
            Top             =   2880
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   2640
            TabIndex        =   22
            Top             =   2595
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   2640
            TabIndex        =   21
            Top             =   2310
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   2640
            TabIndex        =   20
            Top             =   2025
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   2640
            TabIndex        =   19
            Top             =   1740
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   2640
            TabIndex        =   18
            Top             =   1455
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   2640
            TabIndex        =   17
            Top             =   1170
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   2640
            TabIndex        =   16
            Top             =   885
            Width           =   900
         End
         Begin VB.TextBox txtActual 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   15
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   14
            Top             =   1455
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   840
            TabIndex        =   13
            Top             =   3450
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   840
            TabIndex        =   12
            Top             =   3165
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   840
            TabIndex        =   11
            Top             =   2880
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   840
            TabIndex        =   10
            Top             =   2595
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   840
            TabIndex        =   9
            Top             =   2310
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   840
            TabIndex        =   8
            Top             =   2025
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   840
            TabIndex        =   7
            Top             =   1740
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   840
            TabIndex        =   6
            Top             =   1170
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   5
            Top             =   885
            Width           =   780
         End
         Begin VB.TextBox txtDesired 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   780
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   63
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   62
            Top             =   885
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   61
            Top             =   1170
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   60
            Top             =   1455
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   59
            Top             =   1740
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   58
            Top             =   2025
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   57
            Top             =   2310
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   56
            Top             =   2595
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   55
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   54
            Top             =   3165
            Width           =   495
         End
         Begin VB.Label lblPointNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   240
            TabIndex        =   53
            Top             =   3450
            Width           =   495
         End
         Begin VB.Label lblCurrPoints 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Point"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblDesireds 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desired"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   840
            TabIndex        =   51
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblActuals 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Actual"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2640
            TabIndex        =   50
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   49
            Top             =   3450
            Width           =   1020
         End
         Begin VB.Label lblCalibrateds 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Calibrated"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1620
            TabIndex        =   48
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   47
            Top             =   600
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   46
            Top             =   885
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   45
            Top             =   1170
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   44
            Top             =   1455
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   43
            Top             =   1740
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   42
            Top             =   2025
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   41
            Top             =   2310
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   40
            Top             =   2595
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   39
            Top             =   2880
            Width           =   1020
         End
         Begin VB.Label lblCalibrated 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   1620
            TabIndex        =   38
            Top             =   3165
            Width           =   1020
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   37
            Top             =   3450
            Width           =   900
         End
         Begin VB.Label lblPercDiffs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "% Diff"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3540
            TabIndex        =   36
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   35
            Top             =   600
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   34
            Top             =   885
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   33
            Top             =   1170
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   32
            Top             =   1455
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   31
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   30
            Top             =   2025
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   29
            Top             =   2310
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   28
            Top             =   2595
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   27
            Top             =   2880
            Width           =   900
         End
         Begin VB.Label lblPercDiff 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
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
            Left            =   3540
            TabIndex        =   26
            Top             =   3165
            Width           =   900
         End
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "messages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   855
         Left            =   1080
         TabIndex        =   86
         Top             =   9150
         Width           =   5145
      End
      Begin VB.Label lblMessage2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "messages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   2910
         Visible         =   0   'False
         Width           =   7065
      End
   End
End
Attribute VB_Name = "frmCalCheckHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 17359
'
Option Explicit

Const NUMROWS = 11                                  ' Number of calibration check rows
Private aryDesFlow(NUMROWS) As Single               ' the desired flow (from the input table)
Private aryActualFlowSLPM(NUMROWS) As Single        ' the actual flow in Calibrated SLPM (from the input table)
Private aryActualFlowUncal(NUMROWS) As Single       ' the actual flow in Uncalibrated SLPM (from the input table)
Private aryPercDiff(NUMROWS) As Single              ' the difference between desired and actual flow in percent (from the input table)
Private bChanged(NUMROWS) As Boolean                ' the actual flow has been changed since the last save (or load)
Private bAllChanged As Boolean                      ' all the actual flows have been changed since the last save (or load)

Private aryOutputFS(NUMROWS) As Single              ' the output value found using Newton's method

Public SelectedCalCheck As Integer                  ' the selected calcheck index
Public SelectedFunc As Integer                      ' the Station Analog Function index for the selected mass flow controller (from frmMassFlowCal)
Public SelectedMFC As Integer                       ' the selected mass flow controller (from frmMfcCalCheck)
Public SelectedRow As Integer                       ' the selected calibration data entry row
Public SelectedStation As Integer                   ' the selected station (from frmMfcCalCheck)
Private MfcSpan As Single                           ' the selected Mfc's span in Engr Units (max - min)
Private MfcMin As Single                            ' the selected Mfc's minimum value in Engr Units

Private AutoCycleOn As Boolean
Private AutoStepNext As Date
Private AutoStepInterval As Long
Private AutoCycleSP As Single

Private Curr_MfcCal As MfcCalibration
Private calDTS As Date
Private curDTS As Date
Private calchkDtsList(1 To MAXCALCHECKS) As Date
Private NumCalPoints As Integer
Private idxMax As Integer
Private bFormLoaded As Boolean           ' Flag - Whether the form has loaded(can't SetFocus on a TextBox until form is Loaded

' Max Station Index
Const MAXSTN = 9
' Min Station Index
Const MINSTN = 1
' Max MFCs per Station
Const MAXINP = MAXMFC

Private rsCrit As String
Private dbDbase As Database
Private rsTable  As Recordset


Private Sub cmdCalCheckDn_Click()
'
    SelectedCalCheck = IIf((SelectedCalCheck > 1), SelectedCalCheck - 1, MAXCALCHECKS)
    pnlCalCheck.Caption = Format(SelectedCalCheck, "##0")
    DisplayMfcCalCheckData
    UpdateCmdButtons
End Sub

Private Sub cmdCalCheckUp_Click()
'
    SelectedCalCheck = IIf((SelectedCalCheck < MAXCALCHECKS), SelectedCalCheck + 1, 1)
    pnlCalCheck.Caption = Format(SelectedCalCheck, "##0")
    DisplayMfcCalCheckData
    UpdateCmdButtons
End Sub

Private Sub cmdGroupDn_Click()
' This command decrements the station number,
' the displayed name of the station, and triggers an update for
' the values displayed on the form for the current mfc
Dim flag As Boolean
    flag = False
    Do While Not flag
        SelectedStation = SelectedStation - 1
        If SelectedStation < MINSTN Then SelectedStation = MAXSTN
        If SelectedStation <= LAST_STN Then flag = True
    Loop
    ' check for valid mfc in new station
    SelectedMFC = SelectedMFC - 1
    NextInput
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcSelection
    FindCalChecks
    If (SelectedCalCheck <> 0) Then DisplayMfcCalCheckData
    UpdateCmdButtons
    pnlCalCheck.Caption = Format(SelectedCalCheck, "##0")
End Sub

Private Sub cmdGroupUp_Click()
' This command increments the station number,
' the displayed name of the station, and triggers an update for
' the values displayed on the form for the current mfc
Dim flag As Boolean
    flag = False
    Do While Not flag
        SelectedStation = SelectedStation + 1
        If SelectedStation > MAXSTN Then SelectedStation = MINSTN
        If SelectedStation <= LAST_STN Then flag = True
    Loop
    ' check for valid mfc in new group
    SelectedMFC = SelectedMFC - 1
    NextInput
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcSelection
    FindCalChecks
    If (SelectedCalCheck <> 0) Then DisplayMfcCalCheckData
    UpdateCmdButtons
    pnlCalCheck.Caption = Format(SelectedCalCheck, "##0")
End Sub

Public Sub UpdateMfcSelection()
'
    Select Case SelectedMFC
        Case MFCBUTANE
            SelectedFunc = asButaneFlow
        Case MFCNITROGEN
            SelectedFunc = asNitrogenFlow
        Case MFCPURGEAIR
            SelectedFunc = asPurgeAirFlow
        Case MFCORVRBUT
            SelectedFunc = asButaneORVRFlow
        Case MFCORVRNIT
            SelectedFunc = asNitrogenORVRFlow
        Case MFCORVRPRG
            SelectedFunc = asPurgeAirFlow
        Case MFCLIVEFUEL
            SelectedFunc = asLiveFuelVaporFlow
        Case MFCORVRLIVE
            SelectedFunc = asLiveFuelVaporORVRFlow
    End Select
    Curr_MfcCal = Stn_MfcCal(SelectedStation, SelectedMFC)
    NumCalPoints = Curr_MfcCal.NumPoints
    Curr_MfcCal.RawInputType = CalRawAsEU
    calDTS = Curr_MfcCal.Dts
    SelectedCalCheck = 1
    pnlCalCheck.Caption = Format(SelectedCalCheck, "##0")
End Sub

Public Sub DisplayMfcAll()
    ' update the screen
    DisplayMfcSelection
    FindCalChecks
    If (SelectedCalCheck <> 0) Then DisplayMfcCalCheckData
    UpdateCmdButtons
End Sub

Private Sub DisplayMfcSelection()
' DisplayMfcSelection
' Displays Information on the Currently Selected Mass Flow Controller
Dim sGrpDesc As String
Dim sAiDesc As String
Dim sEuMax As Single
Dim sEuMin As Single
Dim sVdcMax As Single
Dim sVdcMin As Single
Dim iRawInputType As Integer
    
    sGrpDesc = "Station #" & Format(SelectedStation, "#0")
    sAiDesc = Stn_AnaDef(SelectedFunc).desc
    sEuMax = Stn_AIO(SelectedStation, SelectedFunc).EuMax
    sEuMin = Stn_AIO(SelectedStation, SelectedFunc).EuMin
    txtDispGrp.text = sGrpDesc
    txtaFuncDesc.text = sAiDesc
    txtaEUMax.text = Format(sEuMax, "####0.0##")
    txtaEUMin.text = Format(sEuMin, "####0.0##")
    
    lblCurCalDts.Caption = "Calibrated  " & Format(calDTS, "YYYY MMM DD  HH:MM:SS")
    
End Sub

Private Sub HideInactiveTableRows()
' Hide calcheck table rows
' in excess of the current number of points
Dim flag As Boolean
Dim iRow As Integer
    For iRow = 1 To MAXLSQCALPOINTS
        ' Visible or NOT ??
        flag = IIf((iRow <= NumCalPoints), True, False)
        ' Set every cell in the row
        lblPointNum(iRow).Visible = flag
        txtDesired(iRow).Visible = flag
        lblCalibrated(iRow).Visible = flag
        txtActual(iRow).Visible = flag
        lblPercDiff(iRow).Visible = flag
        lblPointNum(iRow).Enabled = flag
        txtDesired(iRow).Enabled = flag
        lblCalibrated(iRow).Enabled = flag
        txtActual(iRow).Enabled = flag
        lblPercDiff(iRow).Enabled = flag
    Next iRow
End Sub

Private Sub DisplayMfcCalCheckData()
' Displays calcheck point data
'
Dim iPoint As Integer
Dim percDiff As Single

    HideInactiveTableRows
    
    ' open data table MfcCalCheck
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    rsCrit = "SELECT * FROM [MfcCalCheckData] "
    rsCrit = rsCrit & "WHERE [Station] = " & SelectedStation & " "
    rsCrit = rsCrit & "AND [Mfc] = " & SelectedMFC & " "
    rsCrit = rsCrit & "AND [CalDTS] = #" & calDTS & "#  "
    rsCrit = rsCrit & "AND [CalCheckDTS] = #" & calchkDtsList(SelectedCalCheck) & "#  "
    rsCrit = rsCrit & " ORDER BY [MfcCalCheckData].[Point] DESC"
    
    ' open recordset
    Set rsTable = dbDbase.OpenRecordset(rsCrit, dbOpenDynaset)
    If rsTable.BOF Then
        ' no data
        If (SelectedCalCheck > idxMax) Then
            lblMessage2.ForeColor = Message_ForeColor
            lblMessage2.Caption = vbCrLf & "No more CalChecks"
            lblMessage2.Visible = True
            ' calcheck info
            lblCurrent.Caption = " "
            lblCalCheckDtsMsg.Caption = " "
            lblCalCheckDtsMsg.Visible = False
        Else
            lblMessage2.ForeColor = Message_ForeColor
            lblMessage2.Caption = vbCrLf & "No CalCheck data for the " & Format(calchkDtsList(SelectedCalCheck), "YYYY MMM DD  hh:mm:ss") & " CalCheck"
            lblMessage2.Visible = True
            ' calcheck info
            lblCurrent.Caption = "" & Format(calchkDtsList(SelectedCalCheck), "YYYY MMM DD  hh:mm:ss")
            lblCalCheckDtsMsg.Caption = "CalCheck #" & Format(SelectedCalCheck, "##0") & " of " & Format(rsTable.RecordCount, "##0")
            lblCalCheckDtsMsg.Visible = True
        End If
    Else
        ' at least one data record
        lblMessage2.Visible = False
        rsTable.MoveFirst
        rsTable.MoveLast
        Do While Not rsTable.BOF
            ' display calcheck data
            iPoint = rsTable("Point")
            lblPointNum(iPoint).Caption = Format(iPoint, "#0")
            txtDesired(iPoint).text = IIf(txtDesired(iPoint).Enabled, Format(rsTable("FlowSP"), "####0.0##"), "")
            lblCalibrated(iPoint).Caption = IIf(lblCalibrated(iPoint).Enabled, Format(rsTable("FlowPV"), "####0.0##"), "")
            txtActual(iPoint).text = IIf(txtActual(iPoint).Enabled, Format(rsTable("CalCheckFlow"), "####0.0##"), "")
            percDiff = CSng(100) * ((rsTable("CalCheckFlow") - rsTable("FlowPV")) / (Stn_AIO(SelectedStation, SelectedFunc).EuMax - Stn_AIO(SelectedStation, SelectedFunc).EuMin))
            lblPercDiff(iPoint).Caption = IIf(lblPercDiff(iPoint).Enabled, Format(percDiff, "####0.0##"), "")
            rsTable.MovePrevious
        Loop
        ' calcheck info
        lblCurrent.Caption = "" & Format(calchkDtsList(SelectedCalCheck), "YYYY MMM DD  hh:mm:ss")
        lblCalCheckDtsMsg.Caption = "CalCheck #" & Format(SelectedCalCheck, "##0") & " of " & Format(idxMax, "##0")
        lblCalCheckDtsMsg.Visible = True
    End If
        
    
    ' close db
    rsTable.Close
    dbDbase.Close

End Sub

Private Sub UpdateCmdButtons()
' update the calcheck command buttons
'
    cmdGroupDn.Enabled = IIf((SelectedStation > MINSTN), True, False)
    cmdGroupUp.Enabled = IIf((SelectedStation < LAST_STN), True, False)
    cmdInputDn.Enabled = IIf((SelectedMFC > 0), True, False)
    cmdInputUp.Enabled = IIf((SelectedMFC < MAXMFC), True, False)
    cmdCalCheckDn.Enabled = IIf((SelectedCalCheck > 1), True, False)
    cmdCalCheckUp.Enabled = IIf(((SelectedCalCheck < idxMax) And (SelectedCalCheck > 0)), True, False)
    cmdPrint.Enabled = IIf(PRINTERAVAILABLE, True, False)
    
End Sub

Private Sub cmdInputDn_Click()
'
    PrevInput
End Sub

Private Sub cmdInputUp_Click()
'
    NextInput
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
        SelectedMFC = SelectedMFC - 1
        If SelectedMFC < 0 Then SelectedMFC = iMax
        ' get the station analog function index for the selected MFC
        Select Case SelectedMFC
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
        iAddr = Stn_AIO(SelectedStation, iFunc).addr
        iChan = Stn_AIO(SelectedStation, iFunc).chan
        ' Addr=0 AND Chan=0 is an undefined input
        ' MFC inputs and outputs are not done here
        If ((iAddr <> 0) Or (iChan <> 0)) Then
            Select Case SelectedMFC
                Case MFCBUTANE
                    If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCNITROGEN
                    If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCPURGEAIR
                    If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRBUT
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRNIT
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRPRG
                    ' not used
                Case MFCLIVEFUEL
                    If STN_INFO(SelectedStation).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRLIVE
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
            End Select
        End If
        iCntr = iCntr + 1
        If iCntr > iMax Then calneeded = True
    Loop
'    bUnsavedCal = False
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
        SelectedMFC = SelectedMFC + 1
        If SelectedMFC > iMax Then SelectedMFC = 0
        ' get the station analog function index for the selected MFC
        Select Case SelectedMFC
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
        iAddr = Stn_AIO(SelectedStation, iFunc).addr
        iChan = Stn_AIO(SelectedStation, iFunc).chan
        ' Addr=0 AND Chan=0 is an undefined input
        ' MFC inputs and outputs are not done here
        If ((iAddr <> 0) Or (iChan <> 0)) Then
            Select Case SelectedMFC
                Case MFCBUTANE
                    If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCNITROGEN
                    If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCPURGEAIR
                    If STN_INFO(SelectedStation).Type = STN_REGULAR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRBUT
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRNIT
                    If STN_INFO(SelectedStation).Type = STN_ORVR2_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRPRG
                    ' not used
                Case MFCLIVEFUEL
                    If STN_INFO(SelectedStation).Type = STN_LIVEFUEL_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEREG_TYPE Then calneeded = True
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
                Case MFCORVRLIVE
                    If STN_INFO(SelectedStation).Type = STN_LIVEORVR2_TYPE Then calneeded = True
            End Select
        End If
        iCntr = iCntr + 1
        If iCntr > iMax Then calneeded = True
    Loop
'    bUnsavedCal = False
    ' Update the Display
    UpdateMfcSelection
    DisplayMfcAll
End Sub

Private Sub FindCalChecks()
''
''
Dim idx As Integer
Dim currDTS As Date
Dim lastDTS As Date

    ' clear list of CalCheck DTS's
    For idx = 1 To MAXCALCHECKS
        calchkDtsList(idx) = 0
    Next idx
    ' open database MfcCalCheck
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    rsCrit = "SELECT * FROM [MfcCalCheckData] "
    rsCrit = rsCrit & "WHERE [Station] = " & SelectedStation & " "
    rsCrit = rsCrit & "AND [Mfc] = " & SelectedMFC & " "
    rsCrit = rsCrit & "AND [CalDTS] = #" & calDTS & "#  "
    rsCrit = rsCrit & " ORDER BY [MfcCalCheckData].[CalCheckDTS] ASC"
    Set rsTable = dbDbase.OpenRecordset(rsCrit, dbOpenDynaset)
    If rsTable.BOF Then
        ' no records
        frmCalCheckData.Visible = False
        frmCalCheckSelection.Visible = False
        cmdPrint.Visible = False
        lblMessage2.ForeColor = Message_ForeColor
        lblMessage2.Caption = vbCrLf & "NO CalChecks for the Station #" & Format(SelectedStation, "#0") & " " & Mfc_Description(SelectedMFC) & " MFC"
        lblMessage2.Visible = True
        SelectedCalCheck = 0
    Else
        ' at least one calCheck
        frmCalCheckData.Visible = True
        frmCalCheckSelection.Visible = True
        cmdPrint.Visible = True
        lblMessage2.Visible = False
        rsTable.MoveFirst
        lastDTS = 0
        idx = 0
        Do While Not rsTable.EOF
            currDTS = rsTable("CalCheckDTS")
            If (currDTS <> lastDTS) Then
                idx = idx + 1
                calchkDtsList(idx) = currDTS
                lastDTS = currDTS
            End If
            rsTable.MoveNext
        Loop
        idxMax = idx
        SelectedCalCheck = 1
        rsTable.Close
        dbDbase.Close
    End If
End Sub

Private Sub cmdPrint_Click()
    ' Print Current View
    lblMessage.Caption = ""
    Set pbCapture.Picture = CaptureForm(Me)
    PrintPictureToFitPage Printer, pbCapture.Picture
    Printer.EndDoc
    Set pbCapture.Picture = Nothing
    lblMessage.Font.Size = 9.5
    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = "Current screen sent to" & vbCrLf & PRINTERNAME
End Sub

Private Sub cmdReturn_Click()
    Xit
End Sub

Private Sub Form_Load()
'
Dim idx As Integer
    bFormLoaded = False
    ' set display colors
'    frmCalControls.ForeColor = Titles_ForeColor
    frmGroupSelection.ForeColor = Titles_ForeColor
    frmInputSelection.ForeColor = Titles_ForeColor
    frmCalCheckSelection.ForeColor = Titles_ForeColor
'    frmCalFormula.ForeColor = Titles_ForeColor
'    frmCalInformation.ForeColor = Titles_ForeColor
    frmCalCheckData.ForeColor = Titles_ForeColor
'    frmCalGraph.ForeColor = Titles_ForeColor
'    txtNumCalPts.ForeColor = TitlesData_Forecolor
    txtDispGrp.ForeColor = TitlesData_Forecolor
    lblCurCalDts.ForeColor = TitlesData_Forecolor
    lblCurrent.ForeColor = TitlesData_Forecolor
    lblCalCheckDtsMsg.ForeColor = TitlesLabel_ForeColor
    For idx = 1 To MAXLSQCALPOINTS
        lblPointNum(idx).ForeColor = Black
        txtDesired(idx).ForeColor = TitlesData_Forecolor
        lblCalibrated(idx).ForeColor = Black
        txtActual(idx).ForeColor = TitlesData_Forecolor
        lblPercDiff(idx).ForeColor = Black
    Next idx

    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = " "

    ' Set the current station, mfc and point(row)
    SelectedStation = 1
    SelectedMFC = MAXMFC
    SelectedRow = 1
    ' find the first valid mfc
    NextInput
    ' init cal point columns
'    EnableNewCal False
'    EnablePrevCal False
    ' init Unsaved Calibration flag
'    bUnsavedCal = False
'    blnPrevCalExists = False
    ' hide the "Question" frame
'    frmQuestion.Top = OutOfSight
    ' update the screen
    DisplayMfcAll
    
    bFormLoaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub Xit()
    ' close db
'    rsTable.Close
'    dbDbase.Close
    ' close screen
    Unload Me
    Set frmCalCheckHistory = Nothing
End Sub

'*******************************************************************************************************************************
'*******************************************************************************************************************************
'*******************************************************************************************************************************


