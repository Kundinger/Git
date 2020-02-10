VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAkCmdGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AK FcClient"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16740
   Icon            =   "frmAkCmdGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   16740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   29
      Left            =   20880
      TabIndex        =   137
      Text            =   "SFDA K0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   28
      Left            =   20880
      TabIndex        =   136
      Text            =   "SFDM K0 2"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   27
      Left            =   20880
      TabIndex        =   135
      Text            =   "SADF K0 2"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   26
      Left            =   20880
      TabIndex        =   134
      Text            =   "SADF K0 1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   25
      Left            =   20880
      TabIndex        =   133
      Text            =   "SADO K0 2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   24
      Left            =   20880
      TabIndex        =   132
      Text            =   "SADO K0 1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   23
      Left            =   20880
      TabIndex        =   131
      Text            =   "SVAC K0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      DisabledPicture =   "frmAkCmdGen.frx":57E2
      Height          =   375
      Index           =   1
      Left            =   18840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":5B24
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStart 
      DisabledPicture =   "frmAkCmdGen.frx":5E66
      Height          =   375
      Index           =   1
      Left            =   18000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":61A8
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.Frame frmFlows 
      Caption         =   "Flows"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   80
      Top             =   8040
      Width           =   9315
      Begin VB.TextBox txtTot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2797
         TabIndex        =   96
         Text            =   "0.0"
         Top             =   1050
         Width           =   1000
      End
      Begin VB.TextBox txtTotSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2797
         TabIndex        =   95
         Text            =   "0.0"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkCmdGen.frx":64EA
         Height          =   375
         Index           =   5
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":682C
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   975
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkCmdGen.frx":6B6E
         Height          =   375
         Index           =   5
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":6EB0
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txtFlow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2797
         TabIndex        =   86
         Text            =   "0.0"
         Top             =   510
         Width           =   1000
      End
      Begin VB.TextBox txtFlow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   85
         Text            =   "0.0"
         Top             =   510
         Width           =   1000
      End
      Begin VB.TextBox txtFlowSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   84
         Text            =   "0.0"
         Top             =   780
         Width           =   1000
      End
      Begin VB.TextBox txtFlowSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2797
         TabIndex        =   83
         Text            =   "0.0"
         Top             =   780
         Width           =   1000
      End
      Begin VB.TextBox txtTot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   82
         Text            =   "0.0"
         Top             =   1050
         Width           =   1000
      End
      Begin VB.TextBox txtTotSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   81
         Text            =   "0.0"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.Label lblTempName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Target"
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
         Index           =   7
         Left            =   1080
         TabIndex        =   92
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label lblTempName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Totalizer"
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
         Index           =   6
         Left            =   1080
         TabIndex        =   91
         Top             =   1050
         Width           =   1260
      End
      Begin VB.Label lblTankTemp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Tank"
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
         Index           =   3
         Left            =   2745
         TabIndex        =   90
         Top             =   240
         Width           =   1105
      End
      Begin VB.Label lblTankTemp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Gen"
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
         Index           =   0
         Left            =   5955
         TabIndex        =   89
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblTempName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Flow"
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
         Index           =   5
         Left            =   1080
         TabIndex        =   88
         Top             =   510
         Width           =   1260
      End
      Begin VB.Label lblTempName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Flow SP"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   87
         Top             =   780
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   68
      Top             =   5040
      Width           =   9315
      Begin VB.TextBox txtRemoteLocal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2040
         TabIndex        =   103
         Text            =   "idle"
         Top             =   240
         Width           =   5925
      End
      Begin VB.TextBox txtDispenseStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5895
         TabIndex        =   102
         Text            =   "idle"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtTcStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   6900
         TabIndex        =   101
         Text            =   "idle"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtAdfStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   4890
         TabIndex        =   100
         Text            =   "idle"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtDispenseStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2745
         TabIndex        =   99
         Text            =   "idle"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtTcStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   3750
         TabIndex        =   98
         Text            =   "idle"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtAdfStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   1740
         TabIndex        =   97
         Text            =   "idle"
         Top             =   1320
         Width           =   1000
      End
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkCmdGen.frx":71F2
         Height          =   375
         Index           =   6
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":7534
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkCmdGen.frx":7876
         Height          =   375
         Index           =   6
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":7BB8
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txtSystemStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2040
         TabIndex        =   69
         Text            =   "idle"
         Top             =   495
         Width           =   5925
      End
      Begin VB.Label lblDispense 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dispense"
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
         Index           =   2
         Left            =   5895
         TabIndex        =   77
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblTC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Temp Cntrl"
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
         Index           =   2
         Left            =   6900
         TabIndex        =   76
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblADF 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ADF"
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
         Index           =   2
         Left            =   4890
         TabIndex        =   75
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblTank 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Gen"
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
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   74
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblTank 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Tank"
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
         Height          =   255
         Index           =   1
         Left            =   2745
         TabIndex        =   73
         Top             =   840
         Width           =   1105
      End
      Begin VB.Label lblDispense 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Dispense"
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
         Index           =   1
         Left            =   2745
         TabIndex        =   72
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblTC 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Temp Cntrl"
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
         Index           =   1
         Left            =   3750
         TabIndex        =   71
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblADF 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ADF"
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
         Index           =   1
         Left            =   1740
         TabIndex        =   70
         Top             =   1080
         Width           =   1005
      End
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   22
      Left            =   18000
      TabIndex        =   67
      Text            =   "SVAC K0"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   4
      Left            =   18960
      TabIndex        =   66
      Text            =   "SENT K0 0 0 1"
      Top             =   150
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   20
      Left            =   17880
      TabIndex        =   65
      Text            =   "EDUF K0 6 2.33"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   19
      Left            =   17880
      TabIndex        =   64
      Text            =   "EDUF K0 5 2.33"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame frmTemps 
      Caption         =   "Temperatures"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   52
      Top             =   6840
      Width           =   9315
      Begin VB.TextBox txtTempSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2795
         TabIndex        =   58
         Text            =   "0.0"
         Top             =   780
         Width           =   1005
      End
      Begin VB.TextBox txtTempSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   57
         Text            =   "0.0"
         Top             =   780
         Width           =   1005
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   2
         Left            =   5955
         TabIndex        =   56
         Text            =   "0.0"
         Top             =   510
         Width           =   1005
      End
      Begin VB.TextBox txtTemp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   2795
         TabIndex        =   55
         Text            =   "0.0"
         Top             =   510
         Width           =   1005
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkCmdGen.frx":7EFA
         Height          =   375
         Index           =   3
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":823C
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkCmdGen.frx":857E
         Height          =   375
         Index           =   3
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":88C0
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.Label lblTempName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tank Temp SP"
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
         Index           =   2
         Left            =   1080
         TabIndex        =   62
         Top             =   780
         Width           =   1305
      End
      Begin VB.Label lblTempName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tank Temp"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   61
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label lblTankTemp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Gen"
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
         Index           =   2
         Left            =   5955
         TabIndex        =   60
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblTankTemp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Tank"
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
         Index           =   1
         Left            =   2745
         TabIndex        =   59
         Top             =   240
         Width           =   1105
      End
   End
   Begin VB.Timer tmrScreen 
      Interval        =   333
      Left            =   18720
      Top             =   5640
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   51
      Text            =   "frmAkCmdGen.frx":8C02
      Top             =   60
      Width           =   8175
   End
   Begin VB.CommandButton cmdStart 
      DisabledPicture =   "frmAkCmdGen.frx":8C0A
      Height          =   375
      Index           =   4
      Left            =   18000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":8F4C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   10095
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStop 
      DisabledPicture =   "frmAkCmdGen.frx":928E
      Height          =   375
      Index           =   4
      Left            =   18840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":95D0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10095
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStop 
      DisabledPicture =   "frmAkCmdGen.frx":9912
      Height          =   375
      Index           =   2
      Left            =   18840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":9C54
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStart 
      DisabledPicture =   "frmAkCmdGen.frx":9F96
      Height          =   375
      Index           =   2
      Left            =   18000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":A2D8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "master"
      DisabledPicture =   "frmAkCmdGen.frx":A61A
      DownPicture     =   "frmAkCmdGen.frx":A95C
      Height          =   615
      Index           =   7
      Left            =   17895
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":AC9E
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3315
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   18
      Left            =   19320
      TabIndex        =   43
      Text            =   "ETOT K0 1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   17
      Left            =   19320
      TabIndex        =   42
      Text            =   "SENT K0 1 0 0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   16
      Left            =   17880
      TabIndex        =   40
      Text            =   "SFDS K0 2"
      Top             =   5175
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   15
      Left            =   19320
      TabIndex        =   39
      Text            =   "SPAU K0"
      Top             =   4230
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   14
      Left            =   19320
      TabIndex        =   38
      Text            =   "SFDM K0 1"
      Top             =   2610
      Width           =   1215
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   4
      Interval        =   634
      Left            =   18240
      Top             =   7080
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   4
      Left            =   19320
      TabIndex        =   36
      Text            =   "A065 K0 1"
      Top             =   7095
      Width           =   1215
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   3
      Interval        =   633
      Left            =   18240
      Top             =   6600
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   3
      Left            =   19320
      TabIndex        =   32
      Text            =   "ATEM K0"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   13
      Left            =   19320
      TabIndex        =   31
      Text            =   "SMAN K0"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   12
      Left            =   19320
      TabIndex        =   30
      Text            =   "SSPL K0"
      Top             =   2205
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Index           =   1
      Left            =   9600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   29
      Text            =   "frmAkCmdGen.frx":AFE0
      Top             =   5160
      Width           =   7005
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   11
      Left            =   17880
      TabIndex        =   28
      Text            =   "SFDS K0 1"
      Top             =   4770
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   10
      Left            =   17880
      TabIndex        =   27
      Text            =   "SADS K0 2"
      Top             =   4365
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   9
      Left            =   17880
      TabIndex        =   26
      Text            =   "ETEM K0 2"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   8
      Left            =   17880
      TabIndex        =   25
      Text            =   "EDUF K0 3 2.23"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   7
      Left            =   17880
      TabIndex        =   24
      Text            =   "EDUF K0 2 1.23"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   6
      Left            =   19320
      TabIndex        =   23
      Text            =   "ETEM K0 1"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   5
      Left            =   19320
      TabIndex        =   22
      Text            =   "EDUF K0 1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   21
      Left            =   17880
      TabIndex        =   21
      Text            =   "EDUF K0 1 1.23"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   3
      Left            =   19320
      TabIndex        =   20
      Text            =   "STBY K0"
      Top             =   3015
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   2
      Left            =   17880
      TabIndex        =   19
      Text            =   "SADS K0 1"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   1
      Left            =   19320
      TabIndex        =   18
      Text            =   "EZRO K0"
      Top             =   3825
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":AFEA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   1
      Interval        =   631
      Left            =   18240
      Top             =   5640
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   1
      Left            =   19320
      TabIndex        =   15
      Text            =   "AZET K0"
      Top             =   5655
      Width           =   1215
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   6
      Interval        =   636
      Left            =   18720
      Top             =   7080
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   6
      Left            =   19320
      TabIndex        =   13
      Text            =   "ASTZ K0"
      Top             =   8295
      Width           =   1215
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   5
      Left            =   19320
      TabIndex        =   8
      Text            =   "ADUF K0 1"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   2
      Left            =   19320
      TabIndex        =   7
      Text            =   "ASTF K0"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   5
      Interval        =   635
      Left            =   18720
      Top             =   6585
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   2
      Interval        =   632
      Left            =   18240
      Top             =   6105
   End
   Begin MSWinsockLib.Winsock sockMain 
      Left            =   18720
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   0
      Left            =   19320
      TabIndex        =   6
      Text            =   "SREM K0"
      Top             =   3420
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Index           =   0
      Left            =   9600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmAkCmdGen.frx":B32C
      Top             =   960
      Width           =   7005
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkCmdGen.frx":B336
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "5500"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtHost 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "10.0.0.3"
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame frmControls 
      Caption         =   "Controls"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   63
      Top             =   840
      Width           =   9315
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H0080C0FF&
         Caption         =   "ZERO TOTAL #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   5925
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":B678
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Zero the Fuel Flow Totalizer for Tank #1"
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox txtParam2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   9
         Left            =   8520
         TabIndex        =   130
         Text            =   "0"
         Top             =   3480
         Width           =   600
      End
      Begin VB.TextBox txtParam2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   6
         Left            =   8520
         TabIndex        =   129
         Text            =   "0"
         Top             =   2865
         Width           =   600
      End
      Begin VB.TextBox txtParam1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   9
         Left            =   7920
         TabIndex        =   128
         Text            =   "30"
         Top             =   3480
         Width           =   600
      End
      Begin VB.TextBox txtParam1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   6
         Left            =   7920
         TabIndex        =   127
         Text            =   "30"
         Top             =   2865
         Width           =   600
      End
      Begin VB.TextBox txtParam1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   18
         Left            =   7920
         TabIndex        =   126
         Text            =   "25"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.TextBox txtParam1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   5
         Left            =   7920
         TabIndex        =   125
         Text            =   "2"
         Top             =   405
         Width           =   1215
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START AUTO DISP #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":B9BA
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   1485
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START DISPENSE #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   28
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":BCFC
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   2100
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START DISPENSE #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":C03E
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   1485
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START ADF #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   27
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":C380
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START DRAIN ONLY #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":C6C2
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START DRAIN ONLY #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":CA04
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0FFFF&
         Caption         =   "REMOTE CONTROL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":CD46
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   3330
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "START ADF #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":D088
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0FFFF&
         Caption         =   "LOCAL CONTROL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":D3CA
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   3330
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H0080C0FF&
         Caption         =   "SET TEMP RAMP #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   5925
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":D70C
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Set the Fuel Temperature Ramp (endtemp duration) for Tank #1"
         Top             =   2715
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H0080C0FF&
         Caption         =   "SET TOTAL TARGET #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   5925
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":DA4E
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Set the Fuel Flow Totalizer Target for Tank #1"
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H0080C0FF&
         Caption         =   "SET FLOW SP #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   5925
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":DD90
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Set the Fuel Flow SetPoint for Tank #1"
         Top             =   1485
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H0080C0FF&
         Caption         =   "SET TEMP RAMP #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   5925
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":E0D2
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Set the Fuel Temperature Ramp (endtemp duration) for Tank #2"
         Top             =   3330
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0C000&
         Caption         =   "STANDBY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2055
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":E414
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   2715
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0C000&
         Caption         =   "PAUSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   3990
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":E756
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   2715
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "STOP DISPENSE #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   3990
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":EA98
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   1485
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "STOP ADF #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   3990
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":EDDA
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "STOP ADF #1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   3990
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":F11C
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "STOP DISPENSE #2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   3990
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkCmdGen.frx":F45E
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   2100
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read System Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19560
      TabIndex        =   50
      Top             =   10800
      Width           =   2655
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read MFC Flows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19680
      TabIndex        =   49
      Top             =   10560
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read A065 Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19680
      TabIndex        =   48
      Top             =   10170
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read System Temperatures"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19680
      TabIndex        =   47
      Top             =   9840
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read Errors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19680
      TabIndex        =   46
      Top             =   9555
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read Sample Times"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19680
      TabIndex        =   45
      Top             =   9240
      Width           =   2655
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "version:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12720
      TabIndex        =   41
      Top             =   240
      Width           =   3915
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " @ 1 Hz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   21960
      TabIndex        =   37
      Top             =   10170
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " @ 1 Hz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   21960
      TabIndex        =   33
      Top             =   9960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " @ 1 Hz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   21960
      TabIndex        =   16
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "@ 10 Hz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   22080
      TabIndex        =   14
      Top             =   10800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "@ 10 Hz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   22080
      TabIndex        =   10
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " @ 1 Hz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   21960
      TabIndex        =   9
      Top             =   9555
      Width           =   1095
   End
   Begin VB.Label lblPort 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Port "
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
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   390
      Width           =   855
   End
   Begin VB.Label lblHost 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Server "
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
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   855
   End
End
Attribute VB_Name = "frmAkCmdGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CURVERSION = "18 April 2019"
'
' width Constants
Private Const ShortWidth = 12855
Private Const FullWidth = 18735
' color Constants from Color Pallet
Private Const SOFTBLUE = &HFF6830
Private Const SOFTBLUETOO = &HFF6000
' color constants from Color Table
' ********************************
Private Const AliceBlue = &HFFF8F0
Private Const AntiqueWhite = &HD7EBFA
Private Const Aqua = &HFFFF00
Private Const Aquamarine = &HD4FF7F
Private Const Azure = &HFFFFF0
Private Const Beige = &HDCF5F5
Private Const Bisque = &HC4E4FF
Private Const Black = &H0&
Private Const BlanchedAlmond = &HCDEBFF
Private Const Blue = &HFF0000
Private Const BlueViolet = &HE22B8A
Private Const Brown = &H2A2AA5
Private Const BurlyWood = &H87B8DE
Private Const CadetBlue = &HA09E5F
Private Const Chartreuse = &HFF7F&
Private Const Chocolate = &H1E69D2
Private Const Coral = &H507FFF
Private Const CornFlowerBlue = &HED9564
Private Const CornSilk = &HDCF8FF
Private Const Crimson = &H3C14DC
Private Const Cyan = &HFFFF00
Private Const DarkBlue = &H8B0000
Private Const DarkCyan = &H8B8B00
Private Const DarkGoldenRod = &HB86B8
Private Const DarkGray = &HA9A9A9
Private Const DarkGreen = &H6400&
Private Const DarkKhaki = &H6BB7BD
Private Const DarkMagenta = &H8B008B
Private Const DarkOliveGreen = &H2F6B55
Private Const DarkOrange = &H8CFF&
Private Const DarkOrchid = &HCC3299
Private Const DarkRed = &H8B&
Private Const DarkSalmon = &H7A96E9
Private Const DarkSeaGreen = &H8BBC8F
Private Const DarkSlateBlue = &H8B3D48
Private Const DarkSlateGray = &H4F4F2F
Private Const DarkTurquoise = &HD1CE00
Private Const DarkViolet = &HD30094
Private Const DeepPink = &H9314FF
Private Const DeepSkyBlue = &HFFBF00
Private Const DimGray = &H696969
Private Const DodgerBlue = &HFF901E
Private Const FireBrick = &H2222B2
Private Const FloralWhite = &HF0FAFF
Private Const ForestGreen = &H228B22
Private Const Fuchsia = &HFF00FF
Private Const Gainsboro = &HDCDCDC
Private Const GhostWhite = &HFFF8F8
Private Const Gold = &HD7FF&
Private Const Goldenrod = &H20A5DA
Private Const Gray = &H808080
Private Const green = &H8000&
Private Const GreenYellow = &H2FFFAD
Private Const Honeydew = &HF0FFF0
Private Const HotPink = &HB469FF
Private Const IndianRed = &H5C5CCD
Private Const Indigo = &HB2004B
Private Const Ivory = &HF0FFFF
Private Const Khaki = &H8CE6F0
Private Const Lavender = &HFAE6E6
Private Const LavenderBlush = &HF5F0FF
Private Const LawnGreen = &HFC7C&
Private Const LemonChiffon = &HCDFAFF
Private Const LightBlue = &HE6D8AD
Private Const LightCoral = &H8080F0
Private Const LightCyan = &HFFFFE0
Private Const LightGoldenrodYellow = &HD2FAFA
Private Const LightGreen = &H90EE90
Private Const LightGray = &HD3D3D3
Private Const LightPink = &HC1B6FF
Private Const LightSalmon = &H7AA0FF
Private Const LightSeaGreen = &HAAB220
Private Const LightSkyBlue = &HFACE87
Private Const LightSlateGray = &H998877
Private Const LightSteelBlue = &HDEC4B0
Private Const LightYellow = &HE0FFFF
Private Const Lime = &HFF00&
Private Const LimeGreen = &H32CD32
Private Const Linen = &HE6F0FA
Private Const Magenta = &HFF00FF
Private Const Maroon = &H80&
Private Const MediumAquamarine = &HAACD66
Private Const MediumBlue = &HCD0000
Private Const MediumOrchid = &HD355BA
Private Const MediumPurple = &HDB7093
Private Const MediumSeaGreen = &H71B33C
Private Const MediumSlateBlue = &HEE687B
Private Const MediumSpringGreen = &H9AFA00
Private Const MediumTurquoise = &HCCD148
Private Const MediumVioletRed = &H8515C7
Private Const MidnightBlue = &H701919
Private Const MintCream = &HFAFFF5
Private Const MistyRose = &HE1E4FF
Private Const Moccasin = &HB5E4FF
Private Const NavajoWhite = &HADDEFF
Private Const Navy = &H800000
Private Const Olive = &H8080&
Private Const Olivedrab = &H238E6B
Private Const Orange = &HA5FF&
Private Const OrangeRed = &H45FF&
Private Const Orchid = &HD670DA
Private Const PaleGoldenrod = &HAAE8EE
'Private Const PaleGreen = &H98FB98
Private Const PaleTurquoise = &HEEEEAF
Private Const PaleVioletRed = &H9370DB
Private Const PapayaWhip = &HD5EFFF
Private Const PeachPuff = &HB9DAFF
Private Const Peru = &H3F85CD
Private Const Pink = &HCBC0FF
Private Const Plum = &HDDA0DD
Private Const PowderBlue = &HE6E0B0
Private Const Purple = &H800080
Private Const red = &HFF&
Private Const RosyBrown = &H8F8FBC
Private Const RoyalBlue = &HE16941
Private Const SaddleBrown = &H13458B
Private Const Salmon = &H7280FA
Private Const SandyBrown = &H60A4F4
Private Const SeaGreen = &H578B2E
Private Const Seashell = &HEEF5FF
Private Const Sienna = &H2D52A0
Private Const Silver = &HC0C0C0
Private Const SkyBlue = &HEBCE87
Private Const SlateBlue = &HCD5A6A
Private Const SlateGray = &H908070
Private Const Snow = &HFAFAFF
Private Const SpringGreen = &H7FFF00
Private Const SteelBlue = &HB48246
Private Const Tan = &H8CB4D2
Private Const Teal = &H808000
Private Const Thistle = &HD8BFD8
Private Const Tomato = &H4763FF
Private Const Turquoise = &HD0E040
Private Const Violet = &HEE82EE
Private Const Wheat = &HB3DEF5
Private Const White = &HFFFFFF
Private Const WhiteSmoke = &HF5F5F5
Private Const Yellow = &HFFFF&
Private Const YellowGreen = &H32CD9A

Private OKtoSend(1 To 8) As Boolean
Private strSend(1 To 100) As String
Private strSendOnce As String
Private idxLoad As Integer
Private idxUnload As Integer
Private idxMax As Integer
Private openCmd As Boolean
Private WinsockStateDesc(0 To 9) As String
Private Params_Numeric(1 To 9) As Single
Private Params_String(1 To 9) As String
Private Params_Count As Integer
Private Params_Type As Integer             ' 0 = Floating Point, 1 = Integer, 2 = string
Private RcdCmdCode As String * 4
Private SrvrIP As String

Private Sub cmdConnect_Click()
    sockMain.RemoteHost = txtHost.Text
    sockMain.RemotePort = txtPort.Text
    sockMain.Connect
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
    sockMain.Close
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub


Private Sub cmdSend_Click(Index As Integer)
    Select Case Index
        Case 1
            ' zero total 1
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).Text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        Case 5
            ' flow total 1
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).Text & Chr(32) & txtParam1(Index).Text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        Case 18
            ' flow SP 1
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).Text & Chr(32) & txtParam1(Index).Text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        Case 6
            ' temp ramp 1
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).Text & Chr(32) & txtParam1(Index).Text & Chr(32) & txtParam2(Index).Text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        Case 9
            ' temp ramp 2
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).Text & Chr(32) & txtParam1(Index).Text & Chr(32) & txtParam2(Index).Text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        Case Else
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).Text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
    End Select
    SendData
End Sub

Private Sub cmdStart_Click(Index As Integer)
    OKtoSend(Index) = True
    cmdStart(Index).Enabled = False
    cmdStop(Index).Enabled = True
End Sub

Private Sub cmdStop_Click(Index As Integer)
    OKtoSend(Index) = False
    cmdStart(Index).Enabled = True
    cmdStop(Index).Enabled = False
End Sub

Private Sub Form_Load()
Dim idx As Integer
Dim color As Long
Dim CmdLine As String
'    Form2.Show

    'Get command line arguments.
    CmdLine = Command()
'    If Len(CmdLine) > 0 Then frmAkCmdGen.Caption = frmAkCmdGen.Caption & " #" & Mid(CmdLine, 1, 1)
    If Len(CmdLine) > 8 Then
        frmAkCmdGen.txtHost.Text = Mid(CmdLine, 1, Len(CmdLine))
    Else
        frmAkCmdGen.txtHost.Text = "127.0.0.1"
    End If
    
    For idx = cmdStart.LBound To cmdStart.UBound
        cmdStart(idx).Enabled = True
        cmdStop(idx).Enabled = False
    Next idx
    
    For idx = cmdSend.LBound To cmdSend.UBound
        Select Case idx
            Case 7
                ' ignore; do nothing
            Case 4, 7, 8, 12, 17, 19, 20, 21, 22, 23
                ' no cmdsend; unused; do nothing
            Case Else
                cmdSend(idx).Picture = cmdSend(7).Picture
                cmdSend(idx).DisabledPicture = cmdSend(7).DisabledPicture
                cmdSend(idx).DownPicture = cmdSend(7).DownPicture
        End Select
    Next idx
    
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    
    idxLoad = 1
    idxUnload = 1
    idxMax = 100
    openCmd = False
    
'    color = SteelBlue
'    lblHost.ForeColor = color
'    lblPort.ForeColor = color
'    color = SOFTBLUETOO
    color = RoyalBlue
    txtHost.ForeColor = color
    txtPort.ForeColor = color
    
    color = BurlyWood
    cmdSend(0).BackColor = color
    cmdSend(13).BackColor = color
    color = Moccasin
    cmdSend(1).BackColor = color
    cmdSend(5).BackColor = color
    cmdSend(18).BackColor = color
    cmdSend(6).BackColor = color
    cmdSend(9).BackColor = color
    color = DarkCyan
    cmdSend(3).BackColor = color
    cmdSend(15).BackColor = color
    color = SandyBrown
    cmdSend(26).BackColor = color
    cmdSend(27).BackColor = color
    cmdSend(29).BackColor = color
    cmdSend(24).BackColor = color
    cmdSend(25).BackColor = color
    cmdSend(14).BackColor = color
    cmdSend(28).BackColor = color
    color = CornSilk
    cmdSend(2).BackColor = color
    cmdSend(10).BackColor = color
    cmdSend(11).BackColor = color
    cmdSend(16).BackColor = color
    
    color = Teal
    txtSystemStatus.Text = ""
    txtSystemStatus.ForeColor = color
    color = Goldenrod
    For idx = 1 To 2
        txtAdfStatus(idx).Text = ""
        txtAdfStatus(idx).ForeColor = color
        txtDispenseStatus(idx).Text = ""
        txtDispenseStatus(idx).ForeColor = color
        txtTcStatus(idx).Text = ""
        txtTcStatus(idx).ForeColor = color
    Next idx
    
    color = RoyalBlue
    For idx = 1 To 2
        txtTemp(idx).ForeColor = color
        txtTempSP(idx).ForeColor = color
    Next idx

    color = RoyalBlue
    For idx = 1 To 2
        txtFlow(idx).Text = ""
        txtFlowSP(idx).ForeColor = color
        txtTot(idx).Text = ""
        txtTotSP(idx).ForeColor = color
    Next idx
    
    color = DarkGray
    
    
    color = RoyalBlue
    txtMsg.ForeColor = color
    
    color = SOFTBLUE
    lblVersion.ForeColor = color
    lblVersion.Caption = "version: " & CURVERSION
    
    InitArraysEtc
End Sub

Private Sub frmControls_DblClick()
    frmAkCmdGen.Width = IIf(frmAkCmdGen.Width = FullWidth, ShortWidth, FullWidth)
End Sub

Private Sub sockMain_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    
    sockMain.GetData strData, vbString
    strData = Mid(strData, 3, (Len(strData) - 3))
    RcdData_Parse strData
    RcdData_Read strData
    openCmd = False
    If (idxUnload <> idxLoad) Then SendData
End Sub

Private Sub RcdData_Parse(ByVal rcdStr As String)
Dim paramsStr, tmpStr, curChar As String
Dim iChar, iChar2, iParam, Max As Integer
    For iParam = 1 To 9
        Params_String(iParam) = ""
        Params_Numeric(iParam) = 0
    Next iParam
    paramsStr = ""
    tmpStr = ""
    iParam = 0
    Params_Count = 0
    Params_Type = 1
    RcdCmdCode = Mid(rcdStr, 1, 4)
    If (Len(rcdStr) > 6) Then
        paramsStr = Mid(rcdStr, 7, (Len(rcdStr) - 6))
        Max = Len(paramsStr)
        For iChar = 1 To Max
            curChar = Mid(paramsStr, iChar, 1)
            If (curChar <> " ") Then
                tmpStr = tmpStr & curChar
            End If
            If (iChar = Max) Or (curChar = " ") Then
                If Len(tmpStr) > 0 Then
                    iParam = iParam + 1
                    Params_Count = iParam
                    If IsNumeric(tmpStr) Then
                        For iChar2 = 1 To Len(tmpStr)
                            If Mid(tmpStr, iChar2, 1) = "." Then Params_Type = 0
                        Next iChar2
                        Select Case Params_Type
                            Case 0
                                ' floating point
                                Params_Numeric(iParam) = CDbl(tmpStr)
                            Case 1
                                ' integer
                                Params_Numeric(iParam) = CInt(tmpStr)
                        End Select
                        Params_String(iParam) = Trim(tmpStr)
                    Else
                        ' string
                        Params_String(iParam) = Trim(tmpStr)
                        Params_Type = 2
                    End If
                    tmpStr = ""
                End If
            End If
        Next iChar
    End If
        
End Sub

Private Sub RcdData_Read(ByVal strData As String)
Dim iStatusBox As Integer
Dim idx As Integer
    Select Case RcdCmdCode
        Case "ADUF"
            iStatusBox = 1
            idx = CInt(Params_Numeric(1))
            If ((idx = 1) Or (idx = 2)) Then
                ' Fuel Flow
                txtFlow(idx).Text = Params_String(2)
                ' Fuel Totalizer
                txtTot(idx).Text = Params_String(3)
                ' Fuel Flow SP
                txtFlowSP(idx).Text = Params_String(4)
                ' Totalizer Target
                txtTotSP(idx).Text = Params_String(5)
            End If
        Case "ASTZ"
            iStatusBox = 1
            ' Remote/Local Status
            txtRemoteLocal.Text = IIf(Params_String(7) = "SREM", "Remote Control", "Local Control")
            txtRemoteLocal.ForeColor = IIf(Params_String(7) = "SREM", DarkOrange, SlateGray)
'            txtRemoteLocal.ForeColor = IIf(Params_String(2) = "SREM", Tomato, SlateGray)
'            txtRemoteLocal.ForeColor = IIf(Params_String(2) = "SREM", DeepSkyBlue, SlateGray)
            ' System Status
            txtSystemStatus.Text = Params_String(7)
            txtSystemStatus.ForeColor = txtRemoteLocal.ForeColor
            For idx = 1 To 2
                ' Adf Status
                txtAdfStatus(idx).Text = Params_String(idx + 2)
                ' Dispense Status
                txtDispenseStatus(idx).Text = Params_String(idx)
                ' Temp Control Status
                txtTcStatus(idx).Text = Params_String(idx + 4)
            Next idx
        Case "ATEM"
            iStatusBox = 1
            ' Fuel Flow #1
            txtTemp(1).Text = Params_String(1)
            ' Fuel Flow #2
            txtTemp(2).Text = Params_String(2)

        Case "ATSP"
            iStatusBox = 1
            ' Fuel Flow #1
            txtTempSP(1).Text = Params_String(1)
            ' Fuel Flow #2
            txtTempSP(2).Text = Params_String(2)
        Case "ASTF"
            iStatusBox = 1
        Case Else
            iStatusBox = 0
    End Select
    txtStatus(iStatusBox).Text = NowPrefixString & strData & vbCrLf & txtStatus(iStatusBox).Text
    If (Len(txtStatus(iStatusBox).Text) > 32000) Then txtStatus(iStatusBox).Text = Mid(txtStatus(iStatusBox).Text, 1, 24000)
End Sub

Private Sub tmrScreen_Timer()
    txtMsg.Text = Format(Now(), "YYYY MMMM D   hh:mm:ss") & vbCrLf & WinsockStateDesc(sockMain.State)
End Sub

Private Sub TmrSendRepeat_Timer(Index As Integer)
    If OKtoSend(Index) Then
        strSend(idxLoad) = Chr(2) & Chr(32) & txtSendRepeat(Index).Text & Chr(3)
        idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        If (Index = 3) Then
            If (InStr(txtSendRepeat(Index).Text, "ATEM") > 0) Then
                txtSendRepeat(3).Text = "ATSP K0"
            ElseIf (InStr(txtSendRepeat(Index).Text, "ATSP") > 0) Then
                txtSendRepeat(3).Text = "ATEM K0"
            End If
        ElseIf (Index = 5) Then
            If (InStr(txtSendRepeat(Index).Text, "K0 1") > 0) Then
                txtSendRepeat(5).Text = "ADUF K0 2"
            ElseIf (InStr(txtSendRepeat(Index).Text, "K0 2") > 0) Then
                txtSendRepeat(5).Text = "ADUF K0 1"
            End If
        End If
    End If
    SendData
End Sub

Private Sub SendData()
    If (Not openCmd And (idxUnload <> idxLoad) And (sockMain.State = sckConnected)) Then
        sockMain.SendData strSend(idxUnload)
        idxUnload = IIf((idxUnload < 100), (idxUnload + 1), 1)
        openCmd = True
    End If
End Sub

Private Function NowPrefixString() As String
    Dim strDTS, strMS, strMS2 As String
    Dim dotpos As Integer
    strMS = Format(Timer, "##,##0.000")
    dotpos = InStr(1, strMS, ".")
    strMS2 = Mid(strMS, dotpos, (Len(strMS) - dotpos + 1))
    strDTS = Format(Now(), "YYYY MMM D  hh:mm:ss") & strMS2 & "   "
    NowPrefixString = strDTS
End Function

Private Sub InitArraysEtc()

    ' *************************************************************************
    '
    ' SOCKMAIN state descriptions
    '
    ' value    name                 description
    '  0    sckClosed               connection closed
    '  1    sckOpen                 open
    '  2    sckListening            listening for incoming connections
    '  3    sckConnectionPending    connection pending
    '  4    sckResolvingHost        resolving remote host name
    '  5    sckHostResolved         remote host name successfully resolved
    '  6    sckConnecting           connecting to remote host
    '  7    sckConnected            connected to remote host
    '  8    sckClosing              Connection Is closing
    '  9    sckError                error occured
    '
    ' *************************************************************************
    WinsockStateDesc(sckClosed) = "Connection Closed"
    WinsockStateDesc(sckOpen) = "Open"
    WinsockStateDesc(sckListening) = "Listening for Incoming Connections"
    WinsockStateDesc(sckConnectionPending) = "Connection Pending"
    WinsockStateDesc(sckResolvingHost) = "Resolving Remote Server Name"
    WinsockStateDesc(sckHostResolved) = "Remote Server Name Successfully Resolved"
    WinsockStateDesc(sckConnecting) = "Connecting to Remote Server"
    WinsockStateDesc(sckConnected) = "Connected to Remote Server"
    WinsockStateDesc(sckClosing) = "Connection is Closing"
    WinsockStateDesc(sckError) = "Error Occured"
    
    SrvrIP = "127.0.0.1"
End Sub
