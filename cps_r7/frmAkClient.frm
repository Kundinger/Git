VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAkClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AK Client"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13830
   ControlBox      =   0   'False
   Icon            =   "frmAkClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Picture         =   "frmAkClient.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      DisabledPicture =   "frmAkClient.frx":5B24
      DownPicture     =   "frmAkClient.frx":6226
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12720
      Picture         =   "frmAkClient.frx":6928
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Frame frmQ 
      Caption         =   "Cmd Q"
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
      Height          =   615
      Left            =   720
      TabIndex        =   124
      Top             =   840
      Width           =   3915
      Begin VB.TextBox txtLoadIdx 
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
         Left            =   690
         TabIndex        =   126
         Text            =   "00"
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox txtUnloadIdx 
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
         Left            =   3150
         TabIndex        =   125
         Text            =   "00"
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblLoadidx 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Load:"
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
         Left            =   150
         TabIndex        =   128
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblUnloadidx 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "UnLoad:"
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
         Left            =   2490
         TabIndex        =   127
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdReq 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLEAR REQUEST IN"
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
      Left            =   2700
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":702A
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   4140
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton cmdReq 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SET REQUEST IN"
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
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":736C
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   4140
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame frmRequest 
      Caption         =   "Request Out"
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
      Height          =   1125
      Left            =   720
      TabIndex        =   117
      Top             =   3000
      Width           =   3915
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkClient.frx":76AE
         Height          =   375
         Index           =   1
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":79F0
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkClient.frx":7D32
         Height          =   375
         Index           =   1
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":8074
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.Shape shpRequestOut 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   2790
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblRequestOut 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Request Out"
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
         Left            =   2588
         TabIndex        =   121
         Top             =   360
         Width           =   1125
      End
      Begin VB.Shape shpRequestIn 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   1290
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblRequestIn 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Request In"
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
         Left            =   1148
         TabIndex        =   120
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   29
      Left            =   20880
      TabIndex        =   112
      Text            =   "SFDA K0"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   28
      Left            =   20880
      TabIndex        =   111
      Text            =   "SFDM K0 2"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   27
      Left            =   20880
      TabIndex        =   110
      Text            =   "SADF K0 2"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   26
      Left            =   20880
      TabIndex        =   109
      Text            =   "SADF K0 1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   25
      Left            =   20880
      TabIndex        =   108
      Text            =   "SADO K0 2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   24
      Left            =   20880
      TabIndex        =   107
      Text            =   "SADO K0 1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   23
      Left            =   20880
      TabIndex        =   106
      Text            =   "SVAC K0"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame frmCfg 
      Caption         =   "Configuration"
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
      Height          =   1395
      Left            =   720
      TabIndex        =   70
      Top             =   6240
      Width           =   3915
      Begin VB.TextBox txtCfg 
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
         Index           =   3
         Left            =   2797
         TabIndex        =   80
         Text            =   "0.0"
         Top             =   810
         Width           =   1000
      End
      Begin VB.TextBox txtCfg 
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
         Index           =   4
         Left            =   2797
         TabIndex        =   79
         Text            =   "0.0"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkClient.frx":83B6
         Height          =   375
         Index           =   5
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":86F8
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkClient.frx":8A3A
         Height          =   375
         Index           =   5
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":8D7C
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txtCfg 
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
         TabIndex        =   72
         Text            =   "0.0"
         Top             =   270
         Width           =   1000
      End
      Begin VB.TextBox txtCfg 
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
         Left            =   2797
         TabIndex        =   71
         Text            =   "0.0"
         Top             =   540
         Width           =   1000
      End
      Begin VB.Label lblCfg 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Moisture Tol"
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
         Index           =   4
         Left            =   1080
         TabIndex        =   76
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblCfg 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Temp Tol"
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
         Left            =   1080
         TabIndex        =   75
         Top             =   810
         Width           =   1260
      End
      Begin VB.Label lblCfg 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Temp SP"
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
         TabIndex        =   74
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label lblCfg 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Moisture SP"
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
         TabIndex        =   73
         Top             =   540
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
      Height          =   1455
      Left            =   720
      TabIndex        =   64
      Top             =   1560
      Width           =   3915
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkClient.frx":90BE
         Height          =   375
         Index           =   6
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":9400
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkClient.frx":9742
         Height          =   375
         Index           =   6
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":9A84
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txtPagStatus 
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
         Left            =   990
         TabIndex        =   65
         Text            =   "idle"
         Top             =   495
         Width           =   2800
      End
      Begin VB.Shape shpRdyOut 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   2790
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblRdyOut 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ready Out"
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
         Left            =   2655
         TabIndex        =   116
         Top             =   840
         Width           =   1005
      End
      Begin VB.Shape shpReqIn 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   1290
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lblPagDesc 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "PurgeAir Generator"
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
         Left            =   990
         TabIndex        =   67
         Top             =   240
         Width           =   2800
      End
      Begin VB.Label lblReqIn 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Request In"
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
         Left            =   1155
         TabIndex        =   66
         Top             =   840
         Width           =   1005
      End
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   22
      Left            =   18000
      TabIndex        =   63
      Text            =   "SVAC K0"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   4
      Left            =   18960
      TabIndex        =   62
      Text            =   "SENT K0 0 0 1"
      Top             =   150
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   20
      Left            =   17880
      TabIndex        =   61
      Text            =   "EDUF K0 6 2.33"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Index           =   19
      Left            =   17880
      TabIndex        =   60
      Text            =   "EDUF K0 5 2.33"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame frmActual 
      Caption         =   "Current Values"
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
      Height          =   1155
      Left            =   720
      TabIndex        =   52
      Top             =   5040
      Width           =   3915
      Begin VB.TextBox txtCurrentVals 
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
         Left            =   2790
         TabIndex        =   114
         Text            =   "0.0"
         Top             =   270
         Width           =   1000
      End
      Begin VB.TextBox txtCurrentVals 
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
         Index           =   3
         Left            =   2795
         TabIndex        =   56
         Text            =   "0.0"
         Top             =   810
         Width           =   1005
      End
      Begin VB.TextBox txtCurrentVals 
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
         Left            =   2795
         TabIndex        =   55
         Text            =   "0.0"
         Top             =   540
         Width           =   1005
      End
      Begin VB.CommandButton cmdStop 
         DisabledPicture =   "frmAkClient.frx":9DC6
         Height          =   375
         Index           =   3
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":A108
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdStart 
         DisabledPicture =   "frmAkClient.frx":A44A
         Height          =   375
         Index           =   3
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAkClient.frx":A78C
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.Label lblCurrentVals 
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
         ForeColor       =   &H00404080&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   115
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label lblCurrentVals 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Moisture"
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
         Left            =   1080
         TabIndex        =   58
         Top             =   810
         Width           =   1305
      End
      Begin VB.Label lblCurrentVals 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Humidity"
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
         TabIndex        =   57
         Top             =   540
         Width           =   1305
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
      Text            =   "frmAkClient.frx":AACE
      Top             =   60
      Width           =   7095
   End
   Begin VB.CommandButton cmdStart 
      DisabledPicture =   "frmAkClient.frx":AAD6
      Height          =   375
      Index           =   4
      Left            =   18000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":AE18
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   10095
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStop 
      DisabledPicture =   "frmAkClient.frx":B15A
      Height          =   375
      Index           =   4
      Left            =   18840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":B49C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10095
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStop 
      DisabledPicture =   "frmAkClient.frx":B7DE
      Height          =   375
      Index           =   2
      Left            =   18840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":BB20
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdStart 
      DisabledPicture =   "frmAkClient.frx":BE62
      Height          =   375
      Index           =   2
      Left            =   18000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":C1A4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "master"
      DisabledPicture =   "frmAkClient.frx":C4E6
      DownPicture     =   "frmAkClient.frx":C828
      Height          =   615
      Index           =   7
      Left            =   17895
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":CB6A
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
      Interval        =   333
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Index           =   1
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   29
      Text            =   "frmAkClient.frx":CEAC
      Top             =   3420
      Width           =   8415
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
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAkClient.frx":CEB6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   1
      Interval        =   1031
      Left            =   18240
      Top             =   5640
   End
   Begin VB.TextBox txtSendRepeat 
      Height          =   405
      Index           =   1
      Left            =   19320
      TabIndex        =   15
      Text            =   "SREQ K0"
      Top             =   5655
      Width           =   1215
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   6
      Interval        =   336
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
      Text            =   "ACFG K0"
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
      Interval        =   335
      Left            =   18720
      Top             =   6585
   End
   Begin VB.Timer TmrSendRepeat 
      Index           =   2
      Interval        =   632
      Left            =   18240
      Top             =   6105
   End
   Begin MSWinsockLib.Winsock sockClient 
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmAkClient.frx":D1F8
      Top             =   840
      Width           =   8415
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
      Picture         =   "frmAkClient.frx":D202
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
      Text            =   "5600"
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
      TabIndex        =   59
      Top             =   9960
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
         Picture         =   "frmAkClient.frx":D544
         Style           =   1  'Graphical
         TabIndex        =   113
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         Picture         =   "frmAkClient.frx":D886
         Style           =   1  'Graphical
         TabIndex        =   99
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
         Picture         =   "frmAkClient.frx":DBC8
         Style           =   1  'Graphical
         TabIndex        =   98
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
         Picture         =   "frmAkClient.frx":DF0A
         Style           =   1  'Graphical
         TabIndex        =   97
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
         Picture         =   "frmAkClient.frx":E24C
         Style           =   1  'Graphical
         TabIndex        =   96
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
         Picture         =   "frmAkClient.frx":E58E
         Style           =   1  'Graphical
         TabIndex        =   95
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
         Picture         =   "frmAkClient.frx":E8D0
         Style           =   1  'Graphical
         TabIndex        =   94
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
         Picture         =   "frmAkClient.frx":EC12
         Style           =   1  'Graphical
         TabIndex        =   93
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
         Picture         =   "frmAkClient.frx":EF54
         Style           =   1  'Graphical
         TabIndex        =   92
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
         Picture         =   "frmAkClient.frx":F296
         Style           =   1  'Graphical
         TabIndex        =   91
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
         Picture         =   "frmAkClient.frx":F5D8
         Style           =   1  'Graphical
         TabIndex        =   90
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
         Picture         =   "frmAkClient.frx":F91A
         Style           =   1  'Graphical
         TabIndex        =   89
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
         Picture         =   "frmAkClient.frx":FC5C
         Style           =   1  'Graphical
         TabIndex        =   88
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
         Picture         =   "frmAkClient.frx":FF9E
         Style           =   1  'Graphical
         TabIndex        =   87
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
         Picture         =   "frmAkClient.frx":102E0
         Style           =   1  'Graphical
         TabIndex        =   86
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
         Picture         =   "frmAkClient.frx":10622
         Style           =   1  'Graphical
         TabIndex        =   85
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
         Picture         =   "frmAkClient.frx":10964
         Style           =   1  'Graphical
         TabIndex        =   84
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
         Picture         =   "frmAkClient.frx":10CA6
         Style           =   1  'Graphical
         TabIndex        =   83
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
         Picture         =   "frmAkClient.frx":10FE8
         Style           =   1  'Graphical
         TabIndex        =   82
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
         Picture         =   "frmAkClient.frx":1132A
         Style           =   1  'Graphical
         TabIndex        =   81
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
      Left            =   15640
      TabIndex        =   41
      Top             =   750
      Width           =   7035
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
Attribute VB_Name = "frmAkClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
' width Constants
Private Const ShortWidth = 12855
Private Const FullWidth = 18735
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
Private RcdString As String
Private SrvrIP As String
'
Public AkClientResetReq As Boolean
Public DispConnectFlag As Boolean


Private Sub StartClient()
    ConnectHost
    StartSend (6)
    StartSend (1)
    StartSend (3)
    StartSend (5)
End Sub

Private Sub ConnectHost()
    sockClient.RemoteHost = txtHost.text
    sockClient.RemotePort = txtPort.text
    sockClient.Connect
End Sub

Private Sub DisconnectHost()
    sockClient.Close
End Sub

Private Sub cmdConnect_Click()
    ConnectHost
    cmdConnect.Enabled = False
    cmdReset.Enabled = False
    cmdDisconnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
    DisconnectHost
    cmdConnect.Enabled = True
    cmdReset.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Xit
End Sub

Private Sub cmdReq_Click(Index As Integer)
    LocalPagControl.ReqIn = IIf(Index = 1, True, False)
End Sub

Private Sub cmdReset_Click()
    AkClientResetReq = True
    cmdConnect.Enabled = True
    cmdReset.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub cmdSend_Click(Index As Integer)
    Select Case Index
        Case Else
            strSend(idxLoad) = Chr(2) & Chr(32) & txtSend(Index).text & Chr(3)
            idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
    End Select
    SendData
End Sub

Private Sub StartSend(Idx As Integer)
    OKtoSend(Idx) = True
    cmdStart(Idx).Enabled = False
    cmdStop(Idx).Enabled = True
End Sub

Private Sub cmdStart_Click(Index As Integer)
    StartSend (Index)
End Sub

Private Sub StopSend(Idx As Integer)
    OKtoSend(Idx) = False
    cmdStart(Idx).Enabled = True
    cmdStop(Idx).Enabled = False
End Sub

Private Sub cmdStop_Click(Index As Integer)
    StopSend (Index)
End Sub

Private Sub Form_Load()
Dim Idx As Integer
Dim color As Long


    ' Server IP address
    
    If (Len(PAGSERVERIP) > 6) Then
        frmAkClient.txtHost.text = Mid(PAGSERVERIP, 1, Len(PAGSERVERIP))
    Else
        frmAkClient.txtHost.text = "127.0.0.1"
    End If
    
    For Idx = cmdStart.LBound To cmdStart.UBound
        cmdStart(Idx).Enabled = True
        cmdStop(Idx).Enabled = False
    Next Idx
    
    For Idx = cmdSend.LBound To cmdSend.UBound
        Select Case Idx
            Case 7
                ' ignore; do nothing
            Case 4, 7, 8, 12, 17, 19, 20, 21, 22, 23
                ' no cmdsend; unused; do nothing
            Case Else
                cmdSend(Idx).Picture = cmdSend(7).Picture
                cmdSend(Idx).DisabledPicture = cmdSend(7).DisabledPicture
                cmdSend(Idx).DownPicture = cmdSend(7).DownPicture
        End Select
    Next Idx
    
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
    txtPagStatus.text = ""
    txtPagStatus.ForeColor = color
    
    color = MEDGRAY
    shpReqIn.BackColor = color
    shpRdyOut.BackColor = color
    shpRequestIn.BackColor = color
    shpRequestOut.BackColor = color
    
    color = RoyalBlue
    For Idx = 1 To 3
        txtCurrentVals(Idx).ForeColor = color
    Next Idx

    color = RoyalBlue
    For Idx = 1 To 4
        txtCfg(Idx).ForeColor = color
    Next Idx
    
    color = DarkGray
    
    
    color = RoyalBlue
    txtMsg.ForeColor = color
    
    color = SOFTBLUE
'    lblVersion.ForeColor = color
'    lblVersion.Caption = USINGRELEASEDATE
    
    InitArraysEtc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub frmControls_DblClick()
    frmAkClient.Width = IIf(frmAkClient.Width = FullWidth, ShortWidth, FullWidth)
End Sub

Private Sub sockClient_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
    
    sockClient.GetData strData, vbString
    strData = Mid(strData, 3, (Len(strData) - 3))
    RcdString = strData
    RcdData_Parse strData
    RcdData_Read strData
    openCmd = False
    If (idxUnload <> idxLoad) Then SendData
End Sub

Private Sub RcdData_Parse(ByVal rcdStr As String)
Dim paramsStr, tmpStr, curChar As String
Dim iChar, iChar2, iParam, max As Integer
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
        max = Len(paramsStr)
        For iChar = 1 To max
            curChar = Mid(paramsStr, iChar, 1)
            If (curChar <> " ") Then
                tmpStr = tmpStr & curChar
            End If
            If (iChar = max) Or (curChar = " ") Then
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
Dim Idx As Integer
    Select Case RcdCmdCode
        
        Case "ASTZ"
            iStatusBox = 1
            ' PAG Status
            txtPagStatus.text = Params_String(1)
            MasterPagData.Status = Params_String(1)
            Select Case Params_String(1)
                Case "SRDY"
                    txtPagStatus.ForeColor = DarkOrange
                Case "STBY"
                    txtPagStatus.ForeColor = LTBLUE
                Case "SIDL"
                    txtPagStatus.ForeColor = DKGRAY
                Case "SERR"
                    txtPagStatus.ForeColor = MEDRED
                Case "SOFF"
                    txtPagStatus.ForeColor = SlateGray
            End Select
            ' Request In
            MasterPagData.ReqIn = IIf((CInt(Params_String(2)) = "1"), True, False)
            ' PAG is Ready
            MasterPagData.RdyOut = IIf((CInt(Params_String(3)) = "1"), True, False)
        
        Case "ATEM"
            iStatusBox = 1
            ' Temperature
            txtCurrentVals(1).text = Params_String(1)
            MasterPagData.Temperature = ValueFromText(txtCurrentVals(1).text)
            ' Humidity
            txtCurrentVals(2).text = Params_String(2)
            MasterPagData.Humidity = ValueFromText(txtCurrentVals(2).text)
            ' Moisture
            txtCurrentVals(3).text = Params_String(3)
            MasterPagData.Moisture = ValueFromText(txtCurrentVals(3).text)

        Case "ACFG"
            iStatusBox = 1
            ' Temp SP
            txtCfg(1).text = Params_String(1)
            MasterPagData.TempSP = ValueFromText(txtCfg(1).text)
            SysConfig.Temp_Target = ValueFromText(txtCfg(1).text)
            ' Moisture SP
            txtCfg(2).text = Params_String(2)
            MasterPagData.MoistSP = ValueFromText(txtCfg(2).text)
            SysConfig.Moisture_Target = ValueFromText(txtCfg(2).text)
            ' Temp Tolerance
            txtCfg(3).text = Params_String(3)
            MasterPagData.TempTol = ValueFromText(txtCfg(3).text)
            SysConfig.Tol_Temp = ValueFromText(txtCfg(3).text)
            ' Moisture Tolerance
            txtCfg(4).text = Params_String(4)
            MasterPagData.MoistTol = ValueFromText(txtCfg(4).text)
            SysConfig.Tol_Moisture = ValueFromText(txtCfg(4).text)
            
        Case Else
            iStatusBox = 0
    End Select
    txtStatus(iStatusBox).text = NowPrefixString & strData & vbCrLf & txtStatus(iStatusBox).text
    If (Len(txtStatus(iStatusBox).text) > 32000) Then txtStatus(iStatusBox).text = Mid(txtStatus(iStatusBox).text, 1, 24000)
End Sub

Private Sub tmrScreen_Timer()
Static cntr

    cntr = IIf((cntr < 100), cntr + 1, cntr)
    If (cntr = 7) Then StartClient
    
    cmdConnect.Enabled = IIf((sockClient.State = sckConnected), False, True)
    cmdDisconnect.Enabled = IIf((sockClient.State = sckClosed), False, True)
    
    txtMsg.text = Format(Now(), "YYYY MMMM D   hh:mm:ss") & vbCrLf & WinsockStateDesc(sockClient.State)
    
    shpReqIn.BackColor = IIf(MasterPagData.ReqIn, MEDGREEN, DK3ORANGE)
    shpRdyOut.BackColor = IIf(MasterPagData.RdyOut, MEDGREEN, DK3ORANGE)
    
    shpRequestIn.BackColor = IIf(LocalPagControl.ReqIn, MEDGREEN, DK3ORANGE)
    shpRequestOut.BackColor = IIf(LocalPagControl.ReqOut, MEDGREEN, DK3ORANGE)
    
    txtLoadIdx.text = Format(idxLoad, "##0")
    txtUnloadIdx.text = Format(idxUnload, "##0")
    
    PaComm_Flag = IIf((sockClient.State = sckConnected), True, False)
    
End Sub

Private Sub TmrSendRepeat_Timer(Index As Integer)
    If OKtoSend(Index) Then
        strSend(idxLoad) = Chr(2) & Chr(32) & txtSendRepeat(Index).text & Chr(3)
        idxLoad = IIf((idxLoad < 100), (idxLoad + 1), 1)
        If (Index = 1) Then
            If (LocalPagControl.ReqOut) Then
                txtSendRepeat(1).text = "SREQ K0"
            Else
                txtSendRepeat(1).text = "SNRQ K0"
            End If
        End If
    End If
    SendData
End Sub

Private Sub SendData()
    If (Not openCmd And (idxUnload <> idxLoad) And (sockClient.State = sckConnected)) Then
        sockClient.SendData strSend(idxUnload)
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

Private Sub Xit()
    Select Case LocalPagControl.Type
        Case pagClient
            'using AK Client
            frmAkClient.Hide
        Case pagMaster
            ' no AK Client
            Unload frmAkClient
            Set frmAkClient = Nothing
        Case pagNone, pagAlone
            ' no AK Client
            Unload frmAkClient
            Set frmAkClient = Nothing
    End Select
End Sub

Private Sub InitArraysEtc()

    ' *************************************************************************
    '
    ' sockClient state descriptions
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
    
'*********************************************************************************************
'*********************************************************************************************
'*********************************************************************************************
    SrvrIP = PAGSERVERIP
'*********************************************************************************************
'*********************************************************************************************
'*********************************************************************************************
End Sub

