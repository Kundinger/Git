VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIoMonitor 
   Caption         =   "I/O Monitor"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   525
   ClientWidth     =   15135
   Icon            =   "frmIoMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMystic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7710
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmIoMonitor.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   8415
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdPurgeAir 
      Caption         =   "PURGE AIR SOURCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   219
      ToolTipText     =   "Open Purge Air Source(s) Screen"
      Top             =   8415
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame frmMessage 
      Height          =   855
      Left            =   120
      TabIndex        =   217
      Top             =   7395
      Width           =   10260
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "messages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   30
         TabIndex        =   218
         Top             =   120
         Width           =   10185
      End
   End
   Begin VB.CommandButton cmdChiller 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9060
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmIoMonitor.frx":777C
      Style           =   1  'Graphical
      TabIndex        =   216
      ToolTipText     =   "Chiller Controls"
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.Frame frmCommonAI 
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
      Height          =   1875
      Left            =   10440
      TabIndex        =   159
      Top             =   7395
      Width           =   3675
      Begin MSComctlLib.TabStrip tabsAmbPAS 
         Height          =   735
         Left            =   2790
         TabIndex        =   213
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
         MultiRow        =   -1  'True
         Style           =   1
         TabFixedWidth   =   441
         TabFixedHeight  =   450
         TabMinWidth     =   441
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ambient"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "PAS"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblMoistDispl 
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
         Left            =   1440
         TabIndex        =   174
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label lblBarPresDispl 
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
         Left            =   1440
         TabIndex        =   173
         Top             =   945
         Width           =   705
      End
      Begin VB.Label lblLkChkPresDispl 
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
         Left            =   1440
         TabIndex        =   172
         Top             =   1215
         Width           =   705
      End
      Begin VB.Label lblTempDispl 
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
         Left            =   1440
         TabIndex        =   171
         Top             =   180
         Width           =   705
      End
      Begin VB.Label lblTemp 
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature"
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
         TabIndex        =   170
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lblLkChkPress 
         BackStyle       =   0  'Transparent
         Caption         =   "Pressure"
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
         TabIndex        =   169
         Top             =   1215
         Width           =   1095
      End
      Begin VB.Label lblBarPress 
         BackStyle       =   0  'Transparent
         Caption         =   "Barometer"
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
         TabIndex        =   168
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label lblMoistUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grains/lb"
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
         Left            =   2205
         TabIndex        =   167
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblTempUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "deg x"
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
         Left            =   2205
         TabIndex        =   166
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblHumidUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "% rH"
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
         Left            =   2205
         TabIndex        =   165
         Top             =   450
         Width           =   675
      End
      Begin VB.Label lblHumidDispl 
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
         Left            =   1440
         TabIndex        =   164
         Top             =   450
         Width           =   705
      End
      Begin VB.Label lblHumidity 
         BackStyle       =   0  'Transparent
         Caption         =   "Humidity"
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
         TabIndex        =   163
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lblPressUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "psig"
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
         Left            =   2205
         TabIndex        =   162
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label lblBaroUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "mBAR"
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
         Left            =   2205
         TabIndex        =   161
         Top             =   945
         Width           =   675
      End
      Begin VB.Label lblMoisture 
         BackStyle       =   0  'Transparent
         Caption         =   "Moisture"
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
         TabIndex        =   160
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      DisabledPicture =   "frmIoMonitor.frx":83BE
      DownPicture     =   "frmIoMonitor.frx":9000
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
      Left            =   14115
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmIoMonitor.frx":9C42
      Style           =   1  'Graphical
      TabIndex        =   158
      ToolTipText     =   "Quit"
      Top             =   8430
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14230
      Top             =   7920
   End
   Begin VB.Frame fraAnalogIO 
      Caption         =   "Analog I/O"
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
      Height          =   7155
      Left            =   7815
      TabIndex        =   1
      Top             =   120
      Width           =   7140
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   15
         Left            =   765
         TabIndex        =   152
         Text            =   "328.823"
         Top             =   3390
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   15
         Left            =   1620
         TabIndex        =   151
         Text            =   "199.9"
         Top             =   3390
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   15
         Left            =   2265
         TabIndex        =   150
         Text            =   "7654328"
         Top             =   3390
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   6300
         TabIndex        =   149
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   3360
         Width           =   680
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   14
         Left            =   765
         TabIndex        =   148
         Text            =   "328.823"
         Top             =   3030
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   14
         Left            =   1620
         TabIndex        =   147
         Text            =   "199.9"
         Top             =   3030
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   14
         Left            =   2265
         TabIndex        =   146
         Text            =   "7654328"
         Top             =   3030
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   6300
         TabIndex        =   145
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   3000
         Width           =   680
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   14
         Left            =   5595
         TabIndex        =   144
         Text            =   "199.9"
         Top             =   3030
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   15
         Left            =   5595
         TabIndex        =   143
         Text            =   "199.9"
         Top             =   3390
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   13
         Left            =   765
         TabIndex        =   138
         Text            =   "328.823"
         Top             =   2670
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   13
         Left            =   1620
         TabIndex        =   137
         Text            =   "199.9"
         Top             =   2670
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   13
         Left            =   2265
         TabIndex        =   136
         Text            =   "7654328"
         Top             =   2670
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   6300
         TabIndex        =   135
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   2640
         Width           =   680
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   12
         Left            =   765
         TabIndex        =   134
         Text            =   "328.823"
         Top             =   2310
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   12
         Left            =   1620
         TabIndex        =   133
         Text            =   "199.9"
         Top             =   2310
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   12
         Left            =   2265
         TabIndex        =   132
         Text            =   "7654328"
         Top             =   2310
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   6300
         TabIndex        =   131
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   2280
         Width           =   680
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   12
         Left            =   5595
         TabIndex        =   130
         Text            =   "199.9"
         Top             =   2310
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   13
         Left            =   5595
         TabIndex        =   129
         Text            =   "199.9"
         Top             =   2670
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   11
         Left            =   765
         TabIndex        =   124
         Text            =   "328.823"
         Top             =   1830
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   11
         Left            =   1620
         TabIndex        =   123
         Text            =   "199.9"
         Top             =   1830
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   11
         Left            =   2265
         TabIndex        =   122
         Text            =   "7654328"
         Top             =   1830
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   6300
         TabIndex        =   121
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   1800
         Width           =   680
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   10
         Left            =   765
         TabIndex        =   120
         Text            =   "328.823"
         Top             =   1470
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   10
         Left            =   1620
         TabIndex        =   119
         Text            =   "199.9"
         Top             =   1470
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   10
         Left            =   2265
         TabIndex        =   118
         Text            =   "7654328"
         Top             =   1470
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   6300
         TabIndex        =   117
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   1440
         Width           =   680
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   5595
         TabIndex        =   116
         Text            =   "199.9"
         Top             =   1470
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   11
         Left            =   5595
         TabIndex        =   115
         Text            =   "199.9"
         Top             =   1830
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   9
         Left            =   765
         TabIndex        =   110
         Text            =   "328.823"
         Top             =   1110
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   9
         Left            =   1620
         TabIndex        =   109
         Text            =   "199.9"
         Top             =   1110
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   9
         Left            =   2265
         TabIndex        =   108
         Text            =   "7654328"
         Top             =   1110
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   6300
         TabIndex        =   107
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   1080
         Width           =   680
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   8
         Left            =   765
         TabIndex        =   106
         Text            =   "328.823"
         Top             =   750
         Width           =   860
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   8
         Left            =   1620
         TabIndex        =   105
         Text            =   "199.9"
         Top             =   750
         Width           =   640
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   8
         Left            =   2265
         TabIndex        =   104
         Text            =   "7654328"
         Top             =   750
         Width           =   900
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   6300
         TabIndex        =   103
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   720
         Width           =   680
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   5595
         TabIndex        =   102
         Text            =   "199.9"
         ToolTipText     =   "0-100 Percent of Max"
         Top             =   750
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   5595
         TabIndex        =   101
         Text            =   "199.9"
         Top             =   1110
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   5595
         TabIndex        =   100
         Text            =   "199.9"
         Top             =   4230
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   5595
         TabIndex        =   99
         Text            =   "199.9"
         Top             =   4590
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   5595
         TabIndex        =   98
         Text            =   "199.9"
         Top             =   4950
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   5595
         TabIndex        =   97
         Text            =   "199.9"
         Top             =   3870
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   5595
         TabIndex        =   96
         Text            =   "199.9"
         Top             =   5430
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   5595
         TabIndex        =   95
         Text            =   "199.9"
         Top             =   5790
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   5595
         TabIndex        =   94
         Text            =   "199.9"
         Top             =   6150
         Width           =   640
      End
      Begin VB.TextBox txtWritePerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   5595
         TabIndex        =   93
         Text            =   "199.9"
         Top             =   6510
         Width           =   640
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   6300
         TabIndex        =   92
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   6480
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   7
         Left            =   2265
         TabIndex        =   91
         Text            =   "7654328"
         Top             =   6510
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   7
         Left            =   1620
         TabIndex        =   90
         Text            =   "199.9"
         Top             =   6510
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   7
         Left            =   765
         TabIndex        =   87
         Text            =   "328.823"
         Top             =   6510
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   6300
         TabIndex        =   86
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   6120
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   6
         Left            =   2265
         TabIndex        =   85
         Text            =   "7654328"
         Top             =   6150
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   6
         Left            =   1620
         TabIndex        =   84
         Text            =   "199.9"
         Top             =   6150
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   6
         Left            =   765
         TabIndex        =   81
         Text            =   "328.823"
         Top             =   6150
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   6300
         TabIndex        =   80
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   5760
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   5
         Left            =   2265
         TabIndex        =   79
         Text            =   "7654328"
         Top             =   5790
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   5
         Left            =   1620
         TabIndex        =   78
         Text            =   "199.9"
         Top             =   5790
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   5
         Left            =   765
         TabIndex        =   75
         Text            =   "328.823"
         Top             =   5790
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   6300
         TabIndex        =   74
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   5400
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   4
         Left            =   2265
         TabIndex        =   73
         Text            =   "7654328"
         Top             =   5430
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   4
         Left            =   1620
         TabIndex        =   72
         Text            =   "199.9"
         Top             =   5430
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   4
         Left            =   765
         TabIndex        =   69
         Text            =   "328.823"
         Top             =   5430
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   6300
         TabIndex        =   68
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   3840
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   0
         Left            =   2265
         TabIndex        =   67
         Text            =   "7654328"
         Top             =   3870
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   0
         Left            =   1620
         TabIndex        =   66
         Text            =   "199.9"
         Top             =   3870
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   0
         Left            =   765
         TabIndex        =   63
         Text            =   "328.823"
         Top             =   3870
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6300
         TabIndex        =   62
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   4920
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   3
         Left            =   2265
         TabIndex        =   61
         Text            =   "7654328"
         Top             =   4950
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   3
         Left            =   1620
         TabIndex        =   60
         Text            =   "199.9"
         Top             =   4950
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   3
         Left            =   765
         TabIndex        =   57
         Text            =   "328.823"
         Top             =   4950
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6300
         TabIndex        =   56
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   4560
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   2
         Left            =   2265
         TabIndex        =   55
         Text            =   "7654328"
         Top             =   4590
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   2
         Left            =   1620
         TabIndex        =   54
         Text            =   "199.9"
         Top             =   4590
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Index           =   2
         Left            =   765
         TabIndex        =   51
         Text            =   "328.823"
         Top             =   4590
         Width           =   860
      End
      Begin VB.CommandButton cmdSetOut 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6300
         TabIndex        =   50
         ToolTipText     =   "Set analog output to entered percent value"
         Top             =   4200
         Width           =   680
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Left            =   2265
         TabIndex        =   49
         Text            =   "7654328"
         Top             =   4230
         Width           =   900
      End
      Begin VB.TextBox txtReadPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Left            =   1620
         TabIndex        =   48
         Text            =   "199.9"
         Top             =   4230
         Width           =   640
      End
      Begin VB.TextBox txtReadEU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
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
         Left            =   765
         TabIndex        =   45
         Text            =   "328.823"
         Top             =   4230
         Width           =   860
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   15
         Left            =   3195
         TabIndex        =   156
         Top             =   3420
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   15
         Left            =   180
         TabIndex        =   155
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   14
         Left            =   3195
         TabIndex        =   154
         Top             =   3060
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   14
         Left            =   180
         TabIndex        =   153
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   13
         Left            =   3195
         TabIndex        =   142
         Top             =   2700
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   13
         Left            =   180
         TabIndex        =   141
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   12
         Left            =   3195
         TabIndex        =   140
         Top             =   2340
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   12
         Left            =   180
         TabIndex        =   139
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   11
         Left            =   3195
         TabIndex        =   128
         Top             =   1860
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   11
         Left            =   180
         TabIndex        =   127
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   10
         Left            =   3195
         TabIndex        =   126
         Top             =   1500
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   10
         Left            =   180
         TabIndex        =   125
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   9
         Left            =   3195
         TabIndex        =   114
         Top             =   1140
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   9
         Left            =   180
         TabIndex        =   113
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   8
         Left            =   3195
         TabIndex        =   112
         Top             =   780
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   8
         Left            =   180
         TabIndex        =   111
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   7
         Left            =   180
         TabIndex        =   89
         Top             =   6540
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   7
         Left            =   3195
         TabIndex        =   88
         Top             =   6540
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   6
         Left            =   180
         TabIndex        =   83
         Top             =   6180
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   6
         Left            =   3195
         TabIndex        =   82
         Top             =   6180
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   5
         Left            =   180
         TabIndex        =   77
         Top             =   5820
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   5
         Left            =   3195
         TabIndex        =   76
         Top             =   5820
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   4
         Left            =   180
         TabIndex        =   71
         Top             =   5460
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   4
         Left            =   3195
         TabIndex        =   70
         Top             =   5460
         Width           =   2310
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   2280
         TabIndex        =   41
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblReadPerc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   1785
         TabIndex        =   40
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   0
         Left            =   180
         TabIndex        =   65
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   0
         Left            =   3195
         TabIndex        =   64
         Top             =   3900
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   3
         Left            =   180
         TabIndex        =   59
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   3
         Left            =   3195
         TabIndex        =   58
         Top             =   4980
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   2
         Left            =   180
         TabIndex        =   53
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   2
         Left            =   3195
         TabIndex        =   52
         Top             =   4620
         Width           =   2310
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   1
         Left            =   180
         TabIndex        =   47
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label descAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   1
         Left            =   3195
         TabIndex        =   46
         Top             =   4260
         Width           =   2310
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   180
         TabIndex        =   44
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChanDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   3165
         TabIndex        =   43
         Top             =   480
         Width           =   2205
      End
      Begin VB.Label lblWritePerc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   5790
         TabIndex        =   31
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblSet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   6285
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame fraDigOutp 
      Caption         =   "Digital I/O"
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
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.Frame frmlegend 
         Height          =   495
         Left            =   120
         TabIndex        =   208
         Top             =   6540
         Width           =   7455
         Begin VB.Label lblNoModule 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "no module"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   4440
            TabIndex        =   212
            Top             =   180
            Width           =   975
         End
         Begin VB.Label lblOutput 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "output"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   3300
            TabIndex        =   211
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "input"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   2280
            TabIndex        =   210
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "legend"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   240
            TabIndex        =   209
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   3975
         TabIndex        =   190
         Top             =   6060
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   3975
         TabIndex        =   189
         Top             =   5700
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   3975
         TabIndex        =   188
         Top             =   5340
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   3975
         TabIndex        =   187
         Top             =   4980
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   3975
         TabIndex        =   186
         Top             =   4620
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   3975
         TabIndex        =   185
         Top             =   4260
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   3975
         TabIndex        =   184
         Top             =   3900
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   3975
         TabIndex        =   183
         Top             =   3540
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   3975
         TabIndex        =   182
         Top             =   3180
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   3975
         TabIndex        =   181
         Top             =   2820
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   3975
         TabIndex        =   180
         Top             =   2460
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   3975
         TabIndex        =   179
         Top             =   2100
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   3975
         TabIndex        =   178
         Top             =   1740
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   3975
         TabIndex        =   177
         Top             =   1380
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   3975
         TabIndex        =   176
         Top             =   1020
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   3975
         TabIndex        =   175
         Top             =   660
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   32
         Top             =   6060
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   33
         Top             =   5700
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   34
         Top             =   5340
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "3 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   35
         Top             =   4980
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   27
         Top             =   4620
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   25
         Top             =   4260
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   23
         Top             =   3900
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "2 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   3540
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   3180
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   17
         Top             =   2820
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2460
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "1 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   2100
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1740
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1380
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1020
         Width           =   855
      End
      Begin VB.CommandButton cmdModChan 
         Caption         =   "0 / 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   855
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   19
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   28
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Idle"
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
         Index           =   31
         Left            =   5325
         TabIndex        =   207
         Top             =   6120
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   29
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "ADF Fill Valve"
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
         Index           =   30
         Left            =   5325
         TabIndex        =   206
         Top             =   5760
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   30
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Aux Vent"
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
         Index           =   29
         Left            =   5325
         TabIndex        =   205
         Top             =   5400
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   31
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   6120
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "ADF Pump"
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
         Index           =   28
         Left            =   5325
         TabIndex        =   204
         Top             =   5040
         Width           =   2205
      End
      Begin VB.Label lblModChanTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Addr/Chan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   3960
         TabIndex        =   203
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Pause"
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
         Index           =   27
         Left            =   5325
         TabIndex        =   202
         Top             =   4680
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   27
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2 Purge"
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
         Index           =   26
         Left            =   5325
         TabIndex        =   201
         Top             =   4320
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   26
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2 Vent"
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
         Index           =   25
         Left            =   5325
         TabIndex        =   200
         Top             =   3960
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   25
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2 Load"
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
         Index           =   24
         Left            =   5325
         TabIndex        =   199
         Top             =   3600
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   24
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "ADF Drain Valve"
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
         Index           =   23
         Left            =   5325
         TabIndex        =   198
         Top             =   3240
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   23
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Purge"
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
         Index           =   22
         Left            =   5325
         TabIndex        =   197
         Top             =   2880
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   22
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check"
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
         Index           =   21
         Left            =   5325
         TabIndex        =   196
         Top             =   2520
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   21
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux. Can ( Vent )"
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
         Index           =   20
         Left            =   5325
         TabIndex        =   195
         Top             =   2160
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   20
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Direction (Purge)"
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
         Index           =   19
         Left            =   5325
         TabIndex        =   194
         Top             =   1800
         Width           =   2205
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge"
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
         Index           =   18
         Left            =   5325
         TabIndex        =   193
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   18
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane"
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
         Index           =   17
         Left            =   5325
         TabIndex        =   192
         Top             =   1080
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   17
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   16
         Left            =   5325
         TabIndex        =   191
         Top             =   720
         Width           =   2205
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   16
         Left            =   4980
         Shape           =   3  'Circle
         Top             =   720
         Width           =   255
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   12
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Idle"
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
         Index           =   15
         Left            =   1580
         TabIndex        =   39
         Top             =   6120
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   13
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "ADF Fill Valve"
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
         Index           =   14
         Left            =   1580
         TabIndex        =   38
         Top             =   5760
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   14
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   5760
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Aux Vent"
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
         Index           =   13
         Left            =   1580
         TabIndex        =   37
         Top             =   5400
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   15
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   6120
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "ADF Pump"
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
         Index           =   12
         Left            =   1580
         TabIndex        =   36
         Top             =   5040
         Width           =   2200
      End
      Begin VB.Label lblModChanTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Addr/Chan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Index           =   0
         Left            =   220
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Pause"
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
         Index           =   11
         Left            =   1580
         TabIndex        =   28
         Top             =   4680
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   11
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2 Purge"
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
         Index           =   10
         Left            =   1580
         TabIndex        =   26
         Top             =   4320
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   10
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2 Vent"
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
         Index           =   9
         Left            =   1580
         TabIndex        =   24
         Top             =   3960
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   9
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift 2 Load"
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
         Index           =   8
         Left            =   1580
         TabIndex        =   22
         Top             =   3600
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   8
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "ADF Drain Valve"
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
         Index           =   7
         Left            =   1580
         TabIndex        =   20
         Top             =   3240
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux Purge"
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
         Index           =   6
         Left            =   1580
         TabIndex        =   18
         Top             =   2880
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check"
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
         Index           =   5
         Left            =   1580
         TabIndex        =   16
         Top             =   2520
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Aux. Can ( Vent )"
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
         Index           =   4
         Left            =   1580
         TabIndex        =   14
         Top             =   2160
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Direction (Purge)"
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
         Index           =   3
         Left            =   1580
         TabIndex        =   12
         Top             =   1800
         Width           =   2200
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge"
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
         Index           =   2
         Left            =   1580
         TabIndex        =   10
         Top             =   1440
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane"
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
         Index           =   1
         Left            =   1580
         TabIndex        =   8
         Top             =   1080
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblModChan 
         BackStyle       =   0  'Transparent
         Caption         =   "01234567890123456789"
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
         Index           =   0
         Left            =   1580
         TabIndex        =   6
         Top             =   720
         Width           =   2200
      End
      Begin VB.Shape shpModChan 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   1245
         Shape           =   3  'Circle
         Top             =   720
         Width           =   255
      End
   End
   Begin Threed.SSPanel txtDispStn 
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Station Number Displayed"
      Top             =   8415
      Width           =   2130
      _Version        =   65536
      _ExtentX        =   3757
      _ExtentY        =   1508
      _StockProps     =   15
      Caption         =   "Board #9"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdUp 
      Height          =   855
      Left            =   5370
      TabIndex        =   3
      ToolTipText     =   "Next"
      Top             =   8415
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1508
      _StockProps     =   78
      ForeColor       =   8421504
      BevelWidth      =   3
      Outline         =   0   'False
      AutoSize        =   1
      Picture         =   "frmIoMonitor.frx":A884
   End
   Begin Threed.SSCommand cmdDown 
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Previous"
      Top             =   8415
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1508
      _StockProps     =   78
      ForeColor       =   8421504
      BevelWidth      =   3
      Outline         =   0   'False
      AutoSize        =   1
      Picture         =   "frmIoMonitor.frx":10076
   End
   Begin Threed.SSPanel lblMode 
      Height          =   855
      Left            =   120
      TabIndex        =   157
      Top             =   8415
      Width           =   2100
      _Version        =   65536
      _ExtentX        =   3704
      _ExtentY        =   1508
      _StockProps     =   15
      Caption         =   "Automatic"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
   End
   Begin VB.Label lblDebug2 
      Alignment       =   2  'Center
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
      Left            =   15720
      TabIndex        =   215
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label lblDebug1 
      Alignment       =   2  'Center
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
      Left            =   15720
      TabIndex        =   214
      Top             =   3360
      Width           =   900
   End
End
Attribute VB_Name = "frmIoMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' The current board being displayed
Private CurrentBoard As Integer

Private Sub Clear_Buttons()
    Dim inc As Integer
    For inc = 0 To 16
       shpModChan(inc).BackColor = DK3ORANGE
    Next inc
End Sub

Private Sub cmdChiller_Click()
    lblMessage.Caption = " "
    frmControllers.Show
End Sub

Private Sub cmdDown_Click()
    ' This command decrements the Board number variable,
    ' the Board number displayed, and triggers an update for
    ' the values displayed on the form for the current Board
    If CurrentBoard > 0 Then
        CurrentBoard = CurrentBoard - 1
    Else
        CurrentBoard = NR_STN
    End If
    lblMessage.Caption = " "
    txtDispStn.Caption = "Board #" & Format(CurrentBoard, "0")
    If STN_IOForceMode(CurrentBoard) = VBAUTO Then AutoMode
    If STN_IOForceMode(CurrentBoard) = VBMANUAL Then ManualMode
    SetupButtons (CurrentBoard)
    UpdateReadouts
End Sub

Private Sub cmdExit_Click()
    ExitScreen
End Sub

Private Sub cmdModChan_Click(Index As Integer)
'
' Toggle a Digital Output
'
Dim address As Integer
Dim channel As Integer
    lblMessage.Caption = " "
    If STN_IOForceMode(CurrentBoard) <> VBMANUAL Then
        lblMessage.Caption = vbCrLf & "Must be Idle for Manual Mode"
    Else
        address = CInt(CurrentBoard * 4)
        channel = CInt(Index)
        If Index > 15 Then address = address + 1
        If Index > 15 Then channel = Index - 16
        If OptoDIO(address, channel).Type = 2 Then
            If OptoDIO(address, channel).RawValue Then
                OPTO_WriteDigital address, channel, cOFF
            Else
                OPTO_WriteDigital address, channel, cON
            End If
        Else
            lblMessage.Caption = vbCrLf & "Channel is Not an Output"
        End If
    End If
End Sub

Private Sub cmdMystic_Click()
    lblMessage.Caption = " "
    frmMainForm.Show
End Sub

Private Sub cmdPurgeAir_Click()
    lblMessage.Caption = " "
    frmPrgMonitor.Show
End Sub

Private Sub cmdSetOut_Click(Index As Integer)
    ' The Command Flow button was clicked
    ' If a field is out of range, the field is highlighted
    ' in yellow and an error message appears.
    Dim PercMax As Single
    Dim PercMin As Single
    Dim tempPerc, tempVdc As Single
    Dim tempRaw As Long
    Dim addr, chan As Integer
    
    lblMessage.Caption = " "
    If STN_IOForceMode(CurrentBoard) <> VBMANUAL Then
        lblMessage.Caption = vbCrLf & "Station is Running"
        Exit Sub
    End If
        
    ' ************************************************************
    If IsNumeric(txtWritePerc(Index).text) Then
        txtWritePerc(Index).BackColor = Entry_BackColor
        PercMax = 100#
        PercMin = 0#
        tempPerc = CSng(txtWritePerc(Index).text)
        tempPerc = IIf(tempPerc > PercMax, PercMax, tempPerc)
        tempPerc = IIf(tempPerc < PercMin, PercMin, tempPerc)
        txtWritePerc(Index).text = Format(tempPerc, "##0.0")
        
        Select Case Node_Info(CurrentBoard)
            Case 8
                If Index < 8 Then
                    addr = (4 * CurrentBoard) + 2
                    chan = Index + 8
                Else
                    addr = (4 * CurrentBoard) + 2
                    chan = Index - 8
                End If
            Case 12
                If Index < 8 Then
                    addr = (4 * CurrentBoard) + 3
                    chan = Index
                Else
                    addr = (4 * CurrentBoard) + 2
                    chan = Index
                End If
            Case 16
                If Index < 8 Then
                    addr = (4 * CurrentBoard) + 3
                    chan = Index + 8
                Else
                    addr = (4 * CurrentBoard) + 3
                    chan = Index - 8
                End If
            Case Else
                addr = 0
                chan = 0
        End Select
        ' Convert Percent to Raw Counts
        If (Map_AIO(addr, chan).VdcMax > Map_AIO(addr, chan).VdcMin) Then
            tempVdc = (tempPerc / 100#) * (Map_AIO(addr, chan).VdcMax - Map_AIO(addr, chan).VdcMin)   ' Vdc above VdcMin
            tempVdc = tempVdc + Map_AIO(addr, chan).VdcMin                    ' Vdc out of 0-10Vdc
            tempRaw = (tempVdc / 10#) * FULLSCALE
        Else
            tempRaw = (tempPerc / 100#) * FULLSCALE
        End If
        
        OPTO_WriteAnalog CInt(addr), CInt(chan), CLng(tempRaw)
    
    Else
        ' ***** only numeric values are allowed
        ' Mark the box yellow
        txtWritePerc(Index).BackColor = PALEYELLOW
        txtWritePerc(Index).text = "0.0"
        lblMessage.Caption = vbCrLf & "Only numeric values between 0.0 and 100.0 are accepted"
    
    End If

End Sub

Private Sub cmdUp_Click()
    ' This command increments the Board number variable,
    ' the Board number displayed, and triggers an update for
    ' the values displayed on the form for the current Board
    
    If CurrentBoard < NR_STN Then
        CurrentBoard = CurrentBoard + 1
    Else
        CurrentBoard = 0
    End If
    lblMessage.Caption = " "
    txtDispStn.Caption = "Board #" & Format(CurrentBoard, "0")
    If STN_IOForceMode(CurrentBoard) = VBAUTO Then AutoMode
    If STN_IOForceMode(CurrentBoard) = VBMANUAL Then ManualMode
    SetupButtons (CurrentBoard)
    UpdateReadouts
End Sub

Private Sub Form_Load()
Dim Idx As Integer
    ' Set Title Foreground color
    fraAnalogIO.ForeColor = Titles_ForeColor
    fraDigOutp.ForeColor = Titles_ForeColor
    lblMode.ForeColor = TitlesData_Forecolor
    txtDispStn.ForeColor = TitlesData_Forecolor
    lblMessage.ForeColor = DKPURPLE
    
    ' initialize the message
    lblMessage.Caption = " "
    
    ' Set the current Board number to 0
    CurrentBoard = 0
    txtDispStn.Caption = "Board #" & Format(CurrentBoard, "0")
        
'   Set Mode
    If STN_IOForceMode(CurrentBoard) = VBAUTO Then AutoMode
    If STN_IOForceMode(CurrentBoard) = VBMANUAL Then ManualMode
    
    ' Setup Digital IO Buttons
    SetupButtons CurrentBoard
    
    ' Make Optional Buttons Not Visible (to start with)
    cmdMystic.Visible = False
    cmdPurgeAir.Visible = False
    cmdChiller.Visible = False
    
    ' Set AO Write % to 0.0
    For Idx = txtWritePerc.LBound To txtWritePerc.UBound
        txtWritePerc(Idx).text = "0.0"
    Next Idx
            
    ' TEMPERATURE
    If USINGC Then
        lblTempUnits.Caption = "deg C"
    ElseIf USINGF Then
        lblTempUnits.Caption = "deg F"
    End If
    ' MOISTURE
    If USINGMoist_RH Then
        lblMoistUnits.Caption = "% rH"
    ElseIf USINGMoist_Grains Then
        lblMoistUnits.Caption = "grains/lb"
    End If
    
    ' update readouts
    UpdateReadouts
    
    ' Start the timer
    tmrUpdate.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub ExitScreen()
    IOForceActive = False
    tmrUpdate.Enabled = False
    Unload Me
    Set frmIoMonitor = Nothing
End Sub

Private Sub UpdateReadouts()
    
Dim Index, addr, chan, addroffset As Integer
Dim tempPerc, tempVdc, tempTemp As Single
    
        
    ' no board; nothing to update
    If Node_Info(CurrentBoard) > 0 Then
        fraAnalogIO.Top = 120
        fraDigOutp.Top = 120
    Else
        fraAnalogIO.Top = OutOfSight
        fraDigOutp.Top = OutOfSight
        Exit Sub
    End If
          
    ' Board Base Address = BoardNumber * 4
    ' Base + 0 Digitals
    addr = CurrentBoard * 4
    For chan = 0 To 15
        Index = chan
        If OptoDIO(addr, chan).RawValue Then
            ' Set the color to green if on
            shpModChan(Index).FillColor = MEDGREEN
        Else
            ' Set the color to dark if off
            shpModChan(Index).FillColor = DK3ORANGE
        End If
    Next chan
    
    ' Base + 1 Digitals
    addr = addr + 1
    For chan = 0 To 15
        Index = chan + 16
        If OptoDIO(addr, chan).RawValue Then
            ' Set the color to green if on
            shpModChan(Index).FillColor = MEDGREEN
        Else
            ' Set the color to dark if off
            shpModChan(Index).FillColor = DK3ORANGE
        End If
    Next chan
    

    ' Analogs
    For chan = 0 To 15
        Select Case Node_Info(CurrentBoard)
            Case 8
                ' chan 0-15 on base+2
                addroffset = 2
                addr = (4 * CurrentBoard) + addroffset
                Index = IIf(chan > 7, chan - 8, chan + 8)
            Case 12
                If chan > 7 Then
                    ' chan 8-15 on base+2
                    addroffset = 2
                    addr = (4 * CurrentBoard) + addroffset
                    Index = chan
                Else
                    ' chan 0-7 on base+3
                    addroffset = 3
                    addr = (4 * CurrentBoard) + addroffset
                    Index = chan
                End If
            Case 16
                ' chan 0-15 on base+3
                addroffset = 3
                addr = (4 * CurrentBoard) + addroffset
                Index = IIf(chan > 7, chan - 8, chan + 8)
            Case Else
                Exit For
        End Select
        txtReadRaw(Index) = Format(OptoAIO(addr, chan).RawValue, "#,###,##0")
        Select Case OptoAIO(addr, chan).Type
            Case 3, 4
                ' AI & AO 0-10 vdc
                If (Map_AIO(addr, chan).VdcMax > Map_AIO(addr, chan).VdcMin) Then
                    tempVdc = 10# * (Map_AIO(addr, chan).RawValue / FULLSCALE)          ' Vdc out of 0-10Vdc
                    tempVdc = tempVdc - Map_AIO(addr, chan).VdcMin                      ' Vdc above VdcMin
                    tempPerc = tempVdc / (Map_AIO(addr, chan).VdcMax - Map_AIO(addr, chan).VdcMin)
                    tempPerc = CSng(100# * tempPerc)
                    If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                        txtReadEU(Index).Visible = True
                        txtReadEU(Index).text = Format(Map_AIO(addr, chan).EUValue, "#,##0.000")
                    Else
                        txtReadEU(Index).Visible = False
                    End If
                Else
                    txtReadEU(Index).Visible = False
                    tempPerc = CSng(txtReadRaw(Index)) * (100# / FULLSCALE)
                End If
                txtReadPerc(Index).text = Format(tempPerc, "##0.0")
            Case 5, 6
                ' TCs
                tempTemp = CSng(10 * (OptoAIO(addr, chan).RawValue / FULLSCALE))
                tempTemp = tempTemp + Map_AIO(addr, chan).EuMin
                tempTemp = IIf(USINGC, tempTemp, DegCtoF(tempTemp))
                txtReadEU(Index).text = Format(tempTemp, "#,##0.0")
                txtReadEU(Index).Visible = True
                txtReadPerc(Index).Visible = False
            Case 7
                ' RTD's
                tempTemp = CSng(10 * (OptoAIO(addr, chan).RawValue / FULLSCALE))
                tempTemp = tempTemp + Map_AIO(addr, chan).EuMin
                tempTemp = IIf(USINGC, tempTemp, DegCtoF(tempTemp))
                txtReadEU(Index).text = Format(tempTemp, "#,##0.0")
                txtReadEU(Index).Visible = True
                txtReadPerc(Index).Visible = False
        End Select
    Next chan
        
End Sub

Private Sub ConfigIOButtons(BoolVal As Boolean)
Dim i As Integer
    ' Enable or disable the IO buttons
    For i = cmdModChan.LBound To cmdModChan.UBound
        cmdModChan(i).Enabled = BoolVal
    Next i
End Sub

Private Sub ManualMode()
       
    ' Change the mode label color to slate
    lblMode.BackColor = Warning_ForeColor
    lblMode.Caption = "Manual"
    
    ' Enable the command buttons
    ConfigIOButtons True

    IOForceActive = True

End Sub

Private Sub AutoMode()
    
    ' Change the mode label color to slate
    lblMode.BackColor = Common_BackColor
    lblMode.Caption = "Automatic"
        
    ' Disable the command buttons
    ConfigIOButtons False
    
    IOForceActive = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExitScreen
End Sub


Private Sub tmrUpdate_Timer()
    ' Update and display I/O values for this form
'lblDebug1.Caption = Format(Map_AIO(Stn_AIO(1, asPurgeDiffPress).addr, Stn_AIO(1, asPurgeDiffPress).chan).RawValue, "######0")
'lblDebug2.Caption = Format(Map_AIO(Stn_AIO(2, asPurgeDiffPress).addr, Stn_AIO(2, asPurgeDiffPress).chan).RawValue, "######0")
    ' Opto-22 button for sysdef users
    cmdMystic.Visible = IIf(CheckPass("H", False), True, False)
    ' PurgeMonitor button for controller users
    cmdPurgeAir.Visible = IIf(CheckPass("C", False), True, False)
    ' Chiller Control button for controller users
    cmdChiller.Visible = IIf((USINGWATERBATH And (CheckPass("C", False))), True, False)
    
    Select Case tabsAmbPAS.SelectedItem
        Case "Ambient"
            ' Temperature
            lblTempDispl.Caption = Format(AmbTemp, "##0.0")
            ' Humidity
            lblHumidDispl.Caption = Format(AmbHum, "##0.0")         ' Display Humidity in PerCent RH
            ' Moisture
            lblMoistDispl.Caption = Format(AmbMoisture, "###0.0")   ' Display Moisture as either % rH or grains/lb
        Case "PAS"
            ' Temperature
            lblTempDispl.Caption = Format(PATemp, "##0.0")
            ' Humidity
            lblHumidDispl.Caption = Format(PAHum, "##0.0")          ' Display Humidity in PerCent RH
            ' Moisture
            lblMoistDispl.Caption = Format(PAMoisture, "###0.0")    ' Display Moisture as either % rH or grains/lb
    End Select
    ' Pressure
    lblLkChkPresDispl.Caption = Format(PTinvalue, "###0.00")        ' Display Leak Press in psig
    ' Barometer
    lblBarPresDispl.Caption = Format(AmbBaro, "#000")               ' Display Baro in mBar
    
    ' Update display of analog & digital I/O values for this form
    UpdateReadouts
    

End Sub

Private Sub SetupButtons(Board As Integer)
'   This function sets up the optional I/O buttons
'       and optional analog descriptors, etc.
'
Dim addr As Integer
Dim chan As Integer
Dim chn As Integer
Dim module As Integer
Dim addroffset As Integer
Dim flag As Boolean

    flag = IIf((USINGPASLOCALCONTROL Or (LocalPagControl.Type = pagClient) Or (USINGDRYPURGEAIR And SysConfig.DryAirPurge)), True, False)
    tabsAmbPAS.Visible = flag
    tabsAmbPAS.Enabled = flag
    tabsAmbPAS.SelectedItem = "Ambient"
    
    If STN_IOForceMode(Board) = VBAUTO Then Clear_Buttons
    
    '   Channel Functional Descriptors
    SetupOpto
    '       Digital Channels
    lblInput.ForeColor = DK3CYAN
    lblOutput.ForeColor = DKBLUE
    lblNoModule.ForeColor = DKGRAY
    lblInput.Caption = "input"
    lblOutput.Caption = "output"
    lblNoModule.Caption = "no module"
    For addroffset = 0 To 1
        addr = (4 * Board) + addroffset
        For chan = 0 To MAX_CHAN
            module = Int(chan / 4)
            chn = chan + (addroffset * 16)
            If chn <= cmdModChan.UBound Then
                lblModChan(chn).Caption = Left(OptoChanDesc(addr, chan), 24)
                Select Case OptoDIO(addr, chan).Type
                    Case 0
                        ' no module
                        lblModChan(chn).ForeColor = lblNoModule.ForeColor
                    Case 1
                        ' DI
                        lblModChan(chn).ForeColor = lblInput.ForeColor
                    Case 2
                        ' DO
                        lblModChan(chn).ForeColor = lblOutput.ForeColor
                    Case Else
                        ' unknown
                        lblModChan(chn).ForeColor = PALERED
                End Select
                cmdModChan(chn).Caption = Format(addr, "#0") & " / " & Format(chan, "#0")
            End If
        Next chan
    Next addroffset
        
    
    '       Analog Channels
    For chan = 0 To MAX_CHAN
        Select Case Node_Info(CurrentBoard)
            Case 8
                ' chan 0-15 on base+2
                addr = (4 * Board) + 2
                chn = IIf(chan > 7, chan - 8, chan + 8)
            Case 12
                If chan > 7 Then
                    ' chan 8-15 on base+2
                    addr = (4 * Board) + 2
                    chn = chan
                Else
                    ' chan 0-7 on base+3
                    addr = (4 * Board) + 3
                    chn = chan
                End If
            Case 16
                ' chan 0-15 on base+3
                addr = (4 * Board) + 3
                chn = IIf(chan > 7, chan - 8, chan + 8)
            Case Else
                Exit For
        End Select
        Select Case OptoAIO(addr, chan).Type
            Case 3      ' AI
                descAddrChan(chn).ForeColor = lblInput.ForeColor
                lblAddrChan(chn).Visible = True
                descAddrChan(chn).Visible = True
                txtWritePerc(chn).Visible = False
                cmdSetOut(chn).Visible = False
                txtReadPerc(chn).Visible = True
                txtReadRaw(chn).Visible = True
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU(chn).Visible = True
                Else
                    txtReadEU(chn).Visible = False
                End If
            Case 4      ' AO
                descAddrChan(chn).ForeColor = lblOutput.ForeColor
                lblAddrChan(chn).Visible = True
                descAddrChan(chn).Visible = True
                If STN_IOForceMode(Board) = VBMANUAL Then
                    txtWritePerc(chn).Visible = True
                    cmdSetOut(chn).Visible = True
                Else
                    txtWritePerc(chn).Visible = False
                    cmdSetOut(chn).Visible = False
                End If
                txtReadPerc(chn).Visible = True
                txtReadRaw(chn).Visible = True
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU(chn).Visible = True
                Else
                    txtReadEU(chn).Visible = False
                End If
           Case 5      ' TC Type J
                descAddrChan(chn).ForeColor = lblInput.ForeColor
                lblAddrChan(chn).Visible = True
                descAddrChan(chn).Visible = True
                txtWritePerc(chn).Visible = False
                cmdSetOut(chn).Visible = False
                txtReadPerc(chn).Visible = True
                txtReadRaw(chn).Visible = True
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU(chn).Visible = True
                Else
                    txtReadEU(chn).Visible = False
                End If
            Case 6      ' TC Type K
                descAddrChan(chn).ForeColor = lblInput.ForeColor
                lblAddrChan(chn).Visible = True
                descAddrChan(chn).Visible = True
                txtWritePerc(chn).Visible = False
                cmdSetOut(chn).Visible = False
                txtReadPerc(chn).Visible = True
                txtReadRaw(chn).Visible = True
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU(chn).Visible = True
                Else
                    txtReadEU(chn).Visible = False
                End If
            Case 7      ' RTD 100 ohm
                descAddrChan(chn).ForeColor = lblInput.ForeColor
                lblAddrChan(chn).Visible = True
                descAddrChan(chn).Visible = True
                txtWritePerc(chn).Visible = False
                cmdSetOut(chn).Visible = False
                txtReadPerc(chn).Visible = True
                txtReadRaw(chn).Visible = True
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU(chn).Visible = True
                Else
                    txtReadEU(chn).Visible = False
                End If
            Case Else
                ' Nothing to Display
                txtReadEU(chn).Visible = False
                txtReadPerc(chn).Visible = False
                txtReadRaw(chn).Visible = False
                lblAddrChan(chn).Visible = False
                descAddrChan(chn).Visible = False
                txtWritePerc(chn).Visible = False
                cmdSetOut(chn).Visible = False
        End Select
        descAddrChan(chn).Caption = Left(OptoChanDesc(addr, chan), 24)
        lblAddrChan(chn).Caption = Format(addr, "#0") & "/" & Format(chan, "#0")
    Next chan

End Sub

