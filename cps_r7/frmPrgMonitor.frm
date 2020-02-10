VERSION 5.00
Begin VB.Form frmPrgMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PurgeAir Monitor"
   ClientHeight    =   10770
   ClientLeft      =   1650
   ClientTop       =   1650
   ClientWidth     =   10575
   Icon            =   "frmPrgMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13281.26
   ScaleMode       =   0  'User
   ScaleWidth      =   11803.65
   Begin VB.Frame frmAkXface 
      Caption         =   "AK Request/Ready"
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
      Height          =   735
      Left            =   10920
      TabIndex        =   167
      Top             =   3600
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Shape shpAkRdyIn 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   4650
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PAG Ready "
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
         Left            =   3450
         TabIndex        =   170
         Top             =   360
         Width           =   1200
      End
      Begin VB.Shape shpAkReqOut 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   1380
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblAkReqOut 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Request   Out "
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
         Left            =   120
         TabIndex        =   169
         Top             =   360
         Width           =   1260
      End
      Begin VB.Shape shpAkReqIn 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   2460
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblAkReqIn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " In "
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
         Left            =   2100
         TabIndex        =   168
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      DisabledPicture =   "frmPrgMonitor.frx":57E2
      DownPicture     =   "frmPrgMonitor.frx":6424
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
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrgMonitor.frx":7066
      Style           =   1  'Graphical
      TabIndex        =   162
      ToolTipText     =   "Quit"
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Frame frmPurgeWait 
      Caption         =   "Purge Waiting"
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
      Height          =   735
      Left            =   5880
      TabIndex        =   154
      Top             =   3600
      Width           =   3495
      Begin VB.Label lblSecToGoValue 
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
         Left            =   2880
         TabIndex        =   158
         Top             =   360
         Width           =   525
      End
      Begin VB.Label lblLstPrgStartTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "hh:mm:ss"
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
         Left            =   1080
         TabIndex        =   157
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lblLstPrgStart 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Start"
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
         Left            =   0
         TabIndex        =   156
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblSecToGo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Secs "
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
         Left            =   2280
         TabIndex        =   155
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame frmLocalPAS 
      Caption         =   "Local PAG Control"
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
      Height          =   6180
      Left            =   240
      TabIndex        =   25
      Top             =   4440
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdLoadControllers 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPrgMonitor.frx":7CA8
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Reload Controller Setup & Tuning Parameters"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   720
      End
      Begin VB.CommandButton cmdClrPasZlog 
         Height          =   315
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPrgMonitor.frx":7FEA
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Clear PAS zLog"
         Top             =   5760
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame frmPASControl 
         Caption         =   "PAG Control"
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
         Height          =   5565
         Left            =   5640
         TabIndex        =   81
         Top             =   130
         Width           =   4335
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   147
            Top             =   5280
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   146
            Top             =   5280
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   145
            Top             =   5280
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   144
            Top             =   5040
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   143
            Top             =   5040
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   142
            Top             =   5040
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   141
            Top             =   4800
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   140
            Top             =   4800
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   139
            Top             =   4800
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   138
            Top             =   4560
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   137
            Top             =   4560
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   136
            Top             =   4560
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   135
            Top             =   4320
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   134
            Top             =   4320
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   133
            Top             =   4320
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   132
            Top             =   4080
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   131
            Top             =   4080
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   130
            Top             =   4080
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   129
            Top             =   3840
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   128
            Top             =   3840
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   127
            Top             =   3840
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   126
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   125
            Top             =   3600
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   124
            Top             =   3600
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   123
            Top             =   3360
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   122
            Top             =   3360
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   121
            Top             =   3360
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   120
            Top             =   3120
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   119
            Top             =   3120
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   118
            Top             =   3120
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   117
            Top             =   2880
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   116
            Top             =   2880
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   115
            Top             =   2880
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   114
            Top             =   2640
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   113
            Top             =   2640
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   112
            Top             =   2640
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   111
            Top             =   2400
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   110
            Top             =   2400
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   109
            Top             =   2400
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   108
            Top             =   2160
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   107
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   106
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   105
            Top             =   1920
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   104
            Top             =   1920
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   103
            Top             =   1920
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   102
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   101
            Top             =   1680
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   100
            Top             =   1680
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   99
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   98
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   97
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   96
            Top             =   1200
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   95
            Top             =   1200
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   94
            Top             =   1200
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   93
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   92
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   91
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   90
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   89
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   88
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblItemDescr2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   120
            TabIndex        =   87
            Top             =   480
            Width           =   1500
         End
         Begin VB.Label lblDescr2 
            Alignment       =   2  'Center
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
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMoistureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1800
            TabIndex        =   85
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblMoisture2 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   1800
            TabIndex        =   84
            Top             =   240
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTemperatureValue2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2880
            TabIndex        =   83
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblTemperature2 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   2880
            TabIndex        =   82
            Top             =   240
            Width           =   1200
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frmPasCheck 
         Caption         =   "PAG Check"
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
         Height          =   2215
         Left            =   120
         TabIndex        =   56
         Top             =   3480
         Width           =   5415
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
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
            Left            =   240
            TabIndex        =   80
            Top             =   480
            Width           =   1500
         End
         Begin VB.Label lblDesr 
            Alignment       =   2  'Center
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
            Left            =   240
            TabIndex        =   79
            Top             =   240
            Width           =   1500
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1920
            TabIndex        =   78
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblMoisture 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   1920
            TabIndex        =   77
            Top             =   240
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   3000
            TabIndex        =   76
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label lblTemperature 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00004080&
            Height          =   255
            Left            =   3000
            TabIndex        =   75
            Top             =   240
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
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
            Left            =   240
            TabIndex        =   74
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Left            =   1920
            TabIndex        =   73
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Left            =   3000
            TabIndex        =   72
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "DurationTarget"
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
            Left            =   240
            TabIndex        =   71
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Left            =   1920
            TabIndex        =   70
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Left            =   3000
            TabIndex        =   69
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LastUpdate"
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
            Left            =   240
            TabIndex        =   68
            Top             =   1200
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Left            =   1920
            TabIndex        =   67
            Top             =   1200
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Left            =   3000
            TabIndex        =   66
            Top             =   1200
            Width           =   1200
         End
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TimeOut"
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
            Left            =   240
            TabIndex        =   65
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Left            =   1920
            TabIndex        =   64
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Left            =   3000
            TabIndex        =   63
            Top             =   1440
            Width           =   1200
         End
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TimeOut Duration"
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
            Left            =   240
            TabIndex        =   62
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Left            =   1920
            TabIndex        =   61
            Top             =   1680
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Left            =   3000
            TabIndex        =   60
            Top             =   1680
            Width           =   1200
         End
         Begin VB.Label lblItemDescr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TimeOut Target"
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
            Left            =   240
            TabIndex        =   59
            Top             =   1920
            Width           =   1500
         End
         Begin VB.Label lblMoistureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "6"
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
            Left            =   1920
            TabIndex        =   58
            Top             =   1920
            Width           =   1200
         End
         Begin VB.Label lblTemperatureValue 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "6"
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
            Left            =   3000
            TabIndex        =   57
            Top             =   1920
            Width           =   1200
         End
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
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "328.823"
         Top             =   1185
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "199.9"
         Top             =   1185
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
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "7654328"
         Top             =   1185
         Width           =   900
      End
      Begin VB.Frame frmCommonAI 
         Caption         =   "Common"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   5415
         Begin VB.Label lblActual 
            BackStyle       =   0  'Transparent
            Caption         =   "Actual"
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
            TabIndex        =   153
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblTempUnits2 
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
            Left            =   3825
            TabIndex        =   152
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblMoistUnits2 
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
            Left            =   3825
            TabIndex        =   151
            Top             =   1500
            Width           =   855
         End
         Begin VB.Label lblTempDispl2 
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
            Left            =   3120
            TabIndex        =   150
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lblMoistDispl2 
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
            Left            =   3120
            TabIndex        =   149
            Top             =   1500
            Width           =   645
         End
         Begin VB.Label lblTarget 
            BackStyle       =   0  'Transparent
            Caption         =   "Target"
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
            Left            =   3120
            TabIndex        =   148
            Top             =   240
            Width           =   855
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
            Left            =   900
            TabIndex        =   43
            Top             =   1500
            Width           =   645
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
            Left            =   900
            TabIndex        =   42
            Top             =   1245
            Width           =   645
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
            Left            =   900
            TabIndex        =   41
            Top             =   990
            Width           =   645
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
            Left            =   900
            TabIndex        =   40
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lblTemp 
            BackStyle       =   0  'Transparent
            Caption         =   "Temp."
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
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   855
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
            Left            =   120
            TabIndex        =   38
            Top             =   990
            Width           =   855
         End
         Begin VB.Label lblBarPress 
            BackStyle       =   0  'Transparent
            Caption         =   "Baro."
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
            Left            =   120
            TabIndex        =   37
            Top             =   1245
            Width           =   855
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
            Left            =   1605
            TabIndex        =   36
            Top             =   1500
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
            Left            =   1605
            TabIndex        =   35
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "% R. H."
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
            Left            =   1605
            TabIndex        =   34
            Top             =   735
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
            Left            =   900
            TabIndex        =   33
            Top             =   735
            Width           =   645
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
            Left            =   120
            TabIndex        =   32
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label3 
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
            Left            =   1605
            TabIndex        =   31
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label4 
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
            Left            =   1605
            TabIndex        =   30
            Top             =   1245
            Width           =   675
         End
         Begin VB.Label Label8 
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
            Left            =   120
            TabIndex        =   29
            Top             =   1500
            Width           =   855
         End
      End
      Begin VB.Label lblPasZlog 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Number of PAS zLog entries:"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   5790
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label lblPasZlogNumRecords 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "123456"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   2940
         TabIndex        =   164
         Top             =   5790
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Shape shpRunLocalPas 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   1920
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblRunLocalPas 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PAS Running   "
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
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   1680
      End
      Begin VB.Shape shpHeaterSSR 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   4560
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblHeaterSSR 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Heater SSR Out   "
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
         Left            =   2760
         TabIndex        =   54
         Top             =   600
         Width           =   1680
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
         Left            =   3300
         TabIndex        =   53
         Top             =   1230
         Width           =   2265
      End
      Begin VB.Label lblAddrChan2 
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
         Left            =   225
         TabIndex        =   52
         Top             =   1230
         Width           =   585
      End
      Begin VB.Label lblEu 
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
         Left            =   1155
         TabIndex        =   51
         Top             =   990
         Width           =   360
      End
      Begin VB.Label lblCounts 
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
         Left            =   2445
         TabIndex        =   50
         Top             =   990
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
         Left            =   1950
         TabIndex        =   49
         Top             =   990
         Width           =   255
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
         Left            =   3300
         TabIndex        =   48
         Top             =   990
         Width           =   2205
      End
      Begin VB.Label lblAddrChan 
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
         Left            =   120
         TabIndex        =   47
         Top             =   990
         Width           =   975
      End
      Begin VB.Shape shpPasReq 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   1920
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblPasReq 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Request In   "
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
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1680
      End
      Begin VB.Shape shpPasRdy 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   4560
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblPasRdy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ready Out   "
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
         Left            =   2760
         TabIndex        =   26
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.Frame frmHdwXface 
      Caption         =   "Hardware Request/Ready"
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
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label lblHdwRdy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ready In   "
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
         Left            =   2760
         TabIndex        =   24
         Top             =   360
         Width           =   1680
      End
      Begin VB.Shape shpHdwRdy 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   4560
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblHdwReq 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Request Out  "
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1680
      End
      Begin VB.Shape shpHdwReq 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Left            =   1920
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame frmPrgLogicals 
      Caption         =   "Logical Values"
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
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.CommandButton cmdClrZlog 
         Height          =   315
         Left            =   7920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPrgMonitor.frx":832C
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Clear Purge zLog"
         Top             =   2970
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblZlogNumRecords 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "123456"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   6180
         TabIndex        =   160
         Top             =   3000
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblZlog 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Purge zLog entries:"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   3360
         TabIndex        =   159
         Top             =   3000
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
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
         Left            =   240
         TabIndex        =   21
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   8760
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   7920
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   6360
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   7080
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   5640
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   4080
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   4800
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   3360
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   1800
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   2520
         Top             =   2640
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   9
         Left            =   1080
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
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
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   8760
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   7920
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   6360
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   7080
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   5640
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   4080
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   4800
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   3360
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   1800
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   2520
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   8
         Left            =   1080
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   8760
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   7920
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   6360
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   7080
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   5640
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   4080
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   4800
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   3360
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   1800
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   2520
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   7
         Left            =   1080
         Top             =   2160
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   8760
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   7920
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   6360
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   7080
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   5640
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   4080
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   4800
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   3360
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   1800
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   2520
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   6
         Left            =   1080
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   8760
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   7920
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   6360
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   7080
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   5640
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   4080
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   4800
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   3360
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   1800
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   2520
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   5
         Left            =   1080
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   8760
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   7920
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   6360
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   7080
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   5640
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   4080
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   4800
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   3360
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   1800
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   2520
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   4
         Left            =   1080
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   8760
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   7920
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   6360
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   7080
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   5640
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   4080
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   4800
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   3360
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   1800
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   2520
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   3
         Left            =   1080
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   8760
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   7920
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   6360
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   7080
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   5640
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   4080
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   4800
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   3360
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   1800
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   2520
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   2
         Left            =   1080
         Top             =   960
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
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
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   720
      End
      Begin VB.Shape shpPiab 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   8760
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblPiabSol 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Piab"
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
         Left            =   8760
         TabIndex        =   13
         Top             =   480
         Width           =   720
      End
      Begin VB.Shape shpRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   7920
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblRunning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
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
         Left            =   7920
         TabIndex        =   12
         Top             =   480
         Width           =   720
      End
      Begin VB.Shape shpReqd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   6360
         Top             =   720
         Width           =   720
      End
      Begin VB.Shape shpRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   7080
         Top             =   720
         Width           =   720
      End
      Begin VB.Shape shpStdby 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   5640
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblReady 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rdy"
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
         Left            =   7080
         TabIndex        =   11
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblStdby 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stdby"
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
         Left            =   5640
         TabIndex        =   10
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblRequested 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Req'd"
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
         Left            =   6360
         TabIndex        =   9
         Top             =   480
         Width           =   720
      End
      Begin VB.Shape shpLstRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   4080
         Top             =   720
         Width           =   720
      End
      Begin VB.Shape shpLstRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   4800
         Top             =   720
         Width           =   720
      End
      Begin VB.Shape shpLstStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   3360
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblLstRun 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LstRun"
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
         Left            =   4800
         TabIndex        =   8
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblLstStd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LstStd"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblLstRdy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LstRdy"
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
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   720
      End
      Begin VB.Shape shpReqRdy 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   1800
         Top             =   720
         Width           =   720
      End
      Begin VB.Shape shpReqRun 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   2520
         Top             =   720
         Width           =   720
      End
      Begin VB.Shape shpReqStd 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   255
         Index           =   1
         Left            =   1080
         Top             =   720
         Width           =   720
      End
      Begin VB.Label PrgAirNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblPrgAir 
         BackStyle       =   0  'Transparent
         Caption         =   "Source#"
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
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblReqRun 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ReqRun"
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
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblReqStd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ReqStd"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblReqRdy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ReqRdy"
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
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   0
      Top             =   3720
   End
End
Attribute VB_Name = "frmPrgMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' PURGE AIR MONITOR SCREEN
'
'
Option Explicit

Private Sub cmdClrZlog_Click()
    Debug_ZlogPurge_Clear = True
    Write_Zlog_Purge 1, 0, 0, 0, 0, "Purge zLog Cleared"
End Sub

Private Sub cmdClrPasZlog_Click()
    Debug_ZlogPAS_Clear = True
    Write_Zlog_PAS "PAS zLog Cleared"
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmPrgMonitor = Nothing
End Sub

Private Sub cmdLoadControllers_Click()
    ' Load Controllers Config Data
    Load_Controllers
End Sub

Private Sub Form_Load()

    frmPrgLogicals.ForeColor = Titles_ForeColor
    frmHdwXface.ForeColor = Titles_ForeColor
    frmAkXface.ForeColor = Titles_ForeColor
    frmPurgeWait.ForeColor = Titles_ForeColor
    frmLocalPAS.ForeColor = Titles_ForeColor
    frmPASControl.ForeColor = Titles_ForeColor
    frmCommonAI.ForeColor = Titles_ForeColor
    frmPasCheck.ForeColor = Titles_ForeColor
'    frmPrgMonitor.Height = IIf(USINGPASLOCALCONTROL, 11310, 5250)
    
    Select Case LocalPagControl.Type
        Case pagAlone
            ' stand-alone Master
            frmAkXface.Visible = False
            frmHdwXface.Visible = False
            frmPrgMonitor.Height = 11310
            frmLocalPAS.Visible = True
            cmdClrZlog.Visible = IIf(Not NotDebugPURGE, True, False)
            lblZlog.Visible = IIf(Not NotDebugPURGE, True, False)
            lblZlogNumRecords.Visible = IIf(Not NotDebugPURGE, True, False)
            cmdClrPasZlog.Visible = IIf(Not NotDebugPAS, True, False)
            lblPasZlog.Visible = IIf(Not NotDebugPAS, True, False)
            lblPasZlogNumRecords.Visible = IIf(Not NotDebugPAS, True, False)
            ' TEMPERATURE
            If USINGC Then
                lblTempUnits.Caption = "deg C"
            End If
            If USINGF Then
                lblTempUnits.Caption = "deg F"
            End If
            lblTempUnits2.Caption = lblTempUnits.Caption
            ' MOISTURE
            If USINGMoist_RH Then
                lblMoistUnits.Caption = "% rH"
            End If
            If USINGMoist_Grains Then
                lblMoistUnits.Caption = "grains/lb"
            End If
            lblMoistUnits2.Caption = lblMoistUnits.Caption
            
            lblItemDescr(0).Caption = "OK"
            lblItemDescr(1).Caption = "OK Duration"
            lblItemDescr(2).Caption = "OK Target"
            lblItemDescr(3).Caption = "TimeOut"
            lblItemDescr(4).Caption = "TimeOut Duration"
            lblItemDescr(5).Caption = "TimeOut Target"
            lblItemDescr(6).Caption = "Last Update"
            
            lblItemDescr2(0).Caption = "SP"
            lblItemDescr2(1).Caption = "PV"
            lblItemDescr2(2).Caption = "Error"
            lblItemDescr2(3).Caption = "Cum I"
            lblItemDescr2(4).Caption = "Output"
            lblItemDescr2(5).Caption = "P"
            lblItemDescr2(6).Caption = "I"
            lblItemDescr2(7).Caption = "D"
            lblItemDescr2(8).Caption = "Enable"
            lblItemDescr2(9).Caption = "Inhibit"
            lblItemDescr2(10).Caption = "Rev"
            lblItemDescr2(11).Caption = "Off Timer"
            lblItemDescr2(12).Caption = "OffDuty"
            lblItemDescr2(13).Caption = "OffDuty Mult"
            lblItemDescr2(14).Caption = "OffLimit Delta"
            lblItemDescr2(15).Caption = "On Timer"
            lblItemDescr2(16).Caption = "OnDuty"
            lblItemDescr2(17).Caption = "OnDutyMult"
            lblItemDescr2(18).Caption = "OnLimit Delta"
            lblItemDescr2(19).Caption = "Output"
            lblItemDescr2(20).Caption = "LastUpdate"
        
        Case pagMaster
            ' AK Master
            frmAkXface.Visible = False
            frmHdwXface.Visible = False
            frmPrgMonitor.Height = 11310
            frmLocalPAS.Visible = True
            cmdClrZlog.Visible = IIf(Not NotDebugPURGE, True, False)
            lblZlog.Visible = IIf(Not NotDebugPURGE, True, False)
            lblZlogNumRecords.Visible = IIf(Not NotDebugPURGE, True, False)
            cmdClrPasZlog.Visible = IIf(Not NotDebugPAS, True, False)
            lblPasZlog.Visible = IIf(Not NotDebugPAS, True, False)
            lblPasZlogNumRecords.Visible = IIf(Not NotDebugPAS, True, False)
            ' TEMPERATURE
            If USINGC Then
                lblTempUnits.Caption = "deg C"
            End If
            If USINGF Then
                lblTempUnits.Caption = "deg F"
            End If
            lblTempUnits2.Caption = lblTempUnits.Caption
            ' MOISTURE
            If USINGMoist_RH Then
                lblMoistUnits.Caption = "% rH"
            End If
            If USINGMoist_Grains Then
                lblMoistUnits.Caption = "grains/lb"
            End If
            lblMoistUnits2.Caption = lblMoistUnits.Caption
            
            lblItemDescr(0).Caption = "OK"
            lblItemDescr(1).Caption = "OK Duration"
            lblItemDescr(2).Caption = "OK Target"
            lblItemDescr(3).Caption = "TimeOut"
            lblItemDescr(4).Caption = "TimeOut Duration"
            lblItemDescr(5).Caption = "TimeOut Target"
            lblItemDescr(6).Caption = "Last Update"
            
            lblItemDescr2(0).Caption = "SP"
            lblItemDescr2(1).Caption = "PV"
            lblItemDescr2(2).Caption = "Error"
            lblItemDescr2(3).Caption = "Cum I"
            lblItemDescr2(4).Caption = "Output"
            lblItemDescr2(5).Caption = "P"
            lblItemDescr2(6).Caption = "I"
            lblItemDescr2(7).Caption = "D"
            lblItemDescr2(8).Caption = "Enable"
            lblItemDescr2(9).Caption = "Inhibit"
            lblItemDescr2(10).Caption = "Rev"
            lblItemDescr2(11).Caption = "Off Timer"
            lblItemDescr2(12).Caption = "OffDuty"
            lblItemDescr2(13).Caption = "OffDuty Mult"
            lblItemDescr2(14).Caption = "OffLimit Delta"
            lblItemDescr2(15).Caption = "On Timer"
            lblItemDescr2(16).Caption = "OnDuty"
            lblItemDescr2(17).Caption = "OnDutyMult"
            lblItemDescr2(18).Caption = "OnLimit Delta"
            lblItemDescr2(19).Caption = "Output"
            lblItemDescr2(20).Caption = "LastUpdate"
        
        Case pagClient
            ' AK Client
            frmAkXface.Visible = False
            frmHdwXface.Visible = False
            frmPrgMonitor.Height = 4950
            frmLocalPAS.Visible = True
            frmLocalPAS.Caption = " "
            lblAddrChan.Visible = False
            lblAddrChan2.Visible = False
            lblEu.Visible = False
            lblReadPerc.Visible = False
            lblCounts.Visible = False
            lblAddrChanDesc.Visible = False
            txtReadEU.Visible = False
            txtReadPerc.Visible = False
            txtReadRaw.Visible = False
            lblRunLocalPas.Visible = False
            shpRunLocalPas.Visible = False
            lblHeaterSSR.Visible = False
            shpHeaterSSR.Visible = False
            frmPasCheck.Visible = False
            frmPASControl.Visible = False
            frmPasCheck.Visible = False
            frmPasCheck.Visible = False
            cmdClrZlog.Visible = False
            lblZlog.Visible = False
            lblZlogNumRecords.Visible = False
            cmdClrPasZlog.Visible = False
            lblPasZlog.Visible = False
            lblPasZlogNumRecords.Visible = False
            ' TEMPERATURE
            If USINGC Then
                lblTempUnits.Caption = "deg C"
            End If
            If USINGF Then
                lblTempUnits.Caption = "deg F"
            End If
            lblTempUnits2.Caption = lblTempUnits.Caption
            ' MOISTURE
            If USINGMoist_RH Then
                lblMoistUnits.Caption = "% rH"
            End If
            If USINGMoist_Grains Then
                lblMoistUnits.Caption = "grains/lb"
            End If
            lblMoistUnits2.Caption = lblMoistUnits.Caption
                    
    End Select
End Sub

Private Sub tmrUpdate_Timer()
Dim prg, itm, addr, chan As Integer
Dim tempPerc, tempVdc As Single
Dim usingAkReqRdy As Boolean
Dim usingHdwReqRdy As Boolean

    usingAkReqRdy = False
    usingHdwReqRdy = False
    For prg = 1 To NR_PRGAIR
    
        If PRG_INFO(prg).UsingPrgReqAK Then usingAkReqRdy = True
        If PRG_INFO(prg).UsingPrgReqHdw Then usingHdwReqRdy = True
    
        shpReqStd(prg).BackColor = IIf(PRG_INFO(prg).StandbyRequest, MEDGREEN, DK3ORANGE)
        shpReqRdy(prg).BackColor = IIf(PRG_INFO(prg).RequestRdy, MEDGREEN, DK3ORANGE)
        shpReqRun(prg).BackColor = IIf(PRG_INFO(prg).RequestRun, MEDGREEN, DK3ORANGE)
    
        shpLstStd(prg).BackColor = IIf(PRG_INFO(prg).LastStandbyRequest, MEDGREEN, DK3ORANGE)
        shpLstRdy(prg).BackColor = IIf(PRG_INFO(prg).LastRequestRdy, MEDGREEN, DK3ORANGE)
        shpLstRun(prg).BackColor = IIf(PRG_INFO(prg).LastRequestRun, MEDGREEN, DK3ORANGE)
    
        shpStdby(prg).BackColor = IIf(PRG_INFO(prg).StandingBy, MEDGREEN, DK3ORANGE)
        shpReqd(prg).BackColor = IIf(PRG_INFO(prg).Requested, MEDGREEN, DK3ORANGE)
        shpRdy(prg).BackColor = IIf(PRG_INFO(prg).Ready, MEDGREEN, DK3ORANGE)
    
        shpRun(prg).BackColor = IIf(PRG_INFO(prg).Running, MEDGREEN, DK3ORANGE)
    
        shpPiab(prg).BackColor = IIf(Prg_DIO(prg, ipPiabSol).Value, MEDGREEN, DK3ORANGE)
        
    Next prg
    lblZlogNumRecords.Caption = CStr(Debug_ZlogPurge_NumRecords)
    lblLstPrgStartTime.Caption = Format(LastPurgeStart, "hh:mm:ss")
    lblSecToGoValue.Caption = IIf(DateDiff("s", Now, LastPurgeStart + TimeSerial(0, 5, 0)) > 0, _
                                Format(DateDiff("s", Now, LastPurgeStart + TimeSerial(0, 5, 0)), "####0"), 0)
    Select Case LocalPagControl.Type
        Case pagAlone
            ' stand-alone Master
            If usingHdwReqRdy Then
                frmHdwXface.Left = frmPrgLogicals.Left
                frmAkXface.Left = OutOfSight
                shpHdwReq.BackColor = IIf(Com_DIO(icPurgeRequestOut).Value, MEDGREEN, DK3ORANGE)
                shpHdwRdy.BackColor = IIf(Com_DIO(icPurgeReadyIn).Value, MEDGREEN, DK3ORANGE)
            End If
            ' Temperature
            lblTempDispl.Caption = Format(PATemp, "##0.00")
            ' Humidity
            lblHumidDispl.Caption = Format(PAHum, "##0.00")             ' Display Humindity in PerCent RH
            ' Pressure
            lblLkChkPresDispl.Caption = Format(PTinvalue, "###0.00")    ' Display Leak Press in psig
            ' Barometer
            lblBarPresDispl.Caption = Format(AmbBaro, "#000")           ' Display Baro in mBar
            ' Moisture
            lblMoistDispl.Caption = Format(PAMoisture, "##0.00")
            
            ' Temperature Target
            lblTempDispl2.Caption = Format(SysConfig.Temp_Target, "##0.00")
            ' Moisture Target
            lblMoistDispl2.Caption = Format(SysConfig.Moisture_Target, "##0.00")
            
            shpPasReq.BackColor = IIf((PAG_Request Or MasterPagData.ReqIn), MEDGREEN, DK3ORANGE)
            shpPasRdy.BackColor = IIf(Com_DIO(icPASReadyOut).Value, MEDGREEN, DK3ORANGE)
            shpRunLocalPas.BackColor = IIf(Com_DIO(icPASisRunningIn).Value, MEDGREEN, DK3ORANGE)
            shpHeaterSSR.BackColor = IIf(Com_DIO(icPASHeaterSSR).Value, MEDGREEN, DK3ORANGE)
            
            addr = Com_AIO(acPASMoistCntrlOut).addr
            chan = Com_AIO(acPASMoistCntrlOut).chan
            descAddrChan.Caption = Left(OptoChanDesc(addr, chan), 20)
            lblAddrChan2.Caption = Format(addr, "#0") & "/" & Format(chan, "#0")
            txtReadRaw = Format(OptoAIO(addr, chan).RawValue, "######0")
            If (Map_AIO(addr, chan).VdcMax > Map_AIO(addr, chan).VdcMin) Then
                tempVdc = 10# * (Map_AIO(addr, chan).RawValue / FULLSCALE)          ' Vdc out of 0-10Vdc
                tempVdc = tempVdc - Map_AIO(addr, chan).VdcMin                      ' Vdc above VdcMin
                tempPerc = tempVdc / (Map_AIO(addr, chan).VdcMax - Map_AIO(addr, chan).VdcMin)
                tempPerc = CSng(100# * tempPerc)
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU.Visible = True
                    txtReadEU.text = Format(Map_AIO(addr, chan).EUValue, "##0.000")
                Else
                    txtReadEU.Visible = False
                End If
            Else
                txtReadEU.Visible = False
                tempPerc = CSng(txtReadRaw) * (100# / FULLSCALE)
            End If
            txtReadPerc.text = Format(tempPerc, "##0.0")
            
            lblMoistureValue(0).Caption = IIf(PAS_INFO(pasMOISTURE).Ok, "TRUE", "FALSE")
            lblMoistureValue(0).ForeColor = IIf(PAS_INFO(pasMOISTURE).Ok, DK2GREEN, LTORANGE)
            lblMoistureValue(1).Caption = Format(PAS_INFO(pasMOISTURE).Duration, "#####0.000")
            lblMoistureValue(2).Caption = Format(PAS_INFO(pasMOISTURE).DurationTarget, "#####0.000")
            lblMoistureValue(3).Caption = IIf(PAS_INFO(pasMOISTURE).timeOut, "TRUE", "FALSE")
            lblMoistureValue(3).ForeColor = IIf(PAS_INFO(pasMOISTURE).timeOut, DK2GREEN, LTORANGE)
            lblMoistureValue(4).Caption = Format(PAS_INFO(pasMOISTURE).TimeOutDuration, "#####0.000")
            lblMoistureValue(5).Caption = Format(PAS_INFO(pasMOISTURE).TimeOutTarget, "#####0.000")
            lblMoistureValue(6).Caption = Format(PAS_INFO(pasMOISTURE).LastUpdate, "#####0.000")
            
            lblTemperatureValue(0).Caption = IIf(PAS_INFO(pasTemperature).Ok, "TRUE", "FALSE")
            lblTemperatureValue(0).ForeColor = IIf(PAS_INFO(pasTemperature).Ok, DK2GREEN, LTORANGE)
            lblTemperatureValue(1).Caption = Format(PAS_INFO(pasTemperature).Duration, "#####0.000")
            lblTemperatureValue(2).Caption = Format(PAS_INFO(pasTemperature).DurationTarget, "#####0.000")
            lblTemperatureValue(3).Caption = IIf(PAS_INFO(pasTemperature).timeOut, "TRUE", "FALSE")
            lblTemperatureValue(3).ForeColor = IIf(PAS_INFO(pasTemperature).timeOut, DK2GREEN, LTORANGE)
            lblTemperatureValue(4).Caption = Format(PAS_INFO(pasTemperature).TimeOutDuration, "#####0.000")
            lblTemperatureValue(5).Caption = Format(PAS_INFO(pasTemperature).TimeOutTarget, "#####0.000")
            lblTemperatureValue(6).Caption = Format(PAS_INFO(pasTemperature).LastUpdate, "#####0.000")
        
            lblMoistureValue2(0).Caption = Format(PID_INFO(pasMOISTURE).SP, "####0.00")
            lblMoistureValue2(1).Caption = Format(PID_INFO(pasMOISTURE).PV, "####0.00")
            lblMoistureValue2(2).Caption = Format(PID_INFO(pasMOISTURE).Er, "####0.00")
            lblMoistureValue2(3).Caption = Format(PID_INFO(pasMOISTURE).CumI, "####0.00")
            lblMoistureValue2(4).Caption = Format(PID_INFO(pasMOISTURE).out, "####0.00")
            lblMoistureValue2(5).Caption = Format(PID_INFO(pasMOISTURE).Pgain, "####0.00")
            lblMoistureValue2(6).Caption = Format(PID_INFO(pasMOISTURE).Igain, "####0.00")
            lblMoistureValue2(7).Caption = Format(PID_INFO(pasMOISTURE).Dgain, "####0.00")
            lblMoistureValue2(8).Caption = IIf(PID_INFO(pasMOISTURE).Enable, "TRUE", "FALSE")
            lblMoistureValue2(8).ForeColor = IIf(PID_INFO(pasMOISTURE).Enable, DK2GREEN, LTORANGE)
            lblMoistureValue2(9).Caption = IIf(PID_INFO(pasMOISTURE).Inhibit, "TRUE", "FALSE")
            lblMoistureValue2(9).ForeColor = IIf(PID_INFO(pasMOISTURE).Inhibit, DK2GREEN, LTORANGE)
            lblMoistureValue2(10).Caption = IIf(PID_INFO(pasMOISTURE).Rev, "TRUE", "FALSE")
            lblMoistureValue2(10).ForeColor = IIf(PID_INFO(pasMOISTURE).Rev, DK2GREEN, LTORANGE)
            lblMoistureValue2(11).Caption = Format(PID_INFO(pasMOISTURE).OffTimer, "#####0.000")
            lblMoistureValue2(12).Caption = Format(PID_INFO(pasMOISTURE).OffDuty, "#####0.000")
            lblMoistureValue2(13).Caption = Format(PID_INFO(pasMOISTURE).OffDutyMult, "#####0.000")
            lblMoistureValue2(14).Caption = Format(PID_INFO(pasMOISTURE).OffLimitDelta, "#####0.000")
            lblMoistureValue2(15).Caption = Format(PID_INFO(pasMOISTURE).OnTimer, "#####0.000")
            lblMoistureValue2(16).Caption = Format(PID_INFO(pasMOISTURE).OnDuty, "#####0.000")
            lblMoistureValue2(17).Caption = Format(PID_INFO(pasMOISTURE).OnDutyMult, "#####0.000")
            lblMoistureValue2(18).Caption = Format(PID_INFO(pasMOISTURE).OnLimitDelta, "#####0.000")
            lblMoistureValue2(19).Caption = IIf(PID_INFO(pasMOISTURE).Output, "TRUE", "FALSE")
            lblMoistureValue2(19).ForeColor = IIf(PID_INFO(pasMOISTURE).Output, DK2GREEN, LTORANGE)
            lblMoistureValue2(20).Caption = Format(PID_INFO(pasMOISTURE).LastUpdate, "#####0.000")
        
            lblTemperatureValue2(0).Caption = Format(PID_INFO(pasTemperature).SP, "####0.00")
            lblTemperatureValue2(1).Caption = Format(PID_INFO(pasTemperature).PV, "####0.00")
            lblTemperatureValue2(2).Caption = Format(PID_INFO(pasTemperature).Er, "####0.00")
            lblTemperatureValue2(3).Caption = Format(PID_INFO(pasTemperature).CumI, "####0.00")
            lblTemperatureValue2(4).Caption = Format(PID_INFO(pasTemperature).out, "####0.00")
            lblTemperatureValue2(5).Caption = Format(PID_INFO(pasTemperature).Pgain, "####0.00")
            lblTemperatureValue2(6).Caption = Format(PID_INFO(pasTemperature).Igain, "####0.00")
            lblTemperatureValue2(7).Caption = Format(PID_INFO(pasTemperature).Dgain, "####0.00")
            lblTemperatureValue2(8).Caption = IIf(PID_INFO(pasTemperature).Enable, "TRUE", "FALSE")
            lblTemperatureValue2(8).ForeColor = IIf(PID_INFO(pasTemperature).Enable, DK2GREEN, LTORANGE)
            lblTemperatureValue2(9).Caption = IIf(PID_INFO(pasTemperature).Inhibit, "TRUE", "FALSE")
            lblTemperatureValue2(9).ForeColor = IIf(PID_INFO(pasTemperature).Inhibit, DK2GREEN, LTORANGE)
            lblTemperatureValue2(10).Caption = IIf(PID_INFO(pasTemperature).Rev, "TRUE", "FALSE")
            lblTemperatureValue2(10).ForeColor = IIf(PID_INFO(pasTemperature).Rev, DK2GREEN, LTORANGE)
            lblTemperatureValue2(11).Caption = Format(PID_INFO(pasTemperature).OffTimer, "#####0.000")
            lblTemperatureValue2(12).Caption = Format(PID_INFO(pasTemperature).OffDuty, "#####0.000")
            lblTemperatureValue2(13).Caption = Format(PID_INFO(pasTemperature).OffDutyMult, "#####0.000")
            lblTemperatureValue2(14).Caption = Format(PID_INFO(pasTemperature).OffLimitDelta, "#####0.000")
            lblTemperatureValue2(15).Caption = Format(PID_INFO(pasTemperature).OnTimer, "#####0.000")
            lblTemperatureValue2(16).Caption = Format(PID_INFO(pasTemperature).OnDuty, "#####0.000")
            lblTemperatureValue2(17).Caption = Format(PID_INFO(pasTemperature).OnDutyMult, "#####0.000")
            lblTemperatureValue2(18).Caption = Format(PID_INFO(pasTemperature).OnLimitDelta, "#####0.000")
            lblTemperatureValue2(19).Caption = IIf(PID_INFO(pasTemperature).Output, "TRUE", "FALSE")
            lblTemperatureValue2(19).ForeColor = IIf(PID_INFO(pasTemperature).Output, DK2GREEN, LTORANGE)
            lblTemperatureValue2(20).Caption = Format(PID_INFO(pasTemperature).LastUpdate, "#####0.000")
        
            lblPasZlogNumRecords.Caption = CStr(Debug_ZlogPAS_NumRecords)
    
        Case pagMaster
            ' AK Master
            If usingAkReqRdy Then
                frmHdwXface.Left = OutOfSight
                frmAkXface.Left = frmPrgLogicals.Left
                shpAkReqOut.BackColor = IIf(Com_DIO(icPurgeRequestOut).Value, MEDGREEN, DK3ORANGE)
                shpAkReqIn.BackColor = IIf((Com_DIO(icPurgeRequestOut).Value Or MasterPagData.ReqIn), MEDGREEN, DK3ORANGE)
            End If
            ' Temperature
            lblTempDispl.Caption = Format(PATemp, "##0.00")
            ' Humidity
            lblHumidDispl.Caption = Format(PAHum, "##0.00")             ' Display Humindity in PerCent RH
            ' Pressure
            lblLkChkPresDispl.Caption = Format(PTinvalue, "###0.00")    ' Display Leak Press in psig
            ' Barometer
            lblBarPresDispl.Caption = Format(AmbBaro, "#000")           ' Display Baro in mBar
            ' Moisture
            lblMoistDispl.Caption = Format(PAMoisture, "##0.00")
            
            ' Temperature Target
            lblTempDispl2.Caption = Format(SysConfig.Temp_Target, "##0.00")
            ' Moisture Target
            lblMoistDispl2.Caption = Format(SysConfig.Moisture_Target, "##0.00")
            
            shpPasReq.BackColor = IIf((Com_DIO(icPurgeRequestOut).Value Or MasterPagData.ReqIn), MEDGREEN, DK3ORANGE)
            shpPasRdy.BackColor = IIf(Com_DIO(icPASReadyOut).Value, MEDGREEN, DK3ORANGE)
            shpRunLocalPas.BackColor = IIf(Com_DIO(icPASisRunningIn).Value, MEDGREEN, DK3ORANGE)
            shpHeaterSSR.BackColor = IIf(Com_DIO(icPASHeaterSSR).Value, MEDGREEN, DK3ORANGE)
            
            addr = Com_AIO(acPASMoistCntrlOut).addr
            chan = Com_AIO(acPASMoistCntrlOut).chan
            descAddrChan.Caption = Left(OptoChanDesc(addr, chan), 20)
            lblAddrChan2.Caption = Format(addr, "#0") & "/" & Format(chan, "#0")
            txtReadRaw = Format(OptoAIO(addr, chan).RawValue, "######0")
            If (Map_AIO(addr, chan).VdcMax > Map_AIO(addr, chan).VdcMin) Then
                tempVdc = 10# * (Map_AIO(addr, chan).RawValue / FULLSCALE)          ' Vdc out of 0-10Vdc
                tempVdc = tempVdc - Map_AIO(addr, chan).VdcMin                      ' Vdc above VdcMin
                tempPerc = tempVdc / (Map_AIO(addr, chan).VdcMax - Map_AIO(addr, chan).VdcMin)
                tempPerc = CSng(100# * tempPerc)
                If (Map_AIO(addr, chan).EuMax > Map_AIO(addr, chan).EuMin) Then
                    txtReadEU.Visible = True
                    txtReadEU.text = Format(Map_AIO(addr, chan).EUValue, "##0.000")
                Else
                    txtReadEU.Visible = False
                End If
            Else
                txtReadEU.Visible = False
                tempPerc = CSng(txtReadRaw) * (100# / FULLSCALE)
            End If
            txtReadPerc.text = Format(tempPerc, "##0.0")
            
            lblMoistureValue(0).Caption = IIf(PAS_INFO(pasMOISTURE).Ok, "TRUE", "FALSE")
            lblMoistureValue(0).ForeColor = IIf(PAS_INFO(pasMOISTURE).Ok, DK2GREEN, LTORANGE)
            lblMoistureValue(1).Caption = Format(PAS_INFO(pasMOISTURE).Duration, "#####0.000")
            lblMoistureValue(2).Caption = Format(PAS_INFO(pasMOISTURE).DurationTarget, "#####0.000")
            lblMoistureValue(3).Caption = IIf(PAS_INFO(pasMOISTURE).timeOut, "TRUE", "FALSE")
            lblMoistureValue(3).ForeColor = IIf(PAS_INFO(pasMOISTURE).timeOut, DK2GREEN, LTORANGE)
            lblMoistureValue(4).Caption = Format(PAS_INFO(pasMOISTURE).TimeOutDuration, "#####0.000")
            lblMoistureValue(5).Caption = Format(PAS_INFO(pasMOISTURE).TimeOutTarget, "#####0.000")
            lblMoistureValue(6).Caption = Format(PAS_INFO(pasMOISTURE).LastUpdate, "#####0.000")
            
            lblTemperatureValue(0).Caption = IIf(PAS_INFO(pasTemperature).Ok, "TRUE", "FALSE")
            lblTemperatureValue(0).ForeColor = IIf(PAS_INFO(pasTemperature).Ok, DK2GREEN, LTORANGE)
            lblTemperatureValue(1).Caption = Format(PAS_INFO(pasTemperature).Duration, "#####0.000")
            lblTemperatureValue(2).Caption = Format(PAS_INFO(pasTemperature).DurationTarget, "#####0.000")
            lblTemperatureValue(3).Caption = IIf(PAS_INFO(pasTemperature).timeOut, "TRUE", "FALSE")
            lblTemperatureValue(3).ForeColor = IIf(PAS_INFO(pasTemperature).timeOut, DK2GREEN, LTORANGE)
            lblTemperatureValue(4).Caption = Format(PAS_INFO(pasTemperature).TimeOutDuration, "#####0.000")
            lblTemperatureValue(5).Caption = Format(PAS_INFO(pasTemperature).TimeOutTarget, "#####0.000")
            lblTemperatureValue(6).Caption = Format(PAS_INFO(pasTemperature).LastUpdate, "#####0.000")
        
            lblMoistureValue2(0).Caption = Format(PID_INFO(pasMOISTURE).SP, "####0.00")
            lblMoistureValue2(1).Caption = Format(PID_INFO(pasMOISTURE).PV, "####0.00")
            lblMoistureValue2(2).Caption = Format(PID_INFO(pasMOISTURE).Er, "####0.00")
            lblMoistureValue2(3).Caption = Format(PID_INFO(pasMOISTURE).CumI, "####0.00")
            lblMoistureValue2(4).Caption = Format(PID_INFO(pasMOISTURE).out, "####0.00")
            lblMoistureValue2(5).Caption = Format(PID_INFO(pasMOISTURE).Pgain, "####0.00")
            lblMoistureValue2(6).Caption = Format(PID_INFO(pasMOISTURE).Igain, "####0.00")
            lblMoistureValue2(7).Caption = Format(PID_INFO(pasMOISTURE).Dgain, "####0.00")
            lblMoistureValue2(8).Caption = IIf(PID_INFO(pasMOISTURE).Enable, "TRUE", "FALSE")
            lblMoistureValue2(8).ForeColor = IIf(PID_INFO(pasMOISTURE).Enable, DK2GREEN, LTORANGE)
            lblMoistureValue2(9).Caption = IIf(PID_INFO(pasMOISTURE).Inhibit, "TRUE", "FALSE")
            lblMoistureValue2(9).ForeColor = IIf(PID_INFO(pasMOISTURE).Inhibit, DK2GREEN, LTORANGE)
            lblMoistureValue2(10).Caption = IIf(PID_INFO(pasMOISTURE).Rev, "TRUE", "FALSE")
            lblMoistureValue2(10).ForeColor = IIf(PID_INFO(pasMOISTURE).Rev, DK2GREEN, LTORANGE)
            lblMoistureValue2(11).Caption = Format(PID_INFO(pasMOISTURE).OffTimer, "#####0.000")
            lblMoistureValue2(12).Caption = Format(PID_INFO(pasMOISTURE).OffDuty, "#####0.000")
            lblMoistureValue2(13).Caption = Format(PID_INFO(pasMOISTURE).OffDutyMult, "#####0.000")
            lblMoistureValue2(14).Caption = Format(PID_INFO(pasMOISTURE).OffLimitDelta, "#####0.000")
            lblMoistureValue2(15).Caption = Format(PID_INFO(pasMOISTURE).OnTimer, "#####0.000")
            lblMoistureValue2(16).Caption = Format(PID_INFO(pasMOISTURE).OnDuty, "#####0.000")
            lblMoistureValue2(17).Caption = Format(PID_INFO(pasMOISTURE).OnDutyMult, "#####0.000")
            lblMoistureValue2(18).Caption = Format(PID_INFO(pasMOISTURE).OnLimitDelta, "#####0.000")
            lblMoistureValue2(19).Caption = IIf(PID_INFO(pasMOISTURE).Output, "TRUE", "FALSE")
            lblMoistureValue2(19).ForeColor = IIf(PID_INFO(pasMOISTURE).Output, DK2GREEN, LTORANGE)
            lblMoistureValue2(20).Caption = Format(PID_INFO(pasMOISTURE).LastUpdate, "#####0.000")
        
            lblTemperatureValue2(0).Caption = Format(PID_INFO(pasTemperature).SP, "####0.00")
            lblTemperatureValue2(1).Caption = Format(PID_INFO(pasTemperature).PV, "####0.00")
            lblTemperatureValue2(2).Caption = Format(PID_INFO(pasTemperature).Er, "####0.00")
            lblTemperatureValue2(3).Caption = Format(PID_INFO(pasTemperature).CumI, "####0.00")
            lblTemperatureValue2(4).Caption = Format(PID_INFO(pasTemperature).out, "####0.00")
            lblTemperatureValue2(5).Caption = Format(PID_INFO(pasTemperature).Pgain, "####0.00")
            lblTemperatureValue2(6).Caption = Format(PID_INFO(pasTemperature).Igain, "####0.00")
            lblTemperatureValue2(7).Caption = Format(PID_INFO(pasTemperature).Dgain, "####0.00")
            lblTemperatureValue2(8).Caption = IIf(PID_INFO(pasTemperature).Enable, "TRUE", "FALSE")
            lblTemperatureValue2(8).ForeColor = IIf(PID_INFO(pasTemperature).Enable, DK2GREEN, LTORANGE)
            lblTemperatureValue2(9).Caption = IIf(PID_INFO(pasTemperature).Inhibit, "TRUE", "FALSE")
            lblTemperatureValue2(9).ForeColor = IIf(PID_INFO(pasTemperature).Inhibit, DK2GREEN, LTORANGE)
            lblTemperatureValue2(10).Caption = IIf(PID_INFO(pasTemperature).Rev, "TRUE", "FALSE")
            lblTemperatureValue2(10).ForeColor = IIf(PID_INFO(pasTemperature).Rev, DK2GREEN, LTORANGE)
            lblTemperatureValue2(11).Caption = Format(PID_INFO(pasTemperature).OffTimer, "#####0.000")
            lblTemperatureValue2(12).Caption = Format(PID_INFO(pasTemperature).OffDuty, "#####0.000")
            lblTemperatureValue2(13).Caption = Format(PID_INFO(pasTemperature).OffDutyMult, "#####0.000")
            lblTemperatureValue2(14).Caption = Format(PID_INFO(pasTemperature).OffLimitDelta, "#####0.000")
            lblTemperatureValue2(15).Caption = Format(PID_INFO(pasTemperature).OnTimer, "#####0.000")
            lblTemperatureValue2(16).Caption = Format(PID_INFO(pasTemperature).OnDuty, "#####0.000")
            lblTemperatureValue2(17).Caption = Format(PID_INFO(pasTemperature).OnDutyMult, "#####0.000")
            lblTemperatureValue2(18).Caption = Format(PID_INFO(pasTemperature).OnLimitDelta, "#####0.000")
            lblTemperatureValue2(19).Caption = IIf(PID_INFO(pasTemperature).Output, "TRUE", "FALSE")
            lblTemperatureValue2(19).ForeColor = IIf(PID_INFO(pasTemperature).Output, DK2GREEN, LTORANGE)
            lblTemperatureValue2(20).Caption = Format(PID_INFO(pasTemperature).LastUpdate, "#####0.000")
        
            lblPasZlogNumRecords.Caption = CStr(Debug_ZlogPAS_NumRecords)
    
        Case pagClient
            ' AK Client
            shpHdwReq.BackColor = IIf(Com_DIO(icPurgeRequestOut).Value, MEDGREEN, DK3ORANGE)
            shpHdwRdy.BackColor = IIf(Com_DIO(icPurgeReadyIn).Value, MEDGREEN, DK3ORANGE)
            ' Temperature
            lblTempDispl.Caption = Format(MasterPagData.Temperature, "##0.00")
            ' Humidity
            lblHumidDispl.Caption = Format(MasterPagData.Humidity, "##0.00")
            ' Moisture
            lblMoistDispl.Caption = Format(MasterPagData.Moisture, "##0.00")
            
            ' Temperature Target
            lblTempDispl2.Caption = Format(MasterPagData.TempSP, "##0.00")
            ' Moisture Target
            lblMoistDispl2.Caption = Format(MasterPagData.MoistSP, "##0.00")
            
            shpPasReq.BackColor = IIf((PAG_Request Or MasterPagData.ReqIn), MEDGREEN, DK3ORANGE)
            shpPasRdy.BackColor = IIf(MasterPagData.RdyOut, MEDGREEN, DK3ORANGE)
            
            
        
        
    End Select
    
End Sub

