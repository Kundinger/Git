VERSION 5.00
Begin VB.Form frmCalHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calibration Help"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmCalHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblAcquireActual 
      Caption         =   $"frmCalHelp.frx":57E2
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
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   8535
   End
   Begin VB.Label lblAcquireRaw 
      Caption         =   $"frmCalHelp.frx":5891
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
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   8535
   End
   Begin VB.Label lblActual 
      Caption         =   "ACTUAL - The value from the calibration device."
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
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   8535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRaw 
      Caption         =   "RAW - The value from the device under calibration."
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
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   8535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "CALIBRATION HELP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmCalHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
