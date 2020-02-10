VERSION 5.00
Begin VB.Form frmPortValues 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Values"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frmPortValues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   0
      Top             =   7890
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   16
      Left            =   1680
      TabIndex        =   34
      Text            =   "- 01234.567"
      Top             =   7575
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   15
      Left            =   1680
      TabIndex        =   32
      Text            =   "- 01234.567"
      Top             =   7170
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   14
      Left            =   1680
      TabIndex        =   30
      Text            =   "- 01234.567"
      Top             =   6765
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   13
      Left            =   1680
      TabIndex        =   28
      Text            =   "- 01234.567"
      Top             =   6360
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   12
      Left            =   1680
      TabIndex        =   26
      Text            =   "- 01234.567"
      Top             =   5745
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   11
      Left            =   1680
      TabIndex        =   24
      Text            =   "- 01234.567"
      Top             =   5340
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   10
      Left            =   1680
      TabIndex        =   22
      Text            =   "- 01234.567"
      Top             =   4935
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   9
      Left            =   1680
      TabIndex        =   20
      Text            =   "- 01234.567"
      Top             =   4530
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   8
      Left            =   1680
      TabIndex        =   18
      Text            =   "- 01234.567"
      Top             =   3975
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   7
      Left            =   1680
      TabIndex        =   16
      Text            =   "- 01234.567"
      Top             =   3570
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   6
      Left            =   1680
      TabIndex        =   14
      Text            =   "- 01234.567"
      Top             =   3165
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   5
      Left            =   1680
      TabIndex        =   12
      Text            =   "- 01234.567"
      Top             =   2760
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   4
      Left            =   1680
      TabIndex        =   10
      Text            =   "- 01234.567"
      Top             =   2175
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Text            =   "- 01234.567"
      Top             =   1770
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Text            =   "- 01234.567"
      Top             =   1365
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   0
      Left            =   8595
      TabIndex        =   4
      Text            =   "- 01234.567"
      Top             =   120
      Width           =   1800
   End
   Begin VB.TextBox txtScaleValue 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   1
      Left            =   1680
      TabIndex        =   0
      Text            =   "- 01234.567"
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   3600
      TabIndex        =   53
      Top             =   7605
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   3600
      TabIndex        =   52
      Top             =   7200
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   3600
      TabIndex        =   51
      Top             =   6795
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   3600
      TabIndex        =   50
      Top             =   6390
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   3600
      TabIndex        =   49
      Top             =   5775
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   3600
      TabIndex        =   48
      Top             =   5370
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   3600
      TabIndex        =   47
      Top             =   4965
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   3600
      TabIndex        =   46
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   3600
      TabIndex        =   45
      Top             =   4005
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   3600
      TabIndex        =   44
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   3600
      TabIndex        =   43
      Top             =   3195
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   3600
      TabIndex        =   42
      Top             =   2790
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   3600
      TabIndex        =   41
      Top             =   2205
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   3600
      TabIndex        =   40
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   3600
      TabIndex        =   39
      Top             =   1395
      Width           =   1395
   End
   Begin VB.Label lblScaleType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   480
      Left            =   3600
      TabIndex        =   38
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3600
      TabIndex        =   37
      Top             =   990
      Width           =   1395
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10395
      TabIndex        =   36
      Top             =   150
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   240
      TabIndex        =   35
      Top             =   7605
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   240
      TabIndex        =   33
      Top             =   7200
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   240
      TabIndex        =   31
      Top             =   6795
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   240
      TabIndex        =   29
      Top             =   6390
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   240
      TabIndex        =   27
      Top             =   5775
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   240
      TabIndex        =   25
      Top             =   5370
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   240
      TabIndex        =   23
      Top             =   4965
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   240
      TabIndex        =   21
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   240
      TabIndex        =   19
      Top             =   4005
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   3195
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2790
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2205
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1395
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7200
      TabIndex        =   5
      Top             =   150
      Width           =   1395
   End
   Begin VB.Label lblPortNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   480
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "port #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   990
      Width           =   1395
   End
   Begin VB.Label lblScaleValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   480
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmPortValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 19200 ''''''''''''' Form PortValues '''''''''''''''''''
Option Explicit
'
'Global CommReadBuffer(0 To MAX_COMM) As String
'Global CommReadString(0 To MAX_COMM) As String
'Global Port_In_Use(0 To MAX_COMM) As Boolean
'Global Port_OK(0 To MAX_COMM) As Boolean
'Global Port_Type(0 To MAX_COMM) As String * 1               ' (port) Scale Type
'Global Port_Weight(0 To MAX_COMM) As Single                 ' (port) Scale Weight (string converted to number)
'Global Port_Value(0 To MAX_COMM) As String                  ' (port) Scale Input String
Private iPort As Integer

Public Sub UpdateScreen()
'
    For iPort = 1 To MAX_COMM
        txtScaleValue(iPort).text = Format(Port_Weight(iPort), "#,###,##0.0##")
        txtScaleValue(iPort).ForeColor = IIf((Port_In_Use(iPort)), BarActual_ForeColor, MEDGRAY)
        lblPort(iPort).ForeColor = IIf((Port_In_Use(iPort)), Black, MEDGRAY)
        lblType(iPort).ForeColor = IIf((Port_In_Use(iPort)), (IIf((Port_OK(iPort)), MEDGREEN, Alarm_ForeColor)), MEDGRAY)
    Next iPort
    
End Sub

Private Sub Form_Load()
'
Dim flag As Boolean

    lblScaleType.ForeColor = Titles_ForeColor
    lblPortNum.ForeColor = Titles_ForeColor
    lblScaleValue.ForeColor = Titles_ForeColor
    For iPort = 1 To txtScaleValue.UBound
        flag = IIf((iPort > MAX_COMM), False, True)
        lblPort(iPort).Visible = flag
        txtScaleValue(iPort).Visible = flag
        lblType(iPort).Visible = flag
        lblPort(iPort).Caption = Format(iPort, "#0")
        Select Case Port_Type(iPort)
            Case "A"                    'Acculab
                lblType(iPort).Caption = "Acculab"
            Case "S"                    'Sartorius
                lblType(iPort).Caption = "Sartorius"
            Case "T"                    'Toledo
                lblType(iPort).Caption = "Toledo"
            Case "N"                    'A & D
                lblType(iPort).Caption = "A & D"
            Case "V"                    'Toledo Viper
                lblType(iPort).Caption = "Viper"
            Case Else
                lblType(iPort).Caption = "undefined"
        End Select
    
    Next iPort
    
End Sub

Private Sub tmrUpdate_Timer()
    UpdateScreen
End Sub
