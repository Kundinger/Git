VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmComm8Card 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scale Monitor"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14175
   ControlBox      =   0   'False
   Icon            =   "Comm8Card.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPortValues 
      Caption         =   "Port Values"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Comm8Card.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Show Readings By Port"
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.TextBox txtDebug 
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
      Left            =   0
      TabIndex        =   80
      Text            =   " @+ !123456123456 <> + <> 0000111100001111 <> 7 <> 123.456"
      Top             =   0
      Width           =   7600
   End
   Begin VB.CommandButton cmdClrZlog 
      Height          =   315
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Comm8Card.frx":6424
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Clear Scales zLog"
      Top             =   8805
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      DisabledPicture =   "Comm8Card.frx":6766
      DownPicture     =   "Comm8Card.frx":73A8
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Comm8Card.frx":7FEA
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      DisabledPicture =   "Comm8Card.frx":8C2C
      DownPicture     =   "Comm8Card.frx":986E
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Comm8Card.frx":A4B0
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Save Scale Setup Values"
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   16
      ItemData        =   "Comm8Card.frx":B0F2
      Left            =   10560
      List            =   "Comm8Card.frx":B130
      Style           =   2  'Dropdown List
      TabIndex        =   74
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   6314
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   15
      ItemData        =   "Comm8Card.frx":B1D3
      Left            =   10560
      List            =   "Comm8Card.frx":B211
      Style           =   2  'Dropdown List
      TabIndex        =   73
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   5579
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   14
      ItemData        =   "Comm8Card.frx":B2B4
      Left            =   10560
      List            =   "Comm8Card.frx":B2F2
      Style           =   2  'Dropdown List
      TabIndex        =   72
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   4844
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   13
      ItemData        =   "Comm8Card.frx":B395
      Left            =   10560
      List            =   "Comm8Card.frx":B3D3
      Style           =   2  'Dropdown List
      TabIndex        =   71
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   4109
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   12
      ItemData        =   "Comm8Card.frx":B476
      Left            =   10560
      List            =   "Comm8Card.frx":B4B4
      Style           =   2  'Dropdown List
      TabIndex        =   70
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   3374
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   11
      ItemData        =   "Comm8Card.frx":B557
      Left            =   10560
      List            =   "Comm8Card.frx":B595
      Style           =   2  'Dropdown List
      TabIndex        =   69
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   2639
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   10
      ItemData        =   "Comm8Card.frx":B638
      Left            =   10560
      List            =   "Comm8Card.frx":B676
      Style           =   2  'Dropdown List
      TabIndex        =   68
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   1904
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   9
      ItemData        =   "Comm8Card.frx":B719
      Left            =   10560
      List            =   "Comm8Card.frx":B757
      Style           =   2  'Dropdown List
      TabIndex        =   67
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   1169
      Width           =   1400
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   8
      ItemData        =   "Comm8Card.frx":B7FA
      Left            =   3120
      List            =   "Comm8Card.frx":B838
      Style           =   2  'Dropdown List
      TabIndex        =   66
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   6314
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   7
      ItemData        =   "Comm8Card.frx":B8DB
      Left            =   3120
      List            =   "Comm8Card.frx":B919
      Style           =   2  'Dropdown List
      TabIndex        =   65
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   5579
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   6
      ItemData        =   "Comm8Card.frx":B9BC
      Left            =   3120
      List            =   "Comm8Card.frx":B9FA
      Style           =   2  'Dropdown List
      TabIndex        =   64
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   4844
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   5
      ItemData        =   "Comm8Card.frx":BA9D
      Left            =   3120
      List            =   "Comm8Card.frx":BADB
      Style           =   2  'Dropdown List
      TabIndex        =   63
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   4109
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   4
      ItemData        =   "Comm8Card.frx":BB7E
      Left            =   3120
      List            =   "Comm8Card.frx":BBBC
      Style           =   2  'Dropdown List
      TabIndex        =   62
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   3374
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   3
      ItemData        =   "Comm8Card.frx":BC5F
      Left            =   3120
      List            =   "Comm8Card.frx":BC9D
      Style           =   2  'Dropdown List
      TabIndex        =   61
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   2639
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   2
      ItemData        =   "Comm8Card.frx":BD40
      Left            =   3120
      List            =   "Comm8Card.frx":BD7E
      Style           =   2  'Dropdown List
      TabIndex        =   60
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   1904
      Width           =   1485
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   1
      ItemData        =   "Comm8Card.frx":BE21
      Left            =   3120
      List            =   "Comm8Card.frx":BE5F
      Style           =   2  'Dropdown List
      TabIndex        =   59
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   1169
      Width           =   1485
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   16
      ItemData        =   "Comm8Card.frx":BF02
      Left            =   8880
      List            =   "Comm8Card.frx":BF18
      Style           =   2  'Dropdown List
      TabIndex        =   58
      ToolTipText     =   "Select Scale Type"
      Top             =   6314
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   15
      ItemData        =   "Comm8Card.frx":BF50
      Left            =   8880
      List            =   "Comm8Card.frx":BF66
      Style           =   2  'Dropdown List
      TabIndex        =   57
      ToolTipText     =   "Select Scale Type"
      Top             =   5579
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   14
      ItemData        =   "Comm8Card.frx":BF9E
      Left            =   8880
      List            =   "Comm8Card.frx":BFB4
      Style           =   2  'Dropdown List
      TabIndex        =   56
      ToolTipText     =   "Select Scale Type"
      Top             =   4844
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   13
      ItemData        =   "Comm8Card.frx":BFEC
      Left            =   8880
      List            =   "Comm8Card.frx":C002
      Style           =   2  'Dropdown List
      TabIndex        =   55
      ToolTipText     =   "Select Scale Type"
      Top             =   4109
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   12
      ItemData        =   "Comm8Card.frx":C03A
      Left            =   8880
      List            =   "Comm8Card.frx":C050
      Style           =   2  'Dropdown List
      TabIndex        =   54
      ToolTipText     =   "Select Scale Type"
      Top             =   3374
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   11
      ItemData        =   "Comm8Card.frx":C088
      Left            =   8880
      List            =   "Comm8Card.frx":C09E
      Style           =   2  'Dropdown List
      TabIndex        =   53
      ToolTipText     =   "Select Scale Type"
      Top             =   2639
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   10
      ItemData        =   "Comm8Card.frx":C0D6
      Left            =   8880
      List            =   "Comm8Card.frx":C0EC
      Style           =   2  'Dropdown List
      TabIndex        =   52
      ToolTipText     =   "Select Scale Type"
      Top             =   1904
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   9
      ItemData        =   "Comm8Card.frx":C124
      Left            =   8880
      List            =   "Comm8Card.frx":C13A
      Style           =   2  'Dropdown List
      TabIndex        =   51
      ToolTipText     =   "Select Scale Type"
      Top             =   1169
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   8
      ItemData        =   "Comm8Card.frx":C172
      Left            =   1440
      List            =   "Comm8Card.frx":C188
      Style           =   2  'Dropdown List
      TabIndex        =   50
      ToolTipText     =   "Select Scale Type"
      Top             =   6314
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   7
      ItemData        =   "Comm8Card.frx":C1C0
      Left            =   1440
      List            =   "Comm8Card.frx":C1D6
      Style           =   2  'Dropdown List
      TabIndex        =   49
      ToolTipText     =   "Select Scale Type"
      Top             =   5579
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   6
      ItemData        =   "Comm8Card.frx":C20E
      Left            =   1440
      List            =   "Comm8Card.frx":C224
      Style           =   2  'Dropdown List
      TabIndex        =   48
      ToolTipText     =   "Select Scale Type"
      Top             =   4844
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   5
      ItemData        =   "Comm8Card.frx":C25C
      Left            =   1440
      List            =   "Comm8Card.frx":C272
      Style           =   2  'Dropdown List
      TabIndex        =   47
      ToolTipText     =   "Select Scale Type"
      Top             =   4109
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   4
      ItemData        =   "Comm8Card.frx":C2AA
      Left            =   1440
      List            =   "Comm8Card.frx":C2C0
      Style           =   2  'Dropdown List
      TabIndex        =   46
      ToolTipText     =   "Select Scale Type"
      Top             =   3374
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   3
      ItemData        =   "Comm8Card.frx":C2F8
      Left            =   1440
      List            =   "Comm8Card.frx":C313
      Style           =   2  'Dropdown List
      TabIndex        =   45
      ToolTipText     =   "Select Scale Type"
      Top             =   2639
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   2
      ItemData        =   "Comm8Card.frx":C34B
      Left            =   1440
      List            =   "Comm8Card.frx":C361
      Style           =   2  'Dropdown List
      TabIndex        =   44
      ToolTipText     =   "Select Scale Type"
      Top             =   1904
      Width           =   1605
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   360
      Index           =   1
      ItemData        =   "Comm8Card.frx":C399
      Left            =   1440
      List            =   "Comm8Card.frx":C3AF
      Style           =   2  'Dropdown List
      TabIndex        =   43
      ToolTipText     =   "Select Scale Type"
      Top             =   1169
      Width           =   1605
   End
   Begin VB.ComboBox CommPort 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Index           =   0
      ItemData        =   "Comm8Card.frx":C3E7
      Left            =   7920
      List            =   "Comm8Card.frx":C425
      Style           =   2  'Dropdown List
      TabIndex        =   42
      ToolTipText     =   "Alphanumeric Entry"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ComboBox ScaleType 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Index           =   0
      ItemData        =   "Comm8Card.frx":C4C8
      Left            =   6240
      List            =   "Comm8Card.frx":C4DB
      Style           =   2  'Dropdown List
      TabIndex        =   41
      ToolTipText     =   "Select Scale Type"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1605
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
      Left            =   4680
      TabIndex        =   13
      Text            =   "- 01234.567"
      Top             =   1882
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
      Left            =   4680
      TabIndex        =   10
      Text            =   "- 01234.567"
      Top             =   6292
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
      Left            =   4680
      TabIndex        =   9
      Text            =   "- 01234.567"
      Top             =   5557
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
      Left            =   4680
      TabIndex        =   8
      Text            =   "- 01234.567"
      Top             =   4822
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
      Left            =   4680
      TabIndex        =   7
      Text            =   "- 01234.567"
      Top             =   4087
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
      Left            =   4680
      TabIndex        =   6
      Text            =   "- 01234.567"
      Top             =   3352
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
      Left            =   4680
      TabIndex        =   5
      Text            =   "- 01234.567"
      Top             =   2617
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
      Left            =   4680
      TabIndex        =   4
      Text            =   "- 01234.567"
      Top             =   1147
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
      Index           =   16
      Left            =   12120
      TabIndex        =   20
      Text            =   "- 01234.567"
      Top             =   6292
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
      Left            =   12120
      TabIndex        =   19
      Text            =   "- 01234.567"
      Top             =   5557
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
      Left            =   12120
      TabIndex        =   18
      Text            =   "- 01234.567"
      Top             =   4822
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
      Left            =   12120
      TabIndex        =   17
      Text            =   "- 01234.567"
      Top             =   4087
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
      Left            =   12120
      TabIndex        =   16
      Text            =   "- 01234.567"
      Top             =   3352
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
      Left            =   12120
      TabIndex        =   15
      Text            =   "- 01234.567"
      Top             =   2617
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
      Left            =   12120
      TabIndex        =   14
      Text            =   "- 01234.567"
      Top             =   1882
      Width           =   1800
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   1
      Left            =   1440
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   720
      Top             =   7080
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
      Left            =   12120
      TabIndex        =   11
      Text            =   "- 01234.567"
      Top             =   1147
      Width           =   1800
   End
   Begin VB.TextBox text 
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
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Comm8Card.frx":C50C
      Top             =   7680
      Width           =   6240
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   2
      Left            =   2040
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   3
      Left            =   2640
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   4
      Left            =   3240
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   5
      Left            =   3840
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   6
      Left            =   4440
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   7
      Left            =   5040
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   8
      Left            =   5640
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   8
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   9
      Left            =   8880
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   9
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   10
      Left            =   9480
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   10
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   11
      Left            =   10080
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   11
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   12
      Left            =   10680
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   12
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   13
      Left            =   11280
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   13
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   14
      Left            =   11880
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   15
      Left            =   12480
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   15
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   16
      Left            =   13080
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   16
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1035
      Left            =   -120
      TabIndex        =   82
      Top             =   6720
      Width           =   6810
   End
   Begin VB.Label lblZlog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Scale zLog entries:"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   79
      Top             =   8280
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblZlogNumRecords 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   78
      Top             =   8535
      Visible         =   0   'False
      Width           =   2340
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
      Index           =   0
      Left            =   4320
      TabIndex        =   12
      Top             =   420
      Width           =   2400
   End
   Begin VB.Label lblScalePort 
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
      Index           =   0
      Left            =   3135
      TabIndex        =   3
      Top             =   420
      Width           =   1095
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
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   420
      Width           =   975
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   0
      TabIndex        =   40
      Top             =   6375
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   0
      TabIndex        =   39
      Top             =   5640
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   0
      TabIndex        =   38
      Top             =   4905
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   0
      TabIndex        =   37
      Top             =   4170
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   0
      TabIndex        =   36
      Top             =   3435
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   0
      TabIndex        =   35
      Top             =   2700
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   0
      TabIndex        =   34
      Top             =   1965
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   33
      Top             =   1230
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   7440
      TabIndex        =   32
      Top             =   5640
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   7440
      TabIndex        =   31
      Top             =   4905
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   7440
      TabIndex        =   30
      Top             =   4170
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   7440
      TabIndex        =   29
      Top             =   3435
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   7440
      TabIndex        =   28
      Top             =   2700
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   7440
      TabIndex        =   27
      Top             =   1965
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   7440
      TabIndex        =   26
      Top             =   1230
      Width           =   1395
   End
   Begin VB.Label lblScale 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scale #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   7440
      TabIndex        =   25
      Top             =   6375
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
      Index           =   1
      Left            =   9075
      TabIndex        =   24
      Top             =   420
      Width           =   975
   End
   Begin VB.Label lblScaleNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
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
      Index           =   1
      Left            =   7440
      TabIndex        =   23
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label lblScalePort 
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
      Index           =   1
      Left            =   10335
      TabIndex        =   22
      Top             =   420
      Width           =   1095
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
      Index           =   1
      Left            =   12120
      TabIndex        =   21
      Top             =   420
      Width           =   1800
   End
   Begin VB.Label lblScaleNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
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
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   1395
   End
End
Attribute VB_Name = "frmComm8Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 9600 ''''''''''''' Form Comm8Card.frm '''''''''''''''''''
Option Explicit
'
Private WeightsQueue(1 To MAXWEIGHTQUEUE, 1 To MAX_COMM) As Single
Private InIndex(1 To MAX_COMM) As Integer
Private WeightEntries(1 To MAX_COMM) As Integer
Private NotStableCounter(1 To MAX_COMM) As Integer
Private msgCntr As Integer
Private sMsg As String
Const msgAppend = 0
Const msgReplace = 1
'
'
'
Private Function NewMsg(ByVal sNew As String, ByVal Cntrl As Integer) As String
    Select Case Cntrl
        Case msgAppend
            msgCntr = 10
            NewMsg = sMsg & vbCrLf & sNew
        Case msgReplace
            msgCntr = 0
            NewMsg = sNew
    End Select
End Function

Private Sub cmdClrZlog_Click()
    Debug_ZlogScale_Clear = True
    Write_Zlog_Scales 0, 0, "na", "Scales zLog Cleared"
End Sub

Private Sub cmdExit_Click()
'    Unload Me
'    Set frmComm8Card = Nothing
    frmComm8Card.Hide
End Sub

Private Sub cmdPortValues_Click()
    frmPortValues.Show
End Sub

Private Sub cmdSave_Click()
Dim inct As Integer
Dim error As Integer
Dim station As Integer
Dim Shift As Integer

' All Stations must be Idle to Change Scale Config
For station = 1 To LAST_STN
    For Shift = 1 To NR_SHIFT
       If StationControl(station, Shift).Mode <> VBIDLE And StationControl(station, Shift).Mode <> VBIDLEWAITING Then
          sMsg = NewMsg("Station:" & station & "  is still running in Shift " & Shift, msgReplace)
          Exit Sub
       End If
    Next Shift
Next station

' SCALES SUPPORTED ON THIS SYSTEM
error = 0
For inct = 1 To NR_SCALES
    Scale_Port(inct) = CommPort(inct).ListIndex         ' comm port number
    Select Case ScaleType(inct).ListIndex
        Case 0
            ' scale not installed
            Scale_Type(inct) = "_"
        Case 1
            ' Acculab Scale
            Scale_Type(inct) = "A"
        Case 2
            ' Sartorius Scale
            Scale_Type(inct) = "S"
        Case 3
            ' Toledo Scale
            Scale_Type(inct) = "T"
        Case 4
            ' A & D Scale
            Scale_Type(inct) = "N"
        Case 5
            ' Toledo Viper Scale
            Scale_Type(inct) = "V"
        Case Else
            ' scale not installed
            Scale_Type(inct) = "_"
    End Select
Next inct


' write the config file to save values
If error = 0 Then
    Save_ScaleConfig                ' Save configuration data to disk
    Setup_Scales                    ' Uses saved data to initialize comm modules
    sMsg = NewMsg("Configuration Saved", msgReplace)
End If
End Sub

Function Setup_Scales()
Dim scl As Integer
Dim prt As Integer
Dim Idx As Integer

    ' close the port
    For prt = 1 To MAX_COMM
        If SclComOn Then Close_Port prt
        Port_In_Use(prt) = False
        Port_Type(prt) = "_"
        Port_Weight(prt) = 0#
        Port_Value(prt) = "0.0"
        NotStableCounter(prt) = CInt(0)
        WeightEntries(prt) = CInt(0)
        InIndex(prt) = CInt(1)
        For Idx = 1 To MAXWEIGHTQUEUE
            WeightsQueue(Idx, prt) = CSng(0)
        Next Idx
    Next prt
    ' update the PORT arrays
    For scl = 1 To NR_SCALES
        prt = Scale_Port(scl)
        Port_In_Use(prt) = True
        Port_Type(prt) = Scale_Type(scl)
    Next scl

End Function

Private Sub Form_Load()
Dim inct As Integer
Dim Idx As Integer

    If NotDebugSCALES Then
        txtDebug(0).Left = OutOfSight
    Else
        txtDebug(0).Left = 0
        txtDebug(0).ToolTipText = "ScaleReading <> ascii of StatusWordA <> status word as bits" _
                                            & " <> DecimalPointCode <> Weight as String"
    End If

    For inct = 0 To 1
        lblScaleNum(inct).ForeColor = Titles_ForeColor
        lblScaleType(inct).ForeColor = Titles_ForeColor
        lblScalePort(inct).ForeColor = Titles_ForeColor
        lblScaleValue(inct).ForeColor = Titles_ForeColor
    Next inct

    For inct = 1 To MAX_COMM
        NotStableCounter(inct) = CInt(0)
        WeightEntries(inct) = CInt(0)
        InIndex(inct) = CInt(1)
        For Idx = 1 To MAXWEIGHTQUEUE
            WeightsQueue(Idx, inct) = CSng(0)
        Next Idx
    Next inct
        
    If NR_SCALES > 8 Then
        ' Show two columns of Scales
        frmComm8Card.Width = 14800
        cmdExit.Left = 13320
        text.Left = 3700
        text.Width = 7250
    Else
        ' Show one column of Scales
        frmComm8Card.Width = 7600
        cmdExit.Left = 5860
        text.Left = 100
        text.Width = 7250
    End If
    
    For inct = 1 To MAX_SCALES
        txtScaleValue(inct).ForeColor = TitlesData_Forecolor
        If inct > NR_SCALES Then
            'Don't show controls for scales that don't exist
            lblScale(inct).Visible = False
            ScaleType(inct).Visible = False
            CommPort(inct).Visible = False
            txtScaleValue(inct).Visible = False
        Else
            'Show controls for scales
            If inct < 10 Then
                lblScale(inct).Caption = "scale   " & Format(inct, "0")
            Else
                lblScale(inct).Caption = "scale  " & Format(inct, "#0")
            End If
            lblScale(inct).Visible = True
            ScaleType(inct).Visible = True
            CommPort(inct).Visible = True
            ScaleType(inct).ForeColor = DK3ORANGE
            CommPort(inct).ForeColor = DK3ORANGE
            txtScaleValue(inct).Visible = True
        End If
        CommPort(inct).ListIndex = Scale_Port(inct)
        Select Case Scale_Type(inct)
            Case "_"
                ' Scale Not Installed
                ScaleType(inct).ListIndex = 0
            Case "A"
                ' Acculab Scale
                ScaleType(inct).ListIndex = 1
            Case "S"
                ' Sartorius Scale
                ScaleType(inct).ListIndex = 2
            Case "T"
                ' Toledo Scale
                ScaleType(inct).ListIndex = 3
            Case "N"
                ' A & D Scale
                ScaleType(inct).ListIndex = 4
            Case "V"
                ' Toledo Viper Scale
                ScaleType(inct).ListIndex = 5
            Case Else
                ' Scale Not Installed
                ScaleType(inct).ListIndex = 0
        End Select
    Next inct
    cmdSave.Visible = True
 
    ' zLog Scales
    cmdClrZlog.Visible = IIf(Not NotDebugSCALES, True, False)
    lblZlog.Visible = IIf(Not NotDebugSCALES, True, False)
    lblZlogNumRecords.Visible = IIf(Not NotDebugSCALES, True, False)
    
    ' Optional Debug Text Box
    text.ForeColor = TitlesData_Forecolor
    If UseLocalErrorHandler Then
        text.Visible = False
    Else
        text.Visible = True
    End If
  
    ' message
    sMsg = " "
  
End Sub

Function Close_Scale(True_Scale_Number As Integer)
Dim scale_number As Integer
    If SclComOn Then
        If (True_Scale_Number > VALUE0 And True_Scale_Number <= NR_SCALES) Then
            scale_number = Scale_Port(True_Scale_Number)
            If MSComm(scale_number).PortOpen = True Then
                MSComm(scale_number).PortOpen = False
            End If
        Else
            sMsg = NewMsg("Error ComPort Out of Range for Scale #" & Format(True_Scale_Number, "#0"), msgAppend)
        End If
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub tmrUpdate_Timer()
Dim scl
    For scl = 1 To NR_SCALES
        txtScaleValue(scl).text = Format(Scale_Weight(scl), "######0.00")
    Next scl
    lblZlogNumRecords.Caption = CStr(Debug_ZlogScale_NumRecords)
    sMsg = IIf((msgCntr > 0), sMsg, " ")
    lblMessage.Caption = sMsg
End Sub

Public Function WhichScale(ByVal prt As Integer) As Integer
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9600, 8
Dim Idx, scl As Integer
    scl = 0
    For Idx = 1 To MAX_COMM
        If Scale_Port(Idx) = prt Then scl = Idx
    Next Idx
    WhichScale = scl
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Sub Read_Comm_Port(prt As Integer)
' open port, set speed 9600,e,7,1
' read=good => update weight
' set error if not goood
Dim buffers As String
Dim InChars As String
Dim StatusWdA As String
Dim WdAsBitsA As String
Dim StatusWdB As String
Dim WdAsBitsB As String
Dim StatusWdC As String
Dim WdAsBitsC As String
Dim DecPointCode As Integer
Dim SignChar As String * 1
Dim Idx As Integer
Dim signin As Boolean
Dim errors As Integer
Dim errcount As Integer
Dim deltatime As Double
Dim StabilityFlag As Boolean
Dim bufferLength(0 To 1) As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9600, 3

    CurCommPort = prt
    errcount = 0
    
SetupPort:
' errors = errcount/errors
    
ChgErrModule 9600, 300
    errors = 0
    If MSComm(prt).PortOpen = False Then
ChgErrModule 9600, 301
        MSComm(prt).CommPort = prt
        MSComm(prt).PortOpen = True
        errors = Not MSComm(prt).PortOpen
        MSComm(prt).Settings = "9600,e,7,1"
        MSComm(prt).InputLen = 0
        buffers = ""
        text.text = ""
        InChars = ""
        errors = Not MSComm(prt).PortOpen
        If errors Then
            MSComm(prt).PortOpen = False
            sMsg = NewMsg("Error ComPort " & Format(prt, "#0") & " Can't open port", msgAppend)
            Exit Sub
        End If
        ' Looking for First CrLf
ChgErrModule 9600, 302
        Idx = 0
        Do
            buffers = buffers & MSComm(prt).Input
            If (Idx < 12000) Then                   'Find out if we timed out, But it works
                Idx = Idx + 1
            Else
                MSComm(prt).PortOpen = False        'This is a timeout
ChgErrModule 9600, 303
                CommErrors(prt) = CommErrors(prt) + 1
                If CommErrors(prt) > MAXCOMMERRORS Then sMsg = NewMsg("Error ComPort " & Format(prt, "#0") & " Not receiving data >>" & buffers$ & "<<", msgAppend)
                If ((CommErrors(prt) > MAXCOMMERRORS) And Port_OK(prt)) Then Write_ELog "Error ComPort " & Format(prt, "#0") & " Not receiving data >>" & buffers$ & "<<"
                Port_OK(prt) = False
                Exit Sub
            End If
        Loop Until InStr(buffers, Chr(13))          ' first CrLf
'        Loop Until InStr(buffers, vbCrLf)          ' first CrLf
        ' We have the beginning of input
ChgErrModule 9600, 304
        CommReadBuffer(prt) = buffers
    End If
    
    
ReadChars:
ChgErrModule 9600, 310
    ' NOTE: One Reading = Chars between successive CrLf
    
    '   Read any new Chars
    CommReadBuffer(prt) = CommReadBuffer(prt) & MSComm(prt).Input
    
    '   EXIT (if new char is not CrLf)
    If InStr(CommReadBuffer(prt), vbCrLf) = 0 Then Exit Sub         ' no CrLf yet
    
    ' Read Complete (includes 2nd CrLf)
ChgErrModule 9600, 320
    CommErrors(prt) = 0
    CommReadString(prt) = CommReadBuffer(prt)
    CommReadBuffer(prt) = CommReadBuffer(prt) & MSComm(prt).Input
    CommReadBuffer(prt) = ""
    text.text = CommReadString(prt)
    
    ' Optionally, log New Reading (after trimming 2nd CrLf)
    If Not NotDebugSCALES Then
 ChgErrModule 9600, 321
       Dim newRead, msg As String
        newRead = Mid(text.text, 1, (Len(text.text) - 2))
        msg = "Read from Comm Port #" & Format(prt, "#0")
        Write_Zlog_Scales WhichScale(prt), prt, newRead, msg
    End If
    
    
    ' Extract the Numeric Characters from the Input String
ChgErrModule 9600, 322
    Select Case Port_Type(prt)
        Case "A"                    'Acculab
            signin = IIf((Mid(text.text, 1, 1) = "-"), True, False)
            InChars = Mid(text.text, 3, 8)
            If signin Then
               InChars = InChars * -1
            End If
            StabilityFlag = True
            NotStableCounter(prt) = 0
        
        Case "S"                    'Sartorius
            signin = IIf((Mid(text.text, 7, 1) = "-"), True, False)
            InChars = Mid(text.text, 9, 8)
            If signin Then
                InChars = InChars * -1
            End If
            StabilityFlag = True
            NotStableCounter(prt) = 0
        
        Case "T"                    'Toledo
            'signin = IIf((Mid(text.text, 7, 1) = "-"), True, False)  Floating minus is O.K.
            InChars = Mid(text.text, 6, 8)
            StabilityFlag = True
            NotStableCounter(prt) = 0
        
        Case "N"                    'A & D
            InChars = Mid(text.text, 4, 9)
            StabilityFlag = IIf((Mid(text.text, 1, 2) = "ST"), True, False)
            NotStableCounter(prt) = IIf(StabilityFlag, 0, NotStableCounter(prt) + 1)

        Case "V"                    'Toledo Viper
            ' status word A
            StatusWdA = Mid(text.text, 2, 1)
            WdAsBitsA = Decimal2Binary(Asc(StatusWdA))
            ' extract Decimal Point Location
            DecPointCode = 0
            If (Mid(WdAsBitsA, 13, 1) = "1") Then DecPointCode = DecPointCode + 4
            If (Mid(WdAsBitsA, 14, 1) = "1") Then DecPointCode = DecPointCode + 2
            If (Mid(WdAsBitsA, 15, 1) = "1") Then DecPointCode = DecPointCode + 1
            ' status word B
            StatusWdB = Mid(text.text, 3, 1)
            WdAsBitsB = Decimal2Binary(Asc(StatusWdB))
            ' extract Positive/Negative indicator
            SignChar = "+"
            If (Mid(WdAsBitsB, 14, 1) = "1") Then SignChar = "-"
            ' status word C
            StatusWdC = Mid(text.text, 4, 1)
            WdAsBitsC = Decimal2Binary(Asc(StatusWdC))
            ' set decimal point
            Select Case DecPointCode
                Case 0
                    ' X00
                    InChars = Mid(text.text, 5, 6) & "00"
                Case 1
                    ' X0
                    InChars = Mid(text.text, 5, 6) & "0"
                Case 2
                    ' X
                    InChars = Mid(text.text, 5, 6)
                Case 3
                    ' 0.X
                    InChars = Mid(text.text, 5, 5) & "." & Mid(text.text, 10, 1)
                Case 4
                    ' 0.0X
                    InChars = Mid(text.text, 5, 4) & "." & Mid(text.text, 9, 2)
                Case 5
                    ' 0.00X
                    InChars = Mid(text.text, 5, 3) & "." & Mid(text.text, 8, 3)
                Case 6
                    ' 0.000X
                    InChars = Mid(text.text, 5, 2) & "." & Mid(text.text, 7, 4)
                Case 7
                    ' 0.0000X
                    InChars = Mid(text.text, 5, 1) & "." & Mid(text.text, 6, 5)
            End Select
            InChars = SignChar & InChars
            StabilityFlag = True
            NotStableCounter(prt) = 0
        
        Case Else
            InChars = "0"
            StabilityFlag = True
            NotStableCounter(prt) = 0
      
    End Select
    
    ' debug display
ChgErrModule 9600, 323
    If (Not NotDebugSCALES) Then
        If (Port_Type(prt) = "V") Then
            txtDebug(0).text = "   " & text.text _
                            & " <> " & Format(Asc(StatusWdA), "##0") _
                            & " <> " & WdAsBitsA _
                            & " <> " & Format(DecPointCode, "##0") _
                            & " <> " & Format(Asc(StatusWdB), "##0") _
                            & " <> " & WdAsBitsB _
                            & " <> " & "sign" & SignChar & "sign" _
                            & " <> " & ">" & InChars & "<"
        End If
    End If
    
    ' Convert the Input Characters into a Scale Weight
ChgErrModule 9600, 324
    If IsNumeric(InChars) Then
    
        ' too long since a stable reading ??
        If (NotStableCounter(prt) > MAXNOTSTABLECOUNT) Then
            NotStableCounter(prt) = 0
            StabilityFlag = True
        End If
        ' update a stable weight reading
        If StabilityFlag Then UpdateWeight InChars, prt
        ' Optional Debug Statement
        If Not NotDebugSCALES Then
ChgErrModule 9600, 326
            deltatime = 1000 * (Timer - CommPrevRead(prt))
            text.text = Format(deltatime, " #####0") & " ms to read >>>" & Mid(text.text, 1, (Len(text.text) - 2)) & "<<< from Port #" & Format(prt, "#0")
        End If
        CommPrevRead(prt) = Timer
      
    Else
    
        ' error - non numeric data received
ChgErrModule 9600, 327
        errcount = errcount + 1
        If errcount < 2 Then
            CommReadBuffer(prt) = ""
            GoTo ReadChars                      ' try again
        Else
            Write_ELog "Error Port #" & Format(prt, "#0") & " - Invalid Data >>>" & InChars & "<<<"
            sMsg = NewMsg("Error Port #" & Format(prt, "#0") & " - Invalid Data >>>" & InChars & "<<<", msgAppend)
        End If
        
    End If
      
ResetErrModule
Exit Sub

localhandler:
Dim sMsg2 As String
Dim iresponse As Integer
'*********************************************************************
If ShortTermErrorCounter < ShortTermErrorMax Then ShortTermErrorCounter = ShortTermErrorCounter + 1
' Write to Event Log
'Write_ELog "Error: " & err.Number & _
'  ", M" & ErrModule(0) & "-L" & ErrLevel(0) & " " & err.Description
UnreadProgramErrorMessage = True
Select Case err.Number
     Case 8000 To 8020                                      ' MsComm Errors
        sMsg2 = "An MsComm error for Comm Port #" & CStr(CurCommPort) & "   " & vbCrLf & vbCrLf
        sMsg2 = sMsg2 & "Error " & CStr(err.Number) & " - "
        Select Case err.Number
            Case 8000
                sMsg2 = sMsg2 & "Invalid operation on an opened port"
            Case 8001
                sMsg2 = sMsg2 & "Timeout value must be greater than zero"
            Case 8002
                sMsg2 = sMsg2 & "Invalid port number"
            Case 8003
                sMsg2 = sMsg2 & "Property available only at run-time"
            Case 8004
                sMsg2 = sMsg2 & "Property is read-only at run-time"
            Case 8005
                sMsg2 = sMsg2 & "Port is already open"
            Case 8006
                sMsg2 = sMsg2 & "Device identifier is invalid"
            Case 8007
                sMsg2 = sMsg2 & "Unsupported baud rate"
            Case 8008
                sMsg2 = sMsg2 & "Invalid byte size"
            Case 8009
                sMsg2 = sMsg2 & "Error in default parameters"
            Case 8010
                sMsg2 = sMsg2 & "Hardware is not available"
            Case 8011
                sMsg2 = sMsg2 & "Cannot allocate the queues"
            Case 8012
                sMsg2 = sMsg2 & "Device is not open"
            Case 8013
                sMsg2 = sMsg2 & "Device is already open"
            Case 8014
                sMsg2 = sMsg2 & "Could not enable Comm notification"
            Case 8015
                sMsg2 = sMsg2 & "Could not set Comm state"
            Case 8016
                sMsg2 = sMsg2 & "Could not set Comm event mask"
            Case 8017
                sMsg2 = sMsg2 & "undefined comm error 8017"
            Case 8018
                sMsg2 = sMsg2 & "Operation valid only when the port is open"
            Case 8019
                sMsg2 = sMsg2 & "Device busy"
            Case 8020
                sMsg2 = sMsg2 & "Error reading Comm device"
        End Select
        sMsg2 = sMsg2 & vbCrLf & vbCrLf
        sMsg2 = sMsg2 & "Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf & vbCrLf
        iresponse = vbAbort
        ' Write to Event Log
        Write_ELog sMsg2
    Case Else
        sMsg2 = "An error for Comm Port #" & CStr(CurCommPort) & "   " & vbCrLf & vbCrLf
        sMsg2 = sMsg2 & "Error " & CStr(err.Number) & " - "
        sMsg2 = sMsg2 & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg2 = sMsg2 & err.Description & vbCrLf & vbCrLf
        iresponse = vbAbort
        ' Write to Event Log
        Write_ELog sMsg2
End Select
sMsg = NewMsg(sMsg2, msgAppend)
'*********************************************************************
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Function Close_Port(port_number As Integer)
    If SclComOn Then
        If (port_number > VALUE0 And port_number <= MAX_COMM) Then
            If MSComm(port_number).PortOpen = True Then
                MSComm(port_number).PortOpen = False
            End If
        Else
            sMsg = NewMsg("Error Close ComPort #" & Format(port_number, "#0") & "Out of Range", msgAppend)
        End If
    End If
End Function

Private Sub UpdateWeight(ByVal newscaleread As String, ByVal prt As Integer)
' is new scale value good; dont update if not
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9600, 8
Dim newweight As Single
Dim tmpAvg As Single
Dim tmpSum As Single
Dim Idx As Integer

    newweight = ValueFromText(newscaleread)
    
    ' add new weight to the WeightsQueue
    WeightsQueue(InIndex(prt), prt) = newweight
    InIndex(prt) = IIf((InIndex(prt) < WEIGHTQUEUESIZE), (InIndex(prt) + 1), 1)
    ' calculate running average weight
    If (WeightEntries(prt) < WEIGHTQUEUESIZE) Then
        WeightEntries(prt) = WeightEntries(prt) + 1
        ' no valid average weight; use current reading
        Port_Value(prt) = newscaleread          ' Scale Weight as a String
        Port_Weight(prt) = newweight            ' Scale Weight as a Number
        Port_OK(prt) = True                     ' Back to receiving good data
        msgCntr = IIf((msgCntr > 0), (msgCntr - 1), 0)
    Else
        tmpSum = CSng(0)
        For Idx = 1 To WEIGHTQUEUESIZE
            tmpSum = tmpSum + WeightsQueue(Idx, prt)
        Next Idx
        tmpAvg = tmpSum / CSng(WEIGHTQUEUESIZE)
        ' use average weight
        Port_Value(prt) = Format(tmpAvg, "####0.00")    ' Scale Weight as a String
        Port_Weight(prt) = tmpAvg                       ' Scale Weight as a Number
        Port_OK(prt) = True                             ' Back to receiving good data
        msgCntr = IIf((msgCntr > 0), (msgCntr - 1), 0)
    End If
          
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Public Sub ScaleCommOff()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9600, 1390
    If (SclComOn) Then
        ' turn Off Scale I/O Communications
        SclComOn = False
        ScaleCommOff_Request = False
        Write_ELog "Scale Comm Off"
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

Public Sub ScaleCommOn()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9600, 1380
    If (Not SclComOn) Then
        ' Starting Scale I/O Communications
        SclComOn = True
        ScaleCommOn_Request = False
        Write_ELog "Scale Comm On"
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


