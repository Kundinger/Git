VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmCommonTC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Common TC's"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2760
   ClipControls    =   0   'False
   Icon            =   "frmCommonTC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2760
   Begin VB.Timer tmrUpdate 
      Interval        =   200
      Left            =   840
      Top             =   2040
   End
   Begin Threed.SSPanel pnlTC2 
      Height          =   345
      Left            =   945
      TabIndex        =   0
      ToolTipText     =   "Thermocouple Two Value"
      Top             =   465
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "199.9"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC1 
      Height          =   345
      Left            =   945
      TabIndex        =   1
      ToolTipText     =   "Thermocouple One Value"
      Top             =   120
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "97.4"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC4 
      Height          =   345
      Left            =   945
      TabIndex        =   2
      ToolTipText     =   "Thermocouple Four Value"
      Top             =   1155
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "98.6"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC3 
      Height          =   345
      Left            =   945
      TabIndex        =   3
      ToolTipText     =   "Thermocouple Three Value"
      Top             =   810
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "97.4"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC6 
      Height          =   345
      Left            =   945
      TabIndex        =   4
      ToolTipText     =   "Thermocouple Six Value"
      Top             =   1845
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "98.6"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel pnlTC5 
      Height          =   345
      Left            =   945
      TabIndex        =   5
      ToolTipText     =   "Thermocouple Five Value"
      Top             =   1500
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   609
      _StockProps     =   15
      Caption         =   "97.4"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin VB.Label lblTC 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   17
      Top             =   510
      Width           =   800
   End
   Begin VB.Label lblTC 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   165
      Width           =   800
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   2100
      TabIndex        =   15
      Top             =   142
      Width           =   615
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   2100
      TabIndex        =   14
      Top             =   487
      Width           =   615
   End
   Begin VB.Label lblTC 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   135
      TabIndex        =   13
      Top             =   855
      Width           =   800
   End
   Begin VB.Label lblTC 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   135
      TabIndex        =   12
      Top             =   1200
      Width           =   800
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   2100
      TabIndex        =   11
      Top             =   832
      Width           =   615
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   2100
      TabIndex        =   10
      Top             =   1177
      Width           =   615
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   2100
      TabIndex        =   9
      Top             =   1867
      Width           =   615
   End
   Begin VB.Label lblF1 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "deg F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   2100
      TabIndex        =   8
      Top             =   1522
      Width           =   615
   End
   Begin VB.Label lblTC 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #6"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   135
      TabIndex        =   7
      Top             =   1890
      Width           =   800
   End
   Begin VB.Label lblTC 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "TC #5"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   135
      TabIndex        =   6
      Top             =   1545
      Width           =   800
   End
End
Attribute VB_Name = "frmCommonTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim DataForecolor As Long
Dim idx As Integer
    tmrUpdate.Interval = 100
    tmrUpdate.Enabled = True
    frmCommonTC.Left = frmStnDetail.Left + 11000
    frmCommonTC.Top = frmStnDetail.Top + 8000
    DataForecolor = DataBold_ForeColor
    pnlTC1.ForeColor = DataForecolor
    pnlTC2.ForeColor = DataForecolor
    pnlTC3.ForeColor = DataForecolor
    pnlTC4.ForeColor = DataForecolor
    pnlTC5.ForeColor = DataForecolor
    pnlTC6.ForeColor = DataForecolor
    For idx = 1 To 6
        lblF1(idx).Caption = IIf(USINGC, "deg C", "deg F")
        lblTC(idx).ForeColor = TitlesData_Forecolor
        lblF1(idx).ForeColor = TitlesData_Forecolor
    Next idx
    Update
End Sub

Private Sub tmrUpdate_Timer()
    Update
End Sub

Private Sub Update()
    pnlTC1.Caption = Format(Com_AIO(acCommonTC1).EUValue, "##00.0#")
    pnlTC2.Caption = Format(Com_AIO(acCommonTC2).EUValue, "##00.0#")
    pnlTC3.Caption = Format(Com_AIO(acCommonTC3).EUValue, "##00.0#")
    pnlTC4.Caption = Format(Com_AIO(acCommonTC4).EUValue, "##00.0#")
    pnlTC5.Caption = Format(Com_AIO(acCommonTC5).EUValue, "##00.0#")
    pnlTC6.Caption = Format(Com_AIO(acCommonTC6).EUValue, "##00.0#")
End Sub


