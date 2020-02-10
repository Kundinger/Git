VERSION 5.00
Begin VB.Form frmAKcfg 
   Caption         =   "AK Comm Configuration"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawMode        =   5  'Not Copy Pen
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "frmAKcfg.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   2167
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      DrawMode        =   5  'Not Copy Pen
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "frmAKcfg.frx":02BA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   1620
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "frmAKcfg.frx":0636
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   180
      Width           =   480
   End
   Begin VB.TextBox txtTimeout 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   9
      Text            =   " "
      ToolTipText     =   "No Commands timeout in seconds (5 to 999)"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CheckBox chkLogAK 
      Alignment       =   1  'Right Justify
      Caption         =   "Log AK Commands ?"
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
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      ToolTipText     =   "Turn Logging AK commands ON or OFF"
      Top             =   2280
      Width           =   2565
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      DisabledPicture =   "frmAKcfg.frx":08AD
      DownPicture     =   "frmAKcfg.frx":0FAF
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2430
      Picture         =   "frmAKcfg.frx":16B1
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   950
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      DisabledPicture =   "frmAKcfg.frx":1DB3
      DownPicture     =   "frmAKcfg.frx":24B5
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAKcfg.frx":2BB7
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   950
   End
   Begin VB.TextBox txtSeparaterChar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   4
      Text            =   " "
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDontCareChar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      Text            =   " "
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtPort 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   0
      Text            =   "5600"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTimeout 
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout in seconds"
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
      Left            =   840
      TabIndex        =   10
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label lblSeparaterChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Separater Char (ascii)"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1245
      Width           =   2175
   End
   Begin VB.Label lblDontCareChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't Care Char (ascii)"
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
      Left            =   240
      TabIndex        =   3
      Top             =   765
      Width           =   2175
   End
   Begin VB.Label lblPort 
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
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
      Left            =   840
      TabIndex        =   1
      Top             =   285
      Width           =   615
   End
End
Attribute VB_Name = "frmAKcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLogAK_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    frmAKServer.Show
End Sub

Private Sub cmdSave_Click()
    AK_anychar = frmAKServer.anyChar
    AK_sepchar = frmAKServer.sepChar
    AK_portNumStr = frmAKServer.portNumStr
    If Not IsNumeric(txtTimeout.text) Then
        txtTimeout.text = "30"
    ElseIf CInt(txtTimeout.text) < 5 Then
        txtTimeout.text = "5"
    ElseIf CInt(txtTimeout.text) > 999 Then
        txtTimeout.text = "999"
    End If
    AK_timeout = txtTimeout.text
    LogAkCommands = IIf(chkLogAK.Value = cYES, True, False)
    ' save AK Config parameters with sysdef
    Save_SysDef
    cmdSave.Enabled = False
End Sub

Private Sub Form_Load()
    txtPort.text = frmAKServer.portNumStr
    txtDontCareChar.text = Asc(frmAKServer.anyChar)
    txtSeparaterChar.text = Asc(frmAKServer.sepChar)
    txtTimeout.text = AK_timeout
    chkLogAK.Value = IIf(LogAkCommands, cYES, cNO)
    cmdSave.Enabled = False
End Sub

Private Sub txtDontCareChar_Change()
    If IsNumeric(txtDontCareChar.text) Then
        frmAKServer.anyChar = Chr(Int(txtDontCareChar.text))
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtPort_Change()
    If IsNumeric(txtPort.text) Then
        frmAKServer.portNumStr = txtPort.text
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtSeparaterChar_Change()
    If IsNumeric(txtSeparaterChar.text) Then
        frmAKServer.sepChar = Chr(Int(txtSeparaterChar.text))
        cmdSave.Enabled = True
    End If
End Sub

Private Sub txtTimeout_Change()
    cmdSave.Enabled = True
End Sub
