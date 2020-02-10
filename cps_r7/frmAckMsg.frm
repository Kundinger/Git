VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmAckMsg 
   BackColor       =   &H80000018&
   Caption         =   "Message"
   ClientHeight    =   1530
   ClientLeft      =   2370
   ClientTop       =   3945
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   6
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   Icon            =   "frmAckMsg.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1530
   ScaleWidth      =   4950
   Begin Threed.SSCommand cmdClose 
      Height          =   375
      Left            =   4060
      TabIndex        =   1
      ToolTipText     =   "Close this window."
      Top             =   1080
      Width           =   795
      _Version        =   65536
      _ExtentX        =   1402
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Close"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Please wait a short while"
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmAckMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' no error mod ''''''''''''' Form AckMsg.frm '''''''''''''''''''
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
    Set frmAckMsg = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Set frmAckMsg = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift
End Sub

Private Sub Form_Load()
    cmdClose.ForeColor = Titles_ForeColor
End Sub
