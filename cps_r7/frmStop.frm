VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmStop 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stop / Pause"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   Icon            =   "frmStop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Text            =   "Stop the current station?"
      ToolTipText     =   "Do you really, really want to stop"
      Top             =   825
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Text            =   "--- Please Confirm ---"
      Top             =   390
      Width           =   4005
   End
   Begin Threed.SSCommand cmdNo 
      Height          =   840
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1482
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   0   'False
      Picture         =   "frmStop.frx":57E2
   End
   Begin Threed.SSCommand cmdYes 
      Height          =   840
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1482
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   0   'False
      Picture         =   "frmStop.frx":AFD4
   End
End
Attribute VB_Name = "frmStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim station As Integer
Dim Shiftno As Integer

Private Sub cmdNO_Click()
    ' Delay_Box "Station returning to normal processing", MSGDELAY, msgSHOW
    Unload Me
    Stop_In_Progress = False
End Sub

Private Sub cmdYES_Click()
    StationControl(station, Shiftno).StopRequest = True
    Unload Me
End Sub

Private Sub Form_Load()          ' Save the station and shift off
    Stop_In_Progress = True
    station = DispStn
    Shiftno = DispShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Delay_Box "Station returning to normal processing", MSGDELAY, msgSHOW
    Unload Me
    Stop_In_Progress = False
End Sub
