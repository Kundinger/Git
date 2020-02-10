VERSION 5.00
Begin VB.Form frmDelayBox 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1095
   ClientLeft      =   2355
   ClientTop       =   3540
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000017&
   Icon            =   "frmDelay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1095
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrDelayBox 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4320
      Top             =   0
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Please wait a short while"
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmDelayBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' no error mod ''''''''''''' Form DELAYBOX.frm '''''''''''''''''''
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
    Set frmDelayBox = Nothing
End Sub

Private Sub tmrDelayBox_Timer()
Dim tempp As Integer
    tempp = tmrDelayBox.Interval
    tmrDelayBox.Enabled = False
    Unload Me
    Set frmDelayBox = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Set frmDelayBox = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift
End Sub

