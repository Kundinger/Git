VERSION 5.00
Begin VB.Form frmSelCal 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3855
   ClientLeft      =   1185
   ClientTop       =   1560
   ClientWidth     =   4935
   ForeColor       =   &H8000000E&
   Icon            =   "frmSelCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSettings 
      Height          =   450
      Left            =   4365
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelCal.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Calibration Settings"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton cmdMfcCalCheck 
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
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelCal.frx":070E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Mass Flow Controller Calibration Check"
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdCalibrateAIs 
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
      Left            =   2047
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelCal.frx":1350
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Analog Input Calibration"
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdCalibrateMFCs 
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
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelCal.frx":1F92
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Mass Flow Controller Calibration"
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdCalibrateScales 
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
      Left            =   3975
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelCal.frx":2BD4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Scales Calibration"
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Label lblScales 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblAIs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Analog Inputs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblMfcs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MFCs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "                                                                            Select Device Type to be Calibrated"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   885
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Do you agree with this?"
      Top             =   180
      Width           =   4695
   End
End
Attribute VB_Name = "frmSelCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'no error mod '''''''''''''''Form SELCAL.frm''''''''''
Option Explicit


Private Sub cmdCalibrateAIs_Click()
    If CheckPass("X", True) Then
        frmAnalogInputCal.Show
        Unload Me
        Set frmSelCal = Nothing
    End If
End Sub

Private Sub cmdCalibrateMFCs_Click()
    If CheckPass("X", True) Then
'        frmMassFlowCal.Show
        frmMfcCal.Show
        Unload Me
        Set frmSelCal = Nothing
    End If
End Sub

Private Sub cmdCalibrateScales_Click()
    If CheckPass("X", True) Then
        frmScalesCal.Show
        Unload Me
        Set frmSelCal = Nothing
    End If
End Sub

Private Sub cmdMfcCalCheck_Click()
    frmCalCheck.Show
    frmCalCheck.SetupCalCheck 1, 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSelCal = Nothing
    End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

