VERSION 5.00
Begin VB.Form frmAnalyzer 
   BackColor       =   &H80000005&
   Caption         =   "Analyzer Setup"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2400
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Continue to Setup Screen"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Cancel and Set Entrys to Zero"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtDwellTime 
      Alignment       =   1  'Right Justify
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
      Left            =   4300
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Enter a Number Between 0 and 9999 Minutes"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtTargetConcentration 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Left            =   4300
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Enter a Number Between 0 and 1000000 Parts per Million"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Concentration Dwell Time                Minutes"
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
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Analyzer Target Concentration:           ppm"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 98 '''''''''''' Form ANALYZER.frm '''''''''''''''''''''''
Option Explicit

Private Sub cmdCancel_Click()
    txtTargetConcentration = 0
    txtDwellTime = 0
    frmAnalyzer.Visible = False
    Unload frmAnalyzer
    frmRecipe.chkUseAnalyzer.Value = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmAnalyzer = Nothing     'current form
    End If
End Sub

Private Sub cmdReturn_Click()

Dim errors As Integer
SetErrModule 98, 1

    errors = 0  '0=no errors 1=Low errors  2=High errors
    If UseLocalErrorHandler Then On Error GoTo localhandler
    
    If txtTargetConcentration < 0 Then
       txtTargetConcentration.BackColor = EntryInvalid_BackColor
       errors = 1
    End If
    If txtTargetConcentration > 1000000 Then
       txtTargetConcentration.BackColor = EntryInvalid_BackColor
       errors = 2
    End If
    If txtDwellTime < 0 Then
       txtDwellTime.BackColor = EntryInvalid_BackColor
       errors = 1
    End If
    If txtDwellTime > 9999 Then
       txtDwellTime.BackColor = EntryInvalid_BackColor
       errors = 2
    End If
    If errors = 1 Then
       Delay_Box "Number too small....See tool tips", MSGDELAY, msgSHOW
    End If
    If errors = 2 Then
       Delay_Box "Number too large....See tool tips", MSGDELAY, msgSHOW
    End If
    If errors = 0 Then
       frmAnalyzer.Visible = False
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

Private Sub txtDwellTime_Change()
    txtDwellTime.BackColor = Entry_BackColor
End Sub

Private Sub txtTargetConcentration_Change()
    txtTargetConcentration.BackColor = Entry_BackColor
End Sub
