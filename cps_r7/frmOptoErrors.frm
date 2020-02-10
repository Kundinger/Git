VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmOptoErrors 
   Caption         =   "Opto IO Error Monitor"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   750
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   1995
   End
   Begin VB.TextBox txtErrorMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmOptoErrors.frx":0000
      Top             =   780
      Width           =   6200
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   3240
      Top             =   2040
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   750
   End
   Begin Threed.SSPanel txtDispOpto 
      Height          =   600
      Left            =   990
      TabIndex        =   4
      ToolTipText     =   "Station Number Displayed"
      Top             =   1920
      Width           =   810
      _Version        =   65536
      _ExtentX        =   1429
      _ExtentY        =   1058
      _StockProps     =   15
      Caption         =   "49"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   3
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label lblOptoError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OPTO22 ERROR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   300
      Width           =   6135
   End
End
Attribute VB_Name = "frmOptoErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
   frmMainMenu.Show
End Sub

Private Sub cmdNext_Click()
    Disp_Opto = IIf(Disp_Opto >= ((NR_STN * 4) + 3), 0, Disp_Opto + 1)
    Form_Load
End Sub

Private Sub cmdPrev_Click()
    Disp_Opto = IIf(Disp_Opto <= 0, ((NR_STN * 4) + 3), Disp_Opto - 1)
    Form_Load
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Set frmOptoErrors = Nothing     'current form
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 92, 1

    lblOptoError.ForeColor = Titles_ForeColor
    txtDispOpto.ForeColor = Titles_ForeColor
    txtDispOpto = Disp_Opto
    txtErrorMsg = Opto_COMM_ERROR(Disp_Opto)          ' Current error string from this address
    
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

Private Sub tmrUpdate_Timer()
    If frmOptoErrors.Visible = True Then
       Form_Load
    End If
End Sub
