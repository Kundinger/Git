VERSION 5.00
Begin VB.Form frmCheckIt 
   BackColor       =   &H80000005&
   Caption         =   "  Verify Selection"
   ClientHeight    =   3195
   ClientLeft      =   1200
   ClientTop       =   1695
   ClientWidth     =   6630
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   Icon            =   "frmCheckIt.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3195
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
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
      Picture         =   "frmCheckIt.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit the Program"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCheckIt.frx":1084
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00C000C0&
      Height          =   420
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   4440
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Do you agree with this?"
      Top             =   180
      Width           =   6375
   End
End
Attribute VB_Name = "frmCheckIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'no error mod '''''''''''''''Form CKECKIT.frm''''''''''
Option Explicit

Dim iCommand As Integer

Sub CheckIt(Index As Integer, ByVal sMsg As String)
    iCommand = Index
    lblMessage = sMsg
    lblMessage.Height = TextHeight(sMsg) * 1.5
    DoEvents
    Form_Center Me
    frmCheckIt.Show
End Sub

Private Sub cmdCancel_Click()
    Unload frmCheckIt
    Set frmCheckIt = Nothing
End Sub

Private Sub cmdExit_Click()

Dim iStn As Integer
Dim iShift As Integer

ShuttingDown = True

Select Case iCommand
  Case 1 ' Exit Program
    cmdCancel.Visible = False
    cmdExit.Visible = False
    MousePointer = vbHourglass
    frmCheckIt.Caption = "GoodBye"
    ' save current butane remaining
    Save_ButaneSupply
    lblMessage = vbCrLf & vbCrLf & "CPS System Shutting Down"
    Write_ELog "CPS Reporting System Shut Down"
    ' Deenergize All Valves
    Reset_Valves
    ' Close all stations currently running
    For iStn = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
            If (StationControl(iStn, iShift).Mode <> VBIDLE And StationControl(iStn, iShift).Mode <> VBIDLEWAITING) Then
                ALM_Write iStn, iShift, "Reporting System / Data Collection Halted."
                Stats_Write iStn, iShift
                StationControl(iStn, iShift).End_Time = Now
                DoEvents
                StationControl(iStn, iShift).AbortRequest = True
                Delay_Box "", PAUSEDELAY, msgSHOW
            End If
            If iStn <= NR_SCALES Then frmComm8Card.Close_Scale iStn
        Next iShift
    Next iStn
    ' Reset Mouse Pointer
    MousePointer = vbDefault
    ' Reset Watchdog for Main Board Module 0
    Opto_Send_Data(0) = Val(200)                                        ' set Watchdog Time to 2 sec (200 * 10msec)
    If IoComOn Then frmMainForm.Send_Opto_Command 0, 114, 0, 65535      ' All Off; including Beacon
    ' Pause
    lblMsg.Caption = "All Stations are Shut Down"
    Delay_Box "", PAUSEDELAY, msgSHOW
    ' Shutdown Conn(s) to Remote DB
    If USINGREMCANLOAD Then CloseConnToRemTaskDb
    If USINGREMSTSMON Then CloseConnToRemStatusDb
    If ((USINGREMCANLOAD) Or (USINGREMSTSMON)) Then lblMsg.Caption = "All Remote Connections are Shut Down"
    Delay_Box "", PAUSEDELAY, msgSHOW
    ' Done; End Program
    End
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmCheckIt = Nothing
    End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

