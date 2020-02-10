VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   13260
   ClientLeft      =   1470
   ClientTop       =   3840
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   13260
   ScaleWidth      =   7560
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7320
      ScaleHeight     =   975
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   5520
      Width           =   255
   End
   Begin VB.PictureBox frmMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5625
      Left            =   15
      ScaleHeight     =   5625
      ScaleWidth      =   7545
      TabIndex        =   8
      Top             =   6505
      Width           =   7545
      Begin VB.TextBox txtMessage 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3180
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmAbout.frx":57E2
         Top             =   360
         Width           =   7395
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5595
      Left            =   -20
      Picture         =   "frmAbout.frx":57ED
      ScaleHeight     =   5535
      ScaleWidth      =   7530
      TabIndex        =   0
      Top             =   -40
      Width           =   7595
      Begin VB.Label cmdMMW 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label CfgRevLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cfg/Sysdef Revision Level"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         ToolTipText     =   "Cfg/Sysdef File Revision Level"
         Top             =   4710
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label DbfRevLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DB Files Revision Level"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         ToolTipText     =   "Revision Level"
         Top             =   4455
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label Release 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Release 2.1.2,  November 2004"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   3915
         TabIndex        =   5
         ToolTipText     =   "Release data"
         Top             =   4200
         Width           =   3495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail:  sales@aps-mich.com"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Analytical Process Systems, Inc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1771 Harmon Rd. Auburn Hills,  Michigan  48326"
         Height          =   255
         Left            =   210
         TabIndex        =   2
         Top             =   4380
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1(248)393-0700   FAX 1(248)393-0800"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   4650
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   15
      ScaleHeight     =   975
      ScaleWidth      =   7545
      TabIndex        =   9
      Top             =   5525
      Width           =   7545
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Exit"
         DisabledPicture =   "frmAbout.frx":8E17F
         DownPicture     =   "frmAbout.frx":8EDC1
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
         Left            =   6480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAbout.frx":8FA03
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDocumentation 
         Caption         =   "Manual"
         DisabledPicture =   "frmAbout.frx":90645
         DownPicture     =   "frmAbout.frx":91287
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
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAbout.frx":91EC9
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Setup"
         DisabledPicture =   "frmAbout.frx":92B0B
         DownPicture     =   "frmAbout.frx":9374D
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
         Left            =   3360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAbout.frx":9438F
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   90
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmdRelNotes 
         Caption         =   "Notes"
         DisabledPicture =   "frmAbout.frx":94FD1
         DownPicture     =   "frmAbout.frx":95C13
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
         Left            =   5040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmAbout.frx":96855
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Timer tmrUpdate 
         Interval        =   5
         Left            =   4560
         Top             =   480
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2160
         ScaleHeight     =   735
         ScaleWidth      =   1200
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox pbxYelSub 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6480
         Picture         =   "frmAbout.frx":97497
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   270
         Width           =   495
      End
      Begin VB.Timer tmrStartup 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4320
         Top             =   120
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   60
         Picture         =   "frmAbout.frx":97B99
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label lblMMW 
         BackStyle       =   0  'Transparent
         Caption         =   "Software by Mark Winders"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   7030
         TabIndex        =   11
         Top             =   390
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''' Form ABOUT.FRM ''''''''''''''''''''''''''''
Option Explicit

Dim sStartup As Boolean
Dim onCMD As Boolean
Dim sString As String

Public Sub UpdateMsg(ByVal newtext As String)
    txtMessage.text = txtMessage.text & newtext
    txtMessage.SelStart = IIf(Len(txtMessage.text) > 0, Len(txtMessage.text) - 1, 0)
    txtMessage.Refresh
End Sub

Private Sub cmdMMW_Click()
    onCMD = True
    lblMMW.Width = 2500
End Sub

Private Sub cmdRelNotes_Click()
    ShowDoc "RelNotes"
End Sub

Private Sub cmdReturn_Click()
    If sStartup Then
        ' Startup
        sString = "You are about to Exit the Program!" & vbCrLf & _
        vbCrLf & "Exiting will close all currently running files." & vbCrLf _
        & vbCrLf & "Are you sure you wish to Exit?"
        frmCheckIt.CheckIt 1, sString
    Else
        ' Program is already running
        Unload Me
        Set frmAbout = Nothing
    End If
End Sub

Private Sub cmdSetup_Click()
    ReadyToRun = False
    ' go directly to sysdef if already logged on as aps
    If CheckPass("H", False) Then
        UserLoginOk = True
        ' Show System definition Screen
        frmSysDefMain.Show
    Else
        UserLoginOk = False
        ' Show Password Screen
        frmPassword.Show
    End If
End Sub

Private Sub cmdDocumentation_Click()
    menuOperatorManual
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmAbout = Nothing     'current form
    End If
End Sub

Private Sub Form_Load()
Dim onAPS As Boolean
    onAPS = False
    About_Counter = 0
    If Not ReadyToRun Then
        ' Startup
        sStartup = True
        tmrStartup.Enabled = True
        cmdRelNotes.Visible = False
        txtMessage.ForeColor = Message_ForeColor
        Height = 10535
    Else
        ' Program is already running
        sStartup = False
        tmrStartup.Enabled = False
        cmdSetup.Visible = False
        Height = 6535
        If CheckPass("0", False) Then
            cmdRelNotes.Top = cmdSetup.Top
            cmdRelNotes.Left = cmdSetup.Left
            cmdRelNotes.Visible = True
        Else
            cmdRelNotes.Visible = False
        End If
    End If
    
    KeyPreview = True
    CfgRevLevel.ForeColor = TitlesData_Forecolor
    DbfRevLevel.ForeColor = TitlesData_Forecolor
    If InStr(USINGRELEASEDATE, "Release Version") > 0 Then
        ' Release Version
        Release.ForeColor = TitlesData_Forecolor
    Else
        ' Debug Version
        Release.ForeColor = Warning_ForeColor
    End If
    Release.Caption = USINGRELEASEDATE
    UpdateRev
    Form_Center Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub tmrStartup_Timer()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 995, 995


If VarInitDone Then
    If (About_Counter < 32767) Then About_Counter = About_Counter + 1
End If

If (About_Counter > 2) And UserLoginOk And ReadyToRun Then

    tmrStartup.Enabled = False
    cmdSetup.Visible = False
    
    Select Case LocalPagControl.Type
        Case pagClient
            'Start AK Server
            UpdateMsg "Starting AK Server" & vbCrLf
            frmAKServer.Hide
            'Start AK Client
            UpdateMsg "Starting AK Client" & vbCrLf
            frmAkClient.Hide
        Case pagMaster
            'Start AK Server
            UpdateMsg "Starting AK Server" & vbCrLf
            frmAKServer.Hide
        Case pagNone, pagAlone
            ' no AK client or server
    End Select

    If IoComOn Then
    ' Starting I/O Communications
      frmMainForm.Hide
      frmMainForm.PowerUPClear
      
    ' setup scales
      UpdateMsg "Starting Scale RS232 Communications" & vbCrLf
      frmComm8Card.Hide
      frmComm8Card.Setup_Scales
      Delay_Box "", INTRODELAY, msgNOSHOW

    ' Reset all Valves
      UpdateMsg "Turn Off All Valves" & vbCrLf
      Reset_Valves
      Delay_Box "", INTRODELAY, msgNOSHOW
    
    ' I/O Startup Complete
      UpdateMsg "I/O Startup Complete" & vbCrLf
      Delay_Box "", INTRODELAY, msgNOSHOW
      
    End If
    
    IntroDone = True
        
    ' Start Processes with MainMenu
    ' Process Loops triggered by timers on MainMenu         <<< *************** <<<
    UpdateMsg "Starting Main Screen" & vbCrLf
    frmMainMenu.Hide
    Delay_Box "", PAUSEDELAY, msgNOSHOW
    
           
    UpdateMsg CStr(Now()) & "   Startup Complete" & vbCrLf
    Delay_Box "", PAUSEDELAY, msgNOSHOW
    
    ' Other (major) Screens
    UpdateMsg "Starting Screens" & vbCrLf
    frmReview.Hide
    frmDataWatcher.Hide
    frmJoblist.Hide
    frmStnDetail.Hide
    
    ' Show Overview (MainMenu)
    frmMainMenu.Show

    ' Unload this screen
    Unload Me
    Set frmAbout = Nothing
    
End If

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
'Select Case iresponse
'  Case vbAbort       ' Exit if abort
'    Exit Sub
'  Case vbRetry       ' try error line again
'    Resume
'  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
'End Select
End Sub

Private Sub tmrUpdate_Timer()
    UpdateRev
    If onCMD Then
        pbxYelSub.Left = pbxYelSub.Left - 5
        lblMMW.Left = pbxYelSub.Left + 475
        If pbxYelSub.Left < cmdDocumentation.Left Then
            pbxYelSub.Left = 6480
            lblMMW.Left = 7030
            onCMD = False
        End If
    End If
End Sub

Private Sub UpdateRev()
    If CheckPass("0", False) Or Not VarInitDone Then
        DbfRevLevel.Caption = "DB Files Revision Level = " & Format(DBFREVLVL, "#0")
        DbfRevLevel.Visible = True
        CfgRevLevel.Caption = "Config/Sysdef File Revision Level = " & Format(CfgRevLvl, "#0")
        CfgRevLevel.Visible = True
    Else
        DbfRevLevel.Visible = False
        CfgRevLevel.Visible = False
    End If
End Sub

