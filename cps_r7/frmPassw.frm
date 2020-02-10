VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Logon Screen"
   ClientHeight    =   4530
   ClientLeft      =   915
   ClientTop       =   2490
   ClientWidth     =   4455
   ControlBox      =   0   'False
   Icon            =   "frmPassw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAccess 
      Caption         =   "Access"
      DisabledPicture =   "frmPassw.frx":57E2
      DownPicture     =   "frmPassw.frx":6424
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPassw.frx":7066
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdNewPass 
      Caption         =   "OK"
      DisabledPicture =   "frmPassw.frx":7CA8
      DownPicture     =   "frmPassw.frx":88EA
      BeginProperty Font 
         Name            =   "Calibri"
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
      Picture         =   "frmPassw.frx":952C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Accept Entered User Name & Password"
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Exit"
      DisabledPicture =   "frmPassw.frx":A16E
      DownPicture     =   "frmPassw.frx":ADB0
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3495
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPassw.frx":B9F2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Frame frmMessage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   665
      Left            =   120
      TabIndex        =   10
      Top             =   2100
      Width           =   4215
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   365
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Width           =   3555
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1980
         MaxLength       =   8
         TabIndex        =   0
         ToolTipText     =   "Enter Your User Name"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Enter Your Password"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Left            =   300
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   300
         TabIndex        =   8
         Top             =   660
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   4215
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current User:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Access Allowed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblCurUser 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000017&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Current User's Name"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblCurAccess 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ABCDEFGHIJKLMNPQSTUVWXZ0123456789"
         ForeColor       =   &H80000017&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "User Access List"
         Top             =   1200
         Width           =   3855
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' no error modules '''''''''''''''' Form PASSWORD ''''''''''''''''''''
Option Explicit

Dim bOpen As Boolean

Private Sub cmdAccess_Click()
    If CheckPass("5", True) Then frmPassEdit.Show
End Sub

Private Sub cmdNewPass_Click()

If UseLocalErrorHandler Then On Error GoTo localhandler

Dim goodentry As Boolean
Dim strUser As String
Dim strCode As String
Dim waitStartTimer As Double

strUser = txtUser.text
strCode = txtPassword.text

If SetPass(strUser, strCode) Then
    lblCurUser = CurrentUser.USER
    lblCurAccess = CurrentUser.Access
    cmdAccess.Visible = IIf(CheckPass("5", False), True, False)
    cmdReturn.Visible = IIf(CurrentUser.USER = DefaultUser.USER, False, True)
    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = "Password Accepted!"
    ' Write to Event Log,
    Write_ELog "User: " & CurrentUser.USER & " Logged On."
    UserLoginOk = True
    waitStartTimer = Timer
    While Timer < (waitStartTimer + CDbl(0.001 * MSGDELAY))
        DoEvents
    Wend
    If Not ReadyToRun Then
        ' *** Program Startup ***
        ' Setup (i.e. System Definition) was Selected from Splash Screen
        If CheckPass("H", False) Then
            ' Show System definition Screen and unload this screen
            frmSysDefMain.Show
            Unload Me
            Set frmPassword = Nothing
        Else
            ' Stay on this screen and wait for another user login
            Write_ELog "User: " & strUser & " Access to System Setup Denied"
            lblMessage.ForeColor = MEDRED
            lblMessage.Caption = "Access to System Setup Denied for this User !"
        End If
    Else
        ' Normal Exit
        '   update MainMenu Toolbars & Menus
        frmMainMenu.UpdateNavigateBtns
        '   unload this screen
        Unload Me
        Set frmPassword = Nothing
    End If
Else
    Write_ELog "User: " & strUser & " Failed Log On."
    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = "Invalid Password!"
    lblCurUser = CurrentUser.USER
    lblCurAccess = CurrentUser.Access
    cmdAccess.Visible = IIf(CheckPass("5", False), True, False)
    cmdReturn.Visible = IIf(CurrentUser.USER = DefaultUser.USER, False, True)
End If

Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub cmdReturn_Click()
    UserLoginOk = True
    ReadyToRun = True
    Unload Me
    Set frmPassword = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
Dim ShiftDown, AltDown, CtrlDown, alldown As Boolean
Const vbShiftMask = 1
Const vbCtrlMask = 2
Const vbAltMask = 4

    AltDown = (Shift And vbAltMask) > 0
    
    If KeyCode = 27 Then
      Unload Me
      Set frmPassword = Nothing
    End If
    
    If KeyCode = vbKeyD And AltDown Then
        If CheckPass("5", False) Then
            bOpen = Not bOpen
'            Height = IIf(bOpen, 4860, 3300)
            Height = IIf(bOpen, 5010, 3300)
        Else
            lblMessage.ForeColor = MEDRED
            lblMessage.Caption = "Insufficient Access"
        End If
    End If
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
    
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub Form_Load()
    UserLoginOk = False
    KeyPreview = True
    bOpen = False
    Height = 3300
    Form_Center Me
    Show
    lblCurUser = CurrentUser.USER
    lblCurAccess = CurrentUser.Access
    cmdAccess.Visible = IIf(CheckPass("5", False), True, False)
    cmdReturn.Visible = IIf(CurrentUser.USER = DefaultUser.USER, False, True)
End Sub

Private Sub txtPassword_Change()
    lblMessage.Caption = ""
End Sub

Private Sub txtPASSWORD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If txtUser <> "" And txtPassword <> "" Then
        cmdNewPass_Click
      End If
    End If
End Sub

Private Sub txtUser_Change()
    lblMessage.Caption = ""
End Sub
