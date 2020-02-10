VERSION 5.00
Begin VB.Form PortConfigForm 
   Caption         =   "Communication Port Properties"
   ClientHeight    =   5025
   ClientLeft      =   1935
   ClientTop       =   2265
   ClientWidth     =   4845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5025
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPortCancel 
      Caption         =   "Quit"
      DisabledPicture =   "portform.frx":0000
      DownPicture     =   "portform.frx":0C42
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
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "portform.frx":1884
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton btnPortOk 
      Caption         =   "OK"
      DisabledPicture =   "portform.frx":24C6
      DownPicture     =   "portform.frx":3108
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
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "portform.frx":3D4A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Set Backup Path "
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Frame Rs232Frame 
      Caption         =   "RS232 Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox WinApiBaud 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Text            =   "WinApiBaud"
         Top             =   720
         Width           =   1572
      End
      Begin VB.ComboBox WinApiPort 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Text            =   "WinApiPort"
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Port Number:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame AC37Frame 
      Caption         =   "AC37 Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox Ac37Baud 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Text            =   "Ac37Baud"
         Top             =   720
         Width           =   1572
      End
      Begin VB.ComboBox Ac37IoPort 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Text            =   "Ac37IoPort"
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "I/O Port Address:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame PortTypeFrame 
      Caption         =   "Communication Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.ComboBox DVF 
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Text            =   "DVF"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox editTimeOut 
         Height          =   288
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   17
         Text            =   ".5"
         Top             =   1800
         Width           =   1572
      End
      Begin VB.TextBox editRetry 
         Height          =   288
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "2"
         Top             =   1440
         Width           =   1572
      End
      Begin VB.ComboBox AsciiBinary 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Text            =   "AsciiBinary"
         Top             =   720
         Width           =   1572
      End
      Begin VB.ComboBox PortType 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "PortType"
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label lblDVF 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Verification:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Retry Count:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ascii / Binary:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Port Type:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PortConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module  82 ''''''''''' Form PORTCONFIGFORM.frm ''''''''''''''''
Option Explicit
' temporary parameters for the port configuration
Dim tPortType As Long
Dim tBrick_ioPort As Long
Dim tBrick_Port As Long
Dim tBrick_Baud As Long
Dim tBrick_TimeOut As Single   ' in seconds
Dim tBrick_Retry As Long
Dim tBrick_ProtocolType As Long

Private Sub btnPortCancel_Click()
    
    ' no port selected so quit
    PortConfigForm.Hide
    
End Sub
Private Sub btnPortOk_Click()
          
    SetErrModule 82, 1
    If UseLocalErrorHandler Then On Error GoTo localhandler
    ' be sure the port is closed
    opto22MwdPortClose (Brick_Handle)
    ' get port type
    tPortType = ComboValueGet(PortType)
    ' set the common port parameters
    Brick_TimeOut = Val(editTimeOut.text)
    Brick_Retry = Val(editRetry.text)
    Brick_ProtoType = ComboValueGet(AsciiBinary)
    Brick_CheckType = ComboValueGet(DVF)
    ' be sure error is set
    Brick_Error = -999
    ' set AC37 specific parameters
    If (tPortType = mwdPortPhysTypeAC37) Then
        Brick_ioPort = ComboValueGet(Ac37IoPort)
        Brick_Baud = ComboValueGet(Ac37Baud)
        Brick_Error = opto22MwdPortOpenAC37(Brick_Handle, _
                                            Brick_ioPort, _
                                            Brick_Baud, _
                                            Brick_TimeOut, _
                                            Brick_Retry, _
                                            Brick_ProtoType, _
                                            Brick_CheckType)
        End If  'type AC37
    ' set the RS232 specific parameters
    If (tPortType = mwdPortPhysTypeWinApi) Then
        Brick_Port = ComboValueGet(WinApiPort)
        Brick_Baud = ComboValueGet(WinApiBaud)
        Brick_Error = opto22MwdPortOpenWinApi(Brick_Handle, _
                                              Brick_Port, _
                                              Brick_Baud, _
                                              Brick_TimeOut, _
                                              Brick_Retry, _
                                              Brick_ProtoType, _
                                              Brick_CheckType)
    End If   'Type RS232
    If (Brick_Error = 0) Then
        ' close this window if no error
        PortConfigForm.Hide
        frmMainForm.btnSend.Enabled = True
    Else
        MsgBox "Could not open port.  Error # " & Str$(Brick_Error), 0, "Error"
    End If
    
ResetErrModule
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

    SetErrModule 82, 2
    If UseLocalErrorHandler Then On Error GoTo localhandler
    
    ' Set Title Foreground color
    PortTypeFrame.ForeColor = Titles_ForeColor
    AC37Frame.ForeColor = Titles_ForeColor
    Rs232Frame.ForeColor = Titles_ForeColor
    ' fill combo with port types
    PortType.Clear
    Call ComboValueAdd(PortType, "AC37 Port", mwdPortPhysTypeAC37)
    Call ComboValueAdd(PortType, "RS232 Port", mwdPortPhysTypeWinApi)
    PortType.ListIndex = 1   'default to RS232
    ' fill ascii/binary with options
    AsciiBinary.Clear
    Call ComboValueAdd(AsciiBinary, "ASCII ", ProtocolTypeIOAscii)
    Call ComboValueAdd(AsciiBinary, "Binary (Fast)", mwdProtocolTypeBinary)
    AsciiBinary.ListIndex = 0   'default to ascii
    ' fill DVF with options
    DVF.Clear
    Call ComboValueAdd(DVF, "CRC16", mwdDataCheckTypeCrc16)
    Call ComboValueAdd(DVF, "CheckSum 256", mwdDataCheckTypeCheckSum)
    DVF.ListIndex = 0
    ' fill AC37IoPort with options
    Ac37IoPort.Clear
    Call ComboValueAdd(Ac37IoPort, "3F8 (COM1)", &H3F8)
    Call ComboValueAdd(Ac37IoPort, "2F8 (COM2)", &H2F8)
    Call ComboValueAdd(Ac37IoPort, "348 (COM3 Opto)", &H348)
    Call ComboValueAdd(Ac37IoPort, "340 (COM4 Opto)", &H340)
    Call ComboValueAdd(Ac37IoPort, "248 (COM5 Opto)", &H248)
    Call ComboValueAdd(Ac37IoPort, "240 (COM6 Opto)", &H240)
    Call ComboValueAdd(Ac37IoPort, "3E8 (COM7)", &H3E8)
    Call ComboValueAdd(Ac37IoPort, "2E8 (COM8)", &H2E8)
    Select Case OPTOCOM_PORT
        Case VALUE1       ' Com 1
             Ac37IoPort.ListIndex = 0  'default to COM1
        Case VALUE2       ' Com 2
             Ac37IoPort.ListIndex = 1  'default to COM2
        Case VALUE3       ' Com 3
             Ac37IoPort.ListIndex = 2  'default to COM3
        Case VALUE4       ' Com 4
             Ac37IoPort.ListIndex = 3  'default to COM4
        Case VALUE5       ' Com 5
             Ac37IoPort.ListIndex = 4  'default to COM5
        Case VALUE6       ' Com 6
             Ac37IoPort.ListIndex = 5  'default to COM6
        Case VALUE7       ' Com 7
             Ac37IoPort.ListIndex = 6  'default to COM7 Hex address different
        Case VALUE8       ' Com 8
             Ac37IoPort.ListIndex = 7  'default to COM8 Hex address different
    End Select
    ' fill AC37Baud with options
    Ac37Baud.Clear
    Call ComboValueAdd(Ac37Baud, "  300", 300)
    Call ComboValueAdd(Ac37Baud, "  600", 600)
    Call ComboValueAdd(Ac37Baud, " 1200", 1200)
    Call ComboValueAdd(Ac37Baud, " 2400", 2400)
    Call ComboValueAdd(Ac37Baud, " 4800", 4800)
    Call ComboValueAdd(Ac37Baud, " 9600", 9600)
    Call ComboValueAdd(Ac37Baud, "19200", 19200)
    Call ComboValueAdd(Ac37Baud, "38400", 38400)
    Call ComboValueAdd(Ac37Baud, "57600", 57600)
    Call ComboValueAdd(Ac37Baud, "76800", 76800)
    Call ComboValueAdd(Ac37Baud, "115200", 115200)
    Call ComboValueAdd(Ac37Baud, "172800", 172800)  'Expansion
    Call ComboValueAdd(Ac37Baud, "230400", 230400)  'Expansion
    Ac37Baud.ListIndex = 7   'default to 38400
    WinApiBaud.Clear
    Call ComboValueAdd(WinApiBaud, "  300", 300)
    Call ComboValueAdd(WinApiBaud, "  600", 600)
    Call ComboValueAdd(WinApiBaud, " 1200", 1200)
    Call ComboValueAdd(WinApiBaud, " 2400", 2400)
    Call ComboValueAdd(WinApiBaud, " 4800", 4800)
    Call ComboValueAdd(WinApiBaud, " 9600", 9600)
    Call ComboValueAdd(WinApiBaud, "19200", 19200)
    Call ComboValueAdd(WinApiBaud, "38400", 38400)
    Call ComboValueAdd(WinApiBaud, "115200", 115200)
    WinApiBaud.ListIndex = 7  'default to 38400
    WinApiPort.Clear
    Call ComboValueAdd(WinApiPort, "COM1", 1)
    Call ComboValueAdd(WinApiPort, "COM2", 2)
    Call ComboValueAdd(WinApiPort, "COM3", 3)
    Call ComboValueAdd(WinApiPort, "COM4", 4)
    Call ComboValueAdd(WinApiPort, "COM5", 5)
    Call ComboValueAdd(WinApiPort, "COM6", 6)
    Call ComboValueAdd(WinApiPort, "COM7", 7)
    Call ComboValueAdd(WinApiPort, "COM8", 8)
    Select Case OPTOCOM_PORT
        Case VALUE1       ' Com 1
             WinApiPort.ListIndex = 0  'default to COM1
        Case VALUE2       ' Com 2
             WinApiPort.ListIndex = 1  'default to COM2
        Case VALUE3       ' Com 3
             WinApiPort.ListIndex = 2  'default to COM3
        Case VALUE4       ' Com 4
             WinApiPort.ListIndex = 3  'default to COM4
        Case VALUE5       ' Com 5
             WinApiPort.ListIndex = 4  'default to COM5
        Case VALUE6       ' Com 6
             WinApiPort.ListIndex = 5  'default to COM6
        Case VALUE7       ' Com 7
             WinApiPort.ListIndex = 6  'default to COM7
        Case VALUE8       ' Com8
             WinApiPort.ListIndex = 7  'default to COM8
    End Select

ResetErrModule
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

Private Sub PortType_Click()

  '' Purpose: to highlight/disable pertintent items in this form.
  If PortType.ListIndex < 0 Then
    Exit Sub
  End If
  ' determine port type
  tPortType = ComboValueGet&(PortType)
  If (tPortType = mwdPortPhysTypeAC37) Then
    AC37Frame.Visible = True
    Rs232Frame.Visible = False
  End If
  If (tPortType = mwdPortPhysTypeWinApi) Then
    AC37Frame.Visible = False
    Rs232Frame.Visible = True
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
    HotKeyCheck KeyCode, shift  ' undo rest to display key coads
End Sub
