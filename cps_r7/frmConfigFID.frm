VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmConfigFID 
   BackColor       =   &H80000005&
   Caption         =   "FID I/O Configuration"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbCapture 
      Height          =   1815
      Left            =   6960
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   24
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Frame frmOffsets 
      BackColor       =   &H80000005&
      Caption         =   "Range and Offsets"
      ForeColor       =   &H80000002&
      Height          =   3345
      Left            =   -45
      TabIndex        =   6
      ToolTipText     =   "Display/Enter FID Configuration"
      Top             =   30
      Width           =   5265
      Begin VB.ComboBox combTC_Type 
         Height          =   315
         ItemData        =   "frmConfigFID.frx":0000
         Left            =   3840
         List            =   "frmConfigFID.frx":0016
         TabIndex        =   23
         Text            =   "J"
         ToolTipText     =   "Thermocouple type"
         Top             =   615
         Width           =   555
      End
      Begin VB.Frame frmFidCfg 
         BackColor       =   &H80000005&
         Caption         =   "FID Configuration"
         ForeColor       =   &H80000002&
         Height          =   1815
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   5055
         Begin VB.TextBox txtBoxVolume 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2130
            MaxLength       =   4
            TabIndex        =   20
            Text            =   "1"
            ToolTipText     =   "Enter box volume in cubic feet"
            Top             =   1290
            Width           =   500
         End
         Begin VB.TextBox txtRFactor 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2130
            MaxLength       =   4
            TabIndex        =   18
            Text            =   "1"
            ToolTipText     =   "Enter numeric R Factor"
            Top             =   855
            Width           =   500
         End
         Begin VB.TextBox txtKFactor 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2145
            MaxLength       =   4
            TabIndex        =   16
            Text            =   "1"
            ToolTipText     =   "Enter numeric K Factor "
            Top             =   435
            Width           =   500
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "cu ft"
            Height          =   255
            Left            =   2805
            TabIndex        =   22
            Top             =   1380
            Width           =   495
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Box Volume"
            Height          =   255
            Left            =   345
            TabIndex        =   21
            Top             =   1335
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "'R' Factor"
            Height          =   255
            Left            =   345
            TabIndex        =   19
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "'K' Factor"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.TextBox txtTC_Offset 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2505
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "1"
         ToolTipText     =   "Enter an Offset"
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtFID_Offset 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2505
         TabIndex        =   2
         Text            =   "N/A"
         ToolTipText     =   "Enter an Offset N/A on FIDs"
         Top             =   990
         Width           =   1005
      End
      Begin VB.TextBox txtFID_Range 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "1"
         ToolTipText     =   "Fid output expressed as a percent"
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Range                "
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Offset           "
         Height          =   285
         Left            =   2610
         TabIndex        =   11
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Thermocouple"
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "FID Output"
         Height          =   255
         Left            =   105
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   315
         Left            =   3600
         TabIndex        =   8
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdFID 
      Caption         =   "&FID  I/O Forcing"
      Height          =   825
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Unload FID Analog Range and Load Shed I/O Configuration"
      Top             =   3600
      Width           =   1455
   End
   Begin Threed.SSCommand cmdPrint 
      Height          =   765
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Print FID Configuration"
      Top             =   4680
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Print"
      BevelWidth      =   4
   End
   Begin Threed.SSCommand cmdReturn 
      Height          =   765
      Left            =   3735
      TabIndex        =   3
      ToolTipText     =   "Return to Previous Screen"
      Top             =   4650
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Exit"
      BevelWidth      =   4
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   795
      Left            =   2025
      TabIndex        =   4
      ToolTipText     =   "Save FID Configuration"
      Top             =   4650
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   1402
      _StockProps     =   78
      Caption         =   "&Save"
      BevelWidth      =   4
   End
   Begin Threed.SSPanel pnlTest_Type 
      Height          =   930
      Left            =   2055
      TabIndex        =   13
      ToolTipText     =   "I/OType"
      Top             =   3600
      Width           =   3180
      _Version        =   65536
      _ExtentX        =   5609
      _ExtentY        =   1640
      _StockProps     =   15
      Caption         =   "FID"
      ForeColor       =   -2147483646
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   3
   End
End
Attribute VB_Name = "frmConfigFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 97'''''''Form ConfigFID.frm ''''''''''''''''''''''''
Option Explicit

Dim FidStn As Integer
Dim FidShift As Integer

Sub Refresh_Fid()
'
' Function Name     Refresh_Fid
' Author            Analytical Process Programmer     1/3/2002
' Description
'SHED configurations screen
'

SetErrModule 97, 2
If UseLocalErrorHandler Then On Error GoTo localhandler

    Load_FidConfig                ' Get data off the disk first

    txtTC_Offset = TC_Offset
    combTC_Type = TC_Type
    txtFID_Range = FID_Range
    txtKFactor = KFactor
    txtRFactor = RFactor
    txtBoxVolume = BoxVolume
    
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

Sub Update_Fid()
'
' Function Name     Update_Fid
' Author            Analytical Process Programmer     6/1/2000
' Description
' Copies to station variables the screen values for
' the Fid Configuration Screen, Page 1, updates station
' values for current station shown only.
'
SetErrModule 97, 3
If UseLocalErrorHandler Then On Error GoTo localhandler

  TC_Offset = txtTC_Offset
  TC_Type = combTC_Type
  FID_Range = txtFID_Range
  KFactor = txtKFactor
  RFactor = txtRFactor
  BoxVolume = txtBoxVolume
  
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

Private Sub cmdPrint_Click()
    ' capture the form
    Set pbCapture.Picture = CaptureForm(Me)
    ' print the form
    PrintPictureToFitPage Printer, pbCapture.Picture
    Printer.EndDoc
    ' short delay
    DoEvents
'    Delay_Box "", PAUSEDELAY, msgNOSHOW
    ' clear capture box
    Set pbCapture.Picture = Nothing
    DoEvents
End Sub

Private Sub combTC_Type_Change()
     combTC_Type.BackColor = Entry_BackColor
End Sub

Private Sub cmdFID_Click()
   Unload Me
   frmIoMonitor.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
  Set frmConfigFID = Nothing
End If
End Sub

Private Sub Form_Load()
    frmFidCfg.ForeColor = Titles_ForeColor
    frmOffsets.ForeColor = Titles_ForeColor
    pnlTest_Type.ForeColor = Titles_ForeColor
    KeyPreview = True
    Refresh_Fid
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)

  HotKeyCheck KeyCode, shift  ' undo rest to display key coads

End Sub

Private Sub cmdSave_Click()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 7
Dim errors, station, shift As Integer
    
    For station = 1 To LAST_STN
        For shift = 1 To NR_SHIFT
            If StationControl(station, shift).Mode <> VBIDLE Then
                ' Report an error
                Delay_Box "Station:" & station & "  is still running in Shift " & shift, MSGDELAY, msgSHOW
                Exit Sub
            End If
        Next shift
    Next station

    If CheckPass("P", True) Then
        errors = 0              '0=no errors 1=Low errors 2 =High errors
        If IsNumeric(txtFID_Range) Then
            If txtFID_Range < 1 Then
                errors = 1
                txtFID_Range.BackColor = EntryInvalid_BackColor
            End If
            If txtFID_Range > 100 Then
                errors = 2
                txtFID_Range.BackColor = EntryInvalid_BackColor
            End If
        Else
            errors = 3
            txtFID_Range.BackColor = EntryInvalid_BackColor
        End If
        If combTC_Type <> "J" Then
            errors = 4
            combTC_Type.BackColor = EntryInvalid_BackColor
        End If
        If Not IsNumeric(txtKFactor) Then
            errors = 3
            txtKFactor.BackColor = EntryInvalid_BackColor
        End If
        If Not IsNumeric(txtRFactor) Then
            errors = 3
            txtRFactor.BackColor = EntryInvalid_BackColor
        End If
        If Not IsNumeric(txtBoxVolume) Then
            errors = 3
            txtBoxVolume.BackColor = EntryInvalid_BackColor
        End If
        If errors = 1 Then
            Delay_Box "Number too small....See tool tips", MSGDELAY, msgSHOW
        End If
        If errors = 2 Then
            Delay_Box "Number too large....See tool tips", MSGDELAY, msgSHOW
        End If
        If errors = 3 Then
            Delay_Box "Field MUST be Numeric", MSGDELAY, msgSHOW
        End If
        If errors = 4 Then
            Delay_Box "MUST be Type 'J' Thermocouple ONLY", MSGDELAY, msgSHOW
        End If
        
        If errors = 0 Then
            'Save current screen values
            Update_Fid
            Delay_Box "Fid Configuration Saved", MSGDELAY, msgSHOW
            Save_FidConfig
        End If
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

Private Sub txtBoxVolume_Change()
   txtBoxVolume.BackColor = Entry_BackColor
End Sub

Private Sub txtFID_Range_Change()
   txtFID_Range.BackColor = Entry_BackColor
End Sub

Private Sub txtKFactor_Change()
   txtKFactor.BackColor = Entry_BackColor
End Sub

Private Sub txtRFactor_Change()
   txtRFactor.BackColor = Entry_BackColor
End Sub
