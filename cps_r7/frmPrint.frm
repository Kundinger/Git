VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Files"
   ClientHeight    =   9135
   ClientLeft      =   345
   ClientTop       =   375
   ClientWidth     =   7590
   ControlBox      =   0   'False
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9135
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      DisabledPicture =   "frmPrint.frx":0442
      DownPicture     =   "frmPrint.frx":1084
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
      Picture         =   "frmPrint.frx":1CC6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8190
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Preview"
      DisabledPicture =   "frmPrint.frx":2908
      DownPicture     =   "frmPrint.frx":354A
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
      Left            =   4005
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrint.frx":418C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8190
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdSetFonts 
      Caption         =   "Print Setup"
      DisabledPicture =   "frmPrint.frx":4DCE
      DownPicture     =   "frmPrint.frx":5A10
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
      Left            =   2070
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrint.frx":6652
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8190
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Quit"
      DisabledPicture =   "frmPrint.frx":7294
      DownPicture     =   "frmPrint.frx":7ED6
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
      Left            =   5955
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrint.frx":8B18
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8190
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3855
      TabIndex        =   8
      ToolTipText     =   "Drive List"
      Top             =   6810
      Width           =   3600
   End
   Begin VB.DirListBox Dir1 
      Height          =   5490
      Left            =   3855
      TabIndex        =   7
      ToolTipText     =   "Directory Tree"
      Top             =   840
      Width           =   3600
   End
   Begin VB.FileListBox File1 
      Height          =   5550
      Left            =   120
      MultiSelect     =   2  'Extended
      Pattern         =   "*.??A;*.??B;*.??C"
      TabIndex        =   6
      ToolTipText     =   "File List"
      Top             =   840
      Width           =   3600
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPrint.frx":975A
      Left            =   120
      List            =   "frmPrint.frx":976A
      TabIndex        =   5
      Text            =   "Summary Files (OE2*.??B)"
      ToolTipText     =   "File Mask"
      Top             =   6810
      Width           =   3555
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Print File Name"
      Top             =   360
      Width           =   3600
   End
   Begin VB.Label lblNotify 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Files:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   7305
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Drive:"
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
      Left            =   3855
      TabIndex        =   9
      Top             =   6570
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "List Files of Type:"
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
      TabIndex        =   4
      Top             =   6570
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
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
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Folders:"
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
      Left            =   3855
      TabIndex        =   2
      ToolTipText     =   "Print File Folder"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDirName 
      BackStyle       =   0  'Transparent
      Caption         =   "c:\newcps\data"
      Height          =   255
      Left            =   3855
      TabIndex        =   1
      ToolTipText     =   "Folders to look in"
      Top             =   375
      Width           =   3600
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 91 ''''''''''''' Form PRINT.frm ''''''''''''''''''''''''
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
    Set frmPrint = Nothing
End Sub

Private Sub cmdDisplay_Click()
Dim icnt As Integer
Dim sourcename As String
Dim sTime As Date

SetErrModule 91, 0
If UseLocalErrorHandler Then On Error GoTo localhandler

'Delay_Box "View Selected Files!", MSGDELAY
MousePointer = vbHourglass
DoEvents

For icnt = 0 To File1.ListCount - 1
  If File1.Selected(icnt) Then
'    lblNotify.Visible = True
    lblNotify.ForeColor = DK2PURPLE
    sourcename = Dir1.Path
    If Right(sourcename, 1) = "\" Then
      sourcename = sourcename & File1.List(icnt)
    Else
      sourcename = sourcename & "\" & File1.List(icnt)
    End If
    sTime = Now
    lblNotify = "Opening " & sourcename
    DoEvents
    Shell "notepad " & sourcename
'    lblNotify.Visible = False
  End If
Next icnt

lblNotify.ForeColor = DK2PURPLE
lblNotify = "Files Sent to Notepad!"
DoEvents
MousePointer = vbDefault

Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub cmdPrint_Click()

Dim icnt As Integer
Dim sourcename As String
Dim sTime As Date

SetErrModule 91, 1
If UseLocalErrorHandler Then On Error GoTo localhandler

'lblNotify = "Printing Selected Files!"
MousePointer = vbHourglass
DoEvents

For icnt = 0 To File1.ListCount - 1
  If File1.Selected(icnt) Then
    lblNotify.ForeColor = MEDPURPLE
    sourcename = Dir1.Path
    If Right(sourcename, 1) = "\" Then
      sourcename = sourcename & File1.List(icnt)
    Else
      sourcename = sourcename & "\" & File1.List(icnt)
    End If
    sTime = Now
    lblNotify = "Printing " & sourcename
    DoEvents
    Print_File sourcename
    While Now < sTime + DISPDELAY
      DoEvents
    Wend
    lblNotify.Visible = False
  End If
Next icnt

lblNotify.ForeColor = DK2PURPLE
lblNotify = "Files Sent to Printer!"
DoEvents
MousePointer = vbDefault

Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub cmdSetFonts_Click()
    frmPrintSet.Show
End Sub

Private Sub Combo1_Click()

SetErrModule 91, 2
If UseLocalErrorHandler Then On Error GoTo localhandler

Select Case Combo1.ListIndex
 Case 0
    File1.Pattern = "*.*"
 Case 1
    File1.Pattern = "*_Summary.RPT"
 Case 2
    File1.Pattern = "*_Detail.RPT"
 Case 3
    File1.Pattern = "*.RPT"
End Select

txtFileName = File1.Pattern
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

Private Sub Dir1_Change()

SetErrModule 91, 3
If UseLocalErrorHandler Then On Error GoTo localhandler

File1.Path = Dir1.Path
lblDirName = Dir1.Path

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

Private Sub Drive1_Change()

SetErrModule 91, 4
If UseLocalErrorHandler Then On Error GoTo localhandler

Dir1.Path = Drive1.Drive

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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmPrint = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Load()

SetErrModule 91, 5
If UseLocalErrorHandler Then On Error GoTo localhandler

KeyPreview = True

If CheckPass("W", False) Then
  cmdSetFonts.Visible = True
Else
  cmdSetFonts.Visible = False
End If

Form_Center Me
lblNotify.Caption = ""
lblNotify.ForeColor = MEDPURPLE
Dir1.Path = FILEPATH_reports
lblDirName = Dir1.Path
Combo1.ListIndex = 3

Select Case Combo1.ListIndex
 Case 0
    File1.Pattern = "*.*"
 Case 1
    File1.Pattern = "*_Summary.RPT"
 Case 2
    File1.Pattern = "*_Detail.RPT"
 Case 3
    File1.Pattern = "*.RPT"
End Select

txtFileName = File1.Pattern

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

Private Sub txtFileName_Change()
SetErrModule 91, 6
If UseLocalErrorHandler Then On Error GoTo localhandler
    File1.Pattern = txtFileName
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
