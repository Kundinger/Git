VERSION 5.00
Begin VB.Form frmBackupPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SelectBackup Path"
   ClientHeight    =   4095
   ClientLeft      =   1095
   ClientTop       =   1020
   ClientWidth     =   5775
   Icon            =   "frmBackupPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         DisabledPicture =   "frmBackupPath.frx":57E2
         DownPicture     =   "frmBackupPath.frx":6424
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
         Left            =   4320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBackupPath.frx":7066
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Set Backup Path "
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Quit"
         DisabledPicture =   "frmBackupPath.frx":7CA8
         DownPicture     =   "frmBackupPath.frx":88EA
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
         Left            =   4320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmBackupPath.frx":952C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Drives Available"
         Top             =   3120
         Width           =   3615
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "Current Directory Tree"
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Backup Drive:"
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
         TabIndex        =   5
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label txtBackupPath 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C:\newcps\data"
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
         TabIndex        =   4
         ToolTipText     =   "Backup Path"
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Backup Path:"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmBackupPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 444 ''''''''''''Form BACKUPPATH.frm '''''''''''''''''''''''
Option Explicit

Dim BackupSelect As Integer
Dim pathIn As String

Public Sub ChangeBackupSelect(ByVal BackupSelectCode As Integer)
' 0 = Select Report File Backup Path
' 1 = Select Database File Backup Path
' 2 = Select Joblist Report File Backup Path
' 3 = Select Joblist Database File Backup Path
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 444, 1
Dim msg As String
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
    
    BackupSelect = BackupSelectCode
    Select Case BackupSelect
        Case 0
            ' Report File Backup
            pathIn = frmConfig.txtRptBackupPath
        Case 1
            ' DB File Backup
            pathIn = frmConfig.txtDbfBackupPath
        Case 2
            ' Report File Backup for Joblist
            pathIn = frmJoblist.txtRptBackupPath
        Case 3
            ' DB File Backup for Joblist
            pathIn = frmJoblist.txtDbfBackupPath
        Case Else
            Exit Sub
    End Select
    
    ' Confirm trailing \
    If Mid(pathIn, Len(pathIn), 1) <> "\" Then pathIn = pathIn & "\"
    ' does path exist?
    If Not fs.FolderExists(pathIn) Then
        msg = "Backup Files Path Does Not Exist or is Empty - " & pathIn & vbCrLf
        msg = msg & vbCrLf
        msg = msg & "Using Default Backup Path - " & FILEPATH_backup & vbCrLf
        MsgBox msg, vbInformation, "Backup Path Does Not Exist"
        pathIn = FILEPATH_backup
    End If
    ' Remove trailing \
    If Mid(pathIn, Len(pathIn), 1) = "\" Then pathIn = Mid(pathIn, 1, Len(pathIn) - 1)
    ' setup screen
    txtBackupPath.Caption = pathIn
    Drive1.Drive = Left(pathIn, 1)
    Dir1.Path = pathIn
    
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

Private Sub Dir1_Change()
If UseLocalErrorHandler Then On Error GoTo localhandler
    txtBackupPath = Dir1.Path
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
      Set frmBackupPath = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
    KeyPreview = True
    Form_Center Me
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

Private Sub cmdOK_Click()
    Select Case BackupSelect
        Case 0
            ' Report Backup
            frmConfig.txtRptBackupPath = txtBackupPath & "\"
        Case 1
            ' DB File Backup
            frmConfig.txtDbfBackupPath = txtBackupPath & "\"
        Case 2
            ' Joblist Report File Backup
            frmJoblist.txtRptBackupPath = txtBackupPath & "\"
        Case 3
            ' Joblist DB File Backup
            frmJoblist.txtDbfBackupPath = txtBackupPath & "\"
    End Select
    Unload Me
    Set frmBackupPath = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Set frmBackupPath = Nothing
End Sub

