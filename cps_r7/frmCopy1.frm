VERSION 5.00
Begin VB.Form frmCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Files"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   1170
   ClientWidth     =   12375
   ControlBox      =   0   'False
   Icon            =   "frmCopy1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4935
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      DisabledPicture =   "frmCopy1.frx":0442
      DownPicture     =   "frmCopy1.frx":1084
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
      Picture         =   "frmCopy1.frx":1CC6
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Quit"
      DisabledPicture =   "frmCopy1.frx":2908
      DownPicture     =   "frmCopy1.frx":354A
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
      Left            =   10800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCopy1.frx":418C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destination "
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
      Height          =   3615
      Left            =   8200
      TabIndex        =   11
      Top             =   120
      Width           =   4100
      Begin VB.DirListBox Dir2 
         Height          =   1440
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Destination Folder"
         Top             =   960
         Width           =   3600
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Destination Drive"
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblDir2Name 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination Drive:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source "
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
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   4100
         TabIndex        =   7
         ToolTipText     =   "Folder Selected"
         Top             =   960
         Width           =   3600
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   4100
         TabIndex        =   6
         ToolTipText     =   "Source Drive"
         Top             =   3120
         Width           =   3600
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmCopy1.frx":4DCE
         Left            =   240
         List            =   "frmCopy1.frx":4DE1
         TabIndex        =   5
         Text            =   "CPS Report Files (*.rpt)"
         ToolTipText     =   "Folder Search List"
         Top             =   3120
         Width           =   3615
      End
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         ToolTipText     =   "File List"
         Top             =   960
         Width           =   3600
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Source File Name"
         Top             =   480
         Width           =   3600
      End
      Begin VB.Label Label3 
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
         Left            =   4100
         TabIndex        =   18
         Top             =   240
         Width           =   3735
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
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   15
      End
      Begin VB.Label lblDir1Name 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Source Drive:"
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
         Left            =   4100
         TabIndex        =   8
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   2415
      End
   End
   Begin VB.Label lblNotify 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copying Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   840
      Left            =   1740
      TabIndex        =   17
      Top             =   3960
      Width           =   8955
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 54 ''''''''''''''Form COPY.frm ''''''''''''''''''''
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    Set frmCopy = Nothing
End Sub

Private Sub cmdCopy_Click()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 2

Dim icnt As Integer
Dim sourcename As String
Dim destname As String
Dim sTime As Date

    If Dir1.Path <> Dir2.Path Then
    
        ' Copy Selected Files
        MousePointer = vbHourglass
        For icnt = 0 To File1.ListCount - 1
            If File1.Selected(icnt) Then
            
              lblNotify.ForeColor = DK2PURPLE
              sourcename = Dir1.Path
              If Right(sourcename, 1) = "\" Then
                sourcename = sourcename & File1.List(icnt)
              Else
                sourcename = sourcename & "\" & File1.List(icnt)
              End If
              destname = Dir2.Path
              If Right(destname, 1) = "\" Then
                destname = destname & File1.List(icnt)
              Else
                destname = destname & "\" & File1.List(icnt)
              End If
              sTime = Now
              lblNotify = "Copying" & vbCrLf _
                          & sourcename & vbCrLf _
                          & " ----->" & vbCrLf _
                          & destname
              DoEvents
              FileCopy sourcename, destname
              While Now < sTime + DISPDELAY
                DoEvents
              Wend
              
            End If
        Next icnt
        
        lblNotify = "Files Copied!"
        DoEvents
        MousePointer = vbDefault
        
    Else
    
        ' !!!  Duplicate Folders  !!!
        lblNotify.ForeColor = MEDRED
        lblNotify.Caption = "Directories can not be the same"
        
    End If
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    MousePointer = vbDefault
    ResetErrModule
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub Combo1_Click()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 4

    Select Case Combo1.ListIndex
     Case 0
        File1.Pattern = "*.*"
     Case 1
        File1.Pattern = "*_Summary.RPT"
     Case 2
        File1.Pattern = "*_Detail.RPT"
     Case 3
        File1.Pattern = "*.RPT"
     Case 4
        File1.Pattern = "*" & ReportsXlsFileExt
    End Select
    
    txtFileName = File1.Pattern
    
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

Private Sub Dir1_Change()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 5

    File1.Path = Dir1.Path
    lblDir1Name = Dir1.Path

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

Private Sub Dir2_Change()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 6

    lblDir2Name = Dir2.Path
    
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

Private Sub Drive1_Change()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 7

    Dir1.Path = Drive1.Drive
    lblDir1Name = Drive1.Drive

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

Private Sub Drive2_Change()
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 8

    Dir2.Path = Drive2.Drive
    lblDir2Name = Drive2.Drive
    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmCopy = Nothing
    End If
End Sub

Private Sub Form_Load()
' FORM:     Form Copy
' Author:   Analytical Process Programmer  APS
' Description:
' The copy form allows the user to select a file or files from a soure
' directory and copy to the destination directory.
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 1
    KeyPreview = True

    Form_Center Me
    Frame1.ForeColor = Titles_ForeColor
    Frame2.ForeColor = Titles_ForeColor
    lblNotify.Caption = ""
    lblNotify.ForeColor = DK2PURPLE
    Dir1.Path = FILEPATH_reports
    Dir2.Path = FILEPATH_reports
    lblDir1Name = Dir1.Path
    lblDir2Name = Dir2.Path
    
    Combo1.ListIndex = 0
    
    Select Case Combo1.ListIndex
     Case 0
        File1.Pattern = "*.*"
     Case 1
        File1.Pattern = "*_Summary.RPT"
     Case 2
        File1.Pattern = "*_Detail.RPT"
     Case 3
        File1.Pattern = "*.RPT"
     Case 4
        File1.Pattern = "*" & ReportsXlsFileExt
    End Select
    
    txtFileName = File1.Pattern
    
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

Private Sub txtFileName_Change()

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 54, 3

File1.Pattern = txtFileName

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

