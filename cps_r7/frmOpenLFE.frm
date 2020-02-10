VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmOpenLFE 
   Caption         =   "Open Laminar Flow Element File"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Enter the file name"
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmOpenLFE.frx":0000
      Left            =   240
      List            =   "frmOpenLFE.frx":000A
      TabIndex        =   3
      Text            =   "Laminar Flow Element (*.LFE)"
      ToolTipText     =   "File Mask"
      Top             =   2820
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      MultiSelect     =   2  'Extended
      Pattern         =   "*.LFE"
      TabIndex        =   2
      ToolTipText     =   "File List"
      Top             =   720
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Directory Tree"
      Top             =   720
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      ToolTipText     =   "Drive List"
      Top             =   2820
      Width           =   2655
   End
   Begin Threed.SSCommand cmdSaveAs 
      Height          =   735
      Left            =   5760
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Save As"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   6
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   735
      Left            =   7560
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
   End
   Begin Threed.SSCommand cmdOpen 
      Height          =   735
      Left            =   2040
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Open"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   6
   End
   Begin Threed.SSCommand cmdSave 
      Height          =   735
      Left            =   3840
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   6
   End
   Begin Threed.SSCommand cmdNew 
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   6
   End
   Begin VB.Label lblDirName 
      Caption         =   "c:\newcps\data"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      ToolTipText     =   "Folders to look in"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
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
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Print File Folder"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label4 
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
      TabIndex        =   6
      Top             =   2580
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      Left            =   3240
      TabIndex        =   5
      Top             =   2580
      Width           =   1695
   End
End
Attribute VB_Name = "frmOpenLFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnSaveAfterEnter As Boolean
Private blnCloseAfterLostFocus As Boolean
Private blnNewSelected As Boolean
Private blnSaveAsSelected As Boolean
Private intCurrDrive As Integer

Private Sub cmdCancel_Click()
    ' Close the file dialog and return to the screen that called it
    Unload Me
    Set frmOpenLFE = Nothing
    If frmLFEDef.Visible = False Then
        frmMassFlowCal.Enabled = True
        frmMassFlowCal.SetFocus
    Else
        frmLFEDef.Enabled = True
        frmLFEDef.SetFocus
    End If
End Sub

Private Sub cmdNew_Click()
    ' Prompts user to enter a file name.  The file dialog closes
    ' and the LFE definition screen appears.
    File1.Enabled = False
    Combo1.Enabled = False
    txtFileName.Enabled = True
    txtFileName.SetFocus
    txtFileName.text = "new.LFE"
    cmdNew.Enabled = False
    cmdSave.Enabled = False
    cmdSaveAs.Enabled = False
    cmdCancel.Enabled = False
    Delay_Box "Enter a file name and press enter.", MSGDELAY, msgSHOW

    blnCloseAfterLostFocus = False
    blnNewSelected = True
End Sub

Private Sub cmdOpen_Click()
    ' Open the selected file, switching to the LFE
    ' definition form
    Dim pathname, filename As String
    
    pathname = Dir1.Path
    filename = txtFileName.text
    
    frmLFEDef.Show
    frmLFEDef.Path = pathname
    frmLFEDef.File = filename
    frmLFEDef.Read
    frmLFEDef.ShowData
    Unload Me
    Set frmOpenLFE = Nothing
  
End Sub

Private Sub cmdSave_Click()
    ' Saves the current LFE data to the file specified in the
    ' filename text box
    Dim pathname, filename As String

    pathname = Dir1.Path
    
    filename = txtFileName.text
    frmLFEDef.Path = pathname
    frmLFEDef.File = filename
    
    ' Save the current LFE data to the current file
    frmLFEDef.Save
    Delay_Box "File Saved", MSGDELAY, msgSHOW
    frmLFEDef.Enabled = True
    frmLFEDef.ShowData
    Unload Me
    Set frmOpenLFE = Nothing
    frmLFEDef.SetFocus
End Sub

Private Sub cmdSaveAs_Click()
    ' Prompts user to enter a file name in the
    ' file name text box for saving
    
    ' Save should work after the text is entered
    File1.Enabled = False
    Combo1.Enabled = False
    txtFileName.Enabled = True
    txtFileName.SetFocus
    txtFileName.text = "new.LFE"
    File1.Enabled = False
    Dir1.Enabled = False
    cmdSaveAs.Enabled = False
    Delay_Box "Enter name of file to save to and press enter.", MSGDELAY, msgSHOW
    blnSaveAsSelected = True
    blnCloseAfterLostFocus = False

End Sub

Private Sub Combo1_Click()
    ' Update the file list box for
    ' the pattern selected
    Select Case Combo1.ListIndex
        Case 0
            File1.Pattern = "*.LFE"
        Case 1
            File1.Pattern = "*.*"
    End Select

    ' txtFileName = File1.Pattern
End Sub

Private Sub Dir1_Change()
    ' Update the file list box and the directory label
    ' for a change in the directory list box
    File1.Path = Dir1.Path
    lblDirName = Dir1.Path
End Sub

Private Sub Drive1_Change()
    ' Update the directory list box for a change in
    ' the drive list box.  Also returns to the previously selected
    ' drive if a drive is not available
    Dim strPath As String
    On Error GoTo localhandler

    Dir1.Path = Drive1.Drive
    intCurrDrive = Drive1.ListIndex
    Exit Sub
localhandler:
    ' Return to the old drive if the selected drive is unavailable
    If err.Number = 68 Then
        ' Maintain the current file path
        strPath = Dir1.Path
        Drive1.ListIndex = intCurrDrive
        Dir1.Path = strPath
    End If
    Resume Next
End Sub

Private Sub File1_Click()
    ' Set the name displayed in the filename text box to the file selected
    txtFileName.text = File1.filename
End Sub

Private Sub Form_Load()
    Form_Center Me
    Dir1.Path = FILEPATH_data
    lblDirName = Dir1.Path
    Combo1.ListIndex = 0
    txtFileName.Enabled = False
    Select Case Combo1.ListIndex
        Case 0
            File1.Pattern = "*.LFE"
        Case 1
            File1.Pattern = "*.*"
    End Select
    File1.ListIndex = 0
    txtFileName.text = File1.filename

    intCurrDrive = Drive1.ListIndex
    blnSaveAfterEnter = False
    blnCloseAfterLostFocus = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

HotKeyCheck KeyCode, Shift  ' undo rest to display key coads

End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    ' Completes execution of a New file or Save As command
    ' when the {Enter} key is pressed
    Dim pathname, filename As String
    pathname = Dir1.Path
    
    filename = txtFileName.text
    
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        If blnNewSelected = True Then blnCloseAfterLostFocus = True
        If blnSaveAsSelected = True Then
            frmLFEDef.Path = pathname
            frmLFEDef.File = filename
            frmLFEDef.Save
            blnCloseAfterLostFocus = True
        End If
    End If
End Sub

Private Sub txtFileName_LostFocus()
    ' Switches to the LFE definition form when
    ' blnCloseAfterLostFocus is true
    Dim pathname, filename As String
    pathname = Dir1.Path
    
    filename = txtFileName.text
    
    If blnCloseAfterLostFocus = True Then
        
        frmLFEDef.Show
        frmLFEDef.Enabled = True
        frmLFEDef.Path = pathname
        frmLFEDef.File = filename
        frmLFEDef.ShowData
        
        Unload Me
        Set frmOpenLFE = Nothing
        frmLFEDef.SetFocus
    End If
End Sub

