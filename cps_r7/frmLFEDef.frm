VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmLFEDef 
   BackColor       =   &H80000005&
   Caption         =   "Laminar Flow Element Definition Screen"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLFE_D 
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Text            =   "Text6"
      ToolTipText     =   "Enter the x3 coefficient (D) here"
      Top             =   3900
      Width           =   1935
   End
   Begin VB.TextBox txtLFE_C 
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text5"
      ToolTipText     =   "Enter the x2 coefficient (C) here"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtLFE_B 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text4"
      ToolTipText     =   "Enter the x coefficient (B) here"
      Top             =   3300
      Width           =   1935
   End
   Begin VB.TextBox txtLFE_A 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "Text3"
      ToolTipText     =   "Enter the constant coefficient (A) here"
      Top             =   3000
      Width           =   1935
   End
   Begin Threed.SSCommand cmdSaveAs 
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      ToolTipText     =   "Save the current LFE definition with a different filename"
      Top             =   4680
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
   Begin Threed.SSCommand cmdSave 
      Height          =   735
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Save the current LFE definition"
      Top             =   4680
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
   Begin VB.TextBox txtFileName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtComments 
      Height          =   975
      Left            =   3120
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmLFEDef.frx":0000
      ToolTipText     =   "Enter your coments here"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtSerialNumber 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Enter the serial number here"
      Top             =   660
      Width           =   2655
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   735
      Left            =   7320
      TabIndex        =   10
      ToolTipText     =   "Return to the calibration form"
      Top             =   4680
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
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Proceed with the current LFE definition"
      Top             =   4680
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "Proceed"
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "'D' term"
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
      TabIndex        =   20
      Top             =   3900
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "'C' term"
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
      TabIndex        =   19
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "'B' term"
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
      TabIndex        =   18
      Top             =   3300
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "'A' term"
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
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LFE Coefficients:"
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
      Left            =   720
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "where x=Differential Pressure in Inches of W.C."
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
      Left            =   2160
      TabIndex        =   15
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LFE Flow (CFM) = A + Bx + Cx2 + Dx3"
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
      Left            =   1560
      TabIndex        =   14
      Top             =   2220
      Width           =   3375
   End
   Begin VB.Label lblCommentsTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
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
      Left            =   2040
      TabIndex        =   13
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label lblSerialNumTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LFE Serial Number:"
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
      Left            =   1200
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblFileNameTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LFE / Flow Standard Filename:"
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
      TabIndex        =   11
      Top             =   300
      Width           =   2775
   End
End
Attribute VB_Name = "frmLFEDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strPath As String
Private strFile As String

Private strSerialNum As String
Private strComments As String
Private LFE_A As Single
Private LFE_B As Single
Private LFE_C As Single
Private LFE_D As Single
Private blnSavedAs As Boolean

Private Sub cmdCancel_Click()
    ' Close the LFE definition screen
    Unload Me
    frmMassFlowCal.Enabled = True
    frmMassFlowCal.SetFocus
End Sub

Public Sub ShowData()
    ' Updates the screen with stored data
    txtFileName.text = PathFile
    txtSerialNumber.text = strSerialNum
    txtComments.text = strComments
    txtLFE_A.text = LFE_A
    txtLFE_B.text = LFE_B
    txtLFE_C.text = LFE_C
    txtLFE_D.text = LFE_D
End Sub

Private Sub cmdOpen_Click()
    ' Re-enables the calibration form,
    ' transfers LFE data to the form,
    ' and closes the LFE definition form
    Dim strFileName As String
    
    strFileName = PathFile
    frmMassFlowCal.Enabled = True
    
    frmMassFlowCal.SetLFE_A LFE_A
    frmMassFlowCal.SetLFE_B LFE_B
    frmMassFlowCal.SetLFE_C LFE_C
    frmMassFlowCal.SetLFE_D LFE_D
    frmMassFlowCal.SetLFE_SerialNum strSerialNum
    frmMassFlowCal.SetLFE_FileName strFileName
    frmMassFlowCal.UpdateLFELabel
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Save
End Sub

Private Sub cmdSaveAs_Click()
    ' Display the file dialog box
    Me.Enabled = False
    frmOpenLFE.Show
    frmOpenLFE.cmdNew.Enabled = False
    frmOpenLFE.cmdOpen.Enabled = False
End Sub

Private Sub Form_Load()
    Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Clear()
    strSerialNum = ""
    LFE_A = 0
    LFE_B = 0
    LFE_C = 0
    LFE_D = 0
    strComments = ""
    blnSavedAs = False
    ShowData
End Sub

Public Sub Save()
    On Error GoTo localhandler
    Dim fs, f As Object
    Dim iFileNumber As Integer
    Dim strFileName As String
    txtFileName.BackColor = Entry_BackColor
    Set fs = CreateObject("Scripting.FileSystemObject")
    strFileName = PathFile
    
    ' If the file exists then
    If fs.FileExists(strFileName) Then fs.DeleteFile strFileName

    
    ' Open the calibration file
    iFileNumber = FreeFile
    Open strFileName For Output As iFileNumber
    
    Write #iFileNumber, strSerialNum
    Write #iFileNumber, strComments
    Write #iFileNumber, LFE_A
    Write #iFileNumber, LFE_B
    Write #iFileNumber, LFE_C
    Write #iFileNumber, LFE_D
    
    Close #iFileNumber

    Exit Sub
localhandler:
    If err.Number = 53 Then
        txtFileName.Enabled = True

        txtFileName.BackColor = vbRed
        Delay_Box "Unusable filename", MSGDELAY, msgSHOW

    End If
    Resume Next
End Sub

Public Sub Read()
    Dim iFileNumber As Integer
    Dim strFileName As String
    Dim fs, f As Object
    
    strFileName = PathFile
    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(strFileName) = False Then
        ' Delay_Box "no calibration file", MSGDELAY, msgSHOW
        Exit Sub
    End If
    
    ' Open the calibration file
    iFileNumber = FreeFile
    Open strFileName For Input As iFileNumber
    Input #iFileNumber, strSerialNum
    Input #iFileNumber, strComments
    Input #iFileNumber, LFE_A
    Input #iFileNumber, LFE_B
    Input #iFileNumber, LFE_C
    Input #iFileNumber, LFE_D
    
    ' Close the file
    Close #iFileNumber
End Sub

Private Sub txtComments_Validate(Cancel As Boolean)
    strComments = txtComments.text
End Sub

Private Sub txtLFE_A_Validate(Cancel As Boolean)
    If IsNumeric(txtLFE_A.text) Then
        LFE_A = CSng(txtLFE_A.text)
    End If
    txtLFE_A.text = Format(LFE_A, "00.00#######")
End Sub

Private Sub txtLFE_B_Validate(Cancel As Boolean)
    If IsNumeric(txtLFE_B.text) Then
        LFE_B = CSng(txtLFE_B.text)
    End If
    txtLFE_B.text = Format(LFE_B, "00.00#######")
End Sub

Private Sub txtLFE_C_Validate(Cancel As Boolean)
    If IsNumeric(txtLFE_C.text) Then
        LFE_C = CSng(txtLFE_C.text)
    End If
    txtLFE_C.text = Format(LFE_C, "00.00#######")
End Sub

Private Sub txtLFE_D_Validate(Cancel As Boolean)
    If IsNumeric(txtLFE_D.text) Then
        LFE_D = CSng(txtLFE_D.text)
    End If
    txtLFE_D.text = Format(LFE_D, "00.00#######")
End Sub

Private Sub txtSerialNumber_Validate(Cancel As Boolean)
    strSerialNum = txtSerialNumber.text
End Sub

Public Property Get PathFile() As Variant
    ' Returns the whole filename
    PathFile = strPath & strFile
End Property

Public Property Get Path() As Variant
    ' Returns the path section of the filename
    Path = strPath
End Property

Public Property Get File() As Variant
    ' Returns the file section of the filename
    File = strFile
End Property

Public Property Let File(ByVal vFile As Variant)
    ' Sets the file section of the filename
    strFile = vFile
End Property

Public Property Let Path(ByVal vPath As Variant)
    ' Sets the path section of the filename
    If Right(Trim(vPath), 1) = "\" Then
        strPath = vPath
    Else
        strPath = vPath & "\"
    End If
End Property
