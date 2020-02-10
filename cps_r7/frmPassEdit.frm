VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmPassEdit 
   Caption         =   "Password Access Control Screen"
   ClientHeight    =   4380
   ClientLeft      =   180
   ClientTop       =   975
   ClientWidth     =   9180
   Icon            =   "frmPassEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4380
   ScaleWidth      =   9180
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Exit"
      DisabledPicture =   "frmPassEdit.frx":57E2
      DownPicture     =   "frmPassEdit.frx":6424
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
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPassEdit.frx":7066
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   6495
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmPassEdit.frx":7CA8
      Height          =   3255
      Left            =   60
      Negotiate       =   -1  'True
      OleObjectBlob   =   "frmPassEdit.frx":7CBC
      TabIndex        =   0
      Top             =   60
      Width           =   9075
   End
End
Attribute VB_Name = "frmPassEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' no error modules ''''' Form PASSEDIT.frm
Option Explicit
Private daodb36 As DAO.Database
Private rS As DAO.Recordset
Dim sPath As String

Private Sub cmdReturn_Click()
    Unload frmPassEdit
    Set frmPassEdit = Nothing
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    Dim iresponse As Integer
    iresponse = ErrorHandler(DataErr)
    Select Case iresponse
      Case vbAbort       ' Exit if abort
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmPassEdit = Nothing     'current form
    End If
End Sub

Private Sub Form_Load()
    Dim sPath As String
    sPath = FILEPATH_sysdbf & DATAUSER
    Set daodb36 = DBEngine.OpenDatabase(sPath)
    Set rS = daodb36.OpenRecordset("password")
    Set frmPassEdit.Data1.Recordset = rS
End Sub


