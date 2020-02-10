VERSION 5.00
Begin VB.Form frmHelpForm 
   BackColor       =   &H80000005&
   Caption         =   "Mistuser Help"
   ClientHeight    =   5940
   ClientLeft      =   915
   ClientTop       =   1230
   ClientWidth     =   7095
   Icon            =   "Helpform.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelpPrint 
      Caption         =   "&Print"
      Height          =   372
      Left            =   3960
      TabIndex        =   2
      Top             =   5520
      Width           =   2052
   End
   Begin VB.CommandButton cmdHelpClose 
      Caption         =   "&Close"
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   5520
      Width           =   2052
   End
   Begin VB.TextBox HelpText 
      Height          =   5295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmHelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'no error mod '''''''''''''''''Form HELPFORM.frm '''''''''''''''''''''
Option Explicit
Dim bFileOKFlag As Boolean
Private Sub cmdHelpClose_Click()

    Unload frmHelpForm
    
End Sub
Private Sub cmdHelpPrint_Click()

    'print the file to the default printer.
    Printer.Print HelpText.text
    Printer.NewPage
    Printer.EndDoc

End Sub
Private Sub Form_Activate()

    If bFileOKFlag = False Then
        Unload Me
   End If
   
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

HotKeyCheck KeyCode, Shift  ' undo rest to display key coads

End Sub

Private Sub Form_Load()

    ' declare local variables
    Dim DataString$, TempString$
    Dim FileNum%

    ' just in case the file doesn't exist
    On Error GoTo ErrHandler
    bFileOKFlag = True
    
    'initialize variables
    FileNum = 1

    ' open the file
    Open ViewFile For Input As FileNum

    ' read the file into Datastring
    Do While Not (EOF(FileNum))
        TempString$ = Input$(1, FileNum)
        DataString$ = DataString$ + TempString$
    Loop
    
    ' fill the edit box
    HelpText.text = DataString$

    ' close the file
    Close FileNum
    Exit Sub
    
ErrHandler:
    MsgBox err.Description, vbOKOnly + vbInformation, "Error!"
    bFileOKFlag = False

End Sub
