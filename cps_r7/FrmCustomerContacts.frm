VERSION 5.00
Begin VB.Form FrmCustomerContacts 
   BackColor       =   &H80000005&
   Caption         =   "External Alarm Contacts"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrUpdate 
      Interval        =   10000
      Left            =   7440
      Top             =   5040
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "CONTINUE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3620
      TabIndex        =   1
      Top             =   4800
      Width           =   2505
   End
   Begin VB.TextBox TxtDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3800
      Left            =   620
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "FrmCustomerContacts.frx":0000
      Top             =   620
      Width           =   8500
   End
End
Attribute VB_Name = "FrmCustomerContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private txt As String
Private Sub cmdContinue_Click()
    ' IS THE ALARM RESET
    If Com_DIO(icExtAlmContactSw).Value Then
        ' External Contacts Still Set"
        txt = DESC_EXT_CONTACTS + " STILL Set"
        Delay_Box txt, MSGDELAY, msgSHOW
    Else
        ' They fixed it
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_Load()
    txt = vbCrLf
    txt = txt + DESC_EXT_CONTACTS + " Activated" + vbCrLf
    txt = txt + vbCrLf
    txt = txt + "To clear the alarm:" + vbCrLf
    txt = txt + "First resolve the alarm condition," + vbCrLf
    txt = txt + "Then push the CONTINUE button below." + vbCrLf
    txt = txt + vbCrLf
    txt = txt + "*** REMEMBER ***" + vbCrLf
    txt = txt + "Each station in progress" + vbCrLf
    txt = txt + "must also be continued individually." + vbCrLf
    TxtDisplay.text = txt
End Sub

Private Sub tmrUpdate_Timer()
    ' IS THE ALARM RESET
    If Com_DIO(icExtAlmContactSw).Value Then
        ' External Contacts Still Set"
        FrmCustomerContacts.SetFocus
    Else
        ' They fixed it
        Unload Me
    End If
End Sub
