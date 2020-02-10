VERSION 5.00
Begin VB.Form frmOperComment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Operator Comment"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmOperComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmComment 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
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
         Left            =   5760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOperComment.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add the comment to the Job Log"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1245
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   3
         Text            =   "****************"
         Top             =   398
         Width           =   3960
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "message"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   1800
         Width           =   6465
      End
      Begin VB.Label lblComment 
         BackStyle       =   0  'Transparent
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOperComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
Public WhichStn As Integer
Public WhichShift As Integer

Private Sub cmdAccept_Click()
    ' hose origin
    If (Len(Trim(txtComment.text)) > 1) Then
        ' record comment
        If ((WhichStn > 0) And (WhichShift > 0)) Then
            Write_JLog WhichStn, WhichShift, Trim(txtComment.text)
            lblMessage.ForeColor = Message_ForeColor
            lblMessage.Caption = "Comment has been recorded"
            ' short delay
            DelayBySeconds 0.85
            ' unload this screen
            Unload Me
        Else
            If (WhichStn <= 0) And (WhichShift <= 0) Then
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = "Invalid Station and Shift Number"
            ElseIf (WhichStn <= 0) Then
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = "Invalid Station Number"
            ElseIf (WhichShift <= 0) Then
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = "Invalid Shift Number"
            End If
        End If
    Else
        lblMessage.ForeColor = Warning_ForeColor
        lblMessage.Caption = "Minimum of two characters"
        txtComment.BackColor = EntryInvalid_BackColor
    End If
End Sub

Private Sub Form_Load()
    lblMessage.Caption = " "
    txtComment.text = " "
    txtComment.BackColor = WhiteSmoke
    txtComment.TabIndex = 0
End Sub

Private Sub txtComment_Change()
    lblMessage.Caption = " "
    txtComment.BackColor = WhiteSmoke
End Sub


