VERSION 5.00
Begin VB.Form frmCylinder 
   Caption         =   "Cylinder OK"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmCylinder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNO 
      Caption         =   "No"
      DisabledPicture =   "frmCylinder.frx":058A
      DownPicture     =   "frmCylinder.frx":11CC
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
      Left            =   3450
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCylinder.frx":1E0E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdYES 
      Caption         =   "Yes"
      DisabledPicture =   "frmCylinder.frx":2A50
      DownPicture     =   "frmCylinder.frx":3692
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
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCylinder.frx":42D4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   390
      TabIndex        =   1
      Text            =   "--- Please Confirm ---"
      Top             =   390
      Width           =   3915
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   390
      TabIndex        =   0
      Text            =   "Additional Butane Cylinder?"
      Top             =   825
      Width           =   3915
   End
End
Attribute VB_Name = "frmCylinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 57
Option Explicit


Private Sub cmdNO_Click()
   Unload Me
End Sub

Private Sub cmdYES_Click()
Dim errors As Integer

errors = 0      ' 0 = no errors - 1 = Low errors - 2 = High errors
SetErrModule 57, 1
If UseLocalErrorHandler Then On Error GoTo localhandler
' initialize for first time ever used *******
If FrmButaneVolume.txtButaneCylWeight = Empty Then FrmButaneVolume.txtButaneCylWeight = 15
If FrmButaneVolume.txtButaneCylWeight < 0 Then
   FrmButaneVolume.txtButaneCylWeight.BackColor = EntryInvalid_BackColor
   errors = 1
End If
If FrmButaneVolume.txtButaneCylWeight > 999 Then
   FrmButaneVolume.txtButaneCylWeight.BackColor = EntryInvalid_BackColor
   errors = 2
End If
' initialize for first time through
If FrmButaneVolume.txtWarningSetPoint = Empty Then FrmButaneVolume.txtWarningSetPoint = 1
If FrmButaneVolume.txtWarningSetPoint < 0 Then
   FrmButaneVolume.txtWarningSetPoint.BackColor = EntryInvalid_BackColor
   errors = 1
End If
If FrmButaneVolume.txtWarningSetPoint > 100 Then
   FrmButaneVolume.txtWarningSetPoint.BackColor = EntryInvalid_BackColor
   errors = 2
End If

If errors = 0 Then
  ButaneSupply.WarningSetPoint = FrmButaneVolume.txtWarningSetPoint
  ButaneSupply.CylinderWeight = FrmButaneVolume.txtButaneCylWeight
  ButaneSupply.CurrentOnHand = FrmButaneVolume.txtFullCylinder  ' New Scale values now
  FrmButaneVolume.txtActualButane = FrmButaneVolume.txtFullCylinder
  FrmButaneVolume.txtDate = Now
  ButaneSupply.Date = FrmButaneVolume.txtDate
  FrmButaneVolume.txtPercentLeft = FrmButaneVolume.txtActualButane / FrmButaneVolume.txtFullCylinder * 100
  FrmButaneVolume.Update_Cylinders
  Write_ELog ("Butane Cylinder Changed " & ButaneSupply.Date)

  Save_ButaneSupply
Else
  If errors = 1 Then
    FrmButaneVolume.lblMsg.Caption = "Number too small...See tool tips"
  End If
  If errors = 2 Then
    FrmButaneVolume.lblMsg.Caption = "Number too large...See tool tips"
  End If
End If
Unload Me
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

  Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

