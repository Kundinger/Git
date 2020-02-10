VERSION 5.00
Begin VB.Form FrmButaneVolume 
   Caption         =   "Butane Supply"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   525
   ClientWidth     =   7935
   Icon            =   "FrmButaneVolume.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmCurrent 
      Caption         =   "Current Values"
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
      Height          =   1995
      Left            =   330
      TabIndex        =   13
      Top             =   240
      Width           =   7305
      Begin VB.TextBox txtActualButane 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "xxx,xxx"
         ToolTipText     =   "Leters Remaining in  Cylinder"
         Top             =   990
         Width           =   2475
      End
      Begin VB.TextBox txtPercentLeft 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4980
         MaxLength       =   5
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "000.0"
         ToolTipText     =   "Percent remaining in the Cylinder"
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Approximate Butane Remaining"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   16
         Top             =   450
         Width           =   6315
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   " Liters"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2820
         TabIndex        =   15
         Top             =   1050
         Width           =   1185
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   14
         Top             =   1050
         Width           =   465
      End
   End
   Begin VB.Frame frmBtnSetPnt 
      Caption         =   "Butane Settings"
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
      Height          =   2875
      Left            =   330
      TabIndex        =   6
      Top             =   2460
      Width           =   7305
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Quit"
         DisabledPicture =   "FrmButaneVolume.frx":058A
         DownPicture     =   "FrmButaneVolume.frx":0C8C
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
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmButaneVolume.frx":138E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Settings"
         DisabledPicture =   "FrmButaneVolume.frx":1A90
         DownPicture     =   "FrmButaneVolume.frx":2192
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
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmButaneVolume.frx":2894
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Settings"
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdAddCylinder 
         Caption         =   "Add Cylinder"
         DisabledPicture =   "FrmButaneVolume.frx":2F96
         DownPicture     =   "FrmButaneVolume.frx":3698
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
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmButaneVolume.frx":3D9A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add New Cylinder"
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.TextBox txtWarningSetPoint 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3570
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "12.3"
         ToolTipText     =   "Display Low Butane Warning when calculated percent remaining is less than this percent"
         Top             =   1485
         Width           =   585
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "Date"
         ToolTipText     =   "Date Last Cylinder was changed"
         Top             =   915
         Width           =   2085
      End
      Begin VB.TextBox txtFullCylinder 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3480
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "xxx,xxx"
         ToolTipText     =   "Full Cylinder has this many Liters"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtButaneCylWeight 
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
         Left            =   2130
         TabIndex        =   1
         Text            =   "55"
         ToolTipText     =   "Cylinder Weight in lbs."
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         Caption         =   "msg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   615
         Left            =   2280
         TabIndex        =   17
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   1515
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cylinder Warning Alarm Set Point :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   440
         TabIndex        =   11
         Top             =   1500
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cylinder Last Saved On :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   440
         TabIndex        =   10
         Top             =   960
         Width           =   2355
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Liters of C4H10"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4860
         TabIndex        =   9
         Top             =   405
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "lbs.   ="
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2820
         TabIndex        =   8
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cylinder Weight :"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   440
         TabIndex        =   7
         Top             =   390
         Width           =   1545
      End
   End
End
Attribute VB_Name = "FrmButaneVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' * error module = 56
Option Explicit


Private Sub cmdReturn_Click()
    Unload Me
    Set FrmButaneVolume = Nothing
End Sub

Private Sub cmdSave_Click()
    ButaneSupply.WarningSetPoint = ValueFromText(txtWarningSetPoint.text)
    ButaneSupply.CylinderWeight = ValueFromText(txtButaneCylWeight.text)
    ButaneSupply.CurrentOnHand = ValueFromText(txtActualButane.text)
    ButaneSupply.Date = Trim(txtDate.text)
    Save_ButaneSupply
    Write_Elog "Butane Settings Saved by Operator."
End Sub

Private Sub cmdAddCylinder_Click()
    lblMsg.Caption = ""
    frmCylinder.Show
End Sub

Public Sub Update_Cylinders()
    txtWarningSetPoint.text = Format(ButaneSupply.WarningSetPoint, "##0.0##")
    txtButaneCylWeight.text = Format(ButaneSupply.CylinderWeight, "######0.00")
    txtActualButane.text = Format(ButaneSupply.CurrentOnHand, "######0.00")
    txtDate.text = ButaneSupply.Date
    'Figure out percent
    txtFullCylinder.text = Format(((ButaneSupply.CylinderWeight * 28.317) / 0.1501), "#####0.00")
    If ButaneSupply.CurrentOnHand > 0 Then
        txtPercentLeft.text = Format((ButaneSupply.CurrentOnHand / ValueFromText(txtFullCylinder.text) * 100), "######0.00")
    Else ' Set a default for first time start up.
        ButaneSupply.WarningSetPoint = 1               ' This is to initialise on power up
        ButaneSupply.CylinderWeight = 10               ' And prevent errors
        ButaneSupply.CurrentOnHand = 1
    End If
End Sub

Private Sub txtButaneCylWeight_Change()
    If IsNumeric(txtButaneCylWeight.text) Then
        txtFullCylinder.text = Format(((ValueFromText(txtButaneCylWeight.text) * 28.317) / 0.1501), "#####0.00")  ' Liters
        txtActualButane.text = txtFullCylinder.text
        txtButaneCylWeight.BackColor = Entry_BackColor
        lblMsg.Caption = ""
    Else
        txtButaneCylWeight.BackColor = EntryInvalid_BackColor
        lblMsg.Caption = "Number MUST be Numeric"
    End If
End Sub

Private Sub Form_Load()
    frmCurrent.ForeColor = Titles_ForeColor
    frmBtnSetPnt.ForeColor = Titles_ForeColor
    Update_Cylinders
End Sub

Private Sub txtWarningSetPoint_Change()
    If IsNumeric(txtWarningSetPoint) Then
        txtWarningSetPoint.BackColor = Entry_BackColor
        lblMsg.Caption = ""
    Else
        txtWarningSetPoint.BackColor = EntryInvalid_BackColor
        lblMsg.Caption = "Number MUST be Numeric"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
    HotKeyCheck KeyCode, shift  ' undo rest to display key coads
End Sub

