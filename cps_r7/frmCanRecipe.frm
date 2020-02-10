VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmCanRecipe 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Canister Properties"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "frmCanRecipe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbControlBtns 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7185
      TabIndex        =   15
      Top             =   0
      Width           =   7185
      Begin VB.CommandButton cmdSetDefault 
         Caption         =   "Create New"
         DisabledPicture =   "frmCanRecipe.frx":57E2
         DownPicture     =   "frmCanRecipe.frx":5EE4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   912
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Set WC & Description to default values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close"
         DisabledPicture =   "frmCanRecipe.frx":6CE8
         DownPicture     =   "frmCanRecipe.frx":73EA
         BeginProperty Font 
            Name            =   "Arial"
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
         Picture         =   "frmCanRecipe.frx":7AEC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         DisabledPicture =   "frmCanRecipe.frx":81EE
         DownPicture     =   "frmCanRecipe.frx":88F0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":8FF2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Paste Canister Values from the clipboard"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         DisabledPicture =   "frmCanRecipe.frx":96F4
         DownPicture     =   "frmCanRecipe.frx":9DF6
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":A4F8
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Copy Canister Values to the clipboard"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmCanRecipe.frx":ABFA
         DownPicture     =   "frmCanRecipe.frx":B2FC
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1824
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":B9FE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Canister Values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         CausesValidation=   0   'False
         DisabledPicture =   "frmCanRecipe.frx":C100
         DownPicture     =   "frmCanRecipe.frx":CD42
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3648
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":D984
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print a Listing of all Canisters"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         DisabledPicture =   "frmCanRecipe.frx":E5C6
         DownPicture     =   "frmCanRecipe.frx":ECC8
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":F3CA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Open Master Canister List"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore"
         DisabledPicture =   "frmCanRecipe.frx":FACC
         DownPicture     =   "frmCanRecipe.frx":101CE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2736
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":108D0
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Reload Station Recipe Values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
   End
   Begin VB.PictureBox pbMasterBtns 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7185
      TabIndex        =   8
      Top             =   960
      Width           =   7185
      Begin VB.CommandButton cmdPgUp 
         Caption         =   "Pg Next"
         DisabledPicture =   "frmCanRecipe.frx":10FD2
         DownPicture     =   "frmCanRecipe.frx":116D4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5135
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":11DD6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Next"
         DisabledPicture =   "frmCanRecipe.frx":124D8
         DownPicture     =   "frmCanRecipe.frx":12BDA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4295
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":132DC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPgDn 
         Caption         =   "Pg Prev"
         DisabledPicture =   "frmCanRecipe.frx":139DE
         DownPicture     =   "frmCanRecipe.frx":140E0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":147E2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Prev"
         DisabledPicture =   "frmCanRecipe.frx":14EE4
         DownPicture     =   "frmCanRecipe.frx":155E6
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCanRecipe.frx":15CE8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin Threed.SSPanel txtDispCan 
         Height          =   840
         Left            =   3240
         TabIndex        =   9
         ToolTipText     =   "Click for list of Defined Canisters"
         Top             =   0
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "01"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   24.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   3
      End
      Begin VB.Label lblStnDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "station shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   4200
         TabIndex        =   24
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.TextBox txtWorkingCapacity 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1710
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "0.0"
      ToolTipText     =   "Butane Working Capacity; enter 1 to 1000 grams"
      Top             =   3060
      Width           =   1065
   End
   Begin VB.TextBox txtWCVolume 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   1710
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "0.0"
      ToolTipText     =   "Working Volume; enter 0.01 to 10 Liters"
      Top             =   2460
      Width           =   1065
   End
   Begin VB.TextBox txtCanDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1710
      MaxLength       =   20
      TabIndex        =   2
      ToolTipText     =   "Enter 20 Character Description"
      Top             =   1830
      Width           =   5055
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   3720
      Width           =   7215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Capacity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   3105
      Width           =   1605
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "grams"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   6
      Top             =   3105
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "liters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   4
      Top             =   2505
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   2505
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   1875
      Width           =   1605
   End
End
Attribute VB_Name = "frmCanRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' module 190 ******* Canister Recipe Screen
Option Explicit
Private ScreenBkgdColor As Long
Private ScreenDescription As String
Private ScreenDispFlag As Boolean
Private StnShftDescription As String
Private CanRcpMode As Integer            ' 0=master; 1=station
Private DispCan As Integer               ' Current Master Canister index
Private Chgs As Boolean
Private CalcCanisterWC As Boolean
Private DspCanister As CanisterRecipe
Private MemCanister As CanisterRecipe
Private inct As Integer
Private dbDbase As Database
Private rsRecord  As Recordset
Private Criteria As String

Public Sub ChgCanRcpMode(ByVal NewMode As Integer)
    ' 0=master; 1=station
    CanRcpMode = IIf((NewMode = 0 Or NewMode = 1), NewMode, 0)
    Select Case CanRcpMode
        Case MASTERMODE
            ' station/shift description
            lblStnDesc.Visible = False
            ' screen description
            ScreenDescription = "Master Canister Properties"
            ' screen background color
            ScreenBkgdColor = MasterMode_BackColor
            ' show Recipe # & Arrows
            ScreenDispFlag = True
        Case STATIONMODE
            ' station/shift description
            StnShftDescription = "Station #" & Format(DispStn, "#0")
            If NR_SHIFT > 1 Then StnShftDescription = StnShftDescription & "  Shift #" & Format(DispShift, "#0")
            StnShftDescription = StnShftDescription & "  Canister Properties"
            lblStnDesc.Visible = True
            lblStnDesc.Left = cmdPgDn.Left
            lblStnDesc.ForeColor = TitlesData_Forecolor
            lblStnDesc.Caption = StnShftDescription
            ' screen description
            ScreenDescription = StnShftDescription
            ' screen background color
            ScreenBkgdColor = StationMode_BackColor
            ' hide Recipe # & Arrows
            ScreenDispFlag = False
    End Select
    ' screen description
    frmCanRecipe.Caption = ScreenDescription
    ' set screen background colors
    frmCanRecipe.BackColor = ScreenBkgdColor
    pbMasterBtns.BackColor = ScreenBkgdColor
    pbControlBtns.BackColor = ScreenBkgdColor
    txtDispCan.BackColor = ScreenBkgdColor
    lblMsg.BackColor = ScreenBkgdColor
    ' show Recipe # & Arrows ??
    cmdDown.Visible = ScreenDispFlag
    cmdUp.Visible = ScreenDispFlag
    cmdPgDn.Visible = ScreenDispFlag
    cmdPgUp.Visible = ScreenDispFlag
    txtDispCan.Visible = ScreenDispFlag
End Sub

Public Sub LoadNewCan(ByVal NewCan As Integer)
    DispCan = NewCan
    GetCanRcp MASTERMODE, DispCan, 0
    DspCanToScreen
End Sub

Private Sub DspCanToScreen()
    txtDispCan.Caption = Format(DspCanister.Number, "#00")
    txtCanDescription.text = DspCanister.Description
    txtWorkingCapacity.text = Format(DspCanister.WorkingCapacity, "####0.000")    ' 1 to 1000
    txtWCVolume.text = Format(DspCanister.WorkingVolume, "##0.000")                  ' 0.010 to 10
End Sub

Private Sub DspCanToMemCan()
    MemCanister = DspCanister
End Sub

Private Sub MemCanToDspCan()
    DspCanister = MemCanister
End Sub

Private Sub ScreenToDspCan()
    DspCanister.Number = CInt(txtDispCan.Caption)
    DspCanister.Description = txtCanDescription
    DspCanister.WorkingCapacity = CSng(txtWorkingCapacity.text)
    DspCanister.WorkingVolume = CSng(txtWCVolume.text)
End Sub

Private Sub ExitScreen()
    ' close canister / recipe database
    dbDbase.Close
    ' unload form
    frmCanRecipe.Visible = False
    Unload Me
End Sub

Public Sub CanRcpDisplay_ByNum()
    GetCanRcp MASTERMODE, DispCan, 0
    DspCanToScreen
End Sub

Public Sub CanRcpDisplay_ByStnShift()
    GetCanRcp STATIONMODE, DispStn, DispShift
    DspCanToScreen
End Sub

Private Sub GetCanRcp(ByVal MstStnMode As Integer, ByVal index1 As Integer, ByVal index2 As Integer)
    Select Case MstStnMode
        Case MASTERMODE
            ' Read Master Canister Information Record
            Criteria = "SELECT * FROM [MasterCanister] WHERE [Number] = " & index1 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                DspCanister.Number = index1
                DspCanister.Description = "undefined"
                DspCanister.WorkingCapacity = 0
                DspCanister.WorkingVolume = "0"
                DspCanister.Validated = False
            Else
                DspCanister.Number = rsRecord("Number")
                DspCanister.Description = rsRecord("Description")
                DspCanister.WorkingCapacity = rsRecord("WorkingCapacity")
                DspCanister.WorkingVolume = rsRecord("WCVolume")
                DspCanister.Validated = False
            End If
        Case STATIONMODE
            ' Read Station Canister Information Record
            Criteria = "SELECT * FROM [StationCanister] WHERE [Station] = " & index1 & "  and [Shift] = " & index2 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                DspCanister.Number = 0
                DspCanister.Description = "undefined"
                DspCanister.WorkingCapacity = 0
                DspCanister.WorkingVolume = "0"
                DspCanister.Validated = False
            Else
                DspCanister.Number = rsRecord("Number")
                DspCanister.Description = rsRecord("Description")
                DspCanister.WorkingCapacity = rsRecord("WorkingCapacity")
                DspCanister.WorkingVolume = rsRecord("WCVolume")
                DspCanister.Validated = True
            End If
    End Select
    If DspCanister.WorkingCapacity < 0.01 And DspCanister.WorkingVolume < 0.005 Then
        DspCanister.Description = "undefined"
    End If
    rsRecord.Close
    Chgs = False
End Sub

Private Sub SaveMasterCanRcp(ByVal index1 As Integer)
        ' Read Master Canister Information Record
        Criteria = "SELECT * FROM [MasterCanister] WHERE [Number] = " & index1 & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        If rsRecord.BOF Then
            rsRecord.AddNew
            rsRecord("Number") = index1
        Else
          rsRecord.MoveFirst
          rsRecord.Edit
        End If
           
        ' Update Master Canister Information Record
        rsRecord("Description") = DspCanister.Description
        rsRecord("WorkingCapacity") = DspCanister.WorkingCapacity
        rsRecord("WCVolume") = DspCanister.WorkingVolume
        rsRecord.Update
        rsRecord.Close
End Sub

Public Sub InitCanRcp()
    Select Case CanRcpMode
        Case MASTERMODE
            ' master
            If DispCan < 1 Or DispCan > NR_CAN Then
               DispCan = 1
            End If
            GetCanRcp MASTERMODE, DispCan, 0
            cmdSetDefault.Visible = False
            cmdRestore.Visible = False
            cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
        Case STATIONMODE
            ' station
            If StationCanister(DispStn, DispShift).Number < 0 _
             Or StationCanister(DispStn, DispShift).Number > NR_CAN Then
               DispCan = 0
            Else
               DispCan = StationCanister(DispStn, DispShift).Number
            End If
            GetCanRcp STATIONMODE, DispStn, DispShift
            If StationControl(DispStn, DispShift).Mode <> VBIDLE Then
                cmdSetDefault.Visible = False
                cmdRestore.Visible = False
                cmdSave.Visible = False
            Else
                cmdSetDefault.Visible = True
                cmdRestore.Visible = True
                cmdSave.Visible = True
            End If
            cmdPrint.Visible = False
    End Select
    DspCanToScreen
    Chgs = False
End Sub

Private Function ValidCanister() As Boolean
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 50, 111
Dim errors As Integer

    errors = 0      ' 0 = no errors - 1 = Low errors - 2 = High errors
    
    txtCanDescription.text = Mid$(Trim$(txtCanDescription.text), 1, 40)
    If Len(txtCanDescription.text) < 1 Then txtCanDescription.text = "this canister"
    
    ' working capacity
    If Not IsNumeric(txtWorkingCapacity.text) Then
        txtWorkingCapacity.BackColor = EntryInvalid_BackColor
        errors = 3
    ElseIf CSng(txtWorkingCapacity.text) < 0 Then
        txtWorkingCapacity.BackColor = EntryInvalid_BackColor
        errors = 1
    ElseIf CSng(txtWorkingCapacity.text) = 0 Then
        If (Not ((StationRecipe(DispStn, DispShift).EndMethod = ENDWEIGHTCHG) And (StationRecipe(DispStn, DispShift).UpdateCanWc))) Then
            txtWorkingCapacity.BackColor = EntryInvalid_BackColor
            errors = 3
        End If
    ElseIf CSng(txtWorkingCapacity.text) > 1000 Then
        txtWorkingCapacity.BackColor = EntryInvalid_BackColor
        errors = 2
    End If
    
    ' volume
    If Not IsNumeric(txtWCVolume.text) Then
        txtWCVolume.BackColor = EntryInvalid_BackColor
        errors = 3
    ElseIf CSng(txtWCVolume.text) < 0.005 Then
        txtWCVolume.BackColor = EntryInvalid_BackColor
        errors = 1
    ElseIf CSng(txtWCVolume.text) > 10 Then
        txtWCVolume.BackColor = EntryInvalid_BackColor
        errors = 2
    End If
    
    ' errors ??
    Select Case errors
        Case 0
            ' ok
            ValidCanister = True
        Case 1
            ' low errors
            ValidCanister = False
            lblMsg.FontSize = 12
            lblMsg.Caption = "Number too small...See tool tips"
        Case 2
            ' high errors
            ValidCanister = False
            lblMsg.FontSize = 12
            lblMsg.Caption = "Number too large...See tool tips"
        Case 3
            ' invalid errors
            ValidCanister = False
            lblMsg.FontSize = 12
            lblMsg.Caption = "Not Valid...See tool tips"
    End Select
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Private Sub cmdCopy_Click()
    ScreenToDspCan
    DspCanToMemCan
End Sub

Private Sub cmdPaste_Click()
    MemCanToDspCan
    DspCanister.Number = DispCan
    DspCanToScreen
    Chgs = True
End Sub

Private Sub cmdPrint_Click()
    lblMsg.Caption = " "
    Print_All
'    Delay_Box "Canister Listing Released to the Printer", MSGDELAY, msgSHOW
    lblMsg.Font.Size = 9.5
    lblMsg.ForeColor = DKPURPLE
    lblMsg.Caption = "Canister Listing sent to" & vbCrLf & PRINTERNAME
End Sub

Private Sub cmdOpen_Click()
    frmSearchCan.SetInitialRow (DispCan)
    frmSearchCan.ChgCanRcpMode CanRcpMode
    frmSearchCan.Show
End Sub

Private Sub cmdRestore_Click()
    CanRcpDisplay_ByStnShift
End Sub

Private Sub cmdDown_Click()
    DispCan = IIf(DispCan < 2, NR_CAN, DispCan - 1)
    CanRcpDisplay_ByNum
End Sub

Private Sub cmdSetDefault_Click()
    txtCanDescription.text = "test canister"
    txtWCVolume.text = "0.0"
    txtWorkingCapacity.text = "0.0"
    lblMsg.Caption = "Enter Canister Volume"
    txtWCVolume.SetFocus
End Sub

Private Sub cmdUp_Click()
    DispCan = IIf(DispCan > NR_CAN - 1, 1, DispCan + 1)
    CanRcpDisplay_ByNum
End Sub

Private Sub cmdPgDn_Click()
    DispCan = IIf(DispCan < 12, NR_CAN, DispCan - 10)
    CanRcpDisplay_ByNum
End Sub

Private Sub cmdPgUp_Click()
    DispCan = IIf(DispCan > NR_CAN - 10, 1, DispCan + 10)
    CanRcpDisplay_ByNum
End Sub

Private Sub cmdCancel_Click()
    ExitScreen
End Sub

Private Sub cmdSave_Click()
    SaveCanister
End Sub

Private Sub SaveCanister()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 50, 1
    lblMsg.FontSize = 12
    Select Case CanRcpMode
        Case MASTERMODE
            ' master
            If CheckPass("P", True) Then
                If ValidCanister Then
                    ScreenToDspCan
                    DspCanister.Validated = True
                    ' Save Master Canister Information
                    SaveMasterCanRcp CInt(DspCanister.Number)
                    ' Save Remote Master Canister Information
                    If USINGREMCANLOAD Then
                        ' open master canister / recipe database
                        Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
                        ' open remote database
                        OpenConnToRemoteDb
                        ' update Remote Master Canister Information
                        UpdateRemoteCanisters
                        ' close remote database
                        CloseConnToRemoteDb
                    End If
                    lblMsg.Caption = "Master Canister Information #" + Format(DspCanister.Number, "###0") & vbCrLf & " saved to database"
                    Chgs = False
                End If
            End If
    
        Case STATIONMODE
            ' station
            If StationControl(DispStn, DispShift).Mode = VBIDLE Then
                If ValidCanister Then
                    ScreenToDspCan
                    DspCanister.Validated = True
                    StationCanister(DispStn, DispShift) = DspCanister
                    StationCanister(DispStn, DispShift).Number = IIf(Chgs, CInt(0), DispCan)
                    ' clear Canister LeakCheck Status
                    StationControl(DispStn, DispShift).LeakCheckStatus = NORESULT
                    StationControl(DispStn, DispShift).LcStatusDescription = " "
                    ' save station canister recipes
                    Save_StationCanisters
                    Select Case NR_SHIFT
                        Case 1
                            lblMsg.Caption = "Canister Values saved to Station #" + Format(DispStn, "0")
                        Case 2
                            lblMsg.Caption = "Canister Values saved to Station #" + Format(DispStn, "0") + " / Shift #" + Format(DispShift, "0")
                    End Select
                    ' If Recipe has LoadByWC then reset LoadRate to EPA default of 15 grams/hour
                    If StationRecipe(DispStn, DispShift).Load_MethodSave = LOADBYWC Then
                        StationRecipe(DispStn, DispShift).Load_Rate = CSng(15)
                        StationRecipe(DispStn, DispShift).Load_RateSave = CSng(15)
                        Save_StationRecipes
                    End If
                    ' adjust Y-axis on stn XYGraph for new canister
                    frmStnDetail.Scale_Yaxis StationCanister(DispStn, DispShift).WorkingCapacity
                    ' clear REM Data for this Station/Shift
                    RemData_Clear StnRemoteTask(DispStn, DispShift)
'                    ExitScreen
                End If
            Else
               lblMsg.Caption = "Can Not Change values while station is running"
            End If
        
    End Select

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

Private Sub Form_Load()
    KeyPreview = True
    lblMsg.FontSize = 12
    ' set foreground colors
    lblMsg.ForeColor = Message_ForeColor
    txtDispCan.ForeColor = TitlesData_Forecolor
    txtCanDescription.ForeColor = TitlesData_Forecolor
    txtCanDescription.FontBold = True
    txtWCVolume.ForeColor = Data_ForeColor
    txtWorkingCapacity.ForeColor = Data_ForeColor
    ' show Restore Station Recipe ?
    cmdRestore.Visible = IIf(CanRcpMode = STATIONMODE, True, False)
    ' show print button
    cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
    ' show SetDefault button
    cmdSetDefault.Visible = IIf(CanRcpMode = STATIONMODE, True, False)
    ' open canister / recipe database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
End Sub

Private Sub txtCanDescription_Change()
    lblMsg.FontSize = 12
    lblMsg.Caption = " "
    Chgs = True
End Sub

Private Sub txtDispCan_Click()
    frmSearchCan.SetInitialRow (DispCan)
    frmSearchCan.Show
End Sub

Private Sub txtWCVolume_Change()
    txtWCVolume.BackColor = Entry_BackColor
    lblMsg.FontSize = 12
    lblMsg.Caption = " "
    Chgs = True
End Sub

Private Sub txtWorkingCapacity_Change()
    txtWorkingCapacity.BackColor = Entry_BackColor
    lblMsg.FontSize = 12
    lblMsg.Caption = " "
    Chgs = True
'    cmdSave.Enabled = IIf((ValueFromText(txtWorkingCapacity.text) >= CSng(0)), True, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        ExitScreen
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Print_All()
' print canisters
Dim yThisLine, yUnderline, numPages, inct As Integer
Dim xNum, xDesc, xWC, xVol As Integer
Dim widthData As Single
Dim strNum, strDesc, strWC, strVol As String
Dim oldFont As New StdFont
'object.TextWidth(string)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 190, 3
         
    ' set column positions
    xNum = 1700
    xDesc = 2480
    xWC = 8900
    xVol = 10480
    
    ' Save current printer font
    oldFont = Printer.Font
    Printer.Font = REPORTFONT
    
    ' TITLE, HEADER & COLUMN HEADINGS
    Print_Header
    
    ' number of pages
    numPages = (NR_CAN \ 50)
    If (NR_CAN Mod 50) > 0 Then numPages = numPages + 1
    For inct = 1 To NR_CAN
        GetCanRcp MASTERMODE, inct, 0
        yThisLine = Printer.CurrentY
        strNum = Format(inct, "##0")
        widthData = Printer.TextWidth(strNum)
        Printer.CurrentX = xNum - widthData
        Printer.Print strNum
        strDesc = Trim$(Mid$(DspCanister.Description, 1, 40))
        If Len(strDesc) > 0 And strDesc <> "undefined" Then
            strWC = Format(DspCanister.WorkingCapacity, "###0.0#")
            strVol = Format(DspCanister.WorkingVolume, "##0.00")
            Printer.CurrentY = yThisLine
            Printer.CurrentX = xDesc
            Printer.Print strDesc
            widthData = Printer.TextWidth(strWC)
            Printer.CurrentY = yThisLine
            Printer.CurrentX = xWC - widthData
            Printer.Print strWC
            widthData = Printer.TextWidth(strVol)
            Printer.CurrentY = yThisLine
            Printer.CurrentX = xVol - widthData
            Printer.Print strVol
        Else
            Printer.CurrentY = yThisLine
            Printer.CurrentX = xDesc
            Printer.Print "---   currently undefined   ---"
        End If
        ' more pages?
        If inct = NR_CAN Or (inct Mod 50) = 0 Then
            ' print footer
            Print_Footer numPages
            If inct <> NR_CAN Then
                ' new page
                Printer.NewPage
                ' TITLE, HEADER & COLUMN HEADINGS
                Print_Header
            End If
        End If
     Next inct
    
    Print_Footer numPages
    Printer.EndDoc
    Printer.Font = oldFont

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

Private Sub Print_Header()
' print canisters report header
Dim yThisLine, yUnderline As Integer
Dim xNum1, xDesc1, xWC1, xVol1 As Integer
Dim xNum2, xDesc2, xWC2, xVol2 As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 190, 300

    ' set column positions
    xNum1 = 1200
    xDesc1 = 2480
    xWC1 = 7800
    xVol1 = 9880
    xNum2 = xNum1
    xDesc2 = xDesc1
    xWC2 = 8300
    xVol2 = 9980
         
    ' HEADER & TITLE
    ' font
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    Printer.Font.Underline = False
    ' title
    Print_Center "CANISTER LISTING REPORT"
'    Print_Line ""
    ' header
    Print_Center "Canister Preconditioning System"
    Print_Center Trim$(SysConfig.Heading)
    Print_Center Trim$(SysConfig.Heading2)
    Print_Center (Format(Now, "d mmmm yyyy"))
    ' reset font
    Printer.Font.Size = 10
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    Printer.Font.Underline = False
    ' Print blank line(s)
    Print_Line ""
'    Print_Line ""
    
    'Print Header
    '"123456789^123456789^123456789^123456789^123456789^123456789^123456789^1234567890"
    yThisLine = Printer.CurrentY
    Printer.CurrentX = xNum1
    Printer.Print "Number"
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xDesc1
    Printer.Print "Description"
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xWC1
    Printer.Print "Working Capacity"
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xVol1
    Printer.Print "Volume"
    yThisLine = Printer.CurrentY
    yUnderline = Printer.CurrentY
    Printer.CurrentX = xWC2
    Printer.Print "(grams)"
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xVol2
    Printer.Print "(liters)"
    ' column headers
    Printer.Font.Underline = True
    Printer.CurrentY = yUnderline
    Printer.CurrentX = xNum1
    Printer.Print Space(14)
    Printer.CurrentY = yUnderline
    Printer.CurrentX = xDesc1
    Printer.Print Space(84)
    Printer.CurrentY = yUnderline
    Printer.CurrentX = xWC1
    Printer.Print Space(30)
    Printer.CurrentY = yUnderline
    Printer.CurrentX = xVol1
    Printer.Print Space(14)
    Printer.Font.Underline = False

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

Private Sub SaveRemoteMasterCanister(ByVal iCan As Integer)
'
'        Save Remote Master Canister Information Record
'
'        frmRemoteCan.Hide
        frmRemoteCan.Show
        ' Open existing Remote Master Canister Information Record (if any)
        frmRemoteCan.adoRemoteCanisters.RecordSource = "SELECT * FROM [MasterCanister] WHERE [MasterCanister].[Number] = " & iCan & " "
'        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        frmRemoteCan.adoRemoteCanisters.Refresh
        frmRemoteCan.dgRemoteCanisters.Refresh
        
        With frmRemoteCan.adoRemoteCanisters.Recordset
        
            If .BOF Then
                .AddNew
                .Fields("Number").Value = iCan
            Else
              .MoveLast
              .MoveFirst
            End If
               
            If .RecordCount = 1 Then
                ' Update Remote Master Canister Information Record
                .Fields("Description").Value = DspCanister.Description
                .Fields("WorkingCapacity").Value = DspCanister.WorkingCapacity
                .Fields("WCVolume").Value = DspCanister.WorkingVolume
                .Update
            Else
                Write_ELog "RemoteCan Update Failure - Multiple Records Returned for Can# " & Format(iCan, "#,##0")
            End If
            
        End With
        
        ' reset RemoteCanisters RecordSource
        frmRemoteCan.adoRemoteCanisters.RecordSource = "SELECT * FROM [MasterCanister] ORDER BY [MasterCanister].[Number] ASC"
        frmRemoteCan.adoRemoteCanisters.Refresh
        frmRemoteCan.dgRemoteCanisters.Refresh
                    

        Unload frmRemoteCan

End Sub

