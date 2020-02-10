VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisplayProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Properties"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   Icon            =   "frmDisplayProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetDefaults 
      Caption         =   " Default Colors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDisplayProperties.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Restore Default Screen Colors"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.ComboBox DisplayColors 
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
      ItemData        =   "frmDisplayProperties.frx":5CD4
      Left            =   600
      List            =   "frmDisplayProperties.frx":5CD6
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2280
      Width           =   3045
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      DisabledPicture =   "frmDisplayProperties.frx":5CD8
      DownPicture     =   "frmDisplayProperties.frx":63DA
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
      Left            =   7080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDisplayProperties.frx":6ADC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save Recipe"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      DisabledPicture =   "frmDisplayProperties.frx":71DE
      DownPicture     =   "frmDisplayProperties.frx":78E0
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
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDisplayProperties.frx":7FE2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save Recipe"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Close"
      DisabledPicture =   "frmDisplayProperties.frx":86E4
      DownPicture     =   "frmDisplayProperties.frx":8DE6
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
      Left            =   10200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDisplayProperties.frx":94E8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Quit"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.Timer tmrScreen 
      Interval        =   50
      Left            =   3840
      Top             =   6120
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   480
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "Blue"
      Top             =   4920
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   847
      _Version        =   393216
      LargeChange     =   16
      Max             =   255
      SelStart        =   128
      TickFrequency   =   16
      Value           =   128
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   480
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Red"
      Top             =   3960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   847
      _Version        =   393216
      LargeChange     =   16
      Max             =   255
      SelStart        =   128
      TickFrequency   =   16
      Value           =   128
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   480
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Green"
      Top             =   4440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   847
      _Version        =   393216
      LargeChange     =   16
      Max             =   255
      SelStart        =   128
      TickFrequency   =   16
      Value           =   128
   End
   Begin VB.Label lblColor 
      BackColor       =   &H80000016&
      Caption         =   "RGB"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   6375
   End
End
Attribute VB_Name = "frmDisplayProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ColorDesc As String

Private Sub cmdApply_Click()
Dim clr As Long
    clr = RGBFromRedGreenBlue(Slider1.Value, Slider2.Value, Slider3.Value)
    Select Case DisplayColors.ListIndex
        Case 0
            ' do nothing
        Case 1
            Common_BackColor = clr
        Case 2
            Entry_BackColor = clr
        Case 3
            EntryInvalid_BackColor = clr
        Case 4
            EntryUnsaved_BackColor = clr
        Case 5
            EntryNotChangeable_BackColor = clr
        Case 6
            MasterMode_BackColor = clr
        Case 7
            StationMode_BackColor = clr
        Case 8
            Alarm_ForeColor = clr
        Case 9
            BarActual_ForeColor = clr
        Case 10
            Data_ForeColor = clr
        Case 11
            DataBold_ForeColor = clr
        Case 12
            DataHiLite_ForeColor = clr
        Case 13
            Entry_ForeColor = clr
        Case 14
            Good_ForeColor = clr
        Case 15
            Message_ForeColor = clr
        Case 16
            Titles_ForeColor = clr
        Case 17
            TitlesData_Forecolor = clr
        Case 18
            TitlesLabel_ForeColor = clr
        Case 19
            Warning_ForeColor = clr
        Case Else
            ' do nothing
    End Select
End Sub

Private Sub cmdReturn_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Save_ScreenColors
End Sub

Private Sub cmdSetDefaults_Click()
    SetDefault_ScreenColors
    DisplayColors_Click
End Sub

Private Sub DisplayColors_Click()
Dim clr As Long
    ColorDesc = DisplayColors.List(DisplayColors.ListIndex)
    Select Case DisplayColors.ListIndex
        Case 0
            clr = frmDisplayProperties.Point(90, 90)
        Case 1
            clr = Common_BackColor
        Case 2
            clr = Entry_BackColor
        Case 3
            clr = EntryInvalid_BackColor
        Case 4
            clr = EntryUnsaved_BackColor
        Case 5
            clr = EntryNotChangeable_BackColor
        Case 6
            clr = MasterMode_BackColor
        Case 7
            clr = StationMode_BackColor
        Case 8
            clr = Alarm_ForeColor
        Case 9
            clr = BarActual_ForeColor
        Case 10
            clr = Data_ForeColor
        Case 11
            clr = DataBold_ForeColor
        Case 12
            clr = DataHiLite_ForeColor
        Case 13
            clr = Entry_ForeColor
        Case 14
            clr = Good_ForeColor
        Case 15
            clr = Message_ForeColor
        Case 16
            clr = Titles_ForeColor
        Case 17
            clr = TitlesData_Forecolor
        Case 18
            clr = TitlesLabel_ForeColor
        Case 19
            clr = Warning_ForeColor
        Case Else
            ColorDesc = "Other Color"
            clr = White
    End Select
    Slider1.Value = RedFromRGB(clr)
    Slider2.Value = GreenFromRGB(clr)
    Slider3.Value = BlueFromRGB(clr)
End Sub

Private Sub Form_Load()
    DisplayColors.AddItem "ButtonFace", 0
    DisplayColors.AddItem "Common_BackColor", 1
    DisplayColors.AddItem "Entry_BackColor", 2
    DisplayColors.AddItem "EntryInvalid_BackColor", 3
    DisplayColors.AddItem "EntryUnsaved_BackColor", 4
    DisplayColors.AddItem "EntryNotChangeable_BackColor", 5
    DisplayColors.AddItem "MasterMode_BackColor", 6
    DisplayColors.AddItem "StationMode_BackColor", 7
    DisplayColors.AddItem "Alarm_ForeColor", 8
    DisplayColors.AddItem "BarActual_ForeColor", 9
    DisplayColors.AddItem "Data_ForeColor", 10
    DisplayColors.AddItem "DataBold_ForeColor", 11
    DisplayColors.AddItem "DataHiLite_ForeColor", 12
    DisplayColors.AddItem "Entry_ForeColor", 13
    DisplayColors.AddItem "Good_ForeColor", 14
    DisplayColors.AddItem "Message_ForeColor", 15
    DisplayColors.AddItem "Titles_ForeColor", 16
    DisplayColors.AddItem "TitlesData_ForeColor", 17
    DisplayColors.AddItem "TitlesLabel_ForeColor", 18
    DisplayColors.AddItem "Warning_ForeColor", 19
    DisplayColors.ListIndex = 1
    DisplayColors_Click
    Form_Center Me
End Sub

Private Sub tmrScreen_Timer()
Dim rgbcolor As Long
    rgbcolor = RGBFromRedGreenBlue(CLng(Slider1.Value), CLng(Slider2.Value), CLng(Slider3.Value))
    lblColor.BackColor = rgbcolor
    lblColor.Caption = ColorDesc & vbCrLf & vbCrLf & "Red " & Format(Slider1.Value, "##0") & "   Green " & Format(Slider2.Value, "##0") & "   Blue " & Format(Slider3.Value, "##0")
End Sub


