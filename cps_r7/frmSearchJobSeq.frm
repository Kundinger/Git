VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSearchJobSeq 
   BackColor       =   &H80000005&
   Caption         =   "Master Job Sequences"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "frmSearchJobSeq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgJobSeq 
      Bindings        =   "frmSearchJobSeq.frx":57E2
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Job Sequences"
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "Number"
         Caption         =   "Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Description"
         Caption         =   "Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Courses"
         Caption         =   "Courses"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PriScale"
         Caption         =   "PriScale"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AuxScale"
         Caption         =   "AuxScale"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "IDLoad"
         Caption         =   "IDLoad"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "LoadL"
         Caption         =   "LoadL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "IDPurge"
         Caption         =   "IDPurge"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PurgeL"
         Caption         =   "PurgeL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "IDVent"
         Caption         =   "IDVent"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "VentL"
         Caption         =   "VentL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "LoadV"
         Caption         =   "LoadV"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "PurgeV"
         Caption         =   "PurgeV"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "VentV"
         Caption         =   "VentV"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "EstSeqDuration"
         Caption         =   "EstSeqDuration"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "EstSeqDurDesc"
         Caption         =   "EstSeqDurDesc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "Validated"
         Caption         =   "Validated"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pbxBottom 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   9240
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   1482
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         DisabledPicture =   "frmSearchJobSeq.frx":57FA
         DownPicture     =   "frmSearchJobSeq.frx":643C
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
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchJobSeq.frx":707E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         DisabledPicture =   "frmSearchJobSeq.frx":7CC0
         DownPicture     =   "frmSearchJobSeq.frx":8902
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
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchJobSeq.frx":9544
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmSearchJobSeq.frx":A186
         DownPicture     =   "frmSearchJobSeq.frx":ADC8
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
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchJobSeq.frx":BA0A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCreateNew 
         Caption         =   " Create New"
         DisabledPicture =   "frmSearchJobSeq.frx":C64C
         DownPicture     =   "frmSearchJobSeq.frx":C98E
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
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchJobSeq.frx":CCD0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Timer tmrScreen 
         Interval        =   250
         Left            =   13800
         Top             =   0
      End
      Begin MSAdodcLib.Adodc adoJobSeq 
         Height          =   375
         Left            =   12240
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=CpsRecipes"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "CpsRecipes"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [MasterSequence] ORDER BY [Number] ASC"
         Caption         =   "JobSequences"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc adoCourses 
         Height          =   375
         Left            =   12240
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=CpsRecipes"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "CpsRecipes"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [MasterSequenceCourses] ORDER BY [SeqNum] ASC ,[CourseNumber] ASC"
         Caption         =   "Courses"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   4545
         TabIndex        =   2
         Top             =   120
         Width           =   6495
      End
   End
   Begin MSDataGridLib.DataGrid dgCourses 
      Bindings        =   "frmSearchJobSeq.frx":D012
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Courses"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearchJobSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 3158'''''''''' Form SearchJobSeq.frm '''''''''''''''''''
Option Explicit
Dim sPath As String
Dim rsCrit As String
Private antiRepeatDelete As Boolean
Private searchSeqMsg As String
Private searchSeqMsgColor As Long

Private Sub adoJobSeq_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    antiRepeatDelete = False
    searchSeqMsgColor = Message_ForeColor
    searchSeqMsg = ""
End Sub

Private Sub cmdCreateNew_Click()
    searchSeqMsgColor = Message_ForeColor
    searchSeqMsg = ""
    NewSeq
End Sub

Private Sub cmdClear_Click()
    searchSeqMsgColor = Message_ForeColor
    searchSeqMsg = ""
    dgJobSeq.Height = 4095
    ClearSeq
    dgJobSeq.Height = 9255
End Sub

Private Sub cmdDelete_Click()
    searchSeqMsgColor = Message_ForeColor
    searchSeqMsg = ""
    dgJobSeq.Height = 4095
    DeleteSeq
    dgJobSeq.Height = 9255
End Sub


Private Sub Xit()
    Unload frmSearchJobSeq
    Set frmSearchJobSeq = Nothing
End Sub

Private Sub cmdSelect_Click()
Dim recnum As Integer
    recnum = CInt(dgJobSeq.Columns(0).CellValue(dgJobSeq.GetBookmark(0)))
    frmCourses.Show
    frmCourses.LoadMaster CInt(recnum)
    Unload frmSearchJobSeq
    Set frmSearchJobSeq = Nothing
End Sub

Private Sub dgJobSeq_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex + 1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSearchJobSeq = Nothing
    End If
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3158, 2
Dim flag1 As Boolean
Dim flag2 As Boolean

    KeyPreview = True
    
    flag1 = CheckPass("P", False) And CheckPass("7", False)
    flag2 = CheckPass("P", False) And (CheckPass("8", False) Or CheckPass("7", False))
    cmdClear.Visible = IIf(flag1, True, False)
    cmdCreateNew.Visible = IIf(flag2, True, False)
    cmdDelete.Visible = IIf(flag2, True, False)
    cmdSelect.Visible = IIf(flag2, True, False)
    
    dgJobSeq.AllowRowSizing = False
    
    DisplayData 1
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
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

Private Sub DisplayData(sortCol As Integer)
    ' Select & Sort
    Select Case sortCol
        Case 1
            rsCrit = "SELECT * FROM [MasterSequence] ORDER BY [Number] ASC"
        Case 2
            rsCrit = "SELECT * FROM [MasterSequence] ORDER BY [Description] ASC"
        Case 3
            rsCrit = "SELECT * FROM [MasterSequence] ORDER BY [Courses] DESC"
        Case 15
            rsCrit = "SELECT * FROM [MasterSequence] ORDER BY [EstSeqDuration] DESC"
        Case 16
            rsCrit = "SELECT * FROM [MasterSequence] ORDER BY [EstSeqDuration] DESC"
        Case Else
            sortCol = 1
            rsCrit = "SELECT * FROM [MasterSequence] ORDER BY [Number] ASC"
    End Select
    adoJobSeq.RecordSource = rsCrit
    adoJobSeq.Refresh

    If adoJobSeq.Recordset.BOF Then
        dgJobSeq.Caption = " No Defined Job Sequences"
        ' Set column properties
        dgJobSeq.Columns(0).Width = 760
        dgJobSeq.Columns(1).Width = 4000
        dgJobSeq.Columns(2).Width = 760
        dgJobSeq.Columns(3).Width = 1250
        dgJobSeq.Columns(10).Width = 1650
        dgJobSeq.Columns(11).Width = 1500
    '    dgJobSeq.Columns(49).Width = 2400
        cmdClear.Enabled = False
        cmdDelete.Enabled = False
        cmdSelect.Enabled = False
    Else
        ' Display number of Job Sequences found
        adoJobSeq.Recordset.GetRows
        Select Case adoJobSeq.Recordset.RecordCount
            Case 0
                dgJobSeq.Caption = " No Defined Job Sequences"
                cmdClear.Enabled = False
                cmdSelect.Enabled = False
                cmdSelect.Enabled = False
            Case 1
                dgJobSeq.Caption = Format(adoJobSeq.Recordset.RecordCount, "###0") & " Defined Job Sequence"
                cmdClear.Enabled = False
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
            Case Else
                dgJobSeq.Caption = Format(adoJobSeq.Recordset.RecordCount, "###0") & " Defined Job Sequences"
                cmdClear.Enabled = True
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
        End Select
        dgJobSeq.Refresh
        ' Set column properties
        dgJobSeq.Columns(0).Width = 760
        dgJobSeq.Columns(1).Width = 4000
        dgJobSeq.Columns(2).Width = 760
        dgJobSeq.Columns(3).Width = 1250
        dgJobSeq.Columns(10).Width = 1650
        dgJobSeq.Columns(11).Width = 1500
    '    dgJobSeq.Columns(49).Width = 2400
        
        ' move pointer to first row
        adoJobSeq.Recordset.MoveFirst
        
        ' make the Left-Most column the Sorted-By column
        dgJobSeq.LeftCol = IIf(sortCol > 51, 51, sortCol - 1)
    End If
    
End Sub

Private Sub DeleteSeq()
SetErrModule 3158, 31
If UseLocalErrorHandler Then On Error GoTo localhandler
Dim rS As ADODB.Recordset
Dim seqnum As Integer
    
    If Not antiRepeatDelete Then
        If adoJobSeq.Recordset.BOF Then
            searchSeqMsgColor = MEDRED
            searchSeqMsg = "No Job Sequence Data Available"
        Else
        
            If IsNull(dgJobSeq.Columns(0).CellValue(dgJobSeq.GetBookmark(0))) Or IsEmpty(dgJobSeq.Columns(0).CellValue(dgJobSeq.GetBookmark(0))) Then
                ' Report an error
                searchSeqMsgColor = MEDRED
                searchSeqMsg = "Invalid Job Sequence Number"
                Exit Sub
            End If
            
            ' sequence
            seqnum = dgJobSeq.Columns(0).CellValue(dgJobSeq.GetBookmark(0))
            adoJobSeq.Recordset.Delete
            
            ' courses
            adoCourses.RecordSource = "SELECT * FROM [MasterSequenceCourses] with [SeqNum] = " & seqnum & "  ORDER BY [CourseNumber] ASC"

            Set rS = adoCourses.Recordset
'            rS.Open
            With rS
            
                If Not .BOF And Not .EOF Then
                'Ensure that the recordset contains records
                'If no records the code inside the if...end if
                'statement won't run
                
                    .MoveLast
                    .MoveFirst
                    .MoveLast
                    'Not necessary but good practice
                    
                    Do Until .BOF
                        If .Fields("SeqNum").Value = CSng(seqnum) Then
                            If .Supports(adDelete) Then
                            'It is possible that the record you want to update
                            'is locked by another user. If we don't check before
                            'updating, we will generate an error
                            
                                .Delete
                                .MovePrevious
                            Else
                                searchSeqMsgColor = MEDRED
                                searchSeqMsg = "Unable to Delete Sequence; Course Record Locked"
                                Set rS = Nothing
                                '...and set it to nothing
                                Exit Sub
                            End If
                        Else
                            .MovePrevious
                        End If
                    Loop
                
                Else
                    searchSeqMsgColor = MEDRED
                    searchSeqMsg = "Unable to Delete Sequence; No Course Records"
                    Set rS = Nothing
                    '...and set it to nothing
                    Exit Sub
                End If
                
                ' close recordset
                .Close
            
            End With
            
            adoCourses.RecordSource = "SELECT * FROM [MasterSequenceCourses] ORDER BY [SeqNum] ASC,[CourseNumber] ASC"
            adoJobSeq.RecordSource = "SELECT * FROM [MasterSequence] ORDER BY [Number] ASC"
            dgJobSeq.Refresh
            dgCourses.Refresh
            
            searchSeqMsgColor = Message_ForeColor
            searchSeqMsg = "Job Sequence Deleted"
            antiRepeatDelete = True
           
        End If
    End If
Exit Sub
localhandler:
    searchSeqMsgColor = MEDRED
    searchSeqMsg = "Unable to Delete Job Sequence"
    Set rS = Nothing
    '...and set it to nothing
    Exit Sub
End Sub

Private Sub ClearSeq()
Dim seqnum As Integer
Dim iCourse As Integer

    SetErrModule 3158, 3
    If UseLocalErrorHandler Then On Error GoTo localhandler
    ' Clearing Job Sequence
    searchSeqMsgColor = Message_ForeColor
    searchSeqMsg = "Clearing Job Sequence.. Please Wait"
    adoCourses.Recordset.MoveLast
    
    ' job sequence
    seqnum = dgJobSeq.Columns(0).CellValue(dgJobSeq.GetBookmark(0))
    adoJobSeq.RecordSource = "SELECT * FROM [MasterSequence] with [Number] = " & seqnum & "  ORDER BY [Number] ASC"
'    adoJobSeq.Recordset("Number") = CInt(0)
    adoJobSeq.Recordset("Description") = "default sequence"
    adoJobSeq.Recordset("PriScale") = 0
    adoJobSeq.Recordset("AuxScale") = 0
    adoJobSeq.Recordset("EstSeqDuration") = 0
'    adoJobSeq.Recordset("EstSeqDurDesc") = DurationDescription(adoJobSeq.Recordset("EstSeqDuration"))
    adoJobSeq.Recordset("Courses") = 1
    adoJobSeq.Recordset("IDLoad") = 0
    adoJobSeq.Recordset("IDPurge") = 0
    adoJobSeq.Recordset("IDVent") = 0
    adoJobSeq.Recordset("LoadL") = 0
    adoJobSeq.Recordset("LoadV") = 0
    adoJobSeq.Recordset("PurgeL") = 0
    adoJobSeq.Recordset("PurgeV") = 0
    adoJobSeq.Recordset("VentL") = 0
    adoJobSeq.Recordset("VentV") = 0
    adoJobSeq.Recordset("Validated") = False
    
    ' courses
    adoCourses.RecordSource = "SELECT * FROM [MasterSequenceCourses] with [SeqNum] = " & seqnum & "  ORDER BY [CourseNumber] ASC"
    dgCourses.Refresh
    frmSearchJobSeq.Refresh
    adoCourses.Recordset.MoveLast
    adoCourses.Recordset.MoveFirst
    adoCourses.Recordset.MoveLast
    Do Until adoCourses.Recordset.BOF
        If adoCourses.Recordset.Fields("SeqNum").Value = CSng(seqnum) Then
            If adoCourses.Recordset.Fields("CourseNumber").Value = CSng(1) Then
                adoCourses.Recordset.Fields("Type").Value = courseRecipe
                adoCourses.Recordset.Fields("PauseDuration").Value = 0
                adoCourses.Recordset.Fields("RecipeNumber").Value = 0
                adoCourses.Recordset.Fields("Cycles").Value = 0
                adoCourses.Recordset.Fields("LoadRate").Value = 0
                adoCourses.Recordset.Fields("PurgeRate").Value = 0
                adoCourses.Recordset.Fields("EstCourseDuration").Value = 0
                adoCourses.Recordset.Fields("MsgText").Value = "-na-"
                adoCourses.Recordset.Update
                adoCourses.Recordset.MovePrevious
            Else
                adoCourses.Recordset.Delete
            End If
        Else
            adoCourses.Recordset.MovePrevious
        End If
    Loop
    
    adoCourses.RecordSource = "SELECT * FROM [MasterSequenceSteps] ORDER BY [SeqNum] ASC,[CourseNumber] ASC"
    adoJobSeq.RecordSource = "SELECT * FROM [MasterSequence] ORDER BY [Number] ASC"
    dgJobSeq.Refresh
                        
    searchSeqMsgColor = Message_ForeColor
    searchSeqMsg = "Job Sequence Cleared"
    
ResetErrModule
Exit Sub

localhandler:
    searchSeqMsgColor = MEDRED
    searchSeqMsg = "Unable to Clear Job Sequence"
End Sub

Private Sub NewSeq()
Dim iSeq As Integer
Dim seqnum As Integer

    seqnum = 0
    For iSeq = 1 To NR_JOBSEQ
        If seqnum = 0 Then
            If Not IsDefined(iSeq, adoJobSeq.Recordset) Then
                seqnum = iSeq
            End If
        End If
    Next iSeq
    
    If seqnum > 0 Then
        frmCourses.Show
        frmCourses.LoadMaster CInt(seqnum)
        Unload frmSearchJobSeq
        Set frmSearchJobSeq = Nothing
    Else
        searchSeqMsgColor = MEDRED
        searchSeqMsg = "No undefined Master Seqister"
    End If
End Sub

Private Function IsDefined(ByVal iNum As Integer, ByRef rS As ADODB.Recordset) As Boolean
Dim flag As Boolean

    flag = False
    
    With rS
    
        If Not .BOF Or Not .EOF Then
        
            .MoveLast
            .MoveFirst
            .MoveLast
            
            Do Until (.BOF Or flag)
                If ((iNum = .Fields("Number").Value) And (.Fields("Description").Value <> "undefined") And (Len(Trim(.Fields("Description").Value)) > 0)) Then
                    flag = True
                End If
                .MovePrevious
            Loop
        
        End If
    
    End With
    
    IsDefined = flag
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub tmrScreen_Timer()
    lblMessage.ForeColor = searchSeqMsgColor
    lblMessage.Caption = searchSeqMsg
End Sub


