VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSearchProf 
   BackColor       =   &H80000005&
   Caption         =   "Master Purge Profiles"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "frmSearchProf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1080
   ScaleMode       =   0  'User
   ScaleWidth      =   1920
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgProfiles 
      Bindings        =   "frmSearchProf.frx":57E2
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6588
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
      Caption         =   "Purge Profiles"
      ColumnCount     =   6
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
         DataField       =   "Steps"
         Caption         =   "Steps"
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
         DataField       =   "TotalDuration"
         Caption         =   "TotalDuration"
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
         DataField       =   "ProjectedLiters"
         Caption         =   "ProjectedLiters"
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
         DataField       =   "ProjectedVolumes"
         Caption         =   "ProjectedVolumes"
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
            ColumnWidth     =   80.703
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   222.931
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   78.814
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   132.64
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   144.117
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   172.955
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
         DisabledPicture =   "frmSearchProf.frx":57FC
         DownPicture     =   "frmSearchProf.frx":643E
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
         Picture         =   "frmSearchProf.frx":7080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         DisabledPicture =   "frmSearchProf.frx":7CC2
         DownPicture     =   "frmSearchProf.frx":8904
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
         Picture         =   "frmSearchProf.frx":9546
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmSearchProf.frx":A188
         DownPicture     =   "frmSearchProf.frx":ADCA
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
         Picture         =   "frmSearchProf.frx":BA0C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCreateNew 
         Caption         =   " Create New"
         DisabledPicture =   "frmSearchProf.frx":C64E
         DownPicture     =   "frmSearchProf.frx":C990
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
         Picture         =   "frmSearchProf.frx":CCD2
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
      Begin MSAdodcLib.Adodc adoProfiles 
         Height          =   375
         Left            =   12960
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   "SELECT * FROM [MasterProfiles] ORDER BY [Number] ASC"
         Caption         =   "Profiles"
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
      Begin MSAdodcLib.Adodc adoProfileSteps 
         Height          =   375
         Left            =   12960
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         RecordSource    =   "SELECT * FROM [MasterProfileSteps] ORDER BY [ProfileNumber],[StepNumber] ASC"
         Caption         =   "ProfileSteps"
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
         Left            =   4665
         TabIndex        =   2
         Top             =   120
         Width           =   6495
      End
   End
   Begin MSDataGridLib.DataGrid dgProfileSteps 
      Bindings        =   "frmSearchProf.frx":D014
      Height          =   5460
      Left            =   0
      TabIndex        =   3
      Top             =   3735
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9631
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
      Caption         =   "Purge Profile Steps"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "ProfileNumber"
         Caption         =   "ProfileNumber"
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
         DataField       =   "StepNumber"
         Caption         =   "StepNumber"
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
         DataField       =   "InitialSp"
         Caption         =   "InitialSp"
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
         DataField       =   "Duration"
         Caption         =   "Duration"
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
         DataField       =   "StepType"
         Caption         =   "StepType"
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
         DataField       =   "StepTypeDesc"
         Caption         =   "StepTypeDesc"
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
            ColumnWidth     =   136.49
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   122.979
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   136.49
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   136.49
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   97.991
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   222.931
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearchProf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 358'''''''''' Form SearchProf.frm '''''''''''''''''''
Option Explicit
'
Dim sPath As String
Dim rsCrit As String
Private ProfSelDest As Integer
Private antiRepeatDelete As Boolean
Private searchPrfMsg As String
Private searchPrfMsgColor As Long

Private Sub adoProfiles_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    antiRepeatDelete = False
    searchPrfMsgColor = Message_ForeColor
    searchPrfMsg = ""
End Sub

Private Sub cmdCreateNew_Click()
    searchPrfMsgColor = Message_ForeColor
    searchPrfMsg = ""
    NewPrf
End Sub

Private Sub cmdClear_Click()
    searchPrfMsgColor = Message_ForeColor
    searchPrfMsg = ""
    dgProfiles.Height = 4095
    ClearPrf
    dgProfiles.Height = 9255
End Sub

Private Sub cmdDelete_Click()
    searchPrfMsgColor = Message_ForeColor
    searchPrfMsg = ""
    dgProfiles.Height = 4095
    DeletePrf
    dgProfiles.Height = 9255
End Sub

Public Sub ChgSelectionDestination(ByVal NewDest As Integer)
    ' 1=profile; 2=recipe
    ProfSelDest = IIf((NewDest = profdestProfile Or NewDest = profdestRecipe), NewDest, profdestProfile)
End Sub

Private Sub Xit()
    Unload frmSearchProf
    Set frmSearchProf = Nothing
End Sub

Private Sub cmdSelect_Click()
Dim recnum As Integer
    If Not adoProfiles.Recordset.BOF Then
        recnum = CInt(dgProfiles.Columns(0).CellValue(dgProfiles.GetBookmark(0)))
        Select Case ProfSelDest
            Case profdestProfile
                frmPurgeProfile.Show
                frmPurgeProfile.LoadNewProf CInt(recnum)
            Case profdestRecipe
                frmRecipe.SetPurgeProfile CInt(recnum)
        End Select
        Unload frmSearchProf
        Set frmSearchProf = Nothing
    End If
End Sub

Private Sub dgProfiles_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex + 1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSearchProf = Nothing
    End If
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 358, 2
Dim flag1 As Boolean
Dim flag2 As Boolean

    KeyPreview = True
    
    flag1 = CheckPass("P", False) And CheckPass("7", False)
    flag2 = CheckPass("P", False) And (CheckPass("8", False) Or CheckPass("7", False))
    cmdClear.Visible = IIf(flag1, True, False)
    cmdCreateNew.Visible = IIf(flag2, True, False)
    cmdDelete.Visible = IIf(flag2, True, False)
    cmdSelect.Visible = IIf(flag2, True, False)
    
    dgProfiles.AllowRowSizing = False
'    dgProfiles.Height = 9255
    dgProfiles.Height = 400
    dgProfileSteps.Height = 585
    
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
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [Number] ASC"
        Case 2
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [Description] ASC"
        Case 3
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [Steps] DESC"
        Case 4
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [TotalDuration] DESC"
        Case 5
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [ProjectedLiters] DESC"
        Case 6
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [ProjectedVolumes] DESC"
        Case Else
            sortCol = 1
            rsCrit = "SELECT * FROM [MasterProfiles] ORDER BY [Number] ASC"
    End Select
    adoProfiles.RecordSource = rsCrit
    adoProfiles.Refresh

    If adoProfiles.Recordset.BOF Then
        dgProfiles.Caption = " No Defined Profiles"
        ' Set column properties
        dgProfiles.Columns(0).Width = 760
        dgProfiles.Columns(1).Width = 4000
        dgProfiles.Columns(2).Width = 760
        dgProfiles.Columns(3).Width = 1250
        dgProfiles.Columns(4).Width = 1250
        dgProfiles.Columns(5).Width = 1250
        cmdClear.Enabled = False
        cmdDelete.Enabled = False
        cmdSelect.Enabled = False
    Else
        ' Display number of profiles found
        adoProfiles.Recordset.GetRows
        Select Case adoProfiles.Recordset.RecordCount
            Case 0
                dgProfiles.Caption = " No Defined Profiles"
                cmdClear.Enabled = False
                cmdSelect.Enabled = False
                cmdSelect.Enabled = False
            Case 1
                dgProfiles.Caption = Format(adoProfiles.Recordset.RecordCount, "###0") & " Defined Profile"
                cmdClear.Enabled = False
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
            Case Else
                dgProfiles.Caption = Format(adoProfiles.Recordset.RecordCount, "###0") & " Defined Profiles"
                cmdClear.Enabled = True
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
        End Select
        dgProfiles.Refresh
        ' Set column properties
        dgProfiles.Columns(0).Width = 760
        dgProfiles.Columns(1).Width = 4000
        dgProfiles.Columns(2).Width = 760
        dgProfiles.Columns(3).Width = 1250
        dgProfiles.Columns(4).Width = 1250
        dgProfiles.Columns(5).Width = 1250
        
        ' move pointer to first row
        adoProfiles.Recordset.MoveFirst
        
        ' make the Left-Most column the Sorted-By column
        dgProfiles.LeftCol = IIf(sortCol > 5, 5, sortCol - 1)
        
    End If
    
End Sub

Private Sub DeletePrf()
'
SetErrModule 358, 31
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim rS As ADODB.Recordset
Dim prfnum As Integer

    If Not antiRepeatDelete Then
        If adoProfiles.Recordset.BOF Then
            searchPrfMsgColor = MEDRED
            searchPrfMsg = "No Profile Data Available"
        Else
        
            If IsNull(dgProfiles.Columns(0).CellValue(dgProfiles.GetBookmark(0))) Or IsEmpty(dgProfiles.Columns(0).CellValue(dgProfiles.GetBookmark(0))) Then
                ' Report an error
                searchPrfMsgColor = MEDRED
                searchPrfMsg = "Invalid Profile Number"
                Exit Sub
            End If
            
            ' profile
            prfnum = dgProfiles.Columns(0).CellValue(dgProfiles.GetBookmark(0))
            adoProfiles.Recordset.Delete
            
            ' steps
            adoProfileSteps.RecordSource = "SELECT * FROM [MasterProfileSteps] with [ProfileNumber] = " & prfnum & "  ORDER BY [StepNumber] ASC"

            Set rS = adoProfileSteps.Recordset
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
                        If .Fields("ProfileNumber").Value = CSng(prfnum) Then
                            If .Supports(adDelete) Then
                            'It is possible that the record you want to update
                            'is locked by another user. If we don't check before
                            'updating, we will generate an error
                            
                                .Delete
                                .MovePrevious
                            Else
                                searchPrfMsgColor = MEDRED
                                searchPrfMsg = "Unable to Delete Profile; Profile Steps Record Locked"
                            End If
                        Else
                            .MovePrevious
                        End If
                    Loop
                
                End If
                
                ' close recordset
                .Close
            
            End With
            

            adoProfileSteps.RecordSource = "SELECT * FROM [MasterProfileSteps] ORDER BY [ProfileNumber] ASC,[StepNumber] ASC"
            adoProfiles.RecordSource = "SELECT * FROM [MasterProfiles] ORDER BY [Number] ASC"
            dgProfiles.Refresh
            dgProfileSteps.Refresh
                    
            searchPrfMsgColor = Message_ForeColor
            searchPrfMsg = "Profile Deleted"
            antiRepeatDelete = True
           
        End If
    End If
    
ResetErrModule
Exit Sub

localhandler:
    searchPrfMsgColor = MEDRED
    searchPrfMsg = "Unable to Delete Profile"
    Set rS = Nothing
    '...and set it to nothing
    Exit Sub
End Sub

Private Sub ClearPrf()
Dim prfnum As Integer
Dim iStep As Integer

    SetErrModule 358, 3
    If UseLocalErrorHandler Then On Error GoTo localhandler
    ' Clearing Profile
    searchPrfMsgColor = Message_ForeColor
    searchPrfMsg = "Clearing Profile.. Please Wait"
    
    ' Blank Profile
    prfnum = dgProfiles.Columns(0).CellValue(dgProfiles.GetBookmark(0))
    adoProfiles.RecordSource = "SELECT * FROM [MasterProfiles] with [Number] = " & prfnum & "  ORDER BY [Number] ASC"
'    adoProfiles.Recordset.MoveLast
'    adoProfiles.Recordset.Field("Number").Value = CInt(0)
    adoProfiles.Recordset.Fields("Description").Value = "undefined"
    adoProfiles.Recordset.Fields("TotalDuration").Value = CSng(0)
'    adoProfiles.Recordset.Fields("DurDesc").Value = ProfileDurationDescription(adoProfiles.Recordset("Duration"))
    adoProfiles.Recordset.Fields("Steps").Value = CInt(1)
    adoProfiles.Recordset.Fields("ProjectedLiters").Value = CSng(0)
    adoProfiles.Recordset.Fields("ProjectedVolumes").Value = CSng(0)
'    adoProfiles.Recordset.Fields("Validated").Value = False
    adoProfiles.Recordset.Update
    
    ' steps
    adoProfileSteps.RecordSource = "SELECT * FROM [MasterProfileSteps] with [ProfileNumber] = " & prfnum & "  ORDER BY [StepNumber] ASC"
    dgProfileSteps.Refresh
    frmSearchProf.Refresh
    adoProfileSteps.Recordset.MoveLast
    Do Until adoProfileSteps.Recordset.BOF
        If adoProfileSteps.Recordset.Fields("ProfileNumber").Value = CSng(prfnum) Then
            If adoProfileSteps.Recordset.Fields("StepNumber").Value = CSng(1) Then
                adoProfileSteps.Recordset.Fields("Duration").Value = CSng(0)
                adoProfileSteps.Recordset.Fields("InitialSp").Value = CSng(0)
                adoProfileSteps.Recordset.Fields("StepType").Value = CInt(0)
                adoProfileSteps.Recordset.Fields("StepTypeDesc").Value = "undefined"
                adoProfileSteps.Recordset.Update
                adoProfileSteps.Recordset.MovePrevious
            Else
                adoProfileSteps.Recordset.Delete
            End If
        Else
            adoProfileSteps.Recordset.MovePrevious
        End If
    Loop
    
    adoProfileSteps.RecordSource = "SELECT * FROM [MasterProfileSteps] ORDER BY [ProfileNumber],[StepNumber] ASC"
    adoProfiles.RecordSource = "SELECT * FROM [MasterProfiles] ORDER BY [Number] ASC"
    dgProfiles.Refresh
                    
    searchPrfMsgColor = Message_ForeColor
    searchPrfMsg = "Profile Cleared"
    
ResetErrModule
Exit Sub

localhandler:
    searchPrfMsgColor = MEDRED
    searchPrfMsg = "Unable to Clear Profile"
End Sub

Private Sub NewPrf()
Dim iPrf As Integer
Dim prfnum As Integer

    prfnum = 0
    For iPrf = 1 To MAX_PROFILES
        If prfnum = 0 Then
            If Not IsDefined(iPrf, adoProfiles.Recordset) Then
                prfnum = iPrf
            End If
        End If
    Next iPrf
    
    If prfnum > 0 Then
        frmPurgeProfile.Show
        frmPurgeProfile.LoadNewProf CInt(prfnum)
        Unload frmSearchProf
        Set frmSearchProf = Nothing
    Else
        searchPrfMsgColor = MEDRED
        searchPrfMsg = "No undefined Master Purge Profile"
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
    lblMessage.ForeColor = searchPrfMsgColor
    lblMessage.Caption = searchPrfMsg
End Sub

