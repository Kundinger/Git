VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSearchCan 
   BackColor       =   &H80000005&
   Caption         =   "Master Canisters"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   Icon            =   "frmSearchCan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10830
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgCanisters 
      Bindings        =   "frmSearchCan.frx":57E2
      Height          =   9855
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   17383
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
      Caption         =   "Canisters"
      ColumnCount     =   4
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
         DataField       =   "WorkingCapacity"
         Caption         =   "WorkingCapacity"
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
         DataField       =   "WCVolume"
         Caption         =   "WCVolume"
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
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1049.953
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCanisters 
      Height          =   375
      Left            =   3960
      Top             =   10440
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      RecordSource    =   "SELECT * FROM [MasterCanister] ORDER BY [Number] ASC"
      Caption         =   "Canisters"
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
   Begin Threed.SSPanel pbxBottom 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   9915
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
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
         DisabledPicture =   "frmSearchCan.frx":57FD
         DownPicture     =   "frmSearchCan.frx":643F
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
         Picture         =   "frmSearchCan.frx":7081
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Press to Select highlighted Canister"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         DisabledPicture =   "frmSearchCan.frx":7CC3
         DownPicture     =   "frmSearchCan.frx":8905
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
         Picture         =   "frmSearchCan.frx":9547
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmSearchCan.frx":A189
         DownPicture     =   "frmSearchCan.frx":ADCB
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
         Picture         =   "frmSearchCan.frx":BA0D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCreateNew 
         Caption         =   " Create New"
         DisabledPicture =   "frmSearchCan.frx":C64F
         DownPicture     =   "frmSearchCan.frx":C991
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
         Picture         =   "frmSearchCan.frx":CCD3
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Create a new Canister definition"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Timer tmrScreen 
         Interval        =   250
         Left            =   5280
         Top             =   0
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " message message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   735
         Left            =   3360
         TabIndex        =   2
         Top             =   120
         Width           =   3450
      End
   End
End
Attribute VB_Name = "frmSearchCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 258'''''''''' Form SearchCan.frm '''''''''''''''''''
Option Explicit
Dim sPath As String
Dim rsCrit As String
Dim RowHgt As Single
Dim InitRow As Integer
Private CanRcpMode As Integer            ' 0=master; 1=station
Private antiRepeatDelete As Boolean
Private searchCanMsg As String
Private searchCanMsgColor As Long

Public Sub ChgCanRcpMode(ByVal NewMode As Integer)
    ' 0=master; 1=station
    CanRcpMode = IIf((NewMode = 0 Or NewMode = 1), NewMode, 0)
    Select Case CanRcpMode
        Case MASTERMODE
            ' clear button
            cmdClear.Visible = True
            ' create new button
            cmdCreateNew.Visible = True
            ' delete button
            cmdDelete.Visible = True
        Case STATIONMODE
            ' clear button
            cmdClear.Visible = False
            ' create new button
            cmdCreateNew.Visible = False
            ' delete button
            cmdDelete.Visible = False
    End Select
End Sub

Private Sub adoCanisters_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    antiRepeatDelete = False
    searchCanMsgColor = Message_ForeColor
    searchCanMsg = ""
End Sub

Private Sub cmdClear_Click()
    searchCanMsgColor = Message_ForeColor
    searchCanMsg = ""
    ClearCan
End Sub

Private Sub cmdCreateNew_Click()
    searchCanMsgColor = Message_ForeColor
    searchCanMsg = ""
    NewCan
End Sub

Private Sub cmdDelete_Click()
    searchCanMsgColor = Message_ForeColor
    searchCanMsg = ""
    DeleteCan
End Sub

Private Sub Xit()
    Unload frmSearchCan
    Set frmSearchCan = Nothing
End Sub

Private Sub cmdSelect_Click()
Dim recnum As Integer
    recnum = dgCanisters.Columns(0).CellValue(dgCanisters.GetBookmark(0))
    frmCanRecipe.Show
    frmCanRecipe.LoadNewCan CInt(recnum)
    Unload frmSearchCan
    Set frmSearchCan = Nothing
End Sub

Private Sub dgCanisters_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex + 1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSearchCan = Nothing
    End If
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 258, 2
Dim flag1 As Boolean
Dim flag2 As Boolean

    KeyPreview = True
    
    flag1 = CheckPass("P", False) And CheckPass("7", False)
    flag2 = CheckPass("P", False) And (CheckPass("8", False) Or CheckPass("7", False))
    cmdClear.Visible = IIf(flag1, True, False)
    cmdCreateNew.Visible = IIf(flag2, True, False)
    cmdDelete.Visible = IIf(flag2, True, False)
    cmdSelect.Visible = IIf(flag2, True, False)
    
    dgCanisters.AllowRowSizing = False
    
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
            rsCrit = "SELECT * FROM [MasterCanister] ORDER BY [Number] ASC"
        Case 2
            rsCrit = "SELECT * FROM [MasterCanister] ORDER BY [Description] ASC"
        Case 3
            rsCrit = "SELECT * FROM [MasterCanister] ORDER BY [WorkingCapacity] DESC"
        Case 4
            rsCrit = "SELECT * FROM [MasterCanister] ORDER BY [WCVolume] DESC"
        Case Else
            rsCrit = "SELECT * FROM [MasterCanister] ORDER BY [Number] ASC"
    End Select
    adoCanisters.RecordSource = rsCrit
    adoCanisters.Refresh

    If adoCanisters.Recordset.BOF Then
        dgCanisters.Caption = " No Defined Canisters"
        ' Set column properties
        dgCanisters.Columns(0).Width = 760
        dgCanisters.Columns(1).Width = 1900
        dgCanisters.Columns(2).Width = 1500
        dgCanisters.Columns(3).Width = 1050
        cmdClear.Enabled = False
        cmdDelete.Enabled = False
        cmdSelect.Enabled = False
    Else
        ' Display number of canisters found
        adoCanisters.Recordset.GetRows
        Select Case adoCanisters.Recordset.RecordCount
            Case 0
                dgCanisters.Caption = " No Defined Canisters"
                cmdClear.Enabled = False
                cmdSelect.Enabled = False
                cmdSelect.Enabled = False
            Case 1
                dgCanisters.Caption = Format(adoCanisters.Recordset.RecordCount, "###0") & " Defined Canister"
                cmdClear.Enabled = False
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
            Case Else
                dgCanisters.Caption = Format(adoCanisters.Recordset.RecordCount, "###0") & " Defined Canisters"
                cmdClear.Enabled = True
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
        End Select
        
        ' Set column properties
        dgCanisters.Columns(0).Width = 760
        dgCanisters.Columns(1).Width = 1900
        dgCanisters.Columns(2).Width = 1500
        dgCanisters.Columns(3).Width = 1050
        
        ' move pointer to first row
        adoCanisters.Recordset.MoveFirst
        
        ' make the Left-Most column the Sorted-By column
        dgCanisters.LeftCol = IIf(sortCol > 3, 3, sortCol - 1)
        
    End If
    
End Sub

Public Sub SetInitialRow(rownum As Integer)
    InitRow = rownum
End Sub

Private Sub DeleteCan()
SetErrModule 258, 31
If UseLocalErrorHandler Then On Error GoTo localhandler
    If Not antiRepeatDelete Then
        If adoCanisters.Recordset.BOF Then
            searchCanMsgColor = MEDRED
            searchCanMsg = "No Canister Data Available"
        Else
        
            If IsNull(dgCanisters.Columns(0).CellValue(dgCanisters.GetBookmark(0))) Or IsEmpty(dgCanisters.Columns(0).CellValue(dgCanisters.GetBookmark(0))) Then
                ' Report an error
                searchCanMsgColor = MEDRED
                searchCanMsg = "Invalid Canister Number"
                Exit Sub
            End If
            
            adoCanisters.Recordset.Delete
            searchCanMsgColor = Message_ForeColor
            searchCanMsg = "Canister Deleted"
            antiRepeatDelete = True
           
        End If
    End If
Exit Sub
localhandler:
    searchCanMsgColor = MEDRED
    searchCanMsg = "Unable to Delete Canister"
End Sub

Private Sub ClearCan()
Dim iAux As Integer

    SetErrModule 258, 3
    If UseLocalErrorHandler Then On Error GoTo localhandler
    ' Clearing Canister
    searchCanMsgColor = Message_ForeColor
    searchCanMsg = "Clearing Canister.. Please Wait"
    adoCanisters.RecordSource = "SELECT * FROM [MasterCanister] with [Number] = " & (dgCanisters.Columns(0).CellValue(dgCanisters.GetBookmark(0))) & "  ORDER BY [Number] ASC"
'    adoCanisters.Recordset.MoveLast
    
    ' Blank Canister
    With adoCanisters.Recordset
        .Fields("Description").Value = "undefined"
        .Fields("WorkingCapacity").Value = CSng(0)
        .Fields("WCVolume").Value = CSng(0)
        .Update
    End With
    adoCanisters.RecordSource = "SELECT * FROM [MasterCanister] ORDER BY [Number] ASC"
    dgCanisters.Refresh
                    
    searchCanMsgColor = Message_ForeColor
    searchCanMsg = "Canister Cleared"
    
ResetErrModule
Exit Sub

localhandler:
    searchCanMsgColor = MEDRED
    searchCanMsg = "Unable to Clear Canister"
End Sub

Private Sub NewCan()
Dim iCan As Integer
Dim cannum As Integer

    cannum = 0
    For iCan = 1 To MAX_CANRCP
        If cannum = 0 Then
            If Not IsDefined(iCan, adoCanisters.Recordset) Then
                cannum = iCan
            End If
        End If
    Next iCan
    
    If cannum > 0 Then
        frmCanRecipe.Show
        frmCanRecipe.LoadNewCan CInt(cannum)
        Unload frmSearchCan
        Set frmSearchCan = Nothing
    Else
        searchCanMsgColor = MEDRED
        searchCanMsg = "No undefined Master Canister"
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
    lblMessage.ForeColor = searchCanMsgColor
    lblMessage.Caption = searchCanMsg
End Sub

