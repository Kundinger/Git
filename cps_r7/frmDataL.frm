VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmDataLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Log Form"
   ClientHeight    =   5310
   ClientLeft      =   255
   ClientTop       =   525
   ClientWidth     =   15330
   ClipControls    =   0   'False
   Icon            =   "frmDataL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
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
      Left            =   13320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDataL.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "View More"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More"
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
      Left            =   9600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDataL.frx":6824
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "View More"
      Top             =   5520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdLess 
      Caption         =   "Less"
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
      Left            =   10440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDataL.frx":7866
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "View Less"
      Top             =   5520
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Close"
      DisabledPicture =   "frmDataL.frx":88A8
      DownPicture     =   "frmDataL.frx":94EA
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
      Left            =   14340
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDataL.frx":A12C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      DisabledPicture =   "frmDataL.frx":AD6E
      DownPicture     =   "frmDataL.frx":B9B0
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
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDataL.frx":C5F2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Clear All Log Entries"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   600
      TabIndex        =   1
      Text            =   "No data to report."
      Top             =   6000
      Width           =   5415
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      DisabledPicture =   "frmDataL.frx":D234
      DownPicture     =   "frmDataL.frx":DE76
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
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDataL.frx":EAB8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Log"
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin MSDataGridLib.DataGrid dbgDataLog 
      Bindings        =   "frmDataL.frx":F6FA
      Height          =   3705
      Left            =   120
      TabIndex        =   3
      Top             =   510
      Width           =   15060
      _ExtentX        =   26564
      _ExtentY        =   6535
      _Version        =   393216
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Time"
         Caption         =   "Time"
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
         DataField       =   "Comment"
         Caption         =   "Comment"
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
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   12704.88
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoLogData 
      Height          =   345
      Left            =   120
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   609
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=CpsMaster"
      OLEDBString     =   "DSN=CpsMaster"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM [FileLog] ORDER BY [FileLog].[Time] DESC"
      Caption         =   "LogData"
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
   Begin Threed.SSCommand cmdAlarm 
      Height          =   840
      Left            =   6120
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1482
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   0   'False
      Picture         =   "frmDataL.frx":F713
   End
   Begin Threed.SSCommand cmdFileLog 
      Height          =   840
      Left            =   6960
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1482
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   0   'False
      Picture         =   "frmDataL.frx":14F05
   End
   Begin Threed.SSCommand cmdOOT 
      Height          =   840
      Left            =   7800
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1482
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   0   'False
      Picture         =   "frmDataL.frx":1A6F7
   End
   Begin Threed.SSCommand cmdJobLog 
      Height          =   840
      Left            =   8640
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   840
      _Version        =   65536
      _ExtentX        =   1482
      _ExtentY        =   1482
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Outline         =   0   'False
      Picture         =   "frmDataL.frx":1FEE9
   End
   Begin VB.Label lblDataLog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   9255
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "messages"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   4320
      Width           =   10905
   End
End
Attribute VB_Name = "frmDataLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error mod 55 ''''''''''''' Form DATALOG.frm '''''''''''''''''''
Option Explicit
Private xTime, xComment As Integer
Private ViewMode As Integer                  ' 1 = shorter list; 2 = longer list
Private Const viewLess = 1
Private Const viewMore = 2
Public LogData As String                     ' current user of the DataLog screen (& associated db controls)
Public LogJob As Long                        ' Job# for Alarm or OOT Log on DataLog screen
Public LogStn As Integer                     ' Station for Alarm or OOT Log on DataLog screen
Public LogShift As Integer                   ' Shift for Alarm or OOT Log on DataLog screen

Private Sub Print_DataLog()
' Procedure Name:   Print_DataLog
' Created By:       Brunrose
' Description:
'
Dim yThisLine, numPages, Idx As Integer
Dim oldFont As New StdFont
Dim sdate As String
Dim rs_time, rs_comment As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 55, 2

    ' column positions
    xTime = 40
    xComment = 2300
    ' datetime format
    sdate = "YYYY MMM DD  hh:mm:ss"
    ' number of pages
    numPages = (adoLogData.Recordset.RecordCount \ 50)
    If (adoLogData.Recordset.RecordCount Mod 50) > 0 Then numPages = numPages + 1

    ' Save current printer font
    oldFont = Printer.Font
    Printer.Font.Name = "Arial"
    
    ' TITLE, HEADER & COLUMN HEADINGS
    Print_Header
    'DATA
'    adoLogData.Recordset.MoveFirst
    If adoLogData.Recordset.RecordCount = 0 Then
        ' Print blank line(s)
        Print_Line ""
        Print_Line ""
        ' No Data
        Select Case LogData
            Case "Alarm"
                Print_Center "Alarm Log is Empty"
            Case "File"
                Print_Center "File Activity Log is Empty"
            Case "OOT"
                Print_Center "Out Of Tolerance Log is Empty"
        End Select
    Else
        For Idx = 1 To adoLogData.Recordset.RecordCount
            rs_time = IIf(adoLogData.Recordset("Time") <> "", _
                    Format(adoLogData.Recordset("Time"), sdate), " ")
            rs_comment = IIf(adoLogData.Recordset("comment") <> "", _
                    Mid$(adoLogData.Recordset("comment"), 1, 96), "          ")
            yThisLine = Printer.CurrentY
            Printer.CurrentX = xTime
            Printer.Print rs_time
            Printer.CurrentY = yThisLine
            Printer.CurrentX = xComment
            Printer.Print rs_comment
            ' more pages?
            If Idx = adoLogData.Recordset.RecordCount Or (Idx Mod 50) = 0 Then
                ' print footer
                Print_Footer numPages
                If Idx <> adoLogData.Recordset.RecordCount Then
                    ' new page
                    Printer.NewPage
                    ' TITLE, HEADER & COLUMN HEADINGS
                    Print_Header
                End If
            End If
            adoLogData.Recordset.MoveNext
        Next Idx
        adoLogData.Recordset.MoveFirst
    End If
    ' print footer
    Print_Footer numPages
    
    'DONE
    Printer.EndDoc
    Printer.Font = oldFont
    
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

Private Sub Print_Header()
' Procedure Name:   Print_Header
' Created By:       Brunrose
' Description:
'
Dim yThisLine As Integer
Dim sTitle As String
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 55, 22

    ' font
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    Printer.Font.Underline = False
    
    ' TITLE
    Select Case LogData
        Case "Alarm"
            sTitle = "ALARM LOG FOR JOB #" & Format(LogJob, "000000")
            sTitle = sTitle & " ON STATION " & Format(LogStn, "0")
            If NR_SHIFT > 1 Then sTitle = sTitle & "  SHIFT " & Format(LogShift, "0")
        Case "File"
            sTitle = "FILE ACTIVITY LOG"
        Case "OOT"
            sTitle = "OUT OF TOLERANCE LOG FOR JOB #" & Format(LogJob, "000000")
            sTitle = sTitle & " ON STATION " & Format(LogStn, "0")
            If NR_SHIFT > 1 Then sTitle = sTitle & "  SHIFT " & Format(LogShift, "0")
    End Select
    Print_Center (sTitle)
    ' Print blank line(s)
    Print_Line ""
    
    ' PAGE HEADER
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

    ' COLUMN HEADINGS
    yThisLine = Printer.CurrentY
    Printer.CurrentX = xTime
    Printer.Print "Time"
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xComment
    Printer.Print "Comment"
    Printer.Font.Underline = True
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xTime
    Printer.Print Space(38)
    Printer.CurrentY = yThisLine
    Printer.CurrentX = xComment
    Printer.Print Space(170)
    Printer.Font.Underline = False
    
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

Private Sub cmdPrint_Click()
Dim logName As String
'    If adoLogData.Recordset.RecordCount > 0 Then
        Print_DataLog
        Select Case LogData
            Case "Alarm"
                logName = "Alarm Log"
            Case "File"
                logName = "File Activity Log"
            Case "OOT"
                logName = "Out Of Tolerance Log"
        End Select
        lblMessage.Caption = logName & " Listing sent to" & vbCrLf & PRINTERNAME
'    Else
'        lblMessage.Caption = "nothing to print" & vbCrLf
'    End If
End Sub

Private Sub cmdReturn_Click()
    LogData = " "
    Unload Me
    Set frmDataLog = Nothing
End Sub

Private Sub cmdView_Click()
    ToggleView
End Sub

Private Sub Form_Activate()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 55, 0
    KeyPreview = True
    Form_Center Me
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmDataLog = Nothing
    End If
End Sub

Private Sub Form_Load()
Dim deltaWidth, deltaWidth2 As Single
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 55, 1
    KeyPreview = True
    lblDataLog.ForeColor = Titles_ForeColor
    With frmDataLog
        deltaWidth = 360
        deltaWidth2 = 640
'        .Width = frmStnDetail.Width
'        .dbgDataLog.Width = 15000
'        .dbgDataLog.Columns(0).Width = 1905
'        .dbgDataLog.Columns(1).Width = 13240
'        .cmdReturn.Left = .Width - deltaWidth2
        .cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
        .lblMessage.Font.Size = 9.5
        .lblMessage.ForeColor = DKPURPLE
        .lblMessage.Caption = ""
'        .lblMessage.Width = 10905
        ToggleView
    End With
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

Private Sub Form_Unload(Cancel As Integer)
    LogData = " "
    Unload Me
    Set frmDataLog = Nothing
End Sub

Private Sub ToggleView()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 55, 5525
Dim newView As Integer

    With frmDataLog
        Select Case ViewMode
            Case viewLess
                newView = viewMore
                .Height = 11295
                .dbgDataLog.Height = 9225
                .cmdClear.Top = 9840
                .cmdPrint.Top = 9840
                .cmdReturn.Top = 9840
                .cmdView.Top = 9840
                .cmdView.Picture = cmdLess.Picture
                .cmdView.ToolTipText = "View Less"
                .cmdView.Left = .cmdReturn.Left - 1110
             Case viewMore
                newView = viewLess
                .Height = 5775
                .dbgDataLog.Height = 3705
                .cmdClear.Top = 4320
                .cmdPrint.Top = 4320
                .cmdReturn.Top = 4320
                .cmdView.Top = 4320
                .cmdView.Picture = cmdMore.Picture
                .cmdView.ToolTipText = "View More"
                .cmdView.Left = .cmdReturn.Left - 1110
             Case Else
                newView = viewLess
                .Height = 5775
                .dbgDataLog.Height = 3705
                .cmdClear.Top = 4320
                .cmdPrint.Top = 4320
                .cmdReturn.Top = 4320
                .cmdView.Top = 4320
                .cmdView.Picture = cmdMore.Picture
                .cmdView.ToolTipText = "View More"
                .cmdView.Left = .cmdReturn.Left - 1110
         End Select
       ViewMode = newView
    End With

    frmDataLog.Refresh
    Form_Center Me

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


