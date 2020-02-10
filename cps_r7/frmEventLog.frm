VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmEventLog 
   Caption         =   "Event Log "
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   Icon            =   "frmEventLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmEventLog.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "View More"
      Top             =   5280
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
      Left            =   11880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmEventLog.frx":6824
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "View More"
      Top             =   2040
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
      Left            =   12720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmEventLog.frx":7866
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "View Less"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdToggleEvents 
      Caption         =   "Toggle System/Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Log"
      Top             =   127
      UseMaskColor    =   -1  'True
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid dbgEventLog 
      Bindings        =   "frmEventLog.frx":88A8
      Height          =   4515
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   7964
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoEventData 
      Height          =   345
      Left            =   6840
      Top             =   120
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
      RecordSource    =   "SELECT * FROM [EventLog] ORDER BY [EventLog].[Time] DESC"
      Caption         =   "EventData"
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      DisabledPicture =   "frmEventLog.frx":88C3
      DownPicture     =   "frmEventLog.frx":9505
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
      Picture         =   "frmEventLog.frx":A147
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Log"
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Close"
      DisabledPicture =   "frmEventLog.frx":AD89
      DownPicture     =   "frmEventLog.frx":B9CB
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
      Left            =   8525
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmEventLog.frx":C60D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin MSAdodcLib.Adodc adoReportEvents 
      Height          =   345
      Left            =   0
      Top             =   0
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
      RecordSource    =   "SELECT * FROM [ReportsEventLog] ORDER BY [ReportsEventLog].[Time] DESC"
      Caption         =   "ReportEvents"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   5280
      Width           =   6210
   End
   Begin VB.Label lblEventLog 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System Events Log"
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
      TabIndex        =   2
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 58'''''''''' Form EventLog.frm '''''''''''''''''''
Option Explicit
Private xTime As Integer
Private xComment As Integer
Private ViewMode As Integer                  ' 1 = shorter list; 2 = longer list
Private Const viewLess = 1
Private Const viewMore = 2

Sub Print_EventLog()
'
' Procedure Name:   Print_EventLog
' Created By:       Brunrose
' Description:
' This procedure prints the EventLog File
'
Dim yThisLine As Integer
Dim numPages As Integer
Dim Idx As Integer
Dim oldFont As New StdFont
Dim sdate As String
Dim rs_time As String
Dim rs_comment As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 58, 2

    ' column positions
    xTime = 40
    xComment = 2300
    ' datetime format
    sdate = "YYYY MMM DD  hh:mm:ss"
    ' number of pages
    numPages = (adoEventData.Recordset.RecordCount \ 50)
    If (adoEventData.Recordset.RecordCount Mod 50) > 0 Then numPages = numPages + 1

    ' Save current printer font
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    oldFont = Printer.Font
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Name = "Arial"
    
    ' TITLE, HEADER & COLUMN HEADINGS
    Print_Header
    'DATA
    adoEventData.Recordset.MoveFirst
    If adoEventData.Recordset.RecordCount = 0 Then
        ' Print blank line(s)
        Print_Line ""
        Print_Line ""
        ' No Data
        Print_Center "System Event Log is Empty"
    Else
        For Idx = 1 To adoEventData.Recordset.RecordCount
            rs_time = IIf(adoEventData.Recordset("Time") <> "", _
                    Format(adoEventData.Recordset("Time"), sdate), " ")
            rs_comment = IIf(adoEventData.Recordset("comment") <> "", _
                    Mid$(adoEventData.Recordset("comment"), 1, 96), "          ")
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            yThisLine = Printer.CurrentY
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            Printer.CurrentX = xTime
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            Printer.Print rs_time
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            Printer.CurrentY = yThisLine
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            Printer.CurrentX = xComment
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
            Printer.Print rs_comment
            ' more pages?
            If Idx = adoEventData.Recordset.RecordCount Or (Idx Mod 50) = 0 Then
                ' print footer
                Print_Footer numPages
                If Idx <> adoEventData.Recordset.RecordCount Then
                    ' new page
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
                    Printer.NewPage
                    ' TITLE, HEADER & COLUMN HEADINGS
                    Print_Header
                End If
            End If
            adoEventData.Recordset.MoveNext
        Next Idx
        adoEventData.Recordset.MoveFirst
    End If
    ' print footer
    Print_Footer numPages
    
    'DONE
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.EndDoc
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
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
SetErrModule 58, 22

    ' font
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Size = 12
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Bold = False
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Italic = False
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Underline = False
    
    ' TITLE
    sTitle = "SYSTEM EVENT LOG"
    Print_Center (sTitle)
    ' Print blank line(s)
    Print_Line ""
    
    ' PAGE HEADER
    Print_Center "Canister Preconditioning System"
    Print_Center Trim$(SysConfig.Heading)
    Print_Center Trim$(SysConfig.Heading2)
    Print_Center (Format(Now, "d mmmm yyyy"))
    ' reset font
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Size = 10
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Bold = False
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Italic = False
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Underline = False
    ' Print blank line(s)
    Print_Line ""
'    Print_Line ""

    ' COLUMN HEADINGS
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    yThisLine = Printer.CurrentY
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentX = xTime
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Print "Time"
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentY = yThisLine
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentX = xComment
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Print "Comment"
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Font.Underline = True
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentY = yThisLine
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentX = xTime
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Print Space(38)
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentY = yThisLine
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.CurrentX = xComment
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
    Printer.Print Space(170)
'FIXIT: Printer object and Printers collection not upgraded to Visual Basic .NET by the Upgrade Wizard.     FixIT90210ae-R5481-H1984
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
    Print_EventLog
    lblMessage.Caption = "Event Log Listing sent to" & vbCrLf & PRINTERNAME
End Sub

Private Sub cmdReturn_Click()
    Unload frmEventLog
    Set frmEventLog = Nothing
End Sub

Private Sub cmdToggleEvents_Click()
    If (lblEventLog.Caption = "Reports Event Log") Then
        ' show System Events
        lblEventLog.Caption = "System Events Log"
        Set dbgEventLog.DataSource = adoEventData
        adoEventData.Refresh
    Else
        ' show Report Events
        lblEventLog.Caption = "Reports Event Log"
        Set dbgEventLog.DataSource = adoReportEvents
        adoReportEvents.Refresh
    End If
End Sub

Private Sub cmdView_Click()
    ToggleView
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmEventLog = Nothing
    End If
End Sub

Private Sub Form_Load()
Dim deltaWidth As Single
Dim deltaWidth2 As Single
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 58, 1
    UnreadProgramErrorMessage = False
    KeyPreview = True
    cmdToggleEvents.Visible = IIf(CheckPass("F", False), True, False)
    ' Set Title Foreground color
    lblEventLog.ForeColor = Titles_ForeColor
    frmEventLog.adoEventData.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & FILEPATH_sysdbf & DATAMASTER & ";" _
        & "Persist Security Info=False"
    adoEventData.Refresh
    With frmEventLog
        deltaWidth = 360
        deltaWidth2 = 1080
'        .Width = .Width + deltaWidth2
        .Width = frmStnDetail.Width
        .dbgEventLog.Width = .Width - deltaWidth
        .dbgEventLog.Columns(0).Width = 2000    '1.15 * .dbgEventLog.Columns(0).Width
        .dbgEventLog.Columns(1).Width = 12000   '3.5 * .dbgEventLog.Columns(1).Width
        .cmdReturn.Left = .Width - deltaWidth2
        .lblEventLog.Left = .dbgEventLog.Left
        .lblEventLog.Width = .dbgEventLog.Width
        .lblEventLog.Caption = "System Events Log"
        .cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
        .lblMessage.Font.Size = 9.5
        .lblMessage.ForeColor = DKPURPLE
        .lblMessage.Caption = ""
        ToggleView
    End With
    frmEventLog.dbgEventLog.Refresh
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
    HotKeyCheck KeyCode, Shift  ' undo rest to display key codes
End Sub

Private Sub lblEventLog_DblClick()
    adoEventData.Refresh
End Sub

Private Sub ToggleView()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 58, 5525
Dim newView As Integer

    With frmEventLog
        Select Case ViewMode
            Case viewLess
                newView = viewMore
                .Height = 11295
                .dbgEventLog.Height = 9195
                .cmdPrint.Top = 9900
                .cmdReturn.Top = 9900
                .cmdView.Top = 9900
                .cmdView.Picture = cmdLess.Picture
                .cmdView.ToolTipText = "View Less"
                .cmdView.Left = .cmdReturn.Left - 1110
             Case viewMore
                newView = viewLess
                .Height = 5775
                .dbgEventLog.Height = 3705
                .cmdPrint.Top = 4380
                .cmdReturn.Top = 4380
                .cmdView.Top = 4380
                .cmdView.Picture = cmdMore.Picture
                .cmdView.ToolTipText = "View More"
                .cmdView.Left = .cmdReturn.Left - 1110
             Case Else
                newView = viewLess
                .Height = 5775
                .dbgEventLog.Height = 3705
                .cmdPrint.Top = 4380
                .cmdReturn.Top = 4380
                .cmdView.Top = 4380
                .cmdView.Picture = cmdMore.Picture
                .cmdView.ToolTipText = "View More"
                .cmdView.Left = .cmdReturn.Left - 1110
         End Select
       ViewMode = newView
    End With

    frmEventLog.Refresh
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


