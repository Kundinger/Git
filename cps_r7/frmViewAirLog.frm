VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmViewAirLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AirLog"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14895
   Icon            =   "frmViewAirLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   14895
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart AirLogChart 
      Bindings        =   "frmViewAirLog.frx":57E2
      Height          =   8760
      Left            =   14760
      OleObjectBlob   =   "frmViewAirLog.frx":57FA
      TabIndex        =   8
      Top             =   2280
      Width           =   14685
   End
   Begin MSDataGridLib.DataGrid dgAirLog 
      Bindings        =   "frmViewAirLog.frx":75C2
      Height          =   8775
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   15478
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ForeColor       =   12615680
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
      Caption         =   "AirLog"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "DTS"
         Caption         =   "DTS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "YYYY-MMM-DD  HH:MM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Barometer"
         Caption         =   "Barometer"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Temperature"
         Caption         =   "Temperature"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Humidity"
         Caption         =   "Humidity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Moisture"
         Caption         =   "Moisture"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TemperatureOOT"
         Caption         =   "TemperatureOOT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "OOT"
            FalseValue      =   "ok"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "MoistureOOT"
         Caption         =   "MoistureOOT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "OOT"
            FalseValue      =   "ok"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column07 
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
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4800.189
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pbxControls 
      Height          =   1680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      _Version        =   65536
      _ExtentX        =   26273
      _ExtentY        =   2963
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
      Begin VB.CommandButton cmdChart 
         Caption         =   " Display Chart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   13920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewAirLog.frx":75DA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.TextBox txtFileName 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   5760
         TabIndex        =   5
         Text            =   "none selected"
         ToolTipText     =   "Selected AirLog DB File File Name"
         Top             =   720
         Width           =   4320
      End
      Begin VB.FileListBox flbAirLogFiles 
         Height          =   1065
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "List of AirLog DataBase Files"
         Top             =   480
         Width           =   3600
      End
      Begin VB.CommandButton cmdCurrent 
         Caption         =   " Display Current"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   10560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewAirLog.frx":821C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   " Display Selection"
         DisabledPicture =   "frmViewAirLog.frx":8E5E
         DownPicture     =   "frmViewAirLog.frx":9AA0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewAirLog.frx":A6E2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin MSAdodcLib.Adodc adoAirLog 
         Height          =   375
         Left            =   6600
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         RecordSource    =   "SELECT * FROM [MasterRecipe] "
         Caption         =   "AirLog"
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
      Begin VB.Label lblAirLogFiles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AirLog DB Files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   3600
      End
      Begin VB.Label lblFilename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AirLog File Currently Displayed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Width           =   4320
      End
   End
End
Attribute VB_Name = "frmViewAirLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 278'''''''''' Form SearchCan.frm '''''''''''''''''''
Option Explicit
Dim sPath As String
Dim rsCrit As String
Dim RowHgt As Single
Private DisplayChart As Boolean
Private DspAirLogFile As String
Private Graph() As Single
Private rsAirLog As New ADODB.Recordset
Private cn As New ADODB.Connection

Private Sub cmdChart_Click()
    Select Case DisplayChart
        Case True
            cmdChart.Caption = "Display Chart"
            DisplayChart = False
            AirLogChart.Left = OutOfSight
        Case False
            cmdChart.Caption = "Hide Chart"
            DisplayChart = True
            AirLogChart.Left = dgAirLog.Left
            AirLogChart.Top = dgAirLog.Top + 15
            ConnectToDB DspAirLogFile
    End Select
End Sub

Private Sub cmdCurrent_Click()
Dim idx As Integer
    For idx = 0 To flbAirLogFiles.ListCount - 1
        If (flbAirLogFiles.List(idx) = CurAirLogFile) Then
            flbAirLogFiles.Selected(idx) = True
        Else
            flbAirLogFiles.Selected(idx) = False
        End If
    Next idx
    ConnectToDB CurAirLogFile
End Sub

Sub RefreshChart()
Dim clr As Long
Dim idx As Long
Dim maxDay As Integer
Dim minDay As Integer
Dim maxDTS As Date
Dim minDTS As Date
Dim tTxt As String
Dim strQuery As String ' SQL query string.

    ' First change the path to a valid path for your machine.
    cn.ConnectionString = _
         "Provider=Microsoft.Jet.OLEDB.4.0;" _
         & "Data Source=" & FILEPATH_log & DspAirLogFile & ";" _
         & "Persist Security Info=False"
    
    ' Open the connection.
    cn.Open
    
    ' Create a query that retrieves only four fields.
    strQuery = "SELECT DTS, Temperature, Moisture, Humidity FROM Air_Log ORDER BY [DTS] DESC"
    ' Open the recordset.
    rsAirLog.Open strQuery, cn, adOpenKeyset
    
    If rsAirLog.RecordCount < 1 Then Exit Sub
    
    rsAirLog.MoveFirst
    
    ReDim Graph(rsAirLog.RecordCount, 1 To 3)
    idx = rsAirLog.RecordCount
    maxDay = 1
    minDay = 31
    maxDTS = 0
    minDTS = Now + 1
    While Not rsAirLog.EOF
        Graph(idx, 1) = rsAirLog("Temperature")
        Graph(idx, 2) = rsAirLog("Moisture")
        Graph(idx, 3) = rsAirLog("Humidity")
        If (Day(rsAirLog("DTS")) > maxDay) Then maxDay = Day(rsAirLog("DTS"))
        If (Day(rsAirLog("DTS")) < minDay) Then minDay = Day(rsAirLog("DTS"))
        If (rsAirLog("DTS") > maxDTS) Then maxDTS = rsAirLog("DTS")
        If (rsAirLog("DTS") < minDTS) Then minDTS = rsAirLog("DTS")
        If (Not rsAirLog.EOF) Then
            rsAirLog.MoveNext
            idx = idx - 1
        End If
    Wend
    
'    tTxt = NameOfMonth(CInt(Mid(DspAirLogFile, 12, 2))) + " "
'    tTxt = tTxt + Mid(DspAirLogFile, 8, 4)
'    tTxt = "From "
    tTxt = Format(minDTS, "YYYY MMM D  hh:mm") + "   to   "
    tTxt = tTxt + Format(maxDTS, "YYYY MMM D  hh:mm")
    
    AirLogChart = Graph ' populate chart's data grid using Graph array
    
    
    ' Set the DataSource to the recordset.
    With AirLogChart
        .TitleText = tTxt
        .ShowLegend = True
        .Column = 1
        .ColumnLabel = "Temperature"
        .Column = 2
        .ColumnLabel = "Moisture"
        .Column = 3
        .ColumnLabel = "Humidity"
       
        With AirLogChart.Plot
        
            ' Y axis
            With .Axis(VtChAxisIdY).ValueScale
                .Auto = False
                .Minimum = 0
                .Maximum = 100
                .MajorDivision = 10
                .MinorDivision = 2
            End With
            
            ' X axis
            With .Axis(VtChAxisIdX).ValueScale
'                .Auto = True
                .Auto = False
                .Minimum = minDay
                .Maximum = maxDay
                .MajorDivision = maxDay - minDay + 1
                .MinorDivision = 1
            End With
            
            ' Temperature Pen
            clr = MEDRED
            .SeriesCollection(1).Pen.Style = VtPenStyleSolid
            .SeriesCollection(1).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
            
            ' Moisture Pen
            clr = MEDBLUE
            .SeriesCollection(2).Pen.Style = VtPenStyleSolid
            .SeriesCollection(2).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
            
            ' Humidity Pen
            clr = DK2GREEN
            .SeriesCollection(3).Pen.Style = VtPenStyleSolid
            .SeriesCollection(3).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
            
        End With
    
    End With
    
    AirLogChart.Refresh
    
    rsAirLog.Close
    cn.Close
    
End Sub

Private Sub dgAirLog_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex)
End Sub

Private Sub cmdSelect_Click()
Dim idx As Integer
    For idx = 0 To flbAirLogFiles.ListCount - 1
        If flbAirLogFiles.Selected(idx) Then
            ConnectToDB flbAirLogFiles.List(idx)
        End If
    Next idx
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmViewAirLog = Nothing
    End If
End Sub

Private Sub Form_Load()
    SetErrModule 278, 2
    KeyPreview = True
    
'    frmViewAirLog.adoAirLog.Enabled = True
'    frmViewAirLog.AirLogChart.Enabled = True
'    frmViewAirLog.dgAirLog.Enabled = True
    
    cmdChart.Caption = "Display Chart"
    DisplayChart = False
    AirLogChart.Left = OutOfSight
    
    lblFilename.ForeColor = TitlesLabel_ForeColor
    txtFileName.ForeColor = Data_ForeColor
    flbAirLogFiles.Path = FILEPATH_log
    flbAirLogFiles.Pattern = AccessDbFileExt
    
    cmdCurrent.Visible = True
    cmdSelect.Visible = True
    cmdChart.Visible = True
    
    ConnectToDB CurAirLogFile
    
    ResetErrModule
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub DisplayData(sortCol As Integer)
    ' Select & Sort
    Select Case sortCol
        Case 0
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [DTS] DESC"
        Case 1
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [Barometer] DESC"
        Case 2
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [Temperature] DESC"
        Case 3
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [Humidity] DESC"
        Case 4
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [Moisture] DESC"
        Case 5
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [TemperatureOOT] ASC"
        Case 6
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [MoistureOOT] ASC"
        Case 7
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [Comment] DESC"
        Case Else
            rsCrit = "SELECT * FROM [Air_Log] ORDER BY [DTS] DESC"
    End Select
    adoAirLog.RecordSource = rsCrit
    adoAirLog.Refresh

    ' Display number of AirLog found
    adoAirLog.Recordset.GetRows
    Select Case adoAirLog.Recordset.RecordCount
        Case 0
            dgAirLog.Caption = " No Records "
        Case 1
            dgAirLog.Caption = Format(adoAirLog.Recordset.RecordCount, "###0") & " AirLog Record"
        Case Else
            dgAirLog.Caption = Format(adoAirLog.Recordset.RecordCount, "###,##0") & " AirLog Records"
    End Select
    
    ' move pointer to first row
    adoAirLog.Recordset.MoveFirst
    
End Sub

Sub ConnectToDB(ByVal AirLogFile As String)

    txtFileName.text = AirLogFile
    DspAirLogFile = AirLogFile
    
    adoAirLog.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & FILEPATH_log & DspAirLogFile & ";" _
        & "Persist Security Info=False"
'    adoAirLog.RecordSource = "SELECT * FROM [Air_Log] ORDER BY [DTS]"
'    adoAirLog.Refresh
    
    DisplayData 0
    
    RefreshChart

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmViewAirLog
    Set frmViewAirLog = Nothing
End Sub

