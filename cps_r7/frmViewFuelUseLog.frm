VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmViewFuelUseLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fuel Use by Month"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14895
   Icon            =   "frmViewFuelUseLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   14895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   16400
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   18
      Top             =   3600
      Width           =   3015
   End
   Begin MSChart20Lib.MSChart FuelUseLogChart 
      Bindings        =   "frmViewFuelUseLog.frx":57E2
      Height          =   8775
      Left            =   3840
      OleObjectBlob   =   "frmViewFuelUseLog.frx":57FE
      TabIndex        =   3
      Top             =   1680
      Width           =   10935
   End
   Begin MSDataGridLib.DataGrid dgFuelUseLog 
      Bindings        =   "frmViewFuelUseLog.frx":724B
      Height          =   8775
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "DayOfMonth"
         Caption         =   "Year"
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
         DataField       =   "Month"
         Caption         =   "Month"
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
         DataField       =   "DayOfMonth"
         Caption         =   "Day"
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
         DataField       =   "ButaneTotal"
         Caption         =   "ButaneTotal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "FuelVaporTotal"
         Caption         =   "FuelVaporTotal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1200.189
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
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmViewFuelUseLog.frx":7267
         DownPicture     =   "frmViewFuelUseLog.frx":7EA9
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
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
         Picture         =   "frmViewFuelUseLog.frx":8AEB
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print Current Monthly Log"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdYearUp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewFuelUseLog.frx":972D
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1170
      End
      Begin VB.CommandButton cmdYearDn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewFuelUseLog.frx":EF0F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1035
         Width           =   1170
      End
      Begin VB.CommandButton cmdMonthUp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewFuelUseLog.frx":146F1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   1170
      End
      Begin VB.CommandButton cmdMonthDn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmViewFuelUseLog.frx":19ED3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1035
         Width           =   1170
      End
      Begin MSAdodcLib.Adodc adoFuelUseLog 
         Height          =   375
         Left            =   12120
         Top             =   120
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
         Connect         =   "DSN=cpsMaster"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "cpsMaster"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [FuelUseLog] "
         Caption         =   "FuelUseLog"
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
      Begin Threed.SSPanel pnlMonth 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Selected Month"
         Top             =   660
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "month"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin Threed.SSPanel pnlYear 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Selected Year"
         Top             =   660
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "year"
         ForeColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   1
         BevelInner      =   1
      End
      Begin VB.Label lblFuelVaporTotalUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   8775
         TabIndex        =   17
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lblButaneTotalUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   8775
         TabIndex        =   16
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lblFuelVaporTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "8,888,888"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   7560
         TabIndex        =   15
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label lblButaneTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "8,888,888"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   7560
         TabIndex        =   14
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lblLogName2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "for Month Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   480
         Left            =   3960
         TabIndex        =   13
         Top             =   570
         Width           =   6900
      End
      Begin VB.Label lblFuelVaporTotalDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Vapor Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   5700
         TabIndex        =   12
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label lblButaneTotalDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5700
         TabIndex        =   11
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label lblLogName1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Consumption Log"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   480
         Left            =   3960
         TabIndex        =   2
         Top             =   90
         Width           =   6900
      End
   End
End
Attribute VB_Name = "frmViewFuelUseLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 278'''''''''' Form frmViewFuelUseLog.frm '''''''''''''''''''
Option Explicit
Dim sPath As String
Dim rsCrit As String
Dim RowHgt As Single
Private Graph() As Single
Private rsFuelUseLog As New ADODB.Recordset
Private cn As New ADODB.Connection
Private selectedMonth As Integer
Private selectedYear As Integer

Sub RefreshChart()
'
Dim clr As Long
Dim idx As Long
Dim idx2 As Long
Dim tTxt As String
Dim sGrams As Single
Dim sScaleMax As Single
Dim incButane As Single
Dim incLiveFuel As Single
Dim totalButane As Single
Dim totalLiveFuel As Single
    
    If adoFuelUseLog.Recordset.RecordCount < 1 Then Exit Sub
    
    adoFuelUseLog.Recordset.MoveFirst
    
    ReDim Graph(31, 1 To 2)
    totalButane = 0
    totalLiveFuel = 0
    While Not adoFuelUseLog.Recordset.EOF
        idx = adoFuelUseLog.Recordset("DayOfMonth")
        incButane = adoFuelUseLog.Recordset("ButaneTotal")
        incLiveFuel = adoFuelUseLog.Recordset("FuelVaporTotal")
        totalButane = totalButane + incButane
        totalLiveFuel = totalLiveFuel + incLiveFuel
        Graph(idx, 1) = totalButane
        Graph(idx, 2) = totalLiveFuel
        If (Not adoFuelUseLog.Recordset.EOF) Then
            adoFuelUseLog.Recordset.MoveNext
        End If
    Wend
    
    lblButaneTotal.Caption = Format(totalButane, "#,###,##0")
    lblFuelVaporTotal.Caption = Format(totalLiveFuel, "#,###,##0")
    
    idx2 = 1
    sGrams = totalButane
    If (totalLiveFuel > totalButane) Then sGrams = totalLiveFuel
    Do While (sGrams > CSng(idx2))
        idx2 = idx2 * CLng(10)
    Loop
    sScaleMax = idx2
    
    totalButane = 0
    totalLiveFuel = 0
    For idx = 1 To 31
        If (Graph(idx, 1) = 0) Then Graph(idx, 1) = totalButane
        If (Graph(idx, 2) = 0) Then Graph(idx, 2) = totalLiveFuel
        totalButane = Graph(idx, 1)
        totalLiveFuel = Graph(idx, 2)
    Next idx
    
    FuelUseLogChart = Graph ' populate chart's data grid using Graph array
    
    ' Determine Y-scale Max
    ' Set the DataSource to the recordset.
    With FuelUseLogChart
'        .TitleText = tTxt
        .ShowLegend = False
        .Column = 1
        .ColumnLabel = "Butane"
        .Column = 2
        .ColumnLabel = "Fuel Vapor"
'        .Column = 3
'        .ColumnLabel = "WeightChange - 1"
       
        With FuelUseLogChart.Plot
        
            ' Y axis
            With .Axis(VtChAxisIdY).ValueScale
                .Auto = False
                .Minimum = 0
                .Maximum = sScaleMax
                .MajorDivision = 10
                .MinorDivision = 2
            End With
            
            ' X axis
            With .Axis(VtChAxisIdX).ValueScale
                .Auto = True
'                .Auto = False
'                .Minimum = minDay
'                .Maximum = maxDay
'                .MajorDivision = maxDay - minDay + 1
'                .MinorDivision = 1
            End With
            
            ' Weight Change Pen
            clr = DK2GREEN
            .SeriesCollection(1).Pen.Style = VtPenStyleSolid
            .SeriesCollection(1).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
            
            ' Weight Change Pen + 1
            clr = MEDORANGE
            .SeriesCollection(2).Pen.Style = VtPenStyleSolid
            .SeriesCollection(2).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
            
'            ' Weight Change Pen - 1
'            clr = DK2GREEN
'            .SeriesCollection(3).Pen.Style = VtPenStyleSolid
'            .SeriesCollection(3).Pen.VtColor.Set RedFromRGB(clr), GreenFromRGB(clr), BlueFromRGB(clr)
            
        End With
    
    End With
    
    FuelUseLogChart.Refresh
    
End Sub

Private Sub cmdMonthDn_Click()
    selectedMonth = IIf((selectedMonth = 1), 12, selectedMonth - 1)
    UpdateMoYr
End Sub

Private Sub cmdMonthUp_Click()
    selectedMonth = IIf((selectedMonth = 12), 1, selectedMonth + 1)
    UpdateMoYr
End Sub

Private Sub cmdPrint_Click()
    Set pbCapture.Picture = CaptureForm(Me)
    PrintPictureToFitPage Printer, pbCapture.Picture
    Printer.EndDoc
    Set pbCapture.Picture = Nothing
End Sub

Private Sub cmdYearDn_Click()
    selectedYear = selectedYear - 1
    UpdateMoYr
End Sub

Private Sub cmdYearUp_Click()
    selectedYear = selectedYear + 1
    UpdateMoYr
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmViewFuelUseLog = Nothing
    End If
End Sub

Private Sub Form_Load()
    SetErrModule 278, 2
    KeyPreview = True
    
    lblLogName1.ForeColor = TitlesLabel_ForeColor
    lblLogName2.ForeColor = TitlesLabel_ForeColor
    
    cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)

    selectedMonth = Month(Now)
    selectedYear = Year(Now)
    UpdateMoYr
    
    ResetErrModule
End Sub

Private Sub UpdateMoYr()
    pnlMonth.Caption = Format(selectedMonth, "00")
    pnlYear.Caption = Format(selectedYear, "0000")
    ConnectToDB
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub DisplayData()
'
Dim sText As String
Dim tmpDts As Date
    ' Select & Sort
    rsCrit = "SELECT * FROM [FuelUseLog] WHERE [Year] = " & selectedYear & "  and [Month] = " & selectedMonth & " "
    rsCrit = rsCrit & " ORDER BY [DayOfMonth] ASC"
    adoFuelUseLog.RecordSource = rsCrit
    adoFuelUseLog.Refresh
    
    tmpDts = DateSerial(selectedYear, selectedMonth, CInt(1))
    lblLogName2.Caption = "for " + Format(tmpDts, "MMMM YYYY")
    lblButaneTotal.Caption = " "
    lblFuelVaporTotal.Caption = " "
    
    ' Display number of FuelUseLog entries found
'    If adoFuelUseLog.Recordset.BOF Then
'        dgFuelUseLog.Caption = " No Records "
'    Else
'        adoFuelUseLog.Recordset.MoveFirst
'        adoFuelUseLog.Recordset.MoveLast
        ' get number of records
'        adoFuelUseLog.Recordset.GetRows
'        Select Case adoFuelUseLog.Recordset.RecordCount
'            Case 0
'                dgFuelUseLog.Caption = " No Records "
'            Case 1
'                dgFuelUseLog.Caption = Format(adoFuelUseLog.Recordset.RecordCount, "###0") & " FuelUse Log Record"
'            Case Else
'                dgFuelUseLog.Caption = Format(adoFuelUseLog.Recordset.RecordCount, "###,##0") & " FuelUse Log Records"
'        End Select
        ' move pointer to first row
'        adoFuelUseLog.Recordset.MoveFirst
'        dgFuelUseLog.Scroll 0, (adoFuelUseLog.Recordset.RecordCount - 5)
'        dgFuelUseLog.SelStart
'        dgFuelUseLog.Refresh
'    End If
    
    
End Sub

Sub ConnectToDB()
    adoFuelUseLog.Refresh
    DisplayData
    RefreshChart
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmViewAirLog
    Set frmViewAirLog = Nothing
End Sub

