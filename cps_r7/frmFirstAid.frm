VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmFirstAid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "First Aid"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   Icon            =   "frmFirstAid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmControls 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkSetupFiles 
         Caption         =   "Include System Setup files     (Recommended)"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4215
      End
      Begin VB.CheckBox chkJobFiles 
         Caption         =   "Include files for Selected Job (Recommended)"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         DisabledPicture =   "frmFirstAid.frx":57E2
         DownPicture     =   "frmFirstAid.frx":5EE4
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFirstAid.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton cmdFirstAid 
         Caption         =   "FirstAid Save"
         DisabledPicture =   "frmFirstAid.frx":6CE8
         DownPicture     =   "frmFirstAid.frx":73EA
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFirstAid.frx":7AEC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.Label lblMessages 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "messages"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame frmJobs 
      Caption         =   "Jobs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   10815
      Begin MSDataGridLib.DataGrid dbgJoblist 
         Bindings        =   "frmFirstAid.frx":81EE
         Height          =   5205
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   9181
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         ForeColor       =   -2147483635
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "List of Jobs"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Job Number"
            Caption         =   "Job Number"
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
            Caption         =   "Job Description"
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
            DataField       =   "Vehicle"
            Caption         =   "Vehicle"
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
            DataField       =   "Start Time"
            Caption         =   "Start Time"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "YYYY MMM DD   HH:MM:SS"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Stop Time"
            Caption         =   "Stop Time"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "YYYY MMM DD   HH:MM:SS"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Station"
            Caption         =   "Station"
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
            DataField       =   "Shift"
            Caption         =   "Shift"
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
            DataField       =   "Report Filename"
            Caption         =   "Report Filename"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3465.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2670.236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2564.788
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2564.788
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   4470.236
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoJoblist 
         Height          =   375
         Left            =   120
         Top             =   5280
         Visible         =   0   'False
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
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
         Connect         =   "DSN=CpsMaster"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "CpsMaster"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM [Joblist] ORDER BY [Job Number] DESC"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmFirstAid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                                            frmFirstAid    module 911
Option Explicit
Private daodb36 As DAO.Database
Private rS As DAO.Recordset
Private sourcename, destname As String
Private sPath, rptPath, rptName As String
Private rsCrit As String
Private FirstAidDrive As String
Private FirstAidPath As String
Private FirstAidMsg As String
Private FirstAidMsgColor As Long
Private dbPath, DBFile As String
Private idx, iStn, iShift As Integer

Private Sub chkJobFiles_Click()
    chkJobFiles.ForeColor = IIf(chkJobFiles = cNO, MEDRED, MEDBLUE)
    UpdateScreen
End Sub

Private Sub chkSetupFiles_Click()
    chkSetupFiles.ForeColor = IIf(chkSetupFiles = cNO, MEDRED, MEDBLUE)
    UpdateScreen
End Sub

Private Sub cmdFirstAid_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 911, 911
Dim idx As Integer
Dim iRow As Integer
Dim selRows As Integer
Dim iFileNumber As Integer
Dim sFileName As String
Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    FirstAidPath = filepath & "\backup\first_aid\"
    
    lblMessages.ForeColor = DKPURPLE
    lblMessages.Caption = vbCrLf
    lblMessages.Caption = lblMessages.Caption & "***** PREPARING FIRSTAID ZIP FILE *****" & vbCrLf
    
    ' Clear Folders for FirstAid
    ClearFolder FirstAidPath & "calibrate\"
    ClearFolder FirstAidPath & "config\"
    ClearFolder FirstAidPath & "data"
    ClearFolder FirstAidPath & "recipes\"
    ClearFolder FirstAidPath & "reports\"
    ClearFolder FirstAidPath & "sysdbf\"
    
    lblMessages.ForeColor = DKPURPLE
    lblMessages.Caption = vbCrLf
    lblMessages.Caption = lblMessages.Caption & "***** CREATING FIRSTAID ZIP FILE *****" & vbCrLf
    
    FirstAidMsg = ""
    If adoJoblist.Recordset.BOF Then
        lblMessages.ForeColor = MEDRED
        lblMessages.Caption = vbCrLf & "No Job Data Available"
    Else
    
        If IsNull(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            lblMessages.ForeColor = MEDRED
            lblMessages.Caption = vbCrLf & "Invalid Job Number"
            Exit Sub
        Else
            DBFile = CStr(dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(0)))
            For iStn = 1 To LAST_STN
              For iShift = 1 To NR_SHIFT
                If (StationControl(iStn, iShift).Job_Number) = DBFile Then
                  ' Report an error
                  lblMessages.ForeColor = MEDRED
                  lblMessages.Caption = "That Job#" & DBFile & "  has not been completed yet by Station " & iStn & " Shift " & iShift
                  Exit Sub
                End If
              Next iShift
            Next iStn
        End If
        
        If IsNull(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Or IsEmpty(dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(0))) Then
            ' Report an error
            lblMessages.ForeColor = MEDRED
            lblMessages.Caption = "Invalid filename"
            Exit Sub
        End If
        
    
        If fs.FolderExists(FirstAidPath) Then
            MousePointer = vbHourglass
            
            If chkJobFiles.Value = cYES Then
                selRows = dbgJoblist.SelBookmarks.count
                For idx = 0 To (selRows - 1)
                    iRow = dbgJoblist.SelBookmarks(idx)
                    ' db file(s)
                    DBFile = dbgJoblist.Columns(0).CellValue(dbgJoblist.GetBookmark(iRow - dbgJoblist.SelBookmarks(0)))
                    sourcename = FILEPATH_data & "C" & DBFile & AccessDbFileExt
                    destname = FirstAidPath & "data\" & "C" & DBFile & ".dmb"
                    If fs.FileExists(sourcename) Then FileCopy sourcename, destname
                    ' Report Files
                    ' report filename is column 7
                    rptName = dbgJoblist.Columns(7).CellValue(dbgJoblist.GetBookmark(iRow - dbgJoblist.SelBookmarks(0)))
                    ' summary report
                    sourcename = FILEPATH_reports & rptName & "Summary.RPT"
                    destname = FirstAidPath & "reports\" & Left(rptName, 50) & "Summary.RPT"
                    If fs.FileExists(sourcename) Then FileCopy sourcename, destname
                    ' detail report
                    sourcename = FILEPATH_reports & rptName & "Detail.RPT"
                    destname = FirstAidPath & "reports\" & Left(rptName, 50) & "Detail.RPT"
                    If fs.FileExists(sourcename) Then FileCopy sourcename, destname
                Next idx
            End If
            
            If chkSetupFiles.Value = cYES Then
                Dim sTxt As String
                'configuration files
                Shell FirstAidPath & "CopyCfgFiles.bat"
                'calibration files
                Shell FirstAidPath & "CopyCalFiles.bat"
                ' recipe file
                sourcename = FILEPATH_rcp & "cpsrecipes_rev" & Format(DBFREVLVL, "0#") & AccessDbFileExt
                destname = FirstAidPath & "recipes\cpsrecipes_rev" & Format(DBFREVLVL, "0#") & ".dmb"
                If fs.FileExists(sourcename) Then FileCopy sourcename, destname
                ' model db file
                sourcename = FILEPATH_sysdbf & "cpsmodel_rev" & Format(DBFREVLVL, "0#") & AccessDbFileExt
                destname = FirstAidPath & "sysdbf\cpsmodel_rev" & Format(DBFREVLVL, "0#") & ".dmb"
                If fs.FileExists(sourcename) Then FileCopy sourcename, destname
                ' master db file
'                sourcename = FILEPATH_sysdbf & "cpsmaster_rev" & Format(DBFREVLVL, "0#") & AccessDbFileExt
'                destname = FirstAidPath & "sysdbf\cpsmaster_rev" & Format(DBFREVLVL, "0#") & ".dmb"
'                FileCopy sourcename, destname
                ' sysdef db file
                sourcename = FILEPATH_sysdbf & "cpsSysDef_rev" & Format(DBFREVLVL, "0#") & AccessDbFileExt
                destname = FirstAidPath & "sysdbf\cpsSysDef_rev" & Format(DBFREVLVL, "0#") & ".dmb"
                If fs.FileExists(sourcename) Then FileCopy sourcename, destname
                ' user db file
'                sourcename = FILEPATH_sysdbf & "cpsuser_rev" & Format(DBFREVLVL, "0#") & AccessDbFileExt
'                destname = FirstAidPath & "sysdbf\cpsuser_rev" & Format(DBFREVLVL, "0#") & ".dmb"
'                FileCopy sourcename, destname
                ' zLog db file
                sourcename = FILEPATH_sysdbf & "cpsZlog_rev" & Format(DBFREVLVL, "0#") & AccessDbFileExt
                destname = FirstAidPath & "sysdbf\cpsZlog_rev" & Format(DBFREVLVL, "0#") & ".dmb"
                If fs.FileExists(sourcename) Then FileCopy sourcename, destname
            End If
            
'            ' create bat file to switch to the right drive
'            sFileName = FirstAidPath & "SwitchDrive.bat"
'            iFileNumber = FreeFile
            
'            Open sFileName For Output As #iFileNumber
'            Write #iFileNumber, Chr(Asc(DRIVEPATH))
'            Close #iFileNumber
            
            ' zip files
            destname = filepath & "\" & "FirstAid_" & Format(Now, "YYYY_MM_DD_hh_mm_ss") & ".zip"
            Shell FirstAidPath & "ZipFiles.bat " & destname
            
            ' wait for 7zip to finish
            Delay_Box "none", 5000, msgNOSHOW
                    
            ' Backup Complete
            lblMessages.ForeColor = DKPURPLE
            lblMessages.Caption = vbCrLf & "FirstAid File Save Complete"
            
        Else
            ' backup path doesn't exist
            lblMessages.Caption = "Backup Path >" & FirstAidPath & "< Not defined" & vbCrLf
            lblMessages.Caption = lblMessages.Caption & "ABORTING Backup"
            
        End If
            
                
        DoEvents
        MousePointer = vbDefault
            
    End If
    
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

Private Sub cmdQuit_Click()
    Unload Me
    Set frmFirstAid = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmFirstAid = Nothing
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key codes
End Sub

Private Sub Form_Load()
    
    KeyPreview = True
    
    lblMessages.ForeColor = MEDRED
    lblMessages.FontBold = True
    lblMessages.Caption = vbCrLf
    lblMessages.Caption = lblMessages.Caption & "***** CREATE FIRSTAID FILE FOR APS *****" & vbCrLf
    lblMessages.Caption = lblMessages.Caption & vbCrLf
    lblMessages.Caption = lblMessages.Caption & "Select the Job files to be included" & vbCrLf
'    lblMessages.Caption = lblMessages.Caption & "      OR Select   DO Not Include Any Job Files" & vbCrLf
    lblMessages.Caption = lblMessages.Caption & vbCrLf
    lblMessages.Caption = lblMessages.Caption & "Then press" & vbCrLf
    lblMessages.Caption = lblMessages.Caption & vbCrLf
    lblMessages.Caption = lblMessages.Caption & "FirstAid Save" & vbCrLf
'    lblMessages.Caption = lblMessages.Caption & vbCrLf
'    lblMessages.Caption = lblMessages.Caption & "button"
    chkJobFiles.Value = 1
    chkSetupFiles.Value = 1
   
    adoJoblist.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & FILEPATH_sysdbf & DATAMASTER & ";" _
        & "Persist Security Info=False"
    adoJoblist.Refresh
    
    UpdateScreen

End Sub

Private Sub UpdateScreen()
    cmdFirstAid.Enabled = IIf(chkJobFiles = 1 Or chkSetupFiles = 1, True, False)
    cmdQuit.Enabled = True
End Sub

Private Sub ClearFolder(ByVal sFolder As String)

Dim sPath As String
Dim sDir As String
Dim sFile As String

    sPath = IIf((Mid(sFolder, Len(sFolder), 1) = "\"), sFolder, sFolder & "\")
    
    sDir = sPath & "*.*"
    sFile = Dir(sDir)

    ' delete all files in the directory
    Do While sFile <> ""
        Kill sPath & sFile
        sFile = Dir(sDir)
    Loop

End Sub
