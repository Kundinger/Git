Attribute VB_Name = "Module11"
' error module 11 '''''''''''''''''''program SysDef.bas ''''''''''''''''''''
Option Explicit
    
' Initialize Values

Public Function SetupAccessXface() As Boolean
' Routine Name:     SetupAccessXface
' Description:      This routine checks for the version, etc. of the Access Office Application.
'                   The routine is called once at start up.
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 14100
Dim verstr As String
Dim dbApp As Object
Dim flag As Boolean
Dim msg2 As String

    DoEvents
    flag = True
    
msg2 = "Creating Access Application Object"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' open Access Database
    Set dbApp = CreateObject("Access.Application", "")
    dbApp.Visible = False
    
msg2 = "Access Application Object created"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' which version of Access
    ChgErrModule 11, 14101
    verstr = dbApp.Version          '(Build, Name, Version, Visible)
    AccessVerNum = CSng(val(verstr))
        
msg2 = "Access Version is " & verstr
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    Select Case AccessVerNum
    
        Case 0
            ' Default = Office 2003
            ChgErrModule 11, 14102
            AccessVerDesc = "None->Access(2003)"
            AccessDbFileExt = ".mdb"
            
        Case Is < 7
            ' Earlier than Office 95
            ChgErrModule 11, 14105
            AccessVerDesc = "<95->Access(2003)"
            AccessDbFileExt = ".mdb"
    
        Case 7
            ' Office 95
            ChgErrModule 11, 14107
            AccessVerDesc = "Access 95"
            AccessDbFileExt = ".mdb"
    
        Case 8, 8.5
            ' Office 98
            ChgErrModule 11, 14108
            AccessVerDesc = "Access 98"
            AccessDbFileExt = ".mdb"
            
        Case 9
            ' Office 2000
            ChgErrModule 11, 14109
            AccessVerDesc = "Access 2000"
            AccessDbFileExt = ".mdb"
            
        Case 10
            ' Office XP
            ChgErrModule 11, 14110
            AccessVerDesc = "Access XP"
            AccessDbFileExt = ".mdb"
            
        Case 11
            ' Office 2003
            ChgErrModule 11, 14111
            AccessVerDesc = "Access 2003"
            AccessDbFileExt = ".mdb"
            
        Case 12
            ' Office 2007
            ChgErrModule 11, 14112
            AccessVerDesc = "Access 2007"
            AccessDbFileExt = ".accdb"
        
        Case 14
            ' Office 2010
            ChgErrModule 11, 14114
            AccessVerDesc = "Access 2010"
            AccessDbFileExt = ".accdb"
        
        Case 15
            ' Office 2013
            ChgErrModule 11, 14115
            AccessVerDesc = "Access 2013"
            AccessDbFileExt = ".accdb"
        
        Case 16
            ' Office 2016
            ChgErrModule 11, 14116
            AccessVerDesc = "Access 2016"
            AccessDbFileExt = ".accdb"
        
        Case 17
            ' Office 2019
            ChgErrModule 11, 14117
            AccessVerDesc = "Access 2019"
            AccessDbFileExt = ".accdb"
        
        Case Else
            ' Office later than 2019
            ChgErrModule 11, 14120
            AccessVerDesc = ">Access 2019"
            AccessDbFileExt = ".accdb"
        
    End Select
    
' temporary reversion
' temporary reversion
' temporary reversion
AccessDbFileExt = ".mdb"
' temporary reversion
' temporary reversion
' temporary reversion

msg2 = "Nothinging Access Application Object"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    Set dbApp = Nothing
    SetupAccessXface = flag

msg2 = "Setup Access Interface Done"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
flag = False
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    SetupAccessXface = flag
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function SetupExcelXface() As Boolean
' Routine Name:     SetupExcelXface
' Description:      This routine checks for the version, etc. of the Excel Office Application.
'                   The routine is called once at start up.
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 14300
Dim verstr As String
Dim xlApp As Object
Dim flag As Boolean
Dim msg2 As String

    DoEvents
    flag = True
    
    ' Default DB File extention
    DefaultDbFileExt = ".mdb"
    
msg2 = "Creating Excel Application Object"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' open Excel application
    Set xlApp = CreateObject("Excel.Application", "")
    xlApp.Visible = False
    
msg2 = "Excel Application Object created"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' which version of Excel
    XLverstr = xlApp.Version
    XLvernum = val(XLverstr)
    
msg2 = "Access Version is " & XLverstr
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    Select Case XLvernum
    
        Case 7
            ' Office 95
            ChgErrModule 3, 24241
            XLrecsPerSheet = 65532
            XLverdesc = "Excel 95"
            ReportsXlsFileExt = ".xls"

        Case 8
            ' Office 98
            ChgErrModule 3, 24241
            XLrecsPerSheet = 65532
            XLverdesc = "Excel 98"
            ReportsXlsFileExt = ".xls"
            
        Case 9
            ' Office 2000
            ChgErrModule 3, 24242
            XLrecsPerSheet = 65532
            XLverdesc = "Excel 2000"
            ReportsXlsFileExt = ".xls"
            
        Case 10
            ' Office XP
            ChgErrModule 3, 24242
            XLrecsPerSheet = 65532
            XLverdesc = "Excel XP"
            ReportsXlsFileExt = ".xls"
            
        Case 11
            ' Office 2003
            ChgErrModule 3, 24242
            XLrecsPerSheet = 65532
            XLverdesc = "Excel 2003"
            ReportsXlsFileExt = ".xls"
            
        Case 12
            ' Office 2007
            ChgErrModule 3, 24243
            XLrecsPerSheet = 1048576
            XLverdesc = "Excel 2007"
            ReportsXlsFileExt = ".xlsx"
        
        Case 14
            ' Office 2010
            ChgErrModule 3, 24244
            XLrecsPerSheet = 1048576
            XLverdesc = "Excel 2010"
            ReportsXlsFileExt = ".xlsx"
        
        Case 15
            ' Office 2013
            ChgErrModule 3, 24245
            XLrecsPerSheet = 1048576
            XLverdesc = "Excel 2013"
            ReportsXlsFileExt = ".xlsx"
        
        Case 16
            ' Office 2016
            ChgErrModule 3, 24246
            XLrecsPerSheet = 1048576
            XLverdesc = "Excel 2016"
            ReportsXlsFileExt = ".xlsx"
        
        Case 17
            ' Office 2019
            ChgErrModule 3, 24247
            XLrecsPerSheet = 1048576
            XLverdesc = "Excel 2019"
            ReportsXlsFileExt = ".xlsx"
        
        Case Else
            ' Office later than 2019
            ChgErrModule 3, 24249
            XLrecsPerSheet = 1048576
            XLverdesc = "unknown"
            ReportsXlsFileExt = ".xlsx"
        
    End Select
    
msg2 = "Nothinging Excel Application Object"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    Set xlApp = Nothing
    SetupExcelXface = flag

msg2 = "Setup Excel Interface Done"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
flag = False
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    SetupExcelXface = flag
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Sub FilesAndFolders()
' Routine Name:     FilesAndFolders
' Description:      This routine checks for the presence of required files and folders.
'                   The routine is called once at start up.

Dim EnvString As String
Dim msg As String
Dim msg2 As String
Dim msg3 As String
Dim curFilePath As String
Dim sPath As String
Dim DBmsg As String
Dim XLmsg As String
Dim OSplatform, OSversion, OSbuild As Integer
Dim fs As Object

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 1

Set fs = CreateObject("Scripting.FileSystemObject")

    ' Check for non-standard file locations
    EnvString = Environ("WHEREISCPS")   ' Get environment variable WHEREISCPS.
    If Len(EnvString) > 0 Then
        filepath = EnvString
        msg = "CPS Files Folder = " & EnvString
        frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg) & vbCrLf
'        Delay_Box "", INTRODELAY, msgNOSHOW
    Else
    '   No WHEREISCPS environment variable exists; use default root path
        filepath = "C:\CPS_r7"
    End If
    
msg2 = "Checking for Files and Paths"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Does the root path exist?
    If Not fs.FolderExists(filepath & "\") Then
        msg = filepath & "     CPS ROOT PATH DOES NOT EXIST OR IS EMPTY !!!" & vbCrLf & vbCrLf
        msg = msg & "Cannot continue. Please fix the problem then try again." & vbCrLf
        MsgBox msg, vbCritical, "SHUTTING DOWN!"
        End
    Else
    
msg2 = "Setting File Drive"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        ' Set file drive
        DRIVEPATH = Mid(filepath, 1, 2)
        
msg2 = "Setting Paths"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        ' Set file paths
        FILEPATH_manuals = filepath & "\manuals\"   ' root path for all manuals files
        FILEPATH_reports = filepath & "\reports\"   ' root path for all report files
        FILEPATH_cal = filepath & "\calibrate\"     ' root path for all calibration files
        FILEPATH_cfg = filepath & "\config\"        ' root path for all configuration files
        FILEPATH_log = filepath & "\logs\"          ' root path for all log files
        FILEPATH_rcp = filepath & "\recipes\"       ' root path for all recipe (and canister) files
        FILEPATH_data = filepath & "\data\"         ' root path for all data files
        FILEPATH_sysdbf = filepath & "\sysdbf\"     ' root path for all system database files
        FILEPATH_backup = filepath & "\backup\"     ' default path for backup files
        FILEPATH_avlhost = filepath & "\AVL_Host\"  ' root path for Remote AVL Host files
        FILEPATH_avlstatus = FILEPATH_avlhost & "Status\"  ' path for Remote AVL Host Status files
        FILEPATH_avltasks = FILEPATH_avlhost & "Tasks\"    ' path for Remote AVL Host Task files
    
        
msg2 = "Paths Set; check for Ups-has-low-battery file"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        ' delete UpsAlarmOn.txt file; if it exists
        ' Ups-has-low-battery file
        If fs.FileExists(filepath & "\UpsAlarmOn.txt") Then
            fs.DeleteFile (filepath & "\UpsAlarmOn.txt")
msg2 = "deleted withUps-has-low-battery file"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        End If
        
msg2 = "Done with Ups-has-low-battery file;  do Paths exist?"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        ' Do all these paths exist?
        msg = ""
        ' root path for all manuals files
        If Not fs.FolderExists(FILEPATH_manuals) Then msg = msg & FILEPATH_manuals & "     CPS Manual Files Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all report files
        If Not fs.FolderExists(FILEPATH_reports) Then msg = msg & FILEPATH_reports & "      CPS Report Files Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all calibration files
        If Not fs.FolderExists(FILEPATH_cal) Then msg = msg & FILEPATH_cal & "    CPS Calibration Files Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all configuration files
        If Not fs.FolderExists(FILEPATH_cfg) Then msg = msg & FILEPATH_cfg & "        CPS Configuration Files Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all log files
        If Not fs.FolderExists(FILEPATH_log) Then msg = msg & FILEPATH_log & "        CPS Logs Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all recipe files
        If Not fs.FolderExists(FILEPATH_rcp) Then msg = msg & FILEPATH_rcp & "        CPS Recipe Files Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all data files
        If Not fs.FolderExists(FILEPATH_data) Then msg = msg & FILEPATH_data & "          CPS Data Files Path Does Not Exist or is Empty" & vbCrLf
        ' root path for all system database files
        If Not fs.FolderExists(FILEPATH_sysdbf) Then msg = msg & FILEPATH_sysdbf & "          CPS System Database Files Path Does Not Exist or is Empty" & vbCrLf
        ' default root path for backup files
        If Not fs.FolderExists(FILEPATH_backup) Then msg = msg & FILEPATH_backup & "     CPS Default Files Backup Path Does Not Exist or is Empty" & vbCrLf
        If USINGREMAVLFILES Then
            ' root path for all AVL Host files
            If Not fs.FolderExists(FILEPATH_avlhost) Then msg = msg & FILEPATH_avlhost & "          AVL Host Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Status files
            If Not fs.FolderExists(FILEPATH_avlstatus) Then msg = msg & FILEPATH_avlstatus & "          AVL Host Status Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Task files
            If Not fs.FolderExists(FILEPATH_avltasks) Then msg = msg & FILEPATH_avltasks & "          AVL Host Task Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Task ToDo files
            sPath = FILEPATH_avltasks & "ToDo\"
            If Not fs.FolderExists(sPath) Then msg = msg & sPath & "          AVL Host Task ToDo Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Task OnList files
            sPath = FILEPATH_avltasks & "OnList\"
            If Not fs.FolderExists(sPath) Then msg = msg & sPath & "          AVL Host Task OnList Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Task InProcess files
            sPath = FILEPATH_avltasks & "InProcess\"
            If Not fs.FolderExists(sPath) Then msg = msg & sPath & "          AVL Host Task InProcess Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Task Done files
            sPath = FILEPATH_avltasks & "Done\"
            If Not fs.FolderExists(sPath) Then msg = msg & sPath & "          AVL Host Task Done Files Path Does Not Exist or is Empty" & vbCrLf
            ' path for Remote AVL Host Task Failed files
            sPath = FILEPATH_avltasks & "Failed\"
            If Not fs.FolderExists(sPath) Then msg = msg & sPath & "          AVL Host Task Failed Files Path Does Not Exist or is Empty" & vbCrLf
        End If
        ' OK? else End
        If msg <> "" Then
            msg = msg & vbCrLf
            msg = msg & "Cannot continue. Please fix the path error(s) then try again." & vbCrLf
            MsgBox msg, vbCritical, "SHUTTING DOWN!"
            End
        Else
        
msg2 = "Paths OK"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        End If
    End If
    
    
msg2 = "Check for Access version"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Default DB File extention
    DefaultDbFileExt = ".mdb"
    ' setup Access Interface
    ChgErrModule 11, 10100
    If SetupAccessXface Then
    
        ' Use Installed version of Access interface
        DBmsg = "Using Access Version " & Format(AccessVerNum, "##0.0") & " aka " & AccessVerDesc
        frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(DBmsg) & vbCrLf
        
    Else
    
        ' use default Access interface = Access 2003
        ChgErrModule 11, 10101
        DBmsg = "Access Interface Setup Failed; Using default file version (mdb)."
        frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg) & vbCrLf
        AccessVerDesc = "Default Access(2003)"
        AccessVerNum = 11
        AccessDbFileExt = DefaultDbFileExt
            
    End If
    
msg2 = "Check for Excel version"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' setup Excel Interface
    ChgErrModule 1, 10200
    If SetupExcelXface Then
    
        ' Use Installed version of Excel interface
        XLmsg = "Using Excel Version " & Format(XLvernum, "##0.0") & " aka " & XLverdesc
        frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(XLmsg) & vbCrLf
        
    Else
    
        ' use default Excel interface = Excel 2003
        ChgErrModule 1, 10201
        XLmsg = "Excel Interface Setup Failed; Using default file version (mdb)."
        frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(XLmsg) & vbCrLf
        XLvernum = 11
        XLrecsPerSheet = 65532
        XLverdesc = "Default Excel(2003)"
        ReportsXlsFileExt = ".xls"
            
    End If
    
    
msg2 = "Check for Excel Template File"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Check for excel template file
    curFilePath = FILEPATH_reports & "cpsXlsReport" & ReportsXlsFileExt
    msg3 = ""
    If (Not fs.FileExists(curFilePath)) Then
        msg3 = ">>>" & curFilePath & "<<<   does Not Exist"
        Write_ELog msg3
        ' try using default report file extention
        curFilePath = FILEPATH_reports & "cpsXlsReport" & ".xls"
        If fs.FileExists(curFilePath) Then
            msg3 = ">>>" & curFilePath & "<<<   does Exist; Proceeding"
            Write_ELog msg3
            msg3 = ""
            ReportsXlsFileExt = ".xls"
        Else
            ' no excel template file to use
            msg3 = ">>>" & curFilePath & "<<<   does Not Exist"
            Write_ELog msg3
            msg = msg3 & vbCrLf
            msg = msg & "Cannot continue. Please fix the filepath error then try again." & vbCrLf
            MsgBox msg, vbCritical, "SHUTTING DOWN!"
            End
        End If
    End If
    
msg2 = "Set DB File Names"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Set DB file names
    ChgErrModule 11, 10102
    DATACAL = "cpsCalibrations_rev" & Format(DBFREVLVL, "0#") & Format(DBCALREVLVL, "00") & AccessDbFileExt               ' DataBase File with Recipe and Canister info
    DATAMODEL = "cpsModel_rev" & Format(DBFREVLVL, "0#") & Format(DBMODELREVLVL, "00") & AccessDbFileExt                  ' Template DataBase file
    DATAMASTER = "cpsMaster_rev" & Format(DBFREVLVL, "0#") & Format(DBMASTERREVLVL, "00") & AccessDbFileExt               ' DataBase File with Job Log, Event Log, etc.
    DATAUSER = "cpsUser_rev" & Format(DBFREVLVL, "0#") & Format(DBUSERREVLVL, "00") & AccessDbFileExt                     ' DataBase File with User Logon info
    DATARCP = "cpsRecipes_rev" & Format(DBFREVLVL, "0#") & Format(DBRCPREVLVL, "00") & AccessDbFileExt                    ' DataBase File with Recipe and Canister info
    DATASYSDEF = "cpsSysDef_rev" & Format(DBFREVLVL, "0#") & Format(DBSYSDEFREVLVL, "00") & AccessDbFileExt               ' DataBase File with System Definition info
    DATAAIRLOG = "cpsAirLogModel_rev" & Format(DBFREVLVL, "0#") & Format(DBAIRLOGREVLVL, "00") & AccessDbFileExt          ' DataBase File Master for Air Temp/Rh logs
    DATAREM = "cpsRemote_rev" & Format(DBFREVLVL, "0#") & Format(DBREMREVLVL, "00") & AccessDbFileExt                 ' DataBase File with Remote Tasks
    DATAZLOG = "cpsZlog_rev" & Format(DBFREVLVL, "0#") & Format(DBZLOGREVLVL, "00") & AccessDbFileExt                     ' DataBase File with Zlogs (debug logs)
    DATATOM = "cpsTomCanLoad_rev" & Format(DBFREVLVL, "0#") & Format(DBTOMREVLVL, "00") & ".mdb"                 ' DataBase File with TOM Tasks
    
msg2 = "DB File Names set; do they exist?"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Do these required DB files exist?
    msg = ""
    ' DataBase File with Calibration info
    If Not fs.FileExists(FILEPATH_cal & DATACAL) Then msg = msg & DATACAL & "         CPS Calibration DataBase File does Not Exist" & vbCrLf
    ' Template DataBase file
    If Not fs.FileExists(FILEPATH_sysdbf & DATAMODEL) Then msg = msg & DATAMODEL & "       CPS Template DataBase File does Not Exist" & vbCrLf
    ' DataBase File with Job Log, Event Log, etc.
    If Not fs.FileExists(FILEPATH_sysdbf & DATAMASTER) Then msg = msg & DATAMASTER & "     CPS Master DataBase File does Not Exist" & vbCrLf
    ' DataBase File with User Logon info
    If Not fs.FileExists(FILEPATH_sysdbf & DATAUSER) Then msg = msg & DATAUSER & "         CPS User Logon DataBase File does Not Exist" & vbCrLf
    ' DataBase File with Recipe and Canister info
    If Not fs.FileExists(FILEPATH_rcp & DATARCP) Then msg = msg & DATARCP & "         CPS Recipe DataBase File does Not Exist" & vbCrLf
    ' DataBase File with System info
    If Not fs.FileExists(FILEPATH_sysdbf & DATASYSDEF) Then msg = msg & DATASYSDEF & "         CPS System Definition File does Not Exist" & vbCrLf
    ' DataBase File Master for Air Temp/Rh Logs
    If Not fs.FileExists(FILEPATH_sysdbf & DATAAIRLOG) Then msg = msg & DATAAIRLOG & "     CPS AirLog DataBase File does Not Exist" & vbCrLf
    ' DataBase File with Zlogs
    If Not fs.FileExists(FILEPATH_sysdbf & DATAZLOG) Then msg = msg & DATAZLOG & "     CPS Zlog DataBase File does Not Exist" & vbCrLf
    ' OK?
    If msg <> "" Then
        msg = msg & vbCrLf
        
msg2 = "DB Files Problem "
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    Else
    
msg2 = "DB Files OK"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf

        ElogInitDone = True
    
    End If
    
    
msg2 = "Check for Calibration.xls"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Check for Calibration.xls file
    ChgErrModule 11, 10103
    sPath = FILEPATH_cal
    If Not fs.FileExists(sPath & "CALIBRATION.xls") Then msg = msg & sPath & "CALIBRATION.xls" & "       Calibration spreadsheet does Not Exist" & vbCrLf
    
msg2 = "Calibration.xls OK"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
msg2 = "Check for Opto ToolKit"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
    ' Check for OptoDriverToolkit files
    ChgErrModule 11, 10104
    EnvString = Environ("windir")   ' Get Windows Directory environment variable
    If Len(EnvString) > 0 Then
        sPath = EnvString
    '    Msg = "Windows Folder = " & EnvString
    '    Delay_Box CStr(Msg), MSGDELAY, msgSHOW
    Else
    '   No WINDIR environment variable exists; use default windows path
        sPath = "c:\windows"
    End If
    ChgErrModule 11, 10105
    sPath = sPath & "\system32\"
    ' Do the required OptoDriverToolkit files exist?
    ' Opto22 Error Message DLL
    If Not fs.FileExists(sPath & "OPTOERR.DLL") Then msg = msg & sPath & "OPTOERR.DLL" & "       Opto22 Error Message DLL does Not Exist" & vbCrLf
    ' Opto22 Communications DLL
    If Not fs.FileExists(sPath & "OPTOCOM.DLL") Then msg = msg & sPath & "OPTOCOM.DLL" & "       Opto22 Communications DLL does Not Exist" & vbCrLf
    ' Opto22 MWD DLL
    If Not fs.FileExists(sPath & "OPTOMWD.DLL") Then msg = msg & sPath & "OPTOMWD.DLL" & "       Opto22 MWD DLL does Not Exist" & vbCrLf
    
    ' OK? else End
    ChgErrModule 11, 10109
    If msg <> "" Then
        msg = msg & vbCrLf
        msg = msg & "Cannot continue. Please fix the file error(s) then try again." & vbCrLf
        MsgBox msg, vbCritical, "SHUTTING DOWN!"
        End
    Else
    
msg2 = "Opto ToolKit OK"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
        Write_ELog DBmsg
        Write_ELog XLmsg
    End If
    
msg2 = "End of Files and Folders"
If STARTUPVERBOSE Then frmAbout.txtMessage.text = frmAbout.txtMessage.text & CStr(msg2) & vbCrLf
    
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

Public Sub Init_Vars()
' Function Name:    Init_Vars
' Author:           Analytical Process Programmer         8/96
' Description:      This routine initializes the variables.
'                   The routine is called at power up one time.

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 6
Dim idx1 As Integer
Dim idx2 As Integer
Dim iAux As Integer

    FORMFEED = Chr$(12)
    
    ' System Timer Descriptions
    SystemTimers(1).desc = "Opto I/O Scanner"
    SystemTimers(2).desc = "Read Scales"
    SystemTimers(3).desc = "Alarm and OOT Check"
    SystemTimers(4).desc = "DataLogger"
    SystemTimers(5).desc = "Controllers Logic"
    SystemTimers(6).desc = "Station Logic"
    SystemTimers(7).desc = "System Timers"
    SystemTimers(8).desc = "unused"
    SystemTimers(9).desc = "unused"
    
    ' MFC Calibration File Names
    Mfc_Calib_Filename(MFCBUTANE) = "Butane_"
    Mfc_Calib_Filename(MFCNITROGEN) = "Nitrogen_"
    Mfc_Calib_Filename(MFCPURGEAIR) = "PurgeAir_"
    Mfc_Calib_Filename(MFCORVRBUT) = "ORVRBut_"
    Mfc_Calib_Filename(MFCORVRNIT) = "ORVRNit_"
    Mfc_Calib_Filename(MFCORVRPRG) = "PurgeAir_"
    Mfc_Calib_Filename(MFCLIVEFUEL) = "LiveFuel_"
    Mfc_Calib_Filename(MFCORVRLIVE) = "ORVRLiveFuel_"
    
    ' MFC Descriptions
    Mfc_Description(MFCBUTANE) = "Butane"
    Mfc_Description(MFCNITROGEN) = "Nitrogen"
    Mfc_Description(MFCPURGEAIR) = "PurgeAir"
    Mfc_Description(MFCORVRBUT) = "ORVRButane"
    Mfc_Description(MFCORVRNIT) = "ORVRNitrogen"
    Mfc_Description(MFCORVRPRG) = "ORVRPurgeAir"
    Mfc_Description(MFCLIVEFUEL) = "LiveFuel"
    Mfc_Description(MFCORVRLIVE) = "ORVRLiveFuel"
    
    ' Initialize Station Values on Starting the system from power up
    For idx1 = 1 To NR_STN
        For idx2 = 1 To NR_SHIFT
        
            Station_Clear idx1, idx2
            StationControl(idx1, idx2).Mode = VBIDLE
            Stn_Purge_Log_TestTimer(idx1, idx2) = 0
            Stn_Load_Log_TestTimer(idx1, idx2) = 0
            Stn_Leak_Log_TestTimer(idx1, idx2) = 0
            Stn_LT_Log_TestTimer(idx1, idx2) = 0
            Stn_XYGraph_TestTimer(idx1, idx2) = 0
            ' Clear tolerance Data
            OOTs(idx1, idx2).NitFlowOOT = False
            OOTs(idx1, idx2).BtnFlowOOT = False
            OOTs(idx1, idx2).PurFlowOOT = False
            OOTs(idx1, idx2).FuelTempOOT = False
            OOTs(idx1, idx2).AirTempOOT = False
            OOTs(idx1, idx2).AirMoistOOT = False
            OOTs(idx1, idx2).CanVentOOT = False
            ' Set all OOT counts to zero
            OOTs(idx1, idx2).PurFlowOOTCnt = 0
            OOTs(idx1, idx2).BtnFlowOOTCnt = 0
            OOTs(idx1, idx2).NitFlowOOTCnt = 0
            OOTs(idx1, idx2).FuelTempOOTCnt = 0
            OOTs(idx1, idx2).AirTempOOTCnt = 0
            OOTs(idx1, idx2).AirMoistOOTCnt = 0
            ' Clear Alarm Data
            Alm_LoadPress(idx1, idx2) = False
            Alm_LiveFuelLevel(idx1, idx2) = False
            Alm_LiveFuelHeater(idx1, idx2) = False
            Alm_LiveFuelSafety(idx1, idx2) = False
            ' Initialize Station Scale Values
            StationControl(idx1, idx2).AuxScaleWt = CSng(0)             ' Aux Scale Weight
            StationControl(idx1, idx2).PriScaleWt = CSng(0)             ' Primary Scale Weight
            
            StationControl(idx1, idx2).LiveFuelCycleCount = 0
            StationControl(idx1, idx2).StartRequest = False
            StationControl(idx1, idx2).ContinueRequest = False
            StationControl(idx1, idx2).StopRequest = False
            StationControl(idx1, idx2).AbortRequest = False
            PreviousReportTimer(idx1, idx2) = 0
            DeltaTimer(idx1, idx2) = 0
            PreviousTotalTimer(idx1, idx2) = 0
            DeltaTotalTimer(idx1, idx2) = 0
            Stn_XYChart_Xinterval(idx1, idx2) = 10  ' note: No screens yet
            StationControl(idx1, idx2).EstJobDur = StationSequence(idx1, idx2).EstSeqDuration
            StationControl(idx1, idx2).EstJobDurDesc = StationSequence(idx1, idx2).EstSeqDurDesc
            InIdx(idx1, idx2) = 1
            ' clear Canister LeakCheck Status
            StationControl(idx1, idx2).LeakCheckStatus = NORESULT
            StationControl(idx1, idx2).LcStatusDescription = " "
            
        Next idx2
        
        STN_IOForceMode(idx1) = VBMANUAL
        Stn_ActiveShift(idx1) = 1
        AdfControl(idx1).Enable = False
        AdfControl(idx1).Heater_Enable = False
        AdfControl(idx1).ManScreen_Enable = False
        AdfControl(idx1).ReadyForRefill = False
        AdfControl(idx1).RefillRequest = False
        AdfControl(idx1).LiveFuelDensityOkCnt = 0
        AdfControl(idx1).LiveFuelDensityDeadCnt = False
        AdfControl(idx1).LiveFuelDensityWeakCnt = False
        AdfControl(idx1).LiveFuelState = fuelOK
        
        ' scale simulation
        Sim_AuxWt(idx1) = CSng(0)
        Sim_PriWt(idx1) = CSng(0)
        Sim_AuxWt_Current(idx1) = CSng(0)
        Sim_PriWt_Current(idx1) = CSng(0)
        Sim_AuxWt_Last(idx1) = CSng(0)
        Sim_PriWt_Last(idx1) = CSng(0)
    Next idx1
    
    ' Initialize Scale Values
    For idx1 = 1 To NR_SCALES
        Scale_In_Use(idx1) = False
        Scale_Weight(idx1) = CSng(0)
        Scale_Value(idx1) = Format(Scale_Weight(idx1), "######0.00")
    Next idx1
    
    ' Clear Alarm Data
    Alm_Estop = False
    Alm_Flow = False
    Alm_Btn20 = False
    Alm_N2 = False
    Alm_Doors = False
    
    STN_IOForceMode(0) = VBMANUAL
    IOForceActive = False
    CalibrateActive = False
    OptoReadAllOnce = False
    ScalesReadAllOnce = False
    
    WTCHGQUEUESIZE = 1000
    
    LeakCheckControl.station = 0
    LeakCheckControl.Shift = 0
    ' Common Function Definition Indexes
    DefFunc(0, 0) = 0
    DefFunc(1, 0) = 0
    ' FID Function Definition Indexes
    DefFunc(0, 1) = 0
    DefFunc(1, 1) = 0
    ' Station Function Definition Indexes
    DefFunc(0, 2) = 0
    DefFunc(1, 2) = 0
    ' PurgeAir Function Definition Indexes
    DefFunc(0, 3) = 0
    DefFunc(1, 3) = 0
    ' Station Definition Index
    DefStn = 1
    ' Station Currently on Display
    DispStn = 1
    DispShift = 1
    ' PurgeAir Definition Index
    DefPrg = 1
    LastPurgeStart = Sys_StartTime - TimeSerial(0, 5, 0)
    ' Load MFC Calibration Values (set to linear first)
    frmAbout.txtMessage.text = frmAbout.txtMessage.text & "Loading MFC Calibration Values" & vbCrLf
    Delay_Box "", INTRODELAY, msgNOSHOW
    For idx1 = 1 To LAST_STN
        SetMfcCalLinear idx1
        LoadMfcCalibration idx1
    Next idx1
    ' Load Analog Input Calibration Values (set to linear first)
    frmAbout.txtMessage.text = frmAbout.txtMessage.text & "Loading AI Calibration Values" & vbCrLf
    Delay_Box "", INTRODELAY, msgNOSHOW
    SetAiCalLinear calgrpComm
    LoadAiCalibration calgrpComm
    For idx1 = 1 To LAST_STN
        SetAiCalLinear idx1
        LoadAiCalibration idx1
    Next idx1
    For idx1 = 11 To (NR_PRGAIR + 10)
        SetAiCalLinear idx1
        If USINGPASLOCALCONTROL Then LoadAiCalibration idx1
    Next idx1
    ' Load Scale Calibration Values (set to linear first)
    frmAbout.txtMessage.text = frmAbout.txtMessage.text & "Loading Scale Calibration Values" & vbCrLf
    Delay_Box "", INTRODELAY, msgNOSHOW
    SetSclCalLinear
    LoadSclCalibration
    
    ' Blank Recipe
    EmptyRecipe.Name = "undefined"
    EmptyRecipe.Number = CInt(0)
    
    EmptyRecipe.CycleType = CyclePurgeLoad
    
    EmptyRecipe.Load_Method = NOLOAD
    EmptyRecipe.UseHiRangeMFC = False
    EmptyRecipe.UseLoadRatePID = False
    EmptyRecipe.NitrogenFlow = CSng(0)
    EmptyRecipe.NitrogenFlowSave = CSng(0)
    EmptyRecipe.Load_Rate = 0
    EmptyRecipe.Load_RateSave = 0
    EmptyRecipe.Mix_Percent = 0
    EmptyRecipe.WC_Mult = 0
    EmptyRecipe.EPAFill = 0
    EmptyRecipe.Load_Wt = 0
    EmptyRecipe.LoadBreakthrough = 0
    EmptyRecipe.Load_Time = 0
    
    EmptyRecipe.Purge_Method = NOPURGE
    EmptyRecipe.Purge_Flow = 0
    EmptyRecipe.Purge_Time = 0
    EmptyRecipe.Purge_AuxTime = 0
    EmptyRecipe.Purge_Can_Vol = 0
    EmptyRecipe.Purge_ProfileNumber = 0
    EmptyRecipe.Purge_TargetMode = 0
    EmptyRecipe.Purge_TargetWC = 0
    EmptyRecipe.Purge_TargetWeight = 0
    EmptyRecipe.Purge_MaxVolumes = 0
    EmptyRecipe.Purge_TargetPurge = 0
    EmptyRecipe.Purge_TargetPause = 0
    
    EmptyRecipe.UseAuxScale = False
    EmptyRecipe.PurgeAuxCan = False
    EmptyRecipe.AuxScaleNo = 0
    EmptyRecipe.PauseLeakTime = 0
    EmptyRecipe.PauseLoadTime = 0
    EmptyRecipe.PausePurgeTime = 0
    EmptyRecipe.UsePriScale = False
    EmptyRecipe.PriScaleNo = 0
    EmptyRecipe.PauseAfterLeak = False
    EmptyRecipe.PauseAfterLoad = False
    EmptyRecipe.PauseAfterPurge = False
    EmptyRecipe.LeakCheck = False
    EmptyRecipe.LeakPrimary = False
    EmptyRecipe.LeakAux = False
    EmptyRecipe.MaxLoadTime = 0
    
    EmptyRecipe.IDLoad = 0
    EmptyRecipe.LoadL = 0
    EmptyRecipe.LoadV = 0
    EmptyRecipe.IDPurge = 0
    EmptyRecipe.PurgeL = 0
    EmptyRecipe.PurgeV = 0
    EmptyRecipe.IDVent = 0
    EmptyRecipe.VentL = 0
    EmptyRecipe.VentV = 0
    
    EmptyRecipe.LiveFuel = False
    EmptyRecipe.LiveFuelChgAuto = False
    EmptyRecipe.LiveFuelChgFreq = 0
    EmptyRecipe.ADF_Heater = False
    EmptyRecipe.ADF_HeaterSP = 0
    
    EmptyRecipe.StartMethod = STARTNOW
    EmptyRecipe.StartDelay = 0
    EmptyRecipe.StartDate = Now()
                
    ' end method
    EmptyRecipe.EndMethod = 0
    EmptyRecipe.Cycles = CInt(0)
    EmptyRecipe.EndWeightTolerance = CInt(0)
    EmptyRecipe.EndConsecutiveCycles = CInt(0)
    EmptyRecipe.EndMaximumCycles = CInt(0)
    EmptyRecipe.EndMinimumCycles = CInt(0)

    EmptyRecipe.AuxOutputs = False
    For iAux = 1 To 4
        EmptyRecipe.AuxOutputs_Load(iAux) = False
        EmptyRecipe.AuxOutputs_Purge(iAux) = False
    Next iAux
                    
    ' Blank LeakData
    BlankLeakData.ClkTime = CDate("00:00:00")
    BlankLeakData.Comment = ""
    BlankLeakData.Pressure = 0#
    BlankLeakData.TstTimr = 0#
    BlankLeakData.isBlank = True
    
    ' Blank LoadData
    BlankLoadData.AuxScle = 0#
    BlankLoadData.BtnFlow = 0#
    BlankLoadData.LoadRate = 0#
    BlankLoadData.loadTotalGrams = 0#
    BlankLoadData.ClkTime = CDate("00:00:00")
    BlankLoadData.FuelTmp = 0#
    BlankLoadData.LFcycls = 0
    BlankLoadData.NitFlow = 0#
    BlankLoadData.PriScle = 0#
    BlankLoadData.TstTimr = 0#
    BlankLoadData.isBlank = True
    
    ' Blank PurgeData
    BlankPurgeData.AuxScle = 0#
    BlankPurgeData.ClkTime = CDate("00:00:00")
    BlankPurgeData.PrgFlow = 0#
    BlankPurgeData.PrgHumd = 0#
    BlankPurgeData.PrgTemp = 0#
    BlankPurgeData.PriScle = 0#
    BlankPurgeData.TstTimr = 0#
    BlankPurgeData.isBlank = True
    
    ' Blank LeakTest Data
    BlankLT2_Data.AtmPress = 0#
    BlankLT2_Data.ClkTime = CDate("00:00:00")
    BlankLT2_Data.EffDia = 0#
    BlankLT2_Data.InPress = 0#
    BlankLT2_Data.NitFlow = 0#
    BlankLT2_Data.NitTemp = 0#
    BlankLT2_Data.SecTimer = 0#
    BlankLT2_Data.isBlank = True
        
    
    ' current & last error status
    With ErrorStatus_NoErrors
        .AnyError = False
        .MoistOOT = False
        .MoistTO = False
        .TempOOT = False
        .TempTO = False
        .TestBit = False
    End With
    ErrorStatus_Current = ErrorStatus_NoErrors
    ErrorStatus_Last = ErrorStatus_NoErrors

    ' misc flags, etc.
    Stop_In_Progress = False
    Pause_Alarm = NOTPAUSED
    LastOOTCheckTime = Now()
    ' remote status file update timer
    For idx1 = 1 To NR_STN
        For idx2 = 1 To NR_SHIFT
            Stn_RemStatus_Log_TestTimer(idx1, idx2) = Timer
        Next idx2
    Next idx1
    ' zLog PasLog
    Debug_ZlogPAS_LastUpdate = Timer
    Debug_ZlogPAS_MaxRecords = 99999
    Debug_ZlogPAS_NumRecords = 0
    Debug_ZlogPAS_Clear = False
    ' zLog PurgeLog
    Debug_ZlogPurge_MaxRecords = 9999
    Debug_ZlogPurge_NumRecords = 0
    Debug_ZlogPurge_Clear = False
    ' zLog ScaleLog
    Debug_ZlogScale_MaxRecords = 99999
    Debug_ZlogScale_NumRecords = 0
    Debug_ZlogScale_Clear = False
    '
    '   *****************
    '   PAS LOCAL CONTROL
    '   *****************
    '
    ' AK default values
    AK_portNumStr = "5600"
    AK_timeout = 30
    AK_anychar = " "
    AK_sepchar = " "
    ' PAS Temperature Ready
    PAS_INFO(pasTEMPERATURE).LastUpdate = Timer
    PID_INFO(pasTEMPERATURE).LastUpdate = Timer
    ' PAS Moisture Ready
    PAS_INFO(pasMOISTURE).LastUpdate = Timer
    PID_INFO(pasMOISTURE).LastUpdate = Timer
    
    ' STN LoadRate Ready
    For idx1 = 1 To NR_STN
        PID_INFO(10 + idx1).LastUpdate = Timer
    Next idx1

    ' Chiller SetPoint
    WaterBathSP = IIf(USINGC, DegFtoC(70), 70)  ' 70 degF

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


Public Sub SetupStnModes()
' Function Name:    SetupStnModes
' Author:           Analytical Process Programmer         12/2005
' Description:      This routine initializes the Station Mode Info and Other Arrays
'                   The routine is called at power up one time.
'
'

Dim iCode As Integer
Dim iMode As Integer
Dim Idx As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 611

    ' Initialize Station Mode Information Arrays on startup
    StnDetail_ButtonTop = 6540
    OutOfSight = 15000
    
    ' Set all values to defaults
    For iMode = 0 To MAX_MODE
        ModeDescLong(iMode) = "UNDEFINED"
        ModeDescShort(iMode) = "undefined"
        ModeBackColor(iMode) = Black
        ModeForeColor(iMode) = DKPURPLE
        Stn_ContinueBtn(iMode).Top = CInt(OutOfSight)
        Stn_ContinueBtn(iMode).Enabled = False
        Stn_ContinueBtn(iMode).ToolTipText = "Continue"
    Next iMode
    
    ' Set values for defined modes
    iMode = VBIDLE
    ModeDescLong(iMode) = "IDLE"
    ModeDescShort(iMode) = "Idle"
    ModeBackColor(iMode) = Common_BackColor
    ModeForeColor(iMode) = DKGRAY
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SIDL"
    
    iMode = VBIDLEWAITING
    ModeDescLong(iMode) = "IDLE - WAITING"
    ModeDescShort(iMode) = "Idle - Waiting"
    ModeBackColor(iMode) = Moccasin
    ModeForeColor(iMode) = Black
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SIDW"
    
    iMode = VBCOMPLETE
    'ModeDescLong(iMode) = "COMPLETE"
    'ModeDescShort(iMode) = "Complete"
    ModeDescLong(iMode) = "GENERATING REPORTS"
    ModeDescShort(iMode) = "Generating Reports"
    ModeBackColor(iMode) = Teal
    ModeForeColor(iMode) = Black
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SDNE"
    
    iMode = VBSCALEWAIT
    ModeDescLong(iMode) = "WAITING SCALE"
    ModeDescShort(iMode) = "Waiting for Scale"
    ModeBackColor(iMode) = Goldenrod
    ModeForeColor(iMode) = DK2GRAY
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SWSC"
    
    iMode = VBSHIFTWAIT
    ModeDescLong(iMode) = "WAITING SHIFT"
    ModeDescShort(iMode) = "Waiting Shift"
    ModeBackColor(iMode) = Goldenrod
    ModeForeColor(iMode) = DK2GRAY
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SWSH"
    
    iMode = VBSTARTWAIT
    ModeDescLong(iMode) = "WAITING START"
    ModeDescShort(iMode) = "Waiting to Start"
    ModeBackColor(iMode) = Goldenrod
    ModeForeColor(iMode) = DK2GRAY
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SIDL"
    
    iMode = VBLEAKWAIT
    ModeDescLong(iMode) = "WAITING LEAK"
    ModeDescShort(iMode) = "Waiting to Leak Check"
    ModeBackColor(iMode) = SandyBrown
    ModeForeColor(iMode) = DK2GRAY
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SWLK"
    
    iMode = VBLEAK
    ModeDescLong(iMode) = "LEAK CHECK"
    ModeDescShort(iMode) = "Leak Check"
    ModeBackColor(iMode) = SkyBlue
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SLKR"
    
    iMode = VBPOSTLEAK
    ModeDescLong(iMode) = "LEAK PAUSED"
    ModeDescShort(iMode) = "Leak Paused"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel post Leak Delay"
    AK_ModeDesc(iMode) = "SLKP"
    
    iMode = VBWBPAUSE
    ModeDescLong(iMode) = "PAUSED FOR WATERBATH"
    ModeDescShort(iMode) = "WaterBath Pause"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel WaterBath Pause"
    AK_ModeDesc(iMode) = "SPWB"
    
    iMode = VBLEAKERROR
    ModeDescLong(iMode) = "LEAK ERROR"
    ModeDescShort(iMode) = "Leak Check Error"
    ModeBackColor(iMode) = Alarm_ForeColor
    ModeForeColor(iMode) = Black
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SLER"
    
    iMode = VBLOAD
    ModeDescLong(iMode) = "LOADING"
    ModeDescShort(iMode) = "Loading"
    ModeBackColor(iMode) = Green
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SLDR"
    
    iMode = VBPOSTLOAD
    ModeDescLong(iMode) = "LOAD PAUSED"
    ModeDescShort(iMode) = "Load Paused"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel post Load Delay"
    AK_ModeDesc(iMode) = "SLDP"
    
    iMode = VBPOSTLOADOPER
    ModeDescLong(iMode) = "LOAD WAIT FOR OPER"
    ModeDescShort(iMode) = "Load Oper Pause"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SLDO"
    
    iMode = VBGASPAUSE
    ModeDescLong(iMode) = "FUEL PAUSE"
    ModeDescShort(iMode) = "Fuel Pause"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPGS"
    
    iMode = VBFIDPAUSE
    ModeDescLong(iMode) = "FID PAUSE"
    ModeDescShort(iMode) = "FID Pause"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel FID pause"
    AK_ModeDesc(iMode) = "SPFD"
    
    iMode = VBPRELOAD
    ModeDescLong(iMode) = "PRELOAD N2"
    ModeDescShort(iMode) = "Preloading N2"
    ModeBackColor(iMode) = ForestGreen
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel PreLoad N2 Push"
    AK_ModeDesc(iMode) = "SPLD"
    
    iMode = VBPURGE
    ModeDescLong(iMode) = "PURGING"
    ModeDescShort(iMode) = "Purging"
    ModeBackColor(iMode) = DKPURPLE
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPRG"
    
    iMode = VBPURGECONT
    ModeDescLong(iMode) = "PURGE CONTINUE"
    ModeDescShort(iMode) = "Purge Continue"
    ModeBackColor(iMode) = SandyBrown
    ModeForeColor(iMode) = Black
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPRC"
    
    iMode = VBPURGEWAIT
    ModeDescLong(iMode) = "PURGE WAITING"
    ModeDescShort(iMode) = "Purge Waiting"
    ModeBackColor(iMode) = SandyBrown
    ModeForeColor(iMode) = DK2GRAY
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SWPR"
    
    iMode = VBPOSTPURGE
    ModeDescLong(iMode) = "PURGE PAUSED"
    ModeDescShort(iMode) = "Purge Paused"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel post Purge Delay"
    AK_ModeDesc(iMode) = "SPRP"
    
    iMode = VBPOSTPURGEOPER
    ModeDescLong(iMode) = "PURGE WAIT FOR OPER"
    ModeDescShort(iMode) = "Purge Oper Wait"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPRO"
    
    iMode = VBPAUSEVACSW
    ModeDescLong(iMode) = "SYSTEM VAC SW OFF"
    ModeDescShort(iMode) = "Waiting for System Vac Sw"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SIDL"
    
    iMode = VBPAUSE
    ModeDescLong(iMode) = "PAUSED"
    ModeDescShort(iMode) = "Paused"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPAU"
    
    iMode = VBPAUSEBYUSER
    ModeDescLong(iMode) = "PAUSED BY USER"
    ModeDescShort(iMode) = "Paused By User"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPUU"
    
    iMode = VBPAUSEOOT
    ModeDescLong(iMode) = "PAUSED OOT"
    ModeDescShort(iMode) = "Paused OOT"
    ModeBackColor(iMode) = LightPink
    ModeForeColor(iMode) = Black
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPOT"
    
    iMode = VBPAUSEALARM
    ModeDescLong(iMode) = "PAUSED ALARM"
    ModeDescShort(iMode) = "Paused Alarm"
    ModeBackColor(iMode) = Alarm_ForeColor
    ModeForeColor(iMode) = Black
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "CONTINUE"
    AK_ModeDesc(iMode) = "SPAL"
    
    iMode = VBCOURSEWAIT
    ModeDescLong(iMode) = "WAITING FOR OPERATOR"
    ModeDescShort(iMode) = "Waiting for Oper"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = PALEBLUE
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "OK to CONTINUE"
    AK_ModeDesc(iMode) = "SWCO"
    
    iMode = VBCOURSEPAUSE
    ModeDescLong(iMode) = "PAUSED BY COURSE"
    ModeDescShort(iMode) = "Paused Course"
    ModeBackColor(iMode) = SaddleBrown
    ModeForeColor(iMode) = PALEBLUE
    Stn_ContinueBtn(iMode).Enabled = True
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel Course Pause"
    AK_ModeDesc(iMode) = "SPCO"
    
    iMode = VBLEAKTEST
    ModeDescLong(iMode) = "LEAKTEST"
    ModeDescShort(iMode) = "LeakTest"
    ModeBackColor(iMode) = MidnightBlue
    ModeForeColor(iMode) = White
    Stn_ContinueBtn(iMode).Enabled = False
    Stn_ContinueBtn(iMode).ToolTipText = "Cancel LeakTest"
    AK_ModeDesc(iMode) = "SLTR"
    
    
    '*****************************************************************
    ' Set all descriptions to defaults
    For Idx = 0 To 9
        CycleTypeDesc(Idx) = "undefined"
        LeakPhaseDesc(Idx) = "undefined"
        LoadPhaseDesc(Idx) = "undefined"
        PurgePhaseDesc(Idx) = "undefined"
        LoadTypeDesc(Idx) = "undefined"
        LoadTypeDesc2(Idx) = "undefined"
        LoadTypeDesc3(Idx) = "undefined"
        StartTypeDesc(Idx) = "undefined"
        StartTypeDesc2(Idx) = "undefined"
        StartMethodDesc(Idx) = "undefined"
        EndMethodDesc(Idx) = "undefined"
        LeakMethodDesc(Idx) = "undefined"
        LoadMethodDesc(Idx) = "undefined"
        PurgeMethodDesc(Idx) = "undefined"
        PurgeTargetDesc(Idx) = "undefined"
        PAGtypeDesc(Idx) = "undefined"
    Next Idx
    
    CycleTypeDesc(CyclePurgeLoad) = "Purge - Load"
    CycleTypeDesc(CycleLoadPurge) = "Load - Purge"
    
    LeakPhaseDesc(LeakPurging) = "Purging"
    LeakPhaseDesc(LeakPressurizing) = "Pressurize"
    LeakPhaseDesc(LeakTesting) = "Testing"
    LeakPhaseDesc(LeakComplete) = "Complete"
    
    LoadPhaseDesc(LoadStarting) = "Starting"
    LoadPhaseDesc(LoadLoading) = "Loading"
    LoadPhaseDesc(LoadComplete) = "Complete"
    LoadPhaseDesc(LoadStopping) = "Stopping"
    LoadPhaseDesc(LoadPause) = "Pause"
    LoadPhaseDesc(LoadPrep) = "Preliminaries"
    
    PurgePhaseDesc(PurgeStarting) = "Starting"
    PurgePhaseDesc(PurgePurging) = "Purging"
    PurgePhaseDesc(PurgeComplete) = "Complete"
    PurgePhaseDesc(PurgeStopping) = "Stopping"
    PurgePhaseDesc(PurgePause) = "Pause"
    
    LoadTypeDesc(NOLOAD) = "No Load"
    LoadTypeDesc(LOADBYTIME) = "Loading for "
    LoadTypeDesc(LOADBYWC) = "Load "
    LoadTypeDesc(LOADBYWEIGHT) = "Loading "
    LoadTypeDesc(LOADBYBREAKTHRU) = "Load "
    LoadTypeDesc(LOADBYFID) = "Load "
    LoadTypeDesc2(NOLOAD) = "No Load"
    LoadTypeDesc2(LOADBYTIME) = " min"
    LoadTypeDesc2(LOADBYWC) = "x WC in "
    LoadTypeDesc2(LOADBYWEIGHT) = " grams"
    LoadTypeDesc2(LOADBYBREAKTHRU) = " gram"
    LoadTypeDesc2(LOADBYFID) = " mg"
    LoadTypeDesc3(NOLOAD) = "No Load"
    LoadTypeDesc3(LOADBYTIME) = " "
    LoadTypeDesc3(LOADBYWC) = " hrs"
    LoadTypeDesc3(LOADBYWEIGHT) = " "
    LoadTypeDesc3(LOADBYBREAKTHRU) = " Breakthrough"
    LoadTypeDesc3(LOADBYFID) = " FID"
    
    StartTypeDesc(STARTNOW) = "Start "
    StartTypeDesc(STARTDELAYED) = "Start after "
    StartTypeDesc(STARTATDATE) = "Start at "
    StartTypeDesc2(STARTNOW) = "Now"
    StartTypeDesc2(STARTDELAYED) = " min"
    StartTypeDesc2(STARTATDATE) = " "
    
    EndMethodDesc(ENDCYCLES) = "End after x cycles "
    EndMethodDesc(ENDWEIGHTCHG) = "End after Stable Weight Chg"
    
    StartMethodDesc(STARTNOW) = "Start without delay"
    StartMethodDesc(STARTDELAYED) = "Start after Delay"
    StartMethodDesc(STARTATDATE) = "Start at specified DateTime"
    
    LeakMethodDesc(NOLEAKCHECK) = "No Leakcheck"
    LeakMethodDesc(LEAKCHECKPRI) = "Pri"
    LeakMethodDesc(LEAKCHECKAUX) = "Aux"
    
    LoadMethodDesc(NOLOAD) = "No Load"
    LoadMethodDesc(LOADBYTIME) = "Load by Time"
    LoadMethodDesc(LOADBYWC) = "Load by W.C."
    LoadMethodDesc(LOADBYWEIGHT) = "Load by Weight"
    LoadMethodDesc(LOADBYBREAKTHRU) = "Load by Breakthrough"
    LoadMethodDesc(LOADBYFID) = "Load by FID"
    
    PurgeMethodDesc(NOPURGE) = "No Purge"
    PurgeMethodDesc(PURGEBYTIME) = "Purge by Time"
    PurgeMethodDesc(PURGEBYVOLUME) = "Purge by Volumes"
    PurgeMethodDesc(PURGEAUXONLY) = "Aux Can Only"
    PurgeMethodDesc(PURGEBYPROFILE) = "Purge By Profile"
    PurgeMethodDesc(PURGEBYWC) = "Purge By W.C."
    PurgeMethodDesc(PURGETOTARGET) = "Purge to Target"
    PurgeMethodDesc(PURGETOUNDOLOAD) = "Purge to Undo Load"
    PurgeMethodDesc(PURGEBYLITERS) = "Purge by Liters"
    
    PurgeTargetDesc(NOTARGET) = "No Target"
    PurgeTargetDesc(TARGETCONTINUOUS) = "Continuous Purge"
    PurgeTargetDesc(TARGETPURGEPAUSE) = "Purge/Pause/Repeat"
    
    PAGtypeDesc(pagNone) = "none"
    PAGtypeDesc(pagAlone) = "StandAlone"
    PAGtypeDesc(pagMaster) = "AK master"
    PAGtypeDesc(pagClient) = "AK client"

    For iCode = 0 To 19
        If iCode < 10 Then
            LeakCanisterDesc(iCode) = "undefined"
            LeakResultDesc(iCode) = "undefined"
        End If
        ReportCodeDesc(iCode) = "undefined"
    Next iCode
    
    LeakCanisterDesc(NOLEAKCHECK) = "no leakcheck"
    LeakCanisterDesc(LEAKCHECKPRI) = "Primary"
    LeakCanisterDesc(LEAKCHECKAUX) = "Aux"
    
    LeakResultDesc(NORESULT) = "no result yet"
    LeakResultDesc(RESULTFAIL_PURGETIMEOUT) = "Failed-Purge Timeout"
    LeakResultDesc(RESULTFAIL_PRESSURETIMEOUT) = "Failed-Pressurize Timeout"
    LeakResultDesc(RESULTFAIL_LEAKRATE) = "Failed-Excessive Leak Rate"
    LeakResultDesc(RESULTABORTAUTO) = "Aborted"
    LeakResultDesc(RESULTABORTOPER) = "Aborted by Operator"
    LeakResultDesc(RESULTGOOD) = "PASSED"
    
    ReportCodeDesc(NORMALUPDATE) = " "
    ReportCodeDesc(LOADBEGIN) = "Beginning of Load"
    ReportCodeDesc(LOADDONE) = "Load Done"
    ReportCodeDesc(PURGEBEGIN) = "Beginning of Purge"
    ReportCodeDesc(PURGEDONE) = "Purge Done"
    ReportCodeDesc(OOTPAUSEBEGIN) = "**** OOT Pause Begins ****"
    ReportCodeDesc(OOTPAUSECLEAR) = "Operator Cleared OOT Pause"
    ReportCodeDesc(LCBEGINPHASE0) = "LeakCheck phase 0 (purging) begins"
    ReportCodeDesc(LCBEGINPHASE1) = "LeakCheck phase 1 (pressurizing) begins"
    ReportCodeDesc(LCBEGINPHASE2) = "LeakCheck phase 2 (testing) begins"
    ReportCodeDesc(LCTESTRESULT) = "LeakCheck Result (pass, fail, aborted)"
    ReportCodeDesc(LCOPERCONTINUE) = "Operator pressed CONTINUE after a LeakCheck failed"
    ReportCodeDesc(LCAUTOCONTINUE) = "Automatic CONTINUE after a LeakCheck failed"

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

Public Sub SetupOpto()
' Function Name:    SetupOpto
' Author:           Analytical Process Programmer         12/2005
' Description:      This routine initializes the Opto DIO & AIO Information Arrays
'                   The routine is called at power up and when the system definitions are changed.
'
'               Type    Code    Module
'               1       100     DI
'               2       180     DO
'               3       12      AI
'               4       133     AO
'               5       5       AI TC Type J
'               6       8       AI TC Type K
'               7       10      AI RTD 100 Ohm
'
'
Dim prg As Integer
Dim stn As Integer
Dim slot As Integer
Dim addr As Integer
Dim chan As Integer
Dim func As Integer
Dim baseaddr As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 622

' Refresh Opto Channel Type

For stn = 0 To NR_STN
    baseaddr = stn * 4
    Select Case Node_Info(stn)
        Case 8, 12, 16
            ' 8-Slot, 12-Slot & 16-Slot Boards
            For slot = 0 To MAX_SLOT
                Select Case slot
                    Case 0, 1, 2, 3
                        Select Case Opto_Info(baseaddr, slot)
                            Case optotypeDI, optotypeDO
                                ' Module is a Digital Module @ Baseaddr + 0
                                OptoDIO(baseaddr + 0, ((4 * slot) + 0)).Type = Opto_Info(baseaddr, slot)
                                OptoDIO(baseaddr + 0, ((4 * slot) + 1)).Type = Opto_Info(baseaddr, slot)
                                OptoDIO(baseaddr + 0, ((4 * slot) + 2)).Type = Opto_Info(baseaddr, slot)
                                OptoDIO(baseaddr + 0, ((4 * slot) + 3)).Type = Opto_Info(baseaddr, slot)
                            Case Else
                                ' Module is an Analog Module @ Baseaddr + 2
                                OptoAIO(baseaddr + 2, ((2 * slot) + 0)).Type = Opto_Info(baseaddr, slot)
                                OptoAIO(baseaddr + 2, ((2 * slot) + 1)).Type = Opto_Info(baseaddr, slot)
                        End Select
                    Case 4, 5, 6, 7
                        Select Case Opto_Info(baseaddr, slot)
                            Case optotypeDI, optotypeDO
                                ' Module is a Digital Module @ Baseaddr + 1
                                OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 0)).Type = Opto_Info(baseaddr, slot)
                                OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 1)).Type = Opto_Info(baseaddr, slot)
                                OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 2)).Type = Opto_Info(baseaddr, slot)
                                OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 3)).Type = Opto_Info(baseaddr, slot)
                            Case Else
                                ' Module is an Analog Module @ Baseaddr + 2
                                OptoAIO(baseaddr + 2, ((2 * (slot - 4)) + 8)).Type = Opto_Info(baseaddr, slot)
                                OptoAIO(baseaddr + 2, ((2 * (slot - 4)) + 9)).Type = Opto_Info(baseaddr, slot)
                        End Select
                    Case 8, 9, 10, 11, 12, 13, 14, 15
                        ' Module must be an Analog Module @ Baseaddr + 3
                        OptoAIO(baseaddr + 3, ((2 * (slot - 8)) + 0)).Type = Opto_Info(baseaddr, slot)
                        OptoAIO(baseaddr + 3, ((2 * (slot - 8)) + 1)).Type = Opto_Info(baseaddr, slot)
                End Select
            Next slot
        Case Else
            ' No Board
            For slot = 0 To MAX_SLOT
                Select Case slot
                    Case 0, 1, 2, 3
                        ' Module set to a Digital Module @ Baseaddr
                        OptoDIO(baseaddr, ((4 * slot) + 0)).Type = 0
                        OptoDIO(baseaddr, ((4 * slot) + 1)).Type = 0
                        OptoDIO(baseaddr, ((4 * slot) + 2)).Type = 0
                        OptoDIO(baseaddr, ((4 * slot) + 3)).Type = 0
                    Case 4, 5, 6, 7
                        ' Module set to a Digital Module @ Baseaddr + 1
                        OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 0)).Type = 0
                        OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 1)).Type = 0
                        OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 2)).Type = 0
                        OptoDIO(baseaddr + 1, ((4 * (slot - 4)) + 3)).Type = 0
                    Case 8, 9, 10, 11, 12, 13, 14, 15
                        ' Module set to an Analog Module @ Baseaddr + 3
                        OptoAIO(baseaddr + 3, ((2 * (slot - 8)) + 0)).Type = 0
                        OptoAIO(baseaddr + 3, ((2 * (slot - 8)) + 1)).Type = 0
                End Select
            Next slot
    End Select
Next stn

ChgErrModule 11, 623

' *********************************
' Refresh Opto Channel Information
' *********************************

'   Initialize Channel Information
For addr = 0 To MAX_ADDR
    For chan = 0 To MAX_CHAN
        OptoChanDesc(addr, chan) = "Spare " & Format(addr, "00") & "/" & Format(chan, "00")
        Map_AIO(addr, chan).desc = OptoChanDesc(addr, chan)
        Map_AIO(addr, chan).Type = optotypeNoModule
        Map_AIO(addr, chan).TypeCode = optocodeUndefined
        Map_AIO(addr, chan).RawValue = 0
        Map_AIO(addr, chan).VdcMax = 0
        Map_AIO(addr, chan).VdcMin = 0
        Map_AIO(addr, chan).EuMax = 0
        Map_AIO(addr, chan).EuMin = 0
        Map_AIO(addr, chan).EUValue = 0
        Map_DIO(addr, chan).desc = OptoChanDesc(addr, chan)
        Map_DIO(addr, chan).Type = optotypeNoModule
        Map_DIO(addr, chan).TypeCode = optocodeUndefined
        Map_DIO(addr, chan).RawValue = False
        Map_DIO(addr, chan).UseInverse = False
        Map_DIO(addr, chan).Value = False
    Next chan
Next addr

'   Get Common Analog Information
For func = 0 To MAX_ANA_COM
    If Com_AIO(func).addr <> 0 Or Com_AIO(func).chan <> 0 Then
        addr = Com_AIO(func).addr
        chan = Com_AIO(func).chan
        OptoChanDesc(addr, chan) = Com_AnaDef(func).desc
        Map_AIO(addr, chan).desc = OptoChanDesc(addr, chan)
        Map_AIO(addr, chan).Type = OptoAIO(addr, chan).Type
        Map_AIO(addr, chan).TypeCode = OptoAIO(addr, chan).TypeCode
        Map_AIO(addr, chan).VdcMax = Com_AIO(func).VdcMax
        Map_AIO(addr, chan).VdcMin = Com_AIO(func).VdcMin
        Map_AIO(addr, chan).EuMax = Com_AIO(func).EuMax
        Map_AIO(addr, chan).EuMin = Com_AIO(func).EuMin
    End If
Next func
'   Get Common Digital Information
For func = 0 To MAX_DIG_COM
    If Com_DIO(func).addr <> 0 Or Com_DIO(func).chan <> 0 _
        Or _
        func = 1 And Com_DIO(func).addr = 0 And Com_DIO(func).chan = 0 Then
        addr = Com_DIO(func).addr
        chan = Com_DIO(func).chan
        OptoChanDesc(addr, chan) = Com_DigDef(func).desc
        Map_DIO(addr, chan).desc = OptoChanDesc(addr, chan)
        Map_DIO(addr, chan).Type = OptoDIO(addr, chan).Type
        Map_DIO(addr, chan).TypeCode = OptoDIO(addr, chan).TypeCode
        Map_DIO(addr, chan).UseInverse = Com_DIO(func).UseInverse
    End If
Next func

'   Get PurgeAir Analog Information
For prg = 1 To NR_PRGAIR
    For func = 0 To MAX_ANA_PRG
        If Prg_AIO(prg, func).addr <> 0 Or Prg_AIO(prg, func).chan <> 0 Then
            addr = Prg_AIO(prg, func).addr
            chan = Prg_AIO(prg, func).chan
            OptoChanDesc(addr, chan) = Prg_AnaDef(func).desc & " #" & Format(prg, "0")
            Map_AIO(addr, chan).desc = OptoChanDesc(addr, chan)
            Map_AIO(addr, chan).Type = OptoAIO(addr, chan).Type
            Map_AIO(addr, chan).TypeCode = OptoAIO(addr, chan).TypeCode
            Map_AIO(addr, chan).VdcMax = Prg_AIO(prg, func).VdcMax
            Map_AIO(addr, chan).VdcMin = Prg_AIO(prg, func).VdcMin
            Map_AIO(addr, chan).EuMax = Prg_AIO(prg, func).EuMax
            Map_AIO(addr, chan).EuMin = Prg_AIO(prg, func).EuMin
        End If
    Next func
Next prg
'   Get PurgeAir Digital Information
For prg = 1 To NR_PRGAIR
    For func = 0 To MAX_DIG_PRG
        If Prg_DIO(prg, func).addr <> 0 Or Prg_DIO(prg, func).chan <> 0 Then
            addr = Prg_DIO(prg, func).addr
            chan = Prg_DIO(prg, func).chan
            OptoChanDesc(addr, chan) = Prg_DigDef(func).desc & " #" & Format(prg, "0")
            Map_DIO(addr, chan).desc = OptoChanDesc(addr, chan)
            Map_DIO(addr, chan).Type = OptoDIO(addr, chan).Type
            Map_DIO(addr, chan).TypeCode = OptoDIO(addr, chan).TypeCode
            Map_DIO(addr, chan).UseInverse = Prg_DIO(prg, func).UseInverse
        End If
    Next func
Next prg

'   Get Station Analog Information
For stn = 1 To NR_STN
    For func = 0 To MAX_ANA_STN
        If Stn_AIO(stn, func).addr <> 0 Or Stn_AIO(stn, func).chan <> 0 Then
            addr = Stn_AIO(stn, func).addr
            chan = Stn_AIO(stn, func).chan
            OptoChanDesc(addr, chan) = "s" & Format(stn, 0) & " " & Stn_AnaDef(func).desc
            Map_AIO(addr, chan).desc = OptoChanDesc(addr, chan)
            Map_AIO(addr, chan).Type = OptoAIO(addr, chan).Type
            Map_AIO(addr, chan).TypeCode = OptoAIO(addr, chan).TypeCode
            Map_AIO(addr, chan).VdcMax = Stn_AIO(stn, func).VdcMax
            Map_AIO(addr, chan).VdcMin = Stn_AIO(stn, func).VdcMin
            Map_AIO(addr, chan).EuMax = Stn_AIO(stn, func).EuMax
        Map_AIO(addr, chan).EuMin = Stn_AIO(stn, func).EuMin
        End If
    Next func
Next stn
'   Get Station Digital Information
For stn = 1 To NR_STN
    For func = 0 To MAX_DIG_STN
        If Stn_DIO(stn, func).addr <> 0 Or Stn_DIO(stn, func).chan <> 0 Then
            addr = Stn_DIO(stn, func).addr
            chan = Stn_DIO(stn, func).chan
            OptoChanDesc(addr, chan) = "s" & Format(stn, 0) & " " & Stn_DigDef(func).desc
            Map_DIO(addr, chan).desc = OptoChanDesc(addr, chan)
            Map_DIO(addr, chan).Type = OptoDIO(addr, chan).Type
            Map_DIO(addr, chan).TypeCode = OptoDIO(addr, chan).TypeCode
            Map_DIO(addr, chan).UseInverse = Stn_DIO(stn, func).UseInverse
        End If
    Next func
Next stn

   
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

Public Sub SetupStations()
' Function Name:    SetupStations
' Author:           Analytical Process Programmer         12/2005
' Description:      This routine initializes the Station Info Arrays
'                   The routine is called at power up one time.
'
'
'                   Station Types
'                   1 = Regular
'                   2 = ORVR
'                   3 = Live Fuel
'                   4 = ORVR with 2 MFC's per Load Gas
'                   6 = Regular and  Live Fuel
'                   9 = Dummy
'
'                   ADF_TankTypes
'                   0 = None (i.e. No I/O)                                              (first=before 2004)
'                   1 = ADF #1; Pump, Drain, Fill                                       (first=MarkIV)
'                   2 = ADF #2; Pump, Drain, Fill, Vapor                                (none so far)
'                  11 = ADF #1; Pump, Drain, Fill, with Heater                          (none so far)
'                  12 = ADF #2; Pump, Drain, Fill, Vapor, Bypass, N2Purge with Heater   (first=Mahle)
'                  22 = ADF #2; Pump, Drain, Fill, Vapor, Bypass                        (first=Chrysler)
'
'                   Aspirator Number
'                   1-9 = use PurgeAir Source#
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 11, 600
Dim iStn As Integer

    ' Initialize Station Information on Starting the system from power up/start of job
    
    systemhasAUTODRAINFILL = False
    systemhasADF_HEATER = False
    systemhasADF_WATERBATH = False
    systemhasBUTANE = False
    systemhasLIVEFUEL = False
    systemhasLEAKTEST = False
    systemhasORVR2 = False
    
    For iStn = 1 To NR_STN
        ' station aspirator # valid ??
        If STN_INFO(iStn).AspiratorNum < 1 Then STN_INFO(iStn).AspiratorNum = 1
        If STN_INFO(iStn).AspiratorNum > MAX_PRG Then STN_INFO(iStn).AspiratorNum = 1
        ' using Butane ??
        If STN_INFO(iStn).Type = STN_REGULAR_TYPE Then systemhasBUTANE = True
        If STN_INFO(iStn).Type = STN_LIVEREG_TYPE Then systemhasBUTANE = True
        If STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE Then systemhasBUTANE = True
        If STN_INFO(iStn).Type = STN_ORVR_TYPE Then systemhasBUTANE = True
        If STN_INFO(iStn).Type = STN_ORVR2_TYPE Then systemhasBUTANE = True
        ' using LiveFuel ??
        If STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE Then systemhasLIVEFUEL = True
        If STN_INFO(iStn).Type = STN_LIVEREG_TYPE Then systemhasLIVEFUEL = True
        If STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE Then systemhasLIVEFUEL = True
        ' LeakTest Station ??
        If STN_INFO(iStn).Type = STN_LEAKTEST_TYPE Then systemhasLEAKTEST = True
        ' using Two Butane&Nitrogen MFC's ??
        If STN_INFO(iStn).Type = STN_ORVR2_TYPE Then systemhasORVR2 = True
        If STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE Then systemhasORVR2 = True
        ' using LiveFuel AutoDrainFill ??
        If ((STN_INFO(iStn).ADF_TANKTYPE > 0) And (STN_INFO(iStn).ADF_TANKTYPE < 90)) Then systemhasAUTODRAINFILL = True
        ' using a LiveFuel Heater ??
        If ((STN_INFO(iStn).ADF_TANKTYPE > 10) And (STN_INFO(iStn).ADF_TANKTYPE <= 20)) Then systemhasADF_HEATER = True
        ' using a LiveFuel WaterBath ??
        If ((STN_INFO(iStn).ADF_TANKTYPE = 90)) Then systemhasADF_WATERBATH = True
    Next iStn
    
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

Public Sub SetupFunctionArrays()
' Populate the Various Arrays for the Digital & Analog Functions
'
Dim index1, Index2, index3 As Integer

' COMMON ANALOG IO FUNCTIONS
' Initialize the Common Analog IO Function Definition Array
For index1 = 0 To MAX_ANA_COM
    Com_AnaDef(index1).desc = "com analog unused " & Format(index1, "00")
    Com_AnaDef(index1).UsedIn = False
Next index1

index3 = acPasTempSensor
Com_AnaDef(index3).desc = "PurgeAir Temperature"
Com_AnaDef(index3).UsedIn = False

index3 = acPasHumiditySensor
Com_AnaDef(index3).desc = "PurgeAir Humidity"
Com_AnaDef(index3).UsedIn = False

index3 = acAmbTempSensor
Com_AnaDef(index3).desc = "Ambient Air Temperature"
Com_AnaDef(index3).UsedIn = False

index3 = acAmbHumiditySensor
Com_AnaDef(index3).desc = "Ambient Air Humidity"
Com_AnaDef(index3).UsedIn = False

index3 = acComnPressSensor
Com_AnaDef(index3).desc = "LeakCheck Pressure"
Com_AnaDef(index3).UsedIn = False

index3 = acAmbBaroSensor
Com_AnaDef(index3).desc = "Ambient Barometer"
Com_AnaDef(index3).UsedIn = False

index3 = acCustCalDevice
Com_AnaDef(index3).desc = "Calibration Device"
Com_AnaDef(index3).UsedIn = False

index3 = acCommonTC1
Com_AnaDef(index3).desc = "Common TC #1"
Com_AnaDef(index3).UsedIn = False

index3 = acCommonTC2
Com_AnaDef(index3).desc = "Common TC #2"
Com_AnaDef(index3).UsedIn = False

index3 = acCommonTC3
Com_AnaDef(index3).desc = "Common TC #3"
Com_AnaDef(index3).UsedIn = False

index3 = acCommonTC4
Com_AnaDef(index3).desc = "Common TC #4"
Com_AnaDef(index3).UsedIn = False

index3 = acCommonTC5
Com_AnaDef(index3).desc = "Common TC #5"
Com_AnaDef(index3).UsedIn = False

index3 = acCommonTC6
Com_AnaDef(index3).desc = "Common TC #6"
Com_AnaDef(index3).UsedIn = False

index3 = acPASMoistCntrlOut
Com_AnaDef(index3).desc = "PAS Moisture Pump"
Com_AnaDef(index3).UsedIn = False



' PURGEAIR SOURCE ANALOG IO FUNCTIONS
' Initialize the PurgeAir Source Analog IO Function Definition Array
For index1 = 0 To MAX_ANA_PRG
    Prg_AnaDef(index1).desc = "prg analog unused " & Format(index1, "00")
    Prg_AnaDef(index1).UsedIn = False
Next index1

index3 = apTemp
Prg_AnaDef(index3).desc = "PurgeAir Temperature"
Prg_AnaDef(index3).UsedIn = False

index3 = apHumidity
Prg_AnaDef(index3).desc = "PurgeAir Humidity"
Prg_AnaDef(index3).UsedIn = False

index3 = apBaro
Prg_AnaDef(index3).desc = "PurgeAir Baro"
Prg_AnaDef(index3).UsedIn = False

' STATION ANALOG IO FUNCTIONS
' Initialize the Station Analog IO Function Definition Array
For index1 = 0 To MAX_ANA_STN
    Stn_AnaDef(index1).desc = "stn analog unused " & Format(index1, "00")
    For Index2 = 1 To MAX_STNTYPE
        Stn_AnaDef(index1).UsedIn(Index2) = False
    Next Index2
Next index1

index3 = asNitrogenFlowSP
Stn_AnaDef(index3).desc = "NitrogenMFC SetPoint"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = asButaneFlowSP
Stn_AnaDef(index3).desc = "ButaneMFC SetPoint"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asPurgeAirFlowSP
Stn_AnaDef(index3).desc = "PurgeAirMFC SetPoint"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asNitrogenORVRFlowSP
Stn_AnaDef(index3).desc = "NitrogenORVRMFC SP"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asButaneORVRFlowSP
Stn_AnaDef(index3).desc = "ButaneORVRMFC SP"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asLiveFuelVaporFlowSP
Stn_AnaDef(index3).desc = "LiveFuelMFC SetPoint"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asLiveFuelVaporORVRFlowSP
Stn_AnaDef(index3).desc = "LiveFuelORVRMFC SP"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asPurgeDiffPress
Stn_AnaDef(index3).desc = "Purge Diff.Press"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asNitrogenFlow
Stn_AnaDef(index3).desc = "NitrogenMFC Flow"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = asButaneFlow
Stn_AnaDef(index3).desc = "ButaneMFC Flow"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asPurgeAirFlow
Stn_AnaDef(index3).desc = "PurgeAirMFC Flow"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asNitrogenORVRFlow
Stn_AnaDef(index3).desc = "NitORVRMFC Flow"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asButaneORVRFlow
Stn_AnaDef(index3).desc = "ButORVRMFC Flow"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asLoadPressure
Stn_AnaDef(index3).desc = "Load Pressure"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asPurgeOvenTempSP
Stn_AnaDef(index3).desc = "PurgeOven Temp SP"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asPurgeOvenTemp
Stn_AnaDef(index3).desc = "PurgeOven Temp"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asLiveFuelVaporFlow
Stn_AnaDef(index3).desc = "LiveFuelMFC Flow"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asLiveFuelVaporORVRFlow
Stn_AnaDef(index3).desc = "LiveFuelORVRMFC"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asFuelHeaterTemp
Stn_AnaDef(index3).desc = "LiveFuel Heater Temp"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asFuelTankTemp
Stn_AnaDef(index3).desc = "LiveFuel Tank Temp"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asFuelTankLevel
Stn_AnaDef(index3).desc = "LiveFuel Tank Level"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asFuelVaporTemp
Stn_AnaDef(index3).desc = "LiveFuel Vapor Temp"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asStorageTankLevel
Stn_AnaDef(index3).desc = "Storage Tank Level"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = asLtInletPress
Stn_AnaDef(index3).desc = "LeakTest Inlet Press"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = asLtN2Temp
Stn_AnaDef(index3).desc = "LeakTest Nitrogen Tmp"
Stn_AnaDef(index3).UsedIn(STN_REGULAR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEREG_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_ORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = False
Stn_AnaDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True


' COMMON DIGITAL IO FUNCTIONS

' Initialize the Common Digital IO Function Definition Array
For index1 = 0 To MAX_DIG_COM
    Com_DigDef(index1).desc = "com digital unused " & Format(index1, "00")
    Com_DigDef(index1).UsedIn = False
Next index1

Com_DigDef(icHornSilencePB).desc = "Horn Silence PB"
Com_DigDef(icHornSilencePB).UsedIn = True

Com_DigDef(icExhaustFlowFS).desc = "Exhaust Flow Switch"
Com_DigDef(icExhaustFlowFS).UsedIn = True

Com_DigDef(icEStopSw).desc = "E-Stop Input"
Com_DigDef(icEStopSw).UsedIn = True

Com_DigDef(icDoorSw).desc = "Door Switch Input"
If USINGDOOROPEN Then Com_DigDef(icDoorSw).UsedIn = True

Com_DigDef(icMaintSw).desc = "Maint Mode Input"
Com_DigDef(icMaintSw).UsedIn = True

Com_DigDef(ic20LelGasSw).desc = "20% LEL Alarm Input"
Com_DigDef(ic20LelGasSw).UsedIn = True

Com_DigDef(icUpsFaultSw).desc = "UPS Fault Input"
If USINGUPS <> 0 Then Com_DigDef(icUpsFaultSw).UsedIn = True

Com_DigDef(icUpsActiveSw).desc = "UPS Active Input"
If USINGUPS <> 0 Then Com_DigDef(icUpsActiveSw).UsedIn = True

Com_DigDef(icSystemVacSw).desc = "System Vac Sw Input"
Com_DigDef(icSystemVacSw).UsedIn = True

Com_DigDef(icAlarmBeacon).desc = "Alarm Beacon"
Com_DigDef(icAlarmBeacon).UsedIn = True

Com_DigDef(icAlarmHorn).desc = "Alarm Horn"
Com_DigDef(icAlarmHorn).UsedIn = True

Com_DigDef(icPauseLT).desc = "Pause LT"
Com_DigDef(icPauseLT).UsedIn = True

Com_DigDef(icButaneShutoffSol).desc = "Butane Shutoff Sol"
If systemhasBUTANE Then Com_DigDef(icButaneShutoffSol).UsedIn = True

Com_DigDef(icLeakCheckExhaustSol).desc = "Leak Check Exhaust Sol"
If USINGLEAKCHECKEXHAUSTSOL Then Com_DigDef(icLeakCheckExhaustSol).UsedIn = True

Com_DigDef(icExtAlmContactSw).desc = "Customer Ext Alarm Sw"
If USING_EXT_CONTACTS Then Com_DigDef(icExtAlmContactSw).UsedIn = True

Com_DigDef(icCustLowGasSw).desc = "Customer Low Gas Sw"
If USINGCUSTOMERLOWGAS Then Com_DigDef(icCustLowGasSw).UsedIn = True

Com_DigDef(icPurgeRequestOut).desc = "Request to PAS"
Com_DigDef(icPurgeRequestOut).UsedIn = False

Com_DigDef(icPurgeReadyIn).desc = "Ready from PAS"
Com_DigDef(icPurgeReadyIn).UsedIn = False

Com_DigDef(icPASPowerOnIn).desc = "PAS Power ON"
If USINGPASLOCALCONTROL Then Com_DigDef(icPASPowerOnIn).UsedIn = True

Com_DigDef(icPASReadyOut).desc = "PAS is Ready"
If USINGPASLOCALCONTROL Then Com_DigDef(icPASReadyOut).UsedIn = True

Com_DigDef(icPASisRunningIn).desc = "PAS in Running"
If USINGPASLOCALCONTROL Then Com_DigDef(icPASisRunningIn).UsedIn = True

Com_DigDef(icPASHeaterSSR).desc = "PAS Heater SSR"
If USINGPASLOCALCONTROL Then Com_DigDef(icPASHeaterSSR).UsedIn = True

Com_DigDef(icPurgeDryAirSupplySol).desc = "Purge DryAir Supply Sol"
If USINGDRYPURGEAIR Then Com_DigDef(icPurgeDryAirSupplySol).UsedIn = True

Com_DigDef(icPurgeAirSourceSelectSol).desc = "Purge AirSource Select"
If USINGDRYPURGEAIR Then Com_DigDef(icPurgeAirSourceSelectSol).UsedIn = True

Com_DigDef(icLiveFuelPurgePS).desc = "Live Fuel N2 Purge PS"
If systemhasLIVEFUEL Then Com_DigDef(icLiveFuelPurgePS).UsedIn = True




' PURGEAIR SOURCE DIGITAL IO FUNCTIONS
' Initialize the PurgeAir Source Digital IO Function Definition Array
For index1 = 0 To MAX_DIG_PRG
    Prg_DigDef(index1).desc = "prg digital unused " & Format(index1, "00")
    Prg_DigDef(index1).UsedIn = False
Next index1

index3 = ipPiabSol
Prg_DigDef(index3).desc = "PurgeAir Piab Sol"
Prg_DigDef(index3).UsedIn = True

index3 = ipPurgeVacuumSw
Prg_DigDef(index3).desc = "PurgeAir Vacuum Sw"
Prg_DigDef(index3).UsedIn = True

index3 = ipAuxAirSol
Prg_DigDef(index3).desc = "PurgeAir Aux Air Sol"
Prg_DigDef(index3).UsedIn = False

index3 = ipPosPrsPrgSol
Prg_DigDef(index3).desc = "PurgeAir Pressure Purge Sol"
If USINGPRESSUREPURGE Then Prg_DigDef(index3).UsedIn = True


' STATION DIGITAL IO FUNCTIONS
' Initialize the Station Digital IO Function Definition Array
For index1 = 0 To MAX_DIG_STN
    Stn_DigDef(index1).desc = "stn digital unused " & Format(index1, "00")
    For Index2 = 1 To MAX_STNTYPE
        Stn_DigDef(index1).UsedIn(Index2) = False
    Next Index2
Next index1

index3 = isNitrogenSol
Stn_DigDef(index3).desc = "Nitrogen Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = isButaneSol
Stn_DigDef(index3).desc = "Butane Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isNitrogenOrvrSol
Stn_DigDef(index3).desc = "NitrogenORVR Sol"
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isButaneOrvrSol
Stn_DigDef(index3).desc = "ButaneORVR Sol"
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isPurgeSol
Stn_DigDef(index3).desc = "Purge Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isPriDirectionSol
Stn_DigDef(index3).desc = "Pri Direction Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isAuxCanVentSol
Stn_DigDef(index3).desc = "Aux Can (Vent) Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isLeakCheckSol
Stn_DigDef(index3).desc = "Leak Check Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isAuxPurgeSol
Stn_DigDef(index3).desc = "Aux Purge Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isPriAuxVentSol
Stn_DigDef(index3).desc = "Pri Aux Vent Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isAuxDirectionSol
Stn_DigDef(index3).desc = "Aux Direction Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isAuxLeakCheckSol
Stn_DigDef(index3).desc = "Aux LeakCheck Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isPauseLT
Stn_DigDef(index3).desc = "Pause LT"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = isIdleLT
Stn_DigDef(index3).desc = "Idle LT"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = isPriSeriesPurgeSol
Stn_DigDef(index3).desc = "Pri SeriesPurge Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isAuxSeriesPurgeSol
Stn_DigDef(index3).desc = "Aux SeriesPurge Sol"
Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isLoadShift2Sol
Stn_DigDef(index3).desc = "Shift 2 Load Sol"
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isVentShift2Sol
Stn_DigDef(index3).desc = "Shift 2 Vent Sol"
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isPurgeShift2Sol
Stn_DigDef(index3).desc = "Shift 2 Purge Sol"
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 1 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isLiveFuelOrvrSol
Stn_DigDef(index3).desc = "LiveFuelORVR Sol"
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False


index3 = isLoadTypeSelectSol
Stn_DigDef(index3).desc = "LoadType Select Sol"
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isLiveFuelSol
Stn_DigDef(index3).desc = "Live Fuel Sol"
Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Stn_DigDef(index3).UsedIn(STN_LEAKTEST_TYPE) = False

index3 = isFuelVentSol
Stn_DigDef(index3).desc = "Tank Vent Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelRecircSol
Stn_DigDef(index3).desc = "Tank Recirc Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelDrainSol
Stn_DigDef(index3).desc = "Tank Drain Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelFillSol
Stn_DigDef(index3).desc = "Tank Fill Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelPressSol
Stn_DigDef(index3).desc = "Tank Pressurize Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelVaporSol
Stn_DigDef(index3).desc = "Tank Vapor Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelPumpMotor
Stn_DigDef(index3).desc = "Tank Pump"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelHiHiLevelLS
Stn_DigDef(index3).desc = "Tank HiHi LS"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelHighLevelLS
Stn_DigDef(index3).desc = "Tank High LS"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelLowLevelLS
Stn_DigDef(index3).desc = "Tank Low LS"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelHeaterSSR
Stn_DigDef(index3).desc = "Tank Heater SSR"
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelOverTempSw
Stn_DigDef(index3).desc = "Tank OverTemp Sw"
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isSheathOverTempSw
Stn_DigDef(index3).desc = "Sheath OverTemp Sw"
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelSafetyLevelLS
Stn_DigDef(index3).desc = "Tank Safety LS"
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasADF_HEATER Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelPressPS
Stn_DigDef(index3).desc = "Tank N2 Press Sw"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isStorageHiHiLevelLS
Stn_DigDef(index3).desc = "Storage HiHi LS"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isStorageLowLevelLS
Stn_DigDef(index3).desc = "Storage Low LS"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isStorageDrainSol
Stn_DigDef(index3).desc = "Storage Drain Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isStorageFillSol
Stn_DigDef(index3).desc = "Storage Fill Sol"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isStorageFillRequest
Stn_DigDef(index3).desc = "Storage Fill Req"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isFuelOverTempResetOut
Stn_DigDef(index3).desc = "FuelOT Reset Out"
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If systemhasAUTODRAINFILL Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = isCanVentAlarmSw
Stn_DigDef(index3).desc = "CanVent Alarm Sw"
If USINGCANVENTALARM Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USINGCANVENTALARM Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If USINGCANVENTALARM Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USINGCANVENTALARM Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USINGCANVENTALARM Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USINGCANVENTALARM Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isPurgeLocationSupplySelectSol
Stn_DigDef(index3).desc = "Purge Location Supply  Sol"
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isPurgeLocationVentSelectSol
Stn_DigDef(index3).desc = "Purge Location Vent  Sol"
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USINGPURGEOVEN Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isLoadShift3Sol
Stn_DigDef(index3).desc = "Shift 3 Load Sol"
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isVentShift3Sol
Stn_DigDef(index3).desc = "Shift 3 Vent Sol"
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isPurgeShift3Sol
Stn_DigDef(index3).desc = "Shift 3 Purge Sol"
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 2 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isLoadShift4Sol
Stn_DigDef(index3).desc = "Shift 4 Load Sol"
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isVentShift4Sol
Stn_DigDef(index3).desc = "Shift 4 Vent Sol"
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isPurgeShift4Sol
Stn_DigDef(index3).desc = "Shift 4 Purge Sol"
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If NR_SHIFT > 3 Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = isAuxOutput1
Stn_DigDef(index3).desc = "Aux Output 1"
If (Len(DESC_AUX_OUTPUT1) > 3) Then Stn_DigDef(index3).desc = DESC_AUX_OUTPUT1
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True

index3 = isAuxOutput2
Stn_DigDef(index3).desc = "Aux Output 2"
If (Len(DESC_AUX_OUTPUT2) > 3) Then Stn_DigDef(index3).desc = DESC_AUX_OUTPUT2
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True

index3 = isAuxOutput3
Stn_DigDef(index3).desc = "Aux Output 3"
If (Len(DESC_AUX_OUTPUT3) > 3) Then Stn_DigDef(index3).desc = DESC_AUX_OUTPUT3
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True

index3 = isAuxOutput4
Stn_DigDef(index3).desc = "Aux Output 4"
If (Len(DESC_AUX_OUTPUT4) > 3) Then Stn_DigDef(index3).desc = DESC_AUX_OUTPUT4
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_REGULAR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_ORVR2_TYPE) = True
If USING_AUX_OUTPUTS Then Stn_DigDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True


PurgeProfileStepDesc(NOSTEP) = "undefined"
PurgeProfileStepDesc(STEPSTEP) = "Step SP"
PurgeProfileStepDesc(STEPRAMP) = "Ramp SP"
PurgeProfileStepDesc(STEPLAST) = "Last Step"

' MFC FUNCTIONS
' Initialize the Mfc Definition Array
For index1 = 0 To MAXMFC
    Mfc_FunDef(index1).desc = "mfc unused " & Format(index1, "00")
    For Index2 = 1 To MAX_STNTYPE
        Mfc_FunDef(index1).UsedIn(Index2) = False
    Next Index2
Next index1

index3 = MFCBUTANE
Mfc_FunDef(index3).desc = "Butane MFC"
Mfc_FunDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = MFCNITROGEN
Mfc_FunDef(index3).desc = "Nitrogen MFC"
Mfc_FunDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR2_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LEAKTEST_TYPE) = True

index3 = MFCPURGEAIR
Mfc_FunDef(index3).desc = "PurgeAir MFC"
Mfc_FunDef(index3).UsedIn(STN_REGULAR_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = MFCLIVEFUEL
Mfc_FunDef(index3).desc = "LiveFuel MFC"
Mfc_FunDef(index3).UsedIn(STN_LIVEFUEL_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEREG_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True

index3 = MFCORVRBUT
Mfc_FunDef(index3).desc = "ORVR Butane MFC"
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = MFCORVRNIT
Mfc_FunDef(index3).desc = "ORVR Nitrogen MFC"
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True
Mfc_FunDef(index3).UsedIn(STN_ORVR2_TYPE) = True

index3 = MFCORVRPRG
Mfc_FunDef(index3).desc = "ORVR PurgeAir MFC"
' not used so far

index3 = MFCORVRLIVE
Mfc_FunDef(index3).desc = "ORVR LiveFuel MFC"
Mfc_FunDef(index3).UsedIn(STN_LIVEORVR2_TYPE) = True


End Sub

