Attribute VB_Name = "Module6"
' error Module 6 ''''''''''''''''''''program LOGS.bas ''''''''''''''''
Option Explicit
'
Private daodb36 As DAO.Database
Private rS As DAO.Recordset
Dim sPath As String

Sub ALM_Write(ByVal Index As Integer, ByVal index2 As Integer, ByVal Comment As String)
'
' Function Name:    ALM_Write
' Author:           DJP     8/8/96
' Description:      Updates the data file with alarm info
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 1
Dim dbDbase As Database
Dim rsTable As Recordset
  ' update reports for stations out of tolerance condition/alarm change
If ((Index = 0) Or (index2 = 0)) Then
    ' Common
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
    Set rsTable = dbDbase.OpenRecordset("EventLog")
    rsTable.AddNew
    rsTable("Time") = Now()
    rsTable("Comment") = Left(Comment, 255)
    rsTable.Update
    rsTable.Close
    dbDbase.Close
ElseIf (Len(StationControl(Index, index2).DBFile) <= 3) Then
    ' station (no open job)
    Comment = "Station#" & Format(Index, "#0") & " Shift#" & Format(index2, "0") & " - " & Comment
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
    Set rsTable = dbDbase.OpenRecordset("EventLog")
    rsTable.AddNew
    rsTable("Time") = Now()
    rsTable("Comment") = Left(Comment, 255)
    rsTable.Update
    rsTable.Close
    dbDbase.Close
Else
    ' valid stn/shift & job dbfile name
    Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
    Set rsTable = dbDbase.OpenRecordset("Alarm")
    rsTable.AddNew
    rsTable("Course") = StationControl(Index, index2).Course
    rsTable("Time") = Now()
    rsTable("Comment") = Left(Comment, 255)
    rsTable.Update
    rsTable.Close
    dbDbase.Close
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

Sub FLog_Write(Comment As String)
'
' Function Name:    FLog_Write
' Author:           DJP     8/8/96
' Description:      Updates the data file with file deletion info
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 2
Dim a As Integer
a = Len(Comment)
Dim dbDbase As Database
Dim rsTable As Recordset

' update reports for stations out of tolerance condition change
  Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
  Set rsTable = dbDbase.OpenRecordset("filelog")
  rsTable.AddNew
  rsTable("Time") = Now()
  rsTable("Comment") = Comment
  rsTable.Update
  rsTable.Close
  dbDbase.Close
  
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

Sub OOT_Write(Index As Integer, index2 As Integer, Comment As String)
'
' Function Name:    OOT_Write
' Author:           DJP     8/8/96
' Description:      Updates the data file with out of tolerance info
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 3

Dim dbDbase As Database
Dim rsTable As Recordset
  Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile) ' set station
  Set rsTable = dbDbase.OpenRecordset("Tolerance")
  rsTable.AddNew
  rsTable("Course") = StationControl(Index, index2).Course
  rsTable("Time") = Now()
  rsTable("Comment") = Comment
  rsTable.Update
  rsTable.Close
  dbDbase.Close
  
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

Sub AddNew_Joblist( _
    ByVal JobNumber As String, _
    ByVal JobDesc As String, _
    ByVal StartTime As Date, _
    ByVal station As Integer, _
    ByVal Shift As Integer, _
    ByVal ReportFilename As String)
'
' Function Name:    AddNew_Joblist
' Created By:       MMW
' This function adds a new Job to the Job List.
'
Dim sPath As String
Dim dbDbase As Database
Dim rsTable As Recordset

    sPath = FILEPATH_sysdbf & DATAMASTER
    
    Set daodb36 = DBEngine.OpenDatabase(sPath)
    Set rS = daodb36.OpenRecordset("Joblist")
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
    Set rsTable = dbDbase.OpenRecordset("joblist")
    
    ' Add Job to the Joblist
    rsTable.AddNew
    rsTable("Job Number") = JobNumber
    rsTable("Description") = JobDesc
    rsTable("Vehicle") = JobInfo(station, Shift).Vehicle + " "
    rsTable("Start Time") = StartTime
    rsTable("Station") = Format(station, "0")
    rsTable("Shift") = Format(Shift, "0")
    rsTable("Report Filename") = ReportFilename
    rsTable.Update
    rsTable.Close
    dbDbase.Close
    
    ' Refresh the joblist form
    If frmJoblist.Visible Then frmJoblist.RefreshJoblist

End Sub

Sub Trim_Joblist()
'
' Routine Name: Trim Joblist
' Author:       DJP / APS    10/96
' Description:
' This routine uses the value JOBRECS which represents the maximum desired
' number of records in the Joblist.  (Comes from System Configuration)
' If SysConfig.JobRecs is = 0, then the Joblist is not altered.
' If SysConfig.JobRecs is some other number, this routine removes records from the
' Joblist until the number of records is <= to the maximum.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 4

    ' Don't perform any action if Jobrecs = 0
    If SysConfig.JobRecs = 0 Then
      ResetErrModule
      Exit Sub
    End If
    
    Dim dB As Database
    Dim rS As Recordset
    Dim rsCrit As String
    
    rsCrit = "SELECT * FROM [Joblist] ORDER BY [Job Number] DESC"
    ' Move to last record and delete records until less than JOBRECS
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    If Not rS.BOF Then
      rS.MoveLast
      Do While rS.RecordCount > SysConfig.JobRecs        ' Delete from bottom (oldest)
        rS.Delete
        rS.MoveLast
      Loop
    End If
    
    rS.Close
    dB.Close
    
    ' update the joblist form
    If frmJoblist.Visible Then frmJoblist.RefreshJoblist

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

Sub Update_Joblist(iStn As Integer, iShift As Integer)
' Function Name:    Update_Joblist
' Created By:       DJP / APS
' This function updates the joblist entry for the job which has just
' completed with the Stop Time and Vehicle Number.
'

Dim rsCrit As String
Dim dbDbase As Database
Dim rsTable As Recordset

    sPath = FILEPATH_sysdbf & DATAMASTER
    Set daodb36 = DBEngine.OpenDatabase(sPath)
    Set rS = daodb36.OpenRecordset("joblist")
    
    rsCrit = "SELECT * FROM [Joblist] WHERE [Job Number] = '" & StationControl(iStn, iShift).Job_Number & "'"
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
    Set rsTable = dbDbase.OpenRecordset(rsCrit, dbOpenDynaset)
    
    If rsTable.BOF Then
        Write_ELog "Job #" & StationControl(iStn, iShift).Job_Number & " not in joblist. Can't Update"
    Else
      rsTable.MoveFirst
      rsTable.Edit
      rsTable("Stop Time") = StationControl(iStn, iShift).End_Time
      rsTable("Vehicle") = JobInfo(iStn, iShift).Vehicle + " "
      rsTable.Update
    End If
        
    rsTable.Close
    dbDbase.Close
    
    ' update the joblist form
    If frmJoblist.Visible Then frmJoblist.RefreshJoblist
    
End Sub

Sub View_Alarm(ByVal JobNum As String, ByVal iStn As Integer, ByVal iShift As Integer)
'
' Function Name:    View Alarm
' Created By:       Analytical Process Programmer
' Creation Date:    8/29/96
' Description:      Routine calls the data logger and displays the alarms
'                   logged for that station.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 5
Dim JobFilePath As String

    JobFilePath = FILEPATH_data & "C" & JobNum & AccessDbFileExt

    If JobFilePath <> "" Then
    
        frmDataLog.adoLogData.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Data Source=" & JobFilePath & ";" _
            & "Persist Security Info=False"
        frmDataLog.adoLogData.RecordSource = "SELECT * FROM [Alarm] ORDER BY [Alarm].[Time] DESC"
    
        With frmDataLog
            .LogData = "Alarm"
            .LogJob = JobNum
            .LogStn = iStn
            .LogShift = iShift
            .adoLogData.Refresh
            .Caption = "Alarm Log"
            .lblDataLog.Caption = "Alarm Log    Station " & iStn
            If NR_SHIFT > 1 Then .lblDataLog.Caption = .lblDataLog.Caption & "  Shift " & iShift
            .txtMsg.Left = IIf((.adoLogData.Recordset.RecordCount = 0), (.dbgDataLog.Left + ((.dbgDataLog.Width / 2) - (.txtMsg.Width / 2))), OutOfSight)
            .txtMsg.Top = IIf((.adoLogData.Recordset.RecordCount = 0), (.dbgDataLog.Top + 1850), OutOfSight)
            .txtMsg.text = IIf((.adoLogData.Recordset.RecordCount = 0), "No Alarms to report", "No data to report")
            .cmdClear.Visible = IIf((.adoLogData.Recordset.RecordCount = 0), False, True)
            .Icon = .cmdAlarm.Picture
            .Show
        End With
    Else
        Delay_Box "Station ALARM DB Table not opened yet.", MSGDELAY, msgSHOW
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

Sub View_OOT(ByVal JobNum As String, ByVal iStn As Integer, ByVal iShift As Integer)
'
' Function Name:    View_OOT
' Created By:       Analytical Process Programmer
' Creation Date:    8/29/96
' Description:      Routine calls the data logger and displays the out of
'                   tolerance data logged for that station.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 7
Dim JobFilePath As String

    JobFilePath = FILEPATH_data & "C" & JobNum & AccessDbFileExt

    If JobFilePath <> "" Then
    
        frmDataLog.adoLogData.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Data Source=" & JobFilePath & ";" _
            & "Persist Security Info=False"
        frmDataLog.adoLogData.RecordSource = "SELECT * FROM [Tolerance] ORDER BY [Tolerance].[Time] DESC"
    
        With frmDataLog
            .LogData = "OOT"
            .LogJob = JobNum
            .LogStn = iStn
            .LogShift = iShift
            .adoLogData.Refresh
            .Caption = "Out of Tolerance Log"
            .lblDataLog.Caption = "Out of Tolerance Log    Station " & iStn
            If NR_SHIFT > 1 Then .lblDataLog.Caption = .lblDataLog.Caption & "  Shift " & iShift
            .txtMsg.Left = IIf((.adoLogData.Recordset.RecordCount = 0), (.dbgDataLog.Left + ((.dbgDataLog.Width / 2) - (.txtMsg.Width / 2))), OutOfSight)
            .txtMsg.Top = IIf((.adoLogData.Recordset.RecordCount = 0), (.dbgDataLog.Top + 1850), OutOfSight)
            .txtMsg.text = IIf((.adoLogData.Recordset.RecordCount = 0), "No Out of Tolerances to report", "No data to report")
            .cmdClear.Visible = IIf((.adoLogData.Recordset.RecordCount = 0), False, True)
            .Icon = .cmdOOT.Picture
            .Show
        End With
    Else
        Delay_Box "Station TOLERANCE DB Table not opened yet.", MSGDELAY, msgSHOW
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

Sub Write_ELog(ByVal sMsg As String)
'
' Routine   Write_ELog
' Author    DJP APS   10/96
' Description:
' Writes a message to the Event Log, stamps the message with the current
' time and data.
' The message must be passed as a string in sMsg.
' If sMsg is blank, the message '  - None -  ' is written to the log
' If the current number of Event Log entries is greater than the
' value specified in the system configuration file for Max Event Log Entries
' then the oldest Event Log entries are deleted.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
' do not use SetErrModule for this routine, it cant user errhandler
' because the first line of errhandler writes to Event Log.
' creates a big forever type loop
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim nrRecs As Integer

    If ElogInitDone Then
       
       '   sMsg = "tst>" & sMsg & "<tst"       ' for debugging
        If sMsg = "" Then sMsg = " - None - "
        ' Write a record to the EventLog
        rsCrit = "SELECT * FROM [EventLog] ORDER BY [EventLog].[Time] DESC"
        Set dB = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
        Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
        rS.AddNew
        rS("Time") = Now
        rS("Comment") = Mid(sMsg, 1, 255)
        rS.Update
        ' Check to see if too many entries
        rS.MoveLast
        nrRecs = rS.RecordCount
        rS.Close
        dB.Close
        
        ' Trim the Event Log to max entries
        If (SysConfig.EventRecs = 0) Then Exit Sub
        If nrRecs > SysConfig.EventRecs Then
          rsCrit = "SELECT * FROM [EventLog] ORDER BY [EventLog].[Time] DESC"
          Set dB = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
          Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
          If Not rS.BOF Then
            rS.MoveLast
            Do While rS.RecordCount > SysConfig.EventRecs
              rS.Delete
              rS.MoveLast
            Loop
          End If
          rS.Close
          dB.Close
        End If
    
    Else
    
        ' not ready for elog; display the message
        Delay_Box sMsg, MSGDELAY, msgSHOW
        
    End If
    
Exit Sub

localhandler:
Dim iresponse As Integer
Delay_Box "Error " & err & "Writing to Event Log: " & error$(err), MSGDELAY, msgSHOW
End Sub

Sub Write_AirLog(AirLogFile As String, Optional ByVal msg As String)
'
' Writes a record to the current Air Temp/Rh Log file.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
' does not use SetErrModule for this routine
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim txt As String

    ' Ensure valid values
    '   comment
    If msg = "" Then
        ' expand Null message
        txt = " - None - "
    ElseIf Len(msg) > 255 Then
        ' trim message that is too long
        txt = Mid(msg, 1, 255)
    Else
        ' use msg as is
        txt = msg
    End If
    
    ' Write a record to the Air_Log Table
    rsCrit = "SELECT * FROM [Air_Log] ORDER BY [Air_Log].[DTS] DESC"
    Set dB = OpenDatabase(FILEPATH_log & CurAirLogFile)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    rS.AddNew
        rS("DTS") = Now
        rS("Barometer") = AmbBaro
        rS("Temperature") = AmbTemp
        rS("Humidity") = AmbHum
        rS("Moisture") = AmbMoisture
        rS("TemperatureOOT") = IIf(PAS_INFO(pasTEMPERATURE).Ok, False, True)
        rS("MoistureOOT") = IIf(PAS_INFO(pasMOISTURE).Ok, False, True)
        rS("Comment") = txt
    rS.Update
    
    ' close database
    rS.Close
    dB.Close
    
    If (Not PAS_INFO(pasMOISTURE).Ok) Then
        Write_ELog ("Wrote AirLog Record with Moisture OOT; Moisture = " & Format(PAMoisture, "##0.0"))
    ElseIf (Not PAS_INFO(pasTEMPERATURE).Ok) Then
        Write_ELog ("Wrote AirLog Record with Temp OOT; Temp = " & Format(PATemp, "##0.0"))
    ElseIf SysConfig.TempRhLogVerbose Then
        Write_ELog ("Wrote AirLog Record")
    End If
    
Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Writing to AirLog(" & AirLogFile & "): " & error$(err)
    etxt = Mid(etxt, 1, 255)
    Write_ELog etxt
    If Not ErrorMsgBypassActive Then Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Sub Write_FuelUseLog(ByVal iDate As Date, ByVal sButane As Single, ByVal sFuelVapor As Single)
'
' Writes a record to the FuelUse Log file.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
' does not use SetErrModule for this routine
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim iYear As Integer
Dim iMonth As Integer
Dim iDay As Integer
Dim clipButane As Single
Dim clipFuelVapor As Single
Dim tmpTotal As Single

    ' three digits to the right of the decimal point please
    clipButane = CSng(Format(sButane, "###0.000"))
    clipFuelVapor = CSng(Format(sFuelVapor, "###0.000"))
    
    ' Write a record to the FuelUseLog Table
    iYear = CInt(Year(iDate))
    iMonth = CInt(Month(iDate))
    iDay = CInt(Day(iDate))
    rsCrit = "SELECT * FROM [FuelUseLog] WHERE [Year] = " & iYear & "  and [Month] = " & iMonth & " and [DayOfMonth] = " & iDay & " "
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAMASTER)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    
    If rS.BOF Then
        rS.AddNew
            rS("Year") = iYear
            rS("Month") = iMonth
            rS("DayOfMonth") = iDay
            rS("ButaneTotal") = clipButane
            rS("FuelVaporTotal") = clipFuelVapor
        rS.Update
    Else
        rS.MoveFirst
        rS.Edit
                tmpTotal = rS("ButaneTotal")
            rS("ButaneTotal") = tmpTotal + clipButane
                tmpTotal = rS("FuelVaporTotal")
            rS("FuelVaporTotal") = tmpTotal + clipFuelVapor
        rS.Update
    End If
        
    
    ' close database
    rS.Close
    dB.Close
    
    If (Not NotDebugMMW) Then Write_ELog ("Wrote FuelUseLog Record for " & Format(iDate, "YYYY MMMM") & " of " & Format(sButane, "###0.00") & " / " & Format(sFuelVapor, "###0.00") & " grams of Butane/FuelVapor")
    
Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Writing to FuelUseLog " & error$(err)
    etxt = Mid(etxt, 1, 255)
    Write_ELog etxt
    If Not ErrorMsgBypassActive Then Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Sub Write_JLog(Index As Integer, index2 As Integer, Comment As String)
'
' Function Name:    Write_JLog`
' Author:           Brunrose
' Description:      Updates the data file with job event info
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 17493

Dim dbDbase As Database
Dim rsTable As Recordset

    If (Len(StationControl(Index, index2).DBFile) > 3) Then
        Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
        Set rsTable = dbDbase.OpenRecordset("EventLog")
        rsTable.AddNew
        rsTable("Course") = StationControl(Index, index2).Course
        rsTable("Time") = Now()
        rsTable("Comment") = Comment
        rsTable.Update
        rsTable.Close
        dbDbase.Close
    Else
        Write_ELog "Attempted write to non-existant JobLog for Station #" & Format(Index, "#0") & " Shift #" & Format(index2, "#0") & " - " & Left(Comment, 255)
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

Sub Write_Zlog_Purge(ByVal prg As Integer, ByVal fnc As Integer, ByVal adr As Integer, ByVal chn As Integer, ByVal out As Integer, ByVal msg As String)
'  Desc As String
'  CheckSecs as Integer
'  RequestRdy As Boolean                     ' PurgeAir Supply Ready Request Flag (from station(s)); station will want purgeair soon
'  RequestRun As Boolean                     ' PurgeAir Supply Run Request Flag (from station(s)); station wants purgeair now
'  LastRequestRdy As Boolean                 ' Value of "RequestRdy Flag" Last Time it was checked
'  LastRequestRun As Boolean                 ' Value of "RequestRun Flag" Last Time it was checked
'  StandbyRequest As Boolean                 ' PurgeAir Standby Request Flag (from station(s)); station is testing
'  LastStandbyRequest As Boolean             ' Value of "Standby Request Flag" Last Time it was checked
'  lastTime As Date                          ' DateTime when "Request Flag" was last checked
'  StandingBy As Boolean                     ' PurgeAir Supply is StandingBy (i.e. Ready to be Requested to be Ready)
'  Requested As Boolean                      ' PurgeAir Supply is Requested to be Ready
'  Running As Boolean                        ' PurgeAir Supply is Running
'  Ready As Boolean                          ' PurgeAir Supply Ready Flag (to station(s)); PurgeAir Supply is Ready to Run
'  UsingPrgReqHdw As Boolean                 ' set by PurgDef; this PurgeAir Source uses Purge Request/Ready (DO/DI) Hardware
'  UsingVacSwHdw As Boolean                  ' set by PurgDef; this PurgeAir Source uses Vacuum Switch(s)
'  UsingAuxAirSol As Boolean                 ' set by PurgDef; this PurgeAir Source uses an Aux Air Valve
'  UsingPosPrsPrg As Boolean                 ' set by PurgDef; this PurgeAir Source can perform Positive Pressure Purges
'
' Writes a message to the Zlog Purge Table, stamps the message with the current time and date.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
' does not use SetErrModule for this routine
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim txt As String

' Ensure valid values
If Not IsNumeric(prg) Then prg = 0
If Not IsNumeric(fnc) Then fnc = 0
If Not IsNumeric(adr) Then adr = 0
If Not IsNumeric(chn) Then chn = 0
If Not IsNumeric(out) Then out = 0
If msg = "" Then
    ' expand Null message
    txt = " - None - "
ElseIf Len(msg) > 255 Then
    ' trim message that is too long
    txt = Mid(msg, 1, 255)
Else
    ' use msg as is
    txt = msg
End If


' Clear Log first?
If Debug_ZlogPurge_Clear Then
    rsCrit = "SELECT * FROM [PurgeLog] ORDER BY [PurgeLog].[DTS] DESC"
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    If Not rS.BOF Then
        rS.MoveLast
        Do While Not rS.BOF
            rS.Delete
            rS.MoveLast
        Loop
    End If
    rS.Close
    dB.Close
    Debug_ZlogPurge_Clear = False
End If

' Write a record to the Zlog PurgeLog Table
rsCrit = "SELECT * FROM [PurgeLog] ORDER BY [PurgeLog].[DTS] DESC"
Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
rS.AddNew
    rS("DTS") = Now
    rS("Timer") = Timer
    rS("Number") = prg
    rS("Function") = fnc
    rS("Address") = adr
    rS("Channel") = chn
    rS("Output") = out
    If prg >= 1 And prg <= MAX_PRG Then
        rS("Running") = PRG_INFO(prg).Running
        rS("Ready") = PRG_INFO(prg).Ready
        rS("Requested") = PRG_INFO(prg).Requested
        rS("LastTime") = PRG_INFO(prg).lastTime
        rS("LastStandByRequest") = PRG_INFO(prg).LastStandbyRequest
        rS("StandByRequest") = PRG_INFO(prg).StandbyRequest
        rS("LastRequestRdy") = PRG_INFO(prg).LastRequestRdy
        rS("RequestRdy") = PRG_INFO(prg).RequestRdy
        rS("LastRequestRun") = PRG_INFO(prg).LastRequestRun
        rS("RequestRun") = PRG_INFO(prg).RequestRun
    End If
    rS("Temperature") = PATemp
    rS("Humidity") = PAHum
    rS("Barometer") = AmbBaro
    rS("Moisture") = PAMoisture
    rS("MoistureOK") = PAS_INFO(pasMOISTURE).Ok
    rS("TemperatureOK") = PAS_INFO(pasTEMPERATURE).Ok
    If NR_STN > 0 Then rS("Stn1Active_Mode") = StationControl(1, Stn_ActiveShift(1)).Mode
    If NR_STN > 1 Then rS("Stn2Active_Mode") = StationControl(2, Stn_ActiveShift(2)).Mode
    If NR_STN > 2 Then rS("Stn3Active_Mode") = StationControl(3, Stn_ActiveShift(3)).Mode
    If NR_STN > 3 Then rS("Stn4Active_Mode") = StationControl(4, Stn_ActiveShift(4)).Mode
    If NR_STN > 4 Then rS("Stn5Active_Mode") = StationControl(5, Stn_ActiveShift(5)).Mode
    If NR_STN > 5 Then rS("Stn6Active_Mode") = StationControl(6, Stn_ActiveShift(6)).Mode
    If NR_STN > 6 Then rS("Stn7Active_Mode") = StationControl(7, Stn_ActiveShift(7)).Mode
    If NR_STN > 7 Then rS("Stn8Active_Mode") = StationControl(8, Stn_ActiveShift(8)).Mode
    If NR_STN > 8 Then rS("Stn9Active_Mode") = StationControl(9, Stn_ActiveShift(9)).Mode
    rS("Comment") = txt
rS.Update


' Check to see how many entries in the table
rS.MoveLast
    Debug_ZlogPurge_NumRecords = rS.RecordCount
rS.Close
dB.Close

' Trim the zLog to max entries
If Debug_ZlogPurge_NumRecords > Debug_ZlogPurge_MaxRecords Then
    Dim newmax As Long
    newmax = CLng(0.9 * Debug_ZlogPurge_MaxRecords)
    rsCrit = "SELECT * FROM [PurgeLog] ORDER BY [PurgeLog].[DTS] DESC"
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    If Not rS.BOF Then
        rS.MoveLast
        Do While rS.RecordCount > newmax
            rS.Delete
            rS.MoveLast
        Loop
    End If
    rS.Close
    dB.Close
End If

Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Writing to Zlog Purge: " & error$(err)
    etxt = Mid(etxt, 1, 255)
    Write_ELog etxt
    If Not ErrorMsgBypassActive Then Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Sub Write_Zlog_Scales(ByVal num As Integer, ByVal prt As Integer, ByVal newrd As String, ByVal msg As String)
' New Reading As String
'
' Writes a message to the Zlog Scales Table, stamps the message with the current time and date.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
' does not use SetErrModule for this routine
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim txt, rd As String

    ' Ensure valid values
    '   scale #
    If Not IsNumeric(num) Then num = 0
    '   port #
    If Not IsNumeric(prt) Then prt = 0
    '   new reading
    If newrd = "" Then
        ' expand Null reading
        rd = " - None - "
    ElseIf Len(newrd) > 255 Then
        ' trim reading that is too long
        rd = Mid(newrd, 1, 255)
    Else
        ' use reading as is
        rd = newrd
    End If
    '   comment
    If msg = "" Then
        ' expand Null message
        txt = " - None - "
    ElseIf Len(msg) > 255 Then
        ' trim message that is too long
        txt = Mid(msg, 1, 255)
    Else
        ' use msg as is
        txt = msg
    End If
    
    
    ' Clear Log first?
    If Debug_ZlogScale_Clear Then
        rsCrit = "SELECT * FROM [ScaleLog] ORDER BY [ScaleLog].[DTS] DESC"
        Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
        Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
        If Not rS.BOF Then
            rS.MoveLast
            Do While Not rS.BOF
                rS.Delete
                rS.MoveLast
            Loop
        End If
        rS.Close
        dB.Close
        Debug_ZlogScale_Clear = False
    End If
    
    ' Write a record to the Zlog ScaleLog Table
    rsCrit = "SELECT * FROM [ScaleLog] ORDER BY [ScaleLog].[DTS] DESC"
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    rS.AddNew
        rS("DTS") = Now
        rS("Timer") = Timer
        rS("Number") = num
        rS("Port") = prt
        rS("Reading") = rd
        rS("Comment") = txt
    rS.Update
    
    
    ' Check to see how many entries in the table
    rS.MoveLast
        Debug_ZlogScale_NumRecords = rS.RecordCount
    rS.Close
    dB.Close
    
    ' Trim the zLog to max entries
    If Debug_ZlogScale_NumRecords > Debug_ZlogScale_MaxRecords Then
        Dim newmax As Long
        newmax = CLng(0.75 * Debug_ZlogScale_MaxRecords)
        rsCrit = "SELECT * FROM [ScaleLog] ORDER BY [ScaleLog].[DTS] DESC"
        Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
        Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
        If Not rS.BOF Then
            rS.MoveLast
            Do While rS.RecordCount > newmax
                rS.Delete
                rS.MoveLast
            Loop
        End If
        rS.Close
        dB.Close
    End If
    
Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Writing to Zlog Scales: " & error$(err)
    etxt = Mid(etxt, 1, 255)
    Write_ELog etxt
    If Not ErrorMsgBypassActive Then Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Sub Write_Zlog_PAS(ByVal msg As String)
'
' Writes a message to the Zlog PAS Table, stamps the message with the current time and date.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
' does not use SetErrModule for this routine
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim txt, rd As String
Dim addr, chan As Integer

    ' Ensure valid values
    '   comment
    If msg = "" Then
        ' expand Null message
        txt = " - None - "
    ElseIf Len(msg) > 255 Then
        ' trim message that is too long
        txt = Mid(msg, 1, 255)
    Else
        ' use msg as is
        txt = msg
    End If
    
    
    ' Clear Log first?
    If Debug_ZlogPAS_Clear Then
        rsCrit = "SELECT * FROM [PAS_Log] ORDER BY [PAS_Log].[DTS] DESC"
        Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
        Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
        If Not rS.BOF Then
            rS.MoveLast
            Do While Not rS.BOF
                rS.Delete
                rS.MoveLast
            Loop
        End If
        rS.Close
        dB.Close
        Debug_ZlogPAS_Clear = False
    End If
    
    ' Write a record to the Zlog PAS_Log Table
    rsCrit = "SELECT * FROM [PAS_Log] ORDER BY [PAS_Log].[DTS] DESC"
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    rS.AddNew
        rS("DTS") = Now
        rS("Timer") = Timer
        rS("PAS_LastUpdateTemperature") = PAS_INFO(pasTEMPERATURE).LastUpdate
        rS("PAS_LastUpdateMoisture") = PAS_INFO(pasMOISTURE).LastUpdate
        
        rS("TemperaturePV") = PID_INFO(pasTEMPERATURE).PV
        rS("TemperatureSP") = PID_INFO(pasTEMPERATURE).SP
        rS("TemperatureEr") = PID_INFO(pasTEMPERATURE).Er
        rS("TemperaturePgain") = PID_INFO(pasTEMPERATURE).Pgain
        rS("TemperatureIgain") = PID_INFO(pasTEMPERATURE).Igain
        rS("TemperaturePgain") = PID_INFO(pasTEMPERATURE).Dgain
        rS("TemperatureCumI") = PID_INFO(pasTEMPERATURE).CumI
        rS("TemperatureOut") = PID_INFO(pasTEMPERATURE).out
        rS("TemperatureEnable") = PID_INFO(pasTEMPERATURE).Enable
        rS("TemperatureInhibit") = PID_INFO(pasTEMPERATURE).Inhibit
        rS("TemperatureRev") = PID_INFO(pasTEMPERATURE).Rev
        rS("TemperatureOutput") = PID_INFO(pasTEMPERATURE).Output
        rS("TemperatureOffTimer") = PID_INFO(pasTEMPERATURE).OffTimer
        rS("TemperatureOffDuty") = PID_INFO(pasTEMPERATURE).OffDuty
        rS("TemperatureOffLimitDelta") = PID_INFO(pasTEMPERATURE).OffLimitDelta
        rS("TemperatureOnTimer") = PID_INFO(pasTEMPERATURE).OnTimer
        rS("TemperatureOnDuty") = PID_INFO(pasTEMPERATURE).OnDuty
        rS("TemperatureOnLimitDelta") = PID_INFO(pasTEMPERATURE).OnLimitDelta
        rS("TemperatureLastUpdate") = PID_INFO(pasTEMPERATURE).LastUpdate
        
        rS("MoisturePV") = PID_INFO(pasMOISTURE).PV
        rS("MoistureSP") = PID_INFO(pasMOISTURE).SP
        rS("MoistureEr") = PID_INFO(pasMOISTURE).Er
        rS("MoisturePgain") = PID_INFO(pasMOISTURE).Pgain
        rS("MoistureIgain") = PID_INFO(pasMOISTURE).Igain
        rS("MoistureDgain") = PID_INFO(pasMOISTURE).Dgain
        rS("MoistureCumI") = PID_INFO(pasMOISTURE).CumI
        rS("MoistureOut") = PID_INFO(pasMOISTURE).out
        rS("MoistureEnable") = PID_INFO(pasMOISTURE).Enable
        rS("MoistureInhibit") = PID_INFO(pasMOISTURE).Inhibit
        rS("MoistureRev") = PID_INFO(pasMOISTURE).Rev
        rS("MoistureLastUpdate") = PID_INFO(pasMOISTURE).LastUpdate
        
        rS("TemperatureDuration") = PAS_INFO(pasTEMPERATURE).Duration
        rS("TemperatureTarget") = PAS_INFO(pasTEMPERATURE).DurationTarget
        rS("TemperatureOK") = PAS_INFO(pasTEMPERATURE).Ok
        rS("TemperatureTimeoutDuration") = PAS_INFO(pasTEMPERATURE).TimeOutDuration
        rS("TemperatureTimeoutTarget") = PAS_INFO(pasTEMPERATURE).TimeOutTarget
        rS("TemperatureTimeout") = PAS_INFO(pasTEMPERATURE).timeOut
        
        rS("MoistureDuration") = PAS_INFO(pasMOISTURE).Duration
        rS("MoistureTarget") = PAS_INFO(pasMOISTURE).DurationTarget
        rS("MoistureOK") = PAS_INFO(pasMOISTURE).Ok
        rS("MoistureTimeoutDuration") = PAS_INFO(pasMOISTURE).TimeOutDuration
        rS("MoistureTimeoutTarget") = PAS_INFO(pasMOISTURE).TimeOutTarget
        rS("MoistureTimeout") = PAS_INFO(pasMOISTURE).timeOut
        
        rS("HdwReqOut") = Com_DIO(icPurgeRequestOut).Value
        rS("HdwRdyIn") = Com_DIO(icPurgeReadyIn).Value
        rS("PasReqIn") = Com_DIO(icPASPowerOnIn).Value
        rS("PasRdyOut") = Com_DIO(icPASReadyOut).Value
        rS("RunLocalPasIn") = Com_DIO(icPASisRunningIn).Value
        rS("TemperatureControlOut") = Com_DIO(icPASHeaterSSR).Value
            addr = Com_AIO(acPASMoistCntrlOut).addr
            chan = Com_AIO(acPASMoistCntrlOut).chan
        rS("MoistureControlOut") = Map_AIO(addr, chan).EUValue
        
        rS("Comment") = txt
    rS.Update
    Debug_ZlogPAS_LastUpdate = PAS_INFO(pasTEMPERATURE).LastUpdate
    If PAS_INFO(pasMOISTURE).LastUpdate > Debug_ZlogPAS_LastUpdate Then Debug_ZlogPAS_LastUpdate = PAS_INFO(pasMOISTURE).LastUpdate
    
    ' Check to see how many entries in the table
    rS.MoveLast
        Debug_ZlogPAS_NumRecords = rS.RecordCount
    rS.Close
    dB.Close
    
    ' Trim the zLog to max entries
    If Debug_ZlogPAS_NumRecords > Debug_ZlogPAS_MaxRecords Then
        Dim newmax As Long
        newmax = CLng(0.75 * Debug_ZlogPAS_MaxRecords)
        rsCrit = "SELECT * FROM [PAS_Log] ORDER BY [PAS_Log].[DTS] DESC"
        Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
        Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
        If Not rS.BOF Then
            rS.MoveLast
            Do While rS.RecordCount > newmax
                rS.Delete
                rS.MoveLast
            Loop
        End If
        rS.Close
        dB.Close
    End If

Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Writing to Zlog PAS: " & error$(err)
    etxt = Mid(etxt, 1, 255)
    Write_ELog etxt
    If Not ErrorMsgBypassActive Then Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Function AirLogFileIsReady(ByVal curDate As Date) As Boolean
Dim filename As String
Dim filepath As String
Dim strExists As String
Dim flag As Boolean
    
    ' Test to see if DB File already there.
    filename = "AirLog_"
    filename = filename + Format(Year(curDate), "0000") + Format(Month(curDate), "00")
    filename = filename + "_" + AccessDbFileExt
    filepath = FILEPATH_log & filename
    strExists = Dir(filepath)
    If strExists = filename Then
        ' file already exists
        flag = True
        CurAirLogFile = filename
    Else
        ' create database file
        FileCopy FILEPATH_sysdbf & DATAAIRLOG, filepath
        ' verify copy
        strExists = Dir(filepath)
        If strExists = filename Then
            ' copy was successful
            Write_ELog ("AirLog db file Create Successful for " & filename)
            flag = True
            CurAirLogFile = filename
        Else
            ' copy failed
            Write_ELog ("AirLog db file Create Failed for " & filename)
            flag = False
            CurAirLogFile = "none"
        End If
    End If
    
    AirLogFileIsReady = flag
End Function

Sub View_JobLog(ByVal JobNum As String, ByVal iStn As Integer, ByVal iShift As Integer)
'
' Function Name:    View JobLog
' Created By:       Brunrose
' Description:      Routine calls the data logger and displays the Job Events Log
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 4675
Dim deltaWidth, deltaWidth2 As Single
Dim JobFilePath As String

    JobFilePath = FILEPATH_data & "C" & JobNum & AccessDbFileExt

    If JobFilePath <> "" Then
        
        frmDataLog.adoLogData.ConnectionString = _
                "Provider=Microsoft.Jet.OLEDB.4.0;" _
                & "Data Source=" & JobFilePath & ";" _
                & "Persist Security Info=False"
        frmDataLog.adoLogData.RecordSource = "SELECT * FROM [EventLog] ORDER BY [EventLog].[EventNumber] DESC"
    
        With frmDataLog
            .LogData = "JobLog"
            .LogJob = JobNum
            .LogStn = iStn
            .LogShift = iShift
            .adoLogData.Refresh
            .Caption = "JobEvents Log"
            .lblDataLog.Caption = "JobEvents Log    Station " & iStn
            If NR_SHIFT > 1 Then .lblDataLog.Caption = .lblDataLog.Caption & "  Shift " & iShift
            .txtMsg.Left = IIf((.adoLogData.Recordset.RecordCount = 0), (.dbgDataLog.Left + ((.dbgDataLog.Width / 2) - (.txtMsg.Width / 2))), OutOfSight)
            .txtMsg.Top = IIf((.adoLogData.Recordset.RecordCount = 0), (.dbgDataLog.Top + 1850), OutOfSight)
            .txtMsg.text = IIf((.adoLogData.Recordset.RecordCount = 0), "No Events to report", " ")
            .cmdClear.Visible = IIf((.adoLogData.Recordset.RecordCount = 0), False, True)
            deltaWidth = 0.36 * .Width
            deltaWidth2 = 0.03 * .Width
'            .Width = .Width + deltaWidth2
'            .dbgDataLog.Width = .dbgDataLog.Width + (1.01 * deltaWidth2)
            .dbgDataLog.Columns(0).Visible = True
'            .dbgDataLog.Columns(0).Width = 1900
'            .dbgDataLog.Columns(1).Width = 7200
            .Icon = .cmdJobLog.Picture
            .Show
        End With
        
    Else
        Delay_Box "Job Event Log not opened yet.", MSGDELAY, msgSHOW
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

Sub Write_AkLog(ByVal dts As Date, ByVal tmr As Double, ByVal cmd As String, ByVal rsp As String)
'
' Routine   Write_AkLog
' Author    Brunrose
' Description:
' Writes a data to the AK Command log, stamps the message with the current time and data.
'
' If the current number of log entries is greater than the
' value specified in the system configuration file for Max zLog Entries
' then the oldest error log entries are deleted.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 6, 626

Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim nrRecs, trimRecs As Long
Dim MaxAkLogEntries As Long

'Exit Sub

    ' Write a record to the AK Command Log
    rsCrit = "SELECT * FROM [AK_Log] ORDER BY [AK_Log].[DTS] DESC"
    Set dB = OpenDatabase(FILEPATH_sysdbf & DATAZLOG)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    rS.AddNew
    rS("DTS") = dts
    rS("Timer") = tmr
    rS("Command") = cmd
    rS("Response") = rsp
    rS.Update
    'rS.Refresh
    ' Check to see if too many entries
    rS.MoveLast
    nrRecs = rS.RecordCount
    rS.Close
    dB.Close
    
    MaxAkLogEntries = 100000
'    MaxAkLogEntries = 200
    If nrRecs > MaxAkLogEntries Then
        ' Trim the log to max entries
        trimRecs = (0.6 * MaxAkLogEntries)
        rsCrit = "SELECT * FROM [AK_Log] ORDER BY [AK_Log].[DTS] DESC"
        Set dB = OpenDatabase(FILEPATH_sysdbf & "SCSIIzLog.mdb")
        Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
        If Not rS.BOF Then
            rS.MoveLast
            Do While rS.RecordCount > trimRecs
                rS.Delete
                rS.MoveLast
            Loop
        End If
        rS.Close
        dB.Close
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



