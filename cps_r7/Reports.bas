Attribute VB_Name = "Module3"
'error module 3 ''''''''''''''''''''''REPORTS.bas '''''''''''''''''''''
Option Explicit
'
Private CurrRate(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Private deltaVol(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Private deltaWt(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Private lastSecond As Integer
Private lastAirLogDTS As Date
Private tempstr As String
Private Const No_Rate = 0
Private Const Def_Rate = 1
Private Const Leak_Rate = 2
Private Const Load_Rate = 3
Private Const Purge_Rate = 4
Private Const LT_Rate = 5


Sub Data_Writer()
' Procedure Name:   Data_Writer
' Created by:       Brunrose    Feb. 2007
' Description:      This routine initiates writes of data to the database file for the station.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 222
Dim station, Shift As Integer
Dim livefuelFlag As Boolean
Dim orvr2Flag As Boolean
Dim btnMult As Single
Dim sdate As Date
Static xcntr As Integer


    'Time to log data to db file ??
    For station = 1 To LAST_STN
        For Shift = 1 To NR_SHIFT
        
            Select Case STN_INFO(station).Type
                Case STN_LIVEFUEL_TYPE
                    livefuelFlag = True
                Case STN_LIVEREG_TYPE, STN_LIVEORVR2_TYPE
                    livefuelFlag = IIf(StationRecipe(station, Shift).LiveFuel, True, False)
                Case Else
                    livefuelFlag = False
            End Select
        
            Select Case STN_INFO(station).Type
                Case STN_ORVR2_TYPE, STN_LIVEORVR2_TYPE
                    orvr2Flag = IIf(StationRecipe(station, Shift).UseHiRangeMFC, True, False)
                    btnMult = IIf(StationRecipe(station, Shift).UseHiRangeMFC, STN_INFO(station).ButMfc2DensityMult, STN_INFO(station).ButMfcDensityMult)
                Case Else
                    orvr2Flag = False
                    btnMult = STN_INFO(station).ButMfcDensityMult
            End Select
        
            Select Case StationControl(station, Shift).Mode
            
                Case VBLEAK
                    ' Just Starting Leakcheck Data Logging ???
                    If (CurrRate(station, Shift) <> Leak_Rate) Then
                        CurrRate(station, Shift) = Leak_Rate
                        Stn_Leak_Log_TestTimer(station, Shift) = (StationControl(station, Shift).TestTimer - StationConfig(station, Shift).LeakTotal_Interval)
                    End If
                    ' See if time to log Leak Check data
                    If (Stn_Leak_Log_TestTimer(station, Shift) + StationConfig(station, Shift).LeakTotal_Interval) <= StationControl(station, Shift).TestTimer Then
                        ' reset Leak Check logging timer
                        Stn_Leak_Log_TestTimer(station, Shift) = Stn_Leak_Log_TestTimer(station, Shift) + StationConfig(station, Shift).LeakTotal_Interval     ' # of seconds
                        ' Log Leak Check Data to File only if not paused for anything
                        If Not StationControl(station, Shift).IsPausedInAlarm Then
                            Leak_Write CInt(station), CInt(Shift), NORMALUPDATE, NORESULT
'    tempstr = "Station " & Format(station, "#0") & " Leak Write @ " & Format(Timer, "###,##0.000")
'    Debug.Print tempstr
                        End If
                    End If
                
                Case VBLOAD
                    ' First, Update cumulative Weight Values
                    LoadControl(station, Shift).AuxWtChg = StationControl(station, Shift).AuxScaleWt - LoadControl(station, Shift).AuxWt_Start
                    LoadControl(station, Shift).PriWtChg = StationControl(station, Shift).PriScaleWt - LoadControl(station, Shift).PriWt_Start
                    LoadControl(station, Shift).TotalWtChg = LoadControl(station, Shift).AuxWtChg + LoadControl(station, Shift).PriWtChg
                    LoadControl(station, Shift).CurrWtChgRate = RecentWtChgRate(station, Shift)
                    ' Second, Update elapsed time & weight-change rate
                    Select Case LoadControl(station, Shift).Phase
                        Case Is = LoadPrep
'                            LoadControl(station, Shift).ElapsedHours = CSng(0)
                            LoadControl(station, Shift).ElapsedHours_Prev = CSng(0)
'                            LoadControl(station, Shift).LoadRate = CSng(0)
'                            LoadControl(station, Shift).TotalWtChgRate = CSng(0)
                        Case LoadStarting, LoadLoading, LoadComplete
                            LoadControl(station, Shift).ElapsedHours = LoadControl(station, Shift).ElapsedHours_Prev + CSng(DateDiff("s", LoadControl(station, Shift).ElapsedStartDts, Now)) / CSng(3600)
                            If (LoadControl(station, Shift).ElapsedHours <> 0) Then
                                Select Case livefuelFlag
                                    Case True
                                        LoadControl(station, Shift).LoadRate = LoadControl(station, Shift).TotalWtChg / LoadControl(station, Shift).ElapsedHours
                                    Case False
                                        LoadControl(station, Shift).LoadRate = LoadControl(station, Shift).loadTotalGrams / LoadControl(station, Shift).ElapsedHours
'                                        LoadControl(station, Shift).LoadRate = Stn_Btn_Flow_PV(station, Shift) * GramsPerLiter * btnMult * CSng(60)
                                End Select
                                ' Calculate Current Weight Change Rate
                                LoadControl(station, Shift).TotalWtChgRate = LoadControl(station, Shift).TotalWtChg / LoadControl(station, Shift).ElapsedHours
                                ' Calculate Current gm/liter of vapor flow
                                If (STN_INFO(station).Type = STN_LIVEORVR2_TYPE And StationRecipe(station, Shift).UseHiRangeMFC) Then
                                    If (Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue <> 0) Then LoadControl(station, Shift).CurrLoadDensity = LoadControl(station, Shift).CurrWtChgRate * (1 / 60) * (1 / Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue)
                                Else
                                    If (Stn_AIO(station, asLiveFuelVaporFlow).EUValue <> 0) Then LoadControl(station, Shift).CurrLoadDensity = LoadControl(station, Shift).CurrWtChgRate * (1 / 60) * (1 / Stn_AIO(station, asLiveFuelVaporFlow).EUValue)
                                End If
                            Else
                            End If
                        Case Is > LoadComplete
                            If (LoadControl(station, Shift).ElapsedHours <> 0) Then
                                LoadControl(station, Shift).LoadRate = CSng(0)
                                LoadControl(station, Shift).TotalWtChgRate = LoadControl(station, Shift).TotalWtChg / LoadControl(station, Shift).ElapsedHours
                            End If
                    End Select
                    ' Just Starting Load Data Logging ???
                    If (CurrRate(station, Shift) <> Load_Rate) Then
                        CurrRate(station, Shift) = Load_Rate
                        Stn_Load_Log_TestTimer(station, Shift) = (StationControl(station, Shift).TestTimer - StationConfig(station, Shift).LoadTotal_Interval)
                    End If
                    ' See if time to log Load data
                    If (Stn_Load_Log_TestTimer(station, Shift) + StationConfig(station, Shift).LoadTotal_Interval) <= StationControl(station, Shift).TestTimer Then
                        ' reset Load logging timer
                        Stn_Load_Log_TestTimer(station, Shift) = Stn_Load_Log_TestTimer(station, Shift) + StationConfig(station, Shift).LoadTotal_Interval     ' # of seconds
                        ' log Load Data to db only if not paused for anything
                        If Not StationControl(station, Shift).IsPausedInAlarm Then
                            ' Totalize Load Flow
                            Load_Totalize CInt(station), CInt(Shift)
                            ' Write Load Data to File
                            Load_Write CInt(station), CInt(Shift), NORMALUPDATE
'    tempstr = "Station " & Format(station, "#0") & " Load Write @ " & Format(Timer, "###,##0.000")
'    Debug.Print tempstr
                        End If
                    End If
    
            
                Case VBPURGE
                    ' First, Update cumulative Purge Weight Change
                    PurgeControl(station, Shift).AuxWtChg = StationControl(station, Shift).AuxScaleWt - PurgeControl(station, Shift).AuxWt_Start
                    PurgeControl(station, Shift).PriWtChg = StationControl(station, Shift).PriScaleWt - PurgeControl(station, Shift).PriWt_Start
                    PurgeControl(station, Shift).TotalWtChg = PurgeControl(station, Shift).AuxWtChg + PurgeControl(station, Shift).PriWtChg
                    PurgeControl(station, Shift).CurrWtChgRate = RecentWtChgRate(station, Shift)
                    ' Second, Update elapsed time & weight-change rate
                    Select Case PurgeControl(station, Shift).Phase
                        Case Is < PurgePurging
                            PurgeControl(station, Shift).ElapsedHours = CSng(0)
                            PurgeControl(station, Shift).TotalWtChgRate = CSng(0)
                        Case PurgePurging, PurgeComplete
                            PurgeControl(station, Shift).ElapsedHours = PurgeControl(station, Shift).ElapsedHours_Prev + CSng(DateDiff("s", PurgeControl(station, Shift).PhaseStartDts, Now)) / CSng(3600)
                            If (PurgeControl(station, Shift).ElapsedHours <> 0) Then
                                PurgeControl(station, Shift).TotalWtChgRate = PurgeControl(station, Shift).TotalWtChg / PurgeControl(station, Shift).ElapsedHours
                            End If
                        Case Is > PurgeComplete
                            If (PurgeControl(station, Shift).ElapsedHours <> 0) Then
                                PurgeControl(station, Shift).TotalWtChgRate = PurgeControl(station, Shift).TotalWtChg / PurgeControl(station, Shift).ElapsedHours
                            End If
                    End Select
                    ' Just Starting Purge Data Logging ???
                    If (CurrRate(station, Shift) <> Purge_Rate) Then
                        CurrRate(station, Shift) = Purge_Rate
                        Stn_Purge_Log_TestTimer(station, Shift) = (StationControl(station, Shift).TestTimer - StationConfig(station, Shift).PurgeTotal_Interval)
                    End If
                    ' See if time to log Purge data
                    If (Stn_Purge_Log_TestTimer(station, Shift) + StationConfig(station, Shift).PurgeTotal_Interval) <= StationControl(station, Shift).TestTimer Then
                        ' reset Purge logging timer
                        Stn_Purge_Log_TestTimer(station, Shift) = Stn_Purge_Log_TestTimer(station, Shift) + StationConfig(station, Shift).PurgeTotal_Interval     ' # of seconds
                        ' log Load Data to db only if not paused for anything
                        If Not StationControl(station, Shift).IsPausedInAlarm Then
                            ' Totalize Purge Flow
                            Purge_Totalize CInt(station), CInt(Shift)
                            ' Write Purge Data to File
                            Purge_Write CInt(station), CInt(Shift), NORMALUPDATE
'    tempstr = "Station " & Format(station, "#0") & " Purge Write @ " & Format(Timer, "###,##0.000")
'    Debug.Print tempstr
                        End If
                    End If
            
                    
                Case VBPOSTLOAD, VBPRELOAD, VBPOSTPURGE, VBPURGEWAIT, VBPOSTLEAK, VBGASPAUSE, VBWBPAUSE, VBPURGECONT, VBPAUSE, VBPAUSEBYUSER, VBCOURSEWAIT, VBCOURSEPAUSE
                    ' Just Starting Default Data Logging ???
                    If (CurrRate(station, Shift) <> Def_Rate) Then
                        CurrRate(station, Shift) = Def_Rate
                        Stn_Default_Log_TestTimer(station, Shift) = (StationControl(station, Shift).TestTimer - StationConfig(station, Shift).Default_Interval)
                    End If
                    ' See if time to log Other data (log once per second)
                    If (Stn_Default_Log_TestTimer(station, Shift) + StationConfig(station, Shift).Default_Interval) <= StationControl(station, Shift).TestTimer Then
                        ' reset Other logging timer
                        Stn_Default_Log_TestTimer(station, Shift) = Stn_Default_Log_TestTimer(station, Shift) + StationConfig(station, Shift).Default_Interval
                        ' log Other Data to db only if not paused for anything
                        If Not StationControl(station, Shift).IsPausedInAlarm Then
                            ' Write Default Data to File
                            Default_Write CInt(station), CInt(Shift), NORMALUPDATE
'    tempstr = "Station " & Format(station, "#0") & " Default Write @ " & Format(Timer, "###,##0.000")
'    Debug.Print tempstr
                        End If
                    End If
            
                Case VBLEAKTEST
                    ' Just Starting LeakTest Logging ???
                    If (CurrRate(station, Shift) <> LT_Rate) Then
                        CurrRate(station, Shift) = LT_Rate
                        Stn_LT_Log_TestTimer(station, Shift) = (StationControl(station, Shift).TestTimer - Cfg_LeakTest.ReportInterval)
                    End If
                    ' See if time to log LeakTest data
                    If (Stn_LT_Log_TestTimer(station, Shift) + Cfg_LeakTest.ReportInterval) <= StationControl(station, Shift).TestTimer Then
                        ' reset Other logging timer
                        Stn_LT_Log_TestTimer(station, Shift) = Stn_LT_Log_TestTimer(station, Shift) + Cfg_LeakTest.ReportInterval
                        ' log leaktest Data to db only if not paused for anything
                        If Not StationControl(station, Shift).IsPausedInAlarm Then
                            ' Write LeakTest Data to File
                            LT_Write CInt(station), CInt(Shift), NORMALUPDATE
' tempstr = "Station " & Format(station, "#0") & " LT_Write @ " & Format(StationControl(station, Shift).TestTimer, "###,##0.000")
' Debug.Print tempstr
                        End If
                    End If
            
                Case Else
                    ' nothing to do
                
            End Select
        
            ' See if time to log Remote Status data
            If USINGREMSTSMON Then
                If ((Stn_RemStatus_Log_TestTimer(station, Shift) - Timer) > 86000#) Then
                    Stn_RemStatus_Log_TestTimer(station, Shift) = (Stn_RemStatus_Log_TestTimer(station, Shift) - 86400#)
                End If
                If ((Stn_RemStatus_Log_TestTimer(station, Shift) + SysConfig.RemStatus_Interval) <= Timer) Then
Debug.Print "Station " & Format(station, "0") & " Shift " & Format(Shift, "0") & " RemStatusTimer = " & Format(Stn_RemStatus_Log_TestTimer(station, Shift), "####0.000")
                    ' reset RemStatus logging timer
                    Stn_RemStatus_Log_TestTimer(station, Shift) = Stn_RemStatus_Log_TestTimer(station, Shift) + SysConfig.RemStatus_Interval
                    If (Stn_RemStatus_Log_TestTimer(station, Shift) > 86400) Then
                        Stn_RemStatus_Log_TestTimer(station, Shift) = Stn_RemStatus_Log_TestTimer(station, Shift) - 86400
                    End If
                    If ((Stn_RemStatus_Log_TestTimer(station, Shift) + SysConfig.RemStatus_Interval) > 86399.999) Then
                        Stn_RemStatus_Log_TestTimer(station, Shift) = (Stn_RemStatus_Log_TestTimer(station, Shift) - 86400#)
                    End If
                    ' log RemStatus Data to RemoteDB
                    ' & optionally Write RemoteStatus Data to File
                    RemStatus_Update CInt(station), CInt(Shift)
                End If
            End If
        
        
        Next Shift
    Next station
    
    ' Time to Update Saved Butane Supply Values? (done once per hour at 33 min & 33 seconds after the hour)
    If Second(Now) = 33 Then
        If Minute(Now) = 33 Then
            sdate = CDate(GetSetting("cps_r7", "Butane Supply", "LastUpdate DTS", FormatDateTime(Now)))
            If DateDiff("s", sdate, Now) > 3000 Then Save_ButaneSupply
        End If
    End If
    
    ' Time to Log Air Temp/Humidity Values? (switch db files at midnight on 1st of month)
    If LogTempRh Then
        If Second(Now) <> lastSecond Then
            lastSecond = Second(Now)
            If (Second(Now) Mod 6) = 0 Then
                If (Not USINGPASLOCALCONTROL) Then
                    '   Check PAS Temperature
                    AIR_Check pasTEMPERATURE                    ' Tolerance check for PAS Temp for AirLog
                    '   Check PAS Moisture
                    AIR_Check pasMOISTURE                       ' Tolerance check for PAS Moisture for AirLog
                End If
                If DateDiff("s", lastAirLogDTS, Now) >= (CLng(SysConfig.TempRhLogInterval * 60)) Then
                    ' write a record
                    If AirLogFileIsReady(Now) Then
                        ' airlog db file exists; write to it
                        Write_AirLog CurAirLogFile, " "
                        lastAirLogDTS = Now
                    End If
                End If
            End If
        End If
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

Sub Header_Write(Index As Integer, index2 As Integer)
' Procedure Name:   Header_Write
' Description:      This routine writes the file Header data to the data -
'                   base file for the station.
'
Dim dbDbase As Database
Dim rsTable As Recordset
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 4

If StationControl(Index, index2).DBFile <> "" Then
    'Write Header data to data file
    Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
    Set rsTable = dbDbase.OpenRecordset("header")
    rsTable.AddNew
    
        rsTable("Course") = StationControl(Index, index2).Course
        rsTable("Job Number") = StationControl(Index, index2).Job_Number
        rsTable("Job Description") = StationControl(Index, index2).Job_Description
        rsTable("Report Filename") = StationControl(Index, index2).RptFile
        rsTable("BeginTime") = StationControl(Index, index2).Start_Time
        ' rsTable("EndTime") = StationControl(Index, index2).End_Time
        rsTable("StnDescription") = STN_INFO(Index).desc
        rsTable("StnSysID") = STN_INFO(Index).SysID
        rsTable("StnAbrev") = STN_INFO(Index).Abrev
        rsTable("Station") = Index
        rsTable("Shift") = index2
        rsTable("Engineer") = JobInfo(Index, index2).Engineer + " "           'Prevent zero length string
        rsTable("Vehicle") = JobInfo(Index, index2).Vehicle + " "             'Prevent zero length string
        rsTable("StartOp") = JobInfo(Index, index2).Start_Op + " "            'Prevent zero length string
        rsTable("EndOp") = JobInfo(Index, index2).End_Op + " "                'Prevent zero length string
        If (STN_INFO(Index).Type = STN_LEAKTEST_TYPE) Then
            rsTable("CanId") = "in vehicle"
            rsTable("CanVol") = 0
            rsTable("WorkCap") = 0
        Else
            rsTable("CanId") = StationCanister(Index, index2).Description + " "                'Prevent zero length string
            rsTable("CanVol") = StationCanister(Index, index2).WorkingVolume
            rsTable("WorkCap") = StationCanister(Index, index2).WorkingCapacity
        End If
        rsTable("Comments") = JobInfo(Index, index2).Comment + " "            'Prevent zero length string
        rsTable("BeginBaro") = JobInfo(Index, index2).Start_Baro
        ' rsTable("EndBaro") = JobInfo(Index, index2)._End_Baro
        ' rsTable("EndOK") = JobInfo(Index, index2).End_OK
        rsTable("StnType") = STN_INFO(Index).Type
        rsTable("AdfTankType") = STN_INFO(Index).ADF_TANKTYPE
        rsTable("AdfVaporTankVolume") = StationCfg_ADF(Index, index2).VaporGenTankVol
        rsTable("AdfVaporTankLevelTol") = StationCfg_ADF(Index, index2).VaporGenLevelTol
        rsTable("AdfStorageTankVolume") = StationCfg_ADF(Index, index2).FuelStorageTankVol
        rsTable("AdfStorageTankLevelTol") = StationCfg_ADF(Index, index2).FuelStorageLevelTol
        rsTable("PurgeMfcEUMax") = Stn_AIO(Index, asPurgeAirFlow).EuMax
        rsTable("LiveFuelMfcEUMax") = Stn_AIO(Index, asLiveFuelVaporFlow).EuMax
        rsTable("ButGramsPerLiter") = GramsPerLiter
        If STN_INFO(Index).Type = STN_ORVR2_TYPE And StationRecipe(Index, index2).UseHiRangeMFC Then
            rsTable("ButMfcDensityMult") = STN_INFO(Index).ButMfc2DensityMult
            rsTable("ButMfcEUMax") = Stn_AIO(Index, asButaneORVRFlow).EuMax
            rsTable("NitMfcEUMax") = Stn_AIO(Index, asNitrogenORVRFlow).EuMax
        Else
            rsTable("ButMfcDensityMult") = STN_INFO(Index).ButMfcDensityMult
            rsTable("ButMfcEUMax") = Stn_AIO(Index, asButaneFlow).EuMax
            rsTable("NitMfcEUMax") = Stn_AIO(Index, asNitrogenFlow).EuMax
        End If
        rsTable("RemTaskID") = StnRemoteTask(Index, index2).TaskID
        If USINGREMAVLFILES Then
            rsTable("AVL_FileRoot") = StnRemoteTask(Index, index2).AVL_FileRoot
        Else
            rsTable("AVL_FileRoot") = "na"
        End If
        rsTable("UsingAuxOutputs") = USING_AUX_OUTPUTS
        rsTable("NrAuxOutputs") = NR_AUX_OUTPUTS
        rsTable("AuxOutDesc1") = DESC_AUX_OUTPUT1
        rsTable("AuxOutDesc2") = DESC_AUX_OUTPUT2
        rsTable("AuxOutDesc3") = DESC_AUX_OUTPUT3
        rsTable("AuxOutDesc4") = DESC_AUX_OUTPUT4
        rsTable("UsingCommonTC") = USINGCOMMONTC
        rsTable("UsingLineVolume") = USINGLINEVOLUME
        rsTable("UsingRemCanLoad") = USINGREMCANLOAD
        rsTable("UsingRemStatusMon") = USINGREMSTSMON
        rsTable("UsingAVL_Files") = USINGREMAVLFILES
        rsTable("UsingPurgeSeries") = USINGPURGESERIES
        rsTable("UsingPurgeOven") = USINGPURGEOVEN
        rsTable("UsingWaterBath") = USINGWATERBATH
        rsTable("ActualMaxCycle") = StationControl(Index, index2).CurrCycle
        rsTable("MaxSheathTempForAdfDrain") = MaxSheathTempForAdfDrain
        rsTable("DeadLiveFuelDensity") = DeadLiveFuelDensity
        rsTable("WeakLiveFuelDensity") = WeakLiveFuelDensity
      
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

Sub Header_Update(Index As Integer, index2 As Integer)
' Procedure Name:   Header_Update
' Description:      This routine updates the file Header data
'                   with the End of Test Values.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 44

Dim Criterion, filename As String

Dim dbDbase As Database
Dim rsTable As Recordset

    If StationControl(Index, index2).DBFile <> "" Then
    
        'Update Header data in data file
                                                                                                                    
        Criterion = _
            "SELECT * FROM [Header] WHERE [Header].[Course] = " & StationControl(Index, index2).Course & " "
        Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
        Set rsTable = dbDbase.OpenRecordset(Criterion, dbOpenDynaset)
    
        If rsTable.BOF Then
            rsTable.AddNew
            rsTable("Course") = StationControl(Index, index2).Course
        Else
          rsTable.MoveFirst
          rsTable.Edit
        End If
       
        rsTable("Job Number") = StationControl(Index, index2).Job_Number
        rsTable("Job Description") = StationControl(Index, index2).Job_Description
        rsTable("Report Filename") = StationControl(Index, index2).RptFile
        rsTable("BeginTime") = StationControl(Index, index2).Start_Time
        rsTable("EndTime") = StationControl(Index, index2).End_Time
        rsTable("Station") = Index
        rsTable("Shift") = index2
        rsTable("Engineer") = JobInfo(Index, index2).Engineer + " "           'Prevent zero length string
        rsTable("Vehicle") = JobInfo(Index, index2).Vehicle + " "             'Prevent zero length string
        rsTable("StartOp") = JobInfo(Index, index2).Start_Op + " "            'Prevent zero length string
        rsTable("EndOp") = JobInfo(Index, index2).End_Op + " "                'Prevent zero length string
        If (STN_INFO(Index).Type = STN_LEAKTEST_TYPE) Then
            rsTable("CanId") = "in vehicle"
            rsTable("CanVol") = 0
            rsTable("WorkCap") = 0
        Else
            rsTable("CanId") = StationCanister(Index, index2).Description + " "                'Prevent zero length string
            rsTable("CanVol") = StationCanister(Index, index2).WorkingVolume
            rsTable("WorkCap") = StationCanister(Index, index2).WorkingCapacity
        End If
        rsTable("Comments") = JobInfo(Index, index2).Comment + " "            'Prevent zero length string
        rsTable("BeginBaro") = JobInfo(Index, index2).Start_Baro
        rsTable("EndBaro") = JobInfo(Index, index2).End_Baro
        rsTable("Course") = StationControl(Index, index2).Course
        rsTable("EndOK") = JobInfo(Index, index2).End_OK
        rsTable("StnType") = STN_INFO(Index).Type
        rsTable("AdfTankType") = STN_INFO(Index).ADF_TANKTYPE
        rsTable("AdfVaporTankVolume") = StationCfg_ADF(Index, index2).VaporGenTankVol
        rsTable("AdfVaporTankLevelTol") = StationCfg_ADF(Index, index2).VaporGenLevelTol
        rsTable("AdfStorageTankVolume") = StationCfg_ADF(Index, index2).FuelStorageTankVol
        rsTable("AdfStorageTankLevelTol") = StationCfg_ADF(Index, index2).FuelStorageLevelTol
        rsTable("PurgeMfcEUMax") = Stn_AIO(Index, asPurgeAirFlow).EuMax
        rsTable("LiveFuelMfcEUMax") = Stn_AIO(Index, asLiveFuelVaporFlow).EuMax
        rsTable("ButGramsPerLiter") = GramsPerLiter
        If STN_INFO(Index).Type = STN_ORVR2_TYPE And StationRecipe(Index, index2).UseHiRangeMFC Then
            rsTable("ButMfcDensityMult") = STN_INFO(Index).ButMfc2DensityMult
            rsTable("ButMfcEUMax") = Stn_AIO(Index, asButaneORVRFlow).EuMax
            rsTable("NitMfcEUMax") = Stn_AIO(Index, asNitrogenORVRFlow).EuMax
        Else
            rsTable("ButMfcDensityMult") = STN_INFO(Index).ButMfcDensityMult
            rsTable("ButMfcEUMax") = Stn_AIO(Index, asButaneFlow).EuMax
            rsTable("NitMfcEUMax") = Stn_AIO(Index, asNitrogenFlow).EuMax
        End If
        rsTable("CurVersion") = USINGRELEASEDATE
        rsTable("RemTaskID") = StnRemoteTask(Index, index2).TaskID + " "            'Prevent zero length string
        rsTable("UsingAuxOutputs") = USING_AUX_OUTPUTS
        rsTable("NrAuxOutputs") = NR_AUX_OUTPUTS
        rsTable("AuxOutDesc1") = DESC_AUX_OUTPUT1
        rsTable("AuxOutDesc2") = DESC_AUX_OUTPUT2
        rsTable("AuxOutDesc3") = DESC_AUX_OUTPUT3
        rsTable("AuxOutDesc4") = DESC_AUX_OUTPUT4
        rsTable("UsingCommonTC") = USINGCOMMONTC
        rsTable("UsingLineVolume") = USINGLINEVOLUME
        rsTable("UsingRemCanLoad") = USINGREMCANLOAD
        rsTable("UsingPurgeSeries") = USINGPURGESERIES
        rsTable("ActualMaxCycle") = StationControl(Index, index2).CurrCycle
    
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

Sub Purge_Totalize(Index As Integer, index2 As Integer)
'
' Routine Name:    Purge_Totalize
' Description:     Updates cumulative purge flow
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 66
        
    ' TOTALIZE PURGE FLOW
    '   First, determine Elapsed Time (in Minutes) since previous totalize
    DeltaTotalTimer(Index, index2) = (StationControl(Index, index2).TestTimer - PreviousTotalTimer(Index, index2)) / 60
    PreviousTotalTimer(Index, index2) = StationControl(Index, index2).TestTimer
            
    '   Update cumulative Purge Flow
    deltaVol(Index, index2) = Stn_AIO(Index, asPurgeAirFlow).EUValue * DeltaTotalTimer(Index, index2)
    PurgeControl(Index, index2).Purge_Total = PurgeControl(Index, index2).Purge_Total + deltaVol(Index, index2)
            
    '   Update cumulative Purge Values
    PurgeControl(Index, index2).Purge_Volumes = PurgeControl(Index, index2).Purge_Total / StationCanister(Index, index2).WorkingVolume
    
    ' update (debug only) Net Purge Flow calculations
    If Not NotDebugPURGE Then
        If Not tardone(Index, index2) And DateDiff("s", PurgeControl(Index, index2).StartTime, Now) > 60 Then
            tarmin(Index, index2) = StationControl(Index, index2).TestTimer / 60
            tarvol(Index, index2) = StationControl(Index, index2).Actual
            tardone(Index, index2) = True
        End If
        ' Update Net Elapsed Time(in minutes) for this Purge Cycle
        netmin(Index, index2) = (StationControl(Index, index2).TestTimer / 60) - tarmin(Index, index2)
        ' Update Net Flow Rate for this Purge Cycle
        If netmin(Index, index2) <> 0 Then
            netflow(Index, index2) = (StationControl(Index, index2).Actual - tarvol(Index, index2)) / netmin(Index, index2)
        End If
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

Sub Default_Write(Index As Integer, index2 As Integer, flag As Integer)
' flag = 0 normal write; flag = 1 unused; flag = 2 unused
'
' Function Name:    Default_Write
' Author:           Analytical Process Programmer     9/9/09
' Description:      Updates the data file with other information
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 643

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String
Dim temp As Single

' Using the open DB File
If Len(StationControl(Index, index2).DBFile) > 0 Then

    ' which type of Default_Write is this?
    Select Case flag
        Case NORMALUPDATE
    
        Case Else
            
    End Select
        
  
  ' update reports for running stations that are not purging, loading, or leak checking
  Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
  Set rsTable = dbDbase.OpenRecordset("data")
  rsTable.AddNew
  
    ' GENERAL
    rsTable("Course") = StationControl(Index, index2).Course
    rsTable("Mode") = StationControl(Index, index2).Mode
    rsTable("Phase") = 0
    rsTable("ModeDesc") = ModeDescShort(StationControl(Index, index2).Mode)
    rsTable("Time") = Now
    rsTable("TestTime") = StationControl(Index, index2).TestTimer
    rsTable("Cycle") = StationControl(Index, index2).CurrCycle
    rsTable("Actual") = 0
    
    ' AIR
    rsTable("PATemp") = PATemp
    rsTable("PARH") = PAHum
    rsTable("Moisture") = PAMoisture
    rsTable("Baro") = AmbBaro
    
    ' FLOWS
    rsTable("PurgeFlow") = Stn_AIO(Index, asPurgeAirFlow).EUValue
    rsTable("NitFlow") = Stn_Nit_Flow_PV(Index, index2)
    rsTable("BtnFlow") = Stn_Btn_Flow_PV(Index, index2)
    
    ' SCALES
    If StationRecipe(Index, index2).UseAuxScale = True Then
       rsTable("AuxScale") = StationControl(Index, index2).AuxScaleWt
    Else
       rsTable("AuxScale") = 0
    End If
    If StationRecipe(Index, index2).UsePriScale = True Then
       rsTable("PriScale") = StationControl(Index, index2).PriScaleWt
    Else
       rsTable("PriScale") = 0
    End If
    
'    ' STATION TC's
    If Stn_UseTC(Index, index2) Then
      rsTable("TC1Temp") = Stn_AIO(Index, asStationTC1).EUValue
      rsTable("TC2Temp") = Stn_AIO(Index, asStationTC2).EUValue
    End If
    
    ' COMMON TC's
    If USINGCOMMONTC Then
'       If Stn_CommonTC(Index, index2) = True Then
            rsTable("CommonTC1") = Com_AIO(acCommonTC1).EUValue
            rsTable("CommonTC2") = Com_AIO(acCommonTC2).EUValue
            rsTable("CommonTC3") = Com_AIO(acCommonTC3).EUValue
            rsTable("CommonTC4") = Com_AIO(acCommonTC4).EUValue
            rsTable("CommonTC5") = Com_AIO(acCommonTC5).EUValue
            rsTable("CommonTC6") = Com_AIO(acCommonTC6).EUValue
'       End If
    End If

    If USINGPURGEOVEN And StationRecipe(Index, index2).PurgeOven Then
        rsTable("PurgeOvenTemp") = Stn_AIO(Index, asPurgeOvenTemp).EUValue
    Else
        rsTable("PurgeOvenTemp") = 0
    End If


    ' Live Fuel
    If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE)) Then
        rsTable("LiveFuelCycles") = StationControl(Index, 1).LiveFuelCycleCount
        If STN_INFO(Index).ADF_TANKTYPE > 0 Then
            rsTable("LiveFuelTemp") = Stn_AIO(Index, asFuelTankTemp).EUValue
            rsTable("LiveFuelLevel") = Stn_AIO(Index, asFuelTankLevel).EUValue
            rsTable("FuelStorageLevel") = Stn_AIO(Index, asStorageTankLevel).EUValue
            If ((Stn_AIO(Index, asFuelVaporTemp).addr <> 0) Or (Stn_AIO(Index, asFuelVaporTemp).chan <> 0)) Then
                rsTable("LiveFuelVaporTemp") = Stn_AIO(Index, asFuelVaporTemp).EUValue
            Else
                rsTable("LiveFuelVaporTemp") = 0
            End If
            If USINGWATERBATH And STN_INFO(Index).ADF_DEF.hasADF_WaterBath Then
                rsTable("WaterBathTemp") = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
            Else
                rsTable("WaterBathTemp") = 0
            End If
        Else
            rsTable("LiveFuelTemp") = 0
            rsTable("LiveFuelLevel") = 0
            rsTable("FuelStorageLevel") = 0
            rsTable("LiveFuelVaporTemp") = 0
            rsTable("WaterBathTemp") = 0
        End If
    End If

  ' Flag to indicate that the DB has been updated
  StationControl(Index, index2).NewDataInDB = True

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

Sub LT2_Write(ByVal Index As Integer, ByVal index2 As Integer, ByVal sTxt As String, ByRef sData As LT2_ReportData)
'
' Function Name:    LT2_Write
' Author:           Brunrose     2/2018
' Description:      Updates the data file with leaktest data events
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 876

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String
Dim temp As Single

    ' Using the open DB File
    If Len(StationControl(Index, index2).DBFile) > 0 Then
    
        ' update station that is leaktesting
        Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
        Set rsTable = dbDbase.OpenRecordset("LeakTestEvents")
        rsTable.AddNew
        
        ' GENERAL
        rsTable("DTS") = sData.ClkTime
        rsTable("Timer") = sData.SecTimer
        rsTable("Comment") = sTxt
        
        ' LeakTest
        rsTable("Deff") = sData.EffDia
        rsTable("QN2") = sData.NitFlow
        rsTable("Pin") = sData.InPress
        rsTable("Patm") = sData.AtmPress
        rsTable("TN2") = sData.NitTemp
        rsTable("SGN2") = SGN2
    
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

Sub LT_Write(Index As Integer, index2 As Integer, flag As Integer)
'
' flag
'0 = normal write
'8 = begin-of-leaktest write
'9 = end-of-leaktest write
'
'
' Function Name:    LT_Write
' Author:           Brunrose     1/2018
' Description:      Updates the data file with leaktest setup
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 678

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String
Dim temp As Single

    ' Using the open DB File
    If Len(StationControl(Index, index2).DBFile) > 0 Then
    
        ' which type of LeakTest_Write is this?
        Select Case flag
            Case NORMALUPDATE
        
            Case LT_BEGIN
                
            Case LT_DONE
                
            Case Else
                
        End Select
            
      
        ' Shuffle LeakTest Data sets
        StnLT2Data(Index, index2, 9) = StnLT2Data(Index, index2, 8)
        StnLT2Data(Index, index2, 8) = StnLT2Data(Index, index2, 7)
        StnLT2Data(Index, index2, 7) = StnLT2Data(Index, index2, 6)
        StnLT2Data(Index, index2, 6) = StnLT2Data(Index, index2, 5)
        StnLT2Data(Index, index2, 5) = StnLT2Data(Index, index2, 4)
        StnLT2Data(Index, index2, 4) = StnLT2Data(Index, index2, 3)
        StnLT2Data(Index, index2, 3) = StnLT2Data(Index, index2, 2)
        StnLT2Data(Index, index2, 2) = StnLT2Data(Index, index2, 1)
        StnLT2Data(Index, index2, 1) = StnLT2Data(Index, index2, 0)
    
        ' update reports for running stations that are leaktesting
        Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
        Set rsTable = dbDbase.OpenRecordset("data")
        rsTable.AddNew
      
        ' GENERAL
        rsTable("Course") = StationControl(Index, index2).Course
        rsTable("Mode") = StationControl(Index, index2).Mode
        rsTable("Phase") = 0
        rsTable("ModeDesc") = ModeDescShort(StationControl(Index, index2).Mode)
        rsTable("Time") = Now
        rsTable("TestTime") = StationControl(Index, index2).TestTimer
        rsTable("Cycle") = StationControl(Index, index2).CurrCycle
        rsTable("Actual") = 0
        
        ' AIR
        rsTable("PATemp") = PATemp
        rsTable("PARH") = PAHum
        rsTable("Moisture") = PAMoisture
        rsTable("Baro") = AmbBaro
        
        ' FLOWS
        rsTable("NitFlow") = Stn_AIO(Index, asNitrogenFlow).EUValue
        
        ' STATION TC's
        If Stn_UseTC(Index, index2) Then
          rsTable("TC1Temp") = Stn_AIO(Index, asStationTC1).EUValue
          rsTable("TC2Temp") = Stn_AIO(Index, asStationTC2).EUValue
        End If
        
        ' COMMON TC's
        If USINGCOMMONTC Then
    '       If Stn_CommonTC(Index, index2) = True Then
                rsTable("CommonTC1") = Com_AIO(acCommonTC1).EUValue
                rsTable("CommonTC2") = Com_AIO(acCommonTC2).EUValue
                rsTable("CommonTC3") = Com_AIO(acCommonTC3).EUValue
                rsTable("CommonTC4") = Com_AIO(acCommonTC4).EUValue
                rsTable("CommonTC5") = Com_AIO(acCommonTC5).EUValue
                rsTable("CommonTC6") = Com_AIO(acCommonTC6).EUValue
    '       End If
        End If
    
        ' LeakTest
        rsTable("LT_InletPress") = Stn_AIO(Index, asLtInletPress).EUValue
        rsTable("LT_InletTemp") = Stn_AIO(Index, asLtN2Temp).EUValue
        rsTable("LT_Deff") = Deff
    
         ' Flag to indicate that the DB has been updated
          StationControl(Index, index2).NewDataInDB = True
        
        ' update LeakTest DataSet 0
        StnLT2Data(Index, index2, 0).ClkTime = rsTable("Time")
        StnLT2Data(Index, index2, 0).NitFlow = rsTable("NitFlow")
        StnLT2Data(Index, index2, 0).NitTemp = rsTable("LT_InletTemp")
        StnLT2Data(Index, index2, 0).InPress = rsTable("LT_InletPress")
        StnLT2Data(Index, index2, 0).EffDia = rsTable("LT_Deff")
        StnLT2Data(Index, index2, 0).SecTimer = rsTable("TestTime")
        StnLT2Data(Index, index2, 0).isBlank = False
            
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

Sub Purge_Write(Index As Integer, index2 As Integer, flag As Integer)
'
' flag = 0 normal write; flag = 1 start Purge; flag = 2 end Purge
'
' Function Name:    Purge_Write
' Author:           Analytical Process Programmer     8/8/96
' Description:      Updates the data file with purge information
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 6

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String
Dim temp As Single
Dim auxWt As Single
Dim priWt As Single

    ' Using the open DB File
    If Len(StationControl(Index, index2).DBFile) > 0 Then
    
            
        ' Determine Elapsed Time (in Minutes) since previous Purge_Write
        DeltaTimer(Index, index2) = (StationControl(Index, index2).TestTimer - PreviousReportTimer(Index, index2)) / 60
        PreviousReportTimer(Index, index2) = StationControl(Index, index2).TestTimer
           
        ' Shuffle PurgeData sets
        StnPurgeData(Index, index2, 9) = StnPurgeData(Index, index2, 8)
        StnPurgeData(Index, index2, 8) = StnPurgeData(Index, index2, 7)
        StnPurgeData(Index, index2, 7) = StnPurgeData(Index, index2, 6)
        StnPurgeData(Index, index2, 6) = StnPurgeData(Index, index2, 5)
        StnPurgeData(Index, index2, 5) = StnPurgeData(Index, index2, 4)
        StnPurgeData(Index, index2, 4) = StnPurgeData(Index, index2, 3)
        StnPurgeData(Index, index2, 3) = StnPurgeData(Index, index2, 2)
        StnPurgeData(Index, index2, 2) = StnPurgeData(Index, index2, 1)
        StnPurgeData(Index, index2, 1) = StnPurgeData(Index, index2, 0)
        
        ' update reports for stations purging
        Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
        Set rsTable = dbDbase.OpenRecordset("data")
        rsTable.AddNew
                    
            rsTable("Course") = StationControl(Index, index2).Course
            rsTable("Mode") = VBPURGE
            rsTable("Phase") = PurgeControl(Index, index2).Phase
            rsTable("ModeDesc") = ModeDescShort(VBPURGE)
            rsTable("PhaseDesc") = PurgePhaseDesc(PurgeControl(Index, index2).Phase)
            rsTable("Time") = Now                                       ' time = now
            rsTable("TestTime") = StationControl(Index, index2).TestTimer
            rsTable("Cycle") = StationControl(Index, index2).CurrCycle
            rsTable("Actual") = StationControl(Index, index2).Actual
            rsTable("PurgeFlow") = Stn_AIO(Index, asPurgeAirFlow).EUValue
            rsTable("PATemp") = PATemp
            rsTable("PARH") = PAHum
            rsTable("PurgeVol") = PurgeControl(Index, index2).Purge_Total
            rsTable("PurgeDP") = Stn_AIO(Index, asPurgeDiffPress).EUValue
            '   NOPURGE = 0             ' No Purge
            '   PURGEBYTIME = 1         ' Purge by time
            '   PURGEBYVOLUME = 2       ' Purge by Canister Volumes
            '   PURGEAUXONLY = 3        ' Purge Aux Canister Only
            '   PURGEBYPROFILE = 4      ' Purge by Profile
            '   PURGEBYWC = 5           ' Purge by Canister Working Capacity
            '   PURGETOTARGET = 6       ' Purge to Target Weight
            '   PURGETOUNDOLOAD = 7     ' Purge by Weight (= Weight Loaded Last Cycle)
            '   PURGEBYLITERS = 8       ' Purge by Liters of PurgeAir Flow
            rsTable("PurgeMethod") = StationRecipe(Index, index2).Purge_Method
            
            ' which type of Purge_Write is this?
            Select Case flag
                Case NORMALUPDATE                                       ' all but first and last time
                    ' cumulative Purge Flow updated once per second by Purge_Totalize routine
                    ' scale weights
                    If StationRecipe(Index, index2).UsePriScale Then
                        rsTable("priscale") = StationControl(Index, index2).PriScaleWt
                    Else
                        rsTable("priscale") = CSng(0)
                    End If
                    If StationRecipe(Index, index2).UseAuxScale Then
                        rsTable("auxscale") = StationControl(Index, index2).AuxScaleWt
                    Else
                        rsTable("auxscale") = CSng(0)
                    End If
                
                Case PURGEBEGIN                                          ' first time only
                    ' Initialize cumulative Purge Flow
                    StationControl(Index, index2).Actual = 0
                    PurgeControl(Index, index2).Purge_Total = StationControl(Index, index2).Actual
                    ' scale weights
                    If StationRecipe(Index, index2).UsePriScale Then
                        PurgeControl(Index, index2).PriWt_Start = StationControl(Index, index2).PriScaleWt
                        rsTable("priscale") = PurgeControl(Index, index2).PriWt_Start
                    Else
                        PurgeControl(Index, index2).PriWt_Start = CSng(0)
                        rsTable("priscale") = PurgeControl(Index, index2).PriWt_Start
                    End If
                    StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Purge_StartWeight_Pri = PurgeControl(Index, index2).PriWt_Start
                    If StationRecipe(Index, index2).UseAuxScale Then
                        PurgeControl(Index, index2).AuxWt_Start = StationControl(Index, index2).AuxScaleWt
                        rsTable("auxscale") = PurgeControl(Index, index2).AuxWt_Start
                    Else
                        PurgeControl(Index, index2).AuxWt_Start = CSng(0)
                        rsTable("auxscale") = PurgeControl(Index, index2).AuxWt_Start
                    End If
                    StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Purge_StartWeight_Aux = PurgeControl(Index, index2).AuxWt_Start
                
                Case PURGEDONE                                           ' last time only
                    ' Request final totalization
                    Purge_Totalize Index, index2
                    ' scale weights
                    If StationRecipe(Index, index2).UsePriScale Then
                        PurgeControl(Index, index2).PriWt_End = StationControl(Index, index2).PriScaleWt
                        rsTable("priscale") = PurgeControl(Index, index2).PriWt_End
                    Else
                        PurgeControl(Index, index2).PriWt_End = CSng(0)
                        rsTable("priscale") = PurgeControl(Index, index2).PriWt_End
                    End If
                    StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Purge_EndWeight_Pri = PurgeControl(Index, index2).PriWt_End
                    If StationRecipe(Index, index2).UseAuxScale Then
                        PurgeControl(Index, index2).AuxWt_End = StationControl(Index, index2).AuxScaleWt
                        rsTable("auxscale") = PurgeControl(Index, index2).AuxWt_End
                    Else
                        PurgeControl(Index, index2).AuxWt_End = CSng(0)
                        rsTable("auxscale") = PurgeControl(Index, index2).AuxWt_End
                    End If
                    StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Purge_EndWeight_Aux = PurgeControl(Index, index2).AuxWt_End
            End Select
                
            rsTable("TotalWtChg") = PurgeControl(Index, index2).TotalWtChg
            rsTable("TotalWtChgRate") = PurgeControl(Index, index2).TotalWtChgRate
            
            rsTable("Moisture") = PAMoisture
            rsTable("Baro") = AmbBaro
            If Stn_UseTC(Index, index2) Then
              rsTable("TC1Temp") = Stn_AIO(Index, asStationTC1).EUValue
              rsTable("TC2Temp") = Stn_AIO(Index, asStationTC2).EUValue
            End If
            
            ' Debug Data
            ' Debug Data
            ' Debug Data
            If Not NotDebugMMW Or Not NotDebugPURGE Then
              rsTable("TC3Temp") = PreviousReportTimer(Index, index2)
              rsTable("TC4Temp") = DeltaTimer(Index, index2)
              rsTable("TC5Temp") = flag
            End If
            ' Debug Data
            ' Debug Data
            ' Debug Data
                    
            
            If USINGCOMMONTC Then
            '       If Stn_CommonTC(Index, index2) = True Then
                    rsTable("CommonTC1") = Com_AIO(acCommonTC1).EUValue
                    rsTable("CommonTC2") = Com_AIO(acCommonTC2).EUValue
                    rsTable("CommonTC3") = Com_AIO(acCommonTC3).EUValue
                    rsTable("CommonTC4") = Com_AIO(acCommonTC4).EUValue
                    rsTable("CommonTC5") = Com_AIO(acCommonTC5).EUValue
                    rsTable("CommonTC6") = Com_AIO(acCommonTC6).EUValue
            '       End If
            End If
            
            If USINGPURGEOVEN And StationRecipe(Index, index2).PurgeOven Then
                rsTable("PurgeOvenTemp") = Stn_AIO(Index, asPurgeOvenTemp).EUValue
            Else
                rsTable("PurgeOvenTemp") = 0
            End If
    
            ' Live Fuel
            If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE)) Then
                rsTable("LiveFuelCycles") = StationControl(Index, 1).LiveFuelCycleCount
                If STN_INFO(Index).ADF_TANKTYPE > 0 Then
                    rsTable("LiveFuelTemp") = Stn_AIO(Index, asFuelTankTemp).EUValue
                    rsTable("LiveFuelLevel") = Stn_AIO(Index, asFuelTankLevel).EUValue
                    rsTable("FuelStorageLevel") = Stn_AIO(Index, asStorageTankLevel).EUValue
                    If ((Stn_AIO(Index, asFuelVaporTemp).addr <> 0) Or (Stn_AIO(Index, asFuelVaporTemp).chan <> 0)) Then
                        rsTable("LiveFuelVaporTemp") = Stn_AIO(Index, asFuelVaporTemp).EUValue
                    Else
                        rsTable("LiveFuelVaporTemp") = 0
                    End If
                    If USINGWATERBATH And STN_INFO(Index).ADF_DEF.hasADF_WaterBath Then
                        rsTable("WaterBathTemp") = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
                    Else
                        rsTable("WaterBathTemp") = 0
                    End If
                Else
                    rsTable("LiveFuelTemp") = 0
                    rsTable("LiveFuelLevel") = 0
                    rsTable("FuelStorageLevel") = 0
                    rsTable("LiveFuelVaporTemp") = 0
                    rsTable("WaterBathTemp") = 0
                End If
            End If

            
            ' update PurgeDataSet 0
            StnPurgeData(Index, index2, 0).ClkTime = rsTable("Time")
            StnPurgeData(Index, index2, 0).PrgFlow = rsTable("PurgeFlow")
            StnPurgeData(Index, index2, 0).PrgTemp = rsTable("PATemp")
            StnPurgeData(Index, index2, 0).PrgHumd = rsTable("Moisture")
            StnPurgeData(Index, index2, 0).PriScle = rsTable("PriScale")
            StnPurgeData(Index, index2, 0).AuxScle = rsTable("AuxScale")
            StnPurgeData(Index, index2, 0).WtChange = rsTable("TotalWtChg")
            StnPurgeData(Index, index2, 0).WtChgRate = rsTable("TotalWtChgRate")
            StnPurgeData(Index, index2, 0).VolTotl = rsTable("PurgeVol")
            StnPurgeData(Index, index2, 0).TstTimr = rsTable("TestTime")
            StnPurgeData(Index, index2, 0).isBlank = False
            
            ' Flag to indicate that the DB has been updated
            StationControl(Index, index2).NewDataInDB = True
              
        rsTable.Update
        rsTable.Close
        dbDbase.Close
      
    End If
    
    ' Update Statistics
    Update_Stats Index, index2
    
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

Sub OOT_Write_Data(station As Integer, Shift As Integer, Mode As Integer, errnum As Integer)

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 781

    ' Update data report for OOT station.
    If StationControl(station, Shift).DBFile <> "" Then
        Set dbDbase = OpenDatabase(StationControl(station, Shift).DBFile)
        Set rsTable = dbDbase.OpenRecordset("data")
        rsTable.AddNew
        rsTable("Time") = Now()
        rsTable("TestTime") = StationControl(station, Shift).TestTimer
        rsTable("Course") = StationControl(station, Shift).Course
        rsTable("Cycle") = StationControl(station, Shift).CurrCycle
        rsTable("Mode") = Mode
        rsTable("ReportCode") = errnum
        rsTable("ModeDesc") = ModeDescShort(Mode)
        rsTable("ReportCodeDesc") = ReportCodeDesc(errnum)
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

Sub Leak_Write(Index As Integer, index2 As Integer, reportcode As Integer, resultcode As Integer)
' reportcode = 0 normal write
' reportcode = 800 begin leak check phase 0 (purging)
' reportcode = 801 begin leak check phase 1 (pressurize)
' reportcode = 802 begin leak check phase 2 (testing)
' reportcode = 808 leak check done
' reportcode = 811 operator continue after LeakCheck Failure
' reportcode = 814 automatic continue after LeakCheck Failure
'
' resultcode = 0 no result
' resultcode = 1 failed - Purge Timeout
' resultcode = 2 failed - Pressurize Timeout
' resultcode = 3 failed - Excessive Leak Rate
' resultcode = 8 aborted
' resultcode = 9 passed
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 552

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String

    ' Using the Current DB File
    If StationControl(Index, index2).DBFile <> "" Then
    
        ' Determine Elapsed Time (in Minutes) since previous Leak_Write
        DeltaTimer(Index, index2) = (StationControl(Index, index2).TestTimer - PreviousReportTimer(Index, index2)) / 60
        PreviousReportTimer(Index, index2) = StationControl(Index, index2).TestTimer
        
        If LeakCheckControl.Phase < LeakComplete Then
        
            ' Shuffle LeakData sets
            StnLeakData(Index, index2, 9) = StnLeakData(Index, index2, 8)
            StnLeakData(Index, index2, 8) = StnLeakData(Index, index2, 7)
            StnLeakData(Index, index2, 7) = StnLeakData(Index, index2, 6)
            StnLeakData(Index, index2, 6) = StnLeakData(Index, index2, 5)
            StnLeakData(Index, index2, 5) = StnLeakData(Index, index2, 4)
            StnLeakData(Index, index2, 4) = StnLeakData(Index, index2, 3)
            StnLeakData(Index, index2, 3) = StnLeakData(Index, index2, 2)
            StnLeakData(Index, index2, 2) = StnLeakData(Index, index2, 1)
            StnLeakData(Index, index2, 1) = StnLeakData(Index, index2, 0)
    
            ' Update data report (note: LeakCheckControl.Phase = LeakComplete is closing valves, etc.)
            Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
            Set rsTable = dbDbase.OpenRecordset("data")
            rsTable.AddNew
            ' GENERAL
            rsTable("Course") = StationControl(Index, index2).Course
            rsTable("Mode") = VBLEAK
            rsTable("Phase") = LeakCheckControl.Phase
            rsTable("ModeDesc") = ModeDescShort(VBLEAK)
            rsTable("PhaseDesc") = LeakPhaseDesc(LeakCheckControl.Phase)
            rsTable("Time") = Now()
            rsTable("TestTime") = StationControl(Index, index2).TestTimer
            rsTable("Cycle") = StationControl(Index, index2).CompletedCycles
            
            ' AIR
            rsTable("PATemp") = PATemp
            rsTable("PARH") = PAHum
            rsTable("Moisture") = PAMoisture
            rsTable("Baro") = AmbBaro
                  
            ' SCALES
            If StationRecipe(Index, index2).UseAuxScale = True Then
                rsTable("AuxScale") = StationControl(Index, index2).AuxScaleWt
            Else
                rsTable("AuxScale") = 0
            End If
            If StationRecipe(Index, index2).UsePriScale = True Then
                rsTable("PriScale") = StationControl(Index, index2).PriScaleWt
            Else
                rsTable("PriScale") = 0
            End If
        
            ' LEAKCHECK
            rsTable("Pressure") = Com_AIO(acComnPressSensor).EUValue
            rsTable("ReportCode") = reportcode
            rsTable("LeakCheckResult") = resultcode
            rsTable("ReportCodeDesc") = ReportCodeDesc(reportcode)
            rsTable("LeakCheckResultDesc") = LeakResultDesc(resultcode)
            rsTable("LeakCheckCanister") = LeakCanisterDesc(LeakCheckControl.Method)
            
            ' COMMON TC's
    '        If USINGCOMMONTC And Stn_CommonTC(Index, index2) Then
            If USINGCOMMONTC Then
                rsTable("CommonTC1") = Com_AIO(acCommonTC1).EUValue
                rsTable("CommonTC2") = Com_AIO(acCommonTC2).EUValue
                rsTable("CommonTC3") = Com_AIO(acCommonTC3).EUValue
                rsTable("CommonTC4") = Com_AIO(acCommonTC4).EUValue
                rsTable("CommonTC5") = Com_AIO(acCommonTC5).EUValue
                rsTable("CommonTC6") = Com_AIO(acCommonTC6).EUValue
            Else
                rsTable("CommonTC1") = 0
                rsTable("CommonTC2") = 0
                rsTable("CommonTC3") = 0
                rsTable("CommonTC4") = 0
                rsTable("CommonTC5") = 0
                rsTable("CommonTC6") = 0
            End If
            
            ' Live Fuel
            If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE)) Then
                rsTable("LiveFuelCycles") = StationControl(Index, 1).LiveFuelCycleCount
                If STN_INFO(Index).ADF_TANKTYPE > 0 Then
                    rsTable("LiveFuelTemp") = Stn_AIO(Index, asFuelTankTemp).EUValue
                    rsTable("LiveFuelLevel") = Stn_AIO(Index, asFuelTankLevel).EUValue
                    rsTable("FuelStorageLevel") = Stn_AIO(Index, asStorageTankLevel).EUValue
                Else
                    rsTable("LiveFuelTemp") = 0
                    rsTable("LiveFuelLevel") = 0
                    rsTable("FuelStorageLevel") = 0
                End If
            End If

            ' update LeakDataSet 0
            StnLeakData(Index, index2, 0).ClkTime = rsTable("Time")
            StnLeakData(Index, index2, 0).Pressure = rsTable("Pressure")
            StnLeakData(Index, index2, 0).Comment = " "
            StnLeakData(Index, index2, 0).TstTimr = rsTable("TestTime")
            StnLeakData(Index, index2, 0).isBlank = False
            
            ' Flag to indicate that the DB has been updated
            StationControl(Index, index2).NewDataInDB = True
    
            ' close db
            rsTable.Update
            rsTable.Close
            dbDbase.Close
            
        End If
    
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

Sub Load_Totalize(Index As Integer, index2 As Integer)
'
' Routine Name:    Load_Totalize
' Description:     Updates cumulative load (butane) flow
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 55
Static xcntr As Integer
Static xcntr2 As Integer
    ' TOTALIZE LOAD FLOW
    
'    xcntr = iif((xcntr < 32000), xcntr + 1, 0)
'    if (xcntr = 0) then
'       xcntr2 = iif((xcntr2 < 32000), 0)
'    end if
'Debug.Print "LoadTotalize - " & Format(xcntr, "####0")
    
ChgErrModule 3, 550
    ' First, determine Elapsed Time (in Minutes) since previous totalize
    DeltaTotalTimer(Index, index2) = (StationControl(Index, index2).TestTimer - PreviousTotalTimer(Index, index2)) / CDbl(60)
ChgErrModule 3, 5501
    PreviousTotalTimer(Index, index2) = StationControl(Index, index2).TestTimer
            
    ' Update Cumulative Flow Values
    Select Case STN_INFO(Index).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
ChgErrModule 3, 551
            ' Update Cumulative Butane Mass Flow
            If Stn_Btn_Flow_PV(Index, index2) > 0 Then
ChgErrModule 3, 5511
                ' Calculate the butane used (in grams) during this interval
                deltaWt(Index, index2) = CDbl(SlpmToGramsPerHour(Stn_Btn_Flow_PV(Index, index2), StationControl(Index, index2).BtnDensity) / CDbl(60)) * DeltaTotalTimer(Index, index2)
ChgErrModule 3, 5512
                ' Totalize butane
                LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).loadTotalGrams + CSng(deltaWt(Index, index2))
ChgErrModule 3, 5513
                ' Reduce the total butane capacity left
                ButaneSupply.CurrentOnHand = ButaneSupply.CurrentOnHand - deltaVol(Index, index2)
            End If
            ' Update Cumulative Nitrogen Volumetric Flow
            If Stn_Nit_Flow_PV(Index, index2) > 0 Then
ChgErrModule 3, 5514
                ' Calculate the nitrogen used (in liters) during this interval
                deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
ChgErrModule 3, 5515
                ' Totalize nitrogen
                LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
            End If
                
        Case STN_ORVR2_TYPE
ChgErrModule 3, 552
            ' Update Cumulative Butane Flow
            If Stn_Btn_Flow_PV(Index, index2) > 0 Then
                ' Calculate the butane used (in grams) during this interval
                deltaWt(Index, index2) = (SlpmToGramsPerHour(Stn_Btn_Flow_PV(Index, index2), StationControl(Index, index2).BtnDensity) / 60) * DeltaTotalTimer(Index, index2)
                ' Totalize butane
                LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).loadTotalGrams + deltaWt(Index, index2)
                ' Reduce the total butane capacity left
                ButaneSupply.CurrentOnHand = ButaneSupply.CurrentOnHand - deltaVol(Index, index2)
            End If
            ' Update Cumulative Nitrogen Volumetric Flow
            If Stn_Nit_Flow_PV(Index, index2) > 0 Then
                ' Calculate the nitrogen used (in liters) during this interval
                deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
                ' Totalize nitrogen
                LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
            End If
                
        Case STN_LIVEFUEL_TYPE
ChgErrModule 3, 553
            ' Update Cumulative Vapor Carrier Total
            If Stn_Nit_Flow_PV(Index, index2) > 0 Then
                ' Calculate the VaporCarrier volume during this interval
                deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
                ' Totalize VaporCarrier volume
                LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
            End If
            ' Update Cumulative Live Fuel Vapor Mass (Weight) Flow
            LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).TotalWtChg
        
        Case STN_LIVEREG_TYPE
ChgErrModule 3, 554
            If (StationRecipe(Index, index2).LiveFuel) Then
                ' Update Cumulative Vapor Carrier Volume Total
                If Stn_Nit_Flow_PV(Index, index2) > 0 Then
                    ' Calculate the VaporCarrier volume during this interval
                    deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
                    ' Totalize VaporCarrier volume
                    LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
                End If
                ' Update Cumulative Live Fuel Vapor Mass (Weight) Flow
                LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).TotalWtChg
           Else
                ' Update Cumulative Butane Mass Flow
                If Stn_Btn_Flow_PV(Index, index2) > 0 Then
                    ' Calculate the butane used (in grams) during this interval
                    deltaWt(Index, index2) = (SlpmToGramsPerHour(Stn_Btn_Flow_PV(Index, index2), StationControl(Index, index2).BtnDensity) / 60) * DeltaTotalTimer(Index, index2)
                    ' Totalize butane
                    LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).loadTotalGrams + deltaWt(Index, index2)
                    ' Reduce the total butane capacity left
                    ButaneSupply.CurrentOnHand = ButaneSupply.CurrentOnHand - deltaVol(Index, index2)
                End If
                ' Update Cumulative Nitrogen Volumetric Flow
                If Stn_Nit_Flow_PV(Index, index2) > 0 Then
                    ' Calculate the nitrogen used (in liters) during this interval
                    deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
                    ' Totalize nitrogen
                    LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
                End If
            End If
            
        Case STN_LIVEORVR2_TYPE
ChgErrModule 3, 555
            If (StationRecipe(Index, index2).LiveFuel) Then
                ' Update Cumulative Vapor Carrier Volume Total
                If Stn_Nit_Flow_PV(Index, index2) > 0 Then
ChgErrModule 3, 5551
                    ' Calculate the VaporCarrier volume during this interval
                    deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
ChgErrModule 3, 5552
                    ' Totalize VaporCarrier volume
                    LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
                End If
ChgErrModule 3, 5553
                ' Update Cumulative Live Fuel Vapor Mass (Weight) Flow
                LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).TotalWtChg
           Else
                ' Update Cumulative Butane Flow
                If Stn_Btn_Flow_PV(Index, index2) > 0 Then
ChgErrModule 3, 5554
                    ' Calculate the butane used (in grams) during this interval
                    deltaWt(Index, index2) = (SlpmToGramsPerHour(Stn_Btn_Flow_PV(Index, index2), StationControl(Index, index2).BtnDensity) / 60) * DeltaTotalTimer(Index, index2)
ChgErrModule 3, 5555
                    ' Totalize butane
                    LoadControl(Index, index2).loadTotalGrams = LoadControl(Index, index2).loadTotalGrams + deltaWt(Index, index2)
ChgErrModule 3, 5556
                    ' Reduce the total butane capacity left
                    ButaneSupply.CurrentOnHand = ButaneSupply.CurrentOnHand - deltaVol(Index, index2)
                End If
                ' Update Cumulative Nitrogen Volumetric Flow
                If Stn_Nit_Flow_PV(Index, index2) > 0 Then
ChgErrModule 3, 5557
                    ' Calculate the nitrogen used (in liters) during this interval
                    deltaVol(Index, index2) = Stn_Nit_Flow_PV(Index, index2) * DeltaTotalTimer(Index, index2)
ChgErrModule 3, 5558
                    ' Totalize nitrogen
                    LoadControl(Index, index2).LoadTotalLiters = LoadControl(Index, index2).LoadTotalLiters + deltaVol(Index, index2)
                End If
            End If
                
        Case STN_COMBO3_TYPE
            ' future
ChgErrModule 3, 556
                
        Case Else
            ' Do Nothing
ChgErrModule 3, 557
    End Select
            
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

Sub Load_Write(Index As Integer, index2 As Integer, flag As Integer)
'
' flag = 0 normal write
' flag = 1 start load
' flag = 2 end load
'
' Function Name:    Load_Write
' Author:           Analytical Process Programmer     8/8/96
' Description:      Updates the data file with Load information
'

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String
Dim temp As Single

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 5

    ' Using the open DB File
    If StationControl(Index, index2).DBFile <> "" Then
    
        ' Determine Elapsed Time (in Minutes) since previous Load_Write
        DeltaTimer(Index, index2) = (StationControl(Index, index2).TestTimer - PreviousReportTimer(Index, index2)) / 60
        PreviousReportTimer(Index, index2) = StationControl(Index, index2).TestTimer
        
        ' Shuffle LoadData sets
        StnLoadData(Index, index2, 9) = StnLoadData(Index, index2, 8)
        StnLoadData(Index, index2, 8) = StnLoadData(Index, index2, 7)
        StnLoadData(Index, index2, 7) = StnLoadData(Index, index2, 6)
        StnLoadData(Index, index2, 6) = StnLoadData(Index, index2, 5)
        StnLoadData(Index, index2, 5) = StnLoadData(Index, index2, 4)
        StnLoadData(Index, index2, 4) = StnLoadData(Index, index2, 3)
        StnLoadData(Index, index2, 3) = StnLoadData(Index, index2, 2)
        StnLoadData(Index, index2, 2) = StnLoadData(Index, index2, 1)
        StnLoadData(Index, index2, 1) = StnLoadData(Index, index2, 0)
    
        ' Update Load report for stations Loading.
        Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
        Set rsTable = dbDbase.OpenRecordset("data")
        rsTable.AddNew
        
        ' GENERAL
        rsTable("Course") = StationControl(Index, index2).Course
        rsTable("Mode") = VBLOAD
        rsTable("Phase") = LoadControl(Index, index2).Phase
        rsTable("ModeDesc") = ModeDescShort(VBLOAD)
        rsTable("PhaseDesc") = LoadPhaseDesc(LoadControl(Index, index2).Phase)
        rsTable("Time") = Now()
        rsTable("TestTime") = StationControl(Index, index2).TestTimer
        rsTable("Cycle") = StationControl(Index, index2).CurrCycle
        rsTable("Actual") = StationControl(Index, index2).Actual
        '   NOLOAD = 0             ' No load
        '   LOADBYTIME = 1         ' Load by time
        '   LOADBYWC = 2           ' Load by working capacity
        '   LOADBYWEIGHT = 3       ' Load by weight
        '   LOADBYBREAKTHRU = 4    ' Load by breakthrough
        '   LOADBYFID = 5          ' Load by FID Breakthrough
        rsTable("LoadMethod") = StationRecipe(Index, index2).Load_Method
        
        ' AIR
        rsTable("PATemp") = PATemp
        rsTable("PARH") = PAHum
        rsTable("Moisture") = PAMoisture
        rsTable("Baro") = AmbBaro
            
        rsTable("NitFlow") = Stn_Nit_Flow_PV(Index, index2)
        rsTable("BtnFlow") = Stn_Btn_Flow_PV(Index, index2)
        rsTable("Baro") = AmbBaro
        
        If Stn_Btn_Flow_PV(Index, index2) + Stn_Nit_Flow_PV(Index, index2) <= 0.0001 Then
            rsTable("Mix") = 0
        Else
            rsTable("Mix") = 100 * Stn_Btn_Flow_PV(Index, index2) / _
                (Stn_Btn_Flow_PV(Index, index2) + Stn_Nit_Flow_PV(Index, index2) + 0.00001)
        End If
        
    ' **************************************************************************************************************************************************
        Select Case flag
            Case NORMALUPDATE
                ' normal
                ' Totals are calculated once a second by Load_Totalize routine
                rsTable("LineLoss") = CSng(0)
                ' scale weights
                If StationRecipe(Index, index2).UsePriScale Then
                    rsTable("priscale") = StationControl(Index, index2).PriScaleWt
                Else
                    rsTable("priscale") = CSng(0)
                End If
                If StationRecipe(Index, index2).UseAuxScale Then
                    rsTable("auxscale") = StationControl(Index, index2).AuxScaleWt
                Else
                    rsTable("auxscale") = CSng(0)
                End If
                
            Case LOADBEGIN
                ' start of Load
                rsTable("LineLoss") = CSng(0)
                ' scale weights
                If StationRecipe(Index, index2).UsePriScale Then
                    LoadControl(Index, index2).PriWt_Start = StationControl(Index, index2).PriScaleWt
                    rsTable("priscale") = LoadControl(Index, index2).PriWt_Start
                Else
                    LoadControl(Index, index2).PriWt_Start = CSng(0)
                    rsTable("priscale") = LoadControl(Index, index2).PriWt_Start
                End If
                StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Load_StartWeight_Pri = LoadControl(Index, index2).PriWt_Start
                If StationRecipe(Index, index2).UseAuxScale Then
                    LoadControl(Index, index2).AuxWt_Start = StationControl(Index, index2).AuxScaleWt
                    rsTable("auxscale") = LoadControl(Index, index2).AuxWt_Start
                Else
                    LoadControl(Index, index2).AuxWt_Start = CSng(0)
                    rsTable("auxscale") = LoadControl(Index, index2).AuxWt_Start
                End If
                StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Load_StartWeight_Aux = LoadControl(Index, index2).AuxWt_Start
                CurrRate(Index, index2) = Load_Rate
                
            Case LOADDONE
                ' end of Load
                ' need to account for short interval at the end
                Load_Totalize CInt(Index), CInt(index2)
                ' scale weights
                If StationRecipe(Index, index2).UsePriScale Then
                    LoadControl(Index, index2).PriWt_End = StationControl(Index, index2).PriScaleWt
                    If StationRecipe(Index, index2).UseAuxScale Then
                        LoadControl(Index, index2).AuxWt_End = StationControl(Index, index2).AuxScaleWt
                    Else
                        LoadControl(Index, index2).AuxWt_End = CSng(0)
                    End If
                Else
                    LoadControl(Index, index2).PriWt_End = CSng(0)
                    If StationRecipe(Index, index2).UseAuxScale Then
                        ' aux scale but no primary; use wtchg at EOL not end-of-load
                        LoadControl(Index, index2).AuxWt_End = LoadControl(Index, index2).AuxWt_Start + LoadControl(Index, index2).AuxWtChgAtEOL
                    Else
                        LoadControl(Index, index2).AuxWt_End = CSng(0)
                    End If
                End If
                rsTable("priscale") = LoadControl(Index, index2).PriWt_End
                StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Load_EndWeight_Pri = LoadControl(Index, index2).PriWt_End
                rsTable("auxscale") = LoadControl(Index, index2).AuxWt_End
                StationCycleWeightData(Index, index2, StationControl(Index, index2).CurrCycle).Load_EndWeight_Aux = LoadControl(Index, index2).AuxWt_End
                            
                If USINGLINEVOLUME And StationControl(Index, index2).CompletedCycles = VALUE0 Then
    '   ******************************
    '   *** FORMULA CHECK REQUIRED ***
    '   ******************************
                    If Stn_Btn_FlowSP(Index, index2) > 0 Then
                        Select Case StationRecipe(Index, index2).Load_Method
                            Case LOADBYTIME
                                ' losses       =  line volume / slpm flow
                                '                 *  Load rate / one minute
                                rsTable("LineLoss") = (StationRecipe(Index, index2).LoadV / Stn_Btn_FlowSP(Index, index2)) _
                                           * (StationRecipe(Index, index2).Load_Rate / 60)
                                           
                            Case LOADBYWC
                                ' losses       =  line volume / slpm flow
                                '                 *  Load rate / one minute
                                rsTable("LineLoss") = (StationRecipe(Index, index2).LoadV / Stn_Btn_FlowSP(Index, index2)) _
                                           * (StationRecipe(Index, index2).Load_Rate / 60)
                                           
                            Case LOADBYWEIGHT
                                ' losses       =  line volume / slpm flow
                                '                 *  Load rate / one minute
                                rsTable("LineLoss") = (StationRecipe(Index, index2).LoadV / Stn_Btn_FlowSP(Index, index2)) _
                                           * (StationRecipe(Index, index2).Load_Rate / 60)
                    
                            Case LOADBYBREAKTHRU
                                ' losses       =  load volume + vent volume / slpm flow
                                '                 *  Load rate / one minute
                                rsTable("LineLoss") = ((StationRecipe(Index, index2).LoadV + StationRecipe(Index, index2).VentV) / Stn_Btn_FlowSP(Index, index2)) _
                                           * (StationRecipe(Index, index2).Load_Rate / 60)
                                           
                            Case LOADBYFID
                                ' losses       =  load volume + vent volume / slpm flow
                                '                 *  Load rate / one minute
                                rsTable("LineLoss") = ((StationRecipe(Index, index2).LoadV + StationRecipe(Index, index2).VentV) / Stn_Btn_FlowSP(Index, index2)) _
                                           * (StationRecipe(Index, index2).Load_Rate / 60)
                                           
                        End Select
                    Else
                        ' no divide by zero
                        rsTable("LineLoss") = 0
                    End If
                    
                End If
    
    '   ******************************
    '   *** FORMULA CHECK REQUIRED ***
    '  should this include UsingCanvent?
    '   ******************************
                If USINGLINEVOLUME And StationControl(Index, index2).CompletedCycles > VALUE0 Then
                    If StationRecipe(Index, index2).Load_Method = LOADBYBREAKTHRU Or StationRecipe(Index, index2).Load_Method = LOADBYFID Then
                        If Stn_Btn_FlowSP(Index, index2) > 0 Then
                            ' losses       =   vent volume / slpm flow
                            '                 *  Load rate / one minute
                            rsTable("LineLoss") = (StationRecipe(Index, index2).VentV / Stn_Btn_FlowSP(Index, index2)) _
                                       * (StationRecipe(Index, index2).Load_Rate / 60)
                        Else
                            ' no divide by zero
                            rsTable("LineLoss") = 0
                        End If
                    End If
                End If
                
            Case Else
                ' nothing
        End Select
    ' **************************************************************************************************************************************************
        
        ' Cumulative Load Totals
        rsTable("LoadRate") = LoadControl(Index, index2).LoadRate
        rsTable("LoadTotalGrams") = LoadControl(Index, index2).loadTotalGrams
        rsTable("LoadTotalLiters") = LoadControl(Index, index2).LoadTotalLiters
        rsTable("TotalWtChg") = LoadControl(Index, index2).TotalWtChg
        rsTable("TotalWtChgRate") = LoadControl(Index, index2).TotalWtChgRate
        
        ' TC's
        If Stn_UseTC(Index, index2) Then
            rsTable("TC1Temp") = Stn_AIO(Index, asStationTC1).EUValue
            rsTable("TC2Temp") = Stn_AIO(Index, asStationTC2).EUValue
        End If
    
        ' Debug Data
        ' Debug Data
        ' Debug Data
        If Not NotDebugMMW Then
          rsTable("TC3Temp") = PreviousReportTimer(Index, index2)
          rsTable("TC4Temp") = DeltaTimer(Index, index2)
          rsTable("TC5Temp") = flag
        End If
        ' Debug Data
        ' Debug Data
        ' Debug Data
        
        If USINGCOMMONTC Then
    '        If Stn_CommonTC(Index, index2) = True Then
                rsTable("CommonTC1") = Com_AIO(acCommonTC1).EUValue
                rsTable("CommonTC2") = Com_AIO(acCommonTC2).EUValue
                rsTable("CommonTC3") = Com_AIO(acCommonTC3).EUValue
                rsTable("CommonTC4") = Com_AIO(acCommonTC4).EUValue
                rsTable("CommonTC5") = Com_AIO(acCommonTC5).EUValue
                rsTable("CommonTC6") = Com_AIO(acCommonTC6).EUValue
    '        End If
        End If
        If USINGLOADPRESSURE Then
            rsTable("LoadPressure") = Stn_AIO(Index, asLoadPressure).EUValue
        End If
        If USINGPURGEOVEN And StationRecipe(Index, index2).PurgeOven Then
            rsTable("PurgeOvenTemp") = Stn_AIO(Index, asPurgeOvenTemp).EUValue
        Else
            rsTable("PurgeOvenTemp") = 0
        End If
    
    
        ' Live Fuel
        If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE)) Then
            rsTable("LiveFuelCycles") = StationControl(Index, 1).LiveFuelCycleCount
            If STN_INFO(Index).ADF_TANKTYPE > 0 Then
                rsTable("LiveFuelTemp") = Stn_AIO(Index, asFuelTankTemp).EUValue
                rsTable("LiveFuelLevel") = Stn_AIO(Index, asFuelTankLevel).EUValue
                rsTable("FuelStorageLevel") = Stn_AIO(Index, asStorageTankLevel).EUValue
                If ((Stn_AIO(Index, asFuelVaporTemp).addr <> 0) Or (Stn_AIO(Index, asFuelVaporTemp).chan <> 0)) Then
                    rsTable("LiveFuelVaporTemp") = Stn_AIO(Index, asFuelVaporTemp).EUValue
                Else
                    rsTable("LiveFuelVaporTemp") = 0
                End If
                If USINGWATERBATH And STN_INFO(Index).ADF_DEF.hasADF_WaterBath Then
                    rsTable("WaterBathTemp") = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
                Else
                    rsTable("WaterBathTemp") = 0
                End If
            Else
                rsTable("LiveFuelTemp") = 0
                rsTable("LiveFuelLevel") = 0
                rsTable("FuelStorageLevel") = 0
                rsTable("LiveFuelVaporTemp") = 0
                rsTable("WaterBathTemp") = 0
            End If
        End If

        ' update LoadDataSet 0
        StnLoadData(Index, index2, 0).ClkTime = rsTable("Time")
        StnLoadData(Index, index2, 0).NitFlow = rsTable("NitFlow")
        StnLoadData(Index, index2, 0).BtnFlow = rsTable("BtnFlow")
        StnLoadData(Index, index2, 0).MixPcnt = rsTable("Mix")
        StnLoadData(Index, index2, 0).LoadRate = rsTable("LoadRate")
        StnLoadData(Index, index2, 0).loadTotalGrams = rsTable("LoadTotalGrams")
        StnLoadData(Index, index2, 0).PriScle = rsTable("priscale")
        StnLoadData(Index, index2, 0).AuxScle = rsTable("auxscale")
        StnLoadData(Index, index2, 0).WtChange = rsTable("TotalWtChg")
        StnLoadData(Index, index2, 0).WtChgRate = rsTable("TotalWtChgRate")
        StnLoadData(Index, index2, 0).LFcycls = rsTable("LiveFuelCycles")
        StnLoadData(Index, index2, 0).FuelTmp = rsTable("LiveFuelTemp")
        StnLoadData(Index, index2, 0).TstTimr = rsTable("TestTime")
        StnLoadData(Index, index2, 0).isBlank = False
        
        ' Flag to indicate that the DB has been updated
        StationControl(Index, index2).NewDataInDB = True
    
        rsTable.Update
        rsTable.Close
        dbDbase.Close
    End If
        
    ' Update Statistics only when MFC's have a nonzero SP
    If LoadControl(Index, index2).Phase = LoadLoading Then Update_Stats Index, index2
              
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

Sub Recipe_Write(ByVal pathDB As String, ByVal thisStnType As Integer, thisRcp As Recipe, ByVal iCourse As Integer)
'
' Procedure Name:   Recipe_Write
' Created by:       Analytical Process Programmer 9/96
' Description:      This routine writes the recipe's data to the
'                   specified database file.
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 7
Dim dbDbase As Database
Dim rsTable As Recordset

    If pathDB <> "" Then
    
        If (thisStnType = STN_LEAKTEST_TYPE) Then
        
            ' LeakTest Station
            ' Write LeakTest Recipe to Database
            Set dbDbase = OpenDatabase(pathDB)
            Set rsTable = dbDbase.OpenRecordset("LeakTestSetup")
            rsTable.AddNew
            
            rsTable("Timeout") = Cfg_LeakTest.timeOut
            rsTable("PressurizeTimeout") = Cfg_LeakTest.PressTimeout
            rsTable("PressureTolerance") = Cfg_LeakTest.PressTol
            rsTable("StablePressDuration") = Cfg_LeakTest.PressTolDuration
            rsTable("DeffTolerance") = Cfg_LeakTest.DeffTol
            rsTable("InitialN2Flow") = Cfg_LeakTest.InitialN2Flow
            rsTable("ReportInterval") = Cfg_LeakTest.ReportInterval
        
            rsTable("SgN2") = SGN2
        
            rsTable("TargetPressure") = Rcp_LeakTest.TargetPress
            rsTable("StableDeffDuration") = Rcp_LeakTest.HoldDuration
        
            rsTable.Update
            rsTable.Close
            dbDbase.Close
               
        Else
        
            ' Normal Station
            ' Write Station Recipe to Database
            Set dbDbase = OpenDatabase(pathDB)
            Set rsTable = dbDbase.OpenRecordset("Recipe")
            rsTable.AddNew
            
            rsTable("Course") = iCourse
            If thisRcp.Name <> "" Then
                rsTable("Name") = thisRcp.Name
            Else
                rsTable("Name") = "unnamed recipe"
            End If
            rsTable("Number") = thisRcp.Number
            
            rsTable("CycleType") = thisRcp.CycleType
            rsTable("CycleTypeDesc") = CycleTypeDesc(thisRcp.CycleType)
                
            rsTable("Load_Method") = thisRcp.Load_MethodSave
            rsTable("Load_MethodDesc") = LoadMethodDesc(thisRcp.Load_MethodSave)
            rsTable("UseHiRangeMFC") = thisRcp.UseHiRangeMFC
            rsTable("UseLoadRatePID") = thisRcp.UseLoadRatePID
            rsTable("NitrogenFlow") = thisRcp.NitrogenFlowSave
            rsTable("Load_Rate") = thisRcp.Load_Rate
            rsTable("Mix_Percent") = thisRcp.Mix_Percent
            rsTable("WC_Mult") = thisRcp.WC_MultSave
            rsTable("EPAFill") = thisRcp.EPAFill
            rsTable("Load_Wt") = thisRcp.Load_Wt
            rsTable("LoadBreakthrough") = thisRcp.LoadBreakthrough
            rsTable("Load_Time") = thisRcp.Load_Time
            
            rsTable("Purge_Time") = thisRcp.Purge_Time
            rsTable("Purge_AuxTime") = thisRcp.Purge_AuxTime
            rsTable("Purge_Method") = thisRcp.Purge_Method
            rsTable("Purge_MethodDesc") = PurgeMethodDesc(thisRcp.Purge_Method)
            rsTable("Purge_Flow") = thisRcp.Purge_Flow
            rsTable("Purge_Can_Vol") = thisRcp.Purge_Can_Vol
            rsTable("Purge_ProfileNumber") = thisRcp.Purge_ProfileNumber
            rsTable("Purge_TargetMode") = thisRcp.Purge_TargetMode
            rsTable("Purge_TargetModeDesc") = PurgeTargetDesc(thisRcp.Purge_TargetMode)
            rsTable("Purge_TargetWeight") = thisRcp.Purge_TargetWeight
            rsTable("Purge_MaxVolumes") = thisRcp.Purge_MaxVolumes
            rsTable("Purge_TargetPurge") = thisRcp.Purge_TargetPurge
            rsTable("Purge_TargetPause") = thisRcp.Purge_TargetPause
            
            rsTable("PurgeInOven") = thisRcp.PurgeOven
            rsTable("PurgeOvenSP") = thisRcp.PurgeOvenSP
            
            rsTable("UseAuxScale") = thisRcp.UseAuxScale
            rsTable("PurgeAuxCan") = thisRcp.PurgeAuxCan
            rsTable("AuxScaleNo") = thisRcp.AuxScaleNo
            rsTable("PauseLeakTime") = thisRcp.PauseLeakTime
            rsTable("PauseLoadTime") = thisRcp.PauseLoadTime
            rsTable("PausePurgeTime") = thisRcp.PausePurgeTime
            rsTable("UsePriScale") = thisRcp.UsePriScale
            rsTable("PriScaleNo") = thisRcp.PriScaleNo
            rsTable("PauseAfterLeak") = thisRcp.PauseAfterLeak
            rsTable("PauseAfterLoad") = thisRcp.PauseAfterLoad
            rsTable("PauseAfterLoadForOper") = thisRcp.PauseAfterLoadForOper
            rsTable("PauseAfterPurge") = thisRcp.PauseAfterPurge
            rsTable("PauseAfterPurgeForOper") = thisRcp.PauseAfterPurgeForOper
            rsTable("LeakCheck") = thisRcp.LeakCheck
            rsTable("LeakPrimary") = thisRcp.LeakPrimary
            rsTable("LeakAux") = thisRcp.LeakAux
            rsTable("MaxLoadTime") = thisRcp.MaxLoadTime
            
            rsTable("IDLoad") = thisRcp.IDLoad
            rsTable("LoadL") = thisRcp.LoadL
            rsTable("LoadV") = thisRcp.LoadV
            rsTable("IDPurge") = thisRcp.IDPurge
            rsTable("PurgeL") = thisRcp.PurgeL
            rsTable("PurgeV") = thisRcp.PurgeV
            rsTable("IDVent") = thisRcp.IDVent
            rsTable("VentL") = thisRcp.VentL
            rsTable("VentV") = thisRcp.VentV
            
            ' start method
            rsTable("StartMethod") = thisRcp.StartMethod
            rsTable("StartDelay") = thisRcp.StartDelay
            rsTable("StartDate") = thisRcp.StartDate
            rsTable("StartMethodDesc") = StartMethodDesc(thisRcp.StartMethod)
                    
            ' end method
            rsTable("EndMethod") = thisRcp.EndMethod
            rsTable("EndMaximumCycles") = thisRcp.EndMaximumCycles
            rsTable("EndMinimumCycles") = thisRcp.EndMinimumCycles
            rsTable("EndConsecutiveCycles") = thisRcp.EndConsecutiveCycles
            rsTable("EndWeightTolerance") = thisRcp.EndWeightTolerance
            rsTable("UpdateCanWc") = thisRcp.UpdateCanWc
            rsTable("Cycles") = thisRcp.CyclesSave
            rsTable("EndMethodDesc") = EndMethodDesc(thisRcp.EndMethod)
                        
            ' aux outputs
            rsTable("AuxOutputs") = thisRcp.AuxOutputs
            rsTable("AuxOutput1_Load") = thisRcp.AuxOutputs_Load(1)
            rsTable("AuxOutput2_Load") = thisRcp.AuxOutputs_Load(2)
            rsTable("AuxOutput3_Load") = thisRcp.AuxOutputs_Load(3)
            rsTable("AuxOutput4_Load") = thisRcp.AuxOutputs_Load(4)
            rsTable("AuxOutput1_Purge") = thisRcp.AuxOutputs_Purge(1)
            rsTable("AuxOutput2_Purge") = thisRcp.AuxOutputs_Purge(2)
            rsTable("AuxOutput3_Purge") = thisRcp.AuxOutputs_Purge(3)
            rsTable("AuxOutput4_Purge") = thisRcp.AuxOutputs_Purge(4)
                
           ' LiveFuel Options
            If (systemhasLIVEFUEL And thisRcp.LiveFuel) Then
                ' For LiveFuel Stations
                rsTable("LiveFuel") = thisRcp.LiveFuel
                rsTable("LiveFuelChgFreq") = thisRcp.LiveFuelChgFreq
                If systemhasAUTODRAINFILL Then
                     ' For Auto Drain/Fill Equipped LiveFuel Tank
                        rsTable("LiveFuelChgAuto") = thisRcp.LiveFuelChgAuto
                     If ((systemhasADF_HEATER) Or (systemhasADF_WATERBATH)) Then
                        ' For LiveFuel Tank with Heater or WaterBath
                        rsTable("ADF_Heater") = thisRcp.ADF_Heater
                        rsTable("ADF_HeaterSP") = thisRcp.ADF_HeaterSP
                     Else
                        ' No LiveFuel Tank with Heater
                        rsTable("ADF_Heater") = False
                        rsTable("ADF_HeaterSP") = 0#
                     End If
                ElseIf (systemhasADF_WATERBATH) Then
                   ' For LiveFuel Tank with Heater
                   rsTable("ADF_Heater") = thisRcp.ADF_Heater
                   rsTable("ADF_HeaterSP") = thisRcp.ADF_HeaterSP
                Else
                     ' No Auto Drain/Fill Equipped LiveFuel Tank
                     rsTable("LiveFuelChgAuto") = False
                     rsTable("ADF_Heater") = False
                     rsTable("ADF_HeaterSP") = 0#
                End If
            Else
                 ' Non LiveFuel Stations
                 rsTable("LiveFuel") = False
                 rsTable("LiveFuelChgFreq") = 0
                 rsTable("LiveFuelChgAuto") = False
                 rsTable("ADF_Heater") = False
                 rsTable("ADF_HeaterSP") = 0#
            End If
               
            rsTable.Update
            rsTable.Close
            dbDbase.Close
               
        End If
    
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

Sub Config_Write(Index As Integer, index2 As Integer)
' Procedure Name:   Config_Write
' Created by:       Brunrose     Sept 2008
' Description:      This routine writes the Configuration data to the data -
'                   base file for the station.
'
Dim dbDbase As Database
Dim rsRecord As Recordset


If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 4321

If StationControl(Index, index2).DBFile <> "" Then
    'Write Config data to data file
    Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
    Set rsRecord = dbDbase.OpenRecordset("Configuration")
    rsRecord.AddNew
      
        rsRecord("UpdateDts") = Now()
        rsRecord("Heading") = StationConfig(Index, index2).Heading
        rsRecord("Heading2") = StationConfig(Index, index2).Heading2
        rsRecord("Next_File") = StationConfig(Index, index2).Next_File
        rsRecord("AutoLogon") = StationConfig(Index, index2).AutoLogon
        rsRecord("AutoLogonUser") = StationConfig(Index, index2).AutoLogonUser
        rsRecord("DbFileBackup_Active") = StationConfig(Index, index2).DbFileBackup_Active
        rsRecord("DbFileBackup_Path") = StationConfig(Index, index2).DbFileBackup_Path
        rsRecord("ReportBackup_Active") = StationConfig(Index, index2).ReportBackup_Active
        rsRecord("ReportBackup_Path") = StationConfig(Index, index2).ReportBackup_Path
        rsRecord("EventRecs") = StationConfig(Index, index2).EventRecs
        rsRecord("JobRecs") = StationConfig(Index, index2).JobRecs
        rsRecord("LCMinDelay") = StationConfig(Index, index2).LCMinDelay
        rsRecord("LCSetPoint") = StationConfig(Index, index2).LCSetPoint
        rsRecord("LCTime") = StationConfig(Index, index2).LCTime
        rsRecord("PressureDecay") = StationConfig(Index, index2).PressureDecay
        rsRecord("LeakCheckFailResponse") = StationConfig(Index, index2).LeakCheckFailResponse
        rsRecord("NitrogenPurgeTime") = StationConfig(Index, index2).NitrogenPurgeTime
        rsRecord("CanVent_Delay_Max") = StationConfig(Index, index2).CanVent_Delay_Max
        rsRecord("OOTtimeDelay") = StationConfig(Index, index2).OOTtimeDelay
        rsRecord("PosPressPurge") = StationConfig(Index, index2).PosPressPurge
        rsRecord("DoorOpenDelay") = StationConfig(Index, index2).DoorOpenDelay
        rsRecord("UPSOpenDelay") = StationConfig(Index, index2).UPSOpenDelay
        rsRecord("LoadPressure") = StationConfig(Index, index2).LoadPressure
        rsRecord("ButaneMassLimit") = StationConfig(Index, index2).ButaneMassLimit
        rsRecord("LoadTimeLimit") = StationConfig(Index, index2).LoadTimeLimit
        rsRecord("WaterBathControl") = StationConfig(Index, index2).WaterBathControl
        rsRecord("LeakCheck_Interval") = StationConfig(Index, index2).LeakCheck_Interval
        rsRecord("LeakTotal_Interval") = StationConfig(Index, index2).LeakTotal_Interval
        rsRecord("Load_Interval") = StationConfig(Index, index2).Load_Interval
        rsRecord("Purge_Interval") = StationConfig(Index, index2).Purge_Interval
        rsRecord("LoadTotal_Interval") = StationConfig(Index, index2).LoadTotal_Interval
        rsRecord("PurgeTotal_Interval") = StationConfig(Index, index2).PurgeTotal_Interval
        rsRecord("Tol_Nit_Flow") = StationConfig(Index, index2).Tol_Nit_Flow
        rsRecord("Tol_Btn_Flow") = StationConfig(Index, index2).Tol_Btn_Flow
        rsRecord("Tol_ORVRNit_Flow") = StationConfig(Index, index2).Tol_ORVRNit_Flow
        rsRecord("Tol_ORVRBtn_Flow") = StationConfig(Index, index2).Tol_ORVRBtn_Flow
        rsRecord("Tol_Pur_Flow") = StationConfig(Index, index2).Tol_Pur_Flow
        rsRecord("Tol_Lfv_Flow") = StationConfig(Index, index2).Tol_Lfv_Flow
        rsRecord("Tol_Mix_Ratio") = StationConfig(Index, index2).Tol_Mix_Ratio
        rsRecord("Tol_Temp") = StationConfig(Index, index2).Tol_Temp
        rsRecord("Tol_Moisture") = StationConfig(Index, index2).Tol_Moisture
        rsRecord("Tol_FuelTemp") = StationConfig(Index, index2).Tol_FuelTemp
        rsRecord("Tol_Purge_Total") = StationConfig(Index, index2).Tol_Purge_Total
        rsRecord("Tol_Load_Total") = StationConfig(Index, index2).Tol_Load_Total
        rsRecord("Tol_PurgeOvenTemp") = StationConfig(Index, index2).Tol_PurgeOvenTemp
        rsRecord("Tol_WaterBathTemp") = StationConfig(Index, index2).Tol_WaterBathTemp
        rsRecord("PurgeDP_HiLimit") = StationConfig(Index, index2).PurgeDP_HiLimit
        rsRecord("LoLim_Load_Flow") = StationConfig(Index, index2).LoLim_Load_Flow
        rsRecord("LoLim_Purge_Flow") = StationConfig(Index, index2).LoLim_Purge_Flow
        rsRecord("Temp_Target") = StationConfig(Index, index2).Temp_Target
        rsRecord("Moisture_Target") = StationConfig(Index, index2).Moisture_Target
        rsRecord("LoadSettleTime") = StationConfig(Index, index2).LoadSettleTime
        rsRecord("PurgeSettleTime") = StationConfig(Index, index2).PurgeSettleTime
        rsRecord("ReportFileName1stPart") = StationConfig(Index, index2).ReportFileName1stPart
        rsRecord("ReportFileName2ndPart") = StationConfig(Index, index2).ReportFileName2ndPart
        rsRecord("ReportFileName3rdPart") = StationConfig(Index, index2).ReportFileName3rdPart
    
        rsRecord("CsvEotReporting") = StationConfig(Index, index2).RptConfig.CsvEotReporting
        rsRecord("CsvEotSummary") = StationConfig(Index, index2).RptConfig.CsvEotSummary
        rsRecord("CsvEotDetail") = StationConfig(Index, index2).RptConfig.CsvEotDetail
        rsRecord("CsvGenReporting") = StationConfig(Index, index2).RptConfig.CsvGenReporting
        rsRecord("CsvGenSummary") = StationConfig(Index, index2).RptConfig.CsvGenSummary
        rsRecord("CsvGenDetail") = StationConfig(Index, index2).RptConfig.CsvGenDetail
        rsRecord("TextEotReporting") = StationConfig(Index, index2).RptConfig.TextEotReporting
        rsRecord("TextEotSummary") = StationConfig(Index, index2).RptConfig.TextEotSummary
        rsRecord("TextEotSummary_AutoPrint") = StationConfig(Index, index2).RptConfig.TextEotSummary_AutoPrint
        rsRecord("TextEotDetail") = StationConfig(Index, index2).RptConfig.TextEotDetail
        rsRecord("TextGenReporting") = StationConfig(Index, index2).RptConfig.TextGenReporting
        rsRecord("TextGenSummary") = StationConfig(Index, index2).RptConfig.TextGenSummary
        rsRecord("TextGenDetail") = StationConfig(Index, index2).RptConfig.TextGenDetail
        rsRecord("XlsEotReporting") = StationConfig(Index, index2).RptConfig.XlsEotReporting
        rsRecord("XlsEotSummary") = StationConfig(Index, index2).RptConfig.XlsEotSummary
        rsRecord("XlsEotDetail") = StationConfig(Index, index2).RptConfig.XlsEotDetail
        rsRecord("XlsGenReporting") = StationConfig(Index, index2).RptConfig.XlsGenReporting
        rsRecord("XlsGenSummary") = StationConfig(Index, index2).RptConfig.XlsGenSummary
        rsRecord("XlsGenDetail") = StationConfig(Index, index2).RptConfig.XlsGenDetail
    
        rsRecord("BtnFlowResponse") = StationConfig(Index, index2).BtnFlowResp
        rsRecord("NitFlowResponse") = StationConfig(Index, index2).NitFlowResp
        rsRecord("FuelLevelResponse") = StationConfig(Index, index2).FuelLevelResp
        rsRecord("FuelTempResponse") = StationConfig(Index, index2).FuelTempResp
        rsRecord("PurFlowResponse") = StationConfig(Index, index2).PurFlowResp
        rsRecord("AirMoistResponse") = StationConfig(Index, index2).AirMoistResp
        rsRecord("AirTempResponse") = StationConfig(Index, index2).AirTempResp
        rsRecord("CanVentResponse") = StationConfig(Index, index2).CanVentResp
        rsRecord("LoadRateResponse") = StationConfig(Index, index2).LoadRateResp
        rsRecord("PurgeDpResponse") = StationConfig(Index, index2).PurgeDpResp
        rsRecord("PurgeOvenResponse") = StationConfig(Index, index2).PurgeOvenResp
        rsRecord("WaterBathResponse") = StationConfig(Index, index2).WaterBathResp
    
    rsRecord.Update
    rsRecord.Close
    
    'Write Sysdef data to job data file
    Set rsRecord = dbDbase.OpenRecordset("SystemDefinition")
    rsRecord.AddNew
        rsRecord("UsingC") = USINGC
        rsRecord("UsingF") = USINGF
        rsRecord("UsingMoist_RH") = USINGMoist_RH
        rsRecord("UsingMoist_Grains") = USINGMoist_Grains
        rsRecord("UsingLV_English") = USINGLVol_Engl
        rsRecord("UsingLV_SI") = USINGLVol_SI
    rsRecord.Update
    rsRecord.Close
    
    
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

Sub Sysdef_Write(iStation As Integer, iShift As Integer)
' Procedure Name:   Sysdef_Write
' Created by:       Brunrose
' Description:      This routine writes the Sysdef data to the
'                   job database file for the station.
'
Dim dbDbase As Database
Dim rsTable As Recordset
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 7471

If StationControl(iStation, iShift).DBFile <> "" Then
    'Write Sysdef data to job database
    Set dbDbase = OpenDatabase(StationControl(iStation, iShift).DBFile)
    Set rsTable = dbDbase.OpenRecordset("SystemDefinition")
    rsTable.AddNew
    
        rsTable("UsingC") = SysSysDef.USINGC
        rsTable("UsingF") = SysSysDef.USINGF
        rsTable("UsingMoist_RH") = SysSysDef.USINGMoist_RH
        rsTable("UsingMoist_Grains") = SysSysDef.USINGMoist_Grains
        rsTable("UsingLV_English") = SysSysDef.USINGLVol_Engl
        rsTable("UsingLV_SI") = SysSysDef.USINGLVol_SI
      
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

Sub Sequence_Write(iStation As Integer, iShift As Integer)
' Procedure Name:   Sequence_Write
' Created by:       Brunrose
' Description:      This routine writes the Sequence data to the
'                   job database file for the station.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 1921

Dim dbDbase As Database
Dim rsTable As Recordset
Dim iCourse As Integer
Dim saveNumCourses As Integer
Dim saveCourseData As JobSequenceCourse

    If StationControl(iStation, iShift).DBFile <> "" Then
    
        ' Make adjustments for LeakTest Station
        If (STN_INFO(iStation).Type = STN_LEAKTEST_TYPE) Then
            saveNumCourses = StationSequence(iStation, iShift).NumCourses
            saveCourseData = StationSequence(iStation, iShift).CourseData(1)
            StationSequence(iStation, iShift).NumCourses = 1
            StationSequence(iStation, iShift).CourseData(1).RecipeNumber = 0
            StationSequence(iStation, iShift).CourseData(1).Type = courseRecipe
            StationSequence(iStation, iShift).CourseData(1).Cycles = 1
            StationSequence(iStation, iShift).CourseData(1).CourseNumber = 1
            StationSequence(iStation, iShift).CourseData(1).PauseDuration = 0
            StationSequence(iStation, iShift).CourseData(1).LoadRate = 0
            StationSequence(iStation, iShift).CourseData(1).PurgeRate = 0
            StationSequence(iStation, iShift).CourseData(1).EstCourseDuration = CSng(Rcp_LeakTest.HoldDuration + 15) / 60#
            StationSequence(iStation, iShift).CourseData(1).MsgText = "LeakTest"
        End If
    
        ' Write Sequence data to job database
        Set dbDbase = OpenDatabase(StationControl(iStation, iShift).DBFile)
        Set rsTable = dbDbase.OpenRecordset("Sequence")
        rsTable.AddNew
        
            rsTable("Number") = 0
            rsTable("Description") = StationSequence(iStation, iShift).Description
            rsTable("Courses") = StationSequence(iStation, iShift).NumCourses
            rsTable("PriScale") = StationSequence(iStation, iShift).PriScaleNo
            rsTable("AuxScale") = StationSequence(iStation, iShift).AuxScaleNo
            rsTable("IDLoad") = StationSequence(iStation, iShift).IDLoad
            rsTable("IDPurge") = StationSequence(iStation, iShift).IDPurge
            rsTable("IDVent") = StationSequence(iStation, iShift).IDVent
            rsTable("LoadL") = StationSequence(iStation, iShift).LoadL
            rsTable("LoadV") = StationSequence(iStation, iShift).LoadV
            rsTable("PurgeL") = StationSequence(iStation, iShift).PurgeL
            rsTable("PurgeV") = StationSequence(iStation, iShift).PurgeV
            rsTable("VentL") = StationSequence(iStation, iShift).VentL
            rsTable("VentV") = StationSequence(iStation, iShift).VentV
            rsTable("Validated") = StationSequence(iStation, iShift).Validated
            rsTable("EstSeqDuration") = StationSequence(iStation, iShift).EstSeqDuration
            rsTable("EstSeqDurDesc") = StationSequence(iStation, iShift).EstSeqDurDesc
                
        rsTable.Update
        rsTable.Close
               
        
        ' Write Sequence Course data to job database
        Set dbDbase = OpenDatabase(StationControl(iStation, iShift).DBFile)
        Set rsTable = dbDbase.OpenRecordset("SequenceCourses")
        If rsTable.BOF Then
            iCourse = 1
            Do While (iCourse <= StationSequence(iStation, iShift).NumCourses)
                rsTable.AddNew
                rsTable("CourseNumber") = StationSequence(iStation, iShift).CourseData(iCourse).CourseNumber
                rsTable("Type") = StationSequence(iStation, iShift).CourseData(iCourse).Type
                rsTable("PauseDuration") = StationSequence(iStation, iShift).CourseData(iCourse).PauseDuration
                rsTable("RecipeNumber") = StationSequence(iStation, iShift).CourseData(iCourse).RecipeNumber
                rsTable("Cycles") = StationSequence(iStation, iShift).CourseData(iCourse).Cycles
                rsTable("LoadRate") = StationSequence(iStation, iShift).CourseData(iCourse).LoadRate
                rsTable("PurgeRate") = StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate
                rsTable("EstCourseDuration") = StationSequence(iStation, iShift).CourseData(iCourse).EstCourseDuration
                rsTable("MsgText") = StationSequence(iStation, iShift).CourseData(iCourse).MsgText
                rsTable.Update
                iCourse = iCourse + 1
            Loop
        End If
                   
        rsTable.Close
        dbDbase.Close
        
        ' Restore adjustments for LeakTest Station
        If (STN_INFO(iStation).Type = STN_LEAKTEST_TYPE) Then
            StationSequence(iStation, iShift).NumCourses = saveNumCourses
            StationSequence(iStation, iShift).CourseData(1) = saveCourseData
        End If

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

Sub Weights_Write(iStation As Integer, iShift As Integer)
' Procedure Name:   Weights_Write
' Created by:       Brunrose
' Description:      This routine writes the Cycle Weights data to the
'                   job database file for the station.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 747
Dim dbDbase As Database
Dim rsTable As Recordset
Dim iCycle As Integer

    If StationControl(iStation, iShift).DBFile <> "" Then
        'Write Cycle Weights data to job database
        Set dbDbase = OpenDatabase(StationControl(iStation, iShift).DBFile)
        Set rsTable = dbDbase.OpenRecordset("CycleWeights")
        ' for all actual cycles
        For iCycle = 1 To StationControl(iStation, iShift).CurrCycle
            rsTable.AddNew
                rsTable("Course") = StationControl(iStation, iShift).Course
                rsTable("Cycle") = iCycle
                rsTable("Cycle_StartWeight_Total") = StationCycleWeightData(iStation, iShift, iCycle).Cycle_StartWeight_Total
                rsTable("Cycle_EndWeight_Total") = StationCycleWeightData(iStation, iShift, iCycle).Cycle_EndWeight_Total
                rsTable("Load_StartWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Load_StartWeight_Aux
                rsTable("Load_EndWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Load_EndWeight_Aux
                rsTable("Load_StartWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Load_StartWeight_Pri
                rsTable("Load_EndWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Load_EndWeight_Pri
                rsTable("Purge_StartWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Purge_StartWeight_Aux
                rsTable("Purge_EndWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Purge_EndWeight_Aux
                rsTable("Purge_StartWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Purge_StartWeight_Pri
                rsTable("Purge_EndWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Purge_EndWeight_Pri
                rsTable("Load_TotalGrams") = StationCycleWeightData(iStation, iShift, iCycle).Load_TotalGrams
                rsTable("LoadPause_StartWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Load_StartWeight_Aux
                rsTable("LoadPause_EndWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Load_EndWeight_Aux
                rsTable("LoadPause_StartWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Load_StartWeight_Pri
                rsTable("LoadPause_EndWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Load_EndWeight_Pri
                rsTable("PurgePause_StartWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Purge_StartWeight_Aux
                rsTable("PurgePause_EndWeight_Aux") = StationCycleWeightData(iStation, iShift, iCycle).Purge_EndWeight_Aux
                rsTable("PurgePause_StartWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Purge_StartWeight_Pri
                rsTable("PurgePause_EndWeight_Pri") = StationCycleWeightData(iStation, iShift, iCycle).Purge_EndWeight_Pri
            rsTable.Update
        Next iCycle
        
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

Private Function RecentWtChgRate(ByVal iStn As Integer, ByVal iShift As Integer) As Single
' update the Recent-Weight-Change-Rate calculations
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 9696
Dim newWt As Single
Dim newTime As Single
Dim oldWt As Single
Dim oldTime As Single
Dim deltaWt As Single
Dim deltatime As Single
Dim tmpRate As Single
Dim Idx As Integer

    ' current "load new value" index
    Idx = InIdx(iStn, iShift)
    ' new weight & time
    newWt = 0
    If StationRecipe(iStn, iShift).UseAuxScale Then newWt = newWt + StationControl(iStn, iShift).AuxScaleWt
    If StationRecipe(iStn, iShift).UsePriScale Then newWt = newWt + StationControl(iStn, iShift).PriScaleWt
    newTime = StationControl(iStn, iShift).TestTimer
    ' old weight & time
    oldWt = WtQueue(Idx, iStn, iShift)
    oldTime = TimeQueue(Idx, iStn, iShift)
    
    ' calc recent weight change
    deltaWt = newWt - oldWt
    ' calc time change
    deltatime = newTime - oldTime
    
    ' calc Recent Weight Change Rate
    If (deltatime <> 0) Then
        tmpRate = CSng(3600) * (deltaWt / deltatime)
    Else
        tmpRate = CSng(0)
    End If
    
    ' add new weight & time to the Queues
    WtQueue(Idx, iStn, iShift) = newWt
    TimeQueue(Idx, iStn, iShift) = newTime
    ' increment the "load new value" index
    InIdx(iStn, iShift) = IIf((InIdx(iStn, iShift) < WTCHGQUEUESIZE), (InIdx(iStn, iShift) + 1), 1)
    
    RecentWtChgRate = tmpRate
          
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    RecentWtChgRate = CSng(0)
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

'=====================================================================================
'COMPACT AND REPAIR DATABASE
'=====================================================================================
'    Call CompactRepairAccessDB(DatabasePath, FileName)
Public Sub CompactRepairAccessDB(ByVal sDBPATH As String, ByVal sDBFILE As String)
Dim sDB As String
Dim sDBtmp As String

    sDB = sDBPATH & sDBFILE
    sDBtmp = sDBPATH & "tmp" & sDBFILE

    'Call the statement to execute compact and repair...
    Call DBEngine.CompactDatabase(sDB, sDBtmp)
    'wait for the app to finish
    DoEvents
    'remove the uncompressed original
    Kill sDB
    'rename the compressed file to the original to restore for other functions
    Name sDBtmp As sDB
End Sub

Sub RemStatus_Write(Index As Integer, index2 As Integer, flag As Integer)
' flag = 0 normal write; flag = 1 unused; flag = 2 unused
'
' Function Name:    RemStatus_Write
' Author:           Analytical Process Programmer     9/9/09
' Description:      Updates the RemoteDB UnitStatus table
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 3, 13643

Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String
Dim temp As Single

' Using the open DB File
If Len(StationControl(Index, index2).DBFile) > 0 Then

    ' which type of Default_Write is this?
    Select Case flag
        Case NORMALUPDATE
    
        Case Else
            
    End Select
        
  
  ' update reports for running stations that are not purging, loading, or leak checking
  Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
  Set rsTable = dbDbase.OpenRecordset("data")
  rsTable.AddNew
  
    ' GENERAL
    rsTable("Course") = StationControl(Index, index2).Course
    rsTable("Mode") = StationControl(Index, index2).Mode
    rsTable("Phase") = 0
    rsTable("ModeDesc") = ModeDescShort(StationControl(Index, index2).Mode)
    rsTable("Time") = Now
    rsTable("TestTime") = StationControl(Index, index2).TestTimer
    rsTable("Cycle") = StationControl(Index, index2).CurrCycle
    rsTable("Actual") = 0
    
    ' AIR
    rsTable("PATemp") = PATemp
    rsTable("PARH") = PAHum
    rsTable("Moisture") = PAMoisture
    rsTable("Baro") = AmbBaro
    
    ' FLOWS
    rsTable("PurgeFlow") = Stn_AIO(Index, asPurgeAirFlow).EUValue
    rsTable("NitFlow") = Stn_Nit_Flow_PV(Index, index2)
    rsTable("BtnFlow") = Stn_Btn_Flow_PV(Index, index2)
    
    ' SCALES
    If StationRecipe(Index, index2).UseAuxScale = True Then
       rsTable("AuxScale") = StationControl(Index, index2).AuxScaleWt
    Else
       rsTable("AuxScale") = 0
    End If
    If StationRecipe(Index, index2).UsePriScale = True Then
       rsTable("PriScale") = StationControl(Index, index2).PriScaleWt
    Else
       rsTable("PriScale") = 0
    End If
    
'    ' STATION TC's
    If Stn_UseTC(Index, index2) Then
      rsTable("TC1Temp") = Stn_AIO(Index, asStationTC1).EUValue
      rsTable("TC2Temp") = Stn_AIO(Index, asStationTC2).EUValue
    End If
    
    ' COMMON TC's
    If USINGCOMMONTC Then
'       If Stn_CommonTC(Index, index2) = True Then
            rsTable("CommonTC1") = Com_AIO(acCommonTC1).EUValue
            rsTable("CommonTC2") = Com_AIO(acCommonTC2).EUValue
            rsTable("CommonTC3") = Com_AIO(acCommonTC3).EUValue
            rsTable("CommonTC4") = Com_AIO(acCommonTC4).EUValue
            rsTable("CommonTC5") = Com_AIO(acCommonTC5).EUValue
            rsTable("CommonTC6") = Com_AIO(acCommonTC6).EUValue
'       End If
    End If

    If USINGPURGEOVEN And StationRecipe(Index, index2).PurgeOven Then
        rsTable("PurgeOvenTemp") = Stn_AIO(Index, asPurgeOvenTemp).EUValue
    Else
        rsTable("PurgeOvenTemp") = 0
    End If


    ' Live Fuel
    If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE)) Then
        rsTable("LiveFuelCycles") = StationControl(Index, 1).LiveFuelCycleCount
        If STN_INFO(Index).ADF_TANKTYPE > 0 Then
            rsTable("LiveFuelTemp") = Stn_AIO(Index, asFuelTankTemp).EUValue
            rsTable("LiveFuelLevel") = Stn_AIO(Index, asFuelTankLevel).EUValue
            rsTable("FuelStorageLevel") = Stn_AIO(Index, asStorageTankLevel).EUValue
            If ((Stn_AIO(Index, asFuelVaporTemp).addr <> 0) Or (Stn_AIO(Index, asFuelVaporTemp).chan <> 0)) Then
                rsTable("LiveFuelVaporTemp") = Stn_AIO(Index, asFuelVaporTemp).EUValue
            Else
                rsTable("LiveFuelVaporTemp") = 0
            End If
            If USINGWATERBATH And STN_INFO(Index).ADF_DEF.hasADF_WaterBath Then
                rsTable("WaterBathTemp") = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
            Else
                rsTable("WaterBathTemp") = 0
            End If
        Else
            rsTable("LiveFuelTemp") = 0
            rsTable("LiveFuelLevel") = 0
            rsTable("FuelStorageLevel") = 0
            rsTable("LiveFuelVaporTemp") = 0
            rsTable("WaterBathTemp") = 0
        End If
    End If

  ' Flag to indicate that the DB has been updated
  StationControl(Index, index2).NewDataInDB = True

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


