Attribute VB_Name = "Module2"
' error module 2 ''''''''''''''''''' program TestMgr.bas '''''''''''''''''''''
Option Explicit
'
Private scaleFlag(1 To MAX_STN) As Boolean
Private LoadMinDuration(1 To MAX_STN) As Long   ' minimum Load Duration in seconds

Public Sub Check_Stations()
' Procedure name:   Check_Stations
' Author:           Analytical Process Programmer 7/25/96
' Description:      Check the current status of all stations, sets up
'                   data files, writes report data, shuts down data files
'                   when complete.
'
Dim iStation As Integer
Dim iShift As Integer
Dim tempMode As Integer
Dim shiftcount As Integer
Dim idleshiftcount As Integer
Dim sourcefile As String
Dim destfile As String
Dim strExists As String
Dim MeasuredFlowRate As Single
Dim inc As Integer
Dim Idx As Integer
Dim Nitrogen_Output As Single
Dim StnIsActive As Boolean
Dim ootFlag As Boolean
Dim ootResp As Integer
Dim startTest As Boolean
Dim activeShift As Integer
Dim activeness As Integer
Dim maxactive As Integer
Dim slogmsg As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 20


    ' Common Purge-With-Dry-Air valves
    If (AllStationsIdle And USINGDRYPURGEAIR And SysConfig.DryAirPurge) Then
        Com_OutDigital icPurgeDryAirSupplySol, cOFF
        Com_OutDigital icPurgeAirSourceSelectSol, cOFF
    End If
    
    
    ' Check Status (and respond appropriately) of All Shifts
    If CurChkStn < 1 Then CurChkStn = 1
    iStation = CurChkStn
    SystemTimers(tmrStnLogic).Phase = iStation
    
    ' Is this Station Active, etc ?
    ChgErrModule 2, 2010
    StnIsActive = False
    activeShift = 1     ' default active shift
    maxactive = 0
    For iShift = 1 To NR_SHIFT
        
        ' DETERMINE IF STA IS IDLE_DEBOUNCED
        If StationControl(iStation, iShift).Mode = VBIDLE Then
            If Not StationControl(iStation, iShift).ModeIsIdle_Debounced Then
                StationControl(iStation, iShift).ModeIsIdle_DebounceCount = StationControl(iStation, iShift).ModeIsIdle_DebounceCount + 1
                If StationControl(iStation, iShift).ModeIsIdle_DebounceCount > 3 Then
                    StationControl(iStation, iShift).ModeIsIdle_Debounced = True
                End If
            End If
        Else
            StationControl(iStation, iShift).ModeIsIdle_DebounceCount = 0
            StationControl(iStation, iShift).ModeIsIdle_Debounced = False
        End If
        
        ' DETERMINE IF STATION IS ACTIVE
        If StationControl(iStation, iShift).Mode = VBIDLE _
            Or StationControl(iStation, iShift).Mode = VBIDLEWAITING _
            Or StationControl(iStation, iShift).Mode = VBCOMPLETE _
            Or StationControl(iStation, iShift).Mode = VBSTARTWAIT _
            Or StationControl(iStation, iShift).Mode = VBSCALEWAIT _
            Or StationControl(iStation, iShift).Mode = VBSHIFTWAIT _
            Or StationControl(iStation, iShift).Mode = VBPAUSEALARM _
            Or StationControl(iStation, iShift).Mode = VBPAUSEOOT _
            Or StationControl(iStation, iShift).Mode = VBFIDPAUSE _
            Or StationControl(iStation, iShift).Mode = VBWBPAUSE _
            Or StationControl(iStation, iShift).Mode = VBGASPAUSE Then
            ' Station/Shift is Not Active; OK to Use Station IO Forcing
        Else
            ' Station/Shift Is Active; No Station IO Forcing Allowed
            StnIsActive = True
            ' Continue to Request Aspirator Standby
            PRG_INFO(STN_INFO(iStation).AspiratorNum).StandbyRequest = True
        End If
        
        ' DETERMINE ACTIVE SHIFT
        Select Case StationControl(iStation, iShift).Mode
            Case VBIDLE
                activeness = 0
            Case VBSTARTWAIT, VBCOURSEWAIT, VBCOURSEPAUSE
                activeness = 1
            Case VBSCALEWAIT, VBSHIFTWAIT
                activeness = 2
            Case VBIDLEWAITING, VBCOMPLETE
                activeness = 4
            Case VBPAUSEOOT, VBLEAKERROR
                activeness = 6
            Case VBPAUSEALARM
                Select Case StationControl(iStation, iShift).Mode_PauseSave
                    Case VBIDLE
                        activeness = 0
                    Case VBSTARTWAIT
                        activeness = 1
                    Case VBSCALEWAIT, VBSHIFTWAIT
                        activeness = 2
                    Case VBIDLEWAITING, VBCOMPLETE
                        activeness = 4
                    Case VBPAUSEOOT, VBLEAKERROR
                        activeness = 6
                    Case Else
                        activeness = 9
                End Select
            Case Else
                activeness = 9
        End Select
        If activeness > maxactive Then
            ' this shift is the most active (so far)
            maxactive = activeness
            activeShift = iShift
        End If
    
    Next iShift
    
    ' Set the Active Shift for this Station
    Stn_ActiveShift(iStation) = activeShift
    
    ' check for (& clear) invalid lock on Pressure Transducer
    If LeakCheckControl.station = iStation Then
        If StationControl(iStation, Stn_ActiveShift(iStation)).ModeIsIdle_Debounced _
                Or (StationControl(iStation, Stn_ActiveShift(iStation)).Mode = VBLOAD _
                Or StationControl(iStation, Stn_ActiveShift(iStation)).Mode = VBPURGE) Then
            LeakCheckControl.station = 0
            LeakCheckControl.Shift = 0
            LeakCheckControl.Phase = 0
        End If
    End If
    
    ' Set Mode for IO Forcing
    If StnIsActive Then
        STN_IOForceMode(iStation) = VBAUTO
        IdlePauseCount = IdlePauseCount + 1
    Else
        STN_IOForceMode(iStation) = VBMANUAL
    End If
    
    
    ChgErrModule 2, 2012
    ' ***********************************************
    ' Update Pause & Idle Lights only if no IOForcing
    ' ***********************************************
    
    ' Main Pause LT
    If Not IOForceActive Or STN_IOForceMode(0) = VBAUTO Then
        If Pause_Alarm = SYSTEMPAUSED Then
            Com_OutDigital icPauseLT, cON
        Else
            Com_OutDigital icPauseLT, cOFF
        End If
    End If
    
    ' Station LT's
    If Not IOForceActive Then
    
        ' Look for most active mode on any shift
        tempMode = VBIDLE
        If NR_SHIFT = 1 Then
            ' Only One Shift; Use Shift 1's mode
            tempMode = StationControl(iStation, 1).Mode
        ElseIf StationControl(iStation, 1).Mode = VBPAUSEALARM Then
            ' Shift 1 Is Paused; Use Shift 1's mode
            tempMode = StationControl(iStation, 1).Mode
        ElseIf StationControl(iStation, 2).Mode = VBPAUSEALARM Then
            ' Shift 2 Is Paused; Use Shift 2's mode
            tempMode = StationControl(iStation, 2).Mode
        ElseIf (StationControl(iStation, 2).Mode = VBIDLE Or StationControl(iStation, 2).Mode = VBSTARTWAIT _
            Or StationControl(iStation, 2).Mode = VBSCALEWAIT Or StationControl(iStation, 2).Mode = VBSHIFTWAIT) Then
            ' Shift 2 is Idle or Waiting to Start; Use Shift 1's mode
            tempMode = StationControl(iStation, 1).Mode
        ElseIf (StationControl(iStation, 1).Mode = VBIDLE Or StationControl(iStation, 2).Mode = VBSTARTWAIT _
            Or StationControl(iStation, 1).Mode = VBSCALEWAIT Or StationControl(iStation, 1).Mode = VBSHIFTWAIT) Then
            ' Shift 1 Is Idle or Waiting to Start; Use Shift 2's mode
            tempMode = StationControl(iStation, 2).Mode
        Else
            ' There are 2 Shifts; Neither Is Paused Or Idle Or Waiting to Start; Use Shift 1's mode
            '   (this shouldn't be possible; but just to be safe)
            tempMode = StationControl(iStation, 1).Mode
        End If
        
    
        ' Station Pause LT
        Select Case tempMode
            Case VBPURGEWAIT, VBSTARTWAIT, VBSCALEWAIT, VBSHIFTWAIT, VBLEAKWAIT, VBPOSTPURGE, VBPOSTLOAD, VBPOSTLEAK, VBCOURSEWAIT, VBCOURSEPAUSE
                ' Blink the Station Pause LT
                If Not Stn_DIO(iStation, isPauseLT).Value Then
                    Stn_OutDigital iStation, isPauseLT, cON
                ElseIf Stn_DIO(iStation, isPauseLT).Value Then
                    Stn_OutDigital iStation, isPauseLT, cOFF
                End If
            Case VBPAUSEALARM, VBPAUSEOOT, VBGASPAUSE, VBWBPAUSE, VBFIDPAUSE, VBLEAKERROR
                ' Turn On the Station Pause LT
                If Not Stn_DIO(iStation, isPauseLT).Value Then Stn_OutDigital iStation, isPauseLT, cON
            Case Else
                ' Turn Off the Station Pause LT
                If Stn_DIO(iStation, isPauseLT).Value Then Stn_OutDigital iStation, isPauseLT, cOFF
        End Select
    
        ' Station Idle LT
        If StationControl(iStation, 1).ModeIsIdle_Debounced And (NR_SHIFT = 1 Or StationControl(iStation, 2).ModeIsIdle_Debounced) Then
            If Not Stn_DIO(iStation, isIdleLT).Value Then Stn_OutDigital iStation, isIdleLT, cON
        Else
            If Stn_DIO(iStation, isIdleLT).Value Then Stn_OutDigital iStation, isIdleLT, cOFF
        End If
        
        
    End If
    ' ***********************************************
    ' End of Pause & Idle Lights
    ' ***********************************************
        
        
    For iShift = 1 To NR_SHIFT
                
        ' Need to Stop Station ?
        If StationControl(iStation, iShift).StopRequest Then
            ' Stop Button Pressed
            StationControl(iStation, iShift).StopRequest = False
            Station_Abort iStation, iShift, OPER_STOP                      ' Stop station & shift
        ElseIf StationControl(iStation, iShift).AbortRequest Then
            ' Stopping Station
            StationControl(iStation, iShift).AbortRequest = False
            Station_Abort iStation, iShift, EXIT_STOP                      ' Abort station & shift
        Else
        
            ' Check for OOT Pause or OOT Stop
            ChgErrModule 2, 2011
            If (USINGOOTPAUSE And (Pause_Alarm = NOTPAUSED)) Then
                ' don't check for station OOT if system paused
                If StationControl(iStation, iShift).Mode <> VBPAUSEOOT And StationControl(iStation, iShift).Mode <> VBIDLE _
                    And StationControl(iStation, iShift).Mode <> VBIDLEWAITING And StationControl(iStation, iShift).Mode <> VBCOMPLETE Then
                    
                    ' check for OOTs
                    For Idx = 1 To ootStorageLevel
                    
                        ' which OOT are we looking at ?
                        Select Case Idx
                            Case ootBtnFlow
                                ootFlag = OOTs(iStation, iShift).BtnFlowOOT
                                ootResp = StationConfig(iStation, iShift).BtnFlowResp
                            Case ootNitFlow
                                ootFlag = OOTs(iStation, iShift).NitFlowOOT
                                ootResp = StationConfig(iStation, iShift).BtnFlowResp
                            Case ootFuelTemp
                                ootFlag = OOTs(iStation, iShift).FuelTempOOT
                                ootResp = StationConfig(iStation, iShift).FuelTempResp
                            Case ootPurFlow
                                ootFlag = OOTs(iStation, iShift).PurFlowOOT
                                ootResp = StationConfig(iStation, iShift).PurFlowResp
                            Case ootAirMoist
                                ootFlag = IIf(SysConfig.PosPressPurge, False, OOTs(iStation, iShift).AirMoistOOT)
                                ootResp = StationConfig(iStation, iShift).AirMoistResp
                            Case ootAirTemp
                                ootFlag = IIf(SysConfig.PosPressPurge, False, OOTs(iStation, iShift).AirTempOOT)
                                ootResp = StationConfig(iStation, iShift).AirTempResp
                            Case ootCanVent
                                ootFlag = OOTs(iStation, iShift).CanVentOOT
                                ootResp = StationConfig(iStation, iShift).CanVentResp
                            Case ootLoadRate
                                ootFlag = OOTs(iStation, iShift).LoadRateOOT
                                ootResp = StationConfig(iStation, iShift).LoadRateResp
                            Case ootPurgeDp
                                ootFlag = OOTs(iStation, iShift).PurgeDpOOT
                                ootResp = StationConfig(iStation, iShift).PurgeDpResp
                            Case ootFuelLevel
                                ootFlag = OOTs(iStation, iShift).FuelLevelOOT
                                ootResp = StationConfig(iStation, iShift).FuelLevelResp
                            Case ootStorageLevel
                                ootFlag = OOTs(iStation, iShift).StorageLevelOOT
                                ootResp = StationConfig(iStation, iShift).StorageLevelResp
                            Case ootPurgeOvenTemp
                                ootFlag = OOTs(iStation, iShift).PurgeOvenOOT
                                ootResp = StationConfig(iStation, iShift).PurgeDpResp
                            Case ootWaterBathTemp
                                ootFlag = OOTs(iStation, iShift).WaterBathOOT
                                ootResp = StationConfig(iStation, iShift).PurgeDpResp
                        End Select
                        
                        ' Out-Of-Tolerance Condition Now ?
                        If ootFlag Then
                            ' Configured OOT Response
                            Select Case ootResp
                                Case ootrspPause
                                    ' pause the station
                                    OOT_Pause iStation, iShift, Idx, ootResp
                                Case ootrspStop
                                    ' stop the station
                                    OOT_Pause iStation, iShift, Idx, ootResp
                                Case Else
                                    ' continue or undefined; nothing to do
                            End Select
                        End If
                        
                    Next Idx
                    
                End If
            End If
        
        
            ' What Mode is this Station/Shift In?
            ChgErrModule 2, 2013
            Select Case StationControl(iStation, iShift).Mode
                
                Case VBIDLE                                         ' IDLE
                    
                    ' Start Button Pressed
                    If StationControl(iStation, iShift).StartRequest Then
                        StationControl(iStation, iShift).StartRequest = False
                        If StationControl(iStation, iShift).DBFile = "" Then
                            Station_StartPB iStation, iShift                         ' start station(& shift) sequence
                        Else
                            Delay_Box "Attempted Start of Station with Open DB File >>" & StationControl(iStation, iShift).DBFile & "<<", MSGDELAY, msgSHOW
                        End If
                    End If
                    
                Case VBIDLEWAITING                               ' Reports,etc. are Now Complete
                    StationControl(iStation, iShift).Mode = VBIDLE             ' all done; set station & shift to Idle
                 
                Case VBCOMPLETE                                  ' Testing is COMPLETE; Reports,etc. aren't (yet)
                 
                Case VBLEAK                                      ' LEAK CHECK STATUS
                    LeakCheck_Check iStation, iShift
                    
                Case VBLEAKWAIT                                 ' waiting to LEAK CHECK
                    If Not ShuttingDown Then
                        If LeakCheckControl.station = 0 Then LeakCheck_Start iStation, iShift
                    End If
                    
                Case VBLEAKERROR                                 ' LEAK CHECK ERROR
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        ALM_Write iStation, iShift, "Operator Continued after LC Failure"
                        Leak_Write iStation, iShift, LCOPERCONTINUE, NORESULT
                        Com_OutDigital icAlarmBeacon, cOFF
                        LeakCheck_Next iStation, iShift
                    End If
                    
                Case VBLEAKTEST                                  ' LEAKTEST
                    LeakTest_Check iStation, iShift
                    
                Case VBFIDPAUSE                                  ' FID PAUSE
'                    FID_PauseDelay iStation, iShift
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        ALM_Write iStation, iShift, "Operator Pushed FID Continue"
                        Load_Start iStation, iShift
                    End If
                    
                Case VBPAUSEALARM                                 ' SYSTEM PAUSED
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        StationControl(iStation, iShift).IsPausedInAlarm = False
                        ALM_Write iStation, iShift, "Alarm Pause Reset"
                        Write_ELog "Operator Continued Stn " & Format(iStation, "0") & " after System Pause"
                        Com_OutDigital icAlarmBeacon, cOFF
                        Alarm_Continue iStation, iShift
                    End If
                
                Case VBPOSTPURGE                                ' Pause After Purge
                    Pause_AfterPurge_Check iStation, iShift
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        StationControl(iStation, iShift).End_Time = 0   ' allow continue on pauses
                        ALM_Write iStation, iShift, "Operator Pushed PostPurge Continue"
                    End If
            
                Case VBPOSTPURGEOPER                            ' Pause After Purge for Operator
                    Pause_AfterPurgeForOper_Check iStation, iShift
            
                Case VBPOSTLOAD                                 ' Pause After Load
                    Pause_AfterLoad iStation, iShift
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        StationControl(iStation, iShift).End_Time = 0   ' allow continue on pauses
                        ALM_Write iStation, iShift, "Operator Pushed PostLoad Continue"
                    End If
                                
                Case VBPOSTLOADOPER                             ' Pause After Load for Operator
                    Pause_AfterLoadForOper iStation, iShift
                                
                Case VBPOSTLEAK                                 ' Pause After Leak Check
                    Pause_AfterLeak iStation, iShift
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        StationControl(iStation, iShift).End_Time = 0   ' allow continue on pauses
                        ALM_Write iStation, iShift, "Operator Pushed PostLeak Continue"
                    End If
                                
                Case VBWBPAUSE                                  ' Pause for WaterBath to reach Temp SP
                    ' OK to Load (from WaterBath Controller)
                    If LoadControl(iStation, iShift).WaterBathTempOK Then
                        If (StationControl(iStation, 1).LiveFuelCycleCount < StationRecipe(iStation, iShift).LiveFuelChgFreq) Then
                            ' ready to start a load
                            AdfControl(iStation).ReadyForLoad = True
                            ' start Load
'MsgBox "Load 0", vbInformation, "Info"
                            Load_Start CInt(iStation), CInt(1)
                        Else
                            ' manual gas pause
                            StationControl(iStation, iShift).Mode = VBGASPAUSE
                        End If
                    ElseIf StationControl(iStation, iShift).ContinueRequest Then
                        ' Continue Button Pressed
                        StationControl(iStation, iShift).ContinueRequest = False
                        ' ready to start Load ??
                        StationControl(iStation, iShift).Mode = VBGASPAUSE
                    End If
        
                    
                Case VBGASPAUSE                                  ' Pause for Live Fuel
                    ' Start LoadCycle Requested (probably by LiveFuel ADF Logic)
                    If LoadControl(iStation, iShift).CycleStartRequest Then
                        LoadControl(iStation, iShift).CycleStartRequest = False
                        ' start Load
'MsgBox "Load 1", vbInformation, "Info"
                        Load_Start iStation, iShift
                    ElseIf (STN_INFO(iStation).ADF_TANKTYPE = 90) Then
                        ' manual with waterbath only
'                        If ((StationControl(iStation, iShift).CompletedCycles Mod StationRecipe(iStation, iShift).LiveFuelChgFreq) < StationRecipe(iStation, iShift).LiveFuelChgFreq) Then
                        If (StationControl(iStation, 1).LiveFuelCycleCount < StationRecipe(iStation, iShift).LiveFuelChgFreq) Then
                            ' is WaterBath temp still OK ??
                            If LoadControl(iStation, iShift).WaterBathTempOK Then
                                AdfControl(iStation).ReadyForLoad = True
                                ' start Load
'MsgBox "Load 2", vbInformation, "Info"
                                Load_Start CInt(iStation), CInt(1)
                            End If
                        Else
                            If StationControl(iStation, iShift).ContinueRequest Then
                                StationControl(iStation, iShift).ContinueRequest = False
                                StationControl(DispStn, 1).LiveFuelCycleCount = 0
                                ' is WaterBath temp still OK ??
                                If LoadControl(iStation, iShift).WaterBathTempOK Then
                                    ' ready to start Load
                                    AdfControl(iStation).ReadyForLoad = True
                                    StationControl(iStation, 1).LiveFuelCycleCount = 0
                                    ' start Load
'MsgBox "Load 3", vbInformation, "Info"
                                    Load_Start CInt(iStation), CInt(1)
                                Else
                                    ' wait for waterbath
                                    StationControl(iStation, iShift).Mode = VBWBPAUSE
                                End If
                            End If
                        End If
                    ' Continue Button Pressed
                    ElseIf StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        If StationControl(iStation, iShift).CompletedCycles > CInt(0) Then
                            ' manual or auto drain/fill ??
                            If StationRecipe(iStation, iShift).LiveFuelChgAuto Then
                                ' auto; start AutoDrainFill
                                AdfControl(iStation).Mode = 2
                                AdfControl(iStation).Step = 0
                            Else
                                ' manual drain/fill; show Fuel Supply screen for Operator
                                If (Not frmFuelSupply.Visible) Then frmFuelSupply.Show
                                AdfControl(iStation).Message = "Waiting for Manual Fuel Change"
                                AdfControl(iStation).ButtonVisible_Done = True
                            End If
                        End If
                    ' Using AutoDrainFill ??
                    ElseIf StationRecipe(iStation, iShift).LiveFuelChgAuto Then
                        ' yes; start AutoDrainFill
                        If (AdfControl(iStation).Mode <> 2) Then
                            AdfControl(iStation).Mode = 2
                            AdfControl(iStation).Step = 0
                        End If
                    Else
                        ' no; show Fuel Supply screen for Operator
                        If (Not frmFuelSupply.Visible) Then frmFuelSupply.Show
                        AdfControl(iStation).Message = "Waiting for Manual Fuel Change"
                        AdfControl(iStation).ButtonVisible_Done = True
                    End If
        
                    
                Case VBPAUSEOOT                                  ' Pause due to OOT
                    Stn_Nit_Flow_PV(iStation, iShift) = 0
                    Stn_Btn_Flow_PV(iStation, iShift) = 0
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                        Com_OutDigital icAlarmBeacon, cOFF
                        OOT_Continue iStation, iShift               ' resume
                    End If
            
                Case VBPURGECONT                                 ' Purge continue (from alarm or OOT)
                    ' Is PurgeAir Ready Yet ?
                    If Not ShuttingDown Then Purge_ContinueDelayed iStation, iShift
                    
                Case VBPURGEWAIT                                ' Purge delay state
                    ' Continue Button Pressed
                    If StationControl(iStation, iShift).ContinueRequest Then
                        StationControl(iStation, iShift).ContinueRequest = False
                            If USINGPASLOCALCONTROL And PAS_INFO(pasTEMPERATURE).timeOut Then
                                ' Reset Local PAS Temperature Control Timeout
                                PAS_INFO(pasTEMPERATURE).timeOut = False
                                PAS_INFO(pasTEMPERATURE).TimeOutDuration = 0#
                            ElseIf USINGPASLOCALCONTROL And PAS_INFO(pasMOISTURE).timeOut Then
                                ' Reset Local PAS Moisture Control Timeout
                                PAS_INFO(pasMOISTURE).timeOut = False
                                PAS_INFO(pasMOISTURE).TimeOutDuration = 0#
                            Else
                                ' Not using local PAS control or no timeouts
                            End If
                    End If
                    ' Have we set Auto purge delay and has it timed out yet
                    If Not ShuttingDown Then Purge_StartDelayed iStation, iShift
                    
                Case VBPURGE                                     ' PURGE STATUS
                    ' Is Purge done yet?
                    Purge_Check iStation, iShift
                                    
                Case VBPRELOAD                                   ' PREPARE TO LOAD STATUS
                    ' Is PreLoad done yet?
                    PreLoad_Check iStation, iShift
                    
                Case VBLOAD                                      ' LOAD STATUS
                    ' Is Load done yet?
                    Load_Check iStation, iShift
                    
                Case VBSCALEWAIT                                 ' Waiting for Scale(s)
                    If Not ShuttingDown Then
                        startTest = False
                        ' **** Is a scale used *****
                        If StationRecipe(iStation, iShift).UsePriScale Or StationRecipe(iStation, iShift).UseAuxScale Then
                            ' Using 2 scales ?
                            If StationRecipe(iStation, iShift).UsePriScale And StationRecipe(iStation, iShift).UseAuxScale Then
                                ' Either scale in use ?
                                If Scale_In_Use(StationRecipe(iStation, iShift).PriScaleNo) Or Scale_In_Use(StationRecipe(iStation, iShift).AuxScaleNo) Then
                                    ' keep waiting
                                    StationControl(iStation, iShift).Mode = VBSCALEWAIT
                                Else
                                    ' get started
                                    startTest = True
                                End If
                            Else    '   using 1 scale
                                ' **** using a Primary Scale *****
                                If StationRecipe(iStation, iShift).UsePriScale Then
                                    If Scale_In_Use(StationRecipe(iStation, iShift).PriScaleNo) Then
                                        ' keep waiting
                                        StationControl(iStation, iShift).Mode = VBSCALEWAIT
                                    Else
                                        ' get started
                                        startTest = True
                                    End If
                                End If
                                ' **** using an Aux Scale *****
                                If StationRecipe(iStation, iShift).UseAuxScale Then
                                    If Scale_In_Use(StationRecipe(iStation, iShift).AuxScaleNo) Then
                                        ' keep waiting
                                        StationControl(iStation, iShift).Mode = VBSCALEWAIT
                                    Else
                                        ' get started
                                        startTest = True
                                    End If
                                End If
                            End If
                        Else     ' ***** No scales *****
                            ' get started
                            startTest = True
                        End If   ' is a scale used
                        ' Start Test ?
                        If startTest Then Recipe_Start iStation, iShift
                    End If
            
                    
                Case VBSHIFTWAIT                                 ' waiting for other shift to finish
                    If Not ShuttingDown Then
                        ' If the other shift is idle
                        ' Start this shift's test
                        If (iShift = 1 And StationControl(iStation, 2).ModeIsIdle_Debounced) Or _
                                (iShift = 2 And StationControl(iStation, 1).ModeIsIdle_Debounced) Or _
                                NR_SHIFT = 1 Or (StationControl(iStation, iShift).Mode = VBSHIFTWAIT And StationControl(iStation, iShift).Mode = VBSHIFTWAIT) Then
                            ' get started
                            Recipe_Start iStation, iShift
                        End If
                    End If
                    
                                               
                Case VBSTARTWAIT                                 ' Delayed Start
                    If Not ShuttingDown Then
                        ' How many seconds to go ?
                        ' Start this shift's test
                        StationControl(iStation, iShift).DelayToGo = StationControl(iStation, iShift).DelaySeconds - StationControl(iStation, iShift).TestTimer
                        ' Done yet ?
                        If StationControl(iStation, iShift).DelayToGo <= CDbl(0) Then
                            ' get started
                            StationControl(iStation, iShift).DelayToGo = CDbl(0)
                            Recipe_Start iStation, iShift
                        End If
                    End If
                    
                                               
                Case VBPAUSEVACSW                                   ' System Vacuum Switch Off; Wait for Resume from Operator after Vacuum Switch is On
                    If Not ShuttingDown Then
                        ' Continue Button Pressed
                        If StationControl(iStation, iShift).ContinueRequest Then
                            StationControl(iStation, iShift).ContinueRequest = False
                            StationControl(iStation, iShift).IsPausedInAlarm = False
                            ' update Job EventLog
                            slogmsg = "Operator Continued after System Vacuum Sw Alarm"
                            ' update Job AlarmLog
                            ALM_Write iStation, iShift, slogmsg
                            ' update event log
                            Write_ELog "Station #" & Format(iStation, "0") & " Shift #" & Format(iShift, "0") & " " & slogmsg
                            ' Resume
                            Com_OutDigital icAlarmBeacon, cOFF
                            Alarm_Continue iStation, iShift
                        End If
                    End If
                    
                                               
                Case VBPAUSEBYUSER                                ' Operator pressed Pause; Wait for Resume from Operator
                    If Not ShuttingDown Then
                        ' Continue Button Pressed
                        If StationControl(iStation, iShift).ContinueRequest Then
                            StationControl(iStation, iShift).ContinueRequest = False
                            StationSequence(iStation, iShift).CourseData(StationControl(iStation, iShift).Course).OkToProceed = True
                            ' update JobLog
                            slogmsg = "Operator Pushed OK to Resume for Course #" & Format(StationControl(iStation, iShift).Course, "##0")
                            Write_JLog iStation, iShift, slogmsg
                            ' Resume
                            Station_ContinuePB iStation, iShift
                        End If
                    End If
                    
                                               
                Case VBCOURSEWAIT                                ' JobSequence Course; Wait for OK from Operator
                    If Not ShuttingDown Then
                        ' Continue Button Pressed
                        If StationControl(iStation, iShift).ContinueRequest Then
                            StationControl(iStation, iShift).ContinueRequest = False
                            StationSequence(iStation, iShift).CourseData(StationControl(iStation, iShift).Course).OkToProceed = True
                            ' update JobLog
                            slogmsg = "Operator Pushed OK to Proceed for Course #" & Format(StationControl(iStation, iShift).Course, "##0")
                            Write_JLog iStation, iShift, slogmsg
                            ' Next Course
                            Course_Next iStation, iShift
                        End If
                    End If
                    
                                               
                Case VBCOURSEPAUSE                               ' JobSequence Course; Pause for x minutes
                    If Not ShuttingDown Then
                        ' How many minutes to go ?
                        Dim EndOfCourse As Date
                        Dim iMin, iSec As Integer
                        Dim sMin, sSec As Single
                        sMin = StationSequence(iStation, iShift).CourseData(StationControl(iStation, iShift).Course).PauseDuration
                        iMin = CInt(sMin)
                        sSec = 60 * (CSng(sMin - iMin))
                        iSec = CInt(sSec)
                        EndOfCourse = StationSequence(iStation, iShift).CourseData(StationControl(iStation, iShift).Course).DtsStart + TimeSerial(0, iMin, iSec)
                        ' Done yet ?
                        If Now > EndOfCourse Then
                            ' Next Course
                            Course_Next iStation, iShift
                            
                        ElseIf StationControl(iStation, iShift).ContinueRequest Then
                            ' Continue Button Pressed
                            StationControl(iStation, iShift).ContinueRequest = False
                            StationSequence(iStation, iShift).CourseData(StationControl(iStation, iShift).Course).OkToProceed = True
                            ' update JobLog
                            slogmsg = "Operator Cancelled Pause for Course #" & Format(StationControl(iStation, iShift).Course, "##0")
                            Write_JLog iStation, iShift, slogmsg
                            ' Next Course
                            Course_Next iStation, iShift
                        End If
                    End If
                    
                                           
                Case Else                                      ' Other
            
                    
            End Select
        
            ' CANVENTALARM Override Logic
            If USINGCANVENTALARM Then
        
                ChgErrModule 2, 2014
                
                OOTs(iStation, iShift).CanVent_TimeNow = Now()
                Select Case StationControl(iStation, iShift).Mode
                    Case VBIDLE
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case VBLEAK
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case VBLEAKWAIT
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case VBLEAKERROR
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case VBPRELOAD                                   ' PREPARE TO LOAD STATUS
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                    
                    Case VBLOAD
                        ' Reset count if LOAD just started (or restarted)
                        If StationControl(iStation, iShift).Mode_Last <> StationControl(iStation, iShift).Mode Then
                            OOTs(iStation, iShift).CanVent_DelayCount = 0
                        End If
                        ' Increment count once per second (to max of 29999)
                        OOTs(iStation, iShift).CanVent_TimeDelta = DateDiff("s", OOTs(iStation, iShift).CanVent_TimeLast, OOTs(iStation, iShift).CanVent_TimeNow)
                        If (OOTs(iStation, iShift).CanVent_TimeDelta > 0) And (OOTs(iStation, iShift).CanVent_DelayCount < 30000) And (OOTs(iStation, iShift).CanVent_TimeDelta < 30) Then
                            OOTs(iStation, iShift).CanVent_DelayCount = OOTs(iStation, iShift).CanVent_DelayCount + OOTs(iStation, iShift).CanVent_TimeDelta
                        End If
                        ' Turn On override if count less than max
                        If OOTs(iStation, iShift).CanVent_DelayCount < SysConfig.CanVent_Delay_Max Then
                            OOTs(iStation, iShift).CanVent_DelayOn = 1
                        Else
                            OOTs(iStation, iShift).CanVent_DelayOn = 0
                        End If
                        
                    Case VBPURGE
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case VBPURGECONT
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case VBPURGEWAIT
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
                        OOTs(iStation, iShift).CanVent_DelayCount = 0
                        OOTs(iStation, iShift).CanVent_DelayOn = 0
                        
                    Case Else
                        OOTs(iStation, iShift).CanVent_TimeDelta = 0
        
                        
                End Select
                OOTs(iStation, iShift).CanVent_TimeLast = Now()
                StationControl(iStation, iShift).Mode_Last = StationControl(iStation, iShift).Mode
                
            Else
                OOTs(iStation, iShift).CanVent_TimeDelta = 0
                OOTs(iStation, iShift).CanVent_DelayCount = 0
                OOTs(iStation, iShift).CanVent_DelayOn = 0
            End If
    
        End If  'Stop Button Pressed
    
    Next iShift
    
    DoEvents
    
    ChgErrModule 2, 2019
    If iStation = LAST_STN Then
        ' Set IO Force Mode for Common Valves, Etc.
        If IdlePauseCount = 0 Then
            STN_IOForceMode(0) = VBMANUAL
        Else
            STN_IOForceMode(0) = VBAUTO
        End If
        IdlePauseCount = 0
        ' Set station iStation to 1
        CurChkStn = 1
    Else
        ' Increment station iStation
        CurChkStn = CurChkStn + 1
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

Public Sub LeakCheck_SetupAux(station As Integer, Shift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3232
Dim maxout As Single
Dim span As Single

        Reset_Bar_Graph station, Shift              ' reset the bar graph
        Close_Stn_Valves station, Shift             ' close all valves
        StationControl(station, Shift).Target = SysConfig.LCSetPoint
        
         ' only purge if NOT using a LeakCheck Exhaust Valve
        If USINGLEAKCHECKEXHAUSTSOL Then
            ChgPhase LeakPurging, (Now + TimeSerial(0, 0, 1)), station, Shift
            ' Energize Leak Check Exhaust Valve
            Com_OutDigital icLeakCheckExhaustSol, cON           ' Turn ON LeakCheck Exhaust Valve
                    
        Else
            ChgPhase LeakPurging, (Now + TimeSerial(0, 0, 5)), station, Shift
            ' Open Purge Valves to help vent existing pressure
            span = Stn_AIO(station, asPurgeAirFlowSP).EuMax - Stn_AIO(station, asPurgeAirFlowSP).EuMin
            maxout = Stn_AIO(station, asPurgeAirFlowSP).EuMin + (span * Cal_MfcOutput(CSng(Stn_AIO(station, asPurgeAirFlowSP).EuMax), station, MFCPURGEAIR, Stn_MfcCal(station, MFCPURGEAIR)))
            Stn_OutAnalog station, asPurgeAirFlowSP, maxout, outNORMAL      ' open Purge MFC to max
            Stn_OutDigital station, isPurgeSol, cON                         ' station purge flow valve on
            Stn_OutDigital station, isLeakCheckSol, cON                     ' Turn on LeakCheck
            Stn_OutDigital station, isAuxLeakCheckSol, cON                  ' Turn on Aux LeakCheck
            Stn_OutDigital station, isAuxCanVentSol, cON                    ' Turn on AuxCanVent
            Stn_OutDigital station, isPriDirectionSol, cON                  ' Turn on PriDirection
        
            ' Open Nitrogen & Leakcheck Valves
            Select Case STN_INFO(station).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO             ' 0 slpm
                    Stn_OutDigital station, isNitrogenSol, cON                      ' Turn on nitro
                
                Case STN_ORVR2_TYPE
                    If StationRecipe(station, Shift).UseHiRangeMFC Then
                        ' use higher range MFC
                        Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                        Stn_OutDigital station, isNitrogenOrvrSol, cON              ' Turn on nitro
                    Else
                        ' use lower range MFC
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                        Stn_OutDigital station, isNitrogenSol, cON                  ' Turn on nitro
                    End If
                
                Case STN_LIVEFUEL_TYPE
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                    Stn_OutDigital station, isLiveFuelSol, cON                      ' Turn on live fuel vapor
'                    Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                    
                Case STN_LIVEREG_TYPE
                    If StationRecipe(station, Shift).LiveFuel Then
                        Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                        Stn_OutDigital station, isLiveFuelSol, cON                  ' Turn on live fuel vapor
                        Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                    Else
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                        Stn_OutDigital station, isNitrogenSol, cON                  ' Turn on nitro
                    End If
                    
                Case STN_LIVEORVR2_TYPE
                    If StationRecipe(station, Shift).LiveFuel Then
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO     ' 0 slpm
                            Stn_OutDigital station, isLiveFuelOrvrSol, cON              ' Turn on live fuel vapor carrier
'                            Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO         ' 0 slpm
                            Stn_OutDigital station, isLiveFuelSol, cON                       ' Turn on live fuel vapor
'                            Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                        End If
                    Else
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO          ' 0 slpm
                            Stn_OutDigital station, isNitrogenOrvrSol, cON                   ' Turn on nitro
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO              ' 0 slpm
                            Stn_OutDigital station, isNitrogenSol, cON                       ' Turn on nitro
                        End If
                    End If
                    
                Case STN_COMBO3_TYPE
                    ' future
                    
                Case Else
                    ' Nothing to do
            End Select
            
            ' Shift valves
            Select Case Shift
                Case 1
                    ' nothing to do
                Case 2
                    Stn_OutDigital station, isLoadShift2Sol, cOFF
                    Stn_OutDigital station, isPurgeShift2Sol, cON
                    Stn_OutDigital station, isVentShift2Sol, cON
                Case 3
                    Stn_OutDigital station, isLoadShift3Sol, cOFF
                    Stn_OutDigital station, isLoadShift2Sol, cOFF
                    Stn_OutDigital station, isPurgeShift2Sol, cON
                    Stn_OutDigital station, isVentShift2Sol, cON
                    Stn_OutDigital station, isPurgeShift3Sol, cON
                    Stn_OutDigital station, isVentShift3Sol, cON
                Case 4
                    Stn_OutDigital station, isLoadShift2Sol, cOFF
                    Stn_OutDigital station, isPurgeShift2Sol, cON
                    Stn_OutDigital station, isVentShift2Sol, cON
                    Stn_OutDigital station, isLoadShift4Sol, cOFF
                    Stn_OutDigital station, isPurgeShift4Sol, cON
                    Stn_OutDigital station, isVentShift4Sol, cON
            End Select
            
        End If
    
        ' Reset Report & Totalize Timers
        Stn_Leak_Log_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer           ' do first normal report after one Leak_Interval
        PreviousReportTimer(station, Shift) = StationControl(station, Shift).TestTimer              ' reset report timer
    
        ' Clear LeakData
        StnLeakData(station, Shift, 0) = BlankLeakData
        StnLeakData(station, Shift, 1) = BlankLeakData
        StnLeakData(station, Shift, 2) = BlankLeakData
        StnLeakData(station, Shift, 3) = BlankLeakData
        StnLeakData(station, Shift, 4) = BlankLeakData
        StnLeakData(station, Shift, 5) = BlankLeakData
        StnLeakData(station, Shift, 6) = BlankLeakData
        StnLeakData(station, Shift, 7) = BlankLeakData
        StnLeakData(station, Shift, 8) = BlankLeakData
        StnLeakData(station, Shift, 9) = BlankLeakData
    
        Leak_Write CInt(station), CInt(Shift), LCBEGINPHASE0, NORESULT
        

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

Public Sub LeakCheck_SetupPri(station As Integer, Shift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3231
Dim maxout As Single
Dim span As Single

        Reset_Bar_Graph station, Shift              ' reset the bar graph
        Close_Stn_Valves station, Shift             ' close all valves
        StationControl(station, Shift).Target = SysConfig.LCSetPoint
        
        ' vent via "old style" or a LeakCheck Exhaust Valve
        If USINGLEAKCHECKEXHAUSTSOL Then
            ChgPhase LeakPurging, (Now + TimeSerial(0, 0, 1)), station, Shift
            ' Energize Leak Check Exhaust Valve
            Com_OutDigital icLeakCheckExhaustSol, cON           ' Turn ON LeakCheck Exhaust Valve
        Else
            ChgPhase LeakPurging, (Now + TimeSerial(0, 0, 5)), station, Shift
            ' Open Purge Valves to help vent existing pressure
            span = Stn_AIO(station, asPurgeAirFlowSP).EuMax - Stn_AIO(station, asPurgeAirFlowSP).EuMin
            maxout = Stn_AIO(station, asPurgeAirFlowSP).EuMin + (span * Cal_MfcOutput(CSng(Stn_AIO(station, asPurgeAirFlowSP).EuMax), station, MFCPURGEAIR, Stn_MfcCal(station, MFCPURGEAIR)))
            Stn_OutAnalog station, asPurgeAirFlowSP, maxout, outNORMAL      ' open Purge MFC to max
            Stn_OutDigital station, isPurgeSol, cON                         ' station purge flow valve on
            Stn_OutDigital station, isLeakCheckSol, cON                    ' Turn on LeakCheck
        
            ' Open Nitrogen & Leakcheck Valves
            Select Case STN_INFO(station).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                    Stn_OutDigital station, isNitrogenSol, cON                  ' Turn on nitro
                
                Case STN_ORVR2_TYPE
                    If StationRecipe(station, Shift).UseHiRangeMFC Then
                        ' use higher range MFC
                        Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO         ' 0 slpm
                        Stn_OutDigital station, isNitrogenOrvrSol, cON                  ' Turn on nitro
                    Else
                        ' use lower range MFC
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                        Stn_OutDigital station, isNitrogenSol, cON                  ' Turn on nitro
                    End If
                
                Case STN_LIVEFUEL_TYPE
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                    Stn_OutDigital station, isLiveFuelSol, cON                      ' Turn on live fuel vapor
'                    Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                    
                Case STN_LIVEREG_TYPE
                    If StationRecipe(station, Shift).LiveFuel Then
                        Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                        Stn_OutDigital station, isLiveFuelSol, cON                  ' Turn on live fuel vapor
'                        Stn_OutDigital station, isLoadTypeSelectSol, cON            ' Isolate station from LiveFuel Tank
                    Else
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                        Stn_OutDigital station, isNitrogenSol, cON                  ' Turn on nitro
                    End If
                    
                Case STN_LIVEORVR2_TYPE
                    If StationRecipe(station, Shift).LiveFuel Then
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO     ' 0 slpm
                            Stn_OutDigital station, isLiveFuelOrvrSol, cON                   ' Turn on live fuel vapor carrier
'                            Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO         ' 0 slpm
                            Stn_OutDigital station, isLiveFuelSol, cON                       ' Turn on live fuel vapor
'                            Stn_OutDigital station, isLoadTypeSelectSol, cON                ' Isolate station from LiveFuel Tank
                        End If
                    Else
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO          ' 0 slpm
                            Stn_OutDigital station, isNitrogenOrvrSol, cON                   ' Turn on nitro
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                            Stn_OutDigital station, isNitrogenSol, cON                  ' Turn on nitro
                        End If
                    End If
                    
                Case STN_COMBO3_TYPE
                    ' future
                    
                Case Else
                    ' Nothing to do
            End Select
            
            ' Shift valves
            Select Case Shift
                Case 1
                    ' nothing to do
                Case 2
                    Stn_OutDigital station, isLoadShift2Sol, cOFF
                    Stn_OutDigital station, isPurgeShift2Sol, cON
                    Stn_OutDigital station, isVentShift2Sol, cON
                Case 3
                    Stn_OutDigital station, isLoadShift3Sol, cOFF
                    Stn_OutDigital station, isLoadShift2Sol, cOFF
                    Stn_OutDigital station, isPurgeShift2Sol, cON
                    Stn_OutDigital station, isVentShift2Sol, cON
                    Stn_OutDigital station, isPurgeShift3Sol, cON
                    Stn_OutDigital station, isVentShift3Sol, cON
                Case 4
                    Stn_OutDigital station, isLoadShift2Sol, cOFF
                    Stn_OutDigital station, isPurgeShift2Sol, cON
                    Stn_OutDigital station, isVentShift2Sol, cON
                    Stn_OutDigital station, isLoadShift4Sol, cOFF
                    Stn_OutDigital station, isPurgeShift4Sol, cON
                    Stn_OutDigital station, isVentShift4Sol, cON
            End Select
            
        End If
    
        ' PriAuxVent Valve
        If ((StationRecipe(station, Shift).UsePriScale) And (StationRecipe(station, Shift).PriScaleNo > 0) _
                And (StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE)) Then
            Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cON
        End If
        ' AuxDirection Valve
        If ((StationRecipe(station, Shift).UseAuxScale) And (StationRecipe(station, Shift).AuxScaleNo > 0)) Then
            Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxDirectionSol, cON
        End If
        ' Reset Report & Totalize Timers
        Stn_Leak_Log_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer           ' do first normal report after one Leak_Interval
        PreviousReportTimer(station, Shift) = StationControl(station, Shift).TestTimer              ' reset report timer
    
        ' Clear LeakData
        StnLeakData(station, Shift, 0) = BlankLeakData
        StnLeakData(station, Shift, 1) = BlankLeakData
        StnLeakData(station, Shift, 2) = BlankLeakData
        StnLeakData(station, Shift, 3) = BlankLeakData
        StnLeakData(station, Shift, 4) = BlankLeakData
        StnLeakData(station, Shift, 5) = BlankLeakData
        StnLeakData(station, Shift, 6) = BlankLeakData
        StnLeakData(station, Shift, 7) = BlankLeakData
        StnLeakData(station, Shift, 8) = BlankLeakData
        StnLeakData(station, Shift, 9) = BlankLeakData
    
        Leak_Write CInt(station), CInt(Shift), LCBEGINPHASE0, NORESULT
        
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

Public Sub LeakCheck_Start(station As Integer, Shift As Integer)
'*************************************************************************************
'
' Start of a Leak Check needs to have valves and Mass Air Controllers set up properly
'
'*************************************************************************************
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 32

    If LeakCheckControl.station = 0 Then
        LeakCheckControl.station = station                         ' Starting a process with leak check
        LeakCheckControl.Shift = Shift
        LeakCheckControl.StartTime = Now
        LeakCheckControl.StartTimer = StationControl(station, Shift).TestTimer
        StationControl(station, Shift).IsPausedInAlarm = False
        StationControl(station, Shift).Mode_StartDts = Now()
        StationControl(station, Shift).Mode = VBLEAK
        
        If StationRecipe(station, Shift).LeakPrimary Then
            LeakCheckControl.Method = LEAKCHECKPRI
            LeakCheck_SetupPri station, Shift
        ElseIf StationRecipe(station, Shift).LeakAux Then
            LeakCheckControl.Method = LEAKCHECKAUX
            LeakCheck_SetupAux station, Shift
        Else
            LeakCheckControl.Method = NOLEAKCHECK
            LeakCheck_Next station, Shift
        End If
        
    Else
        ' waiting for purge LCP transducer
        StationControl(station, Shift).Mode = VBLEAKWAIT             ' "Waiting Leak paused"
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

Public Sub LeakCheck_Done(station As Integer, Shift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 33
Dim sMsg As String

    StationControl(station, Shift).End_Time = Now
    Select Case LeakCheckControl.Method
        Case LEAKCHECKPRI
            StationControl(station, Shift).LeakCheckStatus = RESULTGOOD
            StationControl(station, Shift).LcStatusDescription = LeakResultDesc(RESULTGOOD)
        Case Else
            ' Nothing to do
    End Select

    sMsg = LeakCanisterDesc(LeakCheckControl.Method) & " Leakcheck " & LeakResultDesc(RESULTGOOD)
    Reset_Bar_Graph station, Shift                                          ' reset the bar graph
    Leak_Write station, Shift, LCTESTRESULT, RESULTGOOD                     ' passed with no errors
    Write_ELog sMsg & "  Station " & station & "  Shift" & Shift
    Write_JLog station, Shift, sMsg
    ChgPhase LeakComplete, Now, station, Shift
    
    Close_Stn_Valves station, Shift                                         ' close all valves
    Select Case STN_INFO(station).Type
    
      Case STN_REGULAR_TYPE, STN_ORVR_TYPE
          Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO               ' 0 slpm
      
      Case STN_ORVR2_TYPE
          If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
          Else
                ' use lower range MFC
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
          End If
      
      Case STN_LIVEFUEL_TYPE
          Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO          ' 0 slpm
          Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      
      Case STN_LIVEREG_TYPE
          If StationRecipe(station, Shift).LiveFuel Then
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                Stn_OutDigital station, isLoadTypeSelectSol, cOFF
          Else
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
          End If
      
      Case STN_LIVEORVR2_TYPE
          If StationRecipe(station, Shift).LiveFuel Then
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                  ' use higher range MFC
                    Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO    ' 0 slpm
                    Stn_OutDigital station, isLoadTypeSelectSol, cOFF
            Else
                  ' use lower range MFC
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                    Stn_OutDigital station, isLoadTypeSelectSol, cOFF
            End If
          Else
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                  ' use higher range MFC
                  Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
            Else
                  ' use lower range MFC
                  Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
            End If
          End If
      
      Case STN_COMBO3_TYPE
          ' future
      
      Case Else
      
    End Select
    
    Close_Stn_Valves station, Shift             ' Always end by closing all valves & indicators
    If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
            And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
    End If
    ' Turn Off the Leak Check Exhaust Valve
    If (LeakCheckControl.station = station) Or (LeakCheckControl.station = 0) Then
        If USINGLEAKCHECKEXHAUSTSOL Then Com_OutDigital icLeakCheckExhaustSol, cOFF    ' Turn off LeakCheck Exhaust Valve
    End If
    
    ' what to do next
    LeakCheck_Next station, Shift
      
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

Public Sub LeakCheck_Error(station As Integer, Shift As Integer, err As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 34
Dim sMsg As String

    StnRemoteTask(station, Shift).PreviousResult = "LeakCheck Error"
      
    StationControl(station, Shift).End_Time = Now
    StationControl(station, Shift).Mode = VBLEAKERROR      ' Leak check failed
    Select Case LeakCheckControl.Method
        Case LEAKCHECKPRI
            StationControl(station, Shift).LeakCheckStatus = err
            StationControl(station, Shift).LcStatusDescription = LeakResultDesc(err)
        Case Else
            ' Nothing to do
    End Select
     
    sMsg = LeakCanisterDesc(LeakCheckControl.Method) & " Leakcheck " & LeakResultDesc(err)
    ALM_Write station, Shift, LeakResultDesc(err)
    Write_ELog sMsg & "  Station " & station & "  Shift" & Shift
    Write_JLog station, Shift, sMsg
    If DispStn = station And DispShift = Shift Then
        If NR_SHIFT = 1 Then
'            frmStnDetail.NewMessage LeakResultDesc(err) & "  Station " & station
        Else
'            frmStnDetail.NewMessage LeakResultDesc(err) & "  Station " & station & "  Shift" & Shift
        End If
    End If
    
    Reset_Bar_Graph station, Shift              ' reset the bar graph
    Leak_Write station, Shift, LCTESTRESULT, err
    ChgPhase LeakComplete, Now, station, Shift
    
    Close_Stn_Valves station, Shift                                         ' close all valves
    Select Case STN_INFO(station).Type
    
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO             ' 0 slpm
      
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
            Else
                ' use lower range MFC
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
            End If
        
        Case STN_LIVEFUEL_TYPE
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
        
        Case STN_LIVEREG_TYPE
          If StationRecipe(station, Shift).LiveFuel Then
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                Stn_OutDigital station, isLoadTypeSelectSol, cOFF
          Else
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
          End If
        
        Case STN_LIVEORVR2_TYPE
          If StationRecipe(station, Shift).LiveFuel Then
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO    ' 0 slpm
                    Stn_OutDigital station, isLoadTypeSelectSol, cOFF
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                    Stn_OutDigital station, isLoadTypeSelectSol, cOFF
                End If
          Else
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                End If
          End If
        
        Case STN_COMBO3_TYPE
            ' future
        
        Case Else
      
    End Select
    
    Close_Stn_Valves station, Shift             ' Always end by closing all valves & indicators
    If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
      And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
    End If
    ' leakcheck exhaust solenoid
    If (LeakCheckControl.station = station) Or (LeakCheckControl.station = 0) Then
        If USINGLEAKCHECKEXHAUSTSOL Then Com_OutDigital icLeakCheckExhaustSol, cOFF    ' Turn off LeakCheck Exhaust Valve
    End If
    
    ' what to do next
    Select Case StationConfig(station, Shift).LeakCheckFailResponse
    
        Case MANUALSTOP
            ' release resources & wait for operator to press STOP
            If StationRecipe(station, Shift).UseAuxScale Then
                Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = False
            End If
            If StationRecipe(station, Shift).UsePriScale Then
                Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = False
            End If
            If LeakCheckControl.station = station Then
                LeakCheckControl.station = 0
                LeakCheckControl.Shift = 0
                LeakCheckControl.Phase = 0
            End If
    
            
        Case AUTOSTOP
            ' release resources & stop test
            If StationRecipe(station, Shift).UseAuxScale Then
                Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = False
            End If
            If StationRecipe(station, Shift).UsePriScale Then
                Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = False
            End If
            If LeakCheckControl.station = station Then
                LeakCheckControl.station = 0
                LeakCheckControl.Shift = 0
                LeakCheckControl.Phase = 0
            End If
            ' stop the test
            Station_Abort station, Shift, AUTO_STOP
            
        Case AUTOCONTINUE
            ' continue anyway
            ALM_Write station, Shift, "Automatic Continue after LC Failure"
            Leak_Write station, Shift, LCAUTOCONTINUE, NORESULT
            LeakCheck_Next station, Shift
      
            
        Case Else
            ' do not release FID and Scales
            
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

Public Sub LeakCheck_Next(iStation As Integer, iShift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3434
    
    ' what to do next ??
    Select Case LeakCheckControl.Method
        Case LEAKCHECKPRI
            ' finished Primary; Test Aux or Pause or Other ??
            If StationRecipe(iStation, iShift).LeakAux Then
                LeakCheckControl.Method = LEAKCHECKAUX
                LeakCheck_SetupAux iStation, iShift
            Else
                If LeakCheckControl.station = iStation Then
                    LeakCheckControl.station = 0
                    LeakCheckControl.Shift = 0
                    LeakCheckControl.Phase = 0
                End If
                If StationRecipe(iStation, iShift).PauseAfterLeak Then
                    Pause_AfterLeak iStation, iShift
                Else
                    LeakCheck_Continue iStation, iShift
                End If
            End If
        Case LEAKCHECKAUX
            ' finished Aux; Pause or Other ??
            If LeakCheckControl.station = iStation Then
                LeakCheckControl.station = 0
                LeakCheckControl.Shift = 0
                LeakCheckControl.Phase = 0
            End If
            If StationRecipe(iStation, iShift).PauseAfterLeak Then
                Pause_AfterLeak iStation, iShift
            Else
                LeakCheck_Continue iStation, iShift
            End If
        Case Else
            ' do something else
            LeakCheck_Continue iStation, iShift
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

Public Sub LeakCheck_Continue(station As Integer, Shift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3536
    
    ' What is next after this LeakCheck ??, Purge ??, Load ??, Done ??
    Select Case StationRecipe(station, Shift).CycleType
    
        Case CyclePurgeLoad
            ' Purge then Load Cycles
            If StationRecipe(station, Shift).Purge_Method <> NOPURGE Then   ' let's set up for purge next
                ' Start of a Purge
                Purge_Start station, Shift
            Else
                If StationRecipe(station, Shift).Cycles >= 0 Then
                    If StationRecipe(station, Shift).Load_Method <> NOLOAD Then
                        ' LOAD ONLY
                        PreLoad_Start station, Shift
                    Else
                        '  Must be a leak check only
                        JobInfo(station, Shift).End_OK = True
                        ' What's Next ??
                        Course_Next station, Shift
                    End If
                Else
                    '  Must be a leak check only
                    JobInfo(station, Shift).End_OK = True
                    ' What's Next ??
                    Course_Next station, Shift
                End If
            End If
            
        Case CycleLoadPurge
            ' Load then Purge Cycles
            If StationRecipe(station, Shift).Load_Method <> NOLOAD Then   ' let's set up for load next
                ' Start of a Load
                PreLoad_Start station, Shift
            Else
                If StationRecipe(station, Shift).Cycles >= 0 Then
                    If StationRecipe(station, Shift).Purge_Method <> NOPURGE Then
                        ' PURGE ONLY
                        Purge_Start station, Shift
                    Else
                        '  Must be a leak check only
                        JobInfo(station, Shift).End_OK = True
                        ' What's Next ??
                        Course_Next station, Shift
                    End If
                Else
                    '  Must be a leak check only
                    JobInfo(station, Shift).End_OK = True
                    ' What's Next ??
                    Course_Next station, Shift
                End If
            End If
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

Public Sub LeakCheck_Abort(station As Integer, Shift As Integer, code As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 35
Dim iCourse As Integer
Dim sPrint As String
      
    ' time stamp end of current course
    iCourse = StationControl(station, Shift).Course
    StationSequence(station, Shift).CourseData(iCourse).DtsEnd = Now()
    If StationControl(station, Shift).Mode <> VBLEAKERROR Then
        StationControl(station, Shift).End_Time = Now
        StationControl(station, Shift).End_Timer = StationControl(station, Shift).TestTimer
        Leak_Write station, Shift, LCTESTRESULT, RESULTABORTOPER
    End If
      
    ' Update Header data in data file
    Header_Update station, Shift
    ' Write CycleWeights data in data file
    Weights_Write station, Shift
      
    StationControl(station, Shift).IsPausedInAlarm = False
    Reset_Bar_Graph station, Shift              ' reset the bar graph
    Write_ELog LeakResultDesc(code) & "  Station " & station & "  Shift" & Shift
    ChgPhase LeakComplete, Now, station, Shift
    
    Close_Stn_Valves station, Shift                                         ' close all valves
    Select Case STN_INFO(station).Type
    
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO            ' 0 slpm
      
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
            Else
                ' use lower range MFC
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
            End If
        
        Case STN_LIVEFUEL_TYPE
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO       ' 0 slpm
            Stn_OutDigital station, isLoadTypeSelectSol, cOFF
        
        Case STN_LIVEREG_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                  Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO    ' 0 slpm
                  Stn_OutDigital station, isLoadTypeSelectSol, cOFF
            Else
                  Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
            End If
        
        Case STN_LIVEORVR2_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO    ' 0 slpm
'                    Stn_OutDigital station, isLoadTypeSelectSol, cOFF
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
'                    Stn_OutDigital station, isLoadTypeSelectSol, cOFF
                End If
            Else
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                End If
            End If
        
        Case STN_COMBO3_TYPE
            ' future
        
        Case Else
      
    End Select
    
    ' Close valves & indicators
    Close_Stn_Valves station, Shift
    If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
            And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
    End If
    ' Turn Off the Leak Check Exhaust Valve
    If (LeakCheckControl.station = station) Or (LeakCheckControl.station = 0) Then
        If USINGLEAKCHECKEXHAUSTSOL Then Com_OutDigital icLeakCheckExhaustSol, cOFF    ' Turn off LeakCheck Exhaust Valve
    End If
    
    If LeakCheckControl.station = station Then
        LeakCheckControl.station = 0
        LeakCheckControl.Shift = 0
        LeakCheckControl.Phase = 0
    End If
    
    ' write the logs and close off the world
    Station_Finish station, Shift
    
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

Public Sub Pause_AfterPurge_Start(station As Integer, Shift As Integer)
'
    StationControl(station, Shift).Mode = VBPOSTPURGE
    ' net scale PurgePause start weights
    If StationRecipe(station, Shift).UsePriScale Then
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Pri = StationControl(station, Shift).PriScaleWt
    Else
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Pri = 0
    End If
    If StationRecipe(station, Shift).UseAuxScale Then
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Aux = StationControl(station, Shift).AuxScaleWt
    Else
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Aux = 0
    End If
    If USINGCANVENTALARM And StationRecipe(station, Shift).PausePurgeTime > 0 _
        And StationRecipe(station, Shift).Load_Method <> NOLOAD Then
        SEQ_Nmbr(station, Shift) = seqCanVentN2Feed       ' Post Purge N2 Feed
        SEQ_Step(station, Shift) = 1
    Else
        StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PausePurgeTime)
        SEQ_Nmbr(station, Shift) = seqIdle              ' idle
        SEQ_Step(station, Shift) = 0
    End If
End Sub

Public Sub Pause_AfterPurgeForOper_Start(station As Integer, Shift As Integer)
'
    StationControl(station, Shift).Mode = VBPOSTPURGEOPER
    ' net scale PurgePause start weights
    If StationRecipe(station, Shift).UsePriScale Then
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Pri = StationControl(station, Shift).PriScaleWt
    Else
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Pri = 0
    End If
    If StationRecipe(station, Shift).UseAuxScale Then
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Aux = StationControl(station, Shift).AuxScaleWt
    Else
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_StartWeight_Aux = 0
    End If
    If USINGCANVENTALARM And StationRecipe(station, Shift).PausePurgeTime > 0 _
        And StationRecipe(station, Shift).Load_Method <> NOLOAD Then
        SEQ_Nmbr(station, Shift) = seqCanVentN2Feed       ' Post Purge N2 Feed
        SEQ_Step(station, Shift) = 1
    Else
        StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PausePurgeTime)
        SEQ_Nmbr(station, Shift) = seqIdle              ' idle
        SEQ_Step(station, Shift) = 0
    End If
End Sub

Public Sub Pause_AfterPurge_Check(station As Integer, Shift As Integer)
'
Dim tmpWt As Single
    ' Is Post Purge N2 Feed Sequence Running ?
    If SEQ_Nmbr(station, Shift) = seqCanVentN2Feed Then
        StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PausePurgeTime)
        Select Case SEQ_Step(station, Shift)
            
            Case 9
                ' Successful Completion; Reset Sequence Number
                SEQ_Nmbr(station, Shift) = seqIdle              ' idle
                SEQ_Step(station, Shift) = 0
                ' Start Post Purge Delay
                StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PausePurgeTime)
           
            Case 90, 91, 95
                ' Aborted
            
        End Select
        
    ElseIf (Now() >= StationControl(station, Shift).End_Time) Then
            ' net scale PurgePause end weights
            If StationRecipe(station, Shift).UsePriScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Pri = StationControl(station, Shift).PriScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Pri = 0
            End If
            If StationRecipe(station, Shift).UseAuxScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Aux = StationControl(station, Shift).AuxScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Aux = 0
            End If
        ' PauseAfterPurge is complete; is cycle complete ??
        If (StationControl(station, Shift).CompletedLoads = StationControl(station, Shift).CompletedPurges) Then
            ' cycle is complete
            StationControl(station, Shift).CompletedCycles = StationControl(station, Shift).CompletedCycles + 1
            ' net scale end weight
            tmpWt = CSng(0)
            If StationRecipe(station, Shift).UsePriScale Then tmpWt = tmpWt + StationControl(station, Shift).PriScaleWt
            If StationRecipe(station, Shift).UseAuxScale Then tmpWt = tmpWt + StationControl(station, Shift).AuxScaleWt
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total = tmpWt
            ' is recipe complete ??
            If RecipeIsDone(station, Shift) Then
                ' recipe completed OK
                JobInfo(station, Shift).End_OK = True
                ' any more LiveFuel Loads ??
                If Not AnyMoreLiveFuelLoads(station, Shift) Then
                    ' No More LiveFuel Loads; using AutoDrainFill ??
                    If (AdfControl(station).AdfDefinition.hasAUTODRAINFILL And AdfControl(station).LiveFuel And AdfControl(station).LiveFuelChgAuto) Then
                        ' Got to Empty the Live Fuel Tank
                        AdfControl(station).Mode = 1
                        AdfControl(station).Step = 0
                    Else
                        AdfControl(station).Mode = 0
                        AdfControl(station).Step = 0
                    End If
                End If
                ' *****************************************
                ' What's Next ??
                Course_Next station, Shift
                ' *****************************************
            Else
                ' recipe is not complete; new cycle needs a load
                StationControl(station, Shift).CurrCycle = StationControl(station, Shift).CurrCycle + 1
                ' new cycle start weight = last cycle end weight
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).Cycle_StartWeight_Total = _
                    StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total
                If (StationRecipe(station, Shift).Load_Method = NOLOAD) Then
                    ' End a Load
                    Load_Done station, Shift
                Else
                    ' Start a Load
                    PreLoad_Start station, Shift
                End If
            End If
        Else
            ' cycle is not complete; current cycle needs a Load
            If (StationRecipe(station, Shift).Load_Method = NOLOAD) Then
                ' End a Load
                Load_Done station, Shift
            Else
                ' Start a Load
                PreLoad_Start station, Shift
            End If
        End If
        
    End If
    
End Sub

Public Sub Pause_AfterPurgeForOper_Check(station As Integer, Shift As Integer)
'
Dim tmpWt As Single
    ' Is Post Purge N2 Feed Sequence Running ?
    If SEQ_Nmbr(station, Shift) = seqCanVentN2Feed Then
        StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PausePurgeTime)
        Select Case SEQ_Step(station, Shift)
            
            Case 9
                ' Successful Completion; Reset Sequence Number
                SEQ_Nmbr(station, Shift) = seqIdle              ' idle
                SEQ_Step(station, Shift) = 0
                ' Start Post Purge Delay
                StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PausePurgeTime)
           
            Case 90, 91, 95
                ' Aborted
            
        End Select
        
        
    ElseIf StationControl(station, Shift).ContinueRequest Then
            ' Continue Button Pressed
            StationControl(station, Shift).ContinueRequest = False
            ALM_Write station, Shift, "Operator Pushed PostLoadForOper Continue"
            ' net scale PurgePause end weights
            If StationRecipe(station, Shift).UsePriScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Pri = StationControl(station, Shift).PriScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Pri = 0
            End If
            If StationRecipe(station, Shift).UseAuxScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Aux = StationControl(station, Shift).AuxScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).PurgePause_EndWeight_Aux = 0
            End If
        ' PauseAfterPurgeForOper is complete; is cycle complete ??
        If (StationControl(station, Shift).CompletedLoads = StationControl(station, Shift).CompletedPurges) Then
            ' cycle is complete
            StationControl(station, Shift).CompletedCycles = StationControl(station, Shift).CompletedCycles + 1
            ' net scale end weight
            tmpWt = CSng(0)
            If StationRecipe(station, Shift).UsePriScale Then tmpWt = tmpWt + StationControl(station, Shift).PriScaleWt
            If StationRecipe(station, Shift).UseAuxScale Then tmpWt = tmpWt + StationControl(station, Shift).AuxScaleWt
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total = tmpWt
            ' is recipe complete ??
            If RecipeIsDone(station, Shift) Then
                ' recipe completed OK
                JobInfo(station, Shift).End_OK = True
                ' any more LiveFuel Loads ??
                If Not AnyMoreLiveFuelLoads(station, Shift) Then
                    ' No More LiveFuel Loads; using AutoDrainFill ??
                    If (AdfControl(station).AdfDefinition.hasAUTODRAINFILL And AdfControl(station).LiveFuel And AdfControl(station).LiveFuelChgAuto) Then
                        ' Got to Empty the Live Fuel Tank
                        AdfControl(station).Mode = 1
                        AdfControl(station).Step = 0
                    Else
                        AdfControl(station).Mode = 0
                        AdfControl(station).Step = 0
                    End If
                End If
                ' *****************************************
                ' What's Next ??
                Course_Next station, Shift
                ' *****************************************
            Else
                ' recipe is not complete; new cycle needs a load
                StationControl(station, Shift).CurrCycle = StationControl(station, Shift).CurrCycle + 1
                ' new cycle start weight = last cycle end weight
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).Cycle_StartWeight_Total = _
                    StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total
                If (StationRecipe(station, Shift).Load_Method = NOLOAD) Then
                    ' End a Load
                    Load_Done station, Shift
                Else
                    ' Start a Load
                    PreLoad_Start station, Shift
                End If
            End If
        Else
            ' cycle is not complete; current cycle needs a Load
            If (StationRecipe(station, Shift).Load_Method = NOLOAD) Then
                ' End a Load
                Load_Done station, Shift
            Else
                ' Start a Load
                PreLoad_Start station, Shift
            End If
        End If
        
    End If
    
End Sub

Public Sub Pause_AfterLoad(station As Integer, Shift As Integer)
'
Dim tmpWt As Single
    If StationControl(station, Shift).Mode = VBPOSTLOAD Then
        ' is PauseAfterLoad complete ??
        If Now() >= StationControl(station, Shift).End_Time Then
            ' net scale LoadPause end weights
            If StationRecipe(station, Shift).UsePriScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Pri = StationControl(station, Shift).PriScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Pri = 0
            End If
            If StationRecipe(station, Shift).UseAuxScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Aux = StationControl(station, Shift).AuxScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Aux = 0
            End If
            ' PauseAfterLoad is complete; is cycle complete ??
            If (StationControl(station, Shift).CompletedLoads = StationControl(station, Shift).CompletedPurges) Then
                ' cycle is complete
                StationControl(station, Shift).CompletedCycles = StationControl(station, Shift).CompletedCycles + 1
                ' net scale end weight
                tmpWt = CSng(0)
                If StationRecipe(station, Shift).UsePriScale Then tmpWt = tmpWt + StationControl(station, Shift).PriScaleWt
                If StationRecipe(station, Shift).UseAuxScale Then tmpWt = tmpWt + StationControl(station, Shift).AuxScaleWt
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total = tmpWt
                ' is recipe complete ??
                If RecipeIsDone(station, Shift) Then
                    ' recipe completed OK
                    JobInfo(station, Shift).End_OK = True
                    ' any more LiveFuel Loads ??
                    If Not AnyMoreLiveFuelLoads(station, Shift) Then
                        ' No More LiveFuel Loads; using AutoDrainFill ??
                        If (AdfControl(station).AdfDefinition.hasAUTODRAINFILL And AdfControl(station).LiveFuel And AdfControl(station).LiveFuelChgAuto) Then
                            ' Got to Empty the Live Fuel Tank
                            AdfControl(station).Mode = 1
                            AdfControl(station).Step = 0
                        Else
                            AdfControl(station).Mode = 0
                            AdfControl(station).Step = 0
                        End If
                    End If
                    ' *****************************************
                    ' What's Next ??
                    Course_Next station, Shift
                    ' *****************************************
                Else
                    ' recipe is not complete; new cycle needs a purge
                    StationControl(station, Shift).CurrCycle = StationControl(station, Shift).CurrCycle + 1
                    ' new cycle start weight = last cycle end weight
                    StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).Cycle_StartWeight_Total = _
                        StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total
                    If (StationRecipe(station, Shift).Purge_Method = NOPURGE) Then
                        ' End a Purge
                        Purge_Done station, Shift
                    Else
                        ' Start a Purge
                        Purge_Start station, Shift
                    End If
                End If
            Else
                ' cycle is not complete; current cycle needs a Purge
                If (StationRecipe(station, Shift).Purge_Method = NOPURGE) Then
                    ' End a Purge
                    Purge_Done station, Shift
                Else
                    ' Start a Purge
                    Purge_Start station, Shift
                End If
            End If
        
                        
        End If
    Else
        ' net scale LoadPause start weights
        If StationRecipe(station, Shift).UsePriScale Then
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Pri = StationControl(station, Shift).PriScaleWt
        Else
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Pri = 0
        End If
        If StationRecipe(station, Shift).UseAuxScale Then
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Aux = StationControl(station, Shift).AuxScaleWt
        Else
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Aux = 0
        End If
        ' start PauseAfterLoad
        StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PauseLoadTime)
        StationControl(station, Shift).Mode = VBPOSTLOAD
    End If
End Sub

Public Sub Pause_AfterLoadForOper(station As Integer, Shift As Integer)
'
Dim tmpWt As Single
    If StationControl(station, Shift).Mode = VBPOSTLOADOPER Then
        ' is PauseAfterLoad complete ??
        If StationControl(station, Shift).ContinueRequest Then
            ' Continue Button Pressed
            StationControl(station, Shift).ContinueRequest = False
            ALM_Write station, Shift, "Operator Pushed PostLoadForOper Continue"
            ' net scale LoadPause end weights
            If StationRecipe(station, Shift).UsePriScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Pri = StationControl(station, Shift).PriScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Pri = 0
            End If
            If StationRecipe(station, Shift).UseAuxScale Then
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Aux = StationControl(station, Shift).AuxScaleWt
            Else
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_EndWeight_Aux = 0
            End If
            ' PauseAfterLoadForOper is complete; is cycle complete ??
            If (StationControl(station, Shift).CompletedLoads = StationControl(station, Shift).CompletedPurges) Then
                ' cycle is complete
                StationControl(station, Shift).CompletedCycles = StationControl(station, Shift).CompletedCycles + 1
                ' net scale end weight
                tmpWt = CSng(0)
                If StationRecipe(station, Shift).UsePriScale Then tmpWt = tmpWt + StationControl(station, Shift).PriScaleWt
                If StationRecipe(station, Shift).UseAuxScale Then tmpWt = tmpWt + StationControl(station, Shift).AuxScaleWt
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total = tmpWt
                ' is recipe complete ??
                If RecipeIsDone(station, Shift) Then
                    ' recipe completed OK
                    JobInfo(station, Shift).End_OK = True
                    ' any more LiveFuel Loads ??
                    If Not AnyMoreLiveFuelLoads(station, Shift) Then
                        ' No More LiveFuel Loads; using AutoDrainFill ??
                        If (AdfControl(station).AdfDefinition.hasAUTODRAINFILL And AdfControl(station).LiveFuel And AdfControl(station).LiveFuelChgAuto) Then
                            ' Got to Empty the Live Fuel Tank
                            AdfControl(station).Mode = 1
                            AdfControl(station).Step = 0
                        Else
                            AdfControl(station).Mode = 0
                            AdfControl(station).Step = 0
                        End If
                    End If
                    ' *****************************************
                    ' What's Next ??
                    Course_Next station, Shift
                    ' *****************************************
                Else
                    ' recipe is not complete; new cycle needs a purge
                    StationControl(station, Shift).CurrCycle = StationControl(station, Shift).CurrCycle + 1
                    ' new cycle start weight = last cycle end weight
                    StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).Cycle_StartWeight_Total = _
                        StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total
                    If (StationRecipe(station, Shift).Purge_Method = NOPURGE) Then
                        ' End a Purge
                        Purge_Done station, Shift
                    Else
                        ' Start a Purge
                        Purge_Start station, Shift
                    End If
                End If
            Else
                ' cycle is not complete; current cycle needs a Purge
                If (StationRecipe(station, Shift).Purge_Method = NOPURGE) Then
                    ' End a Purge
                    Purge_Done station, Shift
                Else
                    ' Start a Purge
                    Purge_Start station, Shift
                End If
            End If
        
                        
        End If
    Else
        ' net scale LoadPause start weights
        If StationRecipe(station, Shift).UsePriScale Then
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Pri = StationControl(station, Shift).PriScaleWt
        Else
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Pri = 0
        End If
        If StationRecipe(station, Shift).UseAuxScale Then
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Aux = StationControl(station, Shift).AuxScaleWt
        Else
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).LoadPause_StartWeight_Aux = 0
        End If
        ' start PauseAfterLoadForOper
        StationControl(station, Shift).Mode = VBPOSTLOADOPER
    End If
End Sub

Public Sub Pause_AfterLeak(station As Integer, Shift As Integer)
    If StationControl(station, Shift).Mode = VBPOSTLEAK Then
        If Now() >= StationControl(station, Shift).End_Time Then
            LeakCheck_Continue station, Shift
        End If
    Else
        StationControl(station, Shift).End_Time = Now() + MinutesFromNow(StationRecipe(station, Shift).PauseLeakTime)
        StationControl(station, Shift).Mode = VBPOSTLEAK
    End If
End Sub

Public Sub Purge_StartDelayed(ByVal iStation As Integer, ByVal iShift As Integer)
 
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3777
Dim count1 As Integer
Dim count2 As Integer
Dim timenow As Date
Dim SomebodyPurging As Integer
Dim LeakInterlock As Boolean
Dim deltaMin As Long

    ' Request Aspirator for Aux Scale (if needed)(if Aspirator needs time to get Ready)
    If StationRecipe(iStation, iShift).PurgeAuxCan And (StationRecipe(iStation, iShift).AuxScaleNo > 0) Then
        ' Only allowed to purge Aux Scale with a Vacuum Purge
        If Not StationConfig(iStation, iShift).PosPressPurge Then
    '        If PRG_INFO(STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum).UsingPrgReqHdw Thenf
            If (StationRecipe(iStation, iShift).AuxScaleNo <= LAST_STN) Then
                If (StationRecipe(iStation, iShift).AuxScaleNo <= LAST_STN) Then
                    PRG_INFO(STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum).RequestRdy = True
                End If
            End If
    '        End If
        End If
    End If
                                
    ' Request Aspirator for Station (if Aspirator needs time to get Ready)
    'If PRG_INFO(STN_INFO(station).AspiratorNum).UsingPrgReqHdw Then
        PRG_INFO(STN_INFO(iStation).AspiratorNum).RequestRdy = True
    'End If
    
    If Pause_Alarm = NOTPAUSED Then
        timenow = Now()
        SomebodyPurging = 0
        LeakInterlock = False
        For count1 = 1 To LAST_STN
           For count2 = 1 To NR_SHIFT
             If StationControl(count1, count2).Mode = VBLEAK Then
                If (USINGPRESSUREPURGE And SysConfig.PosPressPurge) Then
                    LeakInterlock = True
                End If
             End If
             ' Anyone Started Purging within the last 5 minutes ?
             If StationControl(count1, count2).Mode = VBPURGE And PurgeControl(count1, count2).Phase < PurgePause Then
                deltaMin = DateDiff("n", PurgeControl(count1, count2).StartTime, timenow)
                If deltaMin < 5 Then SomebodyPurging = SomebodyPurging + 1
             End If
           Next count2
        Next count1
    End If
    
    
    '   OK to Start a Purge ?
    If LeakInterlock = False Then
        ' No Leak Interlock
        If SomebodyPurging = 0 Or (timenow >= LastPurgeStart + TimeSerial(0, 5, 0)) Then
            ' No purges running or five minutes has elapsed since most recent station started purging
            If PRG_INFO(STN_INFO(iStation).AspiratorNum).Ready Then
                ' Using a Purge Oven and the Oven Temp is within the desired band
                If (USINGPURGEOVEN And StationRecipe(iStation, iShift).PurgeOven) Then
                    If (PurgeControl(iStation, iShift).PurgeOvenTempOK) Then
                        ' Start one
                        Purge_Start iStation, iShift
                    End If
                Else
                    ' Start one
                    Purge_Start iStation, iShift
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

Public Sub Purge_Start(ByVal iStation As Integer, ByVal iShift As Integer)
'
' Start of a purge cycle needs (iStation, iShift) valves and Mass Air Controllers set up properly
'
'
'******************************************************************************
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 371
Dim Idx As Integer
Dim iAuxOut As Integer
Dim sMsg As String
Dim sMsg1 As String
Dim sMsg2 As String

    If StationControl(iStation, iShift).Mode = VBPURGEWAIT Then     'Start it
    
        If StationRecipe(iStation, iShift).Purge_Method <> NOPURGE Then
        
            StationControl(iStation, iShift).IsPausedInAlarm = False
            StationControl(iStation, iShift).AlarmDelayTime = VALUE0
                        
            ' reset the bar graph (also resets Actual & Target)
            Reset_Bar_Graph iStation, iShift
            
            ' Set Purge Target
            Select Case StationRecipe(iStation, iShift).Purge_Method
                Case PURGEBYTIME
                    ' PURGE BY TIME
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                    StationControl(iStation, iShift).Target = StationRecipe(iStation, iShift).Purge_Time
                Case PURGEAUXONLY
                    ' PURGE AUX CANISTER ONLY
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                    StationControl(iStation, iShift).Target = StationRecipe(iStation, iShift).Purge_AuxTime
                Case PURGEBYPROFILE
                    ' PURGE BY PROFILE
                    PurgeControl(iStation, iShift).curStep = 1
                    PurgeControl(iStation, iShift).ProfileStartDTS = Now()
                    PurgeControl(iStation, iShift).ProfileElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).ProfileElapsedSeconds = CLng(0)
                    PurgeControl(iStation, iShift).StepStartDTS = Now()
                    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
                    PurgeControl(iStation, iShift).CompletedStepMinutes = CSng(0)
                    ' initial Purge MFC set point
                    PurgeControl(iStation, iShift).CurMfcSp = StationProfile(iStation, iShift).StepStartSetpoint(1)
                    PurgeControl(iStation, iShift).InhibitOotCheck = True
                    StationControl(iStation, iShift).Target = CSng(StationProfile(iStation, iShift).Duration)
                Case PURGEBYLITERS
                    ' PURGE BY LITERS
                    ' purge target weight in grams
                    PurgeControl(iStation, iShift).PurgeTargetWt = Scale_Weight(StationRecipe(iStation, iShift).PriScaleNo) - StationCanister(iStation, iShift).WorkingCapacity
                    PurgeControl(iStation, iShift).curCycle = 1
                    PurgeControl(iStation, iShift).curStep = 1
                    PurgeControl(iStation, iShift).StepStartDTS = Now()
                    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
                    ' initial Purge MFC set point
                    PurgeControl(iStation, iShift).CurMfcSp = StationRecipe(iStation, iShift).Purge_Flow
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                    ' station target is Multiple Canister Volumes
                    StationControl(iStation, iShift).Target = StationRecipe(iStation, iShift).Purge_Liters
                Case PURGEBYVOLUME
                    ' PURGE BY VOLUME
                    ' purge target weight in grams
                    PurgeControl(iStation, iShift).PurgeTargetWt = Scale_Weight(StationRecipe(iStation, iShift).PriScaleNo) - StationCanister(iStation, iShift).WorkingCapacity
                    PurgeControl(iStation, iShift).curCycle = 1
                    PurgeControl(iStation, iShift).curStep = 1
                    PurgeControl(iStation, iShift).StepStartDTS = Now()
                    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
                    ' initial Purge MFC set point
                    PurgeControl(iStation, iShift).CurMfcSp = StationRecipe(iStation, iShift).Purge_Flow
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                    ' station target is Multiple Canister Volumes
                    StationControl(iStation, iShift).Target = StationRecipe(iStation, iShift).Purge_Can_Vol
                Case PURGEBYWC
                    ' PURGE BY WORKING CAPACITY
                    ' purge target weight in grams
                    PurgeControl(iStation, iShift).PurgeTargetWt = Scale_Weight(StationRecipe(iStation, iShift).PriScaleNo) - (0.01 * StationRecipe(iStation, iShift).Purge_TargetWC * StationCanister(iStation, iShift).WorkingCapacity)
                    PurgeControl(iStation, iShift).curCycle = 1
                    PurgeControl(iStation, iShift).curStep = 1
                    PurgeControl(iStation, iShift).StepStartDTS = Now()
                    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
                    ' initial Purge MFC set point
                    PurgeControl(iStation, iShift).CurMfcSp = StationRecipe(iStation, iShift).Purge_Flow
                    ' station target is Required Weight Change
                    StationControl(iStation, iShift).Target = StationRecipe(iStation, iShift).Purge_TargetWC
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                Case PURGETOTARGET
                    ' PURGE TO TARGET WEIGHT
                    ' purge target weight in grams
                    PurgeControl(iStation, iShift).PurgeTargetWt = StationRecipe(iStation, iShift).Purge_TargetWeight
                    PurgeControl(iStation, iShift).curCycle = 1
                    PurgeControl(iStation, iShift).curStep = 1
                    PurgeControl(iStation, iShift).StepStartDTS = Now()
                    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
                    ' initial Purge MFC set point
                    PurgeControl(iStation, iShift).CurMfcSp = StationRecipe(iStation, iShift).Purge_Flow
                    ' station target is Required Weight Change
                    StationControl(iStation, iShift).Target = Scale_Weight(StationRecipe(iStation, iShift).PriScaleNo) - StationRecipe(iStation, iShift).Purge_TargetWeight
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                Case PURGETOUNDOLOAD
                    ' PURGE TO UNDO LOAD
                    ' purge target weight in grams
                    If ((StationRecipe(iStation, iShift).CycleType = CyclePurgeLoad) And (StationControl(iStation, iShift).CurrCycle = CInt(1))) Then
                        ' force purge until timeout on first cycle (of PurgeLoad; 1st Purge of LoadPurge is OK)
                        PurgeControl(iStation, iShift).PurgeTargetWt = Scale_Weight(StationRecipe(iStation, iShift).PriScaleNo) - (CSng(2) * StationCanister(iStation, iShift).WorkingCapacity)
                    Else
                        ' purge to undo the last Load
                        PurgeControl(iStation, iShift).PurgeTargetWt = Scale_Weight(StationRecipe(iStation, iShift).PriScaleNo) - LoadControl(iStation, iShift).PriWtChg
                    End If
                    PurgeControl(iStation, iShift).curCycle = 1
                    PurgeControl(iStation, iShift).curStep = 1
                    PurgeControl(iStation, iShift).StepStartDTS = Now()
                    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
                    ' initial Purge MFC set point
                    PurgeControl(iStation, iShift).CurMfcSp = StationRecipe(iStation, iShift).Purge_Flow
                    ' station target is Required Weight Change
                    If (StationControl(iStation, iShift).CurrCycle = CInt(1)) Then
                        ' force purge until timeout on first cycle
                        StationControl(iStation, iShift).Target = CSng(2) * StationCanister(iStation, iShift).WorkingCapacity
                    Else
                        ' purge to undo the last Load
                         StationControl(iStation, iShift).Target = LoadControl(iStation, iShift).PriWtChg
                    End If
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                Case Else
                    ' undefined
                    PurgeControl(iStation, iShift).InhibitOotCheck = False
                    StationControl(iStation, iShift).Target = 0
            End Select
            
            ' Reset Totals
            PurgeControl(iStation, iShift).Purge_Total = CSng(0)
            PurgeControl(iStation, iShift).Purge_Volumes = CSng(0)
            PurgeControl(iStation, iShift).AuxWtChg = CSng(0)
            PurgeControl(iStation, iShift).PriWtChg = CSng(0)
            PurgeControl(iStation, iShift).TotalWtChg = CSng(0)
            PurgeControl(iStation, iShift).TotalWtChgRate = CSng(0)
            
            
            ' Set Purge Target (for statistics)
            PurgeControl(iStation, iShift).Purge_Target = StationControl(iStation, iShift).Target
            Clear_Stats iStation, iShift, 3               ' Clear old stats, if exist, for both load and purge
    
            
            ' Clear PurgeData
            StnPurgeData(iStation, iShift, 0) = BlankPurgeData
            StnPurgeData(iStation, iShift, 1) = BlankPurgeData
            StnPurgeData(iStation, iShift, 2) = BlankPurgeData
            StnPurgeData(iStation, iShift, 3) = BlankPurgeData
            StnPurgeData(iStation, iShift, 4) = BlankPurgeData
            StnPurgeData(iStation, iShift, 5) = BlankPurgeData
            StnPurgeData(iStation, iShift, 6) = BlankPurgeData
            StnPurgeData(iStation, iShift, 7) = BlankPurgeData
            StnPurgeData(iStation, iShift, 8) = BlankPurgeData
            StnPurgeData(iStation, iShift, 9) = BlankPurgeData
    
            tarmin(iStation, iShift) = 0
            tarvol(iStation, iShift) = 0
            tardone(iStation, iShift) = False
            scaleFlag(iStation) = False
            
            
            Stn_Purge_Log_TestTimer(iStation, iShift) = StationControl(iStation, iShift).TestTimer
            PreviousReportTimer(iStation, iShift) = StationControl(iStation, iShift).TestTimer
            PreviousTotalTimer(iStation, iShift) = StationControl(iStation, iShift).TestTimer
            
            
            LastPurgeStart = Now()                      ' Next purge cannot start for five min (donot want to overload exhaust header)
            PurgeControl(iStation, iShift).StartTime = Now()
            PurgeControl(iStation, iShift).StartTimer = StationControl(iStation, iShift).TestTimer
            PreviousNow(iStation, iShift) = Now
            StationControl(iStation, iShift).Mode_StartDts = Now        ' Since just starting purge, set mode and mode start timer
            StationControl(iStation, iShift).Mode = VBPURGE
            ChgPhase PurgeStarting, Now, iStation, iShift
            
            Purge_Write iStation, iShift, PURGEBEGIN
            
            ' Reset "MFC SetPoint is Set" flag
            '   note: SP is actually set in Purge_Check routine
            Stn_MfcSpIsSet(iStation) = False
    
            ' Continue to Request Station Aspirator
            PRG_INFO(STN_INFO(iStation).AspiratorNum).RequestRun = True
    
            ' Continue to Request Aspirator for Aux Scale (if needed)
            If StationRecipe(iStation, iShift).PurgeAuxCan And (StationRecipe(iStation, iShift).AuxScaleNo > 0) Then
                ' Only allowed to purge Aux Scale with a Vacuum Purge
                If Not StationConfig(iStation, iShift).PosPressPurge Then
                    If ((STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum < 1) Or (STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum > NR_PRGAIR)) Then
                        If (Not scaleFlag(iStation)) Then
                            scaleFlag(iStation) = True
                            sMsg1 = "Station " & Format(iStation, "##0") & " Shift " & Format(iShift, "##0")
                            sMsg2 = "Aux Scale #" & Format(StationRecipe(iStation, iShift).AuxScaleNo, "##0") & " has no owner"
                            sMsg = sMsg1 & " - " & sMsg2
                            Write_ELog sMsg
                        End If
                    Else
                        scaleFlag(iStation) = False
                        PRG_INFO(STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum).RequestRun = True
                    End If
                End If
            End If
                                        
            '
            ' open purge valves
            '
            Purge_Valves iStation, iShift
    
            ChgPhase PurgePurging, Now, iStation, iShift
                            
        Else
            ' Must be a load only
            PreLoad_Start iStation, iShift
        End If
    
    Else                                            ' station mode is not VBPURGEWAIT
        StationControl(iStation, iShift).Mode = VBPURGEWAIT
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

Public Sub Purge_ContinueDelayed(ByVal iStation As Integer, ByVal iShift As Integer)
 
 Dim count1 As Integer
 Dim count2 As Integer
 Dim timenow As Date
 Dim deltaMin As Long
 Dim SomebodyPurging As Integer
 Dim LeakInterlock As Boolean
 
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3772


    ' Continue to Request Aspirator for Aux Scale (if needed)
    If StationRecipe(iStation, iShift).PurgeAuxCan And (StationRecipe(iStation, iShift).AuxScaleNo > 0) Then
        ' Only allowed to purge Aux Scale with a Vacuum Purge
        If Not StationConfig(iStation, iShift).PosPressPurge Then
            PRG_INFO(STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum).RequestRdy = True
        End If
    End If
                                
    ' Continue to Request Station Aspirator
    PRG_INFO(STN_INFO(iStation).AspiratorNum).RequestRdy = True
    
    If Pause_Alarm = NOTPAUSED Then
        timenow = Now()
        SomebodyPurging = 0
        LeakInterlock = False
        For count1 = 1 To LAST_STN
           For count2 = 1 To NR_SHIFT
             If StationControl(count1, count2).Mode = VBLEAK Then
                If (USINGPRESSUREPURGE And SysConfig.PosPressPurge) Then
                    LeakInterlock = True
                End If
             End If
             ' Anyone Started Purging within the last 5 minutes ?
             If StationControl(count1, count2).Mode = VBPURGE And PurgeControl(count1, count2).Phase < PurgePause Then
                deltaMin = DateDiff("n", PurgeControl(count1, count2).StartTime, timenow)
                If deltaMin < 5 Then SomebodyPurging = SomebodyPurging + 1
             End If
           Next count2
        Next count1
    End If
    
    
    '   OK to Continue Purge ?
    If LeakInterlock = False Then
        ' No Leak Interlock
        If SomebodyPurging = 0 Or (timenow >= LastPurgeStart + TimeSerial(0, 5, 0)) Then
            ' No purges running or five minutes has elapsed since most recent station started purging
            If PRG_INFO(STN_INFO(iStation).AspiratorNum).Ready Then
                ' Resume(Continue) Purge
                Purge_Continue iStation, iShift
            Else
                ' continue waiting
                StationControl(iStation, iShift).Mode = VBPURGECONT
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

Public Sub Purge_Continue(ByVal iStation As Integer, ByVal iShift As Integer)
'
' Continue of a purge cycle after an alarm
' needs valves and Mass Air Controllers set up properly
'
'   written 8 Mar 2005
'
'******************************************************************************
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3711
Dim cnt1 As Long
Dim Idx As Integer
Dim iAuxOut As Integer
Dim sMsg As String
Dim sMsg1 As String
Dim sMsg2 As String
  
    If StationRecipe(iStation, iShift).UseAuxScale Then
        Scale_In_Use(StationRecipe(iStation, iShift).AuxScaleNo) = True
    End If
    If StationRecipe(iStation, iShift).UsePriScale Then
        Scale_In_Use(StationRecipe(iStation, iShift).PriScaleNo) = True
    End If
                
    
    ' Reset "MFC SetPoint is Set" flag
    '   note: SP is actually set in Purge_Check routine
    Stn_MfcSpIsSet(iStation) = False
    
    ' Continue to Request Aspirator for Aux Scale (if needed)
    If StationRecipe(iStation, iShift).PurgeAuxCan And (StationRecipe(iStation, iShift).AuxScaleNo > 0) Then
        ' Only allowed to purge Aux Scale with a Vacuum Purge
        If Not StationConfig(iStation, iShift).PosPressPurge Then
            If ((STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum < 1) Or (STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum > NR_PRGAIR)) Then
                If (Not scaleFlag(iStation)) Then
                    scaleFlag(iStation) = True
                    sMsg1 = "Station " & Format(iStation, "##0") & " Shift " & Format(iShift, "##0")
                    sMsg2 = "Aux Scale #" & Format(StationRecipe(iStation, iShift).AuxScaleNo, "##0") & " has no owner"
                    sMsg = sMsg1 & " - " & sMsg2
                    Write_ELog sMsg
                End If
            Else
                scaleFlag(iStation) = False
                PRG_INFO(STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum).RequestRun = True
            End If
        End If
    End If
                                        
    ' Continue to Request Station Aspirator
    PRG_INFO(STN_INFO(iStation).AspiratorNum).RequestRun = True
    
    ' Purge Valves
    Purge_Valves iStation, iShift
    
    
    ' Adjust Total Time in Alarm
    cnt1 = CLng(Second(Now - StationControl(iStation, iShift).PauseAlarmStartTime)) + CLng(Minute(Now - StationControl(iStation, iShift).PauseAlarmStartTime) * 60) + CLng(Hour(Now - StationControl(iStation, iShift).PauseAlarmStartTime) * 3600)
    StationControl(iStation, iShift).AlarmDelayTime = StationControl(iStation, iShift).AlarmDelayTime + cnt1    ' in seconds
    
    ' Reset Report & Totalize Timers
    Stn_Purge_Log_TestTimer(iStation, iShift) = StationControl(iStation, iShift).TestTimer                             ' do next Totalize after one SysConfig.PurgeTotal_Interval
    PreviousReportTimer(iStation, iShift) = StationControl(iStation, iShift).TestTimer                                 ' reset report timer
    PreviousTotalTimer(iStation, iShift) = StationControl(iStation, iShift).TestTimer                                  ' reset totalize timer
    
    ' Reset Station Mode
    StationControl(iStation, iShift).Mode = VBPURGE
    ChgPhase PurgePurging, Now, iStation, iShift
    PurgeControl(iStation, iShift).StepStartDTS = Now()
    PurgeControl(iStation, iShift).StepElapsedMinutes = CSng(0)
    PurgeControl(iStation, iShift).StepElapsedSeconds = CLng(0)
    PurgeControl(iStation, iShift).CurMfcSp = StationProfile(iStation, iShift).StepStartSetpoint(PurgeControl(iStation, iShift).curStep)
    
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

Public Sub Purge_Valves(station As Integer, Shift As Integer)
'
' Opens valves for Purge
'
'   written 9 April 2018
'
'******************************************************************************
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1137
Dim cnt1 As Long
Dim Idx As Integer
Dim iAuxOut As Integer
Dim sMsg As String
  
    ' Common Purge-With-Dry-Air valves
    ' note: these valves are turned off in Check_Stations if all stations are idle
    If (USINGDRYPURGEAIR And StationConfig(station, Shift).DryAirPurge) Then
        Com_OutDigital icPurgeDryAirSupplySol, cON
        Com_OutDigital icPurgeAirSourceSelectSol, cON
    End If
    
    If StationRecipe(station, Shift).Purge_Method <> PURGEAUXONLY Then
        ' station direction valve on
        Stn_OutDigital station, isPriDirectionSol, cON
        ' station purge flow valve on
        Stn_OutDigital station, isPurgeSol, cON
        ' Purge Canisters in Series ?
        If (USINGPURGESERIES And (StationRecipe(station, Shift).PurgeCansInSeries)) Then
            ' Only allowed to purge-in-series with a vacuum purge
            If (Not StationConfig(station, Shift).PosPressPurge) Then
                ' inseries valves
                Stn_OutDigital station, isAuxSeriesPurgeSol, cON
                Stn_OutDigital station, isPriSeriesPurgeSol, cON
            Else
                 sMsg = "Cannot Purge In Series because Config is set to PositivePressurePurge"
                 Write_ELog sMsg
'                 MsgBox sMsg, vbInformation, "DEBUG INFO"
            End If
        End If
        ' Purge-In-Oven valves
        If (USINGPURGEOVEN And (StationRecipe(station, Shift).PurgeOven)) Then
            Stn_OutDigital station, isPurgeLocationSupplySelectSol, cON
            Stn_OutDigital station, isPurgeLocationVentSelectSol, cON
        End If
        ' Shift valves
        Select Case Shift
            Case 1
                ' nothing to do
            Case 2
                Stn_OutDigital station, isLoadShift2Sol, cOFF
                Stn_OutDigital station, isPurgeShift2Sol, cON
                Stn_OutDigital station, isVentShift2Sol, cON
            Case 3
                Stn_OutDigital station, isLoadShift2Sol, cOFF
                Stn_OutDigital station, isPurgeShift2Sol, cON
                Stn_OutDigital station, isVentShift2Sol, cON
                Stn_OutDigital station, isLoadShift3Sol, cOFF
                Stn_OutDigital station, isPurgeShift3Sol, cON
                Stn_OutDigital station, isVentShift3Sol, cON
            Case 4
                Stn_OutDigital station, isLoadShift2Sol, cOFF
                Stn_OutDigital station, isPurgeShift2Sol, cON
                Stn_OutDigital station, isVentShift2Sol, cON
                Stn_OutDigital station, isLoadShift4Sol, cOFF
                Stn_OutDigital station, isPurgeShift4Sol, cON
                Stn_OutDigital station, isVentShift4Sol, cON
        End Select
        ' PriAux Vent Valve
'        Delay_Box ("PriScaleStn = " & Format(StationControl(station, Shift).PriScaleStn, "##0")), MSGDELAY, msgSHOW
        If StationRecipe(station, Shift).UsePriScale And StationControl(station, Shift).PriScaleStn > 0 _
          And StationControl(station, Shift).PriScaleStn < FIRST_REMOTESCALE Then
            Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cON  ' Diff station valve
        End If
    End If
    ' Aux Direction & Aux Purge Valves
    If ((StationRecipe(station, Shift).PurgeAuxCan) And (StationControl(station, Shift).AuxScaleStn > 0)) Then
'        sMsg = "Purge Aux Canister"
'        MsgBox sMsg, vbInformation, "DEBUG INFO"
        ' Only allowed to purge Aux Can with a Vacuum Purge
        If (Not StationConfig(station, Shift).PosPressPurge) Then
'            Delay_Box ("AuxScaleStn = " & Format(StationControl(station, Shift).AuxScaleStn, "##0")), MSGDELAY, msgSHOW
             ' Purge Canisters in Series ?
            If ((USINGPURGESERIES) And (StationRecipe(station, Shift).PurgeCansInSeries)) Then
'                sMsg = "Purge In Series"
'                MsgBox sMsg, vbInformation, "DEBUG INFO"
                ' Aux Direction Valve
'                sMsg = "Open Station#" & Format(StationControl(station, Shift).AuxScaleStn, "##0") & " Aux Direction Valve"
'                MsgBox sMsg, vbInformation, "DEBUG INFO"
                Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxDirectionSol, cON
                ' Aux Purge Valve
'                sMsg = "Close Station#" & Format(StationControl(station, Shift).AuxScaleStn, "##0") & " Aux Purge Valve"
'                MsgBox sMsg, vbInformation, "DEBUG INFO"
                Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxPurgeSol, cOFF
            Else
'                sMsg = "Purge In Parallel"
'                MsgBox sMsg, vbInformation, "DEBUG INFO"
                ' Aux Direction Valve
'                sMsg = "Close Station#" & Format(StationControl(station, Shift).AuxScaleStn, "##0") & " Aux Direction Valve"
'                MsgBox sMsg, vbInformation, "DEBUG INFO"
                Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxDirectionSol, cOFF
                ' Aux Purge Valve
'                sMsg = "Open Station#" & Format(StationControl(station, Shift).AuxScaleStn, "##0") & " Aux Purge Valve"
'                MsgBox sMsg, vbInformation, "DEBUG INFO"
                Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxPurgeSol, cON
            End If
        Else
             sMsg = "Cannot Purge Aux Canister because Config is set to PositivePressurePurge"
             Write_ELog sMsg
'             MsgBox sMsg, vbInformation, "DEBUG INFO"
        End If
'    Else
'        If ((Not StationRecipe(station, Shift).PurgeAuxCan) And (Not StationControl(station, Shift).AuxScaleStn > 0)) Then
'            MsgBox "Recipe is not set to Purge the Aux Canister AND Aux Scale# = 0", vbInformation, "DEBUG INFO"
'        ElseIf (Not StationRecipe(station, Shift).PurgeAuxCan) Then
'            MsgBox "Recipe is not set to Purge the Aux Canister", vbInformation, "DEBUG INFO"
'        ElseIf (StationControl(station, Shift).AuxScaleStn = 0) Then
'            MsgBox "Aux Scale# = 0", vbInformation, "DEBUG INFO"
'        End If
    End If
    ' aux outputs
    If (USING_AUX_OUTPUTS And StationRecipe(station, Shift).AuxOutputs) Then
        For Idx = 1 To 4
            If (Idx <= NR_AUX_OUTPUTS) Then
                If (StationRecipe(station, Shift).AuxOutputs_Purge(Idx)) Then
                    iAuxOut = isAuxOutput1 + Idx - 1
                    Stn_OutDigital station, iAuxOut, cON
                End If
            End If
        Next Idx
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

Public Sub Purge_Done(ByVal iStation As Integer, ByVal iShift As Integer)
'
' End of a purge cycle needs (Index, index2) valves and Mass Air Controllers shutdown properly
'
'   rewritten 4 Mar 2005
'
'******************************************************************************
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 379
Dim inc As Integer
Dim inc2 As Integer
Dim tmpWt As Single
  
    ' increment the number of Completed Purges
    StationControl(iStation, iShift).CompletedPurges = StationControl(iStation, iShift).CompletedPurges + 1
    ' did a Purge actually happen ??
    If StationRecipe(iStation, iShift).Purge_Method <> NOPURGE Then
        ' Remember Primary Scale Value at End of Purge
        If StationRecipe(iStation, iShift).UsePriScale Then
            PurgeControl(iStation, iShift).PriWt_End = StationControl(iStation, iShift).PriScaleWt
        Else
            PurgeControl(iStation, iShift).PriWt_End = 0
        End If
        ' Remember Aux Scale Value at End of Purge
        If StationRecipe(iStation, iShift).UseAuxScale Then
            PurgeControl(iStation, iShift).AuxWt_End = StationControl(iStation, iShift).AuxScaleWt
        Else
            PurgeControl(iStation, iShift).AuxWt_End = 0
        End If
        StationControl(iStation, iShift).End_Time = Now
        StationControl(iStation, iShift).End_Timer = StationControl(iStation, iShift).TestTimer
        StationControl(iStation, iShift).IsPausedInAlarm = False
        FirstTime(iStation, iShift) = False               ' clear OOT FirstTime Flag
        scaleFlag(iStation) = False
        Purge_Write iStation, iShift, PURGEDONE
        Stats_Write iStation, iShift
        Reset_Bar_Graph iStation, iShift                  ' reset the bar graphs
    End If
            
            
    ' WHAT TO DO NEXT ??
    ' Need to PauseAfterPurge ??
    If StationRecipe(iStation, iShift).PauseAfterPurge Then
        ' PauseAfterPurge
        Pause_AfterPurge_Start iStation, iShift
    ElseIf StationRecipe(iStation, iShift).PauseAfterPurgeForOper Then
        ' PauseAfterPurgeForOperator
        Pause_AfterPurgeForOper_Start iStation, iShift
    ElseIf (StationControl(iStation, iShift).CompletedPurges = StationControl(iStation, iShift).CompletedLoads) Then
        ' cycle is complete
        StationControl(iStation, iShift).CompletedCycles = StationControl(iStation, iShift).CompletedCycles + 1
        ' net scale end weight
        tmpWt = CSng(0)
        If StationRecipe(iStation, iShift).UsePriScale Then tmpWt = tmpWt + StationControl(iStation, iShift).PriScaleWt
        If StationRecipe(iStation, iShift).UseAuxScale Then tmpWt = tmpWt + StationControl(iStation, iShift).AuxScaleWt
        StationCycleWeightData(iStation, iShift, StationControl(iStation, iShift).CompletedCycles).Cycle_EndWeight_Total = tmpWt
        ' is recipe complete ??
        If RecipeIsDone(iStation, iShift) Then
            ' recipe completed OK
            JobInfo(iStation, iShift).End_OK = True
            ' any more LiveFuel Loads ??
            If Not AnyMoreLiveFuelLoads(iStation, iShift) Then
                ' No More LiveFuel Loads; using AutoDrainFill ??
                If (AdfControl(iStation).AdfDefinition.hasAUTODRAINFILL And AdfControl(iStation).LiveFuel And AdfControl(iStation).LiveFuelChgAuto) Then
                    ' Got to Empty the Live Fuel Tank
                    AdfControl(iStation).Mode = 1
                    AdfControl(iStation).Step = 0
                Else
                    AdfControl(iStation).Mode = 0
                    AdfControl(iStation).Step = 0
                End If
            End If
            ' *****************************************
            ' What's Next ??
            Course_Next iStation, iShift
            ' *****************************************
        Else
            ' recipe is not complete; new cycle needs a load
            StationControl(iStation, iShift).CurrCycle = StationControl(iStation, iShift).CurrCycle + 1
            ' new cycle start weight = last cycle end weight
            StationCycleWeightData(iStation, iShift, StationControl(iStation, iShift).CurrCycle).Cycle_StartWeight_Total = _
                StationCycleWeightData(iStation, iShift, StationControl(iStation, iShift).CompletedCycles).Cycle_EndWeight_Total
            If StationRecipe(iStation, iShift).Load_Method = NOLOAD Then
                ' End a Load
                Load_Done iStation, iShift
            Else
                ' Start a Load
                PreLoad_Start iStation, iShift
            End If
        End If
    Else
        ' cycle is not complete; current cycle needs a load
        If StationRecipe(iStation, iShift).Load_Method = NOLOAD Then
            ' End a Load
            Load_Done iStation, iShift
        Else
            ' Start a Load
            PreLoad_Start iStation, iShift
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

Public Sub Purge_Abort(ByVal iStation As Integer, ByVal iShift As Integer)
'
' Stop of a purge cycle needs (iStation, iShift) valves and Mass Air Controllers shutdown properly
'
'   rewritten 4 Mar 2005
'
'******************************************************************************
      
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 376
Dim iCourse As Integer
Dim sPrint As String
               
    
    ' time stamp end of current course
    iCourse = StationControl(iStation, iShift).Course
    StationSequence(iStation, iShift).CourseData(iCourse).DtsEnd = Now()
    StationControl(iStation, iShift).End_Time = Now
    StationControl(iStation, iShift).End_Timer = StationControl(iStation, iShift).TestTimer
    ChgPhase PurgeStopping, Now, iStation, iShift
      
    ' Remember Primary Scale Value at End of Purge
    If StationRecipe(iStation, iShift).UsePriScale Then
        PurgeControl(iStation, iShift).PriWt_End = StationControl(iStation, iShift).PriScaleWt
    Else
        PurgeControl(iStation, iShift).PriWt_End = 0
    End If
    ' Remember Aux Scale Value at End of Purge
    If StationRecipe(iStation, iShift).UseAuxScale Then
        PurgeControl(iStation, iShift).AuxWt_End = StationControl(iStation, iShift).AuxScaleWt
    Else
        PurgeControl(iStation, iShift).AuxWt_End = 0
    End If
    
    ' Update Header data in data file
    Header_Update iStation, iShift
    Purge_Write iStation, iShift, PURGEDONE
    Stats_Write iStation, iShift
    ' Write CycleWeights data in data file
    Weights_Write iStation, iShift
      
    ALM_Write iStation, iShift, "Purge Cycle Aborted"
    Write_ELog "Purge cycle aborted " & "  Station " & iStation & "  Shift" & iShift
    FirstTime(iStation, iShift) = False               ' clear OOT FirstTime Flag
    scaleFlag(iStation) = False
    StationControl(iStation, iShift).IsPausedInAlarm = False
    Reset_Bar_Graph iStation, iShift                  ' reset the bar graph
 
    ' Shut Off IO
    '   Station Valves
    Close_Stn_Valves iStation, iShift
    '   Scale Valves
    If StationRecipe(iStation, iShift).UsePriScale And StationControl(iStation, iShift).PriScaleStn > 0 _
            And StationControl(iStation, iShift).PriScaleStn < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(iStation, iShift).PriScaleStn, isPriAuxVentSol, cOFF
    End If
    If StationRecipe(iStation, iShift).PurgeAuxCan And StationControl(iStation, iShift).AuxScaleStn > 0 Then
        Stn_OutDigital StationControl(iStation, iShift).AuxScaleStn, isAuxPurgeSol, cOFF
    End If
    
    ' write the logs and close off the world
    Station_Finish iStation, iShift
    
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

Public Sub PurgeController(ByVal iStn As Integer, ByVal iShift As Integer)
'
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 3376
Dim fractionOfStep As Single
Dim deltaSP As Single
Dim profSP As Single
Dim iStep As Integer
               
    ' Control Purge Sequence
    Select Case StationRecipe(iStn, iShift).Purge_Method
        Case PURGEBYTIME
            ' PURGE BY TIME
        Case PURGEAUXONLY
            ' PURGE AUX CANISTER ONLY
        Case PURGEBYPROFILE
            ' PURGE BY PROFILE
            PurgeControl(iStn, iShift).ProfileElapsedSeconds = DateDiff("s", PurgeControl(iStn, iShift).ProfileStartDTS, Now())
            PurgeControl(iStn, iShift).ProfileElapsedMinutes = CSng(PurgeControl(iStn, iShift).ProfileElapsedSeconds) / CSng(60)
            PurgeControl(iStn, iShift).StepElapsedSeconds = DateDiff("s", PurgeControl(iStn, iShift).StepStartDTS, Now())
            PurgeControl(iStn, iShift).StepElapsedMinutes = CSng(PurgeControl(iStn, iShift).StepElapsedSeconds) / CSng(60)
            ' purge profile step control
'            For iStep = 1 To MAX_PROFILESTEPS
'                If (iStep = PurgeControl(iStn, iShift).curStep) Then
                iStep = PurgeControl(iStn, iShift).curStep
                If (StationProfile(iStn, iShift).StepType(iStep) = STEPLAST) Then
                    ' PURGE BY PROFILE IS DONE
                    PurgeControl(iStn, iShift).CurMfcSp = StationProfile(iStn, iShift).StepStartSetpoint(iStep)
'                ElseIf (PurgeControl(iStn, iShift).StepElapsedSeconds >= (CSng(60) * StationProfile(iStn, iShift).StepDuration(iStep))) Then
                ElseIf (PurgeControl(iStn, iShift).ProfileElapsedMinutes >= (PurgeControl(iStn, iShift).CompletedStepMinutes + StationProfile(iStn, iShift).StepDuration(iStep))) Then
                    ' starting a new step
                    PurgeControl(iStn, iShift).curStep = IIf((iStep < MAX_PROFILESTEPS), iStep + 1, iStep)
                    PurgeControl(iStn, iShift).StepStartDTS = Now()
                    PurgeControl(iStn, iShift).StepElapsedMinutes = CSng(0)
                    PurgeControl(iStn, iShift).StepElapsedSeconds = CLng(0)
                    PurgeControl(iStn, iShift).CompletedStepMinutes = PurgeControl(iStn, iShift).CompletedStepMinutes + StationProfile(iStn, iShift).StepDuration(iStep)
                    PurgeControl(iStn, iShift).CurMfcSp = StationProfile(iStn, iShift).StepStartSetpoint(PurgeControl(iStn, iShift).curStep)
                    Stn_MfcSpIsSet(iStn) = False
                Else
                    ' just update the MFC SP
                    Select Case StationProfile(iStn, iShift).StepType(iStep)
                        Case STEPSTEP
                            PurgeControl(iStn, iShift).CurMfcSp = StationProfile(iStn, iShift).StepStartSetpoint(iStep)
                        Case STEPRAMP
                            fractionOfStep = 0
                            If (StationProfile(iStn, iShift).StepDuration(iStep) > 0) Then
                                fractionOfStep = PurgeControl(iStn, iShift).StepElapsedMinutes / StationProfile(iStn, iShift).StepDuration(iStep)
                            End If
                            deltaSP = StationProfile(iStn, iShift).StepStartSetpoint(iStep + 1) - StationProfile(iStn, iShift).StepStartSetpoint(iStep)
                            PurgeControl(iStn, iShift).CurMfcSp = StationProfile(iStn, iShift).StepStartSetpoint(iStep) + (deltaSP * fractionOfStep)
                            Stn_MfcSpIsSet(iStn) = False
                        Case STEPLAST
                            PurgeControl(iStn, iShift).CurMfcSp = StationProfile(iStn, iShift).StepStartSetpoint(iStep)
                    End Select
                End If
'                End If
'            Next iStep
        Case PURGEBYLITERS, PURGEBYVOLUME, PURGEBYWC, PURGETOTARGET, PURGETOUNDOLOAD
            ' PURGE BY LITERS & PURGE BY VOLUME & PURGE BY WORKING CAPACITY & PURGE TO TARGET WEIGHT & PURGE TO UNDO LOAD
            PurgeControl(iStn, iShift).StepElapsedSeconds = DateDiff("s", PurgeControl(iStn, iShift).StepStartDTS, Now())
            PurgeControl(iStn, iShift).StepElapsedMinutes = CSng(PurgeControl(iStn, iShift).StepElapsedSeconds) / CSng(60)
            Select Case StationRecipe(iStn, iShift).Purge_TargetMode
                Case TARGETPURGEPAUSE
                    ' PurgePauseRepeat
                    Select Case PurgeControl(iStn, iShift).curStep
                        Case 0
                            ' Pauseing
                            ' done with this step ?
                            If (PurgeControl(iStn, iShift).StepElapsedMinutes >= StationRecipe(iStn, iShift).Purge_TargetPause) Then
                                ' done with Purge ??
                                If StationControl(iStn, iShift).Actual >= StationControl(iStn, iShift).Target Then
                                    ' all done O.K.
                                    If Not StationControl(iStn, iShift).IsPausedInAlarm Then
                                        ' actual=target AND not paused for anything
                                        If (PurgeControl(iStn, iShift).Phase < PurgeComplete) Then ChgPhase PurgeComplete, Now, iStn, iShift
                                    End If
                                Else
                                    ' switch to purging
                                    PurgeControl(iStn, iShift).StepStartDTS = Now()
                                    PurgeControl(iStn, iShift).curStep = 1
                                    PurgeControl(iStn, iShift).CurMfcSp = CSng(StationRecipe(iStn, iShift).Purge_Flow)
                                    PurgeControl(iStn, iShift).InhibitOotCheck = False
                                    Stn_MfcSpIsSet(iStn) = False
                                    ' increment cycle counter
                                    PurgeControl(iStn, iShift).curCycle = PurgeControl(iStn, iShift).curCycle + 1
                                End If
                            End If
                        Case 1
                            ' Purging
                            ' done with this step OR reached Purge Target ???
                            If ( _
                                (PurgeControl(iStn, iShift).StepElapsedMinutes >= StationRecipe(iStn, iShift).Purge_TargetPurge) _
                                    Or _
                                (StationControl(iStn, iShift).Actual >= StationControl(iStn, iShift).Target) _
                                    ) Then
                                ' switch to pauseing
                                PurgeControl(iStn, iShift).StepStartDTS = Now()
                                PurgeControl(iStn, iShift).curStep = 0
                                PurgeControl(iStn, iShift).CurMfcSp = CSng(0)
                                PurgeControl(iStn, iShift).InhibitOotCheck = True
                                Stn_MfcSpIsSet(iStn) = False
                            End If
                    End Select
                Case TARGETCONTINUOUS
                    ' Continuous Purging
                    PurgeControl(iStn, iShift).CurMfcSp = CSng(StationRecipe(iStn, iShift).Purge_Flow)
            End Select
    
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

Public Sub PreLoad_Abort(station As Integer, Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 200
Dim iCourse As Integer
Dim sPrint As String

    ' time stamp end of current course
    iCourse = StationControl(station, Shift).Course
    StationSequence(station, Shift).CourseData(iCourse).DtsEnd = Now()
    StationControl(station, Shift).End_Time = Now
    StationControl(station, Shift).End_Timer = StationControl(station, Shift).TestTimer
      
    ' Update Header data in data file
    Header_Update station, Shift
    ' Write CycleWeights data in data file
    Weights_Write station, Shift
      
    StationControl(station, Shift).IsPausedInAlarm = False
    ALM_Write station, Shift, "PreLoad Aborted"
    Reset_Bar_Graph station, Shift              ' reset the bar graph
    FirstTime(station, Shift) = False           ' clear OOT FirstTime Flag   2 Mar 2005
    
    ' Close preload MFC
    Select Case STN_INFO(station).Type
    
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
        
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
            Else
                ' use lower range MFC
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
            End If
                
        Case STN_LIVEFUEL_TYPE
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
        
        Case STN_LIVEREG_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
            Else
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
            End If
        
        Case STN_LIVEORVR2_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
            Else
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                End If
            End If
        
        Case STN_COMBO3_TYPE
            ' future
        
        Case Else
      
    End Select
    
    ' Close Valves
    Close_Stn_Valves station, Shift
    If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
      And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
       Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF
    End If
    
    Write_ELog "PreLoad aborted " & "  Station " & station & "  Shift" & Shift
    ' write the logs and close off the world
    Station_Finish station, Shift
   

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

Public Sub PreLoad_Start(station As Integer, Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 201
Dim tempMin, tempSec As Integer
Dim Nitrogen_Rate As Single
Dim Nitrogen_Output As Single
Dim span As Single
    
    tempSec = StationConfig(station, Shift).NitrogenPurgeTime Mod 60
    tempMin = CInt((StationConfig(station, Shift).NitrogenPurgeTime - tempSec) / 60)
    StationControl(station, Shift).End_Time = Now() + TimeSerial(0, tempMin, tempSec)
    StationControl(station, Shift).Mode = VBPRELOAD
    ' open load valves (except butane valves)
    LoadValves_Open station, Shift
    ' Put MFC(s) into operation
    Select Case STN_INFO(station).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
            Nitrogen_Rate = CSng(0.9 * Stn_AIO(station, asNitrogenFlowSP).EuMax)
            If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
            ' set Nitrogen MFC setpoint
            span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
            Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
            Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFCs
                ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
                Nitrogen_Rate = CSng(0.9 * Stn_AIO(station, asNitrogenORVRFlowSP).EuMax)
                If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
                ' set Nitrogen MFC setpoint
                span = Stn_AIO(station, asNitrogenORVRFlowSP).EuMax - Stn_AIO(station, asNitrogenORVRFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asNitrogenORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRNIT, Stn_MfcCal(station, MFCORVRNIT)))
                Stn_OutAnalog station, asNitrogenORVRFlowSP, Nitrogen_Output, outNORMAL
            Else
                ' use lower range MFCs
                ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
                Nitrogen_Rate = CSng(0.9 * Stn_AIO(station, asNitrogenFlowSP).EuMax)
                If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
                ' set Nitrogen MFC setpoint
                span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
                Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
            End If
        Case STN_LIVEFUEL_TYPE
            ' Determine desired Nitrogen flow rates in SLPM
            Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
            ' set Live Fuel Vapor Carrier MFC setpoint
            span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
            Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
        Case STN_LIVEREG_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                ' Determine desired Nitrogen flow rates in SLPM
                Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
                ' set Live Fuel Vapor Carrier MFC setpoint
                span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
            Else
                ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
                Nitrogen_Rate = CSng(0.9 * Stn_AIO(station, asNitrogenFlowSP).EuMax)
                If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
                ' set Nitrogen MFC setpoint
                span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
                Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
            End If
        Case STN_LIVEORVR2_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                ' Determine desired LiveFuel Vapor Carrier flow rates in SLPM
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFCs
                    Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
                    ' set Live Fuel Vapor Carrier MFC setpoint
                    span = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRLIVE, Stn_MfcCal(station, MFCORVRLIVE)))
                    Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, Nitrogen_Output, outNORMAL
                Else
                    ' use lower range MFCs
                    Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
                    ' set Live Fuel Vapor Carrier MFC setpoint
                    span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
                End If
            Else
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFCs
                    ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
                    Nitrogen_Rate = CSng(0.9 * Stn_AIO(station, asNitrogenORVRFlowSP).EuMax)
                    If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
                    ' set Nitrogen MFC setpoint
                    span = Stn_AIO(station, asNitrogenORVRFlowSP).EuMax - Stn_AIO(station, asNitrogenORVRFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asNitrogenORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRNIT, Stn_MfcCal(station, MFCORVRNIT)))
                    Stn_OutAnalog station, asNitrogenORVRFlowSP, Nitrogen_Output, outNORMAL
                Else
                    ' use lower range MFCs
                    ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
                    Nitrogen_Rate = CSng(0.9 * Stn_AIO(station, asNitrogenFlowSP).EuMax)
                    If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
                    ' set Nitrogen MFC setpoint
                    span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
                    Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                End If
            End If
        Case STN_COMBO3_TYPE
            ' future
        Case Else
            ' do nothing
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

Public Sub PreLoad_Check(station As Integer, Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 202
    
    If Now() >= StationControl(station, Shift).End_Time Then
        PreLoad_Done station, Shift
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

Public Sub PreLoad_Done(station As Integer, Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler

    SetErrModule 2, 209
    ' shutdown stn mfc's
    ShutdownStnMFCs station, Shift
    ' close stn valves
    Close_Stn_Valves station, Shift
    ' need to wait for WaterBath ??
    If STN_INFO(station).ADF_DEF.hasADF_WaterBath Then
        StationControl(station, Shift).Mode = VBWBPAUSE
    Else
        ' start Load
        Load_Start station, Shift
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

Public Sub LoadSetPoint_Update(ByVal station As Integer, ByVal Shift As Integer)
'
' special routine to update MFC SP's during a Load Cycle
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 2299
Dim Butane_Rate As Single
Dim Nitrogen_Rate As Single
Dim Butane_Output As Single
Dim Nitrogen_Output As Single
Dim span As Single
    
    ' Update the MFC's SetPoint
    Select Case STN_INFO(station).Type
     
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
        
            ' Determine desired Butane & Nitrogen flow rates in SLPM
            ' Butane flow rate in SLPM
            Butane_Rate = CSng(GramsPerHourToSlpm(StationRecipe(station, Shift).Load_Rate, StationControl(station, Shift).BtnDensity))
            Stn_Btn_FlowSP(station, Shift) = Butane_Rate
            ' Nitrogen Flow rate in SLPM
            Nitrogen_Rate = CSng((100 - StationRecipe(station, Shift).Mix_Percent) * (Butane_Rate / StationRecipe(station, Shift).Mix_Percent))
            Stn_Nit_FlowSP(station, Shift) = Nitrogen_Rate
     
            ' set Butane MFC setpoint
            span = Stn_AIO(station, asButaneFlowSP).EuMax - Stn_AIO(station, asButaneFlowSP).EuMin
            Butane_Output = Stn_AIO(station, asButaneFlowSP).EuMin + (span * Cal_MfcOutput(Butane_Rate, station, MFCBUTANE, Stn_MfcCal(station, MFCBUTANE)))
            Stn_OutAnalog station, asButaneFlowSP, Butane_Output, outNORMAL
       
            ' set Nitrogen MFC setpoint
            span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
            Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
            Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
        
        
        Case STN_ORVR2_TYPE
        
            ' Determine desired Butane & Nitrogen flow rates in SLPM
            ' Butane flow rate in SLPM
            Butane_Rate = CSng(GramsPerHourToSlpm(StationRecipe(station, Shift).Load_Rate, StationControl(station, Shift).BtnDensity))
            Stn_Btn_FlowSP(station, Shift) = Butane_Rate
            ' Nitrogen flow rate in SLPM
            Nitrogen_Rate = CSng((100 - StationRecipe(station, Shift).Mix_Percent) * (Butane_Rate / StationRecipe(station, Shift).Mix_Percent))
            Stn_Nit_FlowSP(station, Shift) = Nitrogen_Rate
     
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFCs
                
                ' set Butane MFC setpoint
                span = Stn_AIO(station, asButaneORVRFlowSP).EuMax - Stn_AIO(station, asButaneORVRFlowSP).EuMin
                Butane_Output = Stn_AIO(station, asButaneORVRFlowSP).EuMin + (span * Cal_MfcOutput(Butane_Rate, station, MFCORVRBUT, Stn_MfcCal(station, MFCORVRBUT)))
                Stn_OutAnalog station, asButaneORVRFlowSP, Butane_Output, outNORMAL
                
                ' set Nitrogen MFC setpoint
                span = Stn_AIO(station, asNitrogenORVRFlowSP).EuMax - Stn_AIO(station, asNitrogenORVRFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asNitrogenORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRNIT, Stn_MfcCal(station, MFCORVRNIT)))
                Stn_OutAnalog station, asNitrogenORVRFlowSP, Nitrogen_Output, outNORMAL
            Else
                ' use lower range MFCs
                
                ' set Butane MFC setpoint
                span = Stn_AIO(station, asButaneFlowSP).EuMax - Stn_AIO(station, asButaneFlowSP).EuMin
                Butane_Output = Stn_AIO(station, asButaneFlowSP).EuMin + (span * Cal_MfcOutput(Butane_Rate, station, MFCBUTANE, Stn_MfcCal(station, MFCBUTANE)))
                Stn_OutAnalog station, asButaneFlowSP, Butane_Output, outNORMAL
                
                ' set Nitrogen MFC setpoint
                span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
                Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
            End If
        
        Case STN_LIVEFUEL_TYPE
        
            Butane_Rate = 0
            Stn_Btn_FlowSP(station, Shift) = 0
            Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
            Stn_Nit_FlowSP(station, Shift) = StationRecipe(station, Shift).NitrogenFlow  'percentage of flow
     
            span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
            Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
                        
        Case STN_LIVEREG_TYPE
        
            If StationRecipe(station, Shift).LiveFuel Then
            
                ' use Live Fuel vapor
                Butane_Rate = 0
                Stn_Btn_FlowSP(station, Shift) = 0
                Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                
                Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
                Stn_Nit_FlowSP(station, Shift) = StationRecipe(station, Shift).NitrogenFlow  'percentage of flow
                
                span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
                        
            Else
                ' use Butane/Nitrogen
                ' Determine desired Butane & Nitrogen flow rates in SLPM
                ' Butane flow rate in SLPM
                Butane_Rate = CSng(GramsPerHourToSlpm(StationRecipe(station, Shift).Load_Rate, StationControl(station, Shift).BtnDensity))
                Stn_Btn_FlowSP(station, Shift) = Butane_Rate
                ' Nitrogen Flow rate in SLPM
                Nitrogen_Rate = CSng((100 - StationRecipe(station, Shift).Mix_Percent) * (Butane_Rate / StationRecipe(station, Shift).Mix_Percent))
                Stn_Nit_FlowSP(station, Shift) = Nitrogen_Rate
                
                ' set Butane MFC setpoint
                span = Stn_AIO(station, asButaneFlowSP).EuMax - Stn_AIO(station, asButaneFlowSP).EuMin
                Butane_Output = Stn_AIO(station, asButaneFlowSP).EuMin + (span * Cal_MfcOutput(Butane_Rate, station, MFCBUTANE, Stn_MfcCal(station, MFCBUTANE)))
                Stn_OutAnalog station, asButaneFlowSP, Butane_Output, outNORMAL
                
                ' set Nitrogen MFC setpoint
                span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
                Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                
            End If
        
        Case STN_LIVEORVR2_TYPE
        
            If StationRecipe(station, Shift).LiveFuel Then
            
                ' use Live Fuel vapor
                Butane_Rate = 0
                Stn_Btn_FlowSP(station, Shift) = 0
                Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
                
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFCs
                    
                    ' set LiveFuel Vapor Carrier MFC setpoint
                    Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
                    Stn_Nit_FlowSP(station, Shift) = StationRecipe(station, Shift).NitrogenFlow  'percentage of flow
                    
                    span = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRLIVE, Stn_MfcCal(station, MFCORVRLIVE)))
                    Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, Nitrogen_Output, outNORMAL
                Else
                    ' use lower range MFCs
                    
                    ' set LiveFuel Vapor Carrier MFC setpoint
                    Nitrogen_Rate = CSng(StationRecipe(station, Shift).NitrogenFlow)
                    Stn_Nit_FlowSP(station, Shift) = StationRecipe(station, Shift).NitrogenFlow  'percentage of flow
                    
                    span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
                End If
                            
            Else
            
                ' use Butane/Nitrogen
                ' Determine desired Butane & Nitrogen flow rates in SLPM
                ' Butane flow rate in SLPM
                Butane_Rate = CSng(GramsPerHourToSlpm(StationRecipe(station, Shift).Load_Rate, StationControl(station, Shift).BtnDensity))
                Stn_Btn_FlowSP(station, Shift) = Butane_Rate
                ' Nitrogen flow rate in SLPM
                Nitrogen_Rate = CSng((100 - StationRecipe(station, Shift).Mix_Percent) * (Butane_Rate / StationRecipe(station, Shift).Mix_Percent))
                Stn_Nit_FlowSP(station, Shift) = Nitrogen_Rate
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFCs
                    ' set Butane MFC setpoint
                    span = Stn_AIO(station, asButaneORVRFlowSP).EuMax - Stn_AIO(station, asButaneORVRFlowSP).EuMin
                    Butane_Output = Stn_AIO(station, asButaneORVRFlowSP).EuMin + (span * Cal_MfcOutput(Butane_Rate, station, MFCORVRBUT, Stn_MfcCal(station, MFCORVRBUT)))
                    Stn_OutAnalog station, asButaneORVRFlowSP, Butane_Output, outNORMAL
                    
                    ' set Nitrogen MFC setpoint
                    span = Stn_AIO(station, asNitrogenORVRFlowSP).EuMax - Stn_AIO(station, asNitrogenORVRFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asNitrogenORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRNIT, Stn_MfcCal(station, MFCORVRNIT)))
                    Stn_OutAnalog station, asNitrogenORVRFlowSP, Nitrogen_Output, outNORMAL
                Else
                    ' use lower range MFCs
                    ' set Butane MFC setpoint
                    span = Stn_AIO(station, asButaneFlowSP).EuMax - Stn_AIO(station, asButaneFlowSP).EuMin
                    Butane_Output = Stn_AIO(station, asButaneFlowSP).EuMin + (span * Cal_MfcOutput(Butane_Rate, station, MFCBUTANE, Stn_MfcCal(station, MFCBUTANE)))
                    Stn_OutAnalog station, asButaneFlowSP, Butane_Output, outNORMAL
                    
                    ' set Nitrogen MFC setpoint
                    span = Stn_AIO(station, asNitrogenFlowSP).EuMax - Stn_AIO(station, asNitrogenFlowSP).EuMin
                    Nitrogen_Output = Stn_AIO(station, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN)))
                    Stn_OutAnalog station, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                End If
                
            End If
        
        Case STN_COMBO3_TYPE
            ' future
                        
        Case Else
        
            ' do nothing
     
    End Select
    

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort
    ' Exit if abort
    ResetErrModule
    Exit Sub
  Case vbRetry
    ' try error line again
    Resume
  Case vbIgnore
    ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub


Public Sub LoadValves_Open(station As Integer, Shift As Integer)
'
' note: switch to correct Stn_Mode before calling this routine
'
If UseLocalErrorHandler Then On Error GoTo localhandler
Dim Idx As Integer
Dim iAuxOut As Integer
    SetErrModule 2, 299
    
    ' Diff station valve
    If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
       And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cON
    End If
    
    ' Shift valves
    Select Case Shift
        Case 1
            ' nothing to do
        Case 2
            Stn_OutDigital station, isLoadShift2Sol, cON
            Stn_OutDigital station, isPurgeShift2Sol, cOFF
            Stn_OutDigital station, isVentShift2Sol, cON
        Case 3
            Stn_OutDigital station, isLoadShift2Sol, cON
            Stn_OutDigital station, isPurgeShift2Sol, cOFF
            Stn_OutDigital station, isVentShift2Sol, cON
            Stn_OutDigital station, isLoadShift3Sol, cON
            Stn_OutDigital station, isPurgeShift3Sol, cOFF
            Stn_OutDigital station, isVentShift3Sol, cON
        Case 4
            Stn_OutDigital station, isLoadShift2Sol, cON
            Stn_OutDigital station, isPurgeShift2Sol, cOFF
            Stn_OutDigital station, isVentShift2Sol, cON
            Stn_OutDigital station, isLoadShift4Sol, cON
            Stn_OutDigital station, isPurgeShift4Sol, cOFF
            Stn_OutDigital station, isVentShift4Sol, cON
    End Select
    
    ' direction valve is off for Load
    Stn_OutDigital station, isPriDirectionSol, cOFF
        
    Select Case STN_INFO(station).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            ' Open Butane valve
            If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isButaneSol, cON
            ' Open Nitrogen valve
            Stn_OutDigital station, isNitrogenSol, cON
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                ' Open Butane valve
                If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isButaneOrvrSol, cON
                ' Open Nitrogen valve
                Stn_OutDigital station, isNitrogenOrvrSol, cON
            Else
                ' use lower range MFC
                ' Open Butane valve
                If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isButaneSol, cON
                ' Open Nitrogen valve
                Stn_OutDigital station, isNitrogenSol, cON
            End If
        Case STN_LIVEFUEL_TYPE
            ' Open LiveFuel valves
            If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isFuelVaporSol, cON
            Stn_OutDigital station, isLiveFuelSol, cON
        Case STN_LIVEREG_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                ' use Live Fuel vapor
                ' Open LiveFuel valves
                If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isFuelVaporSol, cON
                Stn_OutDigital station, isLiveFuelSol, cON
                Stn_OutDigital station, isLoadTypeSelectSol, cON
            Else
                ' use Butane/Nitrogen
                ' Open Butane valve
                If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isButaneSol, cON
                ' Open Nitrogen valve
                Stn_OutDigital station, isNitrogenSol, cON
            End If
        Case STN_LIVEORVR2_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                ' use Live Fuel vapor
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    ' Open LiveFuel valves
                    If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isFuelVaporSol, cON
                    Stn_OutDigital station, isLiveFuelOrvrSol, cON
                    Stn_OutDigital station, isLoadTypeSelectSol, cON
                Else
                    ' use lower range MFC
                    ' Open LiveFuel valves
                    If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isFuelVaporSol, cON
                    Stn_OutDigital station, isLiveFuelSol, cON
                    Stn_OutDigital station, isLoadTypeSelectSol, cON
                End If
            Else
                ' use Butane/Nitrogen
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    ' Open Butane valve
                    If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isButaneOrvrSol, cON
                    ' Open Nitrogen valve
                    Stn_OutDigital station, isNitrogenOrvrSol, cON
                Else
                    ' use lower range MFC
                    ' Open Butane valve
                    If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isButaneSol, cON
                    ' Open Nitrogen valve
                    Stn_OutDigital station, isNitrogenSol, cON
                End If
            End If
        Case STN_COMBO3_TYPE
            ' future
        Case Else
            ' do nothing
    End Select
    
    ' aux canister vent valve
    If StationRecipe(station, Shift).Load_Method = LOADBYBREAKTHRU Then
        If StationRecipe(station, Shift).UseAuxScale = True Then
            Stn_OutDigital station, isAuxCanVentSol, cON
       End If
    End If

    ' aux outputs
    If (USING_AUX_OUTPUTS And StationRecipe(station, Shift).AuxOutputs) Then
        For Idx = 1 To 4
            If (Idx <= NR_AUX_OUTPUTS) Then
                If (StationRecipe(station, Shift).AuxOutputs_Load(Idx)) Then
                    iAuxOut = isAuxOutput1 + Idx - 1
                    Stn_OutDigital station, iAuxOut, cON
                End If
            End If
        Next Idx
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

Public Sub Load_Start(station As Integer, Shift As Integer)
'
'
' Start of a job needs to have valves and Mass Air Controllers set up properly
'
'   If using a FID with live fuel need to wait until operator selects continue after
'   changing the gas can.
'******************************************************************************
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 38
    
' ** Live fuel option allows for multiple passes between gas changes
If systemhasLIVEFUEL Then
    If (AdfControl(station).AdfDefinition.hasLIVEFUEL And StationRecipe(station, Shift).LiveFuel) Then
        If Not AdfControl(station).ReadyForLoad Then
            StationControl(station, Shift).Mode = VBGASPAUSE
'            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
'            Close_Stn_Valves station, Shift
            ResetErrModule
            Exit Sub
        End If
        LoadControl(station, Shift).LoadRateTarget = StationRecipe(station, Shift).Load_RateSave
        LoadControl(station, Shift).TotalWtChg = 0
        LoadControl(station, Shift).TotalWtChgRate = 0
        LoadControl(station, Shift).AuxWt_Start = 0
        LoadControl(station, Shift).PriWt_Start = 0
    End If
End If        ' live fuel
   
' Reset Totals
LoadControl(station, Shift).loadTotalGrams = CSng(0)
LoadControl(station, Shift).LoadTotalLiters = CSng(0)
LoadControl(station, Shift).AuxWtChg = CSng(0)
LoadControl(station, Shift).PriWtChg = CSng(0)
LoadControl(station, Shift).TotalWtChg = CSng(0)
LoadControl(station, Shift).TotalWtChgRate = CSng(0)
LoadControl(station, Shift).loadTotalGrams = 0
LoadControl(station, Shift).ElapsedHours = CSng(0)
LoadControl(station, Shift).LoadRate = CSng(0)
LoadControl(station, Shift).TotalWtChgRate = CSng(0)
                
StationControl(station, Shift).IsPausedInAlarm = False
StationControl(station, Shift).AlarmDelayTime = CLng(0)

If USINGLOADTIMELIMIT Then Stn_LoadLimitStartTime(station, Shift) = Now

Reset_Bar_Graph station, Shift              ' reset the bar graph (also resets Actual & Target)
        
Close_Stn_Valves station, Shift             ' Always start with all valves closed

' Reset Report & Totalize Timers
'Stn_Load_Log_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer          ' do first Totalize after one SysConfig.LoadTotal_Interval
'PreviousReportTimer(station, Shift) = StationControl(station, Shift).TestTimer             ' reset report timer
'PreviousTotalTimer(station, Shift) = StationControl(station, Shift).TestTimer              ' reset totalize timer

' Reset Station Mode
PreviousNow(station, Shift) = Now
StationControl(station, Shift).Mode_StartDts = Now
StationControl(station, Shift).Mode = VBLOAD

ChgErrModule 2, 3837

' Clear LoadData
StnLoadData(station, Shift, 0) = BlankLoadData
StnLoadData(station, Shift, 1) = BlankLoadData
StnLoadData(station, Shift, 2) = BlankLoadData
StnLoadData(station, Shift, 3) = BlankLoadData
StnLoadData(station, Shift, 4) = BlankLoadData
StnLoadData(station, Shift, 5) = BlankLoadData
StnLoadData(station, Shift, 6) = BlankLoadData
StnLoadData(station, Shift, 7) = BlankLoadData
StnLoadData(station, Shift, 8) = BlankLoadData
StnLoadData(station, Shift, 9) = BlankLoadData

' Clear stats, for load and purge if new cycle, for load only if not
If StationRecipe(station, Shift).Purge_Method <> NOPURGE Then
    ' Purge & Load
    Clear_Stats station, Shift, 3
Else
    ' Load Only
    Clear_Stats station, Shift, 2
End If

ChgErrModule 2, 3830

' Set Load Target
Select Case StationRecipe(station, Shift).Load_Method
  
    Case LOADBYTIME
        ChgErrModule 2, 3831
        ' Set Target
        StationControl(station, Shift).Target = StationRecipe(station, Shift).Load_Time
       
    Case LOADBYWC
        ChgErrModule 2, 3832
        ' Determine the working capacity load time
        LoadControl(station, Shift).WC_Load_Time = (StationCanister(station, Shift).WorkingCapacity * StationRecipe(station, Shift).WC_Mult) / StationRecipe(station, Shift).Load_Rate
        ' Determine the working capacity load rate
        If LoadControl(station, Shift).WC_Load_Time > StationRecipe(station, Shift).EPAFill And StationRecipe(station, Shift).EPAFill > 0 Then
            LoadControl(station, Shift).WC_Load_Rate = (StationCanister(station, Shift).WorkingCapacity * StationRecipe(station, Shift).WC_Mult) / StationRecipe(station, Shift).EPAFill
        Else
            LoadControl(station, Shift).WC_Load_Rate = StationRecipe(station, Shift).Load_Rate
        End If
        StationRecipe(station, Shift).WC_Mult = StationRecipe(station, Shift).WC_MultSave
        ' Set Target
        StationControl(station, Shift).Target = StationCanister(station, Shift).WorkingCapacity * StationRecipe(station, Shift).WC_Mult
        
    Case LOADBYWEIGHT
        ChgErrModule 2, 3833
        ' Set Target
        StationControl(station, Shift).Target = StationRecipe(station, Shift).Load_Wt
    
    Case LOADBYBREAKTHRU
        ChgErrModule 2, 3834
        ' Set Target
        StationControl(station, Shift).Target = StationRecipe(station, Shift).LoadBreakthrough
        ' turn on AuxCanVent Sol (if required)
        If StationRecipe(station, Shift).UseAuxScale = True Then
            Stn_OutDigital station, isAuxCanVentSol, cON
        End If
      
    Case LOADBYFID
        ChgErrModule 2, 3835
        ' Set Target
        StationControl(station, Shift).Target = StationRecipe(station, Shift).FIDmg
'        Background = FIDOutputReadAdjusted   'This is the load tar value
    
End Select
    
' set Load Minimum Duration (in seconds)
' LoadMinDuration(station) = 10
' LoadMinDuration(station) = ((60# / 10#) * (CSng(Can.WorkingCapacity) / CSng(Rcp.Load_Rate)))  ' 10% of est. load duration
LoadMinDuration(station) = CLng(LoadEqlDelayTime)

' set Load target (for statistics)
LoadControl(station, Shift).LoadTarget = StationControl(station, Shift).Target

ChgErrModule 2, 3836
 
' clear XY Graph values at start of first Load cycle
If ((StationControl(station, Shift).CurrCycle = 1) And (StationControl(station, Shift).Course = 1)) Then
    ' Reset XY Graph reference values for this station's Scale(s)
    If StationControl(station, Shift).PriScaleWt > Stn_PriScale_RefValues(station, Shift) + (0.2 * StationCanister(station, Shift).WorkingCapacity) Then
        ' resync "zero" weight to current weight
        Stn_PriScale_RefValues(station, Shift) = StationControl(station, Shift).PriScaleWt
    ElseIf StationControl(station, Shift).PriScaleWt < Stn_PriScale_RefValues(station, Shift) - (0.25 * StationCanister(station, Shift).WorkingCapacity) Then
        ' resync "zero" weight to current weight
        Stn_PriScale_RefValues(station, Shift) = StationControl(station, Shift).PriScaleWt
    End If
    If StationControl(station, Shift).AuxScaleWt > Stn_AuxScale_RefValues(station, Shift) + (0.2 * StationCanister(station, Shift).WorkingCapacity) Then
        ' resync "zero" weight to current weight
        Stn_AuxScale_RefValues(station, Shift) = StationControl(station, Shift).AuxScaleWt
    ElseIf StationControl(station, Shift).AuxScaleWt < Stn_AuxScale_RefValues(station, Shift) - (0.25 * StationCanister(station, Shift).WorkingCapacity) Then
        ' resync "zero" weight to current weight
        Stn_AuxScale_RefValues(station, Shift) = StationControl(station, Shift).AuxScaleWt
    End If
    If StationControl(station, Shift).Course = 1 Then
'        ' clear XY Graph values
'        frmStnDetail.ClearXYvalues station, shift
        ' shift Y Graph values
        frmStnDetail.Shift_Yvalues station, Shift
    End If
End If

                
' Reset Report & Totalize Timers
Stn_Load_Log_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer          ' do first Totalize after one SysConfig.LoadTotal_Interval
PreviousReportTimer(station, Shift) = StationControl(station, Shift).TestTimer             ' reset report timer
PreviousTotalTimer(station, Shift) = StationControl(station, Shift).TestTimer              ' reset totalize timer


ChgErrModule 2, 3838
' start loading
ChgPhase LoadPrep, (Now + TimeSerial(0, 0, LoadMfcDelayTime)), station, Shift

' Write Load Data to File
Load_Write station, Shift, LOADBEGIN

ChgErrModule 2, 3839
' Load Valves
LoadValves_Open station, Shift
If systemhasLIVEFUEL Then
    If (AdfControl(station).AdfDefinition.hasLIVEFUEL And StationRecipe(station, Shift).LiveFuel) Then
        If (AdfControl(station).AdfDefinition.hasLIVEFUEL And StationRecipe(station, Shift).LiveFuel) Then
            Stn_OutDigital station, isFuelVentSol, cON
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


Public Sub Load_Continue(station As Integer, Shift As Integer)
'
'
' Continue of a job needs to have valves and Mass Air Controllers set up properly
'
'   If using live fuel need to wait until operator selects continue after
'   changing the gas can.
'******************************************************************************
Dim cnt1 As Long
' If UseLocalErrorHandler Then On Error GoTo localhandler
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 381
    
If USINGLOADTIMELIMIT Then Stn_LoadLimitStartTime(station, Shift) = Now

If StationRecipe(station, Shift).UseAuxScale Then
    Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = True
End If
If StationRecipe(station, Shift).UsePriScale Then
    Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = True
End If

' Adjust Total Time in Alarm
If StationRecipe(station, Shift).Load_Method = LOADBYTIME Or StationRecipe(station, Shift).Load_Method = LOADBYWC Then
    cnt1 = CLng(Second(Now - StationControl(station, Shift).PauseAlarmStartTime)) + CLng(Minute(Now - StationControl(station, Shift).PauseAlarmStartTime) * 60) + CLng(Hour(Now - StationControl(station, Shift).PauseAlarmStartTime) * 3600)
    StationControl(station, Shift).AlarmDelayTime = StationControl(station, Shift).AlarmDelayTime + cnt1
Else
    StationControl(station, Shift).AlarmDelayTime = CLng(0)
End If


' Reset Report & Totalize Timers
Stn_Load_Log_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer                                  ' do next Totalize after one SysConfig.LoadTotal_Interval
PreviousReportTimer(station, Shift) = StationControl(station, Shift).TestTimer                                     ' reset report timer
PreviousTotalTimer(station, Shift) = StationControl(station, Shift).TestTimer                                      ' reset totalize timer

' Reset Station Mode
PreviousNow(station, Shift) = Now
StationControl(station, Shift).Mode_StartDts = Now
StationControl(station, Shift).Mode = VBLOAD

' start loading
ChgPhase LoadStarting, (Now + TimeSerial(0, 0, LoadMfcDelayTime)), station, Shift

' open valves
LoadValves_Open station, Shift
' update mfc(s)
LoadSetPoint_Update station, Shift

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

Public Sub Load_Done(station As Integer, Shift As Integer)
Dim Result As Variant
Dim sButane As Single
Dim sFuelVapor As Single
Dim tmpWt As Single
Const done As Integer = 2
If UseLocalErrorHandler Then On Error GoTo localhandler
    SetErrModule 2, 40
    
    ' increment the number of Completed Loads
    StationControl(station, Shift).CompletedLoads = StationControl(station, Shift).CompletedLoads + 1
    ' did a Load actually happen ??
    If (StationRecipe(station, Shift).Load_Method <> NOLOAD) Then
        ' a Load actually happened
        ' Remember Primary Scale Value at End of Load
        If StationRecipe(station, Shift).UsePriScale Then
            LoadControl(station, Shift).PriWt_End = StationControl(station, Shift).PriScaleWt
        Else
            LoadControl(station, Shift).PriWt_End = 0
        End If
        ' Remember Aux Scale Value at End of Purge
        If StationRecipe(station, Shift).UseAuxScale Then
            LoadControl(station, Shift).AuxWt_End = StationControl(station, Shift).AuxScaleWt
        Else
            LoadControl(station, Shift).AuxWt_End = 0
        End If
        StationControl(station, Shift).End_Time = Now
        StationControl(station, Shift).End_Timer = StationControl(station, Shift).TestTimer
        StationControl(station, Shift).IsPausedInAlarm = False
        Load_Write station, Shift, LOADDONE
        Stats_Write station, Shift
        Select Case StationRecipe(station, Shift).LiveFuel
            Case True
                ' using LiveFuel
                sButane = 0
                sFuelVapor = LoadControl(station, Shift).TotalWtChg
            Case False
                ' using Butane
                sButane = LoadControl(station, Shift).loadTotalGrams
                sFuelVapor = 0
        End Select
        Write_FuelUseLog Now, sButane, sFuelVapor
        Reset_Bar_Graph station, Shift                      ' reset the bar graph
        
        FirstTime(station, Shift) = False                   ' clear OOT FirstTime Flag   2 Mar 2005
        
        StationRecipe(station, Shift).Load_Rate = StationRecipe(station, Shift).Load_RateSave
        If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
           And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
            Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
        End If
        If StationRecipe(station, Shift).LiveFuel Then
            StationControl(station, 1).LiveFuelCycleCount = StationControl(station, 1).LiveFuelCycleCount + 1
        End If
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).Load_TotalGrams = LoadControl(station, Shift).loadTotalGrams
        LoadControl(station, Shift).loadTotalGrams = 0
    End If
    
    
    ' WHAT IS NEXT ??
    ' Need to PauseAfterLoad ??
    If StationRecipe(station, Shift).PauseAfterLoad Then
        ' cycle is not complete; PauseAfterLoad
        Pause_AfterLoad station, Shift
    ElseIf StationRecipe(station, Shift).PauseAfterLoadForOper Then
        ' cycle is not complete; PauseAfterLoadForOperator
        Pause_AfterLoadForOper station, Shift
    ElseIf (StationControl(station, Shift).CompletedLoads = StationControl(station, Shift).CompletedPurges) Then
        ' cycle is complete
        StationControl(station, Shift).CompletedCycles = StationControl(station, Shift).CompletedCycles + 1
        ' net scale end weight
        tmpWt = CSng(0)
        If StationRecipe(station, Shift).UsePriScale Then tmpWt = tmpWt + StationControl(station, Shift).PriScaleWt
        If StationRecipe(station, Shift).UseAuxScale Then tmpWt = tmpWt + StationControl(station, Shift).AuxScaleWt
        StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total = tmpWt
        ' is recipe complete ??
        If RecipeIsDone(station, Shift) Then
            ' recipe completed OK
            JobInfo(station, Shift).End_OK = True
            ' any more LiveFuel Loads ??
            If Not AnyMoreLiveFuelLoads(station, Shift) Then
                ' No More LiveFuel Loads; using AutoDrainFill ??
                If (AdfControl(station).AdfDefinition.hasAUTODRAINFILL And AdfControl(station).LiveFuel And AdfControl(station).LiveFuelChgAuto) Then
                    ' Got to Empty the Live Fuel Tank
                    AdfControl(station).Mode = 1
                    AdfControl(station).Step = 0
                Else
                    AdfControl(station).Mode = 0
                    AdfControl(station).Step = 0
                End If
            End If
            ' *****************************************
            ' What's Next ??
            Course_Next station, Shift
            ' *****************************************
        Else
            ' recipe is not complete; new cycle needs a purge
            StationControl(station, Shift).CurrCycle = StationControl(station, Shift).CurrCycle + 1
            ' new cycle start weight = last cycle end weight
            StationCycleWeightData(station, Shift, StationControl(station, Shift).CurrCycle).Cycle_StartWeight_Total = _
                StationCycleWeightData(station, Shift, StationControl(station, Shift).CompletedCycles).Cycle_EndWeight_Total
            ' a Purge is next
            If (StationRecipe(station, Shift).Purge_Method = NOPURGE) Then
                ' End a Purge
                Purge_Done station, Shift
            Else
                ' Start a Purge
                Purge_Start station, Shift
            End If
        End If
    Else
        ' cycle is not complete; current cycle needs a Purge
        If (StationRecipe(station, Shift).Purge_Method = NOPURGE) Then
            ' End a Purge
            Purge_Done station, Shift
        Else
            ' Start a Purge
            Purge_Start station, Shift
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

Public Sub LoadController(ByVal iStn As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1414

   
    Controller_PID (iStn + 10)

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Sub Load_Abort(station As Integer, Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 41
Dim iCourse As Integer
Dim sButane As Single
Dim sFuelVapor As Single
Dim sPrint As String
    
    ' time stamp end of current course
    iCourse = StationControl(station, Shift).Course
    StationSequence(station, Shift).CourseData(iCourse).DtsEnd = Now()
    StationControl(station, Shift).End_Time = Now
    StationControl(station, Shift).End_Timer = StationControl(station, Shift).TestTimer
      
    ' Remember Primary Scale Value at End of Load
    If StationRecipe(station, Shift).UsePriScale Then
        LoadControl(station, Shift).PriWt_End = StationControl(station, Shift).PriScaleWt
    Else
        LoadControl(station, Shift).PriWt_End = 0
    End If
    ' Remember Aux Scale Value at End of Purge
    If StationRecipe(station, Shift).UseAuxScale Then
        LoadControl(station, Shift).AuxWt_End = StationControl(station, Shift).AuxScaleWt
    Else
        LoadControl(station, Shift).AuxWt_End = 0
    End If
    
    ' Update Header data in data file
    Header_Update station, Shift
    Load_Write station, Shift, LOADDONE
    Stats_Write station, Shift
    ' Write CycleWeights data in data file
    Weights_Write station, Shift
      
    Select Case StationRecipe(station, Shift).LiveFuel
        Case True
            ' using LiveFuel
            sButane = 0
            sFuelVapor = LoadControl(station, Shift).TotalWtChg
        Case False
            ' using Butane
            sButane = LoadControl(station, Shift).loadTotalGrams
            sFuelVapor = 0
    End Select
    Write_FuelUseLog Now, sButane, sFuelVapor
    
    StationControl(station, Shift).IsPausedInAlarm = False
    Write_JLog station, Shift, "Load #" & Format(StationControl(station, Shift).CurrCycle, "###0") & "Aborted"
    ' reset the bar graph
    Reset_Bar_Graph station, Shift
    ' clear OOT FirstTime Flag
    FirstTime(station, Shift) = False
    
    StationRecipe(station, Shift).Load_Rate = StationRecipe(station, Shift).Load_RateSave
    Select Case STN_INFO(station).Type
    
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO                 'close mfc
            Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
        
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
                Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
            Else
                ' use lower range MFC
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
            End If
                
        Case STN_LIVEFUEL_TYPE
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO            'close mfc
        
        Case STN_LIVEREG_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO        'close mfc
                Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
            Else
                Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO             'close mfc
                Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
            End If
        
        Case STN_LIVEORVR2_TYPE
            If StationRecipe(station, Shift).LiveFuel Then
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO        'close mfc
                    Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO        'close mfc
                    Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                End If
            Else
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
                    Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                    Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                End If
            End If
        
        Case STN_COMBO3_TYPE
            ' future
        
        Case Else
      
    End Select
    
    ' Close Valves
    Close_Stn_Valves station, Shift
    If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).PriScaleNo > 0 _
      And StationRecipe(station, Shift).PriScaleNo < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF
    End If
    ' clear Total MFC grams loaded
    LoadControl(station, Shift).loadTotalGrams = 0
    ' update event log
    Write_ELog "Load cycle aborted " & "  Station " & station & "  Shift" & Shift
    ' write the logs and close off the world
    Station_Finish station, Shift
       

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

Sub LiveFuel_Init(station As Integer, Shift As Integer)
    '  Initialize LF/ADF Variables at Start of a Job
    '
    AdfControl(station).InitialFill_Complete = False
    AdfControl(station).ReadyForLoad = False
    AdfControl(station).ReadyForRefill = False
    AdfControl(station).RefillRequest = False
    AdfControl(station).Enable = False
    AdfControl(station).Heater_Enable = False
    AdfControl(station).ManScreen_Enable = False
    AdfControl(station).ButtonVisible_Done = False
    AdfControl(station).ButtonVisible_Retry = False
    AdfControl(station).ButtonVisible_Stop = False
    
    StationControl(station, Shift).LiveFuelCycleCount = 0
    AdfControl(station).Mode = 0
    AdfControl(station).Step = 0
    AdfControl(station).Message = ""
    AdfControl(station).TempOK = False
'    AdfControl(station).LiveFuelState = fuelOk
'    AdfControl(station).LiveFuelDensityOkCnt = 0
'    AdfControl(station).LiveFuelDensityDeadCnt = 0
'    AdfControl(station).LiveFuelDensityWeakCnt = 0
    
End Sub

Sub LiveFuel_Update(ByVal station As Integer, ByVal Shift As Integer)
    '  Update LF/ADF Variables at Start of a Recipe
    '
    If AdfControl(station).AdfDefinition.hasLIVEFUEL Then
        If (StationRecipe(station, 1).Load_Method = NOLOAD) Then
            AdfControl(station).LiveFuel = False
            AdfControl(station).LiveFuelChgAuto = False
            AdfControl(station).LiveFuelChgFreq = 0
        Else
            AdfControl(station).LiveFuel = StationRecipe(station, 1).LiveFuel
            If AdfControl(station).AdfDefinition.hasAUTODRAINFILL Then
                AdfControl(station).LiveFuelChgAuto = StationRecipe(station, 1).LiveFuelChgAuto
                AdfControl(station).LiveFuelChgFreq = StationRecipe(station, 1).LiveFuelChgFreq
                If AdfControl(station).AdfDefinition.hasADF_Heater Then
                    AdfControl(station).Heater = StationRecipe(station, 1).ADF_Heater
                    AdfControl(station).HeaterSP = StationRecipe(station, 1).ADF_HeaterSP
                Else
                    AdfControl(station).Heater = False
                End If
            Else
                AdfControl(station).LiveFuelChgAuto = False
            End If
        End If
    Else
        AdfControl(station).LiveFuel = False
    End If
End Sub

Sub Station_Abort(station As Integer, Shift As Integer, stopcode As Integer)
' Function Name:    Station_Abort
' Author:           Brunrose         Feb 2007
' Description:      This routine aborts  the current process.
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 33
Dim WhoAborted As Integer
Dim sPrint As String

    ' Update RemCanLoad, if required
    If USINGREMCANLOAD Then
        If (Len(StnRemoteTask(station, Shift).TaskID) > 2) Then
            ' Remote Task did not complete successfully
            StnRemoteTask(station, Shift).PreviousResult = ModeDescShort(StationControl(station, Shift).Mode) & " Abort"
            ' set Remote Task status in DB to Failed(AVL Files) OR Ready(other) and add PreviousResult
            Select Case USINGREMAVLFILES
                Case True
                    RemTask_Update station, Shift, "Failed", StnRemoteTask(station, Shift).PreviousResult
                Case False
                    RemTask_Update station, Shift, "Ready", StnRemoteTask(station, Shift).PreviousResult
            End Select
            ' update event log
            sPrint = "Remote Task >" & StnRemoteTask(station, Shift).TaskID
            sPrint = sPrint & "< (Job# " & StationControl(station, Shift).Job_Number
            sPrint = sPrint & ") did not complete successfully."
            Write_ELog sPrint
        End If
    End If
      
    WhoAborted = RESULTABORTAUTO - 1
    Select Case stopcode
        Case AUTO_STOP
            WhoAborted = RESULTABORTAUTO
            If frmStnDetail.Visible _
                And DispStn = station And DispShift = Shift Then
                    If NR_SHIFT = 1 Then
                        frmStnDetail.txtStnDtlMsg.text = vbCrLf & "Station " & Format(station, "0") _
                                & " Stopped"
                    Else
                        frmStnDetail.txtStnDtlMsg.text = vbCrLf & "Station " & Format(station, "0") _
                                & " Shift " & Format(Shift, "0") _
                                & " Stopped"
                    End If
            End If
        Case EXIT_STOP
            WhoAborted = RESULTABORTAUTO
            If NR_SHIFT = 1 Then
                frmCheckIt.lblMsg.Caption = vbCrLf & "Station " & Format(station, "0") _
                        & " Stopped"
            Else
                frmCheckIt.lblMsg.Caption = vbCrLf & "Station " & Format(station, "0") _
                        & " Shift " & Format(Shift, "0") _
                        & " Stopped"
            End If
        Case OPER_STOP
            WhoAborted = RESULTABORTOPER
            frmStnDetail.txtStnDtlMsg.text = vbCrLf & "Pushed Stop (Abort) Button"
        Case Else
            ' don't do anything
    End Select
    
    If StationControl(station, Shift).Mode = VBPAUSEALARM Or StationControl(station, Shift).Mode = VBPAUSEOOT Then
        If StationControl(station, Shift).IsPausedInAlarm Then                          ' station was paused for alarm
            StationControl(station, Shift).IsPausedInAlarm = False                      ' reset station alarm indicator
        End If
        StationControl(station, Shift).Mode = StationControl(station, Shift).Mode_PauseSave
        StationControl(station, Shift).Mode_PauseSave = VBIDLE
    End If
    
'    StationControl(station, shift).Course = StationSequence(station, shift).NumCourses
    
    Select Case StationControl(station, Shift).Mode
        Case VBLEAK, VBLEAKERROR
            LeakCheck_Abort station, Shift, WhoAborted
            
        Case VBPURGE, VBPURGECONT, VBPURGEWAIT, VBPOSTPURGE
            Purge_Abort station, Shift
            
        Case VBLOAD, VBPOSTLOAD
            Load_Abort station, Shift
            
        Case VBFIDPAUSE, VBGASPAUSE
            Load_Abort station, Shift
            
        Case VBPRELOAD
            PreLoad_Abort station, Shift
            
        Case VBLEAKTEST
            ' Update Header data in data file
            JobInfo(station, Shift).End_OK = False
            StationControl(station, Shift).End_Time = Now
            JobInfo(station, Shift).End_Baro = AmbBaro
            Header_Update station, Shift
            LT2_Write station, Shift, "Stopped by Operator", CurrLT2_Data
            SEQ_Step(station, Shift) = 90
            
        Case Else
            ' If a DB file is still open, Close It
            If Len(StationControl(station, Shift).DBFile) > 0 Or StationControl(station, Shift).Mode <> VBIDLEWAITING Then
                ' write the logs and close off the world
                Station_Finish station, Shift
            End If
    End Select
    
    ' LiveFuel Vapor Generator Tank AutoDrainFill
    If AdfControl(station).Mode <> 0 Then AdfControl(station).Step = 90  ' Abort
    ' LiveFuel Fuel Storage Tank Drain & Fill
    If FstControl(station).Mode <> 0 Then FstControl(station).Step = 90  ' Abort
      
    Stop_In_Progress = False
    
ResetErrModule
Exit Sub

    Select Case StationConfig(station, Shift).LeakCheckFailResponse
    
        Case MANUALSTOP
            ' release resources & wait for operator to press STOP
            If StationRecipe(station, Shift).UseAuxScale Then
                Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = False
            End If
            If StationRecipe(station, Shift).UsePriScale Then
                Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = False
            End If
            
        Case AUTOSTOP
            ' release resources & stop test
            If StationRecipe(station, Shift).UseAuxScale Then
                Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = False
            End If
            If StationRecipe(station, Shift).UsePriScale Then
                Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = False
            End If
            ' stop the test
            Station_Abort station, Shift, AUTO_STOP
            
        Case AUTOCONTINUE
            ' continue anyway
            ALM_Write station, Shift, "Automatic Continue after LC Failure"
            Leak_Write station, Shift, LCAUTOCONTINUE, NORESULT
            LeakCheck_Next station, Shift
      
            
        Case Else
            ' do not release FID and Scales
            
    End Select
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

Sub Station_StartPB(station As Integer, Shift As Integer)
' Function Name:    Station_StartPB
' Author:           Brunrose         2008
' Description:      This routine responds to the Start Pushbutton
'                   A DB File is opened and
'                   test variables are initialized.
'                   If a Test Start delay is in the recipe for this test,
'                   the delay is started.
'                   Otherwise the test is started.
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1

    '*********************
    ' Start button pushed
    '*********************
    
    ' open database/clear fields
    Station_Init station, Shift
    
    ' initialize live fuel (esp. ADF) variables
    LiveFuel_Init station, Shift

    ' clear sequence variables
    SEQ_Nmbr(station, Shift) = seqIdle
    SEQ_Step(station, Shift) = 0
    SEQ_Alarm(station, Shift) = False
    SEQ_OOT(station, Shift) = False
    
    '   reset (optional) Local PAS Timeout Timers
    If USINGPASLOCALCONTROL Then
        If PAS_INFO(pasTEMPERATURE).timeOut Then
            PAS_INFO(pasTEMPERATURE).TimeOutDuration = 0#
            PAS_INFO(pasTEMPERATURE).timeOut = False
        End If
        If PAS_INFO(pasMOISTURE).timeOut Then
            PAS_INFO(pasMOISTURE).TimeOutDuration = 0#
            PAS_INFO(pasMOISTURE).timeOut = False
        End If
    End If
               
    ' scale ownership
    If USINGHARDPIPEDSCALES Then
        ' two scales per station, fixed assignments for pri & aux for each station; stn#1 pri = 1, stn#1 aux = 2, etc.
        StationControl(station, Shift).PriScaleStn = station
        StationControl(station, Shift).AuxScaleStn = station
    ElseIf (StationSequence(station, Shift).NumCourses > 1) Then
        ' multiple courses; scale assignments are made for entire JobSequence; one scale per station, selectable as pri or aux by any station
        If (StationSequence(station, Shift).PriScaleNo <> 0) Then
            StationControl(station, Shift).PriScaleStn = StationSequence(station, Shift).PriScaleNo
        Else
            StationControl(station, Shift).PriScaleStn = 0
        End If
        If (StationSequence(station, Shift).AuxScaleNo <> 0) Then
            StationControl(station, Shift).AuxScaleStn = StationSequence(station, Shift).AuxScaleNo
        Else
            StationControl(station, Shift).AuxScaleStn = 0
        End If
    Else
        ' default course; scale assignments are made by Recipe; one scale per station, selectable as pri or aux by any station
        If (StationRecipe(station, Shift).UsePriScale) Then
            StationControl(station, Shift).PriScaleStn = StationRecipe(station, Shift).PriScaleNo
        Else
            StationControl(station, Shift).PriScaleStn = 0
        End If
        If (StationRecipe(station, Shift).UseAuxScale) Then
            StationControl(station, Shift).AuxScaleStn = StationRecipe(station, Shift).AuxScaleNo
        Else
            StationControl(station, Shift).AuxScaleStn = 0
        End If
    End If
    
    If (STN_INFO(station).Type = STN_LEAKTEST_TYPE) Then
        ' LeakTest station
        StationControl(station, Shift).Mode = VBLEAKTEST
        StationControl(station, Shift).Course = 1
        StationControl(station, Shift).Job_Description = "40 CFR 1066.885 LeakTest"
        StationControl(station, Shift).Start_Time = Now
        StationControl(station, Shift).End_Time = Now
        ' initialize sequence variables
        SEQ_Nmbr(station, Shift) = seqLeakTest
        ' Reset Report & Totalize Timers
        Stn_LT_Log_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer             ' do first normal report after one Leak_Interval
        PreviousReportTimer(station, Shift) = StationControl(station, Shift).TestTimer              ' reset report timer
        ' Clear LeakTest Data
        CurrLT2_Data = BlankLT2_Data
        ' Set barometer values
        JobInfo(station, Shift).Start_Baro = AmbBaro
        JobInfo(station, Shift).End_Baro = AmbBaro
        ' default result = Failed; set to True, i.e. Passed, when job completes successfully
        JobInfo(station, Shift).End_OK = False
        ' Write Header data to data file
        Header_Write station, Shift

    Else
        ' setup first job sequence course
        Course_Init station, Shift, CInt(1)
    End If

    ' Add Job to the Joblist
    AddNew_Joblist _
        StationControl(station, Shift).Job_Number, _
        StationControl(station, Shift).Job_Description, _
        StationControl(station, Shift).Start_Time, _
        station, _
        Shift, _
        StationControl(station, Shift).RptFile
    
    ' Reset stnXYGraph
    Stn_XYGraph_TestTimer(station, Shift) = StationControl(station, Shift).TestTimer
    frmStnDetail.SetXYtimeInterval station, Shift
    ' clear XY Graph values
    frmStnDetail.ClearXYvalues station, Shift
    ' Slide stn XY Graph to the left 3 time divisions
    frmStnDetail.SlideXYgraph station, Shift
    frmStnDetail.SlideXYgraph station, Shift
    frmStnDetail.SlideXYgraph station, Shift
    
    Stn_RemStatus_Log_TestTimer(station, Shift) = Timer
        
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

Sub Recipe_Start(station As Integer, Shift As Integer)
' Function Name:    RecipeStart
' Description:      This routine starts a test
'
'                   Some tests start from idle to PURGE
'                   Others go from idle to leak check or just leak/purge check
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 2
Dim iSta As Integer
Dim iShift As Integer
Dim leekin As Integer
Dim presspurge As Integer
Dim tmpWt As Single

'****************************************************************************************************
' Start button pushed ( and No Start delay) Or End of Start Delay Or End of (Scale Or Shift) Waiting
'****************************************************************************************************

leekin = 0
presspurge = 0

' With Positive Pressure Purge - Can't Leak & Purge at same time
If USINGPRESSUREPURGE And StationConfig(station, Shift).PosPressPurge Then
    For iShift = 1 To NR_SHIFT
        For iSta = 1 To LAST_STN
            If StationControl(iSta, iShift).Mode = VBLEAK Then
                leekin = leekin + 1
            End If
            If StationControl(iSta, iShift).Mode = VBPURGE Or StationControl(iSta, iShift).Mode = VBPURGEWAIT Or StationControl(iSta, iShift).Mode = VBPURGECONT Then
                presspurge = presspurge + 1
            End If
        Next iSta
    Next iShift
End If
    

' *************NEW*****************
' **** wait for other shift ? *****
' *********************************
If NR_SHIFT > 1 Then
    ' more than one shift
    For iShift = 1 To NR_SHIFT
        ' don't check yourself
        If (iShift <> Shift) Then
            If Not StationControl(station, iShift).ModeIsIdle_Debounced Then
                ' other shift is not Idle
                If StationControl(station, iShift).Mode <> VBSHIFTWAIT And StationControl(station, iShift).Mode <> VBSCALEWAIT Then
                    If StationControl(station, iShift).Mode <> VBSTARTWAIT And StationControl(station, iShift).Mode <> VBLEAKERROR Then
                        ' ************************
                        ' wait for the other shift
                        ' ************************
                        StationControl(station, Shift).Mode = VBSHIFTWAIT
                        ' release any InUse scales
                        If StationControl(station, Shift).ScalesInUse Then
                            ' have already placed scale(s) InUse
                            ' **** using a Primary Scale ? *****
                            If StationRecipe(station, Shift).UsePriScale Then
                                If Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) Then
                                    ' Release the Primary Scale
                                    Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = False
                                End If
                            End If
                            ' **** using an Aux Scale ? *****
                            If StationRecipe(station, Shift).UseAuxScale Then
                                If Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) Then
                                    ' Release the Aux Scale
                                    Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = False
                                End If
                            End If
                            ' scales released
                            StationControl(station, Shift).ScalesInUse = True
                        End If
                        ResetErrModule
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next iShift
End If


' ******************************
' **** wait for scale(s) ? *****
' ******************************
If Not StationControl(station, Shift).ScalesInUse Then
    ' have not yet placed scales InUse
    If StationRecipe(station, Shift).UsePriScale Or StationRecipe(station, Shift).UseAuxScale Then
        ' Using 2 scales ?
        If StationRecipe(station, Shift).UsePriScale And StationRecipe(station, Shift).UseAuxScale Then
            ' Either scale in use ?
            If Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) Or Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) Then
                ' ******************
                ' wait for the scale
                ' ******************
                StationControl(station, Shift).Mode = VBSCALEWAIT
                ResetErrModule
                Exit Sub
            Else
                ' Place Both Scales InUse
                StationControl(station, Shift).ScalesInUse = True
                Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = True
                Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = True
            End If
        Else    '   using 1 scale
            ' **** using a Primary Scale *****
            If StationRecipe(station, Shift).UsePriScale Then
                If Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) Then
                    ' ******************
                    ' wait for the scale
                    ' ******************
                    StationControl(station, Shift).Mode = VBSCALEWAIT
                    ResetErrModule
                    Exit Sub
                Else
                    ' Place the Primary Scale InUse
                    StationControl(station, Shift).ScalesInUse = True
                    Scale_In_Use(StationRecipe(station, Shift).PriScaleNo) = True
                End If
            End If
            ' **** using an Aux Scale *****
            If StationRecipe(station, Shift).UseAuxScale Then
                If Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) Then
                    ' ******************
                    ' wait for the scale
                    ' ******************
                    StationControl(station, Shift).Mode = VBSCALEWAIT
                    ResetErrModule
                    Exit Sub
                Else
                    ' Place the Aux Scale InUse
                    StationControl(station, Shift).ScalesInUse = True
                    Scale_In_Use(StationRecipe(station, Shift).AuxScaleNo) = True
                End If
            End If
        End If
    End If
End If
        
' ****************************
'      READY TO START
' ****************************
' Simulation - Aux & Pri Scale Starting Weights
'If (StationControl(station, Shift).Course = 1) Then
'    Sim_AuxWt_Current(station) = CSng(0.01) * Sim_AuxCan_JobStartPercentLoaded(station) * StationCanister(station, Shift).WorkingCapacity
'    Sim_PriWt_Current(station) = CSng(0.01) * Sim_PriCan_JobStartPercentLoaded(station) * StationCanister(station, Shift).WorkingCapacity
'End If
' Total Scale Starting Weights
tmpWt = 0
If StationRecipe(station, Shift).UsePriScale Then
    tmpWt = tmpWt + Scale_Weight(StationRecipe(station, Shift).PriScaleNo)
End If
If StationRecipe(station, Shift).UseAuxScale Then
    tmpWt = tmpWt + Scale_Weight(StationRecipe(station, Shift).AuxScaleNo)
End If
StationCycleWeightData(station, Shift, 0).Cycle_EndWeight_Total = tmpWt
StationCycleWeightData(station, Shift, 1).Cycle_StartWeight_Total = tmpWt
' ****************************
'      What to do ??
' ****************************
If (STN_INFO(station).Type = STN_LEAKTEST_TYPE) Then
    ' this is a LeakTest station
    StationControl(station, Shift).Mode = VBLEAKTEST
ElseIf StationRecipe(station, Shift).LeakCheck Then
    ' Need to Start with a Leak Check
    If LeakCheckControl.station = 0 And presspurge = 0 Then    'Do leak check when nobody else is using PT
        ' ****************
        ' start Leak Check
        ' ****************
        LeakCheck_Start station, Shift
    Else
        ' *******************************************
        ' Somebody else is leaking or PressurePurging
        ' keep waiting for opportunity to leak check
        ' *******************************************
        StationControl(station, Shift).Mode = VBLEAKWAIT
        ResetErrModule
        Exit Sub
    End If
Else
    ' No Leak Check Needed; Start a Purge ?? or a Load ??
    Select Case StationRecipe(station, Shift).CycleType
        Case CyclePurgeLoad
            ' Purge then Load Cycles
            If StationRecipe(station, Shift).Purge_Method <> NOPURGE Then
                ' Start a Purge
                Purge_Start station, Shift
            Else
                ' End a Purge
                Purge_Done station, Shift
            End If
        Case CycleLoadPurge
            ' Load then Purge Cycles
            If StationRecipe(station, Shift).Load_Method <> NOLOAD Then
                ' Start of a Load
                PreLoad_Start station, Shift
            Else
                ' End a Load
                Load_Done station, Shift
            End If
    End Select
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

Public Sub Load_Check(station As Integer, Shift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1004
Dim Nitrogen_Rate As Single
Dim Butane_Rate As Single
Dim NitrogenORVR_Rate As Single
Dim ButaneORVR_Rate As Single
Dim Nitrogen_Output As Single
Dim Butane_Output As Single
Dim ButaneORVR_Output As Single
Dim NitrogenORVR_Output As Single
Dim LiveFuel_Rate As Single
Dim LiveFuel_Output As Single
Dim Result As Variant
Dim span As Single
Dim time_seconds As Long
Dim secDelay As Double
Dim minrate As Single
Dim maxmass As Single
Dim maxtime As Single
Dim deltatime As Single
Dim iController As Integer

    ' Continuously update baro.
    JobInfo(station, Shift).End_Baro = AmbBaro
    ' Set LoadRate Controller index
    iController = station + 10
    
    Select Case LoadControl(station, Shift).Phase
    
        Case LoadPrep
            '
            ' Preliminaries
            '
            If Now > LoadControl(station, Shift).PhaseDts Then
                ' Delay is done
                ' reset LoadRate Controller PID (the PID control will set Enable true)
                PID_INFO(iController).Enable = False
                ' Put MFC(s) into operation
                LoadSetPoint_Update station, Shift
                ' set Concordance Start Time (actuall*y Timer)
                secDelay = CDbl(0.1 * (60 * EstimatedLoadDuration(StationRecipe(station, Shift), StationCanister(station, Shift))))
                If (LoadEqlDelayTime < secDelay) Then secDelay = LoadEqlDelayTime
                Stn_LoadEql_StartTimer(station, Shift) = StationControl(station, Shift).TestTimer + secDelay
'                PreviousTotalTimer(station, Shift) = StationControl(station, Shift).TestTimer              ' reset totalize timer
                ' change load phase to "getting started"
                ChgPhase LoadStarting, (Now + TimeSerial(0, 0, MFC_Settle_Time)), station, Shift
            End If

    
        Case LoadStarting
            '
            ' Getting Started
            '
            If Now > LoadControl(station, Shift).PhaseDts Then
                ' Delay is done
                ' if live fuel, vapor flow ok, close vent
                If systemhasLIVEFUEL Then
                    If (AdfControl(station).AdfDefinition.hasLIVEFUEL And StationRecipe(station, Shift).LiveFuel) Then
                        Stn_OutDigital station, isFuelVentSol, cOFF
                    End If
                End If
                ' change load phase to "are we done yet?"
                ChgPhase LoadLoading, (Now + TimeSerial(0, 0, LoadMinDuration(station))), station, Shift
            End If
            
        Case LoadLoading
            '
            ' LOAD CYCLE    "are we done yet?"
            '
    
            ' *** Load Pressure fault test
            If USINGBUTANEMASSLIMIT Then
                If (LoadControl(station, Shift).loadTotalGrams > (StationCanister(station, Shift).WorkingCapacity * StationConfig(station, Shift).ButaneMassLimit)) Then
                    ALM_Write station, Shift, "Butane Mass Limit Exceeded"
                    Write_ELog "Butane Mass Limit Exceeded...Aborting " & "  Station " & station & "  Shift" & Shift
                    Delay_Box "Butane Mass Limit Exceeded....Aborting", MSGDELAY, msgSHOW
                    Load_Abort station, Shift
                    Exit Sub
                End If
            End If
            If USINGLOADTIMELIMIT Then
                ' using load time limit
                ' calculate elapsed time
                deltatime = DateDiff("n", Now, Stn_LoadLimitStartTime(station, Shift))
                ' calculate time limit
                Select Case StationRecipe(station, Shift).Load_Method
                    Case LOADBYTIME
                        maxtime = StationRecipe(station, Shift).Load_Time
                    Case LOADBYWC
                        maxtime = 60 * StationRecipe(station, Shift).EPAFill
                    Case LOADBYWEIGHT
                        maxmass = StationRecipe(station, Shift).Load_Wt
                        minrate = StationRecipe(station, Shift).Load_Rate - SysConfig.Tol_Btn_Flow
                        maxtime = (maxmass / minrate) * 60
                    Case LOADBYBREAKTHRU
                        maxmass = StationCanister(station, Shift).WorkingCapacity + StationRecipe(station, Shift).LoadBreakthrough
                        minrate = StationRecipe(station, Shift).Load_Rate - SysConfig.Tol_Btn_Flow
                        maxtime = (maxmass / minrate) * 60
                    Case LOADBYFID
    '                    maxmass = StationCanister(station, Shift).WorkingCapacity + (ValueFromText(txtFIDmg.text) / 1000)
                        maxmass = StationCanister(station, Shift).WorkingCapacity + StationRecipe(station, Shift).LoadBreakthrough
                        minrate = StationRecipe(station, Shift).Load_Rate - SysConfig.Tol_Btn_Flow
                        maxtime = (maxmass / minrate) * 60
                    Case Else
                        maxtime = deltatime
                End Select
                If (deltatime > (maxtime * StationConfig(station, Shift).LoadTimeLimit)) Then
                    ALM_Write station, Shift, "Load Time Limit Exceeded"
                    Write_ELog "Load Time Limit Exceeded...Aborting " & "  Station " & station & "  Shift" & Shift
                    Delay_Box "Load Time Limit Exceeded....Aborting", MSGDELAY, msgSHOW
                    Load_Abort station, Shift
                    Exit Sub
                End If
            End If
            
            ' check for valid AIO and clip-to-zero
            Select Case STN_INFO(station).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                    ChgErrModule 2, 10041
                    If Stn_AIO(station, asNitrogenFlow).EUValue <> Empty Then
                        If Stn_AIO(station, asNitrogenFlow).EUValue < -100 Or Stn_AIO(station, asNitrogenFlow).EUValue > 400000 Then
                            If Pause_Alarm = NOTPAUSED Then
                                Write_ELog "Questionable nitrogen flow rate? " & "  Station " & station & "  Shift" & Shift
                            End If
                        Else
                            If Stn_AIO(station, asNitrogenFlow).EUValue < 0 Then Stn_AIO(station, asNitrogenFlow).EUValue = 0
                        End If
                    End If
                    If Stn_AIO(station, asButaneFlow).EUValue <> Empty Then
                        If Stn_AIO(station, asButaneFlow).EUValue < -100 Or Stn_AIO(station, asButaneFlow).EUValue > 400000 Then
                            If Pause_Alarm = NOTPAUSED Then
                                Write_ELog "Questionable butane flow rate? " & "  Station " & station & "  Shift" & Shift
                            End If
                        Else
                            If Stn_AIO(station, asButaneFlow).EUValue < 0 Then Stn_AIO(station, asButaneFlow).EUValue = 0
                        End If
                    End If
                    
                Case STN_ORVR2_TYPE
                    ChgErrModule 2, 10042
                    If StationRecipe(station, Shift).UseHiRangeMFC Then
                        ' use higher range MFC
                        If Stn_AIO(station, asNitrogenORVRFlow).EUValue <> Empty Then
                            If Stn_AIO(station, asNitrogenORVRFlow).EUValue < -100 Or Stn_AIO(station, asNitrogenORVRFlow).EUValue > 400000 Then
                                If Pause_Alarm = NOTPAUSED Then
                                    Write_ELog "Questionable nitrogen flow rate? " & "  Station " & station & "  Shift" & Shift
                                End If
                            Else
                                If Stn_AIO(station, asNitrogenORVRFlow).EUValue < 0 Then Stn_AIO(station, asNitrogenORVRFlow).EUValue = 0
                            End If
                        End If
                        If Stn_AIO(station, asButaneORVRFlow).EUValue <> Empty Then
                            If Stn_AIO(station, asButaneORVRFlow).EUValue < -100 Or Stn_AIO(station, asButaneORVRFlow).EUValue > 400000 Then
                                If Pause_Alarm = NOTPAUSED Then
                                    Write_ELog "Questionable butane flow rate? " & "  Station " & station & "  Shift" & Shift
                                End If
                            Else
                                If Stn_AIO(station, asButaneORVRFlow).EUValue < 0 Then Stn_AIO(station, asButaneORVRFlow).EUValue = 0
                            End If
                        End If
                    Else
                        ' use lower range MFC
                        If Stn_AIO(station, asNitrogenFlow).EUValue <> Empty Then
                            If Stn_AIO(station, asNitrogenFlow).EUValue < -100 Or Stn_AIO(station, asNitrogenFlow).EUValue > 400000 Then
                                If Pause_Alarm = NOTPAUSED Then
                                    Write_ELog "Questionable nitrogen flow rate? " & "  Station " & station & "  Shift" & Shift
                                End If
                            Else
                                If Stn_AIO(station, asNitrogenFlow).EUValue < 0 Then Stn_AIO(station, asNitrogenFlow).EUValue = 0
                            End If
                        End If
                        If Stn_AIO(station, asButaneFlow).EUValue <> Empty Then
                            If Stn_AIO(station, asButaneFlow).EUValue < -100 Or Stn_AIO(station, asButaneFlow).EUValue > 400000 Then
                                If Pause_Alarm = NOTPAUSED Then
                                    Write_ELog "Questionable butane flow rate? " & "  Station " & station & "  Shift" & Shift
                                End If
                            Else
                                If Stn_AIO(station, asButaneFlow).EUValue < 0 Then Stn_AIO(station, asButaneFlow).EUValue = 0
                            End If
                        End If
                    End If
                        
                Case STN_LIVEFUEL_TYPE
                    ChgErrModule 2, 10043
                    '   Station uses Live Fuel
                    If StationRecipe(station, Shift).UseLoadRatePID Then LoadController station
                    If Stn_AIO(station, asLiveFuelVaporFlow).EUValue < 0 Then Stn_AIO(station, asLiveFuelVaporFlow).EUValue = 0
                
                Case STN_LIVEREG_TYPE
                    ChgErrModule 2, 10044
                    '   Station uses Live Fuel AND Nit/Btn
                    If StationRecipe(station, Shift).LiveFuel Then
                        ' using live fuel
                        If StationRecipe(station, Shift).UseLoadRatePID Then LoadController station
                        If Stn_AIO(station, asLiveFuelVaporFlow).EUValue < 0 Then Stn_AIO(station, asLiveFuelVaporFlow).EUValue = 0
                    Else
                        ' using nitrogen/butane
                        If Stn_AIO(station, asNitrogenFlow).EUValue <> Empty Then
                            If Stn_AIO(station, asNitrogenFlow).EUValue < -100 Or Stn_AIO(station, asNitrogenFlow).EUValue > 400000 Then
                                If Pause_Alarm = NOTPAUSED Then
                                    Write_ELog "Questionable nitrogen flow rate? " & "  Station " & station & "  Shift" & Shift
                                End If
                            Else
                                If Stn_AIO(station, asNitrogenFlow).EUValue < 0 Then Stn_AIO(station, asNitrogenFlow).EUValue = 0
                            End If
                        End If
                        If Stn_AIO(station, asButaneFlow).EUValue <> Empty Then
                            If Stn_AIO(station, asButaneFlow).EUValue < -100 Or Stn_AIO(station, asButaneFlow).EUValue > 400000 Then
                                If Pause_Alarm = NOTPAUSED Then
                                    Write_ELog "Questionable butane flow rate? " & "  Station " & station & "  Shift" & Shift
                                End If
                            Else
                                If Stn_AIO(station, asButaneFlow).EUValue < 0 Then Stn_AIO(station, asButaneFlow).EUValue = 0
                            End If
                        End If
                    End If
                
                Case STN_LIVEORVR2_TYPE
                    ChgErrModule 2, 10045
                    '   Station uses Live Fuel AND Nit/Btn
                    If StationRecipe(station, Shift).LiveFuel Then
                        ' using live fuel
                        If StationRecipe(station, Shift).UseLoadRatePID Then LoadController station
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            If Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue <> Empty Then
                                If Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue < -100 Or Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue > 400000 Then
                                    If Pause_Alarm = NOTPAUSED Then
                                        Write_ELog "Questionable LiveFuelVapor ORVR flow rate? " & "  Station " & station & "  Shift" & Shift
                                    End If
                                Else
                                    If Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue < 0 Then Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue = 0
                                End If
                            End If
                        Else
                            ' use lower range MFC
                            If Stn_AIO(station, asLiveFuelVaporFlow).EUValue <> Empty Then
                                If Stn_AIO(station, asLiveFuelVaporFlow).EUValue < -100 Or Stn_AIO(station, asLiveFuelVaporFlow).EUValue > 400000 Then
                                    If Pause_Alarm = NOTPAUSED Then
                                        Write_ELog "Questionable LiveFuelVapor flow rate? " & "  Station " & station & "  Shift" & Shift
                                    End If
                                Else
                                    If Stn_AIO(station, asLiveFuelVaporFlow).EUValue < 0 Then Stn_AIO(station, asLiveFuelVaporFlow).EUValue = 0
                                End If
                            End If
                        End If
                    Else
                        ' using nitrogen/butane
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            If Stn_AIO(station, asNitrogenORVRFlow).EUValue <> Empty Then
                                If Stn_AIO(station, asNitrogenORVRFlow).EUValue < -100 Or Stn_AIO(station, asNitrogenORVRFlow).EUValue > 400000 Then
                                    If Pause_Alarm = NOTPAUSED Then
                                        Write_ELog "Questionable nitrogen flow rate? " & "  Station " & station & "  Shift" & Shift
                                    End If
                                Else
                                    If Stn_AIO(station, asNitrogenORVRFlow).EUValue < 0 Then Stn_AIO(station, asNitrogenORVRFlow).EUValue = 0
                                End If
                            End If
                            If Stn_AIO(station, asButaneORVRFlow).EUValue <> Empty Then
                                If Stn_AIO(station, asButaneORVRFlow).EUValue < -100 Or Stn_AIO(station, asButaneORVRFlow).EUValue > 400000 Then
                                    If Pause_Alarm = NOTPAUSED Then
                                        Write_ELog "Questionable butane flow rate? " & "  Station " & station & "  Shift" & Shift
                                    End If
                                Else
                                    If Stn_AIO(station, asButaneORVRFlow).EUValue < 0 Then Stn_AIO(station, asButaneORVRFlow).EUValue = 0
                                End If
                            End If
                        Else
                            ' use lower range MFC
                            If Stn_AIO(station, asNitrogenFlow).EUValue <> Empty Then
                                If Stn_AIO(station, asNitrogenFlow).EUValue < -100 Or Stn_AIO(station, asNitrogenFlow).EUValue > 400000 Then
                                    If Pause_Alarm = NOTPAUSED Then
                                        Write_ELog "Questionable nitrogen flow rate? " & "  Station " & station & "  Shift" & Shift
                                    End If
                                Else
                                    If Stn_AIO(station, asNitrogenFlow).EUValue < 0 Then Stn_AIO(station, asNitrogenFlow).EUValue = 0
                                End If
                            End If
                            If Stn_AIO(station, asButaneFlow).EUValue <> Empty Then
                                If Stn_AIO(station, asButaneFlow).EUValue < -100 Or Stn_AIO(station, asButaneFlow).EUValue > 400000 Then
                                    If Pause_Alarm = NOTPAUSED Then
                                        Write_ELog "Questionable butane flow rate? " & "  Station " & station & "  Shift" & Shift
                                    End If
                                Else
                                    If Stn_AIO(station, asButaneFlow).EUValue < 0 Then Stn_AIO(station, asButaneFlow).EUValue = 0
                                End If
                            End If
                        End If
                    End If
                
                Case STN_COMBO3_TYPE
                    ChgErrModule 2, 10046
                    '   future
                
                Case Else
                ' Do Nothing
            End Select
            
            ChgErrModule 2, 10047
    
            ' Update Concordance at Equalibrium start Values
            If StationControl(station, Shift).TestTimer <= Stn_LoadEql_StartTimer(station, Shift) Then
                ' not at equalibrium yet
                ' update initial values
                Stn_LoadEql_StartAuxWt(station, Shift) = StationControl(station, Shift).AuxScaleWt
                Stn_LoadEql_StartPriWt(station, Shift) = StationControl(station, Shift).PriScaleWt
                Stn_LoadEql_StartLoadTotal(station, Shift) = LoadControl(station, Shift).loadTotalGrams
            End If
            
            ChgErrModule 2, 10147
    
            ' Check for "dead" LiveFuel
            If StationControl(station, Shift).TestTimer <= Stn_LoadEql_StartTimer(station, Shift) Then
                ' not at equalibrium yet
                ' update initial values
                Stn_LoadEql_StartAuxWt(station, Shift) = StationControl(station, Shift).AuxScaleWt
                Stn_LoadEql_StartPriWt(station, Shift) = StationControl(station, Shift).PriScaleWt
                Stn_LoadEql_StartLoadTotal(station, Shift) = LoadControl(station, Shift).loadTotalGrams
            End If
                       
            
            ' Are we done yet?
            Select Case StationRecipe(station, Shift).Load_Method
                Case LOADBYTIME
                    ChgErrModule 2, 10048
                    ' If time is up then end load by time
                    time_seconds = 0#
                    time_seconds = DateDiff("s", PreviousNow(station, Shift), Now()) ' Date diff in seconds
                    If time_seconds > 0# Then
                        PreviousNow(station, Shift) = Now()
                        StationControl(station, Shift).Actual = StationControl(station, Shift).Actual + (time_seconds / 60)
                    End If
                    If StationControl(station, Shift).Actual > StationControl(station, Shift).Target Then     'all done O.K.
                        StationControl(station, Shift).ActualAtEnd = StationControl(station, Shift).Actual
                        LoadControl(station, Shift).AuxWtChgAtEOL = LoadControl(station, Shift).AuxWtChg
                        LoadControl(station, Shift).PriWtChgAtEOL = LoadControl(station, Shift).PriWtChg
                        LoadControl(station, Shift).TotalWtChgAtEOL = LoadControl(station, Shift).TotalWtChg
                        ChgPhase LoadComplete, Now, station, Shift
                    End If
                Case LOADBYWC
                    ChgErrModule 2, 10049
                    ' If W.C. amount used then
                    StationControl(station, Shift).Actual = LoadControl(station, Shift).loadTotalGrams
                    If LoadControl(station, Shift).loadTotalGrams >= ((StationCanister(station, Shift).WorkingCapacity * StationRecipe(station, Shift).WC_Mult)) Then
                        StationControl(station, Shift).ActualAtEnd = StationControl(station, Shift).Actual
                        LoadControl(station, Shift).AuxWtChgAtEOL = LoadControl(station, Shift).AuxWtChg
                        LoadControl(station, Shift).PriWtChgAtEOL = LoadControl(station, Shift).PriWtChg
                        LoadControl(station, Shift).TotalWtChgAtEOL = LoadControl(station, Shift).TotalWtChg
                        ChgPhase LoadComplete, Now, station, Shift
                    End If
                Case LOADBYWEIGHT
                    ChgErrModule 2, 10050
                    ' Check Pri Scale Value
                    StationControl(station, Shift).Actual = LoadControl(station, Shift).PriWtChg
                    ' Done Yet?
                    If Now > LoadControl(station, Shift).PhaseDts Then
                        If StationControl(station, Shift).Actual >= StationRecipe(station, Shift).Load_Wt Then
                            StationControl(station, Shift).ActualAtEnd = StationControl(station, Shift).Actual
                            LoadControl(station, Shift).AuxWtChgAtEOL = LoadControl(station, Shift).AuxWtChg
                            LoadControl(station, Shift).PriWtChgAtEOL = LoadControl(station, Shift).PriWtChg
                            LoadControl(station, Shift).TotalWtChgAtEOL = LoadControl(station, Shift).TotalWtChg
                            ' Write Load Data to File
                            Load_Write station, Shift, NORMALUPDATE
                            ChgPhase LoadComplete, Now, station, Shift
                        End If
                    End If
                Case LOADBYBREAKTHRU
                    ChgErrModule 2, 10051
                    ' Check Aux Scale Value
                    StationControl(station, Shift).Actual = LoadControl(station, Shift).AuxWtChg
                    ' Done Yet?
                    If Now > LoadControl(station, Shift).PhaseDts Then
                        If StationControl(station, Shift).Actual >= StationRecipe(station, Shift).LoadBreakthrough Then
                            StationControl(station, Shift).ActualAtEnd = StationControl(station, Shift).Actual
                            LoadControl(station, Shift).AuxWtChgAtEOL = LoadControl(station, Shift).AuxWtChg
                            LoadControl(station, Shift).PriWtChgAtEOL = LoadControl(station, Shift).PriWtChg
                            LoadControl(station, Shift).TotalWtChgAtEOL = LoadControl(station, Shift).TotalWtChg
                            Write_ELog "Station #" & Format(station, "#0") & " Breakthrough Actual = " & Format(StationControl(station, Shift).Actual, "##0.000") & " grams"
                            Write_ELog "Station #" & Format(station, "#0") & " Aux Wt Chg Actual = " & Format(LoadControl(station, Shift).AuxWtChgAtEOL, "##0.000") & " grams"
                            ' Write Load Data to File
                            Load_Write station, Shift, NORMALUPDATE
                            ChgPhase LoadComplete, Now, station, Shift
                        End If
                    End If
                Case LOADBYFID
                    ChgErrModule 2, 10052
'                    TestValueGrams = ((((KFactor * (AmbBaro / 33.865)) * BoxVolume * 0.001) / _
'                        (Fid_AIO(aaFIDTC).EUValue + 459.67)) * ((FIDOutputReadAdjusted - Background) - RFactor)) * 1000
'                    StationControl(station, Shift).Actual = TestValueGrams
'                    If TestValueGrams >= CInt(StationRecipe(station, Shift).FIDmg) Then
'                        StationControl(station, Shift).ActualAtEnd = StationControl(station, Shift).Actual
'                        LoadControl(station, Shift).TotalWtChgAtEOL = LoadControl(station, Shift).TotalWtChg
'                        ChgPhase LoadComplete, Now, station, Shift
'                    End If
            End Select
            
        Case LoadComplete
            '
            ' turn off load cycle mfc's
            '
            ' reset LoadRate Controller PID (the PID control will set Enable true)
            PID_INFO(iController).Enable = False
            ' station type variations
            Select Case STN_INFO(station).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                    Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                    Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                Case STN_ORVR2_TYPE
                    If StationRecipe(station, Shift).UseHiRangeMFC Then
                        ' use higher range MFC
                        Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
                        Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
                    Else
                        ' use lower range MFC
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                        Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                    End If
                Case STN_LIVEFUEL_TYPE
                    Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
                Case STN_LIVEREG_TYPE
                    If StationRecipe(station, Shift).LiveFuel Then
                        ' use Live Fuel
                        Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                        Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                    Else
                        ' use Butane/Nitrogen
                        Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
                        Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                        Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                    End If
                Case STN_LIVEORVR2_TYPE
                    If StationRecipe(station, Shift).LiveFuel Then
                        ' use Live Fuel
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, 0, outZERO
'                            Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
'                            Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog station, asLiveFuelVaporFlowSP, 0, outZERO
'                            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
'                            Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                        End If
                    Else
                        ' use Butane/Nitrogen
                        If StationRecipe(station, Shift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog station, asNitrogenORVRFlowSP, 0, outZERO
                            Stn_OutAnalog station, asButaneORVRFlowSP, 0, outZERO
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog station, asNitrogenFlowSP, 0, outZERO
                            Stn_OutAnalog station, asButaneFlowSP, 0, outZERO
                        End If
                    End If
                Case STN_COMBO3_TYPE
                    ' future
                Case Else
                    ' nothing to do
            End Select
            ' Write Load Data to File
            Load_Write station, Shift, NORMALUPDATE
            ' update LoadControlBlock
            ChgPhase LoadStopping, (Now + TimeSerial(0, 0, LoadMfcDelayTime)), station, Shift
           
        Case LoadStopping
            '
            ' after delay, turn off load cycle valves
            '
            If Now > LoadControl(station, Shift).PhaseDts Then
                Close_Stn_Valves station, Shift
                ChgPhase LoadPause, (Now + MinutesFromNow(StationConfig(station, Shift).LoadSettleTime)), station, Shift
            End If
        
        Case LoadPause
            '
            ' after scale values settle, end this load cycle
            '
            If Now > LoadControl(station, Shift).PhaseDts Then
                Load_Done station, Shift
            End If
    
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
    
Sub ChgPhase(ByVal newPhase As Integer, ByVal newPhaseDts As Date, ByVal iStn As Integer, ByVal iShift As Integer)
'
    Select Case StationControl(iStn, iShift).Mode
        Case VBLEAK
            With LeakCheckControl
                .Phase = newPhase
                .PhaseDts = newPhaseDts
                .PhaseStartDts = Now
                .PhaseStartTimer = StationControl(iStn, iShift).TestTimer
            End With
        Case VBLOAD
            With LoadControl(iStn, iShift)
                .Phase = newPhase
                .PhaseDts = newPhaseDts
                .PhaseStartDts = Now
                .PhaseStartTimer = StationControl(iStn, iShift).TestTimer
                If (newPhase = LoadStarting) Then LoadControl(iStn, iShift).ElapsedStartDts = Now
            End With
        Case VBPURGE
            With PurgeControl(iStn, iShift)
                .Phase = newPhase
                .PhaseDts = newPhaseDts
                .PhaseStartDts = Now
                .PhaseStartTimer = StationControl(iStn, iShift).TestTimer
            End With
        Case Else
            ' nothing to do
            Exit Sub
    End Select
    
End Sub

Public Sub Station_Finish(ByVal iStn As Integer, ByVal iShift As Integer)
' Routine Name: Station_Finish
' Function:
' This routine finishes all activites on the station after the station
' has been reset to the idle mode, or if the station has been closed
' by terminating the program.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' We finished a Process produce reports for that one
' We are all done
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rptgenCmdInt As Integer
Dim cmdLine As String
Dim rptgenCmdCode
Dim cfgBits(0 To 9) As Boolean
Dim sourcefile As String
Dim destfile As String
Dim filename As String
Dim sPrint As String
Dim Idx, idx2 As Integer
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")

SetErrModule 2, 9
If UseLocalErrorHandler Then On Error GoTo localhandler

'Debug_Dts(1) = Now()
'Debug_Comment(1) = "Start Finish for Station #" & Format(iStn, "0") & " and Shift " & Format(iShift, "0")
'Debug_Comment(2) = "   Start Recipe_Write"
'Debug_Comment(3) = "   Start Update_Joblist"
'Debug_Comment(4) = "   Start Header_Write"
'Debug_Comment(5) = "   Start Report_Detail"
'Debug_Comment(6) = "   Start Report_FID"
'Debug_Comment(7) = "   Start Report_Thermocouple"
'Debug_Comment(8) = "   Start Report_Summary"
'Debug_Comment(9) = "   Start Backup_Active"
'Debug_Comment(10) = "   Start FileCopy"
'Debug_Comment(11) = "   Start Finish_Clear"
'Debug_Comment(12) = "   Start StationRecipe(station, Shift).UseAnalyzer??"
'Debug_Comment(13) = "   Start zReport_Debug"
'Debug_Comment(14) = ""
'Debug_Comment(15) = ""
'Debug_Comment(16) = ""
'Debug_Comment(17) = ""
'Debug_Comment(18) = ""
'Debug_Comment(19) = ""
filename = Left(StationControl(iStn, iShift).RptFile, 50)
    
'  *****  ADDED  *****************
' update TOM Task Status to Done(if required)
If (Len(StnTomTask(iStn, iShift).TaskID) > 2) Then
    If (StnTomTask(iStn, iShift).TaskStatus = "Active") Then
        If (JobInfo(iStn, iShift).End_OK) Then
            ' update TOM Task status in DB
            TomTask_Update iStn, iShift, "Done", "na"
        Else
            ' update TOM Task status in DB
            TomTask_Update iStn, iShift, "Ready", StnTomTask(iStn, iShift).PreviousResult
        End If
    Else
        ' update TOM Task status in DB
        TomTask_Update iStn, iShift, "Ready", StnTomTask(iStn, iShift).PreviousResult
    End If
End If
'  *****  ADDED  *****************

' Testing is COMPLETE; reports,etc. aren't (yet)
StationControl(iStn, iShift).Mode = VBCOMPLETE
StationControl(iStn, iShift).TestTimerIsRunning = False
' ****  R5  StationControl(iStn, iShift).End_Timer = TestTimer(iStn, iShift)
' ****  R5 TestTimerIsRunning(iStn, iShift) = False
    
    
' No Name; No Reports; No DB updates
If (Len(StationControl(iStn, iShift).RptFile) <> 0 And StationControl(iStn, iShift).DBFile <> "") Then
    
    ' Make sure Alarm or OOT log screen isn't up;  DB errors if so
    If frmDataLog.Visible Then
        If frmDataLog.LogData = "Alarm" Then frmDataLog.cmdReturn.Value = True
        If frmDataLog.LogData = "OOT" Then frmDataLog.cmdReturn.Value = True
        DoEvents
    End If
        
    ' Update joblist with stop time, etc.
    Debug_Dts(3) = Now()
    Update_Joblist iStn, iShift
        
        
        ' Slide stn XY Graph to the left
        frmStnDetail.SlideXYgraph iStn, iShift
        
'  ****  ADDED
    ' Update TomCanLoad, if required
    If (USINGTOMCANLOAD) Then
        ' Was the TOM Task completed successfully
        If (Len(StnTomTask(iStn, iShift).TaskID) > 3) Then
            If Not JobInfo(iStn, iShift).End_OK Then
                ' not completed; update event log
                sPrint = "TOM Task >" & StnTomTask(iStn, iShift).TaskID
                sPrint = sPrint & "< (Job# " & StationControl(iStn, iShift).Job_Number
                sPrint = sPrint & ") did not complete successfully."
                Write_ELog sPrint
            Else
                ' TOM Task completed successfully; update event log
                sPrint = "TOM Task " & StnTomTask(iStn, iShift).TaskID
                sPrint = sPrint & " (Job# " & StationControl(iStn, iShift).Job_Number
                sPrint = sPrint & ") completed successfully."
                Write_ELog sPrint
                ' clear any duplicate TomTaskIDs & VINs on other station/shift's
                For Idx = 1 To NR_STN
                    For idx2 = 1 To NR_SHIFT
                        If (StnTomTask(Idx, idx2).TaskID = StnTomTask(iStn, iShift).TaskID) Then
                            StnTomTask(Idx, idx2).TaskID = "na"
                            StnTomTask(Idx, idx2).VIN = "na"
                            JobInfo(Idx, idx2).Vehicle = "na"
                        End If
                    Next idx2
                Next Idx
                ' clear TOM Data for this Station/Shift
                StnTomTask(iStn, iShift) = TomData_Clear
            End If
        End If
    End If
'  ****  ADDED
        
        ChgErrModule 2, 9990
        If StationConfig(iStn, iShift).DbFileBackup_Active Then
            If fs.FolderExists(StationConfig(iStn, iShift).DbFileBackup_Path) Then
                
                ChgErrModule 2, 9991
                ' Avoid DB errors if Review screen still has DB open
                If frmReview.Visible Then
                    ChgErrModule 2, 9992
                    frmReview.JobComplete iStn, iShift
                    ChgErrModule 2, 9993
                    DoEvents
                End If
        
                ChgErrModule 2, 9994
                ' db file
                sourcefile = FILEPATH_data & "C" & StationControl(iStn, iShift).Job_Number & AccessDbFileExt
                destfile = StationConfig(iStn, iShift).DbFileBackup_Path & "C" & StationControl(iStn, iShift).Job_Number & AccessDbFileExt
    '            Debug_Dts(10) = Now()
                FileCopy sourcefile, destfile
                
            Else  ' backup path doesn't exist
                
                ChgErrModule 2, 9997
                Delay_Box "Backup Path >" & StationConfig(iStn, iShift).DbFileBackup_Path & "< Not defined; ABORTING copy", MSGDELAY, msgSHOW
    '            ALM_Write iStn, iShift, "DB File Backup path not defined >" & StationConfig(iStn, iShift).DbFileBackup_Path & "< DB File " & "C" & StationControl(iStn, iShift).Job_Number & " not copied "
            
            End If
        
        End If
    
        ChgErrModule 2, 9998
        ' print reports, if required
        If (Not PRINTERAVAILABLE And StationConfig(iStn, iShift).RptConfig.TextEotSummary_AutoPrint) Then
            ' no printer
    '        Delay_Box "No Printer Available for AutoPrint", MSGDELAY, msgSHOW
            Write_ELog "No Printer Available for AutoPrint of Job#" & StationControl(iStn, iShift).Job_Number & " on Station " & Format(iStn, "0") & " Shift " & Format(iShift, "0")
        End If
        ' combine report config Bits
        '
        '   bit 0 = TextReporting
        '   bit 1 = TextSummary
        '   bit 2 = TextSummary_AutoPrint
        '   bit 3 = TextDetail
        '
        '   bit 4 = XlsReporting
        '   bit 5 = XlsSummary
        '   bit 6 = XlsDetail
        '
        '   bit 7 = CsvReporting
        '   bit 8 = CsvSummary
        '   bit 9 = CsvDetail
        '
        '   bit 10 = unused
        '   bit 11 = unused
        '   bit 12 = unused
        '
        With StationConfig(iStn, iShift).RptConfig
            cfgBits(0) = .TextEotReporting
            cfgBits(1) = .TextEotSummary
            cfgBits(2) = IIf(PRINTERAVAILABLE, .TextEotSummary_AutoPrint, False)
            cfgBits(3) = .TextEotDetail
            cfgBits(4) = .XlsEotReporting
            cfgBits(5) = .XlsEotSummary
            cfgBits(6) = .XlsEotDetail
            cfgBits(7) = .CsvEotReporting
            cfgBits(8) = .CsvEotSummary
            cfgBits(9) = .CsvEotDetail
            rptgenCmdInt = Bits_Pack(cfgBits(0), cfgBits(1), cfgBits(2), cfgBits(3), cfgBits(4), cfgBits(5), cfgBits(6), cfgBits(7), cfgBits(8), cfgBits(9), False, False, False, False, False)
        End With
        rptgenCmdCode = Format(rptgenCmdInt, "000000")
    
            
        ' request reports be generated
        cmdLine = filepath & "\cps_r7_Reporter.exe  " & StationControl(iStn, iShift).Job_Number & "  " & rptgenCmdCode
        Shell cmdLine, vbNormalFocus
                            
        frmStnDetail.txtStnDtlMsg.text = vbCrLf & Trim(StationControl(iStn, iShift).Job_Description) & " has ended"
    
    End If
    
    ChgErrModule 2, 998

    ' update JobLog
    sPrint = "Recipe completed."
    Write_JLog iStn, iShift, sPrint

    ' Some general clears at end of a job
    Station_Clear iStn, iShift
    ' almost done
    StationControl(iStn, iShift).Mode = VBIDLEWAITING


ResetErrModule
Exit Sub

localhandler:
If err = 52 Then    ' Backup Path Not Found
    Delay_Box "Backup Path Not Found while copying file " & sourcefile, MSGDELAY, msgSHOW
    ALM_Write iStn, iShift, "Backup Path Not Found while copying file " & sourcefile
    Write_ELog "Error: " & err & ", M" & ErrModule(0) & "-L" & ErrLevel(0) & " " & error$(err)
    Write_ELog "Backup Path Not Found while copying file " & sourcefile
'    Finish_Clear iStn, iShift
'    ResetErrModule
'    Exit Sub
    Resume Next
ElseIf err = 70 Then    ' Permission Denied (usually the db file)
    Delay_Box "Permission Denied while copying file " & sourcefile, MSGDELAY, msgSHOW
    ALM_Write iStn, iShift, "Permission Denied while copying file " & sourcefile
    Write_ELog "Error: " & err & ", M" & ErrModule(0) & "-L" & ErrLevel(0) & " " & error$(err)
    Write_ELog "Permission Denied while copying file " & sourcefile
'    Finish_Clear iStn, iShift
'    ResetErrModule
'    Exit Sub
    Resume Next
Else
    Dim iresponse As Integer
    iresponse = ErrorHandler(err)
    Select Case iresponse
      Case vbAbort       ' Exit if abort
        ResetErrModule
        frmMainMenu.MousePointer = vbDefault
        frmDelayBox.MousePointer = vbDefault
        Exit Sub
      Case vbRetry       ' try error line again
        Resume
      Case vbIgnore      ' Skip to next line, try to ignore
        Resume Next
    End Select
End If

End Sub

Sub LoadRecipeToStation(ByVal iRecipe As Integer, ByVal iStation As Integer, ByVal iShift As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim iAux As Integer
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1616

    ' open canister / recipe database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
            
    ' Read Master Recipe Record
    Criteria = "SELECT * FROM [MasterRecipe] WHERE [Number] = " & iRecipe & " "
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
    If Not rsRecord.BOF Then
        ' Load Recipe Record to StationRecipe
        StationRecipe(iStation, iShift).Number = rsRecord("Number")
        StationRecipe(iStation, iShift).Name = rsRecord("Name")
        
        StationRecipe(iStation, iShift).Load_Method = rsRecord("Load_Method")
        StationRecipe(iStation, iShift).Load_MethodSave = StationRecipe(iStation, iShift).Load_Method
        StationRecipe(iStation, iShift).NitrogenFlow = rsRecord("NitrogenFlow")
        StationRecipe(iStation, iShift).NitrogenFlowSave = StationRecipe(iStation, iShift).NitrogenFlow
        StationRecipe(iStation, iShift).Load_Rate = rsRecord("Load_Rate")
        StationRecipe(iStation, iShift).Load_RateSave = StationRecipe(iStation, iShift).Load_Rate
        StationRecipe(iStation, iShift).Mix_Percent = rsRecord("Mix_Percent")
        StationRecipe(iStation, iShift).WC_Mult = rsRecord("WC_Mult")
        StationRecipe(iStation, iShift).WC_MultSave = StationRecipe(iStation, iShift).WC_Mult
        StationRecipe(iStation, iShift).EPAFill = rsRecord("EPAFill")
        StationRecipe(iStation, iShift).Load_Wt = rsRecord("Load_Wt")
        StationRecipe(iStation, iShift).LoadBreakthrough = rsRecord("LoadBreakthrough")
        StationRecipe(iStation, iShift).FIDmg = rsRecord("FIDmg")
        StationRecipe(iStation, iShift).Load_Time = rsRecord("Load_Time")
        If IsNumeric(rsRecord("Purge_Method")) Then
            StationRecipe(iStation, iShift).Purge_Method = rsRecord("Purge_Method")
        Else
            StationRecipe(iStation, iShift).Purge_Method = PURGEBYVOLUME
        End If
        If IsNumeric(rsRecord("Purge_AuxTime")) Then
            StationRecipe(iStation, iShift).Purge_AuxTime = rsRecord("Purge_AuxTime")
        Else
            StationRecipe(iStation, iShift).Purge_AuxTime = 0
        End If
        StationRecipe(iStation, iShift).Purge_Time = rsRecord("Purge_Time")
        StationRecipe(iStation, iShift).Purge_Flow = rsRecord("Purge_Flow")
        StationRecipe(iStation, iShift).Purge_Can_Vol = rsRecord("Purge_Can_Vol")
        StationRecipe(iStation, iShift).Purge_ProfileNumber = rsRecord("Purge_ProfileNumber")
        StationRecipe(iStation, iShift).Purge_TargetWC = rsRecord("Purge_TargetWC")
        StationRecipe(iStation, iShift).Purge_TargetWeight = rsRecord("Purge_TargetWeight")
        StationRecipe(iStation, iShift).Purge_MaxVolumes = rsRecord("Purge_MaxVolumes")
        StationRecipe(iStation, iShift).Purge_TargetPurge = rsRecord("Purge_TargetPurge")
        StationRecipe(iStation, iShift).Purge_TargetPause = rsRecord("Purge_TargetPause")
    
        StationRecipe(iStation, iShift).UseAuxScale = rsRecord("UseAuxScale")
        StationRecipe(iStation, iShift).PurgeAuxCan = rsRecord("PurgeAuxCan")
        StationRecipe(iStation, iShift).AuxScaleNo = rsRecord("AuxScaleNo")
        StationRecipe(iStation, iShift).PauseLeakTime = rsRecord("PauseLeakTime")
        StationRecipe(iStation, iShift).PauseLoadTime = rsRecord("PauseLoadTime")
        StationRecipe(iStation, iShift).PausePurgeTime = rsRecord("PausePurgeTime")
        StationRecipe(iStation, iShift).UsePriScale = rsRecord("UsePriScale")
        StationRecipe(iStation, iShift).PriScaleNo = rsRecord("PriScaleNo")
        StationRecipe(iStation, iShift).PauseAfterLeak = rsRecord("PauseAfterLeak")
        StationRecipe(iStation, iShift).PauseAfterLoad = rsRecord("PauseAfterLoad")
        StationRecipe(iStation, iShift).PauseAfterPurge = rsRecord("PauseAfterPurge")
'        StationRecipe(iStation, iShift).TargetConcentration = rsRecord("TargetConcentration")
'        StationRecipe(iStation, iShift).DwellTime = rsRecord("DwellTime")
        StationRecipe(iStation, iShift).LeakCheck = rsRecord("LeakCheck")
        StationRecipe(iStation, iShift).LeakPrimary = rsRecord("LeakPrimary")
        StationRecipe(iStation, iShift).LeakAux = rsRecord("LeakAux")
'        StationRecipe(iStation, iShift).UseAnalyzer = rsRecord("UseAnalyzer")
        StationRecipe(iStation, iShift).MaxLoadTime = rsRecord("MaxLoadTime")
        StationRecipe(iStation, iShift).UseHiRangeMFC = rsRecord("UseHiRangeMFC")
        StationRecipe(iStation, iShift).UseLoadRatePID = rsRecord("UseLoadRatePID")
        
        StationRecipe(iStation, iShift).IDLoad = rsRecord("IDLoad")
        StationRecipe(iStation, iShift).LoadL = rsRecord("LoadL")
        StationRecipe(iStation, iShift).LoadV = rsRecord("LoadV")
        StationRecipe(iStation, iShift).IDPurge = rsRecord("IDPurge")
        StationRecipe(iStation, iShift).PurgeL = rsRecord("PurgeL")
        StationRecipe(iStation, iShift).PurgeV = rsRecord("PurgeV")
        StationRecipe(iStation, iShift).IDVent = rsRecord("IDVent")
        StationRecipe(iStation, iShift).VentL = rsRecord("VentL")
        StationRecipe(iStation, iShift).VentV = rsRecord("VentV")
        
        StationRecipe(iStation, iShift).LiveFuel = rsRecord("LiveFuel")
        StationRecipe(iStation, iShift).LiveFuelChgAuto = rsRecord("LiveFuelChgAuto")
        StationRecipe(iStation, iShift).LiveFuelChgFreq = rsRecord("LiveFuelChgFreq")
        StationRecipe(iStation, iShift).ADF_Heater = rsRecord("ADF_Heater")
        StationRecipe(iStation, iShift).ADF_HeaterSP = rsRecord("ADF_HeaterSP")
        
        ' start method
        StationRecipe(iStation, iShift).StartMethod = rsRecord("StartMethod")
        StationRecipe(iStation, iShift).StartDelay = rsRecord("StartDelay")
        StationRecipe(iStation, iShift).StartDate = rsRecord("StartDate")
        
        ' end method
        StationRecipe(iStation, iShift).EndMethod = rsRecord("EndMethod")
        StationRecipe(iStation, iShift).Cycles = rsRecord("Cycles")
        StationRecipe(iStation, iShift).CyclesSave = StationRecipe(iStation, iShift).Cycles
        StationRecipe(iStation, iShift).EndWeightTolerance = rsRecord("EndWeightTolerance")
        StationRecipe(iStation, iShift).EndConsecutiveCycles = rsRecord("EndConsecutiveCycles")
        StationRecipe(iStation, iShift).EndMinimumCycles = rsRecord("EndMinimumCycles")
        
        ' cycle type
        StationRecipe(iStation, iShift).CycleType = rsRecord("CycleType")
       
        ' aux outputs
        StationRecipe(iStation, iShift).AuxOutputs = False
        For iAux = 1 To 4
            StationRecipe(iStation, iShift).AuxOutputs_Load(iAux) = False
            StationRecipe(iStation, iShift).AuxOutputs_Purge(iAux) = False
        Next iAux
        
    End If
    rsRecord.Close
    ' close canister / recipe database
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
  Case vbIgnore      ' close down
    Station_Clear iStation, iShift
    StationControl(iStation, iShift).Mode = VBIDLEWAITING
    Exit Sub
End Select
End Sub

Sub LoadProfileToStation(ByVal iProfile As Integer, ByVal iStation As Integer, ByVal iShift As Integer)
Dim dbDbase As Database
Dim rsProfile  As Recordset
Dim rsSteps  As Recordset
Dim Criteria As String
Dim iStep As Integer
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1617

    ' clear station profile
    StationProfile(iStation, iShift).Number = CInt(0)
    StationProfile(iStation, iShift).Description = "undefined"
    StationProfile(iStation, iShift).Duration = CSng(0)
    StationProfile(iStation, iShift).DurDesc = "undefined"
    StationProfile(iStation, iShift).EndStep = CInt(0)
    StationProfile(iStation, iShift).ProjectedLiters = CSng(0)
    StationProfile(iStation, iShift).ProjectedVolumes = CSng(0)
    StationProfile(iStation, iShift).Validated = False
    ' steps
    For iStep = 1 To MAX_PROFILESTEPS
        StationProfile(iStation, iShift).StepDuration(iStep) = CSng(0)
        StationProfile(iStation, iShift).StepStartSetpoint(iStep) = CSng(0)
        StationProfile(iStation, iShift).StepType(iStep) = CInt(0)
    Next iStep


    ' open canister / recipe database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
            
    ' Read Master Profile Record
    Criteria = "SELECT * FROM [MasterProfiles] WHERE [Number] = " & iProfile & " "
    Set rsProfile = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    If Not rsProfile.BOF Then
        
        StationProfile(iStation, iShift).Number = rsProfile("Number")
        StationProfile(iStation, iShift).Description = rsProfile("Description")
        
        StationProfile(iStation, iShift).Duration = rsProfile("TotalDuration")
        StationProfile(iStation, iShift).DurDesc = DurationDescription(StationProfile(iStation, iShift).Duration)
        StationProfile(iStation, iShift).EndStep = rsProfile("Steps")
        StationProfile(iStation, iShift).ProjectedLiters = rsProfile("ProjectedLiters")
        StationProfile(iStation, iShift).ProjectedVolumes = rsProfile("ProjectedVolumes")
        StationProfile(iStation, iShift).Validated = True
        
        ' Read Master Profile Steps Information Records
        Criteria = "SELECT * FROM [MasterProfileSteps] WHERE [ProfileNumber] = " & iProfile & " "
        Set rsSteps = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        If (Not rsSteps.BOF) Then
            rsSteps.MoveFirst
            While Not rsSteps.EOF
                iStep = rsSteps("StepNumber")
                StationProfile(iStation, iShift).StepDuration(iStep) = rsSteps("Duration")
                StationProfile(iStation, iShift).StepStartSetpoint(iStep) = rsSteps("InitialSP")
                StationProfile(iStation, iShift).StepType(iStep) = rsSteps("StepType")
                rsSteps.MoveNext
            Wend
        End If
        
        rsSteps.Close

    End If
       
    rsProfile.Close
    
    ' close canister / recipe database
    dbDbase.Close

    ' save Station Purge Profiles
    Save_StationProfiles

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
  Case vbIgnore      ' close down
    Station_Clear iStation, iShift
    StationControl(iStation, iShift).Mode = VBIDLEWAITING
    Exit Sub
End Select
End Sub

Sub Station_Clear(ByVal iStn As Integer, ByVal iShift As Integer)
'
' Some general clears at end of a process
'
'
    If StationControl(iStn, iShift).ScalesInUse Then
        If StationRecipe(iStn, iShift).UseAuxScale Then
           Scale_In_Use(StationRecipe(iStn, iShift).AuxScaleNo) = False
        End If
        If StationRecipe(iStn, iShift).UsePriScale Then
           Scale_In_Use(StationRecipe(iStn, iShift).PriScaleNo) = False
        End If
        StationControl(iStn, iShift).ScalesInUse = False
    End If
    StationControl(iStn, iShift).DBFile = ""
    StationControl(iStn, iShift).RptFile = ""
    StationControl(iStn, iShift).Job_Number = ""
    JobInfo(iStn, iShift).Engineer = ""
    JobInfo(iStn, iShift).Vehicle = ""
    JobInfo(iStn, iShift).Start_Op = ""
    JobInfo(iStn, iShift).End_Op = ""
    JobInfo(iStn, iShift).Comment = ""
    Stn_UseTC(iStn, iShift) = False
    SEQ_Nmbr(iStn, iShift) = seqIdle              ' idle
    SEQ_Step(iStn, iShift) = 0
    SEQ_Alarm(iStn, iShift) = False
    SEQ_OOT(iStn, iShift) = False
    Clear_Stats iStn, iShift, 3
    Reset_Bar_Graph iStn, iShift              ' reset the bar graphs
    
End Sub

Sub Course_Init(ByVal iStation As Integer, ByVal iShift As Integer, ByVal iCourse As Integer)
' Function Name:    Course_Init
' Author:           Brunrose
' Description:      Initializes a Course.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1515
Dim iRecipe As Integer
Dim iProfile As Integer
Dim sMsg As String

    StationControl(iStation, iShift).Course = iCourse
    StationSequence(iStation, iShift).CourseData(iCourse).DtsStart = Now()
    Select Case StationSequence(iStation, iShift).CourseData(iCourse).Type
        Case courseWait
            ' Wait for Operator OK
            StationSequence(iStation, iShift).CourseData(iCourse).OkToProceed = False
            StationControl(iStation, iShift).Mode = VBCOURSEWAIT
            ' Job Description
            StationControl(iStation, iShift).Job_Description = Trim(StationSequence(iStation, iShift).Description)
            ' estimated Job Duration
            StationControl(iStation, iShift).EstJobDur = StationSequence(iStation, iShift).EstSeqDuration
            StationControl(iStation, iShift).EstJobDurDesc = DurationDescription(StationControl(iStation, iShift).EstJobDur)
        Case coursePause
            ' Pause for x Minutes
            StationSequence(iStation, iShift).CourseData(iCourse).OkToProceed = False
            StationControl(iStation, iShift).Mode = VBCOURSEPAUSE
            ' Job Description
            StationControl(iStation, iShift).Job_Description = Trim(StationSequence(iStation, iShift).Description)
            ' estimated Job Duration
            StationControl(iStation, iShift).EstJobDur = StationSequence(iStation, iShift).EstSeqDuration
            StationControl(iStation, iShift).EstJobDurDesc = DurationDescription(StationControl(iStation, iShift).EstJobDur)
        Case courseRecipe
            ' Run a Recipe
            iRecipe = StationSequence(iStation, iShift).CourseData(iCourse).RecipeNumber
            ' recipe 0 means use existing station recipe
            If iRecipe = 0 Then
                ' using existing Station Recipe
                ' if only 1 course then use recipe name for job description & use recipe duration for job duration
                Select Case StationSequence(iStation, iShift).NumCourses
                    Case 1
                        ' Job Description
                        If (STN_INFO(iStation).Type = STN_LEAKTEST_TYPE) Then
                            ' leaktest station type
                            StationControl(iStation, iShift).Job_Description = "40 CPR 1066.985 LeakTest"
                            ' estimated Job Duration
                            StationControl(iStation, iShift).EstJobDur = CSng(Rcp_LeakTest.HoldDuration + 15) / 60#
                            StationControl(iStation, iShift).EstJobDurDesc = DurationDescription(StationControl(iStation, iShift).EstJobDur)
                        Else
                            ' normal station type
                            StationControl(iStation, iShift).Job_Description = Trim(StationRecipe(iStation, iShift).Name)
                            ' optional changes
                            If (StationSequence(iStation, iShift).CourseData(iCourse).Cycles > 0) Then
                                StationRecipe(iStation, iShift).Cycles = StationSequence(iStation, iShift).CourseData(iCourse).Cycles
                                StationRecipe(iStation, iShift).CyclesSave = StationRecipe(iStation, iShift).Cycles
                            End If
                            If (StationSequence(iStation, iShift).CourseData(iCourse).LoadRate > 0) Then
                                StationRecipe(iStation, iShift).Load_Rate = StationSequence(iStation, iShift).CourseData(iCourse).LoadRate
                                StationRecipe(iStation, iShift).Load_RateSave = StationRecipe(iStation, iShift).Load_Rate
                            End If
                            If (StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate > 0) Then
                                StationRecipe(iStation, iShift).Purge_Flow = StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate
                            End If
                            ' estimated Job Duration
                            StationControl(iStation, iShift).EstJobDur = EstimatedRcpDuration(StationRecipe(iStation, iShift), StationCanister(iStation, iShift), StationProfile(iStation, iShift))
                            StationControl(iStation, iShift).EstJobDurDesc = DurationDescription(StationControl(iStation, iShift).EstJobDur)
                        End If
                    Case Else
                        ' Job Description
                        StationControl(iStation, iShift).Job_Description = Trim(StationSequence(iStation, iShift).Description)
                        ' optional changes
                        If (StationSequence(iStation, iShift).CourseData(iCourse).Cycles > 0) Then
                            StationRecipe(iStation, iShift).Cycles = StationSequence(iStation, iShift).CourseData(iCourse).Cycles
                            StationRecipe(iStation, iShift).CyclesSave = StationRecipe(iStation, iShift).Cycles
                        End If
                        If (StationSequence(iStation, iShift).CourseData(iCourse).LoadRate > 0) Then
                            StationRecipe(iStation, iShift).Load_Rate = StationSequence(iStation, iShift).CourseData(iCourse).LoadRate
                            StationRecipe(iStation, iShift).Load_RateSave = StationRecipe(iStation, iShift).Load_Rate
                        End If
                        If (StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate > 0) Then
                            StationRecipe(iStation, iShift).Purge_Flow = StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate
                        End If
                        ' estimated Job Duration
                        StationControl(iStation, iShift).EstJobDur = StationSequence(iStation, iShift).EstSeqDuration
                        StationControl(iStation, iShift).EstJobDurDesc = DurationDescription(StationControl(iStation, iShift).EstJobDur)
                End Select
                ' save station recipes
                Save_StationRecipes
                ' Update Station Recipe descriptors
                UpdateStnRcpDsc iStation, iShift
                ' Update Station LiveFuel setup
                LiveFuel_Update iStation, iShift
            Else
                ' normal job sequence
                ' replace the Station Recipe with a Master Recipe
                LoadRecipeToStation iRecipe, iStation, iShift
                ' replace the Station PurgeProfile with a Master PurgeProfile
                If (StationRecipe(iStation, iShift).Purge_Method = PURGEBYPROFILE) Then
                    iProfile = StationRecipe(iStation, iShift).Purge_ProfileNumber
                    LoadProfileToStation iProfile, iStation, iShift
                End If
                ' make Course adjustments to the Station Recipe
                StationRecipe(iStation, iShift).Number = iRecipe
                StationRecipe(iStation, iShift).PriScaleNo = StationSequence(iStation, iShift).PriScaleNo
                StationRecipe(iStation, iShift).AuxScaleNo = StationSequence(iStation, iShift).AuxScaleNo
                StationRecipe(iStation, iShift).UsePriScale = IIf(StationSequence(iStation, iShift).PriScaleNo <> 0, True, False)
                StationRecipe(iStation, iShift).UseAuxScale = IIf(StationSequence(iStation, iShift).AuxScaleNo <> 0, True, False)
                ' optional changes
                If (StationSequence(iStation, iShift).CourseData(iCourse).Cycles > 0) Then
                    StationRecipe(iStation, iShift).Cycles = StationSequence(iStation, iShift).CourseData(iCourse).Cycles
                    StationRecipe(iStation, iShift).CyclesSave = StationRecipe(iStation, iShift).Cycles
                End If
                If (StationSequence(iStation, iShift).CourseData(iCourse).LoadRate > 0) Then
                    StationRecipe(iStation, iShift).Load_Rate = StationSequence(iStation, iShift).CourseData(iCourse).LoadRate
                    StationRecipe(iStation, iShift).Load_RateSave = StationRecipe(iStation, iShift).Load_Rate
                End If
                If (StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate > 0) Then
                    StationRecipe(iStation, iShift).Purge_Flow = StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate
                End If
                ' save station recipes
                Save_StationRecipes
                ' Update Station Recipe descriptors
                UpdateStnRcpDsc iStation, iShift
                ' Update Station LiveFuel setup
                LiveFuel_Update iStation, iShift
                ' Job Description
                StationControl(iStation, iShift).Job_Description = StationSequence(iStation, iShift).Description
                ' estimated Job Duration
                StationControl(iStation, iShift).EstJobDur = StationSequence(iStation, iShift).EstSeqDuration
                StationControl(iStation, iShift).EstJobDurDesc = DurationDescription(StationControl(iStation, iShift).EstJobDur)
            End If
            ' initialize recipe control
            Recipe_Init iStation, iShift
            ' How does this recipe start?
            Select Case StationRecipe(iStation, iShift).StartMethod
                Case STARTNOW
                    ' No Delay
                    Recipe_Start iStation, iShift
                Case STARTDELAYED
                    ' Start after delay (in minutes)
                    StationControl(iStation, iShift).DelaySeconds = CDbl(60) * StationRecipe(iStation, iShift).StartDelay
                    StationControl(iStation, iShift).DelayToGo = CDbl(0)
                    StationControl(iStation, iShift).Mode = VBSTARTWAIT
                Case STARTATDATE
                    ' Start At Date
                    If StationRecipe(iStation, iShift).StartDate > Now() Then
                        ' calc seconds until start datetime
                        StationControl(iStation, iShift).DelaySeconds = CDbl(DateDiff("s", Now(), StationRecipe(iStation, iShift).StartDate))
                        StationControl(iStation, iShift).DelayToGo = CDbl(0)
                        StationControl(iStation, iShift).Mode = VBSTARTWAIT
                    Else
                        ' start test now
                        Recipe_Start iStation, iShift
                    End If
                Case Else
                    ' No Delay
                    Recipe_Start iStation, iShift
            End Select

        Case Else
            ' nothing to do; invalid Type; return station to idle
            StationSequence(iStation, iShift).Validated = False
    End Select
    
    ' Write Header data to data file
    Header_Write iStation, iShift

    ' update JobLog
    sMsg = "Course #" & StationControl(iStation, iShift).Course & " started."
    Write_JLog iStation, iShift, sMsg

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
  Case vbIgnore      ' close down
    Station_Clear iStation, iShift
    StationControl(iStation, iShift).Mode = VBIDLEWAITING
    Exit Sub
End Select
End Sub

Sub Course_Next(ByVal iStation As Integer, iShift As Integer)    ' Station iStation passed along
' Function Name:    Course_Next
' Author:           Brunrose
' Description:      checks for end of job sequence,
'                   otherwise proceeds to the next course
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1516
Dim iCourse As Integer
Dim Idx As Integer
Dim idx2 As Integer
Dim sPrint As String

    ' time stamp end of current course
    iCourse = StationControl(iStation, iShift).Course
    StationSequence(iStation, iShift).CourseData(iCourse).DtsEnd = Now()
        
    Select Case StationSequence(iStation, iShift).CourseData(iCourse).Type
        Case courseWait
            ' Wait for Operator OK
        Case coursePause
            ' Pause for x Minutes
        Case courseRecipe
            ' Run a Recipe
            StationControl(iStation, iShift).End_Time = Now
            StationControl(iStation, iShift).End_Timer = StationControl(iStation, iShift).TestTimer
            ' Update TomCanLoad, if required
            If USINGREMCANLOAD Then
                If (Len(StnRemoteTask(iStation, iShift).TaskID) > 2) Then
                    If (StnRemoteTask(iStation, iShift).TaskStatus = "InProcess") Then
                        If (JobInfo(iStation, iShift).End_OK) Then
                            ' Remote Task completed successfully
                            ' update Remote Task status in DB to Done
                            RemTask_Update iStation, iShift, "Done", "na"
                            ' update event log
                            sPrint = "Remote Task " & StnRemoteTask(iStation, iShift).TaskID
                            sPrint = sPrint & " (Job# " & StationControl(iStation, iShift).Job_Number
                            sPrint = sPrint & ") completed successfully."
                            Write_ELog sPrint
                            ' clear any duplicate RemTaskIDs & VINs on other station/shift's
                            For Idx = 1 To LAST_STN
                                For idx2 = 1 To NR_SHIFT
                                    If (StnRemoteTask(Idx, idx2).TaskID = StnRemoteTask(iStation, iShift).TaskID) Then
                                        RemData_Clear StnRemoteTask(Idx, idx2)
                                    End If
                                Next idx2
                            Next Idx
                            ' clear TOM Data for this Station/Shift
                            RemData_Clear StnRemoteTask(iStation, iShift)
                        Else
                            ' Remote Task did not complete successfully
                            ' reset Remote Task status in DB to Ready and add PreviousResult
                            Select Case USINGREMAVLFILES
                                Case True
                                    RemTask_Update iStation, iShift, "Failed", StnRemoteTask(iStation, iShift).PreviousResult
                                Case False
                                    RemTask_Update iStation, iShift, "Ready", StnRemoteTask(iStation, iShift).PreviousResult
                            End Select
                            ' update event log
                            sPrint = "Remote Task >" & StnRemoteTask(iStation, iShift).TaskID
                            sPrint = sPrint & "< (Job# " & StationControl(iStation, iShift).Job_Number
                            sPrint = sPrint & ") did not complete successfully."
                            Write_ELog sPrint
                        End If
                    Else
                        ' Remote Task did not complete successfully
                        ' reset Remote Task status in DB to Ready and add PreviousResult
                        Select Case USINGREMAVLFILES
                            Case True
                                RemTask_Update iStation, iShift, "Failed", StnRemoteTask(iStation, iShift).PreviousResult
                            Case False
                                RemTask_Update iStation, iShift, "Ready", StnRemoteTask(iStation, iShift).PreviousResult
                        End Select
                        ' update event log
                        sPrint = "Remote Task >" & StnRemoteTask(iStation, iShift).TaskID
                        sPrint = sPrint & "< (Job# " & StationControl(iStation, iShift).Job_Number
                        sPrint = sPrint & ") did not complete successfully."
                        Write_ELog sPrint
                    End If
                End If
            End If


        Case Else
            ' nothing to do; invalid Type
    End Select
    
    ' Update Header data in data file
    Header_Update iStation, iShift
    ' Write CycleWeights data in data file
    Weights_Write iStation, iShift
    
    ' sequence done yet?
    If StationSequence(iStation, iShift).NumCourses > StationControl(iStation, iShift).Course Then
        
        ' keep going; switch to next course
        iCourse = StationControl(iStation, iShift).Course + 1
        Course_Init iStation, iShift, iCourse
        
    Else
    
        ' done
        Station_Finish iStation, iShift
        
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
  Case vbIgnore      ' close down
    Station_Clear iStation, iShift
    StationControl(iStation, iShift).Mode = VBIDLEWAITING
    Exit Sub
End Select
End Sub

Sub Recipe_Init(ByVal iStation As Integer, iShift As Integer)
' Function Name:    Recipe_Init
' Description:      Initializes recipe variables.
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 15351
Dim sPrint As String
Dim Idx As Integer

    '  Restore the Recipe's Original Load Method, WC Multiplier & # of Cycles
    StationRecipe(iStation, iShift).Load_Method = StationRecipe(iStation, iShift).Load_MethodSave
    StationRecipe(iStation, iShift).NitrogenFlow = StationRecipe(iStation, iShift).NitrogenFlowSave
    StationRecipe(iStation, iShift).Cycles = StationRecipe(iStation, iShift).CyclesSave
    StationRecipe(iStation, iShift).WC_Mult = StationRecipe(iStation, iShift).WC_MultSave
        
    ' reset the bar graphs
    Reset_Bar_Graph iStation, iShift

    ' reset cycle counters
    StationControl(iStation, iShift).CurrCycle = CInt(1)
    StationControl(iStation, iShift).CompletedCycles = CInt(0)
    StationControl(iStation, iShift).CompletedLoads = CInt(0)
    StationControl(iStation, iShift).CompletedPurges = CInt(0)
    
    ' clear Canister LeakCheck Status
    StationControl(iStation, iShift).LeakCheckStatus = NORESULT
    
    ' Reset Timers
    Stn_Default_Log_TestTimer(iStation, iShift) = 0
    Stn_LT_Log_TestTimer(iStation, iShift) = 0
    Stn_Leak_Log_TestTimer(iStation, iShift) = 0
    Stn_Load_Log_TestTimer(iStation, iShift) = 0
    Stn_Purge_Log_TestTimer(iStation, iShift) = 0
    PreviousReportTimer(iStation, iShift) = 0
    PreviousTotalTimer(iStation, iShift) = 0
'    StationControl(iStation, iShift).TestTimer = 0
    StationControl(iStation, iShift).TestTimerIsRunning = True
    StationControl(iStation, iShift).Start_Time = Now()
    StationControl(iStation, iShift).End_Time = 0
    StationControl(iStation, iShift).End_Timer = 0
    
    ' Butane Density for active Butane Mfc
    If STN_INFO(iStation).Type = STN_ORVR2_TYPE And StationRecipe(iStation, iShift).UseHiRangeMFC Then
        StationControl(iStation, iShift).BtnDensity = GramsPerLiter * STN_INFO(iStation).ButMfc2DensityMult
    Else
        StationControl(iStation, iShift).BtnDensity = GramsPerLiter * STN_INFO(iStation).ButMfcDensityMult
    End If
    
    ' reset elapsed hours
    LoadControl(iStation, iShift).ElapsedHours = 0
    LoadControl(iStation, iShift).ElapsedHours_Prev = 0
    PurgeControl(iStation, iShift).ElapsedHours = 0
    PurgeControl(iStation, iShift).ElapsedHours_Prev = 0
    ' reset cumulative totals
    LoadControl(iStation, iShift).loadTotalGrams = 0
    LoadControl(iStation, iShift).LoadTotalLiters = 0
    ' reset wt rate of change
    LoadControl(iStation, iShift).TotalWtChg = 0
    LoadControl(iStation, iShift).TotalWtChgRate = 0
    PurgeControl(iStation, iShift).TotalWtChg = 0
    PurgeControl(iStation, iShift).TotalWtChgRate = 0
    InIdx(iStation, iShift) = 1
    ' reset weight change
    For Idx = 0 To MAX_CYCLES
        StationCycleWeightData(iStation, iShift, Idx).Cycle_StartWeight_Total = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Cycle_EndWeight_Total = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Load_StartWeight_Aux = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Load_EndWeight_Aux = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Load_StartWeight_Pri = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Load_EndWeight_Pri = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Purge_StartWeight_Aux = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Purge_EndWeight_Aux = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Purge_StartWeight_Pri = CSng(0)
        StationCycleWeightData(iStation, iShift, Idx).Purge_EndWeight_Pri = CSng(0)
    Next Idx
    
    ' Clear both Load and Purge Stats
    Clear_Stats iStation, iShift, 3
    
    ' Set barometer values
    JobInfo(iStation, iShift).Start_Baro = AmbBaro
    JobInfo(iStation, iShift).End_Baro = AmbBaro

    ' default result = Failed; set to True, i.e. Passed, when job completes successfully
    JobInfo(iStation, iShift).End_OK = False
    
    ' Update other Station Recipe descriptors
    UpdateStnRcpDsc iStation, iShift

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
  Case vbIgnore      ' close down
    Station_Clear iStation, iShift
    StationControl(iStation, iShift).Mode = VBIDLEWAITING
    Exit Sub
End Select
End Sub

Sub Station_Init(ByVal iStation As Integer, iShift As Integer)
' Function Name:    Station_Init
' Description:      Initializes station variables.
'                   Clears labels and creates database file.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 15
Dim strExists As String
Dim filename As String
Dim sPrint As String
Dim iSeq As JobSequence
Dim iCourse As Integer

    ' initialize station control block
'    StationControl(iStation, iShift).AbortRequest = False
    StationControl(iStation, iShift).Actual = 0
    StationControl(iStation, iShift).AlarmDelayTime = 0
    StationControl(iStation, iShift).AuxScaleStn = 0
    StationControl(iStation, iShift).AuxScaleWt = 0
    StationControl(iStation, iShift).AuxTare = 0
    StationControl(iStation, iShift).AuxWt_End = 0
    StationControl(iStation, iShift).AuxWt_Start = 0
    StationControl(iStation, iShift).BtnDensity = 0
'    StationControl(iStation, iShift).ContinueRequest = False
    StationControl(iStation, iShift).Course = 0
    StationControl(iStation, iShift).CurrCycle = 1
    StationControl(iStation, iShift).CompletedCycles = 0
    StationControl(iStation, iShift).DBFile = ""
    StationControl(iStation, iShift).DelaySeconds = 0
    StationControl(iStation, iShift).DelayToGo = 0
    StationControl(iStation, iShift).End_Time = Now
    StationControl(iStation, iShift).End_Timer = 0
    StationControl(iStation, iShift).EstJobDur = 0
    StationControl(iStation, iShift).EstJobDurDesc = "?"
    StationControl(iStation, iShift).IsPausedInAlarm = False
    StationControl(iStation, iShift).Job_Description = "?"
    StationControl(iStation, iShift).Job_Number = "000000"
    StationControl(iStation, iShift).Mode = VBIDLE
    StationControl(iStation, iShift).Mode_Last = VBIDLE
    StationControl(iStation, iShift).Mode_PauseSave = VBIDLE
    StationControl(iStation, iShift).Mode_StartDts = Now
    StationControl(iStation, iShift).ModeIsIdle_DebounceCount = 0
    StationControl(iStation, iShift).ModeIsIdle_Debounced = False
    StationControl(iStation, iShift).NewDataInDB = False
    StationControl(iStation, iShift).OotCurrent = ootNone
    StationControl(iStation, iShift).PausedDts = Now
    StationControl(iStation, iShift).OotResponse = ootrspUndefined
    StationControl(iStation, iShift).PauseAlarmStartTime = Now
    StationControl(iStation, iShift).PausedDts = Now
    ' Station Paused Message
    StationControl(iStation, iShift).PauseMessage = ""
    StationControl(iStation, iShift).PriScaleStn = 0
    StationControl(iStation, iShift).PriScaleWt = 0
    StationControl(iStation, iShift).PriTare = 0
    StationControl(iStation, iShift).PriWt_End = 0
    StationControl(iStation, iShift).PriWt_Start = 0
    StationControl(iStation, iShift).RptFile = ""
    StationControl(iStation, iShift).Scale_OK = False
    StationControl(iStation, iShift).ScalesInUse = False
    StationControl(iStation, iShift).Start_Time = Now
    StationControl(iStation, iShift).StartMethod = 0
'    StationControl(iStation, iShift).StartRequest = False
'    StationControl(iStation, iShift).StopRequest = False
    StationControl(iStation, iShift).Target = 0
    StationControl(iStation, iShift).TestTimer = 0
    StationControl(iStation, iShift).TestTimerIsRunning = True
    StationRemStatusControl(iStation, iShift).Mode_LastStatus = VBIDLE
    StationRemStatusControl(iStation, iShift).Phase_LastStatus = 0
    StationRemStatusControl(iStation, iShift).Cycle_LastStatus = 0
    
    SEQ_Step(iStation, iShift) = 0

    
    ' open job database file
goodentry:
    StationControl(iStation, iShift).Job_Number = Format(SysConfig.Next_File, "000000")
    
    ' Test to see if DB File already there.
    ' It should not be unless file numbers reused
    StationControl(iStation, iShift).DBFile = FILEPATH_data & "C" + StationControl(iStation, iShift).Job_Number + AccessDbFileExt
    strExists = Dir(StationControl(iStation, iShift).DBFile)
    If strExists <> "" Then
        Write_ELog ("Data Base in use: Trying Another C" & Format(SysConfig.Next_File, "000000"))
        Delay_Box "File " & strExists & " Data-Base Already Exists for Station " & iStation & " Shift " & iShift, MSGDELAY, msgSHOW
        ' increment next file number
        SysConfig.Next_File = SysConfig.Next_File + 1
        GoTo goodentry
    End If
    
    ' Create Database file and filename
    filename = "C" + StationControl(iStation, iShift).Job_Number + AccessDbFileExt
    StationControl(iStation, iShift).DBFile = FILEPATH_data & filename
    FileCopy FILEPATH_sysdbf & DATAMODEL, StationControl(iStation, iShift).DBFile
    ' verify copy
    strExists = Dir(StationControl(iStation, iShift).DBFile)
    If strExists <> filename Then
        Write_ELog ("Data Base File Create Failed (" & StationControl(iStation, iShift).DBFile & "): Trying Another C" & Format(SysConfig.Next_File, "000000"))
        Delay_Box "Data Base File Create Failed (" & StationControl(iStation, iShift).DBFile & ") for Station " & iStation & " Shift " & iShift, MSGDELAY, msgSHOW
        GoTo goodentry
    End If
    
    
    ' increment next file number
    SysConfig.Next_File = SysConfig.Next_File + 1
    ' save next file number in config file.
    Save_Config
    ' copy system configuration values
    StationConfig(iStation, iShift) = SysConfig
    ' Write configuration data to data file
    Config_Write iStation, iShift
    ' Write Sysdef data to data file
    Sysdef_Write iStation, iShift
    ' Write Job Sequence to data file
    Sequence_Write iStation, iShift
    
    
    ' Create Report Filename Kernel
    filename = ""
    '   1st Part
    Select Case StationConfig(iStation, iShift).ReportFileName1stPart
        Case RPT_JOBNUMBER
            ' Job #
            filename = "Job" & Format(StationControl(iStation, iShift).Job_Number, "000000") & "_"
        Case RPT_STARTDTS
            ' Start DateTime
            filename = Format(StationControl(iStation, iShift).Start_Time, "YYYYMMDD_HHMMSS") & "_"
        Case RPT_STNSHIFT
            ' Station Shift
            filename = "Station" & Format(iStation, "0") & "_Shift" & Format(iShift, "0") & "_"
        Case RPT_OPERENTER
            ' Operator Entry
            filename = frmStnDetail.txtRptName1 & "_"
        Case RPT_REMTASKID
            ' Remote Task Order ID
            filename = StnRemoteTask(iStation, iShift).TaskID & "_"
        Case Else
            ' Default = Job #
            filename = "Job" & Format(StationControl(iStation, iShift).Job_Number, "00000") & "_"
    End Select
    '   2nd Part
    Select Case StationConfig(iStation, iShift).ReportFileName2ndPart
        Case RPT_NOTHING
            ' nothing
        Case RPT_JOBNUMBER
            ' Job #
            filename = filename & "Job" & Format(StationControl(iStation, iShift).Job_Number, "000000") & "_"
        Case RPT_STARTDTS
            ' Start DateTime
            filename = filename & Format(StationControl(iStation, iShift).Start_Time, "YYYYMMDD_HHMMSS") & "_"
        Case RPT_STNSHIFT
            ' Station Shift
            filename = filename & "Station" & Format(iStation, "0") & "_Shift" & Format(iShift, "0") & "_"
        Case RPT_OPERENTER
            ' Operator Entry
            filename = filename & frmStnDetail.txtRptName2 & "_"
        Case RPT_REMTASKID
            ' Remote Task Order ID
            filename = filename & StnRemoteTask(iStation, iShift).TaskID & "_"
        Case Else
            ' Default = Job #
            filename = filename & "Job" & Format(StationControl(iStation, iShift).Job_Number, "00000") & "_"
    End Select
    '   3rd Part
    Select Case StationConfig(iStation, iShift).ReportFileName3rdPart
        Case RPT_NOTHING
            ' nothing
        Case RPT_JOBNUMBER
            ' Job #
            filename = filename & "Job" & Format(StationControl(iStation, iShift).Job_Number, "000000") & "_"
        Case RPT_STARTDTS
            ' Start DateTime
            filename = filename & Format(StationControl(iStation, iShift).Start_Time, "YYYYMMDD_HHMMSS") & "_"
        Case RPT_STNSHIFT
            ' Station Shift
            filename = filename & "Station" & Format(iStation, "0") & "_Shift" & Format(iShift, "0") & "_"
        Case RPT_OPERENTER
            ' Operator Entry
            filename = filename & frmStnDetail.txtRptName3 & "_"
        Case RPT_REMTASKID
            ' Remote Task Order ID
            filename = filename & StnRemoteTask(iStation, iShift).TaskID & "_"
        Case Else
            ' Default = Job #
            filename = filename & "Job" & Format(StationControl(iStation, iShift).Job_Number, "00000") & "_"
    End Select
    '   Set Report File Name Kernel
    StationControl(iStation, iShift).RptFile = Left(filename, 50)
    '   Reset Valid Report File Name Flag
    Stn_OperReportNameIsValid = False
    
    
    
    ' Make sure log screen isn't up;  DB errors if it is
    If frmDataLog.Visible Then
        frmDataLog.cmdReturn.Value = True
        DoEvents
    End If
    
    ' Notify the Review Screen that a New Job has Started
    If frmReview.Visible Then
        frmReview.JobStart iStation, iShift
        DoEvents
    End If
    
    ' OOT Response
    StationConfig(iStation, iShift).BtnFlowResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).BtnFlowResp, ootrspContinue)
    StationConfig(iStation, iShift).NitFlowResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).NitFlowResp, ootrspContinue)
    StationConfig(iStation, iShift).FuelTempResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).FuelTempResp, ootrspContinue)
    StationConfig(iStation, iShift).PurFlowResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).PurFlowResp, ootrspContinue)
    StationConfig(iStation, iShift).AirMoistResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).AirMoistResp, ootrspContinue)
    StationConfig(iStation, iShift).AirTempResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).AirTempResp, ootrspContinue)
    StationConfig(iStation, iShift).CanVentResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).CanVentResp, ootrspContinue)
    StationConfig(iStation, iShift).LoadRateResp = IIf(USINGOOTPAUSE, StationConfig(iStation, iShift).LoadRateResp, ootrspContinue)
    
    ' *****  ADDED  *************
    ' Log TOM Task Start
    ' ******************
    If (USINGTOMCANLOAD) Then
        If ( _
            (Len(StnTomTask(iStation, iShift).TaskID) > 2) _
                And _
            (StnTomTask(iStation, iShift).TaskStatus <> "Done") _
            ) _
        Then
            ' update TOM Task Status to Active(if required)
            TomTask_Update iStation, iShift, "Active", "na"
            JobInfo(iStation, iShift).Comment = JobInfo(iStation, iShift).Comment & vbCrLf & "  TOM Task " & StnTomTask(iStation, iShift).TaskID
            ' update event log
            sPrint = "TOM Task >" & StnTomTask(iStation, iShift).TaskID
            sPrint = sPrint & "< (Job# " & StationControl(iStation, iShift).Job_Number
            sPrint = sPrint & ") started."
            Write_ELog sPrint
        End If
    End If
' *****  ADDED  *************
   
    
    ' ******************
    ' Log Remote Task Start
    ' ******************
    If (USINGREMCANLOAD) Then
        If ( _
            (Len(StnRemoteTask(iStation, iShift).TaskID) > 2) _
                And _
            (StnRemoteTask(iStation, iShift).TaskStatus <> "Done") _
            ) _
        Then
            ' update Remote Task Status to Active(if required)
            RemTask_Update iStation, iShift, "InProcess", "na"
            ' update event log
            sPrint = "Remote Task >" & StnRemoteTask(iStation, iShift).TaskID
            sPrint = sPrint & "< (Job# " & StationControl(iStation, iShift).Job_Number
            sPrint = sPrint & ") started."
            Write_ELog sPrint
        End If
    End If

    If (STN_INFO(iStation).Type = STN_LEAKTEST_TYPE) Then
        ' LeakTest station
        ' Write Recipe to data file
        Recipe_Write StationControl(iStation, iShift).DBFile, STN_INFO(iStation).Type, CourseRecipes(iStation, iShift, 1), 1
    Else
        ' "normal" station
        ' build array of Recipes (by Course)
        iSeq = StationSequence(iStation, iShift)
        frmRecipe.Show
        frmRecipe.tmrUpdate.Enabled = True
        frmRecipe.ChgRecipeMode STATIONMODE
        For iCourse = 1 To iSeq.NumCourses
            Select Case iSeq.CourseData(iCourse).Type
                Case courseWait, coursePause
                    ' Wait for operator OK
                    '     OR
                    ' Pause for x minutes
                    CourseRecipes(iStation, iShift, iCourse) = EmptyRecipe
                Case courseRecipe
                    ' run Recipe x
                    ' which Recipe ??
                    Select Case iSeq.CourseData(iCourse).RecipeNumber
                        Case 0
                            ' run current station recipe with optional changes
                            frmRecipe.InitRecipe
                            With frmRecipe
                                ' optional changes to recipe
                                If (iSeq.CourseData(iCourse).Cycles > 0) Then .txtPFCycle.text = Format(iSeq.CourseData(iCourse).Cycles, "##0")
                                If (iSeq.CourseData(iCourse).LoadRate > 0) Then .txtLoadRate.text = Format(iSeq.CourseData(iCourse).LoadRate, "##0.0##")
                                If (iSeq.CourseData(iCourse).PurgeRate > 0) Then .txtPurgeFlow.text = Format(iSeq.CourseData(iCourse).PurgeRate, "##0.0##")
                            End With
                            If Not frmRecipe.OkToRunRecipeInStation Then
                                ' recipe failed validation
                                ' update event log
                                sPrint = "Station Recipe Validation Failed"
                                sPrint = sPrint & " on Course #" + Format(iCourse, "#0")
                                sPrint = sPrint & " for Station #" + Format(iStation, "0")
                                If (NR_SHIFT > 1) Then sPrint = sPrint & " Shift #" + Format(iShift, "0")
                                Write_ELog sPrint
                            End If
                        Case Else
                            ' run master recipe with changes, some optional
                            frmRecipe.LoadNewRcp iSeq.CourseData(iCourse).RecipeNumber
                            With frmRecipe
                                ' optional changes to recipe
                                If (iSeq.CourseData(iCourse).Cycles > 0) Then .txtPFCycle.text = Format(iSeq.CourseData(iCourse).Cycles, "##0")
                                If (iSeq.CourseData(iCourse).LoadRate > 0) Then .txtLoadRate.text = Format(iSeq.CourseData(iCourse).LoadRate, "##0.0##")
                                If (iSeq.CourseData(iCourse).PurgeRate > 0) Then .txtPurgeFlow.text = Format(iSeq.CourseData(iCourse).PurgeRate, "##0.0##")
                                ' job sequence changes to recipe
                                .chkPrimaryScale.Value = IIf((iSeq.PriScaleNo > 0), cYES, cNO)
                                .chkUseAuxScale = IIf((iSeq.AuxScaleNo > 0), cYES, cNO)
                                .txtPrimaryScaleNo.text = Format(iSeq.PriScaleNo, "#0")
                                .txtAuxScaleNo.text = Format(iSeq.AuxScaleNo, "#0")
                                .txtIDLoad.text = Format(iSeq.IDLoad, "#0.00")
                                .txtIDPurge.text = Format(iSeq.IDPurge, "#0.00")
                                .txtIDVent.text = Format(iSeq.IDVent, "#0.00")
                                .txtLoadL.text = Format(iSeq.LoadL, "##0.00")
                                .txtLoadV.text = Format(iSeq.LoadV, "##0.00")
                                .txtPurgeL.text = Format(iSeq.PurgeL, "##0.00")
                                .txtPurgeV.text = Format(iSeq.PurgeV, "##0.00")
                                .txtVentL.text = Format(iSeq.VentL, "##0.00")
                                .txtVentV.text = Format(iSeq.VentV, "##0.00")
                            End With
                            If Not frmRecipe.OkToRunRecipeInStation Then
                                ' recipe failed validation
                                ' update event log
                                sPrint = "Master Recipe #" + Format(iSeq.CourseData(iCourse).RecipeNumber, "##0") & " Validation Failed"
                                sPrint = sPrint & " on Course #" + Format(iCourse, "#0")
                                sPrint = sPrint & " for Station #" + Format(iStation, "0")
                                If (NR_SHIFT > 1) Then sPrint = sPrint & " Shift #" + Format(iShift, "0")
                                Write_ELog sPrint
                            End If
                    End Select
                    frmRecipe.ExportRecipe
                    CourseRecipes(iStation, iShift, iCourse) = ExportedRecipe
                Case Else
                    ' invalid course type
                    CourseRecipes(iStation, iShift, iCourse) = EmptyRecipe
                    ' update event log
                    sPrint = "Course Type of " + Format(iSeq.CourseData(iCourse).Type, "##0") & " is Invalid"
                    sPrint = sPrint & " on Course #" + Format(iCourse, "#0")
                    sPrint = sPrint & " for Station #" + Format(iStation, "0")
                    If (NR_SHIFT > 1) Then sPrint = sPrint & " Shift #" + Format(iShift, "0")
                    Write_ELog sPrint
            End Select
        
            ' Write Recipe to data file
            Recipe_Write StationControl(iStation, iShift).DBFile, STN_INFO(iStation).Type, CourseRecipes(iStation, iShift, iCourse), iCourse
        Next iCourse
        ' close recipe screen
        frmRecipe.ExitScreen
    End If
    
    
    ' update JobLog
    sPrint = "Job #" & StationControl(iStation, iShift).Job_Number & " started."
    Write_JLog iStation, iShift, sPrint
    
    ' using simulation ??
    If (USINGSIMULATION And (Not IoComOn)) Then
        sPrint = "USING SIMULATED I/O"
        Write_JLog iStation, iShift, sPrint
    End If
    If (USINGSIMULATION And (Not SclComOn)) Then
        sPrint = "USING SIMULATED SCALES"
        Write_JLog iStation, iShift, sPrint
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
  Case vbIgnore      ' close down
    Station_Clear iStation, iShift
    StationControl(iStation, iShift).Mode = VBIDLEWAITING
    Exit Sub
End Select
End Sub

Public Sub Purge_Check(iStation As Integer, iShift As Integer)

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1005
Dim Purge_Output As Single
Dim Purge_Rate As Single
Dim span As Single
Dim time_seconds As Long
Dim curStep As Integer
Dim sMsg As String
Dim sMsg1 As String
Dim sMsg2 As String

    ' Continue to Request Station Aspirator
    PRG_INFO(STN_INFO(iStation).AspiratorNum).RequestRun = True
    
    ' Continue to Request Aspirator for Aux Scale (if needed)
    If StationRecipe(iStation, iShift).PurgeAuxCan And (StationRecipe(iStation, iShift).AuxScaleNo > 0) Then
        ' Only allowed to purge Aux Scale with a Vacuum Purge
        If Not StationConfig(iStation, iShift).PosPressPurge Then
            If ((STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum < 1) Or (STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum > NR_PRGAIR)) Then
                If (Not scaleFlag(iStation)) Then
                    scaleFlag(iStation) = True
                    sMsg1 = "Station " & Format(iStation, "##0") & " Shift " & Format(iShift, "##0")
                    sMsg2 = "Aux Scale #" & Format(StationRecipe(iStation, iShift).AuxScaleNo, "##0") & " has no owner"
                    sMsg = sMsg1 & " - " & sMsg2
                    Write_ELog sMsg
                End If
            Else
                scaleFlag(iStation) = False
                PRG_INFO(STN_INFO(StationRecipe(iStation, iShift).AuxScaleNo).AspiratorNum).RequestRun = True
            End If
        End If
    End If
                                
            ' Put Purge MFC in operation
    '        PurgeFlow_Rate = CSng(STN_Purge_FlowSP(iStation, iShift))
    '        span = Stn_AIO(station, asPurgeAirFlowSP).EUMax - Stn_AIO(station, asPurgeAirFlowSP).EUMin
    '        Purge_Output = Stn_AIO(station, asPurgeAirFlowSP).EUMin +  (span * Cal_MfcOutput(PurgeFlow_Rate, iStation, MFCPURGEAIR, Stn_MfcCal(station, MFCPURGEAIR)))
    '        Stn_OutAnalog station, asPurgeAirFlowSP, 1, Purge_Output
    
    ' Determine MFC SetPoint based on Purge Method
    Select Case StationRecipe(iStation, iShift).Purge_Method
        Case PURGEBYTIME, PURGEAUXONLY
            ' CONTINUOUS PURGES
            Purge_Rate = CSng(StationRecipe(iStation, iShift).Purge_Flow)
        Case PURGEBYPROFILE
            ' PURGE BY PROFILE
            PurgeController iStation, iShift
            Purge_Rate = CSng(PurgeControl(iStation, iShift).CurMfcSp)
        Case PURGEBYLITERS
            ' PURGE BY LITERS
            PurgeController iStation, iShift
            Purge_Rate = CSng(PurgeControl(iStation, iShift).CurMfcSp)
        Case PURGEBYVOLUME
            ' PURGE BY VOLUME
            PurgeController iStation, iShift
            Purge_Rate = CSng(PurgeControl(iStation, iShift).CurMfcSp)
        Case PURGEBYWC
            ' PURGE BY WORKING CAPACITY
            PurgeController iStation, iShift
            Purge_Rate = CSng(PurgeControl(iStation, iShift).CurMfcSp)
        Case PURGETOTARGET
            ' PURGE TO TARGET WEIGHT
            PurgeController iStation, iShift
            Purge_Rate = CSng(PurgeControl(iStation, iShift).CurMfcSp)
        Case PURGETOUNDOLOAD
            ' PURGE TO UNDO LOAD
            PurgeController iStation, iShift
            Purge_Rate = CSng(PurgeControl(iStation, iShift).CurMfcSp)
    End Select
    
    ' Set MFC SetPoint when Aspirator is on
    If Prg_DIO(STN_INFO(iStation).AspiratorNum, ipPiabSol).Value Then
        If Not Stn_MfcSpIsSet(iStation) Then
    '        Purge_Rate = CSng(StationRecipe(iStation, iShift).Purge_Flow)
    '        Purge_Output =  (span * Cal_MfcOutput(Purge_Rate, iStation, MFCPURGEAIR, Stn_MfcCal(station, MFCPURGEAIR)))
    '        Stn_OutAnalog station, asPurgeAirFlowSP, Purge_Output, outNORMAL
    '        Stn_MfcSpIsSet(station) = True
            span = Stn_AIO(iStation, asPurgeAirFlowSP).EuMax - Stn_AIO(iStation, asPurgeAirFlowSP).EuMin
            Purge_Output = Stn_AIO(iStation, asPurgeAirFlowSP).EuMin + (span * Cal_MfcOutput(Purge_Rate, iStation, MFCPURGEAIR, Stn_MfcCal(iStation, MFCPURGEAIR)))
            Stn_OutAnalog iStation, asPurgeAirFlowSP, Purge_Output, outNORMAL
            Stn_MfcSpIsSet(iStation) = True
        End If
    Else
        Stn_OutAnalog iStation, asPurgeAirFlowSP, 0#, outZERO
        Stn_MfcSpIsSet(iStation) = False
    End If
    
    ' update end-of-test baro
    JobInfo(iStation, iShift).End_Baro = AmbBaro
    
    
    
    Select Case PurgeControl(iStation, iShift).Phase
        Case PurgeStarting
            ChgPhase PurgePurging, Now, iStation, iShift
        
        Case PurgePurging
            ' Update Actual Progress
            Select Case StationRecipe(iStation, iShift).Purge_Method
                Case PURGEBYTIME, PURGEAUXONLY
                    ChgErrModule 2, 10051
                    time_seconds = 0#
                    time_seconds = DateDiff("s", PreviousNow(iStation, iShift), Now()) ' Date diff in seconds
                    If time_seconds > 0# Then
                        PreviousNow(iStation, iShift) = Now()
                        StationControl(iStation, iShift).Actual = StationControl(iStation, iShift).Actual + (time_seconds / 60)
                    End If
                Case PURGEBYLITERS
                    ChgErrModule 2, 10152
'                    If Not USINGLINEVOLUME Then
                        StationControl(iStation, iShift).Actual = PurgeControl(iStation, iShift).Purge_Total
'                    Else
'                        Dim LineVolAdjustment As Single
'                        LineVolAdjustment = (StationRecipe(iStation, iShift).VentV + StationRecipe(iStation, iShift).PurgeV) / StationRecipe(iStation, iShift).Purge_Can_Vol
'                        StationControl(iStation, iShift).Actual = PurgeControl(iStation, iShift).Purge_Total - LineVolAdjustment
'                    End If
                Case PURGEBYVOLUME
                    ChgErrModule 2, 10052
                    If Not USINGLINEVOLUME Then
                        StationControl(iStation, iShift).Actual = PurgeControl(iStation, iShift).Purge_Total / StationCanister(iStation, iShift).WorkingVolume
                    Else
                        Dim LineVolAdjustment As Single
                        LineVolAdjustment = (StationRecipe(iStation, iShift).VentV + StationRecipe(iStation, iShift).PurgeV) / StationRecipe(iStation, iShift).Purge_Can_Vol
                        StationControl(iStation, iShift).Actual = PurgeControl(iStation, iShift).Purge_Total / (StationCanister(iStation, iShift).WorkingVolume + LineVolAdjustment)
                    End If
                Case PURGEBYPROFILE
                    ' PURGE BY PROFILE
                    ChgErrModule 2, 10053
                    curStep = PurgeControl(iStation, iShift).curStep
                    If (StationProfile(iStation, iShift).StepType(curStep) = STEPLAST) Then
                        StationControl(iStation, iShift).Actual = StationControl(iStation, iShift).Target
                    Else
                        StationControl(iStation, iShift).Actual = PurgeControl(iStation, iShift).CompletedStepMinutes + PurgeControl(iStation, iShift).StepElapsedMinutes
                    End If
                Case PURGEBYWC
                    ' PURGE BY WORKING CAPACITY
                    ChgErrModule 2, 10054
                    StationControl(iStation, iShift).Actual = CSng(-100) * (PurgeControl(iStation, iShift).PriWtChg / StationCanister(iStation, iShift).WorkingCapacity)
                    If PurgeControl(iStation, iShift).Purge_Volumes >= StationRecipe(iStation, iShift).Purge_MaxVolumes Then
                        ' exceeded max canister volumes ??
                        If Not StationControl(iStation, iShift).IsPausedInAlarm Then
                            ' exceeded limit AND not paused for anything
                            sMsg = "Purge #" & Format(StationControl(iStation, iShift).CurrCycle, "##0")
                            sMsg1 = sMsg & " for Station " & Format(iStation, "0") & "  Shift " & Format(iShift, "0") & " failed to reach Target. "
                            sMsg2 = sMsg & " failed to reach Target. "
                            sMsg = "Actual=" & Format(StationControl(iStation, iShift).Actual, "###0.0")
                            sMsg = sMsg & ";Target=" & Format(StationControl(iStation, iShift).Target, "###0")
                            sMsg = sMsg & "% of Can.WC"
                            Write_ELog sMsg1 & sMsg
                            OOT_Write iStation, iShift, sMsg2 & sMsg
                            ChgPhase PurgeComplete, Now, iStation, iShift
                        End If
                    End If
                Case PURGETOTARGET
                    ' PURGE TO TARGET WEIGHT
                    ChgErrModule 2, 10055
                    StationControl(iStation, iShift).Actual = CSng(-1) * PurgeControl(iStation, iShift).PriWtChg
                    If PurgeControl(iStation, iShift).Purge_Volumes >= StationRecipe(iStation, iShift).Purge_MaxVolumes Then
                        ' exceeded max canister volumes ??
                        If Not StationControl(iStation, iShift).IsPausedInAlarm Then
                            ' exceeded limit AND not paused for anything
                            sMsg = "Purge #" & Format(StationControl(iStation, iShift).CurrCycle, "##0")
                            sMsg1 = sMsg & " for Station " & Format(iStation, "0") & "  Shift " & Format(iShift, "0") & " failed to reach Target. "
                            sMsg2 = sMsg & " failed to reach Target. "
                            sMsg = "Actual=" & Format(StationControl(iStation, iShift).Actual, "###0.0")
                            sMsg = sMsg & ";Target=" & Format(StationControl(iStation, iShift).Target, "###0")
                            sMsg = sMsg & "grams"
                            Write_ELog sMsg1 & sMsg
                            OOT_Write iStation, iShift, sMsg2 & sMsg
                            ChgPhase PurgeComplete, Now, iStation, iShift
                        End If
                    End If
                Case PURGETOUNDOLOAD
                    ' PURGE TO UNDO LOAD
                    ChgErrModule 2, 10056
                    StationControl(iStation, iShift).Actual = CSng(-1) * PurgeControl(iStation, iShift).PriWtChg
                    If PurgeControl(iStation, iShift).Purge_Volumes >= StationRecipe(iStation, iShift).Purge_MaxVolumes Then
                        ' exceeded max canister volumes ??
                        If Not StationControl(iStation, iShift).IsPausedInAlarm Then
                            ' exceeded limit AND not paused for anything
                            sMsg = "Purge #" & Format(StationControl(iStation, iShift).CurrCycle, "##0")
                            sMsg1 = sMsg & " for Station " & Format(iStation, "0") & "  Shift " & Format(iShift, "0") & " failed to reach Target. "
                            sMsg2 = sMsg & " failed to reach Target. "
                            sMsg = "Actual=" & Format(StationControl(iStation, iShift).Actual, "###0.0")
                            sMsg = sMsg & ";Target=" & Format(StationControl(iStation, iShift).Target, "###0")
                            sMsg = sMsg & "grams"
                            Write_ELog sMsg1 & sMsg
                            OOT_Write iStation, iShift, sMsg2 & sMsg
                            ChgPhase PurgeComplete, Now, iStation, iShift
                        End If
                    End If
            End Select
            ' Are We Done Yet??
            ChgErrModule 2, 10059
            Select Case StationRecipe(iStation, iShift).Purge_Method
                Case PURGEBYTIME, PURGEAUXONLY, PURGEBYLITERS, PURGEBYVOLUME, PURGEBYPROFILE
                    If StationControl(iStation, iShift).Actual >= StationControl(iStation, iShift).Target Then
                        ' all done O.K.
                        If Not StationControl(iStation, iShift).IsPausedInAlarm Then
                            ' actual=target AND not paused for anything
                            PurgeControl(iStation, iShift).TotalWtChgAtEOP = PurgeControl(iStation, iShift).TotalWtChg
                            ChgPhase PurgeComplete, Now, iStation, iShift
                        End If
                    End If
                Case PURGEBYWC, PURGETOTARGET, PURGETOUNDOLOAD
                    If StationRecipe(iStation, iShift).Purge_TargetMode = TARGETPURGEPAUSE Then
                        ' PurgeController makes "Done??" decision; Purge can only end at the end of a Pause
                    Else
                        If StationRecipe(iStation, iShift).EndMethod = ENDWEIGHTCHG Then
                            ' calcwc(stable weight change) needs target AND low weight-rate-of-change(on 1st cycle only)
                            If (StationControl(iStation, iShift).Actual >= StationControl(iStation, iShift).Target) Then
                                If ((StationControl(iStation, iShift).CurrCycle > 1) Or (Abs(PurgeControl(iStation, iShift).TotalWtChgRate) < 1)) Then
                                    ' all done O.K.
                                    If Not StationControl(iStation, iShift).IsPausedInAlarm Then
                                        ' actual=target AND not paused for anything
                                        PurgeControl(iStation, iShift).TotalWtChgAtEOP = PurgeControl(iStation, iShift).TotalWtChg
                                        ChgPhase PurgeComplete, Now, iStation, iShift
                                    End If
                                End If
                            End If
                        Else
                            ' only need target
                            If (StationControl(iStation, iShift).Actual >= StationControl(iStation, iShift).Target) Then
                                ' all done O.K.
                                If Not StationControl(iStation, iShift).IsPausedInAlarm Then
                                    ' actual=target AND not paused for anything
                                    PurgeControl(iStation, iShift).TotalWtChgAtEOP = PurgeControl(iStation, iShift).TotalWtChg
                                    ChgPhase PurgeComplete, Now, iStation, iShift
                                End If
                            End If
                        End If
                    End If
            End Select
            
        Case PurgeComplete
            ' Shut Off IO
            '   Station Valves
            Close_Stn_Valves iStation, iShift
            '   Scale Valves
            If StationRecipe(iStation, iShift).UsePriScale And StationControl(iStation, iShift).PriScaleStn > 0 _
                    And StationControl(iStation, iShift).PriScaleStn < FIRST_REMOTESCALE Then
                Stn_OutDigital StationControl(iStation, iShift).PriScaleStn, isPriAuxVentSol, cOFF
            End If
            If StationRecipe(iStation, iShift).PurgeAuxCan And StationControl(iStation, iShift).AuxScaleStn > 0 Then
                Stn_OutDigital StationControl(iStation, iShift).AuxScaleStn, isAuxPurgeSol, cOFF
            End If
            ' purge is complete
            ChgPhase PurgeStopping, Now, iStation, iShift
        
        Case PurgeStopping
            ' begin settling time
            ChgPhase PurgePause, (Now + MinutesFromNow(StationConfig(iStation, iShift).PurgeSettleTime)), iStation, iShift
        
        Case PurgePause
            '
            ' after scale values settle, end this purge cycle
            '
            If (Now > PurgeControl(iStation, iShift).PhaseDts) Then
                Purge_Done iStation, iShift
            End If
    
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

Public Sub LeakCheck_Check(iStation As Integer, iShift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1006

    '******************************************
    ' CHECK LEAK MODE   ***** in leak mode ****
    '******************************************
    
    Select Case LeakCheckControl.Method
        Case LEAKCHECKPRI
            LeakCheck_CheckPri iStation, iShift
        Case LEAKCHECKAUX
            LeakCheck_CheckAux iStation, iShift
        Case Else
            ' Nothing to do
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

Public Sub LeakCheck_CheckPri(iStation As Integer, iShift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 10061
Dim outEU As Single

    '******************************************
    '      LEAKCHECK PRIMARY CANISTER
    '      LEAKCHECK PRIMARY CANISTER
    '      LEAKCHECK PRIMARY CANISTER
    '******************************************
    StationControl(iStation, iShift).Actual = PTinvalue
    
    Select Case LeakCheckControl.Phase
    
        Case LeakPurging
            ' Getting Started
            If (Now > StationControl(iStation, iShift).Mode_StartDts + TimeSerial(0, 0, 5)) Or (USINGLEAKCHECKEXHAUSTSOL And Now > StationControl(iStation, iShift).Mode_StartDts + TimeSerial(0, 0, 1)) Then
                If PTinvalue < (0.5 * StationConfig(iStation, iShift).LCSetPoint) Or USINGLEAKCHECKEXHAUSTSOL Then
                
                    ' Ready to Build Pressure
                    ChgPhase LeakPressurizing, (Now + TimeSerial(0, 0, StationConfig(iStation, iShift).LCMinDelay)), iStation, iShift
                    Leak_Write CInt(iStation), CInt(iShift), LCBEGINPHASE1, NORESULT
                        
                    ' Energize Leak Check Exhaust Valve
                    If USINGLEAKCHECKEXHAUSTSOL Then Com_OutDigital icLeakCheckExhaustSol, cON  ' Turn ON LeakCheck Exhaust Valve
                    
                    Stn_OutAnalog iStation, asPurgeAirFlowSP, 0, outZERO                        ' close Purge MFC
                    Stn_OutDigital iStation, isPurgeSol, cOFF                                   ' station purge flow valve Off
                    Stn_OutDigital iStation, isLeakCheckSol, cON                                ' Turn on LeakCheck
                       
                    Select Case STN_INFO(iStation).Type
                        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                            outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                            Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL          ' fullscale (or 5) slpm
                            Stn_OutDigital iStation, isNitrogenSol, cON                         ' Turn on nitro
                        
                        Case STN_ORVR2_TYPE
                            If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                ' use higher range MFC
                                outEU = IIf(Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asNitrogenORVRFlowSP, outEU, outNORMAL   ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isNitrogenOrvrSol, cON                  ' Turn on nitro
                           Else
                                ' use lower range MFC
                                outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL       ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isNitrogenSol, cON                      ' Turn on nitro
                            End If
                        
                        Case STN_LIVEFUEL_TYPE
                            outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                            Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                            Stn_OutDigital iStation, isLiveFuelSol, cON                          ' Turn on live fuel vapor
'                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                    ' Isolate station from LiveFuel Tank
                            
                        Case STN_LIVEREG_TYPE
                            If StationRecipe(iStation, iShift).LiveFuel Then
                                outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isLiveFuelSol, cON                          ' Turn on live fuel vapor
'                                Stn_OutDigital iStation, isLoadTypeSelectSol, cON                   ' Isolate station from LiveFuel Tank
                            Else
                                outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL          ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isNitrogenSol, cON                         ' Turn on nitro
                            End If
                            
                        Case STN_LIVEORVR2_TYPE
                            If StationRecipe(iStation, iShift).LiveFuel Then
                                If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporORVRFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporORVRFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asLiveFuelVaporORVRFlowSP, outEU, outNORMAL     ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isLiveFuelOrvrSol, cON                         ' Turn on live fuel vapor carrier
'                                    Stn_OutDigital iStation, isLoadTypeSelectSol, cON                      ' Isolate station from LiveFuel Tank
                                Else
                                    ' use lower range MFC
                                    outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isLiveFuelSol, cON                              ' Turn on live fuel vapor carrier
'                                    Stn_OutDigital iStation, isLoadTypeSelectSol, cON                      ' Isolate station from LiveFuel Tank
                                End If
                            Else
                                If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    outEU = IIf(Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asNitrogenORVRFlowSP, outEU, outNORMAL   ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isNitrogenOrvrSol, cON                  ' Turn on nitro
                                Else
                                    ' use lower range MFC
                                    outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL       ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isNitrogenSol, cON                      ' Turn on nitro
                                End If
                            End If
                            
                        Case STN_COMBO3_TYPE
                            ' future
                            
                        Case Else
                            ' Nothing to do; invalid station type
                    End Select
                    
                    ' Shift valves
                    Select Case iShift
                        Case 1
                            ' nothing to do
                        Case 2
                            Stn_OutDigital iStation, isLoadShift2Sol, cON
                            Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                            Stn_OutDigital iStation, isVentShift2Sol, cON
                        Case 3
                            Stn_OutDigital iStation, isLoadShift2Sol, cON
                            Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                            Stn_OutDigital iStation, isVentShift2Sol, cON
                            Stn_OutDigital iStation, isLoadShift3Sol, cON
                            Stn_OutDigital iStation, isPurgeShift3Sol, cOFF
                            Stn_OutDigital iStation, isVentShift3Sol, cON
                        Case 4
                            Stn_OutDigital iStation, isLoadShift2Sol, cON
                            Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                            Stn_OutDigital iStation, isVentShift2Sol, cON
                            Stn_OutDigital iStation, isLoadShift4Sol, cON
                            Stn_OutDigital iStation, isPurgeShift4Sol, cOFF
                            Stn_OutDigital iStation, isVentShift4Sol, cON
                    End Select
                    
                    
                End If
            ' Else
                ' Keep On Waiting
            End If
        
        
        Case LeakPressurizing
            ' Pressurizing
            If PTinvalue >= StationConfig(iStation, iShift).LCSetPoint Then               ' We are there
                ChgPhase LeakTesting, (Now + TimeSerial(0, 0, StationConfig(iStation, iShift).LCTime)), iStation, iShift
                Leak_Write CInt(iStation), CInt(iShift), LCBEGINPHASE2, NORESULT
                Select Case STN_INFO(iStation).Type
                    Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                        Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO             ' 0 slpm
                        Stn_OutDigital iStation, isNitrogenSol, cOFF                     ' Turn off nitro
                        Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
                    Case STN_ORVR2_TYPE
                        If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog iStation, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                            Stn_OutDigital iStation, isNitrogenOrvrSol, cOFF             ' Turn off nitro
                            Stn_OutDigital iStation, isLeakCheckSol, cON                 ' Keep On LeakCheck
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                            Stn_OutDigital iStation, isNitrogenSol, cOFF                 ' Turn off nitro
                            Stn_OutDigital iStation, isLeakCheckSol, cON                 ' Keep On LeakCheck
                        End If
                    Case STN_LIVEFUEL_TYPE
                        Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                        Stn_OutDigital iStation, isLiveFuelSol, cOFF                     ' Turn off live fuel vapor
                        Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
                        Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                    Case STN_LIVEREG_TYPE
                        If StationRecipe(iStation, iShift).LiveFuel Then
                            Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                            Stn_OutDigital iStation, isLiveFuelSol, cOFF                     ' Turn off live fuel vapor
                            Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                        Else
                            Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO             ' 0 slpm
                            Stn_OutDigital iStation, isNitrogenSol, cOFF                     ' Turn off nitro
                            Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
                        End If
                    Case STN_LIVEORVR2_TYPE
                        If StationRecipe(iStation, iShift).LiveFuel Then
                            If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                ' use higher range MFC
                                Stn_OutAnalog iStation, asLiveFuelVaporORVRFlowSP, 0, outZERO        ' 0 slpm
                                Stn_OutDigital iStation, isLiveFuelOrvrSol, cOFF                     ' Turn off live fuel vapor
                                Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
    '                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                            Else
                                ' use lower range MFC
                                Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                                Stn_OutDigital iStation, isLiveFuelSol, cOFF                     ' Turn off live fuel vapor
                                Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
    '                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                            End If
                        Else
                            If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                ' use higher range MFC
                                Stn_OutAnalog iStation, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                                Stn_OutDigital iStation, isNitrogenOrvrSol, cOFF             ' Turn off nitro
                                Stn_OutDigital iStation, isLeakCheckSol, cON                 ' Keep On LeakCheck
                            Else
                                ' use lower range MFC
                                Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                                Stn_OutDigital iStation, isNitrogenSol, cOFF                 ' Turn off nitro
                                Stn_OutDigital iStation, isLeakCheckSol, cON                 ' Keep On LeakCheck
                            End If
                        End If
                    Case STN_COMBO3_TYPE
                        ' future
                    Case Else
                        ' Nothing to do
                End Select
            ElseIf (Now > StationControl(iStation, iShift).Mode_StartDts + TimeSerial(0, 0, StationConfig(iStation, iShift).LCMinDelay)) Then
                LeakCheck_Error iStation, iShift, RESULTFAIL_PRESSURETIMEOUT              ' error past min delay and not entered phase 2
            End If
        
        
        Case LeakTesting
            ' Watching Pressure Decay
            If PTinvalue < (StationConfig(iStation, iShift).LCSetPoint - (StationConfig(iStation, iShift).PressureDecay / 100) * StationConfig(iStation, iShift).LCSetPoint) Then
                ' error - leaked too fast
                LeakCheck_Error iStation, iShift, RESULTFAIL_LEAKRATE
            ElseIf (Now > LeakCheckControl.PhaseDts) Then
                'Good run
                LeakCheck_Done iStation, iShift
            End If
            ' Otherwise, keep on watching
            
        Case LeakComplete
            ' What to do next
            LeakCheck_Continue iStation, iShift
        Case Else
            ' Nothing to do
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

Public Sub LeakCheck_CheckAux(iStation As Integer, iShift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 10062
Dim outEU As Single

    '******************************************
    '        LEAKCHECK AUX CANISTER
    '        LEAKCHECK AUX CANISTER
    '        LEAKCHECK AUX CANISTER
    '******************************************
    StationControl(iStation, iShift).Actual = PTinvalue
    
    Select Case LeakCheckControl.Phase
    
        Case LeakPurging
            ' Getting Started
            If (Now > LeakCheckControl.PhaseDts) Then
                If PTinvalue < (0.5 * StationConfig(iStation, iShift).LCSetPoint) Or USINGLEAKCHECKEXHAUSTSOL Then
                
                    ' Ready to Build Pressure
                    ChgPhase LeakPressurizing, (Now + TimeSerial(0, 0, StationConfig(iStation, iShift).LCMinDelay)), iStation, iShift
                    Leak_Write CInt(iStation), CInt(iShift), LCBEGINPHASE1, NORESULT
                        
                    ' Energize Leak Check Exhaust Valve
                    If USINGLEAKCHECKEXHAUSTSOL Then Com_OutDigital icLeakCheckExhaustSol, cON  ' Turn ON LeakCheck Exhaust Valve
                    
                    Stn_OutAnalog iStation, asPurgeAirFlowSP, 0, outZERO                        ' close Purge MFC
                    Stn_OutDigital iStation, isPurgeSol, cOFF                                   ' station purge flow valve Off
                    Stn_OutDigital iStation, isLeakCheckSol, cON                                ' Turn on LeakCheck
                    Stn_OutDigital iStation, isAuxLeakCheckSol, cON                             ' Turn on Aux LeakCheck
                    Stn_OutDigital iStation, isAuxCanVentSol, cON                               ' Turn on AuxCanVent
                    Stn_OutDigital iStation, isPriDirectionSol, cON                             ' Turn on PriDirection
                       
                    Select Case STN_INFO(iStation).Type
                        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                            outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                            Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL          ' fullscale (or 5) slpm
                            Stn_OutDigital iStation, isNitrogenSol, cON                         ' Turn on nitro
                        
                        Case STN_ORVR2_TYPE
                            If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                ' use higher range MFC
                                outEU = IIf(Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asNitrogenORVRFlowSP, outEU, outNORMAL   ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isNitrogenOrvrSol, cON                  ' Turn on nitro
                            Else
                                ' use lower range MFC
                                outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL       ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isNitrogenSol, cON                      ' Turn on nitro
                            End If
                        
                        Case STN_LIVEFUEL_TYPE
                            outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                            Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                            Stn_OutDigital iStation, isLiveFuelSol, cON                          ' Turn on live fuel vapor
                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                        ' Isolate station from LiveFuel Tank
                            
                        Case STN_LIVEREG_TYPE
                            If StationRecipe(iStation, iShift).LiveFuel Then
                                outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isLiveFuelSol, cON                          ' Turn on live fuel vapor
                                Stn_OutDigital iStation, isLoadTypeSelectSol, cON                        ' Isolate station from LiveFuel Tank
                            Else
                                outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                                Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL          ' fullscale (or 5) slpm
                                Stn_OutDigital iStation, isNitrogenSol, cON                         ' Turn on nitro
                            End If
                            
                        Case STN_LIVEORVR2_TYPE
                            If StationRecipe(iStation, iShift).LiveFuel Then
                                If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporORVRFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asLiveFuelVaporORVRFlowSP, outEU, outNORMAL  ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isLiveFuelOrvrSol, cON                 ' Turn on live fuel vapor carrier
'                                    Stn_OutDigital iStation, isLoadTypeSelectSol, cON                   ' Isolate station from LiveFuel Tank
                                Else
                                    ' use lower range MFC
                                    outEU = IIf(Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax < 5#, Stn_AIO(iStation, asLiveFuelVaporFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isLiveFuelSol, cON                          ' Turn on live fuel vapor carrier
'                                    Stn_OutDigital iStation, isLoadTypeSelectSol, cON                   ' Isolate station from LiveFuel Tank
                                End If
                            Else
                                If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    outEU = IIf(Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenORVRFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asNitrogenORVRFlowSP, outEU, outNORMAL      ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isNitrogenOrvrSol, cON                     ' Turn on nitro
                                Else
                                    ' use lower range MFC
                                    outEU = IIf(Stn_AIO(iStation, asNitrogenFlowSP).EuMax < 5#, Stn_AIO(iStation, asNitrogenFlowSP).EuMax, 5#)
                                    Stn_OutAnalog iStation, asNitrogenFlowSP, outEU, outNORMAL          ' fullscale (or 5) slpm
                                    Stn_OutDigital iStation, isNitrogenSol, cON                         ' Turn on nitro
                                End If
                            End If
                            
                        Case STN_COMBO3_TYPE
                            ' future
                            
                        Case Else
                            ' Nothing to do; invalid station type
                    End Select
                    
                    ' Shift valves
                    Select Case iShift
                        Case 1
                            ' nothing to do
                        Case 2
                            Stn_OutDigital iStation, isLoadShift2Sol, cON
                            Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                            Stn_OutDigital iStation, isVentShift2Sol, cON
                        Case 3
                            Stn_OutDigital iStation, isLoadShift2Sol, cON
                            Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                            Stn_OutDigital iStation, isVentShift2Sol, cON
                            Stn_OutDigital iStation, isLoadShift3Sol, cON
                            Stn_OutDigital iStation, isPurgeShift3Sol, cOFF
                            Stn_OutDigital iStation, isVentShift3Sol, cON
                        Case 4
                            Stn_OutDigital iStation, isLoadShift2Sol, cON
                            Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                            Stn_OutDigital iStation, isVentShift2Sol, cON
                            Stn_OutDigital iStation, isLoadShift4Sol, cON
                            Stn_OutDigital iStation, isPurgeShift4Sol, cOFF
                            Stn_OutDigital iStation, isVentShift4Sol, cON
                    End Select
                    
                End If
            ' Else
                ' Keep On Waiting
            End If
        
        
        Case LeakPressurizing
            ' Pressurizing
            If PTinvalue >= StationConfig(iStation, iShift).LCSetPoint Then               ' We are there
                ChgPhase LeakTesting, (Now + TimeSerial(0, 0, StationConfig(iStation, iShift).LCTime)), iStation, iShift
                Leak_Write CInt(iStation), CInt(iShift), LCBEGINPHASE2, NORESULT
                Stn_OutDigital iStation, isLeakCheckSol, cON                                ' Turn on LeakCheck
                Stn_OutDigital iStation, isAuxLeakCheckSol, cON                             ' Turn on Aux LeakCheck
                Stn_OutDigital iStation, isAuxCanVentSol, cON                               ' Turn on AuxCanVent
                Stn_OutDigital iStation, isPriDirectionSol, cON                             ' Turn on PriDirection
                Select Case STN_INFO(iStation).Type
                    Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                        Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO             ' 0 slpm
                        Stn_OutDigital iStation, isNitrogenSol, cOFF                     ' Turn off nitro
                    Case STN_ORVR2_TYPE
                        If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                            ' use higher range MFC
                            Stn_OutAnalog iStation, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                            Stn_OutDigital iStation, isNitrogenOrvrSol, cOFF             ' Turn off nitro
                        Else
                            ' use lower range MFC
                            Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                            Stn_OutDigital iStation, isNitrogenSol, cOFF                 ' Turn off nitro
                        End If
                    Case STN_LIVEFUEL_TYPE
                        Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                        Stn_OutDigital iStation, isLiveFuelSol, cOFF                     ' Turn off live fuel vapor
                        Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                    Case STN_LIVEREG_TYPE
                        If StationRecipe(iStation, iShift).LiveFuel Then
                            Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                            Stn_OutDigital iStation, isLiveFuelSol, cOFF                     ' Turn off live fuel vapor
                            Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                        Else
                            Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO             ' 0 slpm
                            Stn_OutDigital iStation, isNitrogenSol, cOFF                     ' Turn off nitro
                            Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
                        End If
                    Case STN_LIVEORVR2_TYPE
                        If StationRecipe(iStation, iShift).LiveFuel Then
                             If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                ' use higher range MFC
                                Stn_OutAnalog iStation, asLiveFuelVaporORVRFlowSP, 0, outZERO        ' 0 slpm
                                Stn_OutDigital iStation, isLiveFuelOrvrSol, cOFF                     ' Turn off live fuel vapor
                                Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
    '                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                            Else
                                ' use lower range MFC
                                Stn_OutAnalog iStation, asLiveFuelVaporFlowSP, 0, outZERO        ' 0 slpm
                                Stn_OutDigital iStation, isLiveFuelSol, cOFF                     ' Turn off live fuel vapor
                                Stn_OutDigital iStation, isLeakCheckSol, cON                     ' Keep On LeakCheck
    '                            Stn_OutDigital iStation, isLoadTypeSelectSol, cON                ' Maintain isolation from LiveFuel Tank
                            End If
                        Else
                            If StationRecipe(iStation, iShift).UseHiRangeMFC Then
                                ' use higher range MFC
                                Stn_OutAnalog iStation, asNitrogenORVRFlowSP, 0, outZERO     ' 0 slpm
                                Stn_OutDigital iStation, isNitrogenOrvrSol, cOFF             ' Turn off nitro
                            Else
                                ' use lower range MFC
                                Stn_OutAnalog iStation, asNitrogenFlowSP, 0, outZERO         ' 0 slpm
                                Stn_OutDigital iStation, isNitrogenSol, cOFF                 ' Turn off nitro
                            End If
                        End If
                    Case STN_COMBO3_TYPE
                        ' future
                    Case Else
                        ' Nothing to do
                End Select
            ElseIf (Now > LeakCheckControl.PhaseDts) Then
                LeakCheck_Error iStation, iShift, RESULTFAIL_PRESSURETIMEOUT              ' error past min delay and not entered phase 2
            End If
        
        
        Case LeakTesting
            ' Watching Pressure Decay
            If PTinvalue < (StationConfig(iStation, iShift).LCSetPoint - (StationConfig(iStation, iShift).PressureDecay / 100) * StationConfig(iStation, iShift).LCSetPoint) Then
                ' error - leaked too fast
                LeakCheck_Error iStation, iShift, RESULTFAIL_LEAKRATE
            ElseIf (Now > LeakCheckControl.PhaseDts) Then
                'Good run
                LeakCheck_Done iStation, iShift
            End If
            ' Otherwise, keep on watching
            
        Case LeakComplete
            ' What to do next
            LeakCheck_Continue iStation, iShift
        Case Else
            ' Nothing to do
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

Public Sub LeakTest_Check(ByVal iStation As Integer, ByVal iShift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 1066

    '*******************************************************
    '**** CHECK LEAKTEST MODE    ***** in leaktest mode ****
    '** note: UpdateLeakInputs is done in map-stn-analogs **
    '*******************************************************
    
    ' Continuously update baro.
    JobInfo(iStation, iShift).End_Baro = AmbBaro
    
    ' LeakTest Sequence Manager
    SEQ_Nmbr(iStation, iShift) = seqLeakTest
    Select Case SEQ_Step(iStation, iShift)
        Case 0
            ' idle; start sequence
            SEQ_Step(iStation, iShift) = 1
            ' Clear LeakData
            StnLT2Data(iStation, iShift, 0) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 1) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 2) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 3) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 4) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 5) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 6) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 7) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 8) = BlankLT2_Data
            StnLT2Data(iStation, iShift, 9) = BlankLT2_Data
    
        Case 1 To 8
            ' running
            CalcEffLeakDia iStation
        Case 9
            ' done
            CalcEffLeakDia iStation
        Case 10
            ' reset
            Station_Finish iStation, iShift
        Case Is > 10
            ' aborted
            Station_Finish iStation, iShift
    End Select
    
    ' LeakTest Flow Supervisory Controller
    Controller_PID (iStation + 20)
    
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

Public Function EstimatedLoadDuration(Rcp As Recipe, Can As CanisterRecipe) As Single
' Routine Name: EstimatedLoadDuration
' Created by:   Brunrose
' Function:
' This routine estimates the duration of a single load cycle in minutes of a recipe & a canister.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 9091
Dim sLoad As Single
Dim sCanWC As Single

    ' Canister Butane Working Capacity
    If (Can.WorkingCapacity = 0) Then
        ' estimated working capacity
        sCanWC = DefCanVol2CanWcMult * Can.WorkingVolume
    Else
        ' actual working capacity
        sCanWC = Can.WorkingCapacity
    End If
    

    ' LOAD
    sLoad = 1.35 / CSng(60)                                          ' misc delays (estimated)
    sLoad = sLoad + SysConfig.NitrogenPurgeTime / CSng(60)           ' N2 Push just before load
    sLoad = sLoad + (CSng(2) * (LoadMfcDelayTime / CSng(60)))        ' Valve/Mfc on/off delay
    Select Case Rcp.Load_Method
        Case NOLOAD
            sLoad = CSng(0)
        Case LOADBYTIME
            sLoad = sLoad + Rcp.Load_Time
            sLoad = sLoad + SysConfig.LoadSettleTime
        Case LOADBYWC
            sLoad = sLoad + (CSng(60) * CSng(Rcp.EPAFill))
            sLoad = sLoad + SysConfig.LoadSettleTime
        Case LOADBYWEIGHT
            If CSng(Rcp.Load_Rate) = CSng(0) Then
                sLoad = CSng(0)
            Else
                sLoad = sLoad + (CSng(60) * (CSng(Rcp.Load_Wt) / CSng(Rcp.Load_Rate)))
                sLoad = sLoad + SysConfig.LoadSettleTime
            End If
        Case LOADBYBREAKTHRU
            If CSng(Rcp.Load_Rate) = CSng(0) Then
                sLoad = CSng(0)
            Else
                sLoad = sLoad + (CSng(60) * (CSng(sCanWC) / CSng(Rcp.Load_Rate)))
                sLoad = sLoad + (CSng(1.5) * (CSng(60) * (CSng(Rcp.LoadBreakthrough) / CSng(Rcp.Load_Rate))))
                sLoad = sLoad + SysConfig.LoadSettleTime
            End If
        Case LOADBYFID
            sLoad = sLoad   ' not defined
            sLoad = sLoad + SysConfig.LoadSettleTime
    End Select
    
    ' TOTAL
    EstimatedLoadDuration = sLoad

ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function EstimatedRcpDuration(Rcp As Recipe, Can As CanisterRecipe, Prf As PurgeProfileType) As Single
' Routine Name: EstimatedRcpDuration
' Created by:   Brunrose
' Function:
' This routine estimates the duration in minutes of a recipe & a canister.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 9092
Dim iStation As Integer
Dim iHr As Integer
Dim iMin As Integer
Dim iSec As Integer
Dim iFactor As Integer
Dim sHr As Single
Dim sMin As Single
Dim sSec As Single
Dim sFactor As Single
Dim sLeak As Single
Dim sLoad As Single
Dim sPurge As Single
Dim sCycle As Single
Dim sNumCycles As Single
Dim sMult As Single
Dim sTime As Single
Dim sCanWC As Single

    ' Canister Butane Working Capacity
    If (Can.WorkingCapacity = 0) Then
        ' estimated working capacity (for CalcWC end method)
        sCanWC = DefCanVol2CanWcMult * Can.WorkingVolume
    Else
        ' actual working capacity
        sCanWC = Can.WorkingCapacity
    End If
    
    ' number of cycles
    Select Case Rcp.EndMethod
        Case ENDCYCLES
            sNumCycles = CSng(Rcp.Cycles)
        Case ENDWEIGHTCHG
            sNumCycles = CSng(Rcp.EndMinimumCycles + 1)
        Case Else
            sNumCycles = CSng(Rcp.Cycles)
    End Select
    
    ' LEAK
    sLeak = CSng(1.45) / CSng(60)                                          ' misc delays (estimated)
    Select Case Rcp.LeakCheck
        Case True
            sLeak = sLeak + ((CSng(SysConfig.LCMinDelay) / CSng(5)) / CSng(60))          ' assume pressurize in 1/5 of Timeout
            sLeak = sLeak + (CSng(SysConfig.LCTime) / CSng(60))
            If (Rcp.LeakPrimary And Rcp.LeakAux) Then sLeak = sLeak + sLeak
            If Rcp.PauseAfterLeak Then sLeak = sLeak + Rcp.PauseLeakTime
        Case False
            sLeak = CSng(0)
    End Select
    
    ' PURGE
    sPurge = CSng(2.65) / CSng(60)                                         ' misc delays (estimated)
    Select Case Rcp.Purge_Method
        Case NOPURGE
            sPurge = CSng(0)
        Case PURGEBYTIME
            sPurge = sPurge + CSng(Rcp.Purge_Time)
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGEAUXONLY
            sPurge = sPurge + CSng(Rcp.Purge_AuxTime)
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGEBYLITERS
            sPurge = sPurge + EstimatedPurgeDuration(Rcp, Can)
            If Rcp.Purge_TargetMode = TARGETPURGEPAUSE Then
                sMult = EstimatedPurgeDuration(Rcp, Can) / CSng(Rcp.Purge_TargetPurge)
                If ((sMult - CSng(CLng(sMult))) > CSng(0)) Then sMult = sMult + 1
                sPurge = sPurge + (CSng(CLng(sMult)) * CSng(Rcp.Purge_TargetPause))
            End If
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGEBYVOLUME
            sPurge = sPurge + EstimatedPurgeDuration(Rcp, Can)
            If Rcp.Purge_TargetMode = TARGETPURGEPAUSE Then
                sMult = EstimatedPurgeDuration(Rcp, Can) / CSng(Rcp.Purge_TargetPurge)
                If ((sMult - CSng(CLng(sMult))) > CSng(0)) Then sMult = sMult + 1
                sPurge = sPurge + (CSng(CLng(sMult)) * CSng(Rcp.Purge_TargetPause))
            End If
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGEBYPROFILE
            sPurge = sPurge + CSng(Prf.Duration)
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGEBYWC
            ' assume 300 Volumes to reduce Can Wt by WC
            sMult = CSng(300) * (CSng(0.01) * CSng(Rcp.Purge_TargetWC))
            sTime = (sMult * CSng(Can.WorkingVolume)) / CSng(Rcp.Purge_Flow)
            sPurge = sPurge + sTime
            If (Rcp.Purge_TargetMode = TARGETPURGEPAUSE) Then
                If (Rcp.Purge_TargetPurge <> 0) Then
                    sMult = sTime / CSng(Rcp.Purge_TargetPurge)
                    If ((sMult - CSng(CLng(sMult))) > CSng(0)) Then sMult = sMult + 1
                    sPurge = sPurge + (CSng(CLng(sMult)) * CSng(Rcp.Purge_TargetPause))
                End If
            End If
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGETOTARGET
            ' assume 300 Volumes to reduce Can Wt to Target
            sMult = CSng(300)
            sTime = (sMult * CSng(Can.WorkingVolume)) / CSng(Rcp.Purge_Flow)
            sPurge = sPurge + sTime
            If Rcp.Purge_TargetMode = TARGETPURGEPAUSE Then
                sMult = CSng(sTime) / CSng(Rcp.Purge_TargetPurge)
                If ((sMult - CSng(CLng(sMult))) > CSng(0)) Then sMult = sMult + 1
                sPurge = sPurge + (CSng(CLng(sMult)) * CSng(Rcp.Purge_TargetPause))
            End If
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
        Case PURGETOUNDOLOAD
            ' assume 300 Volumes to reduce Can Wt to UndoLoad
            sMult = CSng(300)
            sTime = (sMult * CSng(Can.WorkingVolume)) / CSng(Rcp.Purge_Flow)
            sPurge = sPurge + sTime
            If Rcp.Purge_TargetMode = TARGETPURGEPAUSE Then
                sMult = CSng(sTime) / CSng(Rcp.Purge_TargetPurge)
                If ((sMult - CSng(CLng(sMult))) > CSng(0)) Then sMult = sMult + 1
                sPurge = sPurge + (CSng(CLng(sMult)) * CSng(Rcp.Purge_TargetPause))
            End If
            If Rcp.PauseAfterPurge Then sPurge = sPurge + Rcp.PausePurgeTime
            sPurge = sPurge + SysConfig.PurgeSettleTime
    End Select
    
    ' LOAD
    sLoad = CSng(1.35) / CSng(60)                                           ' misc delays (estimated)
    sLoad = sLoad + SysConfig.NitrogenPurgeTime / CSng(60)                  ' N2 Push just before load
    sLoad = sLoad + (CSng(2) * (LoadMfcDelayTime / CSng(60)))               ' Valve/Mfc on/off delay
    Select Case Rcp.Load_Method
        Case NOLOAD
            sLoad = CSng(0)
        Case LOADBYTIME
            sLoad = sLoad + CSng(Rcp.Load_Time)
            sLoad = sLoad + SysConfig.LoadSettleTime
        Case LOADBYWC
            sLoad = sLoad + (CSng(60) * ((sCanWC * CSng(Rcp.WC_Mult)) / CSng(Rcp.Load_Rate)))
            sLoad = sLoad + SysConfig.LoadSettleTime
        Case LOADBYWEIGHT
            If CSng(Rcp.Load_Rate) = 0 Then
                sLoad = CSng(0)
            Else
                sLoad = sLoad + (CSng(60) * (CSng(Rcp.Load_Wt) / CSng(Rcp.Load_Rate)))
                sLoad = sLoad + SysConfig.LoadSettleTime
            End If
        Case LOADBYBREAKTHRU
            If CSng(Rcp.Load_Rate) = 0 Then
                sLoad = CSng(0)
            Else
                sLoad = sLoad + (CSng(60) * (CSng(sCanWC) / CSng(Rcp.Load_Rate)))
                sLoad = sLoad + (CSng(1.5) * (CSng(60) * (CSng(Rcp.LoadBreakthrough) / CSng(Rcp.Load_Rate))))
                sLoad = sLoad + SysConfig.LoadSettleTime
            End If
        Case LOADBYFID
            sLoad = sLoad   ' not defined
            sLoad = sLoad + SysConfig.LoadSettleTime
    End Select
    If Rcp.PauseAfterLoad Then sLoad = sLoad + Rcp.PauseLoadTime
    
    ' Interference Facter
    sFactor = CSng(0)
    For iStation = 1 To LAST_STN
        If (Stn_ActiveShift(iStation) = 0) Then Stn_ActiveShift(iStation) = 1
        If (StationControl(iStation, Stn_ActiveShift(iStation)).Mode <> VBIDLE) Then sFactor = sFactor + (CSng(1) - (sFactor / CSng(6)))
    Next iStation
    sFactor = (sFactor / CSng(2)) * sNumCycles
        
    ' TOTAL
    EstimatedRcpDuration = sFactor + sLeak + (sNumCycles * (sPurge + sLoad))

ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function DurationDescription(ByVal minutesDur As Single) As String
' Routine Name: DurationDescription
' Created by:   Brunrose
' Function:
' This routine converts a Duration (in minutes) into a string description.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 9093
Dim Hrs, Days As Integer
Dim remainDur As Single
Dim sDesc As String

    Hrs = 0
    Days = 0
    remainDur = minutesDur
    Do While remainDur >= 60
        remainDur = remainDur - 60
        Hrs = Hrs + 1
    Loop
    Do While Hrs >= 24
        Hrs = Hrs - 24
        Days = Days + 1
    Loop
    
    sDesc = ""
    If Days > 1 Then
        sDesc = sDesc & Format(Days, "###0") & " days   "
    ElseIf Days > 0 Then
        sDesc = sDesc & Format(Days, "###0") & " day   "
    End If
    If Hrs > 1 Then
        ' duration >= 2 hours
        sDesc = sDesc & Format(Hrs, "#0") & " hrs   "
        sDesc = sDesc & Format(remainDur, "#0") & " min"
    ElseIf Hrs > 0 Then
        ' duration >= 1 hour
        sDesc = sDesc & Format(Hrs, "#0") & " hr   "
        sDesc = sDesc & Format(remainDur, "#0") & " min"
    ElseIf minutesDur > 1.5 Then
        ' duration > 90 seconds
        sDesc = sDesc & Format(Int(remainDur), "#0") & " min "
        sDesc = sDesc & Format((60 * (remainDur - Int(remainDur))), "#0") & " sec"
    Else
        ' duration <= 90 seconds
        sDesc = sDesc & Format((60 * remainDur), "##0") & " sec"
    End If

    ' Result
    DurationDescription = sDesc

ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function RecipeIsDone(ByVal station As Integer, ByVal Shift As Integer) As Boolean
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 4099
Dim flag As Boolean
Dim iCyc As Integer
Dim Idx As Integer
Dim AvgChg As Single
Dim CurChg As Single
Dim SumChg As Single
Dim ActTol As Single
Dim calcBWC As Single
Dim sMsg As String

    flag = False
    ' is recipe complete ??
    Select Case StationRecipe(station, Shift).EndMethod
        Case ENDCYCLES
'            If (StationControl(station, Shift).CompletedLoads >= StationRecipe(station, Shift).Cycles) Then flag = True
            If (StationControl(station, Shift).CompletedCycles >= StationRecipe(station, Shift).Cycles) Then flag = True
        Case ENDWEIGHTCHG
            If (StationControl(station, Shift).CompletedCycles >= StationRecipe(station, Shift).EndMinimumCycles) Then
                iCyc = StationControl(station, Shift).CompletedCycles
'                CurChg = StationCycleWeightData(station, Shift,iCyc).Load_TotalGrams
                CurChg = StationCycleWeightData(station, Shift, iCyc).Load_EndWeight_Pri - StationCycleWeightData(station, Shift, iCyc).Load_StartWeight_Pri
                SumChg = CSng(0)
                For Idx = 1 To StationRecipe(station, Shift).EndConsecutiveCycles
                    iCyc = StationControl(station, Shift).CompletedCycles - Idx
'                    SumChg = SumChg + StationCycleWeightData(station, Shift,iCyc).Load_TotalGrams
                    SumChg = SumChg + (StationCycleWeightData(station, Shift, iCyc).Load_EndWeight_Pri - StationCycleWeightData(station, Shift, iCyc).Load_StartWeight_Pri)
                Next Idx
                AvgChg = SumChg / StationRecipe(station, Shift).EndConsecutiveCycles
                ActTol = CSng(100) * Abs((CurChg - AvgChg) / CurChg)
                If (ActTol <= Abs(StationRecipe(DispStn, DispShift).EndWeightTolerance)) Then
                    ' Recipe is Complete
                    flag = True
                    ' Update Canister WC ??
                    If StationRecipe(station, Shift).UpdateCanWc Then
                        ' set Canister Working Capacity
                        calcBWC = ((StationRecipe(station, Shift).EndConsecutiveCycles * Abs(AvgChg)) + Abs(AvgChg)) / (StationRecipe(station, Shift).EndConsecutiveCycles + 1)
                        StationCanister(station, Shift).WorkingCapacity = calcBWC
                        ' save the canister values
                        Save_StationCanisters
                        ' log the results
                        sMsg = "Canister Working Capacity set to " & Format(calcBWC, "##,###,##0.0##") & " grams"
                        ' job log
                        Write_JLog station, Shift, sMsg
                        ' system event log
                        sMsg = "Station #" & Format(station, "0") & " Shift #" & Format(Shift, "0") & " " & sMsg
                        Write_ELog sMsg
                    End If
                End If
            End If
        Case Else
            flag = True
    End Select
    RecipeIsDone = flag
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function AnyMoreLiveFuelLoads(ByVal station As Integer, ByVal Shift As Integer) As Boolean
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 2, 4199
Dim flag As Boolean
Dim CoursesToGo As Integer
Dim iCourse As Integer
Dim iCycle As Integer

    flag = False
    
    If StationControl(station, Shift).Course > 0 Then
        CoursesToGo = StationSequence(station, Shift).NumCourses - StationControl(station, Shift).Course
        
        Select Case CoursesToGo
            Case Is > 0
                ' More Courses
                ' check current course
                iCourse = StationControl(station, Shift).Course
                For iCycle = (StationControl(station, Shift).CurrCycle) To CourseRecipes(station, Shift, iCourse).CyclesSave
                    ' does CourseRecipe do a Load??
                    If CourseRecipes(station, Shift, iCourse).Load_Method <> NOLOAD Then
                        ' does Load Use LiveFuel ??
                        If CourseRecipes(station, Shift, iCourse).LiveFuel Then flag = True
                    End If
                Next iCycle
                ' check all remaining courses
                For iCourse = (StationControl(station, Shift).Course + 1) To StationSequence(station, Shift).NumCourses
                    ' does CourseRecipe do a Load??
                    If CourseRecipes(station, Shift, iCourse).Load_Method <> NOLOAD Then
                        ' does Load Use LiveFuel ??
                        If CourseRecipes(station, Shift, iCourse).LiveFuel Then flag = True
                    End If
                Next iCourse
            Case Is = 0
                ' Last Course
                If (Not RecipeIsDone(station, Shift)) Then
                    ' check current course
                    iCourse = StationControl(station, Shift).Course
                    For iCycle = (StationControl(station, Shift).CurrCycle) To CourseRecipes(station, Shift, iCourse).CyclesSave
                        ' does CourseRecipe do a Load??
                        If CourseRecipes(station, Shift, iCourse).Load_Method <> NOLOAD Then
                            ' Any More Loads for current Recipe
                            If (StationControl(station, Shift).CompletedLoads < StationRecipe(station, Shift).Cycles) Then
                                ' does Load Use LiveFuel ??
                                If CourseRecipes(station, Shift, iCourse).LiveFuel Then flag = True
                            End If
                        End If
                    Next iCycle
                End If
                Case Else
                ' invalid case
        End Select
        
    End If
    AnyMoreLiveFuelLoads = flag
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function AllStationsIdle() As Boolean
Dim iStn As Integer
Dim idleCount As Integer

    idleCount = 0
    ' check for idle stations
    For iStn = 1 To LAST_STN
        If StationControl(iStn, Stn_ActiveShift(iStn)).Mode = VBIDLE _
            Or StationControl(iStn, Stn_ActiveShift(iStn)).Mode = VBIDLEWAITING Then
            idleCount = idleCount + 1
        End If
    Next iStn
    
    ' return result
    AllStationsIdle = IIf((idleCount = LAST_STN), True, False)

End Function

Sub Station_ContinuePB(station As Integer, Shift As Integer)
'
'  Only get here from some station in Operator Pause
'
'
'******************************************************************************


    If (Pause_Alarm = SYSTEMPAUSED) Then Exit Sub                ' There is a system wide pause


    ' is a sequence running ?
    Select Case SEQ_Nmbr(station, Shift)
        Case seqCanVentN2Feed
            ' ************************
            ' CanVent N2 Feed Sequence
            ' ************************
            Select Case SEQ_Step(station, Shift)
                Case 9
                    ' Successful Completion; Reset Sequence Number
                    SEQ_Nmbr(station, Shift) = seqIdle
                    SEQ_Step(station, Shift) = 0
                    ' Start Post Purge Delay
                    StationControl(station, Shift).End_Time = Now() + TimeSerial(0, StationRecipe(station, Shift).PausePurgeTime, 0)
                Case 90, 91, 95
                    ' Aborted; Restart Sequence
                    SEQ_Step(station, Shift) = 1
            End Select
        Case Else
            ' ************************
            '   no sequence running
            ' ************************
            ' nothing to do
    End Select
    
    If StationControl(station, Shift).Mode_PauseSave = VBPURGE Then
        If StationRecipe(station, Shift).UseAuxScale Then
            Com_OutDigital (icScale01AuxAirSol + StationRecipe(station, Shift).AuxScaleNo - 1), cON
        End If
    End If
    ' adjust station time value if required
    If StationControl(station, Shift).Mode_PauseSave = VBLOAD Then
        If StationRecipe(station, Shift).Load_Method = LOADBYWC Then
            LoadControl(station, Shift).WC_Load_Time = _
                LoadControl(station, Shift).WC_Load_Time + (Now() - StationControl(station, Shift).PausedDts)
        End If
        If StationRecipe(station, Shift).Load_Method = LOADBYTIME Then
            StationControl(station, Shift).Mode_StartDts = _
                StationControl(station, Shift).Mode_StartDts + (Now() - StationControl(station, Shift).PausedDts)
        End If
    End If
    
    ' Station Paused Message
    StationControl(station, Shift).PauseMessage = ""
    StationControl(station, Shift).PausedDts = 0
    
    Write_ELog "Operator Continued Station #" & Format(station, "0") & " Shift #" & Format(Shift, "0")
    Write_JLog station, Shift, "Operator Continued Station"
    
    If StationControl(station, Shift).Mode_PauseSave = VBLOAD Then
        Load_Continue station, Shift                    ' Continue(Resume) Load
    ElseIf StationControl(station, Shift).Mode_PauseSave = VBPURGE Then
        Purge_ContinueDelayed station, Shift            ' Continue(Resume) Purge
    ElseIf StationControl(station, Shift).Mode_PauseSave = VBPURGECONT Then
        Purge_ContinueDelayed station, Shift            ' Continue(Resume) Purge
    ElseIf StationControl(station, Shift).Mode_PauseSave = VBPRELOAD Then
        PreLoad_Start station, Shift                    ' Restart PreLoad
    ElseIf StationControl(station, Shift).Mode_PauseSave = VBLEAK Then
        LeakCheck_Start station, Shift                  ' Restart LeakCheck
    Else
        StationControl(station, Shift).Mode = StationControl(station, Shift).Mode_PauseSave
    End If
End Sub

Private Function EstimatedPurgeDuration(Rcp As Recipe, Can As CanisterRecipe) As Single
Dim tmpTime As Single
Dim canVol As Single
Dim rcpFlow As Single
Dim rcpVols As Single
Dim volPurge As Single
Dim volVent As Single

        If Can.Validated Then
            canVol = Can.WorkingVolume
            rcpFlow = Rcp.Purge_Flow
            rcpVols = Rcp.Purge_Can_Vol
            If (rcpFlow > 0) Then
                If USINGLINEVOLUME Then
                    ' Using Line Volume
                    volPurge = Rcp.PurgeV
                    volVent = Rcp.VentV
                    tmpTime = (((canVol * rcpVols) + volVent + volPurge) / rcpFlow)
                Else
                    ' not using line volume
                    tmpTime = ((canVol * rcpVols) / rcpFlow)
                End If
            Else
                ' Purge Flow = 0
                tmpTime = 0
            End If
        Else
            ' Invalid Canister Values
            tmpTime = 0
        End If
        
        EstimatedPurgeDuration = tmpTime

End Function


