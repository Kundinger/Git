Attribute VB_Name = "Module13"
' error module 13 ''''''''''''''''program ALARMS.bas '''''''''''''''''''''
Option Explicit
'
Private Const maxSystemVacSwCount = 2
Private Const tankVaporGen = 1
Private Const tankFuelStorage = 2
Private Const tankLevelSteady = 0
Private Const tankLevelRising = 1
Private Const tankLevelFalling = 2
Private LastTankLevel(2) As Single
Private MaxTankLevel(2) As Single
Private MinTankLevel(2) As Single
Private TankLevelChange(2) As Single
Private TankLevelTrend(2) As Integer
'
'
Sub OOT_Check()
'
' Function Name:    OOT_Check
' Author:           Analytical Process Programmer         8/8/96
' Description:      This routine checks flow and environment variables
'                   for out of tolerance condition.  IF out, sets correct
'                   out of tolerance status flag and updates out of
'                   tolerance report.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 3
Dim CheckInterval As Integer
Dim Index As Integer
Dim index2 As Integer
Dim bAboveLoLim As Boolean
Dim bOutHigh, bOutLow, bOutZero As Boolean
Dim deadweakMult As Single
Dim deltaHours As Single
Dim delta As Single
Dim nitSP As Single
Dim nitPV As Single
Dim ootcomment As String
Dim LoLim_pur_flow1 As Single   ' Low Limit of Purge Flow Tolerance Checking in slpm
Dim LoLim_btn_flow1 As Single   ' Low Limit of Butane Flow Tolerance Checking in slpm
Dim LoLim_nit_flow1 As Single   ' Low Limit of Nitrogen Flow Tolerance Checking in slpm
Dim Tol_pur_flow1 As Single     ' Purge Flow Tolerance in slpm
Dim Tol_btn_flow1 As Single     ' Butane Flow Tolerance in slpm
Dim Tol_nit_flow1 As Single     ' Nitrogen Flow Tolerance in slpm
Dim LoLim_Tank_Level As Single  ' Low Limit of Tank Level  in %FS(for Tolerance Checking)
Dim Tol_Tank_Level As Single    ' Tank Level Tolerance in %FS
Dim sDwell As Single
Dim valPV As Single             ' temp value
Dim valSP As Single             ' temp value
Dim valTol As Single            ' temp value
Dim sGramsPerLiter As Single    ' Specific MFC's Butane Density in GramsPerLiter


    ' Don't Check if System Paused
    If (Pause_Alarm = SYSTEMPAUSED) Then Exit Sub
        
    ' USE NEXT LINE TO DISABLE PURGE AIR TEMP, HUMIDITY TILL CONNECTED
    Const CHECKPAVALS = True
        
    ' ******************************************************************
        
    ' Check every 2 seconds
    CheckInterval = 2
    If DateDiff("s", LastOOTCheckTime, Now) >= CheckInterval Then
        LastOOTCheckTime = Now
                    
        ChgErrModule 13, 3131
    
        ' Who is OOT and Who is Not
        For Index = 1 To LAST_STN
            For index2 = 1 To NR_SHIFT
                        
                ' Only do OOT checks if station is not PAUSED-IN-ALARM
                If Not StationControl(Index, index2).IsPausedInAlarm Then
                    ' Only do checks if not already in PAUSED-DUE-TO-OOT
                    If StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                    
                        ' CHECK LIVEFUEL LEVEL(S)
                        If (((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE)) And (STN_INFO(Index).ADF_DEF.hasADF_LT)) Then
                            If USINGFUELLEVELOOT Then
                                ' tank level checking (i.e. tank leak checking) is enabled
                                If ((StationControl(Index, index2).Mode <> VBLOAD) And Not (Stn_DIO(Index, isFuelPumpMotor).Value)) Then
                                    ' Vapor Generator Tank
                                    If (Stn_AIO(Index, asFuelTankLevel).EUValue > MaxTankLevel(tankVaporGen)) Then
                                        MaxTankLevel(tankVaporGen) = Stn_AIO(Index, asFuelTankLevel).EUValue
                                    ElseIf (Stn_AIO(Index, asFuelTankLevel).EUValue < MinTankLevel(tankVaporGen)) Then
                                        MinTankLevel(tankVaporGen) = Stn_AIO(Index, asFuelTankLevel).EUValue
                                    End If
                                    LastTankLevel(tankVaporGen) = Stn_AIO(Index, asFuelTankLevel).EUValue
                                    ' Check Vapor Generator Tank Level
                                    Tol_Tank_Level = StationCfg_ADF(Index, index2).VaporGenLevelTol
                                    bOutHigh = IIf(Stn_AIO(Index, asFuelTankLevel).EUValue > AdfControl(Index).LevelSP + Tol_Tank_Level, True, False)
                                    bOutLow = IIf(Stn_AIO(Index, asFuelTankLevel).EUValue < AdfControl(Index).LevelSP - Tol_Tank_Level, True, False)
                                    ' Vapor Generator Tank Level OOTcnt
                                    If bOutHigh _
                                             Or _
                                            bOutLow Then
                                        If OOTs(Index, index2).FuelLevelOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).FuelLevelOOTCnt = OOTs(Index, index2).FuelLevelOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).FuelLevelOOTCnt = 0
                                    End If
                                    
                                    
                                    ' Fuel Storage Tank
                                    If STN_INFO(Index).ADF_DEF.hasADF_FST Then
                                        If (Stn_AIO(Index, asStorageTankLevel).EUValue > MaxTankLevel(tankFuelStorage)) Then
                                            MaxTankLevel(tankFuelStorage) = Stn_AIO(Index, asStorageTankLevel).EUValue
                                        ElseIf (Stn_AIO(Index, asStorageTankLevel).EUValue < MinTankLevel(tankFuelStorage)) Then
                                            MinTankLevel(tankFuelStorage) = Stn_AIO(Index, asStorageTankLevel).EUValue
                                        End If
                                        LastTankLevel(tankFuelStorage) = Stn_AIO(Index, asStorageTankLevel).EUValue
                                        ' Check Fuel Storage Tank Level
                                        Tol_Tank_Level = StationCfg_ADF(Index, index2).FuelStorageLevelTol
                                        ' check Fuel Storage Tank for leak
                                        bOutHigh = IIf(Stn_AIO(Index, asStorageTankLevel).EUValue > FstControl(Index).LevelSP + Tol_Tank_Level, True, False)
                                        bOutLow = IIf(Stn_AIO(Index, asStorageTankLevel).EUValue < FstControl(Index).LevelSP - Tol_Tank_Level, True, False)
                                        ' Fuel Storage Tank Level OOTcnt
                                        If bOutHigh _
                                                 Or _
                                                bOutLow Then
                                            If OOTs(Index, index2).StorageLevelOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                                OOTs(Index, index2).StorageLevelOOTCnt = OOTs(Index, index2).StorageLevelOOTCnt + CheckInterval
                                            End If
                                        Else
                                            OOTs(Index, index2).StorageLevelOOTCnt = 0
                                        End If
                                    Else
                                        OOTs(Index, index2).StorageLevelOOTCnt = 0
                                    End If
                                Else
                                    ' Set counts to zero
                                    OOTs(Index, index2).FuelLevelOOTCnt = 0
                                    OOTs(Index, index2).StorageLevelOOTCnt = 0
                                End If
                            Else
                                ' Set counts to zero
                                OOTs(Index, index2).FuelLevelOOTCnt = 0
                                OOTs(Index, index2).StorageLevelOOTCnt = 0
                            End If
                        Else
                            ' Set counts to zero
                            OOTs(Index, index2).FuelLevelOOTCnt = 0
                            OOTs(Index, index2).StorageLevelOOTCnt = 0
                        End If
                  
                        ' Only do all other checks if shift is active
'                        If Stn_ActiveShift(Index) = index2 Then
                        If (StationControl(Index, index2).Mode <> VBIDLE) Then
                                               
                            ' Tolerances in Engr Units (if reqd, convert from percent-of-fullscale)
                            Tol_pur_flow1 = (StationConfig(Index, index2).Tol_Pur_Flow / 100#) * (Stn_AIO(Index, asPurgeAirFlow).EuMax - Stn_AIO(Index, asPurgeAirFlow).EuMin)
                            LoLim_pur_flow1 = (StationConfig(Index, index2).LoLim_Purge_Flow / 100#) * (Stn_AIO(Index, asPurgeAirFlow).EuMax - Stn_AIO(Index, asPurgeAirFlow).EuMin)
                            Select Case STN_INFO(Index).Type
                                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                                    sGramsPerLiter = IIf(StationControl(Index, index2).Mode = VBLOAD, StationControl(Index, index2).BtnDensity, (GramsPerLiter * STN_INFO(Index).ButMfcDensityMult))
                                    Tol_btn_flow1 = GramsPerHourToSlpm(StationConfig(Index, index2).Tol_Btn_Flow, sGramsPerLiter)
                                    LoLim_btn_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asButaneFlow).EuMax - Stn_AIO(Index, asButaneFlow).EuMin)
                                    Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Nit_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                    LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                Case STN_ORVR2_TYPE
                                    If StationRecipe(Index, index2).UseHiRangeMFC Then
                                        ' use higher range MFC
                                        sGramsPerLiter = IIf(StationControl(Index, index2).Mode = VBLOAD, StationControl(Index, index2).BtnDensity, (GramsPerLiter * STN_INFO(Index).ButMfc2DensityMult))
                                        Tol_btn_flow1 = GramsPerHourToSlpm(StationConfig(Index, index2).Tol_ORVRBtn_Flow, sGramsPerLiter)
                                        LoLim_btn_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asButaneORVRFlow).EuMax - Stn_AIO(Index, asButaneORVRFlow).EuMin)
                                        Tol_nit_flow1 = (StationConfig(Index, index2).Tol_ORVRNit_Flow / 100#) * (Stn_AIO(Index, asNitrogenORVRFlow).EuMax - Stn_AIO(Index, asNitrogenORVRFlow).EuMin)
                                        LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asNitrogenORVRFlow).EuMax - Stn_AIO(Index, asNitrogenORVRFlow).EuMin)
                                    Else
                                        ' use lower range MFC
                                        sGramsPerLiter = IIf(StationControl(Index, index2).Mode = VBLOAD, StationControl(Index, index2).BtnDensity, (GramsPerLiter * STN_INFO(Index).ButMfcDensityMult))
                                        Tol_btn_flow1 = GramsPerHourToSlpm(StationConfig(Index, index2).Tol_Btn_Flow, sGramsPerLiter)
                                        LoLim_btn_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asButaneFlow).EuMax - Stn_AIO(Index, asButaneFlow).EuMin)
                                        Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Nit_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                        LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                    End If
                                Case STN_LIVEFUEL_TYPE
                                    Tol_btn_flow1 = 0
                                    LoLim_btn_flow1 = 0
                                    Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Lfv_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporFlow).EuMin)
                                    LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporFlow).EuMin)
                                Case STN_LIVEREG_TYPE
                                    If StationRecipe(Index, index2).LiveFuel Then
                                        Tol_btn_flow1 = 0
                                        LoLim_btn_flow1 = 0
                                        Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Lfv_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporFlow).EuMin)
                                        LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporFlow).EuMin)
                                    Else
                                        sGramsPerLiter = IIf(StationControl(Index, index2).Mode = VBLOAD, StationControl(Index, index2).BtnDensity, (GramsPerLiter * STN_INFO(Index).ButMfcDensityMult))
                                        Tol_btn_flow1 = GramsPerHourToSlpm(StationConfig(Index, index2).Tol_Btn_Flow, sGramsPerLiter)
                                        LoLim_btn_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asButaneFlow).EuMax - Stn_AIO(Index, asButaneFlow).EuMin)
                                        Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Nit_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                        LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                    End If
                                Case STN_LIVEORVR2_TYPE
                                    If StationRecipe(Index, index2).LiveFuel Then
                                        Tol_btn_flow1 = 0
                                        LoLim_btn_flow1 = 0
                                        If StationRecipe(Index, index2).UseHiRangeMFC Then
                                            ' use higher range MFC
                                            Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Lfv_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporORVRFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporORVRFlow).EuMin)
                                            LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporORVRFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporORVRFlow).EuMin)
                                        Else
                                            ' use lower range MFC
                                            Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Lfv_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporFlow).EuMin)
                                            LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asLiveFuelVaporFlow).EuMax - Stn_AIO(Index, asLiveFuelVaporFlow).EuMin)
                                        End If
                                    Else
                                        If StationRecipe(Index, index2).UseHiRangeMFC Then
                                            ' use higher range MFC
                                            sGramsPerLiter = IIf(StationControl(Index, index2).Mode = VBLOAD, StationControl(Index, index2).BtnDensity, (GramsPerLiter * STN_INFO(Index).ButMfc2DensityMult))
                                            Tol_btn_flow1 = GramsPerHourToSlpm(StationConfig(Index, index2).Tol_ORVRBtn_Flow, sGramsPerLiter)
                                            LoLim_btn_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asButaneORVRFlow).EuMax - Stn_AIO(Index, asButaneORVRFlow).EuMin)
                                            Tol_nit_flow1 = (StationConfig(Index, index2).Tol_ORVRNit_Flow / 100#) * (Stn_AIO(Index, asNitrogenORVRFlow).EuMax - Stn_AIO(Index, asNitrogenORVRFlow).EuMin)
                                            LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asNitrogenORVRFlow).EuMax - Stn_AIO(Index, asNitrogenORVRFlow).EuMin)
                                        Else
                                            ' use lower range MFC
                                            sGramsPerLiter = IIf(StationControl(Index, index2).Mode = VBLOAD, StationControl(Index, index2).BtnDensity, (GramsPerLiter * STN_INFO(Index).ButMfcDensityMult))
                                            Tol_btn_flow1 = GramsPerHourToSlpm(StationConfig(Index, index2).Tol_Btn_Flow, sGramsPerLiter)
                                            LoLim_btn_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asButaneFlow).EuMax - Stn_AIO(Index, asButaneFlow).EuMin)
                                            Tol_nit_flow1 = (StationConfig(Index, index2).Tol_Nit_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                            LoLim_nit_flow1 = (StationConfig(Index, index2).LoLim_Load_Flow / 100#) * (Stn_AIO(Index, asNitrogenFlow).EuMax - Stn_AIO(Index, asNitrogenFlow).EuMin)
                                        End If
                                    End If
                                Case STN_COMBO3_TYPE
                                    ' future
                                Case Else
                                    ' do nothing
                            End Select
                            
                            
                            ' CHECK PURGE FLOW
                            bOutHigh = IIf(Stn_AIO(Index, asPurgeAirFlow).EUValue > Stn_AIO(Index, asPurgeAirFlowSP).EUValue + Tol_pur_flow1, True, False)
                            bOutLow = IIf(Stn_AIO(Index, asPurgeAirFlow).EUValue < Stn_AIO(Index, asPurgeAirFlowSP).EUValue - Tol_pur_flow1, True, False)
                            If (StationControl(Index, index2).Mode = VBPURGE And PurgeControl(Index, index2).Phase = PurgePurging And Not (PurgeControl(Index, index2).InhibitOotCheck)) Then
                                bAboveLoLim = True
                                bOutZero = IIf(Stn_AIO(Index, asPurgeAirFlowSP).EUValue < StationRecipe(Index, index2).Purge_Flow - Tol_pur_flow1, True, False)
                            Else
                                bAboveLoLim = IIf(Stn_AIO(Index, asPurgeAirFlow).EUValue > LoLim_pur_flow1, True, False)
                                bOutZero = False
                            End If
                            If bOutZero _
                                     Or _
                                    (bAboveLoLim And (bOutHigh Or bOutLow)) Then
                                If OOTs(Index, index2).PurFlowOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                    OOTs(Index, index2).PurFlowOOTCnt = OOTs(Index, index2).PurFlowOOTCnt + CheckInterval
                                End If
                            Else
                                OOTs(Index, index2).PurFlowOOTCnt = 0
                            End If
                            
                            ' CHECK BUTANE FLOW
                            If STN_INFO(Index).Type = STN_REGULAR_TYPE _
                                    Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (Not StationRecipe(Index, index2).LiveFuel)) _
                                    Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (Not StationRecipe(Index, index2).LiveFuel)) _
                                    Or STN_INFO(Index).Type = STN_ORVR_TYPE _
                                    Or STN_INFO(Index).Type = STN_ORVR2_TYPE Then
                                If StationRecipe(Index, index2).UseHiRangeMFC Then
                                    ' use higher range MFC
                                    bOutHigh = IIf(Stn_AIO(Index, asButaneORVRFlow).EUValue > Stn_AIO(Index, asButaneORVRFlowSP).EUValue + Tol_btn_flow1, True, False)
                                    bOutLow = IIf(Stn_AIO(Index, asButaneORVRFlow).EUValue < Stn_AIO(Index, asButaneORVRFlowSP).EUValue - Tol_btn_flow1, True, False)
                                    If StationControl(Index, index2).Mode = VBLOAD And LoadControl(Index, index2).Phase = LoadLoading Then
                                        bAboveLoLim = True
                                        bOutZero = IIf((Stn_AIO(Index, asButaneORVRFlowSP).EUValue < (Stn_Btn_FlowSP(Index, index2) - Tol_btn_flow1)), True, False)
                                    Else
                                        bAboveLoLim = IIf((Stn_AIO(Index, asButaneORVRFlow).EUValue > LoLim_btn_flow1), True, False)
                                        bOutZero = False
                                    End If
                                Else
                                    ' use lower range MFC
                                    bOutHigh = IIf((Stn_AIO(Index, asButaneFlow).EUValue > (Stn_AIO(Index, asButaneFlowSP).EUValue + Tol_btn_flow1)), True, False)
                                    bOutLow = IIf((Stn_AIO(Index, asButaneFlow).EUValue < (Stn_AIO(Index, asButaneFlowSP).EUValue - Tol_btn_flow1)), True, False)
                                    If StationControl(Index, index2).Mode = VBLOAD And LoadControl(Index, index2).Phase = LoadLoading Then
                                        bAboveLoLim = True
                                        bOutZero = IIf((Stn_AIO(Index, asButaneFlowSP).EUValue < (Stn_Btn_FlowSP(Index, index2) - Tol_btn_flow1)), True, False)
                                    Else
                                        bAboveLoLim = IIf((Stn_AIO(Index, asButaneFlow).EUValue > LoLim_btn_flow1), True, False)
                                        bOutZero = False
                                    End If
                                End If
                                If bOutZero _
                                         Or _
                                        (bAboveLoLim And (bOutHigh Or bOutLow)) Then
                                    If OOTs(Index, index2).BtnFlowOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                        OOTs(Index, index2).BtnFlowOOTCnt = OOTs(Index, index2).BtnFlowOOTCnt + CheckInterval
                                    End If
                                Else
                                    OOTs(Index, index2).BtnFlowOOTCnt = 0
                                End If
                            Else
                                OOTs(Index, index2).BtnFlowOOTCnt = 0
                            End If
                        
                            ' CHECK NITROGEN/LIVEFUEL FLOW
                            If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (StationRecipe(Index, index2).LiveFuel)) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, index2).LiveFuel))) Then
                                ' use live fuel
                                If ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, index2).UseHiRangeMFC)) Then
                                    'use higher range live fuel
                                    nitSP = Stn_AIO(Index, asLiveFuelVaporORVRFlowSP).EUValue
                                    nitPV = Stn_AIO(Index, asLiveFuelVaporORVRFlow).EUValue
                                Else
                                    'use lower range live fuel
                                    nitSP = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                                    nitPV = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                                End If
                            ElseIf ((STN_INFO(Index).Type = STN_ORVR2_TYPE And StationRecipe(Index, index2).UseHiRangeMFC) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE And StationRecipe(Index, index2).UseHiRangeMFC)) Then
                                'use higher range nitrogen
                                nitSP = Stn_AIO(Index, asNitrogenORVRFlowSP).EUValue
                                nitPV = Stn_AIO(Index, asNitrogenORVRFlow).EUValue
                            Else
                                'use lower range nitrogen
                                nitSP = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                                nitPV = Stn_AIO(Index, asNitrogenFlow).EUValue
                            End If
                            bOutHigh = IIf(nitPV > nitSP + Tol_nit_flow1, True, False)
                            bOutLow = IIf(nitPV < nitSP - Tol_nit_flow1, True, False)
                            If StationControl(Index, index2).Mode = VBLOAD And LoadControl(Index, index2).Phase = LoadLoading Then
                                bAboveLoLim = True
                                bOutZero = IIf(nitSP < Stn_Nit_FlowSP(Index, index2) - Tol_nit_flow1, True, False)
                            Else
                                bAboveLoLim = IIf(nitPV > LoLim_nit_flow1, True, False)
                                bOutZero = False
                            End If
                            If bOutZero _
                                     Or _
                                    (bAboveLoLim And (bOutHigh Or bOutLow)) Then
                                If OOTs(Index, index2).NitFlowOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                    OOTs(Index, index2).NitFlowOOTCnt = OOTs(Index, index2).NitFlowOOTCnt + CheckInterval
                                End If
                            Else
                                OOTs(Index, index2).NitFlowOOTCnt = 0
                            End If
                    
                    
                            ' USING LIVEFUEL
                            If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And (StationRecipe(Index, index2).LiveFuel)) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And (StationRecipe(Index, index2).LiveFuel))) Then
                                ' CHECK LIVEFUEL TEMP
                                If ((STN_INFO(Index).ADF_TANKTYPE > 10) And (STN_INFO(Index).ADF_TANKTYPE <= 20)) Then
                                    bOutHigh = IIf(Stn_AIO(Index, asFuelTankTemp).EUValue > StationRecipe(Index, index2).ADF_HeaterSP + StationConfig(Index, index2).Tol_FuelTemp, True, False)
                                    bOutLow = IIf(Stn_AIO(Index, asFuelTankTemp).EUValue < StationRecipe(Index, index2).ADF_HeaterSP - StationConfig(Index, index2).Tol_FuelTemp, True, False)
                                    If (bOutHigh Or bOutLow) And StationControl(Index, index2).Mode = VBLOAD And StationRecipe(Index, index2).ADF_Heater Then
                                        If OOTs(Index, index2).FuelTempOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).FuelTempOOTCnt = OOTs(Index, index2).FuelTempOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).FuelTempOOTCnt = 0
                                    End If
                                Else
                                    OOTs(Index, index2).FuelTempOOTCnt = 0
                                End If
                                ' CHECK WATERBATH TEMP
                                If (STN_INFO(Index).ADF_TANKTYPE = 90) Then
                                    valPV = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
                                    bOutHigh = IIf(valPV > StationRecipe(Index, index2).ADF_HeaterSP + StationConfig(Index, index2).Tol_WaterBathTemp, True, False)
                                    bOutLow = IIf(valPV < StationRecipe(Index, index2).ADF_HeaterSP - StationConfig(Index, index2).Tol_WaterBathTemp, True, False)
                                    If (bOutHigh Or bOutLow) And StationControl(Index, index2).Mode = VBLOAD And StationRecipe(Index, index2).ADF_Heater Then
                                        If OOTs(Index, index2).WaterBathOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).WaterBathOOTCnt = OOTs(Index, index2).WaterBathOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).WaterBathOOTCnt = 0
                                    End If
                                Else
                                    OOTs(Index, index2).WaterBathOOTCnt = 0
                                End If
                                ' CHECK LOADRATE-OUT-OF-CONTROL
                                If (StationRecipe(Index, index2).UseLoadRatePID And (StationControl(Index, index2).Mode = VBLOAD)) Then
'                                    Select Case STN_INFO(Index).Type
'                                        Case STN_LIVEFUEL_TYPE, STN_LIVEREG_TYPE
'                                            valSP = CSng(0.99) * Stn_AIO(Index, asLiveFuelVaporFlowSP).EuMax
'                                            valPV = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
'                                        Case STN_LIVEORVR2_TYPE
'                                            If (StationRecipe(Index, index2).UseHiRangeMFC) Then
'                                                valSP = CSng(0.99) * Stn_AIO(Index, asLiveFuelVaporORVRFlowSP).EuMax
'                                                valPV = Stn_AIO(Index, asLiveFuelVaporORVRFlowSP).EUValue
'                                            Else
'                                                valSP = CSng(0.99) * Stn_AIO(Index, asLiveFuelVaporFlowSP).EuMax
'                                                valPV = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
'                                            End If
'                                    End Select
'                                    bOutHigh = IIf((valPV >= valSP), True, False)
'                                    bOutLow = False
                                    bOutHigh = IIf(LoadControl(Index, index2).LoadRate > LoadControl(Index, index2).LoadRateTarget * (1 + (StationConfig(Index, index2).Tol_Load_Total / 100)), True, False)
                                    bOutLow = IIf(LoadControl(Index, index2).LoadRate < LoadControl(Index, index2).LoadRateTarget * (1 - (StationConfig(Index, index2).Tol_Load_Total / 100)), True, False)
                                    If ((bOutHigh Or bOutLow) And (LoadControl(Index, index2).LoadRate <> 0) And (LoadControl(Index, index2).LoadRateTarget <> 0)) Then
                                        If OOTs(Index, index2).LoadRateOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).LoadRateOOTCnt = OOTs(Index, index2).LoadRateOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).LoadRateOOTCnt = 0
                                    End If
                                Else
                                    OOTs(Index, index2).LoadRateOOTCnt = 0
                                End If
                                ' CHECK LIVEFUEL DENSITY
                                deadweakMult = 1.25
                                If (StationControl(Index, index2).Mode = VBLOAD) Then
                                    If (((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And StationRecipe(Index, index2).UseHiRangeMFC And (Stn_AIO(Index, asLiveFuelVaporORVRFlow).EUValue > (0.05 * Stn_AIO(Index, asLiveFuelVaporORVRFlow).EuMax))) _
                                            Or _
                                        ((Not StationRecipe(Index, index2).UseHiRangeMFC) And (Stn_AIO(Index, asLiveFuelVaporFlow).EUValue > (0.05 * Stn_AIO(Index, asLiveFuelVaporFlow).EuMax)))) _
                                            Then
                                        bAboveLoLim = IIf((LoadControl(Index, index2).CurrLoadDensity > DeadLiveFuelDensity), True, False)
                                        If (Not bAboveLoLim) Then
                                            AdfControl(Index).LiveFuelDensityOkCnt = 0
                                            If AdfControl(Index).LiveFuelDensityDeadCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                                AdfControl(Index).LiveFuelDensityDeadCnt = AdfControl(Index).LiveFuelDensityDeadCnt + CheckInterval
                                            End If
                                            If AdfControl(Index).LiveFuelDensityWeakCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                                AdfControl(Index).LiveFuelDensityWeakCnt = AdfControl(Index).LiveFuelDensityWeakCnt + CheckInterval
                                            End If
                                        Else
                                            AdfControl(Index).LiveFuelDensityDeadCnt = 0
                                            bOutLow = IIf((LoadControl(Index, index2).CurrLoadDensity < (WeakLiveFuelDensity)), True, False)
                                            If (bOutLow) Then
                                                AdfControl(Index).LiveFuelDensityOkCnt = 0
                                                If AdfControl(Index).LiveFuelDensityWeakCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                                    AdfControl(Index).LiveFuelDensityWeakCnt = AdfControl(Index).LiveFuelDensityWeakCnt + CheckInterval
                                                End If
                                            Else
                                                If AdfControl(Index).LiveFuelDensityOkCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                                    AdfControl(Index).LiveFuelDensityOkCnt = AdfControl(Index).LiveFuelDensityOkCnt + CheckInterval
                                                End If
                                                AdfControl(Index).LiveFuelDensityWeakCnt = 0
                                            End If
                                        End If
                                    Else
                                        AdfControl(Index).LiveFuelDensityOkCnt = 0
                                        AdfControl(Index).LiveFuelDensityDeadCnt = 0
                                        AdfControl(Index).LiveFuelDensityWeakCnt = 0
                                    End If
                                Else
                                    AdfControl(Index).LiveFuelDensityOkCnt = 0
                                    AdfControl(Index).LiveFuelDensityDeadCnt = 0
                                    AdfControl(Index).LiveFuelDensityWeakCnt = 0
                                End If
                            Else
                                AdfControl(Index).LiveFuelDensityOkCnt = 0
                                AdfControl(Index).LiveFuelDensityDeadCnt = 0
                                AdfControl(Index).LiveFuelDensityWeakCnt = 0
                                OOTs(Index, index2).FuelTempOOTCnt = 0
                                OOTs(Index, index2).LoadRateOOTCnt = 0
                                OOTs(Index, index2).WaterBathOOTCnt = 0
                            End If
                            
                            ' CHECK PURGE AIR TEMP AND MOISTURE
                            If CHECKPAVALS And (StationControl(Index, index2).Mode = VBPURGE) Then
                                If Not PRG_INFO(STN_INFO(Index).AspiratorNum).UsingPrgReqHdw Or Not StationConfig(Index, index2).PosPressPurge Then 'don't check if using either remote purge cab or pospresspurge
                                    If Not Check_Tol(PATemp, StationConfig(Index, index2).Temp_Target, StationConfig(Index, index2).Tol_Temp) Then
                                        If OOTs(Index, index2).AirTempOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).AirTempOOTCnt = OOTs(Index, index2).AirTempOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).AirTempOOTCnt = 0
                                    End If
                                    If Not Check_Tol(PAMoisture, StationConfig(Index, index2).Moisture_Target, StationConfig(Index, index2).Tol_Moisture) Then
                                        If OOTs(Index, index2).AirMoistOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).AirMoistOOTCnt = OOTs(Index, index2).AirMoistOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).AirMoistOOTCnt = 0
                                    End If
                                Else
                                    OOTs(Index, index2).AirTempOOTCnt = 0
                                    OOTs(Index, index2).AirMoistOOTCnt = 0
                                End If
                            Else
                                OOTs(Index, index2).AirTempOOTCnt = 0
                                OOTs(Index, index2).AirMoistOOTCnt = 0
                            End If
            
                            ' CHECK PURGE DIFFERENTIAL PRESSURE
                            If USINGPURGEDP And (StationControl(Index, index2).Mode = VBPURGE) Then
                                ' don't check if using Positive Pressure Purge
                                If Not StationConfig(Index, index2).PosPressPurge Then
                                    If (Stn_AIO(Index, asPurgeDiffPress).EUValue > StationConfig(Index, index2).PurgeDP_HiLimit) Then
                                        If OOTs(Index, index2).PurgeDpOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                            OOTs(Index, index2).PurgeDpOOTCnt = OOTs(Index, index2).PurgeDpOOTCnt + CheckInterval
                                        End If
                                    Else
                                        OOTs(Index, index2).PurgeDpOOTCnt = 0
                                    End If
                                Else
                                    OOTs(Index, index2).PurgeDpOOTCnt = 0
                                End If
                            Else
                                OOTs(Index, index2).PurgeDpOOTCnt = 0
                            End If
            
                            ' CHECK PURGE OVEN TEMPERATURE
                            If USINGPURGEOVEN Then
                                bOutHigh = IIf(Stn_AIO(Index, asPurgeOvenTemp).EUValue > StationRecipe(Index, index2).PurgeOvenSP + StationConfig(Index, index2).Tol_PurgeOvenTemp, True, False)
                                bOutLow = IIf(Stn_AIO(Index, asPurgeOvenTemp).EUValue < StationRecipe(Index, index2).PurgeOvenSP - StationConfig(Index, index2).Tol_PurgeOvenTemp, True, False)
                                If (bOutHigh Or bOutLow) And StationControl(Index, index2).Mode = VBPURGE And StationRecipe(Index, index2).PurgeOven Then
                                    If OOTs(Index, index2).PurgeOvenOOTCnt <= (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                        OOTs(Index, index2).PurgeOvenOOTCnt = OOTs(Index, index2).PurgeOvenOOTCnt + CheckInterval
                                    End If
                                Else
                                    OOTs(Index, index2).PurgeOvenOOTCnt = 0
                                End If
                            Else
                                OOTs(Index, index2).PurgeOvenOOTCnt = 0
                            End If
            
                                
                        Else
                            ' Set all counts to zero (if Idle)
                            OOTs(Index, index2).PurFlowOOTCnt = 0
                            OOTs(Index, index2).BtnFlowOOTCnt = 0
                            OOTs(Index, index2).NitFlowOOTCnt = 0
                            OOTs(Index, index2).FuelTempOOTCnt = 0
                            OOTs(Index, index2).AirTempOOTCnt = 0
                            OOTs(Index, index2).AirMoistOOTCnt = 0
                            OOTs(Index, index2).LoadRateOOTCnt = 0
                            OOTs(Index, index2).PurgeDpOOTCnt = 0
                            OOTs(Index, index2).PurgeOvenOOTCnt = 0
                            AdfControl(Index).LiveFuelDensityDeadCnt = 0
                            AdfControl(Index).LiveFuelDensityWeakCnt = 0
                        End If  ' Not Idle
                        
                        ' Don't change OOT counts if in OOT
                    End If  ' Not OOT
                    
                Else
                    ' Set all counts to zero (if in Alarm)
                    OOTs(Index, index2).PurFlowOOTCnt = 0
                    OOTs(Index, index2).BtnFlowOOTCnt = 0
                    OOTs(Index, index2).NitFlowOOTCnt = 0
                    OOTs(Index, index2).FuelTempOOTCnt = 0
                    OOTs(Index, index2).AirTempOOTCnt = 0
                    OOTs(Index, index2).AirMoistOOTCnt = 0
                    OOTs(Index, index2).LoadRateOOTCnt = 0
                    OOTs(Index, index2).PurgeDpOOTCnt = 0
                    OOTs(Index, index2).PurgeOvenOOTCnt = 0
                    OOTs(Index, index2).FuelLevelOOTCnt = 0
                    OOTs(Index, index2).StorageLevelOOTCnt = 0
                    OOTs(Index, index2).WaterBathOOTCnt = 0
                    AdfControl(Index).LiveFuelDensityDeadCnt = 0
                    AdfControl(Index).LiveFuelDensityWeakCnt = 0
                End If  'Alarm Pause
            
            Next index2
        Next Index
                    
                
        ChgErrModule 13, 3132
                
        ' Set/Reset OOT Flags & Write to OOT Log
        For Index = 1 To LAST_STN
            For index2 = 1 To NR_SHIFT
                
                ' Only do OOT checks if station is not paused
                If Not StationControl(Index, index2).IsPausedInAlarm Then
                    ' Only do checks if not already in OOT
                    If StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                    
                        ' Live Fuel STORAGE TANK LEVEL OOT
                        If OOTs(Index, index2).StorageLevelOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                            If OOTs(Index, index2).StorageLevelOOT = False Then
                                OOTs(Index, index2).StorageLevelOOT = True
                                ootcomment = "LiveFuel Storage Tank Level OOT (SP=" & Format(FstControl(Index).LevelSP, "##0.0") _
                                                & " PV=" & Format(Stn_AIO(Index, asFuelTankTemp).EUValue, "##0.0") & ")"
                                If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                            End If
                        Else
                            If OOTs(Index, index2).StorageLevelOOT = True Then
                                OOTs(Index, index2).StorageLevelOOT = False
                                ootcomment = "LiveFuel Storage Tank Level Back in Tolerance"
                                If StationControl(Index, index2).Mode = VBLOAD Then
                                    ootcomment = "LiveFuel Storage Tank Level Back in Tolerance (PV=" & Format(Stn_AIO(Index, asStorageTankLevel).EUValue, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                    ootcomment = "LiveFuel Storage Tank Level Not OOT (PV=" & Format(Stn_AIO(Index, asStorageTankLevel).EUValue, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            End If
                        End If
                    
                        ' Live Fuel VAPOR TANK LEVEL OOT
                        If OOTs(Index, index2).FuelLevelOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                            If OOTs(Index, index2).FuelLevelOOT = False Then
                                OOTs(Index, index2).FuelLevelOOT = True
                                ootcomment = "LiveFuel Vapor Tank Level OOT (SP=" & Format(AdfControl(Index).LevelSP, "##0.0") _
                                                & " PV=" & Format(Stn_AIO(Index, asFuelTankLevel).EUValue, "##0.0") & ")"
                                If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                            End If
                        Else
                            If OOTs(Index, index2).FuelLevelOOT = True Then
                                OOTs(Index, index2).FuelLevelOOT = False
                                ootcomment = "LiveFuel Vapor Tank Level Back in Tolerance"
                                If StationControl(Index, index2).Mode = VBLOAD Then
                                    ootcomment = "LiveFuel Vapor Tank Level Back in Tolerance (PV=" & Format(Stn_AIO(Index, asFuelTankLevel).EUValue, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                    ootcomment = "LiveFuel Vapor Tank Level Not OOT (PV=" & Format(Stn_AIO(Index, asFuelTankLevel).EUValue, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            End If
                        End If
                    
                        ' Only do other checks if shift is active
                        If Stn_ActiveShift(Index) = index2 Then
                                
                            ' PURGE FLOW OOT
                            If OOTs(Index, index2).PurFlowOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).PurFlowOOT = False Then
                                    OOTs(Index, index2).PurFlowOOT = True
                                    ootcomment = "PurgeFlow OOT (SP=" & Format(Stn_AIO(Index, asPurgeAirFlowSP).EUValue, "##0.000") _
                                                    & " PV=" & Format(Stn_AIO(Index, asPurgeAirFlow).EUValue, "##0.000") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).PurFlowOOT = True Then
                                    OOTs(Index, index2).PurFlowOOT = False
                                    If StationControl(Index, index2).Mode = VBPURGE Then
                                        ootcomment = "Purge Flow Back in Tolerance (PV=" & Format(Stn_AIO(Index, asPurgeAirFlow).EUValue, "##0.000") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "Purge Flow Not OOT (PV=" & Format(Stn_AIO(Index, asPurgeAirFlow).EUValue, "##0.000") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' BUTANE FLOW OOT
                            Select Case STN_INFO(Index).Type
                                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                                    valPV = Stn_AIO(Index, asButaneFlow).EUValue
                                    valSP = Stn_AIO(Index, asButaneFlowSP).EUValue
                                Case STN_ORVR2_TYPE
                                    If StationRecipe(Index, index2).UseHiRangeMFC Then
                                        ' use higher range MFC
                                        valPV = Stn_AIO(Index, asButaneORVRFlow).EUValue
                                        valSP = Stn_AIO(Index, asButaneORVRFlowSP).EUValue
                                    Else
                                        ' use lower range MFC
                                        valPV = Stn_AIO(Index, asButaneFlow).EUValue
                                        valSP = Stn_AIO(Index, asButaneFlowSP).EUValue
                                    End If
                                Case STN_LIVEFUEL_TYPE
                                    valPV = 0
                                    valSP = 0
                                Case STN_LIVEREG_TYPE
                                    If StationRecipe(Index, index2).LiveFuel Then
                                        valPV = 0
                                        valSP = 0
                                    Else
                                        valPV = Stn_AIO(Index, asButaneFlow).EUValue
                                        valSP = Stn_AIO(Index, asButaneFlowSP).EUValue
                                    End If
                                Case STN_LIVEORVR2_TYPE
                                    If StationRecipe(Index, index2).LiveFuel Then
                                        valPV = 0
                                        valSP = 0
                                    Else
                                        If StationRecipe(Index, index2).UseHiRangeMFC Then
                                            ' use higher range MFC
                                            valPV = Stn_AIO(Index, asButaneORVRFlow).EUValue
                                            valSP = Stn_AIO(Index, asButaneORVRFlowSP).EUValue
                                        Else
                                            ' use lower range MFC
                                            valPV = Stn_AIO(Index, asButaneFlow).EUValue
                                            valSP = Stn_AIO(Index, asButaneFlowSP).EUValue
                                        End If
                                    End If
                                Case STN_COMBO3_TYPE
                                    ' future
                                Case Else
                                    ' do nothing
                            End Select
                            If OOTs(Index, index2).BtnFlowOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).BtnFlowOOT = False Then
                                    OOTs(Index, index2).BtnFlowOOT = True
                                    ootcomment = "ButaneFlow OOT (SP=" & Format(valSP, "##0.000") _
                                                    & " PV=" & Format(valPV, "##0.000") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).BtnFlowOOT = True Then
                                    OOTs(Index, index2).BtnFlowOOT = False
                                    If StationControl(Index, index2).Mode = VBLOAD Then
                                        ootcomment = "Butane Flow Back in Tolerance (PV=" & Format(valPV, "##0.000") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "Butane Flow Not OOT (PV=" & Format(valPV, "##0.000") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' NITROGEN FLOW OOT  (or LIVEFUEL, if a Live Fuel Station)
                            Select Case STN_INFO(Index).Type
                                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                                    ootcomment = "Nitrogen Flow"
                                    valPV = Stn_AIO(Index, asNitrogenFlow).EUValue
                                    valSP = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                                Case STN_ORVR2_TYPE
                                    ootcomment = "Nitrogen Flow"
                                    If StationRecipe(Index, index2).UseHiRangeMFC Then
                                        ' use higher range MFC
                                        valPV = Stn_AIO(Index, asNitrogenORVRFlow).EUValue
                                        valSP = Stn_AIO(Index, asNitrogenORVRFlowSP).EUValue
                                    Else
                                        ' use lower range MFC
                                        valPV = Stn_AIO(Index, asNitrogenFlow).EUValue
                                        valSP = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                                    End If
                                Case STN_LIVEFUEL_TYPE
                                    ootcomment = "Vapor Carrier Flow"
                                    valPV = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                                    valSP = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                                Case STN_LIVEREG_TYPE
                                    If StationRecipe(Index, index2).LiveFuel Then
                                        ootcomment = "Vapor Carrier Flow"
                                        valPV = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                                        valSP = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                                    Else
                                        ootcomment = "Nitrogen Flow"
                                        valPV = Stn_AIO(Index, asNitrogenFlow).EUValue
                                        valSP = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                                    End If
                                Case STN_LIVEORVR2_TYPE
                                    If StationRecipe(Index, index2).LiveFuel Then
                                        ootcomment = "Vapor Carrier Flow"
                                        If StationRecipe(Index, index2).UseHiRangeMFC Then
                                            ' use higher range MFC
                                            valPV = Stn_AIO(Index, asLiveFuelVaporORVRFlow).EUValue
                                            valSP = Stn_AIO(Index, asLiveFuelVaporORVRFlowSP).EUValue
                                        Else
                                            ' use lower range MFC
                                            valPV = Stn_AIO(Index, asLiveFuelVaporFlow).EUValue
                                            valSP = Stn_AIO(Index, asLiveFuelVaporFlowSP).EUValue
                                        End If
                                    Else
                                        ootcomment = "Nitrogen Flow"
                                        If StationRecipe(Index, index2).UseHiRangeMFC Then
                                            ' use higher range MFC
                                            valPV = Stn_AIO(Index, asNitrogenORVRFlow).EUValue
                                            valSP = Stn_AIO(Index, asNitrogenORVRFlowSP).EUValue
                                        Else
                                            ' use lower range MFC
                                            valPV = Stn_AIO(Index, asNitrogenFlow).EUValue
                                            valSP = Stn_AIO(Index, asNitrogenFlowSP).EUValue
                                        End If
                                    End If
                                Case STN_COMBO3_TYPE
                                    ' future
                                Case Else
                                    ' do nothing
                            End Select
                            If OOTs(Index, index2).NitFlowOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).NitFlowOOT = False Then
                                    OOTs(Index, index2).NitFlowOOT = True
                                    ootcomment = ootcomment & " OOT (SP=" & Format(valSP, "##0.000") _
                                                    & " PV=" & Format(valPV, "##0.000") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).NitFlowOOT = True Then
                                    OOTs(Index, index2).NitFlowOOT = False
                                    If StationControl(Index, index2).Mode = VBLOAD Then
                                        ootcomment = ootcomment & " Back in Tolerance (PV=" & Format(valPV, "##0.000") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = ootcomment & " Not OOT (PV=" & Format(valPV, "##0.000") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' Live Fuel TANK TEMP OOT
                            If OOTs(Index, index2).FuelTempOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).FuelTempOOT = False Then
                                    OOTs(Index, index2).FuelTempOOT = True
                                    ootcomment = "FuelTemp OOT (SP=" & Format(StationRecipe(Index, index2).ADF_HeaterSP, "##0.0") _
                                                    & " PV=" & Format(Stn_AIO(Index, asFuelTankTemp).EUValue, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).FuelTempOOT = True Then
                                    OOTs(Index, index2).FuelTempOOT = False
                                    ootcomment = "LiveFuel Tank Temp Back in Tolerance"
                                    If StationControl(Index, index2).Mode = VBLOAD Then
                                        ootcomment = "LiveFuel Tank Temp Back in Tolerance (PV=" & Format(Stn_AIO(Index, asFuelTankTemp).EUValue, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "LiveFuel Tank Temp Not OOT (PV=" & Format(Stn_AIO(Index, asFuelTankTemp).EUValue, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' Live Fuel LOADRATE OOT (Out-Of-Control)
                            sDwell = CSng(600)
                            If (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) > sDwell Then sDwell = (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time)
                            If (OOTs(Index, index2).LoadRateOOTCnt > sDwell) Then
                                If OOTs(Index, index2).LoadRateOOT = False Then
                                    OOTs(Index, index2).LoadRateOOT = True
                                    ootcomment = "LoadRate OOT (SP=" & Format(LoadControl(Index, index2).LoadRateTarget, "##0.0") _
                                                    & " PV=" & Format(LoadControl(Index, index2).LoadRate, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).LoadRateOOT = True Then
                                    OOTs(Index, index2).LoadRateOOT = False
                                    ootcomment = "LiveFuel LoadRate Back in Tolerance"
                                    If StationControl(Index, index2).Mode = VBLOAD Then
                                        ootcomment = "LiveFuel LoadRate Back in Tolerance (PV=" & Format(LoadControl(Index, index2).LoadRate, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "LiveFuel LoadRate Not OOT; Not Loading"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' WATERBATH TEMP OOT
                            If OOTs(Index, index2).WaterBathOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).WaterBathOOT = False Then
                                    OOTs(Index, index2).WaterBathOOT = True
                                    valPV = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
                                    ootcomment = "WaterBath OOT (SP=" & Format(StationRecipe(Index, index2).ADF_HeaterSP, "##0.0") _
                                                    & " PV=" & Format(valPV, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).WaterBathOOT = True Then
                                    OOTs(Index, index2).WaterBathOOT = False
                                    ootcomment = "WaterBath Temp Back in Tolerance"
                                    valPV = IIf(USINGC, LF_Chiller.PvIn, DegCtoF(LF_Chiller.PvIn))
                                    If StationControl(Index, index2).Mode = VBLOAD Then
                                        ootcomment = "WaterBath Temp Back in Tolerance (PV=" & Format(valPV, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "WaterBath Temp Not OOT (PV=" & Format(valPV, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' PURGE AIR TEMP OOT
                            If OOTs(Index, index2).AirTempOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).AirTempOOT = False Then
                                    OOTs(Index, index2).AirTempOOT = True
                                    ootcomment = "Purge Temp OOT (SP=" & Format(StationConfig(Index, index2).Temp_Target, "##0.0") _
                                                    & " PV=" & Format(PATemp, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).AirTempOOT = True Then
                                    OOTs(Index, index2).AirTempOOT = False
                                    If StationControl(Index, index2).Mode = VBPURGE Then
                                        ootcomment = "Purge Air Temp Back in Tolerance (PV=" & Format(PATemp, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "Purge Air Temp Not OOT (PV=" & Format(PATemp, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' PURGE AIR MOISTURE OOT
                            If OOTs(Index, index2).AirMoistOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).AirMoistOOT = False Then
                                    OOTs(Index, index2).AirMoistOOT = True
                                    ootcomment = "Purge Air Moisture OOT (SP=" & Format(StationConfig(Index, index2).Moisture_Target, "##0.0") _
                                                    & " PV=" & Format(PAMoisture, "##0.0") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).AirMoistOOT = True Then
                                    OOTs(Index, index2).AirMoistOOT = False
                                    If StationControl(Index, index2).Mode = VBPURGE Then
                                        ootcomment = "Purge Moisture Back in Tolerance (PV=" & Format(PAMoisture, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "Purge Moisture Not OOT (PV=" & Format(PAMoisture, "##0.0") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                
                            ' PURGE DIFFERENTIAL PRESSURE OOT
                            If OOTs(Index, index2).PurgeDpOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).PurgeDpOOT = False Then
                                    OOTs(Index, index2).PurgeDpOOT = True
                                    ootcomment = "Purge DP OOT (Limit=" & Format(StationConfig(Index, index2).PurgeDP_HiLimit, "##0.0#") _
                                                    & " DP=" & Format(Stn_AIO(Index, asPurgeDiffPress).EUValue, "##0.0#") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).PurgeDpOOT = True Then
                                    OOTs(Index, index2).PurgeDpOOT = False
                                    If StationControl(Index, index2).Mode = VBPURGE Then
                                        ootcomment = "Purge DP Back in Tolerance (DP=" & Format(Stn_AIO(Index, asPurgeDiffPress).EUValue, "##0.0#") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "Purge DP Not OOT (DP=" & Format(Stn_AIO(Index, asPurgeDiffPress).EUValue, "##0.0#") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' PURGE OVEN TEMPERATURE OOT
                            If OOTs(Index, index2).PurgeOvenOOTCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If OOTs(Index, index2).PurgeOvenOOT = False Then
                                    OOTs(Index, index2).PurgeOvenOOT = True
                                    ootcomment = "Purge Oven OOT (Tol=" & Format(StationConfig(Index, index2).Tol_PurgeOvenTemp, "###0.0#") _
                                                    & " Temp=" & Format(Stn_AIO(Index, asPurgeOvenTemp).EUValue, "###0.0#") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            Else
                                If OOTs(Index, index2).PurgeOvenOOT = True Then
                                    OOTs(Index, index2).PurgeOvenOOT = False
                                    If StationControl(Index, index2).Mode = VBPURGE Then
                                        ootcomment = "Purge Oven Back in Tolerance (Temp=" & Format(Stn_AIO(Index, asPurgeOvenTemp).EUValue, "###0.0#") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    ElseIf StationControl(Index, index2).Mode <> VBPAUSEOOT Then
                                        ootcomment = "Purge Oven Not OOT (Temp=" & Format(Stn_AIO(Index, asPurgeOvenTemp).EUValue, "###0.0#") & ")"
                                        If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                    End If
                                End If
                            End If
                    
                            ' Live Fuel Density is very low (fuel is dead)
                            sDwell = CSng(120)
                            If (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) > sDwell Then sDwell = (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time)
                            If (AdfControl(Index).LiveFuelDensityDeadCnt > sDwell) Then
                                If AdfControl(Index).LiveFuelState <> fuelDead Then
                                    AdfControl(Index).LiveFuelState = fuelDead
                                    ootcomment = "LF Density is Dead (SP=" & Format(DeadLiveFuelDensity, "##0.0##") _
                                                    & " PV=" & Format(LoadControl(Index, index2).CurrLoadDensity, "##0.0##") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            'End If
                            ' Live Fuel Density is low (fuel is weak)
                            ElseIf (AdfControl(Index).LiveFuelDensityWeakCnt > sDwell) Then
                                If AdfControl(Index).LiveFuelState <> fuelWeak Then
                                    AdfControl(Index).LiveFuelState = fuelWeak
                                    ootcomment = "LF Density is Weak (SP=" & Format(WeakLiveFuelDensity, "##0.0##") _
                                                    & " PV=" & Format(LoadControl(Index, index2).CurrLoadDensity, "##0.0##") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            End If
                            ' Live Fuel Density is OK
                            If AdfControl(Index).LiveFuelDensityOkCnt > (StationConfig(Index, index2).OOTtimeDelay + MFC_Settle_Time) Then
                                If AdfControl(Index).LiveFuelState <> fuelOK Then
                                    AdfControl(Index).LiveFuelState = fuelOK
                                    ootcomment = "LF Density is Ok (SP=" & Format(WeakLiveFuelDensity, "##0.0##") _
                                                    & " PV=" & Format(LoadControl(Index, index2).CurrLoadDensity, "##0.0##") & ")"
                                    If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, ootcomment
                                End If
                            End If
                    
                        
                        Else
                        
                            ' Set All OOT's False(if Mode=Idle)
                            OOTs(Index, index2).PurFlowOOT = False
                            OOTs(Index, index2).PurFlowOOT = False
                            OOTs(Index, index2).BtnFlowOOT = False
                            OOTs(Index, index2).NitFlowOOT = False
                            OOTs(Index, index2).FuelTempOOT = False
                            OOTs(Index, index2).AirTempOOT = False
                            OOTs(Index, index2).AirMoistOOT = False
                            OOTs(Index, index2).LoadRateOOT = False
                            OOTs(Index, index2).FuelTempOOT = False
                            OOTs(Index, index2).FuelLevelOOT = False
                            OOTs(Index, index2).StorageLevelOOT = False
                            AdfControl(Index).LiveFuelState = fuelOK
                        
                        End If
                        
                        ' Don't change OOTs if Mode=PauseOOT
                    End If  ' Not OOT
                    
                Else
                
                    ' Set all OOTs to False (if in Alarm)
                    OOTs(Index, index2).PurFlowOOT = False
                    OOTs(Index, index2).PurFlowOOT = False
                    OOTs(Index, index2).BtnFlowOOT = False
                    OOTs(Index, index2).NitFlowOOT = False
                    OOTs(Index, index2).FuelTempOOT = False
                    OOTs(Index, index2).AirTempOOT = False
                    OOTs(Index, index2).AirMoistOOT = False
                    OOTs(Index, index2).LoadRateOOT = False
                    OOTs(Index, index2).PurgeDpOOT = False
                    OOTs(Index, index2).FuelLevelOOT = False
                    OOTs(Index, index2).StorageLevelOOT = False
'                    AdfControl(Index).LiveFuelState = fuelOK
                        
                End If  'Alarm Pause
            
            Next index2
        Next Index
        
        ' ******************************************************************
        
        ChgErrModule 13, 3133
        
        For Index = 1 To LAST_STN
            For index2 = 1 To NR_SHIFT
                ' Only do OOT checks if station is not paused
                If Not StationControl(Index, index2).IsPausedInAlarm Then
                    
                
                    '  option USINGCANVENT first user = Carb 5/6/03
                    '  test only when in load mode after OOT delay timed out and OOT count to
                    '  eliminate contact bounce
                    If USINGCANVENTALARM Then
                        If StationControl(Index, index2).Mode = VBLOAD Then
                        
                            ' Wait for OOT Time Delay before checking for Canvent OOT
                            If (StationControl(Index, index2).Mode_StartDts + TimeSerial(0, 0, StationConfig(Index, index2).OOTtimeDelay) < Now) Then
                                ' Wait for Canvent FS Override Delay before checking for Canvent OOT
                                If Not OOTs(Index, index2).CanVent_DelayOn Then
                                
                                    ' Check the Canvent Flow Switch
                                    If Not Stn_DIO(Index, isCanVentAlarmSw).Value Then    ' Switch not made
                                        If (OOTs(Index, index2).CanVentOOTCnt >= StationConfig(Index, index2).OOTtimeDelay) Then  ' Not made for too long
                                            If (Not OOTs(Index, index2).CanVentOOT) Then
                                                Write_ELog "Canister Vent, Stn " & Index & " ALARM"
                                                If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, "Canister Vent Alarm Set"
                                                OOTs(Index, index2).CanVentOOT = True
                                            End If
                                        Else                                        ' got alarm for how long
                                            OOTs(Index, index2).CanVentOOTCnt = OOTs(Index, index2).CanVentOOTCnt + 1
                                        End If
                                    Else                                        ' not in alarm but were we
                                        If OOTs(Index, index2).CanVentOOT Then
                                            Write_ELog "Canister Vent, Stn " & Index & " Cleared"
                                            If Len(StationControl(Index, index2).DBFile) > 0 Then OOT_Write Index, index2, "Canister Vent Alarm Cleared"
                                            OOTs(Index, index2).CanVentOOTCnt = 0
                                            OOTs(Index, index2).CanVentOOT = False
                                        End If
                                    End If
                                    
                                Else
                                    OOTs(Index, index2).CanVentOOTCnt = 0
                                End If                                      ' Canvent FS Override Delay
                            Else
                                OOTs(Index, index2).CanVentOOTCnt = 0
                            End If                                          ' OOT Time Delay
                            
                        Else
                            OOTs(Index, index2).CanVentOOTCnt = 0
                        End If                                              ' StationControl(Index, index2).Mode = VBLOAD
                        
                    Else
                        OOTs(Index, index2).CanVentOOTCnt = 0
                    End If                                                  ' using canister vent alarm
                
                Else
                    OOTs(Index, index2).CanVentOOTCnt = 0
                End If
                
            Next index2
        Next Index
        
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

Function Check_Tol(Value As Single, Target As Single, tol As Single) As Boolean
'
' Function Name:    Check_Tol
' Author:           Analytical Process Programmer     8/8/96
' Description:      Checks to see if value is within tolerance.
'                   If value is out of tolerance, result is false.
'                   If value is in tolerance, result is true.
'
SetErrModule 13, 2
Dim tempVal As Boolean
    tempVal = True
    If Value > Target + tol Then tempVal = False
    If Value < Target - tol Then tempVal = False
    Check_Tol = tempVal
ResetErrModule
End Function

Sub Alarm_Check()
'
' Function Name:    Check Alarms
' Author:           Brunrose
' Description:      This routine checks alarm conditions and
'                   sets alarm status flags for each station.
'                   Updates the alarm log for each station
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 1
Dim Index As Integer
Dim index2 As Integer
Dim SomebodyPurging As Integer
Dim SomebodyUsingFID As Integer
Dim SomebodyAlarmOOT As Integer
Dim mixstation As Integer
Dim mixshift As Integer
Dim allIdleFlag As Boolean
Dim TestMixBlower As Boolean
Dim tempstr As String
Static horndisabled As Integer
Static purgevacuumtoggle As Integer
Static SafetySwCounter As Integer

' GoTo bypass2
    
    If OptoReadAllOnce Then
        
        'Check for display error conditions open door, customer low pressure,
        ' 1/2 customer input, load/purge delays , Low Butane gas, High Live Fuel Level
        If USINGDOOROPEN Then DoorOpened
        If USING_EXT_CONTACTS Then ExtAlarmInput
        If USINGCUSTOMERLOWGAS Then CustomerLowGas
        If USINGSYSTEMVACSW Then SystemVacSw
        If USINGUPS > 0 Then UpsTest
        
        ' check for Maintenance Mode
        If ((Com_DIO(icMaintSw).addr + Com_DIO(icMaintSw).chan) > 0) Then
            ' Maintenance Mode Input is in use
            If (Com_DIO(icMaintSw).Value) Then
                ' Maintenance Mode is ON
                If Not MaintMode Then
                    ' just turned on
                    MaintMode = True
                    tempstr = "Maintenance Mode On"
'                    tempstr = "Maintenance Mode On" & " @ " & Format(Timer, "###,##0.000")
'    Debug.Print tempstr
                    ' document it
                    Write_ELog tempstr
                    For Index = 1 To LAST_STN
                        For index2 = 1 To NR_SHIFT
                            If (StationControl(Index, index2).TestTimerIsRunning) Then Write_JLog Index, index2, tempstr
                        Next index2
                    Next Index
                End If
            Else
                ' Maintenance Mode is off
                If MaintMode Then
                    ' just turned off
                    MaintMode = False
                    tempstr = "Maintenance Mode Off"
'                    tempstr = "Maintenance Mode Off" & " @ " & Format(Timer, "###,##0.000")
'    Debug.Print tempstr
                    ' document it
                    Write_ELog tempstr
                    For Index = 1 To LAST_STN
                        For index2 = 1 To NR_SHIFT
                            If (StationControl(Index, index2).TestTimerIsRunning) Then Write_JLog Index, index2, tempstr
                        Next index2
                    Next Index
                End If
            End If
        Else
            ' not using Maintenance Mode
            MaintMode = False
        End If
        
        ' Always look to see if the bottle is empty yet
        If systemhasBUTANE Then LowButaneGas                        ' Always test butane usage
            
        ' Index 0 is overall alarm for main menu screen
        
        ' Clear Alarm Data
        ' true = kill noise
        If (Com_DIO(icHornSilencePB).Value Or SilenceHornRequest) Then
            SilenceHornRequest = False
            Com_OutDigital icAlarmHorn, cOFF                                  ' ok Main box Alarm Push button pressed reset alarms
            horndisabled = 1
        End If
         
        ' EStop true is good
        If USING_ESTOP_INPUT Then
            If Alm_Estop And Com_DIO(icEStopSw).Value Then                  'reset
                 ALM_Write 1, 1, "Emergency Stop Cleared"
                 Alm_Estop = False
            Else
                If Not Alm_Estop And Not Com_DIO(icEStopSw).Value Then    'Set
                    ALM_Write 1, 1, "Emergency Stop Active"
                    Alm_Estop = True
                End If
            End If
        Else
            Alm_Estop = False
        End If
          
        ' main20lel true is good
        If Alm_Btn20 And Com_DIO(ic20LelGasSw).Value Then
            ALM_Write 0, 1, "20 Percent LEL Cleared"
            Alm_Btn20 = False
        Else
            If Not Alm_Btn20 And Not Com_DIO(ic20LelGasSw).Value Then
                ALM_Write 0, 1, "20 Percent LEL Active"
                Alm_Btn20 = True
            End If
        End If
          
        ' Main N2 (Only for Live Fuel with Heaters)
        If (systemhasADF_HEATER And ((Com_DIO(icLiveFuelPurgePS).addr + Com_DIO(icLiveFuelPurgePS).chan) > 0)) Then
            If Alm_N2 And Com_DIO(icLiveFuelPurgePS).Value Then                  'reset
                 ALM_Write 0, 1, "Main N2 Pressure Switch Back On"
                 Alm_N2 = False
            Else
                If Not Alm_N2 And Not Com_DIO(icLiveFuelPurgePS).Value Then    'Set
                    ALM_Write 0, 1, "Main N2 Pressure Switch Off"
                    Alm_N2 = True
                End If
            End If
        Else
            Alm_N2 = False
        End If
          
    ' GoTo bypassAlmFlow
          
        ' opto_mainexaustflowsw true is good
        If Alm_Flow And Com_DIO(icExhaustFlowFS).Value Then
            ALM_Write 0, 1, "Exhaust Flow Alarm Cleared"
            Alm_Flow = False
        Else
            If Not Alm_Flow And Not Com_DIO(icExhaustFlowFS).Value Then
                ALM_Write 0, 1, "Exhaust Flow Alarm Active"
                Alm_Flow = True
            End If
        End If
        
bypassAlmFlow:
    ' GoTo bypassAlmVac
        
        ' Check Purge Vacuum Switch(s)
        For Index = 1 To NR_PRGAIR
            If PRG_INFO(Index).UsingVacSwHdw Then                      ' Vacuum Switch(s) in use on this PurgeAir Source
                    
                ' Check Vac Switch Only if PurgeAir Source Aspirator is On
                If Prg_DIO(Index, ipPiabSol).Value And Not Prg_DIO(Index, ipPurgeVacuumSw).Value Then
                    Alm_Vac_Count(Index) = Alm_Vac_Count(Index) + 1
                Else
                    Alm_Vac_Count(Index) = 0
                End If
                   
            Else
            
                Alm_Vac(Index) = False
        
            End If
        Next Index
         
        ' Check Purge Vacuum Sw Alarms
        For Index = 1 To NR_PRGAIR
            If PRG_INFO(Index).UsingVacSwHdw Then                      ' Vacuum Switch(s) in use on this PurgeAir Source
                        
                ' Set/Reset Vac Alarm
                If Alm_Vac_Count(Index) < 10 Then
                    ' No Vac Switch Alarm Present Now
                    If Alm_Vac(Index) Then
                        If Prg_DIO(Index, ipPiabSol).Value Then
                            ALM_Write CInt(Index), 1, "Purge Vacuum Switch Alarm Cleared"
                        Else
                            ALM_Write CInt(Index), 1, "PIAB Turned Off"
                        End If
                        Alm_Vac(Index) = False
                    End If
                Else
                   ' Vacuum Switch Alarm Present Now
                    If Not Alm_Vac(Index) Then
                        ALM_Write CInt(Index), 1, "Purge Vacuum Switch Alarm Active"
                        Alm_Vac(Index) = True
                    End If
                End If
                
            Else
            
                Alm_Vac(Index) = False
                
            End If
        Next Index
         
         
bypassAlmVac:
    ' GoTo adf
         
         
adf:
        
        For Index = 1 To LAST_STN
            index2 = IIf((Stn_ActiveShift(Index) > 0), Stn_ActiveShift(Index), 1)
            If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(Index).Type = STN_LIVEREG_TYPE And StationRecipe(Index, index2).LiveFuel) Or (STN_INFO(Index).Type = STN_LIVEORVR2_TYPE And StationRecipe(Index, index2).LiveFuel)) Then
                
                ' HiHi Level Alarm
                If ((Stn_DIO(Index, isFuelHiHiLevelLS).addr + Stn_DIO(Index, isFuelHiHiLevelLS).chan) > 0) Then
                    ' using Tank HiHi LS
                    If Stn_DIO(Index, isFuelHiHiLevelLS).Value Then
                        If Not Alm_LiveFuelLevel(Index, 1) Then
                            Alm_LiveFuelLevel(Index, 1) = True
                            tempstr = "High Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " Level Alarm Active"
                            If StationControl(Index, index2).TestTimerIsRunning Then
                                ' open job; write to Job Alarms Log
                                ALM_Write CInt(Index), 1, tempstr
                            Else
                                ' no open job; write to System Events Log
                                Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                            End If
                        End If
                    Else
                        If Alm_LiveFuelLevel(Index, 1) Then
                            Alm_LiveFuelLevel(Index, 1) = False
                            tempstr = "High Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " Level Alarm Cleared"
                            If StationControl(Index, index2).TestTimerIsRunning Then
                                ' open job; write to Job Alarms Log
                                ALM_Write CInt(Index), 1, tempstr
                            Else
                                ' no open job; write to System Events Log
                                Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                            End If
                        End If
                    End If
                Else
                    ' Tank HiHi LS not in use
                    Alm_LiveFuelLevel(Index, 1) = False
                End If
                
                If (STN_INFO(Index).ADF_DEF.hasADF_Heater) And StationRecipe(Index, 1).ADF_Heater Then
                    ' Fuel & Sheath OverTemp Alarms
                    If Stn_DIO(Index, isFuelOverTempSw).Value Then
                        If Not Alm_LiveFuelHeater(Index, 1) Then
                            Alm_LiveFuelHeater(Index, 1) = True
                            tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " Fuel OverTemp Alarm Active"
                            If StationControl(Index, index2).TestTimerIsRunning Then
                                ' open job; write to Job Alarms Log
                                ALM_Write CInt(Index), 1, tempstr
                            Else
                                ' no open job; write to System Events Log
                                Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                            End If
                        Else
                            If Alm_LiveFuelHeater(Index, 1) Then
                                Alm_LiveFuelHeater(Index, 1) = False
                                tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " OverTemp Alarm Cleared"
                                If StationControl(Index, index2).TestTimerIsRunning Then
                                    ' open job; write to Job Alarms Log
                                    ALM_Write CInt(Index), 1, tempstr
                                Else
                                    ' no open job; write to System Events Log
                                    Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                                End If
                            End If
                        End If
                    ElseIf Stn_DIO(Index, isSheathOverTempSw).Value Then
                        If Not Alm_LiveFuelHeater(Index, 1) Then
                            Alm_LiveFuelHeater(Index, 1) = True
                            tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " Sheath OverTemp Alarm Active"
                            If StationControl(Index, index2).TestTimerIsRunning Then
                                ' open job; write to Job Alarms Log
                                ALM_Write CInt(Index), 1, tempstr
                            Else
                                ' no open job; write to System Events Log
                                Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                            End If
                        Else
                            If Alm_LiveFuelHeater(Index, 1) Then
                                Alm_LiveFuelHeater(Index, 1) = False
                                tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " OverTemp Alarm Cleared"
                                If StationControl(Index, index2).TestTimerIsRunning Then
                                    ' open job; write to Job Alarms Log
                                    ALM_Write CInt(Index), 1, tempstr
                                Else
                                    ' no open job; write to System Events Log
                                    Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                                End If
                            End If
                        End If
                    Else
                        If Alm_LiveFuelHeater(Index, 1) Then
                            Alm_LiveFuelHeater(Index, 1) = False
                            tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " OverTemp Alarm Cleared"
                            If StationControl(Index, index2).TestTimerIsRunning Then
                                ' open job; write to Job Alarms Log
                                ALM_Write CInt(Index), 1, tempstr
                            Else
                                ' no open job; write to System Events Log
                                Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                            End If
                        End If
                    End If
                    ' Safety Switch Alarm
                    If (AdfControl(Index).Step = 39 Or AdfControl(Index).Step >= 99) Then
                        ' heating is happening
                        If Not Stn_DIO(Index, isFuelSafetyLevelLS).Value Then
                            If (StationConfig(Index, 1).OOTtimeDelay > SafetySwCounter) Then SafetySwCounter = SafetySwCounter + 1
                        Else
                            SafetySwCounter = 0
                        End If
                        If (SafetySwCounter >= StationConfig(Index, 1).OOTtimeDelay) Then
                            If Not Alm_LiveFuelSafety(Index, 1) Then
                                Alm_LiveFuelSafety(Index, 1) = True
                                tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " SafetyLevel Alarm Active"
                                If StationControl(Index, index2).TestTimerIsRunning Then
                                    ' a job is open; write to Job Alarms Log
                                    ALM_Write CInt(Index), 1, tempstr
                                Else
                                    ' no open job; write to System Events Log
                                    Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                                End If
                            End If
                        Else
                            If Alm_LiveFuelSafety(Index, 1) Then
                                Alm_LiveFuelSafety(Index, 1) = False
                                tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " SafetyLevel Alarm Cleared"
                                If StationControl(Index, index2).TestTimerIsRunning Then
                                    ' a job is open; write to Job Alarms Log
                                    ALM_Write CInt(Index), 1, tempstr
                                Else
                                    ' no open job; write to System Events Log
                                    Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                                End If
                            End If
                        End If
                    ElseIf Alm_LiveFuelSafety(Index, 1) Then
                        ' turn off the safety sw alarm
                        Alm_LiveFuelSafety(Index, 1) = False
                        tempstr = "Live Fuel Tank #" + Format(STN_INFO(Index).ADF_StnNum, "0") + " SafetyLevel Alarm NA (not Load or Purge)"
                        If StationControl(Index, index2).TestTimerIsRunning Then
                            ' open job; write to Job Alarms Log
                            ALM_Write CInt(Index), 1, tempstr
                        Else
                            ' no open job; write to System Events Log
                            Write_ELog "Station #" & Format(Index, "#0") & " " & tempstr
                        End If
                    Else
                        ' no heating; safety switch doesn't matter
                        Alm_LiveFuelSafety(Index, 1) = False
                    End If
                    
                Else
                    Alm_LiveFuelHeater(Index, 1) = False
                    Alm_LiveFuelSafety(Index, 1) = False
                End If
                                
            Else
                Alm_LiveFuelLevel(Index, 1) = False
                Alm_LiveFuelHeater(Index, 1) = False
                Alm_LiveFuelSafety(Index, 1) = False
            End If
        
        Next Index
        
        
bypassadf:
    ' GoTo bypass2
        
         
        For Index = 1 To LAST_STN
            For index2 = 1 To NR_SHIFT
                If USINGLOADPRESSURE Then
                    If StationControl(Index, index2).Mode = VBLOAD Then
                        If Stn_AIO(Index, asLoadPressure).EUValue > StationConfig(Index, index2).LoadPressure Then
                            If Alm_LoadPress(Index, index2) = False Then
                                Alm_LoadPress(Index, index2) = True
                                ALM_Write Index, index2, "Load Pressure Alarm"
                            End If
                        Else
                            If Alm_LoadPress(Index, index2) Then
                                Alm_LoadPress(Index, index2) = False
                                ALM_Write Index, index2, "Load Pressure Alarm Cleared"
                            End If
                        End If
                    Else
                        If Alm_LoadPress(Index, index2) Then
                            Alm_LoadPress(Index, index2) = False
                            ALM_Write Index, index2, "Load Pressure Alarm Cleared"
                        End If
                    End If
                Else
                    Alm_LoadPress(Index, index2) = False
                End If
            Next index2
        Next Index
        
        
bypass2:
        
        ' COMMON ALARMS/ ALARM HORN
        If Alm_Estop Or Alm_Btn20 Or Alm_Flow Or Alm_N2 Or Alm_Doors _
                Or Alm_ExtContacts Then
            If Pause_Alarm = NOTPAUSED Then
                Pause_Alarm = SYSTEMPAUSED
                Alarm_Pause
                If horndisabled = 0 Then
                    Com_OutDigital icAlarmHorn, cON
                End If
            End If
        Else
            If Pause_Alarm = SYSTEMPAUSED Then
                Pause_Alarm = NOTPAUSED
                For Index = 1 To LAST_STN
                    For index2 = 1 To NR_SHIFT
                        If StationControl(Index, index2).Mode = VBPAUSEALARM _
                            And StationControl(Index, index2).Mode_PauseSave = VBIDLE _
                            And AdfControl(Index).Mode = 0 Then
                            ' Automatically Continue Stations that were Idle
                            Alarm_Continue Index, index2
                        End If
                    Next index2
                Next Index
            End If
            If Not IOForceActive Then Com_OutDigital icAlarmHorn, cOFF
            horndisabled = 0
        End If
        
        
        
        ' STATION ALARMS
        SomebodyAlarmOOT = 0
        For Index = 1 To LAST_STN
            For index2 = 1 To NR_SHIFT
                If (Alm_Vac(STN_INFO(Index).AspiratorNum) And Not StationControl(Index, index2).ModeIsIdle_Debounced) _
                   Or Alm_LoadPress(Index, index2) _
                   Or Alm_LiveFuelLevel(Index, index2) _
                   Or Alm_LiveFuelSafety(Index, index2) _
                   Or Alm_LiveFuelHeater(Index, index2) _
                   Or (Alm_SystemVacSw And (StationControl(Index, index2).Mode = VBPURGEWAIT)) Then
                   
                   If (Alm_SystemVacSw And (StationControl(Index, index2).Mode = VBPURGEWAIT)) Then
                        If horndisabled = 0 Then
                            Com_OutDigital icAlarmHorn, cON
                        End If
                   End If
                   Alarm_PauseStation Index, index2
                    
                End If
            Next index2
        Next Index
                
        ' ALARM BEACON
        allIdleFlag = True
        SomebodyAlarmOOT = NOTPAUSED
        For Index = 1 To LAST_STN
            For index2 = 1 To NR_SHIFT
                If StationControl(Index, index2).Mode <> VBIDLE Then allIdleFlag = False
                If StationControl(Index, index2).Mode = VBPAUSEOOT _
                   Or StationControl(Index, index2).Mode = VBPAUSEALARM _
                   Or StationControl(Index, index2).Mode = VBPAUSEVACSW _
                   Or StationControl(Index, index2).Mode = VBLEAKERROR _
                   Or StationControl(Index, index2).Mode = VBCOURSEWAIT Then
                   
                    SomebodyAlarmOOT = SomebodyAlarmOOT + 1
                    
                End If
            Next index2
        Next Index
        If LogTempRh Then
            If (Not USINGPASLOCALCONTROL) Then
                If Not PAS_INFO(pasTEMPERATURE).Ok Then SomebodyAlarmOOT = SomebodyAlarmOOT + 1
                If Not PAS_INFO(pasMOISTURE).Ok Then SomebodyAlarmOOT = SomebodyAlarmOOT + 1
            End If
        End If
        If ((SomebodyAlarmOOT <> NOTPAUSED) Or (Pause_Alarm <> NOTPAUSED)) Then
            Com_OutDigital icAlarmBeacon, cON
        Else
    '        If Not IOForceActive Then Com_OutDigital icAlarmBeacon, cOFF
            Com_OutDigital icAlarmBeacon, cOFF
        End If
        
    End If
    
bypass:

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

Sub OOT_Continue(station As Integer, Shift As Integer)
'
'  Only get here from some station in OOT
'
'
'******************************************************************************


    If (Pause_Alarm = SYSTEMPAUSED) Then Exit Sub                ' There is a system wide pause


    If USINGOOTPAUSE Then

            ' RESET oot COUNT and condition
            FirstTime(station, Shift) = False
            OOTs(station, Shift).BtnFlowOOTCnt = 0
            OOTs(station, Shift).NitFlowOOTCnt = 0
            OOTs(station, Shift).PurFlowOOTCnt = 0
            OOTs(station, Shift).FuelTempOOTCnt = 0
            OOTs(station, Shift).AirTempOOTCnt = 0
            OOTs(station, Shift).AirMoistOOTCnt = 0
            OOTs(station, Shift).CanVentOOTCnt = 0
            OOTs(station, Shift).LoadRateOOTCnt = 0
            OOTs(station, Shift).PurgeDpOOTCnt = 0
            OOTs(station, Shift).FuelLevelOOTCnt = 0
            OOTs(station, Shift).StorageLevelOOTCnt = 0
            AdfControl(station).LiveFuelDensityDeadCnt = 0
            AdfControl(station).LiveFuelDensityWeakCnt = 0
            OOTs(station, Shift).BtnFlowOOT = False
            OOTs(station, Shift).NitFlowOOT = False
            OOTs(station, Shift).PurFlowOOT = False
            OOTs(station, Shift).FuelTempOOT = False
            OOTs(station, Shift).AirTempOOT = False
            OOTs(station, Shift).AirMoistOOT = False
            OOTs(station, Shift).CanVentOOT = False
            OOTs(station, Shift).LoadRateOOT = False
            OOTs(station, Shift).PurgeDpOOT = False
            OOTs(station, Shift).FuelLevelOOT = False
            OOTs(station, Shift).StorageLevelOOT = False
 
 
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
                Case seqLeakTest
                    ' ************************
                    ' LeakTest Sequence
                    ' ************************
                    Select Case SEQ_Step(station, Shift)
                        Case 9
                            ' Successful Completion; Reset Sequence Number
                            SEQ_Nmbr(station, Shift) = seqIdle
                            SEQ_Step(station, Shift) = 0
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
            
            StationControl(station, Shift).OotCurrent = ootNone
            StationControl(station, Shift).PausedDts = 0
            StationControl(station, Shift).OotResponse = ootrspUndefined
        
            OOT_Write_Data station, Shift, StationControl(station, Shift).Mode_PauseSave, OOTPAUSECLEAR      ' OOT Cleared by Operator
            ALM_Write station, Shift, "Operator Continued Station after OOT"
            Write_ELog "Operator Continued, Stn " & station & " after OOT"
     
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
  
    End If                                          ' Not Using OOT Pause

End Sub

Sub OOT_Pause(ByVal station As Integer, ByVal Shift As Integer, ByVal iOOT As Integer, ByVal respOOT As Integer)
'
'  Pause a station with OOT error(s)
'
      
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 3762
      
    ' start pause time
    StationControl(station, Shift).OotCurrent = iOOT
    StationControl(station, Shift).PausedDts = Now
    StationControl(station, Shift).OotResponse = respOOT
    OOT_Write_Data station, Shift, StationControl(station, Shift).Mode, OOTPAUSEBEGIN
    ALM_Write station, Shift, "Station  " & station & " Paused for OOT Condition"
    
    ' Save the current mode for the continue button
    StationControl(station, Shift).Mode_PauseSave = StationControl(station, Shift).Mode
    ' save elapsed hours so far
    Select Case StationControl(station, Shift).Mode
        Case VBLEAK
            LeakCheckControl.ElapsedHours_Prev = LeakCheckControl.ElapsedHours
        Case VBLOAD
            LoadControl(station, Shift).ElapsedHours_Prev = LoadControl(station, Shift).ElapsedHours
        Case VBPURGE
            PurgeControl(station, Shift).ElapsedHours_Prev = PurgeControl(station, Shift).ElapsedHours
        Case Else
    End Select
    ' set mode to OOT Paused
    StationControl(station, Shift).Mode = VBPAUSEOOT
          
    If StationControl(station, Shift).Mode_PauseSave = VBLOAD Then           ' station was loading before OOT
        If StationRecipe(station, Shift).Load_Method = LOADBYTIME Or StationRecipe(station, Shift).Load_Method = LOADBYWC Then
            StationControl(station, Shift).PauseAlarmStartTime = Now                  ' save pause time on load by time
        End If
    End If
   
    If StationControl(station, Shift).Mode_PauseSave = VBPURGE Then          ' station was purging before OOT
        StationControl(station, Shift).PauseAlarmStartTime = Now                      ' save pause time
    End If
    
   
    '  Turn Off Station Related Outputs
   
    '   Station MFCs
    ShutdownStnMFCs station, Shift
    '   Station Valves
    Close_Stn_Valves station, Shift
    '   Scale Valves
    If StationRecipe(station, Shift).UsePriScale And StationControl(station, Shift).PriScaleStn > 0 _
            And StationControl(station, Shift).PriScaleStn < FIRST_REMOTESCALE Then
        Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF
    End If
    If StationRecipe(station, Shift).PurgeAuxCan And StationControl(station, Shift).AuxScaleStn > 0 Then
        Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxPurgeSol, cOFF
    End If
    ' Release Common (Leak) Pressure Transducer (if this station is using it)
    If LeakCheckControl.station = station Then
        LeakCheckControl.station = 0
        LeakCheckControl.Shift = 0
        LeakCheckControl.Phase = 0
        LeakCheckControl.ElapsedHours = 0
        LeakCheckControl.ElapsedHours_Prev = 0
    End If
    
'    Delay_Box "Station " & station & " Shift " & Shift & " Paused for OOT", MSGDELAY, msgSHOW

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

Sub Alarm_Pause()
'
' Significant Problem; Save status as required & then shut-off valves
'
'   rewritten 4 Mar 2005
'
'******************************************************************************

' Must have seen a very significent error to get here
' Build one table set bit for valves on

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 31
Dim cnt1 As Integer
Dim cnt2 As Integer
Dim stnCounter As Integer
Dim pauseMsg As String

    stnCounter = 0
    pauseMsg = "unknown pause"
    For cnt1 = 1 To LAST_STN                                    ' Save all stations doing an activity
        For cnt2 = 1 To NR_SHIFT                                ' and shifts
        
            If (Not StationControl(cnt1, cnt2).IsPausedInAlarm) Then
                 
                If (StationControl(cnt1, cnt2).Mode = VBLEAK Or StationControl(cnt1, cnt2).Mode = VBPURGE Or StationControl(cnt1, cnt2).Mode = VBLOAD) Then       ' only states in which valves used
                     
                    ' Pause LiveFuel ADF Sequence; if active
                    If AdfControl(cnt1).Step <> 0 Then
                        AdfControl(cnt1).StepBeforePause = AdfControl(cnt1).Step
                        AdfControl(cnt1).Step = 95
                    End If
                    ' Pause LiveFuel FST Sequence; if active
                    If FstControl(cnt1).Step <> 0 Then
                        FstControl(cnt1).StepBeforePause = FstControl(cnt1).Step
                        FstControl(cnt1).Step = 95
                    End If
                           
                    ' running stations counter
                    stnCounter = stnCounter + 1
                    
                    '  Turn Off Station Related Outputs
                    '   Station MFCs
                    ShutdownStnMFCs cnt1, cnt2
                    '   Station Valves
                    Close_Stn_Valves cnt1, cnt2
                    '   Scale Valves
                    If StationRecipe(cnt1, cnt2).UsePriScale And StationControl(cnt1, cnt2).PriScaleStn > 0 _
                            And StationControl(cnt1, cnt2).PriScaleStn < FIRST_REMOTESCALE Then
                        Stn_OutDigital StationControl(cnt1, cnt2).PriScaleStn, isPriAuxVentSol, cOFF
                    End If
                    If StationRecipe(cnt1, cnt2).PurgeAuxCan And StationControl(cnt1, cnt2).AuxScaleStn > 0 Then
                        Stn_OutDigital StationControl(cnt1, cnt2).AuxScaleStn, isAuxPurgeSol, cOFF
                    End If
                    ' reset Leakcheck control
                    If LeakCheckControl.station = cnt1 Then
                       LeakCheckControl.station = 0
                       LeakCheckControl.Shift = 0
                       LeakCheckControl.Phase = 0
                       LeakCheckControl.ElapsedHours = 0
                       LeakCheckControl.ElapsedHours_Prev = 0
                    End If
    
                End If                                          ' Leak, Load or Purge
                 
                ' Save the current mode for the continue button (unless station was already paused)
                If StationControl(cnt1, cnt2).Mode <> VBPAUSEALARM And StationControl(cnt1, cnt2).Mode <> VBPAUSEOOT Then
                   StationControl(cnt1, cnt2).Mode_PauseSave = StationControl(cnt1, cnt2).Mode
                End If
                ' set mode to Paused
                StationControl(cnt1, cnt2).Mode = VBPAUSEALARM
                ' start pause time
                StationControl(cnt1, cnt2).PauseAlarmStartTime = Now
                StationControl(cnt1, cnt2).IsPausedInAlarm = True
        
                ' Station Paused Message
                If Alm_Estop Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for ESTOP"
                ElseIf Alm_Btn20 Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for Btn20 Alarm"
                ElseIf Alm_Flow Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for No Flow Switch"
                ElseIf Alm_Doors Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for Doors Open Too Long"
                ElseIf Alm_N2 Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for No N2 Press Switch"
                ElseIf Alm_ExtContacts Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for " & DESC_EXT_CONTACTS
                ElseIf Alm_UPS Then
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused for UPS"
                Else
                   StationControl(cnt1, cnt2).PauseMessage = "All Stations Paused"
                End If
                If (pauseMsg = "unknown pause") Then pauseMsg = StationControl(cnt1, cnt2).PauseMessage
    
                ' Write to Alarm Log
                If (StationControl(cnt1, cnt2).Mode_PauseSave <> VBIDLE) Then
                   ALM_Write cnt1, cnt2, StationControl(cnt1, cnt2).PauseMessage
                Else
                   Write_ELog "Station #" & Format(cnt1, "#0") & " Shift #" & Format(cnt2, "#0") & " reports - " & StationControl(cnt1, cnt2).PauseMessage
            
                End If
            End If
                
        Next cnt2
    Next cnt1
             
    ' Write to Event Log
    Write_ELog pauseMsg
             
    ' Anybody running ??
    If stnCounter > 0 Then
        '  turn off common valves
        Close_Main_Valves
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

Sub Alarm_PauseStation(station As Integer, Shift As Integer)
'
'   Significant Problem; Save status as required & then shut-off valves
'
'   written 10 Jan 2006
'
'******************************************************************************
' Must have seen a very significent error to get here

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 1331
Dim tempstr As String
   
    If Not StationControl(station, Shift).IsPausedInAlarm Then
    
        ' Save the current mode for the continue button
        StationControl(station, Shift).Mode_PauseSave = StationControl(station, Shift).Mode
        ' save elapsed hours so far
        Select Case StationControl(station, Shift).Mode
            Case VBLEAK
                LeakCheckControl.ElapsedHours_Prev = LeakCheckControl.ElapsedHours
            Case VBLOAD
                LoadControl(station, Shift).ElapsedHours_Prev = LoadControl(station, Shift).ElapsedHours
            Case VBPURGE
                PurgeControl(station, Shift).ElapsedHours_Prev = PurgeControl(station, Shift).ElapsedHours
            Case Else
        End Select
        ' start pause time
        StationControl(station, Shift).PauseAlarmStartTime = Now
        StationControl(station, Shift).IsPausedInAlarm = True         ' remember that station was paused for alarm
        ' Which Station Alarm ?
        If Alm_LiveFuelLevel(station, Shift) Then
            tempstr = "Paused for Live Fuel Level Alarm"
            ' set mode to Paused for Alarm
            StationControl(station, Shift).Mode = VBPAUSEALARM
        ElseIf Alm_LiveFuelSafety(station, Shift) Then
            tempstr = "Paused for Live Fuel Safety Level Alarm"
            ' set mode to Paused for Alarm
            StationControl(station, Shift).Mode = VBPAUSEALARM
        ElseIf Alm_LiveFuelHeater(station, Shift) Then
            tempstr = "Paused for Live Fuel Heater OverTemp"
            ' set mode to Paused for Alarm
            StationControl(station, Shift).Mode = VBPAUSEALARM
        ElseIf (Alm_Vac(STN_INFO(station).AspiratorNum) And Not StationControl(station, Shift).ModeIsIdle_Debounced) Then
            tempstr = "Paused for Vacuum Switch Alarm"
            ' set mode to Paused for Alarm
            StationControl(station, Shift).Mode = VBPAUSEALARM
        ElseIf Alm_LoadPress(station, Shift) Then
            tempstr = "Paused for Load Pressure Alarm"
            ' set mode to Paused for Alarm
            StationControl(station, Shift).Mode = VBPAUSEALARM
        ElseIf (Alm_SystemVacSw And (StationControl(station, Shift).Mode = VBPURGEWAIT)) Then
            tempstr = "Paused for System Vacuum Switch Alarm"
            ' set mode to Paused for SystemVacuum Switch Off
            StationControl(station, Shift).Mode = VBPAUSEVACSW
        Else
            tempstr = "Paused"
        End If
        ' Station Paused Message
        StationControl(station, Shift).PauseMessage = "Station " + tempstr
        ' Write to Logs
        If StationControl(station, Shift).TestTimerIsRunning Then ALM_Write station, Shift, tempstr
        Write_ELog "Station #" + Format(station, "0") + " Shift #" + Format(Shift, "0") + " " + tempstr & " @ " & Format(Timer, "###,##0.000")
        
      
        ' Turn Off Live Fuel Vapor Generator Tank Valves & Pump
        If AdfControl(station).Step <> 0 Then
            AdfControl(station).StepBeforePause = AdfControl(station).Step
            AdfControl(station).Step = 95         ' Pause ADF Sequence & Turn Off Valves & Pump
        End If
        ' Turn Off Live Fuel Storage Tank Valves & Pump
        If FstControl(station).Step <> 0 Then
            FstControl(station).StepBeforePause = FstControl(station).Step
            FstControl(station).Step = 95         ' Pause FST Sequence & Turn Off Valves & Pump
        End If
    
                 
        '  Turn Off Station Related Outputs
        
        '   Station Valves
        Close_Stn_Valves station, Shift
        '   Scale Valves
        If StationRecipe(station, Shift).UsePriScale And StationControl(station, Shift).PriScaleStn > 0 _
                And StationControl(station, Shift).PriScaleStn < FIRST_REMOTESCALE Then
            Stn_OutDigital StationControl(station, Shift).PriScaleStn, isPriAuxVentSol, cOFF
        End If
        If StationRecipe(station, Shift).PurgeAuxCan And StationControl(station, Shift).AuxScaleStn > 0 Then
            Stn_OutDigital StationControl(station, Shift).AuxScaleStn, isAuxPurgeSol, cOFF
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

Sub Alarm_Continue(station As Integer, Shift As Integer)
'
' Continue after Significant Problem; reset valves as required
'
'
'***********************************************************************************

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 32

    '
    ' Operator pushed continue on station screen
    ' so we will restart that station
    '*******************************************************************

    Pause_Alarm = NOTPAUSED
    
    If station > 0 And station <= LAST_STN And Shift > 0 And Shift <= NR_SHIFT Then
    
        
        If StationControl(station, Shift).IsPausedInAlarm Then                          ' station was paused for alarm
            StationControl(station, Shift).IsPausedInAlarm = False                             ' reset station alarm indicator
            ' Station Paused Message
            StationControl(station, Shift).PauseMessage = ""
        End If
        
        ' LiveFuel Vapor Generator Tank AutoDrainFill
        If AdfControl(station).Mode <> 0 Then
            Select Case AdfControl(station).Mode
                Case 1      ' ADF Mode 1 is Drain Only
                    ' Restart Drain Sequence
                    AdfControl(station).StepBeforePause = 0
                    AdfControl(station).Step = 0
                Case 2      ' ADF Mode 2 is Drain and Fill (and maybe Temp Control)
                    Select Case AdfControl(station).StepBeforePause
                        Case 0 To 19
                            ' Restart Drain Sequence
                            AdfControl(station).Step = 0
                            AdfControl(station).StepBeforePause = 0
                        Case 21 To 38
                            ' Restart Fill Sequence
                            AdfControl(station).Step = 21
                            AdfControl(station).StepBeforePause = 0
                        Case 39
                            ' Restart Heat to Temp
                            AdfControl(station).Step_Time = Now() + TimeSerial(0, StationCfg_ADF(station, 1).HeaterTimeout, 0)
                            AdfControl(station).Step = AdfControl(station).StepBeforePause
                            AdfControl(station).StepBeforePause = 0
                        Case 49
                            ' Drain/Fill Sequence is Complete; Resume
                            AdfControl(station).Step = AdfControl(station).StepBeforePause
                            AdfControl(station).StepBeforePause = 0
                        Case 90 To 94
                            ' Resume "Aborted"
                            AdfControl(station).Step = AdfControl(station).StepBeforePause
                            AdfControl(station).StepBeforePause = 0
                        Case 96
                            ' Resume "Aborted"
                            AdfControl(station).Step = AdfControl(station).StepBeforePause
                            AdfControl(station).StepBeforePause = 0
                        Case 99 To 109
                            ' Restart Temp Control
                            AdfControl(station).Step = 39
                            AdfControl(station).StepBeforePause = 0
                        Case Else
                            ' Should never get here; Reset Everything
                            Write_ELog "Can't Resume ADF @ Mode" & AdfControl(station).Mode & " Step" & AdfControl(station).Step & " for " & station & " Station " & "  " & Shift & " Shift"
                            AdfControl(station).Mode = 0
                            AdfControl(station).Step = 0
                            AdfControl(station).StepBeforePause = 0
                    End Select
            End Select
        End If
                       
        ' LiveFuel Storage Tank AutoDrainFill
        If FstControl(station).Mode <> 0 Then
            Select Case FstControl(station).Mode
                Case 1      ' FST Mode 1 is Drain
                    ' Restart Drain Sequence
                    FstControl(station).StepBeforePause = 0
                    FstControl(station).Step = 0
                Case 2      ' FST Mode 2 is Fill
                    Select Case FstControl(station).StepBeforePause
                        Case 21 To 29
                            ' Restart Fill Sequence
                            FstControl(station).Step = 21
                            FstControl(station).StepBeforePause = 0
                        Case 90 To 94
                            ' Resume "Aborted"
                            FstControl(station).Step = FstControl(station).StepBeforePause
                            FstControl(station).StepBeforePause = 0
                        Case Else
                            ' Should never get here; Reset Everything
                            Write_ELog "Can't Resume FST @ Mode" & FstControl(station).Mode & " Step" & FstControl(station).Step & " for " & station & " Station " & "  " & Shift & " Shift"
                            FstControl(station).Mode = 0
                            FstControl(station).Step = 0
                            FstControl(station).StepBeforePause = 0
                    End Select
            End Select
        End If
                       
        ' Station Control
        Select Case StationControl(station, Shift).Mode_PauseSave
        
            Case VBLEAK
                ' Restart Leak Check
                LeakCheck_Start station, Shift
                
            Case VBLOAD
                ' Resume Load
                Load_Continue station, Shift
                
            Case VBPRELOAD
                ' Restart PreLoad N2 Push
                PreLoad_Start station, Shift
                
            Case VBPURGE
                ' Resume Purge
                Purge_Continue station, Shift
                       
            Case VBPOSTPURGE
                ' Reset Station Mode
                StationControl(station, Shift).Mode = StationControl(station, Shift).Mode_PauseSave
                ' Is a Sequence Running ?
                Select Case SEQ_Nmbr(station, Shift)
                    Case seqCanVentN2Feed
                        ' Post Purge N2 Feed Sequence
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
                        ' Restart Post Purge Delay
                        StationControl(station, Shift).End_Time = Now() + TimeSerial(0, StationRecipe(station, Shift).PausePurgeTime, 0)
                End Select
          
            
            Case Else
                ' Reset Station Mode
                StationControl(station, Shift).Mode = StationControl(station, Shift).Mode_PauseSave
                
        End Select
                   
    
    End If                                                                  ' station > 0
        
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

Sub DoorOpened()
'
'   The door is ajar and we will pause all the tests
'   when we find the set time equal to the current time
'
Dim iStn As Integer
Dim istn2 As Integer
  
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 34
 
    If Com_DIO(icDoorSw).Value Then
        Alm_Doors = False
        Alm_Doors_Time = 0
        Alm_Doors_Count = 0
        Alm_Doors_FirstTime = False
    Else
        If Alm_Doors_FirstTime Then
            If Not Alm_Doors And Now > Alm_Doors_Time + TimeSerial(0, Alm_Doors_Count, 0) Then
'                Delay_Box "Door Open...System will Pause All Tests in " & StationConfig(Index, index2).DoorOpenDelay - Alm_Doors_Count & " minutes", MSGDELAY, msgSHOW
                Alm_Doors_Count = Alm_Doors_Count + 1
                If Now > Alm_Doors_Time + TimeSerial(0, SysConfig.DoorOpenDelay, 0) Then
                    Write_ELog "CPS Pausing All Tests..Door Opened too Long"
                    Alm_Doors = True
                End If
            End If
        Else        ' setup for first time through
            Alm_Doors = False
            Alm_Doors_FirstTime = True
            Alm_Doors_Time = Now
            Alm_Doors_Count = 1
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

Sub ExtAlarmInput()
'
'   Customer provided two inputs that will stop all current running processes
'
'   Mainly used so far at TOYOTA
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 35
Static ExtAlarmFirstTime As Boolean                     ' Used to make sure it's not a bouncing contact
Static ExtAlarmTime As Date
  

    If USING_EXT_CONTACTS Then
        If Not Com_DIO(icExtAlmContactSw).Value Then
            Alm_ExtContacts = False
            ExtAlarmFirstTime = False
        Else
            If ExtAlarmFirstTime Then
                If Not Alm_ExtContacts Then
                    Alm_ExtContacts = True
                    ' display error and don't return without OK
                    FrmCustomerContacts.Show
                End If
            Else
                ' Yes it's still set
                ExtAlarmFirstTime = True
            End If
        End If
    Else
        Alm_ExtContacts = False
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

Sub SystemVacSw()
'
'   System Vacuum Switch; can't Purge without it
'
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 3556
Static SystemVacSwOnCount As Integer                  ' Used to make sure it's not a bouncing contact
Static SystemVacSwOffCount As Integer                 ' Used to make sure it's not a bouncing contact
Dim iStn As Integer
Dim iShift As Integer
Dim sMsg As String
Dim logFlag As Boolean

    If USINGSYSTEMVACSW Then
        ' is System Vacuum Switch True ??
        If Com_DIO(icSystemVacSw).Value Then
            ' switch is true
            SystemVacSwOffCount = 0
            If (SystemVacSwOnCount < 999) Then SystemVacSwOnCount = SystemVacSwOnCount + 1
            ' switch count > max ??
            If (SystemVacSwOnCount > maxSystemVacSwCount) Then
                ' still in alarm ??
                If Alm_SystemVacSw Then
                    ' turn off alarm
                    Alm_SystemVacSw = False
                    ' update logs
                    sMsg = "System Vacuum Switch is now On"
                    logFlag = True
                Else
                    ' nothing to update
                    logFlag = False
                End If
            End If
        Else
            ' switch is false
            SystemVacSwOnCount = 0
            If (SystemVacSwOffCount < 999) Then SystemVacSwOffCount = SystemVacSwOffCount + 1
            ' switch count > max ??
            If (SystemVacSwOffCount > maxSystemVacSwCount) Then
                ' not in alarm yet ??
                If Not Alm_SystemVacSw Then
                    ' turn on alarm
                    Alm_SystemVacSw = True
                    ' update logs
                    sMsg = "System Vacuum Switch is now OFF"
                    logFlag = True
                Else
                    ' nothing to update
                    logFlag = False
                End If
            End If
        End If
        ' update logs
        If logFlag Then
            ' update system events log
            Write_ELog sMsg
            ' update any (open)job event logs
            For iStn = 1 To LAST_STN
                For iShift = 1 To NR_SHIFT
                    If (Len(StationControl(iStn, iShift).DBFile) > 3) Then Write_JLog iStn, iShift, sMsg
                Next iShift
            Next iStn
        End If
    Else
        ' alarm is OFF
        Alm_SystemVacSw = False
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

Sub UpsTest()
'
'   The system is protected by an UPS system
'
'   If it is down too long do an orderly shut down
'
'   Display a one minute message any time UPS is active
'
'   UPS type 1 - Run system until timer exceeded, then do an orderly shut down
'   UPS type 2 - Pause system until timer exceeded and then do an orderly shut down
'
'   If power restored....
'     UPS type 1 Continue as usual
'     UPS type 2 Operator must continue each process in progress
'
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 36
 
' Ups-has-low-battery file and/or Ups-has-low-battery input
'If Not fs.FileExists(FILEPATH & "\UpsAlarmOn.txt") Then
'If Not Com_DIO(icUpsActiveSw).Value Then   ' false is good; we are not in UPS mode
If Not fs.FileExists(filepath & "\UpsAlarmOn.txt") And Not Com_DIO(icUpsActiveSw).Value Then
    Alm_UPS = False
    If Alm_Ups_FirstTime Then   ' UPS has just resumed AC-Input Operation
        Write_ELog "Ups is Running on AC Input now."
'        Delay_Box "UPS on AC Input: System shut down cancelled.", MSGDELAY, msgSHOW
    End If
    Alm_Ups_Time = 0
    Alm_Ups_Count = 0
    Alm_Ups_FirstTime = False
Else
    Alm_UPS = True
    If USINGUPS = 2 Then                        ' Small UPS => shut down now
        If Pause_Alarm = NOTPAUSED Then         ' not in alarm yet
            Alarm_Pause                         ' We have a very bad power condition => Pause the System
        End If
    End If
    If Alm_Ups_FirstTime Then   ' set up first timers,counters ect.
        If Now > Alm_Ups_Time + TimeSerial(0, Alm_Ups_Count, 0) Then
            Alm_Ups_Count = Alm_Ups_Count + 1
            If Now > Alm_Ups_Time + TimeSerial(0, SysConfig.UPSOpenDelay, 0) Then
                Alm_UPS = True
                Down_Now
                'Orderly shut down now QUIT.
            End If
        End If
    Else        ' setup for first time through
        Write_ELog "Ups is Running on Batteries now."
        Alm_Ups_FirstTime = True
        Alm_Ups_Time = Now
        Alm_Ups_Count = 1
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

Sub Down_Now()

Dim iStn As Integer
Dim iShift As Integer

' orderly shutdown time UPS fault
frmMainMenu.MousePointer = vbHourglass
' save current system condition (esp. butane remaining)
Save_ButaneSupply
Write_ELog "CPS Reporting System Shutting Down..Ups Battery Time Limit"
Reset_Valves
' Close all stations currently running
For iStn = 1 To LAST_STN
    For iShift = 1 To NR_SHIFT
        If (StationControl(iStn, iShift).Mode <> VBIDLE And StationControl(iStn, iShift).Mode <> VBIDLEWAITING) Then
            ALM_Write iStn, iShift, "UPS Time Limit Exceeded / Data Collection Halted."
            Stats_Write iStn, iShift
            StationControl(iStn, iShift).End_Time = Now
            Delay_Box "Stopping Station " & Format(iStn, "0") & " Shift " & Format(iShift, "0"), MSGDELAY, msgSHOW
            StnRemoteTask(iStn, iShift).PreviousResult = "CPS Shutting Down"
            DoEvents
            Station_Finish iStn, iShift
        End If
        If iStn <= NR_SCALES Then frmComm8Card.Close_Scale iStn
    Next iShift
Next iStn

frmMainMenu.MousePointer = vbDefault
' Reset Watchdog for Main Board Module 0
Opto_Send_Data(0) = val(200)                                        ' set Watchdog Time to 2 sec (200 * 10msec)
If IoComOn Then frmMainForm.Send_Opto_Command 0, 114, 0, 65535      ' All Off; including Beacon
' Done; End Program
End                 'Orderly shut down now QUIT.

End Sub

Sub CustomerLowGas()
'
'   Customer has a contact in the BUTANE room that when closed
'   means the gas bottle is low.  Just a warning.... No shut down
'
'
'
Static CustomerLowGasFirstTime As Boolean
Static CustomerLowGasInc As Integer
Static CustomerLowGasTime As Date
Dim iStn As Integer
Dim istn2 As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 37
 
    If Not Com_DIO(icCustLowGasSw).Value Then   'reset
        CustomerLowGasTime = 0
        CustomerLowGasInc = 0
        CustomerLowGasFirstTime = False
    Else
        If CustomerLowGasFirstTime Then     'set
            If Now > CustomerLowGasTime + TimeSerial(0, CustomerLowGasInc, 0) Then
                Delay_Box "Low Gas Pressure Detected", MSGDELAY, msgSHOW
                CustomerLowGasInc = CustomerLowGasInc + 1
            End If
        Else        ' setup for first time through
            CustomerLowGasFirstTime = True
            CustomerLowGasTime = Now
            CustomerLowGasInc = 1
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

Sub LowButaneGas()
'
'       Check for low butane
'
Dim PercentRemaining As Single
Dim LowButane As Boolean

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 13, 38

    If ((ButaneSupply.CurrentOnHand = Empty) Or (ButaneSupply.CylinderWeight = Empty)) Then
        ' power up AND no set values
        ButaneSupply.CurrentOnHand = 1
        ButaneSupply.CylinderWeight = 1
        ButaneSupply.WarningSetPoint = 10
        Save_ButaneSupply
    End If
    PercentRemaining = 100 * (ButaneSupply.CurrentOnHand / ((ButaneSupply.CylinderWeight * 28.317) / 0.1501))
    LowButane = IIf(PercentRemaining < ButaneSupply.WarningSetPoint, True, False)
    If LowButane And Not ButaneSupply.WarningActive Then Write_ELog "Low Butane Warning."
    If Not LowButane And ButaneSupply.WarningActive Then Write_ELog "Reset Low Butane Warning."
    ButaneSupply.WarningActive = IIf(LowButane, True, False)
    
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


