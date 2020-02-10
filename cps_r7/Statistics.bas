Attribute VB_Name = "Module12"
'error module 12 ''''''''''''program STATISTICS.bas ''''''''''''''''''''''''
Option Explicit
' Set Temp Statistic Values
'
Type TempStat
  LastVal(1 To 4) As Single
  FirstTime As Boolean
End Type

Global PurStat(1 To MAX_STN, 1 To MAX_SHIFT) As TempStat
Global BtnStat(1 To MAX_STN, 1 To MAX_SHIFT) As TempStat
Global NitStat(1 To MAX_STN, 1 To MAX_SHIFT) As TempStat
Global MixStat(1 To MAX_STN, 1 To MAX_SHIFT) As TempStat

Sub Stats_Write(Index As Integer, index2 As Integer)
'
' Module Name:  Stats_Write
' Author:       Analytical Process Programmer 9/96
' Description:  This routine writes the statistical data into the data
'               base file for the selected station.
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 12, 2

Dim Butane_Rate As Single
Dim dbDbase As Database
Dim rsTable As Recordset
Dim Criterion As String

    Select Case StationControl(Index, index2).Mode
      
      Case VBPURGE
        ' Write Statistics for Purge to file
          Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
          Set rsTable = dbDbase.OpenRecordset("Stats")
          rsTable.AddNew
            rsTable("PAFlowMin") = StationStatistics(Index, index2).Pur.sMin
            rsTable("PAFlowAvg") = StationStatistics(Index, index2).Pur.sAvg
            rsTable("PAFlowMax") = StationStatistics(Index, index2).Pur.sMax
            rsTable("PAMoistMin") = StationStatistics(Index, index2).AirMoist.sMin
            rsTable("PAMoistAvg") = StationStatistics(Index, index2).AirMoist.sAvg
            rsTable("PAMoistMax") = StationStatistics(Index, index2).AirMoist.sMax
            rsTable("PATempMin") = StationStatistics(Index, index2).AirTemp.sMin
            rsTable("PATempAvg") = StationStatistics(Index, index2).AirTemp.sAvg
            rsTable("PATempMax") = StationStatistics(Index, index2).AirTemp.sMax
            rsTable("Mode") = StationControl(Index, index2).Mode
            rsTable("ModeDesc") = ModeDescShort(VBPURGE)
            rsTable("Course") = StationControl(Index, index2).Course
            rsTable("Cycle") = StationControl(Index, index2).CurrCycle
            rsTable("PATotal") = PurgeControl(Index, index2).Purge_Total
            rsTable("PATarget") = PurgeControl(Index, index2).Purge_Target
            rsTable("StartTime") = StationControl(Index, index2).Mode_StartDts
            rsTable("EndTime") = Now()
            rsTable("WtChgTotal") = PurgeControl(Index, index2).TotalWtChgAtEOP
            '
          rsTable.Update
          rsTable.Close
          dbDbase.Close
      Case VBLOAD
        ' Save Load Statistics
         Set dbDbase = OpenDatabase(StationControl(Index, index2).DBFile)
         Set rsTable = dbDbase.OpenRecordset("Stats")
         rsTable.AddNew
           rsTable("BtnFlowMin") = StationStatistics(Index, index2).Btn.sMin
           rsTable("BtnFlowAvg") = StationStatistics(Index, index2).Btn.sAvg
           rsTable("BtnFlowMax") = StationStatistics(Index, index2).Btn.sMax
           rsTable("NitFlowMin") = StationStatistics(Index, index2).Nit.sMin
           rsTable("NitFlowAvg") = StationStatistics(Index, index2).Nit.sAvg
           rsTable("NitFlowMax") = StationStatistics(Index, index2).Nit.sMax
           rsTable("MixMin") = StationStatistics(Index, index2).Mix.sMin
           rsTable("MixAvg") = StationStatistics(Index, index2).Mix.sAvg
           rsTable("MixMax") = StationStatistics(Index, index2).Mix.sMax
           rsTable("FuelTempMin") = StationStatistics(Index, index2).FuelTemp.sMin
           rsTable("FuelTempAvg") = StationStatistics(Index, index2).FuelTemp.sAvg
           rsTable("FuelTempMax") = StationStatistics(Index, index2).FuelTemp.sMax
           rsTable("Mode") = StationControl(Index, index2).Mode
           rsTable("ModeDesc") = ModeDescShort(VBLOAD)
           rsTable("Course") = StationControl(Index, index2).Course
           rsTable("Cycle") = StationControl(Index, index2).CurrCycle
    '      For Regular Stations   - "LoadTotal" = Butane Total (in grams)
    '      For Live Fuel Stations - "LoadTotal" = Vapor Carrier Flow (in liters)
           If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And StationRecipe(Index, index2).LiveFuel) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And StationRecipe(Index, index2).LiveFuel)) Then
                rsTable("LiveFuelCycles") = StationControl(Index, 1).LiveFuelCycleCount
           Else
                rsTable("LiveFuelCycles") = 0
           End If
           rsTable("LoadTotalGrams") = LoadControl(Index, index2).loadTotalGrams
           rsTable("LoadTotalLiters") = LoadControl(Index, index2).LoadTotalLiters
           rsTable("WtChgTotal") = LoadControl(Index, index2).TotalWtChgAtEOL
           rsTable("LoadTarget") = LoadControl(Index, index2).LoadTarget
           rsTable("StartTime") = StationControl(Index, index2).Mode_StartDts
           rsTable("EndTime") = Now()
           
           Butane_Rate = CSng(GramsPerHourToSlpm(StationRecipe(Index, index2).Load_Rate, StationControl(Index, index2).BtnDensity))
    '   ******************************
    '   *** FORMULA CHECK REQUIRED ***
    '   ******************************
           If USINGLINEVOLUME And StationControl(Index, index2).CompletedCycles = CInt(0) And Stn_Btn_FlowSP(Index, index2) > 0 Then
                    
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
                rsTable("LineLoss") = VALUE0
           End If
    '   ******************************
    '   *** FORMULA CHECK REQUIRED ***
    '   CanVent???
    '   ******************************
           If USINGLINEVOLUME And StationControl(Index, index2).CompletedCycles > VALUE0 Then
                If StationRecipe(Index, index2).Load_Method = LOADBYBREAKTHRU Or StationRecipe(Index, index2).Load_Method = LOADBYFID Then
                    If Butane_Rate > 0 Then
                        ' losses       =   vent volume / slpm flow
                        '                 *  Load rate / one minute
                        rsTable("LineLoss") = (StationRecipe(Index, index2).VentV / Butane_Rate) _
                                               * (StationRecipe(Index, index2).Load_Rate / 60)
                    Else
                        ' avoid divide by zero
                        rsTable("LineLoss") = 0
                    End If
                End If
           End If
           
           ' NOLOAD = 0             ' No load
           ' LOADBYTIME = 1         ' Load by time
           ' LOADBYWC = 2           ' Load by working capacity
           ' LOADBYWEIGHT = 3       ' Load by weight
           ' LOADBYBREAKTHRU = 4    ' Load by breakthrough
           ' LOADBYFID = 5          ' Load by FID Breakthrough
           rsTable("LoadMethod") = StationRecipe(Index, index2).Load_Method
           
           rsTable.Update
           rsTable.Close
           dbDbase.Close
         
       Case Else
         ' do nothing during pause or idle
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

Sub Clear_Stats(Index As Integer, index2 As Integer, Optional Mode)
'
' Module Name:  Clear_Stats
' Author:       Analytical Process Programmer 9/96
' Description:  This routine resets the statistic counters.
'
'               To prevent this from affecting the statistics by reading
'               a bad value when the mode has changed, the statistics have
'               a 'push-down' delay array where values are not immediately
'               read into the statistics until they are four read cycles
'               old.  Once the mode changes, all of the values are cleared
'               preventing old values from being read in the new mode.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 12, 1

    If IsMissing(Mode) Then Mode = 3
    ' Mode 1 = clear purge
    ' Mode 2 = clear load
    ' Mode 3 = clear all
    If Mode = 2 Or Mode = 3 Then
        StationStatistics(Index, index2).Btn.sMin = 0
        StationStatistics(Index, index2).Btn.sMax = 0
        StationStatistics(Index, index2).Btn.sAvg = 0
        StationStatistics(Index, index2).Btn.sCnt = 0
        StationStatistics(Index, index2).Nit.sMin = 0
        StationStatistics(Index, index2).Nit.sMax = 0
        StationStatistics(Index, index2).Nit.sAvg = 0
        StationStatistics(Index, index2).Nit.sCnt = 0
        StationStatistics(Index, index2).Mix.sMin = 0
        StationStatistics(Index, index2).Mix.sMax = 0
        StationStatistics(Index, index2).Mix.sAvg = 0
        StationStatistics(Index, index2).Mix.sCnt = 0
        StationStatistics(Index, index2).FuelTemp.sMin = 0
        StationStatistics(Index, index2).FuelTemp.sMax = 0
        StationStatistics(Index, index2).FuelTemp.sAvg = 0
        StationStatistics(Index, index2).FuelTemp.sCnt = 0
        BtnStat(Index, index2).LastVal(1) = 0
        BtnStat(Index, index2).LastVal(2) = 0
        BtnStat(Index, index2).LastVal(3) = 0
        BtnStat(Index, index2).LastVal(4) = 0
        BtnStat(Index, index2).FirstTime = True
        NitStat(Index, index2).LastVal(1) = 0
        NitStat(Index, index2).LastVal(2) = 0
        NitStat(Index, index2).LastVal(3) = 0
        NitStat(Index, index2).LastVal(4) = 0
        NitStat(Index, index2).FirstTime = True
        MixStat(Index, index2).LastVal(1) = 0
        MixStat(Index, index2).LastVal(2) = 0
        MixStat(Index, index2).LastVal(3) = 0
        MixStat(Index, index2).LastVal(4) = 0
        MixStat(Index, index2).FirstTime = True
    End If
    If Mode = 1 Or Mode = 3 Then
        PurStat(Index, index2).LastVal(1) = 0
        PurStat(Index, index2).LastVal(2) = 0
        PurStat(Index, index2).LastVal(3) = 0
        PurStat(Index, index2).LastVal(4) = 0
        PurStat(Index, index2).FirstTime = True
        StationStatistics(Index, index2).Pur.sMin = 0
        StationStatistics(Index, index2).Pur.sMax = 0
        StationStatistics(Index, index2).Pur.sAvg = 0
        StationStatistics(Index, index2).Pur.sCnt = 0
        StationStatistics(Index, index2).AirMoist.sMin = 0
        StationStatistics(Index, index2).AirMoist.sMax = 0
        StationStatistics(Index, index2).AirMoist.sAvg = 0
        StationStatistics(Index, index2).AirMoist.sCnt = 0
        StationStatistics(Index, index2).AirTemp.sMin = 0
        StationStatistics(Index, index2).AirTemp.sMax = 0
        StationStatistics(Index, index2).AirTemp.sAvg = 0
        StationStatistics(Index, index2).AirTemp.sCnt = 0
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

Sub Update_Stats(ByVal Index As Integer, index2 As Integer)
' Module Name:  Update_Stats
' Author:       Analytical Process Programmer 9/96
' Description:
'
' Max, Min and Average Values are maintained for each cycle.
' To not read values until MFCs have settled, Statistics are not updated until
' MFCs settle, and then are placed in a shift register where samples used
' to update the Stats are 4 cycles old.
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 12, 3

Dim tvalue, mvalue As Single

    ' SEE IF SETTLING TIME IS COMPLETE FOR CURRENT MODE
    If ((StationControl(Index, index2).Mode_StartDts + TimeSerial(0, 0, MFC_Settle_Time)) < Now) Then
        If (StationControl(Index, index2).Mode = VBPURGE And PurgeControl(Index, index2).Phase = PurgePurging And Not (PurgeControl(Index, index2).InhibitOotCheck)) Then
            ' On first entry, write same value to all variables
            If PurStat(Index, index2).FirstTime Then
                PurStat(Index, index2).LastVal(1) = Stn_AIO(Index, asPurgeAirFlow).EUValue
                PurStat(Index, index2).LastVal(2) = Stn_AIO(Index, asPurgeAirFlow).EUValue
                PurStat(Index, index2).LastVal(3) = Stn_AIO(Index, asPurgeAirFlow).EUValue
                PurStat(Index, index2).LastVal(4) = Stn_AIO(Index, asPurgeAirFlow).EUValue
                PurStat(Index, index2).FirstTime = False
            Else
                ' On subsequent cycles, rotate values through
                ' Use the oldest value to update stats
                ' Oldest value is # 1, most recent value is #4
                tvalue = PurStat(Index, index2).LastVal(1)
                PurStat(Index, index2).LastVal(1) = PurStat(Index, index2).LastVal(2)
                PurStat(Index, index2).LastVal(2) = PurStat(Index, index2).LastVal(3)
                PurStat(Index, index2).LastVal(3) = PurStat(Index, index2).LastVal(4)
                PurStat(Index, index2).LastVal(4) = Stn_AIO(Index, asPurgeAirFlow).EUValue
                With StationStatistics(Index, index2).Pur
                    If .sCnt = 0 Then
                        .sMin = tvalue
                        .sMax = tvalue
                        .sAvg = tvalue
                        .sCnt = .sCnt + 1
                    Else
                        .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                        .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                        .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                        .sCnt = .sCnt + 1
                    End If
                End With
            End If
            tvalue = PAMoisture
            With StationStatistics(Index, index2).AirMoist
                If .sCnt = 0 Then
                    .sMin = tvalue
                    .sMax = tvalue
                    .sAvg = tvalue
                    .sCnt = .sCnt + 1
                Else
                    .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                    .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                    .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                    .sCnt = .sCnt + 1
                End If
            End With
            tvalue = PATemp
            With StationStatistics(Index, index2).AirTemp
                 If .sCnt = 0 Then
                    .sMin = tvalue
                    .sMax = tvalue
                    .sAvg = tvalue
                    .sCnt = .sCnt + 1
                Else
                    .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                    .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                    .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                    .sCnt = .sCnt + 1
                End If
            End With
        Else
            If StationControl(Index, index2).Mode = VBLOAD And LoadControl(Index, index2).Phase = LoadLoading Then
                ' On first entry, write same value to all variables
                If BtnStat(Index, index2).FirstTime Then
                    BtnStat(Index, index2).LastVal(1) = Stn_Btn_Flow_PV(Index, index2)
                    BtnStat(Index, index2).LastVal(2) = Stn_Btn_Flow_PV(Index, index2)
                    BtnStat(Index, index2).LastVal(3) = Stn_Btn_Flow_PV(Index, index2)
                    BtnStat(Index, index2).LastVal(4) = Stn_Btn_Flow_PV(Index, index2)
                    BtnStat(Index, index2).FirstTime = False
                Else
                    ' On subsequent cycles, rotate values through
                    ' Use the oldest value to update stats
                    ' Oldest value is # 1, most recent value is #4
                    tvalue = BtnStat(Index, index2).LastVal(1)
                    BtnStat(Index, index2).LastVal(1) = BtnStat(Index, index2).LastVal(2)
                    BtnStat(Index, index2).LastVal(2) = BtnStat(Index, index2).LastVal(3)
                    BtnStat(Index, index2).LastVal(3) = BtnStat(Index, index2).LastVal(4)
                    BtnStat(Index, index2).LastVal(4) = Stn_Btn_Flow_PV(Index, index2)
                    With StationStatistics(Index, index2).Btn
                        If .sCnt = 0 Then
                            .sMin = tvalue
                            .sMax = tvalue
                            .sAvg = tvalue
                            .sCnt = .sCnt + 1
                        Else
                            .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                            .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                            .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                            .sCnt = .sCnt + 1
                        End If
                    End With
                End If
                If NitStat(Index, index2).FirstTime Then
                    ' On first entry, write same value to all variables
                    NitStat(Index, index2).LastVal(1) = Stn_Nit_Flow_PV(Index, index2)
                    NitStat(Index, index2).LastVal(2) = Stn_Nit_Flow_PV(Index, index2)
                    NitStat(Index, index2).LastVal(3) = Stn_Nit_Flow_PV(Index, index2)
                    NitStat(Index, index2).LastVal(4) = Stn_Nit_Flow_PV(Index, index2)
                    NitStat(Index, index2).FirstTime = False
                Else
                    ' On subsequent cycles, rotate values through
                    ' Use the oldest value to update stats
                    ' Oldest value is # 1, most recent value is #4
                    tvalue = NitStat(Index, index2).LastVal(1)
                    NitStat(Index, index2).LastVal(1) = NitStat(Index, index2).LastVal(2)
                    NitStat(Index, index2).LastVal(2) = NitStat(Index, index2).LastVal(3)
                    NitStat(Index, index2).LastVal(3) = NitStat(Index, index2).LastVal(4)
                    NitStat(Index, index2).LastVal(4) = Stn_Nit_Flow_PV(Index, index2)
                    With StationStatistics(Index, index2).Nit
                        If .sCnt = 0 Then
                            .sMin = tvalue
                            .sMax = tvalue
                            .sAvg = tvalue
                            .sCnt = .sCnt + 1
                        Else
                            .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                            .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                            .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                            .sCnt = .sCnt + 1
                        End If
                    End With
                End If
                ' On first entry, write same value to all variables
                If MixStat(Index, index2).FirstTime Then
                    mvalue = 100 * (Stn_Btn_Flow_PV(Index, index2)) / _
                     (Stn_Btn_Flow_PV(Index, index2) + Stn_Nit_Flow_PV(Index, index2) + 0.0001)
                    MixStat(Index, index2).LastVal(1) = mvalue
                    MixStat(Index, index2).LastVal(2) = mvalue
                    MixStat(Index, index2).LastVal(3) = mvalue
                    MixStat(Index, index2).LastVal(4) = mvalue
                    MixStat(Index, index2).FirstTime = False
                Else
                    ' On subsequent cycles, rotate values through
                    ' Use the oldest value to update stats
                    ' Oldest value is # 1, most recent value is #4
                    tvalue = MixStat(Index, index2).LastVal(1)
                    mvalue = 100 * (Stn_Btn_Flow_PV(Index, index2)) / _
                     (Stn_Btn_Flow_PV(Index, index2) + Stn_Nit_Flow_PV(Index, index2) + 0.0001)
                    MixStat(Index, index2).LastVal(1) = MixStat(Index, index2).LastVal(2)
                    MixStat(Index, index2).LastVal(2) = MixStat(Index, index2).LastVal(3)
                    MixStat(Index, index2).LastVal(3) = MixStat(Index, index2).LastVal(4)
                    MixStat(Index, index2).LastVal(4) = mvalue
                    With StationStatistics(Index, index2).Mix
                        If .sCnt = 0 Then
                            .sMin = tvalue
                            .sMax = tvalue
                            .sAvg = tvalue
                            .sCnt = .sCnt + 1
                        Else
                            .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                            .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                            .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                            .sCnt = .sCnt + 1
                        End If
                    End With
                End If
              
                ' LiveFuel Tank Temperature
                tvalue = Stn_AIO(Index, asFuelTankTemp).EUValue
                With StationStatistics(Index, index2).FuelTemp
                    If ((STN_INFO(Index).Type = STN_LIVEFUEL_TYPE) Or ((STN_INFO(Index).Type = STN_LIVEREG_TYPE) And StationRecipe(Index, index2).LiveFuel) Or ((STN_INFO(Index).Type = STN_LIVEORVR2_TYPE) And StationRecipe(Index, index2).LiveFuel)) Then
                        If .sCnt = 0 Then
                            .sMin = tvalue
                            .sMax = tvalue
                            .sAvg = tvalue
                            .sCnt = .sCnt + 1
                        Else
                            .sMin = IIf(tvalue < .sMin, tvalue, .sMin)
                            .sMax = IIf(tvalue > .sMax, tvalue, .sMax)
                            .sAvg = ((.sAvg * .sCnt) + tvalue) / (.sCnt + 1)
                            .sCnt = .sCnt + 1
                        End If
                    Else
                        .sMin = 0
                        .sMax = 0
                        .sAvg = 0
                        .sCnt = 0
                    End If
                End With
              
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
