Attribute VB_Name = "Module8"
' error module 8 ''''''''''''''''''''' controls, controllers and sequencers ''''''''''''''''''
Option Explicit
Private firstPassSim As Boolean
Private antiRepeat(0 To MAX_STN, 0 To MAX_SHIFT) As Boolean
'
'
'
Function ADF_Heater(station As Integer) '
' Routine Name:  AutoDrainFill Heater Control
' Author:        MMW
' Description:
' Controls the AutoDrainFill Heater
'   by requesting that the ADF_Sequence routine
'   actually turn the Heater On and Off
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 1111
Dim delta As Single
Dim delta2 As Single
Dim delta3 As Single
Dim delta4 As Single
Dim Target As Single
Dim LoadRate As Single
Dim temptime As Date

    If AdfControl(station).Heater_Enable Then
    
        If Stn_DIO(station, isFuelHeaterSSR).Value Then
            AdfControl(station).HeaterOff = 0
            If AdfControl(station).HeaterOn < 9999 Then AdfControl(station).HeaterOn = AdfControl(station).HeaterOn + 2
        Else
            If AdfControl(station).HeaterOff < 9999 Then AdfControl(station).HeaterOff = AdfControl(station).HeaterOff + 2
            AdfControl(station).HeaterOn = 0
        End If
        
        ' Check for Temperature within Desired Zone
        LoadRate = StationRecipe(station, 1).NitrogenFlow / Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax
        If LoadRate > 0.55 Then
            delta = StationConfig(station, 1).Tol_FuelTemp
            delta2 = 0.55 * StationConfig(station, 1).Tol_FuelTemp
        ElseIf LoadRate > 0.3 Then
            delta = StationConfig(station, 1).Tol_FuelTemp
            delta2 = 0.15 * StationConfig(station, 1).Tol_FuelTemp
        ElseIf LoadRate > 0.1 Then
            delta = StationConfig(station, 1).Tol_FuelTemp
            delta2 = 0#
        Else
            delta = StationConfig(station, 1).Tol_FuelTemp
            delta2 = 0#
        End If
        Target = AdfControl(station).HeaterSP
        
        If Stn_AIO(station, asFuelTankTemp).EUValue < (Target + delta) _
          And Stn_AIO(station, asFuelTankTemp).EUValue > (Target + delta2) Then
            AdfControl(station).TempinTol = AdfControl(station).TempinTol + 2
        Else
            AdfControl(station).TempinTol = 0
        End If
        
    Else
    
        AdfControl(station).HeaterOff = 0
        AdfControl(station).HeaterOn = 0
        AdfControl(station).TempinTol = 0
    
    End If
            
    ' Is Temperature in Zone & Stable Enough?
    If AdfControl(station).TempinTol > 20 Then
        AdfControl(station).TempOK = True
    Else
        AdfControl(station).TempOK = False
    End If
    
        
    '   HEATER ON/OFF REQUEST CONTROL
    '
    '       Note: Actual HeaterSSR On & Off is done by the ADF_Sequence routine
    '
    If StationControl(station, 1).Mode <> VBIDLE _
        And StationControl(station, 1).Mode <> VBPAUSEALARM _
        And AdfControl(station).Heater And AdfControl(station).Heater_Enable _
        Then
    
        LoadRate = StationRecipe(station, 1).NitrogenFlow / Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax
        If LoadRate > 0.5 Then
            delta3 = 0.75 * StationConfig(station, 1).Tol_FuelTemp
            delta4 = 0.65 * StationConfig(station, 1).Tol_FuelTemp
        ElseIf LoadRate > 0.3 Then
            delta3 = 0.45 * StationConfig(station, 1).Tol_FuelTemp
            delta4 = 0.15 * StationConfig(station, 1).Tol_FuelTemp
        ElseIf LoadRate > 0.1 Then
            delta3 = 0.15 * StationConfig(station, 1).Tol_FuelTemp
            delta4 = 0#
        Else
            delta3 = 0#
            delta4 = 0#
        End If
        Target = AdfControl(station).HeaterSP
        
        ' Check Safeties / Permissives
        If Stn_DIO(station, isFuelSafetyLevelLS).Value And Stn_DIO(station, isFuelLowLevelLS).Value _
            And Not Stn_DIO(station, isFuelHiHiLevelLS).Value And Not Stn_DIO(station, isFuelOverTempSw).Value Then
            
            Select Case StationControl(station, 1).Mode
            
                Case VBLOAD
                    ' Only Turn Heater Off if Fuel Temp is approaching Max Fuel Temp Tolerance Limit
                    If Stn_AIO(station, asFuelTankTemp).EUValue > Target + delta3 Then
                        AdfControl(station).TurnHeaterOn = False        ' Turn Heater Off
                    Else
                        AdfControl(station).TurnHeaterOn = True         ' Turn Heater On
                    End If
                
                Case Else
                    If Stn_DIO(station, isFuelHeaterSSR).Value Then
                        ' Turn Heater Off when Fuel Temp is above Target
                        If Stn_AIO(station, asFuelTankTemp).EUValue > Target + delta4 Then AdfControl(station).TurnHeaterOn = False        ' Turn Heater Off                                ' Turn Heater Off
                    Else
                        ' Wait at least 12 seconds before turning the Heater back On
                        If Stn_AIO(station, asFuelTankTemp).EUValue < Target + delta4 And AdfControl(station).HeaterOff > 12 Then AdfControl(station).TurnHeaterOn = True          ' Turn Heater On
                    End If
                
            End Select
        
        Else
            ' Turn Heater Off, if it is On
            If Stn_DIO(station, isFuelHeaterSSR).Value Then AdfControl(station).TurnHeaterOn = False        ' Turn Heater Off
        End If
        
    Else    ' Not Using Heater Or Don't Have One
        If AdfControl(station).AdfDefinition.hasADF_Heater Then
            ' Turn Heater Off, if it is On
            If Stn_DIO(station, isFuelHeaterSSR).Value Then AdfControl(station).TurnHeaterOn = False        ' Turn Heater Off
        End If
    End If
    
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

Sub ADF_Sequence(station As Integer)

' Routine Name:  AutoDrainFill Sequence
' Author:        MMW
' Description:
' Controls the AutoDrainFill Sequence and the (Optional) Heater SSR.
'
'   ADF_Mode
'   0       Idle
'   1       Drain Only
'   2       Drain then Fill
'   3       Heater Only
'   4       WaterBath Only
'


Dim HighLevelSw As Boolean
Dim LowLevelSw As Boolean
Dim delta, Target, sheathmax As Single
Dim temptime As Date
Dim tempSec As Integer
Dim flag As Boolean
Dim Shift As Integer
Dim Nitrogen_Rate As Single
Dim Nitrogen_Output As Single
Dim span As Single

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 1121

'   Which Type of Sequence are we Executing?
If AdfControl(station).Mode = 0 Then
    If (STN_INFO(station).ADF_TANKTYPE = 0 Or Not AdfControl(station).LiveFuelChgAuto) Then
        AdfControl(station).Task = "Manual Drain/Fill"
        AdfControl(station).Message = "Idle"
    ElseIf (STN_INFO(station).ADF_TANKTYPE <> 0) Then
        If AdfControl(station).Heater Then
            If (STN_INFO(station).ADF_TANKTYPE = 90) Then
                ' WaterBath Only, No ADF
                If AdfControl(station).Heater_Enable Then
                    AdfControl(station).Task = "Maintaining WaterBath Temperature"
                Else
                    AdfControl(station).Task = "WaterBath Idle"
                    AdfControl(station).Message = "Idle"
                End If
            Else
                ' Electric Heater
                If AdfControl(station).Heater_Enable Then
                    AdfControl(station).Task = "Maintaining Fuel Temperature"
                Else
                    AdfControl(station).Task = "Auto Drain/Fill Idle"
                    AdfControl(station).Message = "Idle"
                End If
            End If
        Else
            AdfControl(station).Task = "Auto Drain/Fill Idle"
            AdfControl(station).Message = "Idle"
        End If
    End If
End If
If AdfControl(station).Mode = 1 Then AdfControl(station).Task = "Drain Only in Progress"
If AdfControl(station).Mode = 2 Then AdfControl(station).Task = "Auto Refill Control in Progress"
If AdfControl(station).Mode = 3 Then AdfControl(station).Task = "Maintaining Fuel Temperature"
If AdfControl(station).Mode = 4 Then AdfControl(station).Task = "Maintaining WaterBath Temperature"

' Turn (Optional)Heater Off unless Temp Control is Active
If AdfControl(station).AdfDefinition.hasADF_Heater And AdfControl(station).Step <> 39 And AdfControl(station).Step <> 99 Then
    If Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cOFF    ' Turn Heater Off
End If

'   If Mode = 0 then Nothing to do
If AdfControl(station).Mode = 0 Then
    AdfControl(station).Heater_Enable = False
    AdfControl(station).ButtonVisible_Done = False
    AdfControl(station).ButtonVisible_Retry = False
    AdfControl(station).ButtonVisible_Stop = False
    Exit Sub
End If
'   Auto Drain/Fill Sequence
'
Select Case AdfControl(station).Step

  Case 0
      AdfControl(station).Message = "Start Drain"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).ButtonVisible_Done = False
      AdfControl(station).ButtonVisible_Retry = False
      AdfControl(station).ButtonVisible_Stop = False
      If AdfControl(station).AdfDefinition.hasADF_Heater Then
          Select Case STN_INFO(station).ADF_TANKTYPE
            Case 20
                ' Stant     (LT, no FST, Heater)
                AdfControl(station).Step = 4           ' Has Heater; wait cool sheath
            Case 12
                ' Mahle     (no LT, no FST, Heater)
                AdfControl(station).Step = 1           ' Has Heater; wait for N2 pressurize/purge & cool sheath
          End Select
      Else
          AdfControl(station).Step = 11          ' No Heater; start drain
      End If
      
  Case 1
      AdfControl(station).Message = "Energize N2 Pressurize Sol"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).ReadyForLoad = False
      AdfControl(station).ButtonVisible_Done = False
      AdfControl(station).ButtonVisible_Retry = False
      AdfControl(station).ButtonVisible_Stop = True
      ' Mahle     (no LT, no FST, Heater)
      Stn_OutDigital station, isFuelPressSol, cON
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).PurgeTimeout)    ' setup timeout for Purge
      AdfControl(station).Step = 2
      
  Case 2
      AdfControl(station).Message = "Waiting for N2 Pressure Switch"
      AdfControl(station).Heater_Enable = False
      If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      If Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cOFF
      If Stn_DIO(station, isFuelVaporSol).Value Then Stn_OutDigital station, isFuelVaporSol, cOFF
      If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
      If Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cOFF
      If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If (Now() > AdfControl(station).Step_Time) Or Stn_DIO(station, isFuelPressPS).Value Then
          If Now() > AdfControl(station).Step_Time Then
              AdfControl(station).Step = 94              ' Timeout
          Else
              AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).PurgeDrainDelay)
              AdfControl(station).Step = 3               ' Continue
          End If
                      
      End If
      
  Case 3
      temptime = AdfControl(station).Step_Time - Now()
      tempSec = (60 * Minute(temptime)) + Second(temptime)
      AdfControl(station).Message = "Waiting for N2 Pressurize Delay - " & Format(tempSec, "##0") & " sec"
      AdfControl(station).Heater_Enable = False
      If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If Not Stn_DIO(station, isFuelPressPS).Value Then
          AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
          AdfControl(station).Step = 2               ' Go Back and Wait for N2 PS
      End If
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step = 4               ' Continue
      End If
     
  Case 4
      If (STN_INFO(station).ADF_TANKTYPE = 12) Then
        ' Mahle     (no LT, no FST, Heater)
        sheathmax = AdfControl(station).HeaterSP + StationConfig(station, 1).Tol_FuelTemp
      Else
        If USINGC Then
            ' deg C
            sheathmax = MaxSheathTempForAdfDrain
            If sheathmax > 70# Then sheathmax = 70#
            If sheathmax < AmbTemp Then sheathmax = AmbTemp
            AdfControl(station).Message = "Waiting for Sheath Temp Below " & Format(sheathmax, "#00") & " deg C"
        ElseIf USINGF Then
            ' deg F
            sheathmax = DegCtoF(MaxSheathTempForAdfDrain)
            If sheathmax > 160# Then sheathmax = 160#
            If sheathmax < AmbTemp Then sheathmax = AmbTemp
            AdfControl(station).Message = "Waiting for Sheath Temp Below " & Format(sheathmax, "#00") & " deg F"
        End If
      End If
      AdfControl(station).Heater_Enable = False
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      ' Need Circulation ?
      If StationControl(station, 1).Mode <> VBLOAD Then
          ' Circulate if Heater is On OR if Heater Sheath Temp is Too High
          If ((Stn_DIO(station, isFuelHeaterSSR).Value) _
              Or _
            (Stn_AIO(station, asFuelHeaterTemp).EUValue > sheathmax)) Then
              ' CIRCULATE
              If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
              If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
              If Not Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cON
              If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
              If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
          Else
              ' DO NOT CIRCULATE
              If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
              If Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cOFF
              If Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cOFF
              If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
              If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
          End If
      End If
      If (STN_INFO(station).ADF_TANKTYPE = 12) Then
        ' Mahle     (no LT, no FST, Heater)
        If Not Stn_DIO(station, isFuelPressPS).Value Then
            AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
            AdfControl(station).Step = 2               ' Go Back and Wait for N2 PS
        End If
      End If
      If Stn_AIO(station, asFuelHeaterTemp).EUValue < sheathmax Then
          AdfControl(station).Step = 11              ' Continue
      End If
  
  Case 11
      AdfControl(station).Message = "Energize Pump and Drain"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).ReadyForLoad = False
      AdfControl(station).ButtonVisible_Done = False
      AdfControl(station).ButtonVisible_Retry = False
      AdfControl(station).ButtonVisible_Stop = True
      Stn_OutDigital station, isFuelDrainSol, cON
      Stn_OutDigital station, isFuelPumpMotor, cON
      If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      If Not Stn_DIO(station, isFuelVentSol).Value Then Stn_OutDigital station, isFuelVentSol, cON
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).DrainTimeout)   ' setup timeout for Drain
      AdfControl(station).Step = 12
      
  Case 12
      Select Case STN_INFO(station).ADF_DEF.hasADF_LT
        Case False
            AdfControl(station).Message = "Waiting for Low Level Switch"
            flag = Not Stn_DIO(station, isFuelLowLevelLS).Value
            AdfControl(station).Heater_Enable = False
            If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
            If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
            If STN_INFO(station).ADF_TANKTYPE = 12 Then
              ' Mahle
              If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
              If Not Stn_DIO(station, isFuelPressPS).Value Then
                  AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
                  AdfControl(station).Step = 2                                   ' Go Back and Wait for N2 PS
              End If
            End If
        Case True
            AdfControl(station).Message = "Waiting for Low (" & Format(StationCfg_ADF(station, 1).DrainShutOff, "###0.0") & " %) Level"
            flag = IIf((Stn_AIO(station, asFuelTankLevel).EUValue <= StationCfg_ADF(station, 1).DrainShutOff), True, False)
            AdfControl(station).Heater_Enable = False
            If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
            If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
        Case Else
            AdfControl(station).Message = "Waiting for Low Level Switch"
            flag = Not Stn_DIO(station, isFuelLowLevelLS).Value
            AdfControl(station).Heater_Enable = False
            If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
            If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
            If STN_INFO(station).ADF_TANKTYPE = 12 Then
              ' Mahle
              If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
              If Not Stn_DIO(station, isFuelPressPS).Value Then
                  AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
                  AdfControl(station).Step = 2                                   ' Go Back and Wait for N2 PS
              End If
            End If
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If (Now() > AdfControl(station).Step_Time) Or flag Then
        If Now() > AdfControl(station).Step_Time Then
            AdfControl(station).Step = 91              ' Timeout
        Else
            AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).DrainDelay)
            AdfControl(station).Step = 13              ' Continue
        End If
      End If
  
  Case 13
      temptime = AdfControl(station).Step_Time - Now()
      tempSec = (60 * Minute(temptime)) + Second(temptime)
      AdfControl(station).Message = "Pumping to Clear Piping - " & Format(tempSec, "##0") & " sec"
      AdfControl(station).Heater_Enable = False
      If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
      If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
      If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If AdfControl(station).AdfDefinition.hasADF_Heater Then
            If STN_INFO(station).ADF_TANKTYPE = 12 Then
              ' Mahle
              If Not Stn_DIO(station, isFuelPressPS).Value Then
                AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
                AdfControl(station).Step = 2                                   ' Go Back and Wait for N2 PS
              End If
            End If
      End If
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step = 15              ' Continue
      End If
     
  Case 15
      AdfControl(station).Message = "Deenergize Pump and Drain"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      If Not Stn_DIO(station, isFuelVentSol).Value Then Stn_OutDigital station, isFuelVentSol, cON
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).DrainDelay)
      AdfControl(station).Step = 16                  ' Continue
  
  Case 16
      temptime = AdfControl(station).Step_Time - Now()
      tempSec = (60 * Minute(temptime)) + Second(temptime)
      If Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cOFF
      If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
      If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If AdfControl(station).AdfDefinition.hasADF_Heater Then
        If (STN_INFO(station).ADF_TANKTYPE = 12) Then
          ' Mahle     (no LT, no FST, Heater)
          If Not Stn_DIO(station, isFuelPressPS).Value Then
            AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
            AdfControl(station).Step = 2                                   ' Go Back and Wait for N2 PS
          End If
        End If
      End If
      Select Case STN_INFO(station).ADF_DEF.hasADF_LT
        Case False
            AdfControl(station).Message = "Checking All Level Sw. Off - " & Format(tempSec, "##0") & " sec"
            AdfControl(station).Heater_Enable = False
            If Now() > AdfControl(station).Step_Time Then
                If Not Stn_DIO(station, isFuelLowLevelLS).Value And Not Stn_DIO(station, isFuelHighLevelLS).Value Then
                    ' Tank is Empty; Deenergize N2 Pressurization
                    AdfControl(station).Step = 18
                Else
                    AdfControl(station).Step = 11       ' Resume Draining
                End If
            End If
        Case True
            AdfControl(station).Message = "Checking Level - " & Format(tempSec, "##0") & " sec"
            AdfControl(station).Heater_Enable = False
            If Now() > AdfControl(station).Step_Time Then
                If Not Stn_DIO(station, isFuelLowLevelLS).Value And (Stn_AIO(station, asFuelTankLevel).EUValue <= StationCfg_ADF(station, 1).DrainShutOff) Then
                    ' Tank is Empty
                    AdfControl(station).Step = 19
                Else
                    AdfControl(station).Step = 11       ' Resume Draining
                End If
            End If
      End Select
  
  Case 18
      AdfControl(station).Message = "Deenergize N2 Pressurize Sol"
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step = 19
         
  Case 19
      AdfControl(station).Message = "Drain Complete"
      ' request set of LiveFuel State to "OK"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If AdfControl(station).Mode = 1 Then AdfControl(station).Step = 0
      If AdfControl(station).Mode = 1 Then AdfControl(station).Mode = 0
      If AdfControl(station).Mode = 2 Then AdfControl(station).Step = 20
         
  Case 20
      AdfControl(station).Message = "Checking Storage Tank Level"
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If AdfControl(station).Mode = 2 Then
        If AdfControl(station).AdfDefinition.hasADF_FST Then
            If Stn_DIO(station, isStorageLowLevelLS).Value Then
                AdfControl(station).Step = 21   ' Proceed to Fill Operation
            Else
                AdfControl(station).Step = 97   ' Storage Tank Level Too Low
            End If
        Else
            AdfControl(station).Step = 21       ' Proceed to Fill Operation
        End If
      End If
         
  Case 21
      AdfControl(station).Message = "Fill Vapor Tank"
      AdfControl(station).ButtonVisible_Done = False
      AdfControl(station).ButtonVisible_Retry = False
      AdfControl(station).ButtonVisible_Stop = True
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelRecircSol, cON
      Stn_OutDigital station, isFuelFillSol, cON
      Stn_OutDigital station, isFuelVentSol, cON
      If AdfControl(station).AdfDefinition.hasADF_Heater Then
          Stn_OutDigital station, isFuelPumpMotor, cON
      End If
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).FillTimeout)    ' setup timeout for Fill
      AdfControl(station).Step = 22
  
  Case 22
      Select Case STN_INFO(station).ADF_DEF.hasADF_LT
        Case False
            AdfControl(station).Message = "Waiting for High Level Switch"
            flag = Stn_DIO(station, isFuelHighLevelLS).Value
            AdfControl(station).Heater_Enable = False
            If Not Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cON
            If Not Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cON
            If Not Stn_DIO(station, isFuelVentSol).Value Then Stn_OutDigital station, isFuelVentSol, cON
            If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
            If STN_INFO(station).ADF_DEF.hasADF_PS Then
                If STN_INFO(station).ADF_TANKTYPE = 12 Then
                  ' Mahle
                  If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
                  If Not Stn_DIO(station, isFuelPressPS).Value Then
                    AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 1)    ' reset timeout to 1 sec
                    AdfControl(station).Step = 2                                   ' Go Back and Wait for N2 PS
                  End If
                End If
            End If
        Case True
            ' waiting for LT to reach desired value
            AdfControl(station).Message = "Waiting for High (" & Format(StationCfg_ADF(station, 1).FillShutOff, "###0.0") & " %) Level"
            flag = IIf((Stn_AIO(station, asFuelTankLevel).EUValue >= StationCfg_ADF(station, 1).FillShutOff), True, False)
            AdfControl(station).Heater_Enable = False
            If Not Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cON
            If Not Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cON
            If Not Stn_DIO(station, isFuelVentSol).Value Then Stn_OutDigital station, isFuelVentSol, cON
            If AdfControl(station).AdfDefinition.hasADF_Heater Then
                If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
            End If
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If flag Then
        AdfControl(station).Step = 23  ' Continue
      ElseIf (Now() > AdfControl(station).Step_Time) Then
        AdfControl(station).Step = 92  ' Timeout
      End If
  
  Case 23
      AdfControl(station).Message = "Deenergize Pump and Fill"
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      AdfControl(station).Heater_Enable = False
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).FillDelay)
      AdfControl(station).Step = 24
  
  Case 24
      temptime = AdfControl(station).Step_Time - Now()
      tempSec = (60 * Minute(temptime)) + Second(temptime)
      AdfControl(station).Message = "Waiting for Fill Delay - " & Format(tempSec, "##0") & " sec"
      If Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cOFF
      If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
      If Stn_DIO(station, isFuelVentSol).Value Then Stn_OutDigital station, isFuelVentSol, cOFF
      If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
      AdfControl(station).Heater_Enable = False
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step_Time = Now() + TimeSerial(0, StationCfg_ADF(station, 1).HeaterTimeout, 0)
          AdfControl(station).Step = 25
      End If
  
  Case 25
    ' double check desired level (and check that LowLevel Sitch is ON)
      Select Case STN_INFO(station).ADF_DEF.hasADF_LT
        Case False
            AdfControl(station).Message = "Checking All Level Sw. On"
            AdfControl(station).Heater_Enable = False
            If STN_INFO(station).ADF_DEF.hasADF_Heater Then
                flag = Stn_DIO(station, isFuelLowLevelLS).Value And Stn_DIO(station, isFuelHighLevelLS).Value And Stn_DIO(station, isFuelSafetyLevelLS).Value
            Else
                flag = Stn_DIO(station, isFuelLowLevelLS).Value And Stn_DIO(station, isFuelHighLevelLS).Value
            End If
        Case True
            AdfControl(station).Message = "Checking Level"
            AdfControl(station).Heater_Enable = False
            flag = Stn_DIO(station, isFuelLowLevelLS).Value And (Stn_AIO(station, asFuelTankLevel).EUValue >= StationCfg_ADF(station, 1).FillShutOff)
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If flag Then
            AdfControl(station).SetOkRequest = True
            StationControl(station, 1).LiveFuelCycleCount = 0
            AdfControl(station).InitialFill_Complete = True
            AdfControl(station).ReadyForRefill = False
            If AdfControl(station).Heater Then
              AdfControl(station).Heater_Enable = True
              AdfControl(station).Step = 31   ' Continue; Heater in Use
            Else
              AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 5)         ' delay for level to settle
              AdfControl(station).Step = 49   ' Fill Complete;No Heater; Monitor Level
            End If
      ElseIf (Stn_AIO(station, asFuelTankLevel).EUValue >= StationCfg_ADF(station, 1).FillShutOff) Then
            ' level is good, LowLevel Switch Isn't
            AdfControl(station).Step = 61     ' Ask Operator How to Proceed
      Else
            AdfControl(station).Step = 20     ' Resume Filling
      End If
  
  Case 31
      AdfControl(station).Message = "Energize N2 Pressurize"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).ReadyForLoad = False
      AdfControl(station).ButtonVisible_Done = False
      AdfControl(station).ButtonVisible_Retry = False
      AdfControl(station).ButtonVisible_Stop = True
      Shift = Stn_ActiveShift(station)
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' use HIGHER RANGE MFCs
                ' use HIGHER RANGE MFCs
                ' use HIGHER RANGE MFCs
                ' open LiveFuel ORVR valves
                Stn_OutDigital station, isFuelVentSol, cOFF
                If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isFuelVaporSol, cON
                Stn_OutDigital station, isLiveFuelOrvrSol, cON
                Stn_OutDigital station, isLoadTypeSelectSol, cON
                ' set LiveFuel Vapor Carrier ORVR MFC setpoint
                Nitrogen_Rate = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMax / 2
                Stn_Nit_FlowSP(station, Shift) = Nitrogen_Rate
                
                span = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCORVRLIVE, Stn_MfcCal(station, MFCORVRLIVE)))
                Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, Nitrogen_Output, outNORMAL
            Else
                ' use LOWER RANGE MFCs
                ' use LOWER RANGE MFCs
                ' use LOWER RANGE MFCs
                ' open LiveFuel valves
                Stn_OutDigital station, isFuelVentSol, cOFF
                If StationControl(station, Shift).Mode = VBLOAD Then Stn_OutDigital station, isFuelVaporSol, cON
                Stn_OutDigital station, isLiveFuelSol, cON
                Stn_OutDigital station, isLoadTypeSelectSol, cON
                ' set LiveFuel Vapor Carrier MFC setpoint
                Nitrogen_Rate = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax / 2
                Stn_Nit_FlowSP(station, Shift) = Nitrogen_Rate
                
                span = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin
                Nitrogen_Output = Stn_AIO(station, asLiveFuelVaporFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL)))
                Stn_OutAnalog station, asLiveFuelVaporFlowSP, Nitrogen_Output, outNORMAL
            End If
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            Stn_OutDigital station, isFuelPressSol, cON
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If AdfControl(station).AdfDefinition.hasADF_Heater Then
          AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 2 * StationCfg_ADF(station, 1).PurgeTimeout)   ' longer purge timeout if Heater Only
      Else
          AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).PurgeTimeout)       ' setup timeout for Purge
      End If
      AdfControl(station).Step = 32                        ' Continue
      
  Case 32
      AdfControl(station).Message = "Waiting for N2 Pressure Switch"
      AdfControl(station).Heater_Enable = False
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If (Now() > AdfControl(station).Step_Time) Then
         AdfControl(station).Step = 94                ' Timeout
      ElseIf (Stn_DIO(station, isFuelPressPS).Value) Then
        If AdfControl(station).AdfDefinition.hasADF_Heater Then
            AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, 2 * StationCfg_ADF(station, 1).PurgeFillDelay)   ' longer purge timeout if Heater Only
        Else
            AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).PurgeFillDelay)       ' setup timeout for Purge
        End If
        AdfControl(station).Step = 33                ' Continue
      End If
      
  Case 33
      temptime = AdfControl(station).Step_Time - Now()
      tempSec = (60 * Minute(temptime)) + Second(temptime)
      AdfControl(station).Message = "Waiting for N2 Pressurize Delay - " & Format(tempSec, "##0") & " sec"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      End Select
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step = 34
      End If
     
  Case 34
      AdfControl(station).Message = "Deenergize N2 Pressurize"
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
            Stn_OutAnalog station, asLiveFuelVaporFlowSP, CSng(0), outZERO
            Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, CSng(0), outZERO
            Stn_OutDigital station, isLiveFuelSol, cOFF
            Stn_OutDigital station, isLiveFuelOrvrSol, cOFF
            Stn_OutDigital station, isLoadTypeSelectSol, cOFF
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            Stn_OutDigital station, isFuelPressSol, cOFF
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, StationCfg_ADF(station, 1).HeaterTimeout, 0)
      AdfControl(station).Step = 39
         
  Case 39
      AdfControl(station).Message = "Waiting for Tank Temp"
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Heater_Enable = True
      ' Heater Requested ?
      If AdfControl(station).TurnHeaterOn Then
        ' Check N2 Pressure Switch before turning Heater On
        ' except During Load when don't care about the N2 PS
        If AdfControl(station).HeaterOn = 0 And StationControl(station, 1).Mode <> VBLOAD Then
            ' Ok to turn Heater On ?
            Select Case STN_INFO(station).ADF_TANKTYPE
              Case 20
                  ' Stant     (LT, no FST, Heater)
                  ' Turn the Heater On
                  If Not Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cON    ' Turn Heater On
              Case 12
                  ' Mahle     (no LT, no FST, Heater)
                  If Not Stn_DIO(station, isFuelPressPS).Value Then
                    ' N2 Pressure Switch is Not On
                    If Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cOFF    ' Turn Heater Off
                    AdfControl(station).Step = 101         ' Repressurize the Tank
                  Else
                    ' N2 Pressure Switch is On; Turn the Heater On
                    If Not Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cON    ' Turn Heater On
                  End If
            End Select
        Else
            ' Heater is already On; Keep it On (Or Load is in Progress)
            If Not Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cON    ' Turn Heater On
        End If
      Else
        ' Heater Not Requested
        If Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cOFF    ' Turn Heater Off
      End If
      ' Need Circulation ?
      If StationControl(station, 1).Mode <> VBLOAD Then
          delta = Stn_AIO(station, asFuelHeaterTemp).EUValue - Stn_AIO(station, asFuelTankTemp).EUValue
          If USINGC Then Target = 3#
          If USINGF Then Target = 5#
          ' Circulate if Heater is On OR if Heater Sheath Temp is Too High
          If Stn_DIO(station, isFuelHeaterSSR).Value _
              Or _
                  (delta > Target) And (Stn_DIO(station, isFuelPumpMotor).Value _
                      Or _
                  (delta > (1.5 * Target))) Then
              ' CIRCULATE
              If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
              If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
              If Not Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cON
              If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
              If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
          Else
              ' DO NOT CIRCULATE
              If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
              If Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cOFF
              If Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cOFF
              If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
              If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
          End If
      End If
      ' Is Tank @ Temp ?
      If AdfControl(station).TempOK Then
          AdfControl(station).ReadyForLoad = True
          AdfControl(station).Step = 99  ' Tank @ Temp
          If StationControl(station, 1).Mode = VBGASPAUSE Then LoadControl(station, 1).CycleStartRequest = True
      End If
      ' Timeout ?
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step = 93  ' Timeout
      End If
      AdfControl(station).ButtonVisible_Retry = False
      AdfControl(station).ButtonVisible_Stop = True
  
  Case 49
      AdfControl(station).Message = "Fill Cycle Complete"
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
'      AdfControl(station).Mode = 0
      ' Delay Timeout ?
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step = 100        ' monitor level
      End If
      AdfControl(station).Heater_Enable = False
      AdfControl(station).ReadyForLoad = True
      If StationControl(station, 1).Mode = VBGASPAUSE Then LoadControl(station, 1).CycleStartRequest = True
  
  Case 61
      AdfControl(station).Message = "LowLevel Switch is Off with Fuel in the Tank"
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Ignore = True
      AdfControl(station).ButtonVisible_Retry = True
      AdfControl(station).ButtonVisible_Stop = True
         
  Case 89
      AdfControl(station).Message = "Paused by Operator"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
      AdfControl(station).ButtonVisible_Stop = True
  
  Case 90
      AdfControl(station).Message = "Aborted - Station Stopped"
      AdfControl(station).Mode = 0
      AdfControl(station).Step = 0
      AdfControl(station).StepBeforePause = 0
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = False
  
  Case 91
      AdfControl(station).Message = "Aborted - Too Long to Drain"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 92
      AdfControl(station).Message = "Aborted - Too Long to Fill"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 93
      AdfControl(station).Message = "Aborted - Too Long to Reach Temp"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 94
      AdfControl(station).Message = "Aborted - Too Long to Turn On N2 PS"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 95
      AdfControl(station).Message = "Aborted - Station Paused"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = False
  
  Case 96
      AdfControl(station).Message = "Aborted - Failed To Maintain N2 Pressure"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 97
      AdfControl(station).Message = "Aborted - Storage Tank Level is Too Low"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 98
      AdfControl(station).Message = "Aborted - Vapor Tank Level is Too Low"
      AdfControl(station).Heater_Enable = False
      Stn_OutDigital station, isFuelPumpMotor, cOFF
      Stn_OutDigital station, isFuelRecircSol, cOFF
      Stn_OutDigital station, isFuelPressSol, cOFF
      Stn_OutDigital station, isFuelDrainSol, cOFF
      Stn_OutDigital station, isFuelFillSol, cOFF
      Stn_OutDigital station, isFuelHeaterSSR, cOFF
      Stn_OutDigital station, isFuelVaporSol, cOFF
      Stn_OutDigital station, isFuelVentSol, cOFF
      Stn_OutDigital station, isLiveFuelSol, cOFF
      Stn_OutDigital station, isLoadTypeSelectSol, cOFF
      AdfControl(station).ButtonVisible_Retry = True
  
  Case 99
      AdfControl(station).Message = "Maintaining Fuel Temperature"
      If AdfControl(station).TempOK Then
          AdfControl(station).ReadyForLoad = True
          If StationControl(station, 1).Mode = VBGASPAUSE Then LoadControl(station, 1).CycleStartRequest = True
      End If
      AdfControl(station).Heater_Enable = True
      ' Heater Requested ?
      If AdfControl(station).TurnHeaterOn Then
        ' Check N2 Pressure Switch before turning Heater On
        ' except During Load when don't care about the N2 PS
        If AdfControl(station).HeaterOn = 0 And StationControl(station, 1).Mode <> VBLOAD Then
            ' Ok to turn Heater On ?
            Select Case STN_INFO(station).ADF_TANKTYPE
                Case 20     ' Stant     (LT, no FST, Heater)
                  ' Turn the Heater On
                  If Not Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cON    ' Turn Heater On
              Case 12       ' Mahle     (no LT, no FST, Heater)
                  If Not Stn_DIO(station, isFuelPressPS).Value Then
                    ' N2 Pressure Switch is Not On
                    If Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cOFF    ' Turn Heater Off
                    AdfControl(station).Step = 101         ' Repressurize the Tank
                  Else
                    ' N2 Pressure Switch is On; Turn the Heater On
                    If Not Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cON    ' Turn Heater On
                  End If
            End Select
        Else
            ' Heater is already On; Keep it On (Or Load is in Progress)
            If Not Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cON    ' Turn Heater On
        End If
      Else
        ' Heater Not Requested
        If Stn_DIO(station, isFuelHeaterSSR).Value Then Stn_OutDigital station, isFuelHeaterSSR, cOFF    ' Turn Heater Off
      End If
      ' Need Circulation ?
      If StationControl(station, 1).Mode <> VBLOAD Then
          delta = Stn_AIO(station, asFuelHeaterTemp).EUValue - Stn_AIO(station, asFuelTankTemp).EUValue
          If USINGC Then Target = 3#
          If USINGF Then Target = 5#
          ' Circulate if Heater is On OR if Heater Sheath Temp is Too High
          If Stn_DIO(station, isFuelHeaterSSR).Value _
              Or _
                  (delta > Target) And (Stn_DIO(station, isFuelPumpMotor).Value _
                      Or _
                  (delta > (1.5 * Target))) Then
              ' CIRCULATE
              If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
              If Not Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cON
              If Not Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cON
              If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
              If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON
          Else
              ' DO NOT CIRCULATE
              If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
              If Stn_DIO(station, isFuelDrainSol).Value Then Stn_OutDigital station, isFuelDrainSol, cOFF
              If Stn_DIO(station, isFuelRecircSol).Value Then Stn_OutDigital station, isFuelRecircSol, cOFF
              If Stn_DIO(station, isFuelFillSol).Value Then Stn_OutDigital station, isFuelFillSol, cOFF
              If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
          End If
      End If
      AdfControl(station).ButtonVisible_Retry = False
    
  Case 100
      AdfControl(station).Message = "Monitoring Vapor Tank Level"
      If (AdfControl(station).Heater) And (Not Stn_DIO(station, isFuelSafetyLevelLS).Value) Then
        Write_ELog "Tank Level below Sheath during Load for Station " & Format(station, "0")
        AdfControl(station).ReadyForLoad = False
        AdfControl(station).Step = 98  ' Vapor Tank Level Too Low
      End If
  
  Case 101
      AdfControl(station).Message = "Energize N2 Pressurize"
      Write_ELog "Repressurize Live Fuel Tank for Station " & Format(station, "0")
      AdfControl(station).Heater_Enable = False
      AdfControl(station).ReadyForLoad = False
      AdfControl(station).ButtonVisible_Done = False
      AdfControl(station).ButtonVisible_Retry = False
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
            Stn_OutDigital station, isLiveFuelSol, cON
            Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, (Stn_AIO(station, asLiveFuelVaporORVRFlowSP).EuMax / 2), outNORMAL
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            Stn_OutDigital station, isFuelPressSol, cON
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).PurgeTimeout)       ' setup timeout for Pressurize
      AdfControl(station).Step = 102                        ' Continue
      
  Case 102
      AdfControl(station).Message = "Waiting for N2 Pressure Switch"
      AdfControl(station).Heater_Enable = False
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
            If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isLiveFuelSol, cON
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            If Not Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cON
      End Select
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      If (Now() > AdfControl(station).Step_Time) Or Stn_DIO(station, isFuelPressPS).Value Then
          If Now() > AdfControl(station).Step_Time Then
              AdfControl(station).Step = 94                    ' Timeout
          Else
              AdfControl(station).Step = 103                   ' Continue
          End If
      End If
      
  Case 103
      AdfControl(station).Message = "Deenergize N2 Pressurize"
      AdfControl(station).Heater_Enable = False
      AdfControl(station).LevelSP = Stn_AIO(station, asFuelTankLevel).EUValue
      AdfControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).PurgeFillDelay)         ' pressurize delay for Heater
      Select Case STN_INFO(station).ADF_TANKTYPE
        Case 20
            ' Stant     (LT, no FST, Heater)
            Stn_OutAnalog station, asLiveFuelVaporORVRFlowSP, CSng(0), outZERO
            Stn_OutDigital station, isLiveFuelSol, cOFF
            AdfControl(station).Step = 99
        Case 12
            ' Mahle     (no LT, no FST, Heater)
            Stn_OutDigital station, isFuelPressSol, cOFF
            AdfControl(station).Step = 104
      End Select
         
  Case 104
      temptime = AdfControl(station).Step_Time - Now()
      tempSec = (60 * Minute(temptime)) + Second(temptime)
      AdfControl(station).Message = "Watching N2 Pressure Switch - " & Format(tempSec, "##0") & " sec"
      AdfControl(station).Heater_Enable = False
      If Stn_DIO(station, isFuelPressSol).Value Then Stn_OutDigital station, isFuelPressSol, cOFF
      If Now() > AdfControl(station).Step_Time Then
          AdfControl(station).Step = 104
      End If
      If (Now() > AdfControl(station).Step_Time) Or Not Stn_DIO(station, isFuelPressPS).Value Then
          If Now() > AdfControl(station).Step_Time Then
              AdfControl(station).Step = 99                    ' Continue
          Else
              ALM_Write station, 1, "Live Fuel Tank Failed to Hold N2 Pressure"
              AdfControl(station).Step = 96                    ' Abort, Failed to Hold N2 Pressure
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

    
Sub FST_Sequence(station As Integer)
' Routine Name:  Fuel Storage Tank Drain & Fill Sequence
' Author:        MMW
' Description:
' Controls the Drain & Fill Sequence of the (Live)Fuel Storage Tank.
'
'   FST_Mode
'   0       Idle
'   1       Drain
'   2       Fill
'

Dim temptime As Date
Dim tempSec As Integer
Dim flag As Boolean

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 12321

    If FstControl(station).Mode = 0 Then FstControl(station).Task = "Idle"
    If FstControl(station).Mode = 1 Then FstControl(station).Task = "Drain in Progress"
    If FstControl(station).Mode = 2 Then FstControl(station).Task = "Fill in Progress"
    
    If ((STN_INFO(station).ADF_TANKTYPE > 20) And (STN_INFO(station).ADF_TANKTYPE < 90)) Then
    
        '   Fuel Storage Tank Drain & Fill Sequence
        '
        Select Case FstControl(station).Step
        
            Case 0
                Select Case FstControl(station).Mode
                    Case 1
                        ' drain
                        FstControl(station).Message = "Start Drain"
                        FstControl(station).ButtonVisible_Drain = False
                        FstControl(station).ButtonVisible_Fill = False
                        FstControl(station).ButtonVisible_Stop = True
                        FstControl(station).Step = 11                               ' start drain
                    Case 2
                        ' fill
                        FstControl(station).Message = "Start Fill"
                        FstControl(station).ButtonVisible_Drain = False
                        FstControl(station).ButtonVisible_Fill = False
                        FstControl(station).ButtonVisible_Stop = True
                        FstControl(station).Step = 21                               ' start fill
                    Case Else
                        ' idle
                        FstControl(station).Message = "Idle"
                        FstControl(station).ButtonVisible_Drain = True
                        FstControl(station).ButtonVisible_Fill = True
                        FstControl(station).ButtonVisible_Stop = False
                End Select
        
            Case 11
                FstControl(station).Message = "Pump and Drain On"
                FstControl(station).ButtonVisible_Drain = False
                FstControl(station).ButtonVisible_Fill = False
                FstControl(station).ButtonVisible_Stop = True
                Stn_OutDigital station, isStorageDrainSol, cON              ' Open Drain Valve
                Stn_OutDigital station, isFuelPumpMotor, cON                ' Turn Pump ON
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                ' setup timeout for Drain
                FstControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).FstDrainTimeout)
                FstControl(station).Step = 12
              
            Case 12
                FstControl(station).Message = "Wait Low (" & Format(StationCfg_ADF(station, 1).FstDrainShutOff, "###0.0") & " %) Level"
                flag = IIf((Stn_AIO(station, asStorageTankLevel).EUValue <= StationCfg_ADF(station, 1).FstDrainShutOff), True, False)
                If Not Stn_DIO(station, isStorageDrainSol).Value Then Stn_OutDigital station, isStorageDrainSol, cON            ' Open Drain Valve
                If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON                ' Turn Pump ON
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If (Now() > FstControl(station).Step_Time) Or flag Then
                    If Now() > FstControl(station).Step_Time Then
                        FstControl(station).Step = 91                       ' Timeout
                    Else
                        ' setup Drain Delay
                        FstControl(station).Step_Time = Now() + TimeSerial(0, 0, 2 * StationCfg_ADF(station, 1).FstDrainDelay)
                        FstControl(station).Step = 13                       ' Continue
                    End If
                End If
          
            Case 13
                temptime = FstControl(station).Step_Time - Now()
                tempSec = (60 * Minute(temptime)) + Second(temptime)
                FstControl(station).Message = "Clear Pipes - " & Format(tempSec, "##0") & " sec"
                If Not Stn_DIO(station, isStorageDrainSol).Value Then Stn_OutDigital station, isStorageDrainSol, cON            ' Open Drain Valve
                If Not Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cON                ' Turn Pump ON
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If (Now() > FstControl(station).Step_Time) Then
                    FstControl(station).Step = 14                           ' Continue
                End If
            
            Case 14
                FstControl(station).Message = "Pump and Drain Off"
                Stn_OutDigital station, isFuelPumpMotor, cOFF
                Stn_OutDigital station, isStorageDrainSol, cOFF
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                ' setup Drain Delay, again
                FstControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).FstDrainDelay)
                FstControl(station).Step = 15                  ' Continue
          
            Case 15
                temptime = FstControl(station).Step_Time - Now()
                tempSec = (60 * Minute(temptime)) + Second(temptime)
                FstControl(station).Message = "Wait Drain Delay - " & Format(tempSec, "##0") & " sec"
                If Stn_DIO(station, isStorageDrainSol).Value Then Stn_OutDigital station, isStorageDrainSol, cOFF
                If Stn_DIO(station, isFuelPumpMotor).Value Then Stn_OutDigital station, isFuelPumpMotor, cOFF
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If Now() > FstControl(station).Step_Time Then
                    FstControl(station).Step = 16              ' Continue
                End If
             
            Case 16
                FstControl(station).Message = "Check Level"
                flag = Not Stn_DIO(station, isStorageLowLevelLS).Value And (Stn_AIO(station, asStorageTankLevel).EUValue <= StationCfg_ADF(station, 1).FstDrainShutOff)
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If flag Then
                    ' Tank is Empty
                    FstControl(station).Step_Time = Now() + TimeSerial(0, 0, 2)
                    FstControl(station).Step = 19
                Else
                    FstControl(station).Step = 11       ' Resume Draining
                End If
          
            Case 19
                FstControl(station).Message = "Drain Done"
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                ' is brief delay done??
                If Now() > FstControl(station).Step_Time Then
                    FstControl(station).Mode = 0
                    FstControl(station).Step = 0
                    FstControl(station).ButtonVisible_Drain = True
                    FstControl(station).ButtonVisible_Fill = True
                    FstControl(station).ButtonVisible_Stop = False
                End If
                   
        
            Case 21
                FstControl(station).Message = "Fill On"
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If Stn_DIO(station, isStorageHiHiLevelLS).Value Then
                    ' done; already full
                    FstControl(station).Step = 29
                Else
                    ' begin filling
                    FstControl(station).ButtonVisible_Drain = False
                    FstControl(station).ButtonVisible_Fill = False
                    FstControl(station).ButtonVisible_Stop = True
                    Stn_OutDigital station, isStorageFillSol, cON               ' Open Fill Valve
                    Stn_OutDigital station, isStorageFillRequest, cON           ' Turn FillRequest ON
                    ' setup timeout for Fill
                    FstControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).FstFillTimeout)
                    FstControl(station).Step = 22
                End If
      
            Case 22
                FstControl(station).Message = "Wait High (" & Format(StationCfg_ADF(station, 1).FstFillShutOff, "###0.0") & " %) Level"
                flag = IIf((Stn_AIO(station, asStorageTankLevel).EUValue >= StationCfg_ADF(station, 1).FstFillShutOff), True, False)
                If Not Stn_DIO(station, isStorageFillSol).Value Then Stn_OutDigital station, isStorageFillSol, cON              ' Open Fill Valve
                If Not Stn_DIO(station, isStorageFillRequest).Value Then Stn_OutDigital station, isStorageFillRequest, cON      ' Turn FillRequest ON
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If (flag Or (Now() > FstControl(station).Step_Time)) Then
                    If flag Then
                        ' Continue
                        FstControl(station).Step = 23
                    Else
                        ' Timeout
                        FstControl(station).Step = 92
                    End If
                End If
      
            Case 23
                FstControl(station).Message = "Fill Off"
                Stn_OutDigital station, isStorageFillSol, cOFF
                Stn_OutDigital station, isStorageFillRequest, cOFF
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                FstControl(station).Step_Time = Now() + TimeSerial(0, 0, StationCfg_ADF(station, 1).FstFillDelay)
                FstControl(station).Step = 24
      
            Case 24
                temptime = FstControl(station).Step_Time - Now()
                tempSec = (60 * Minute(temptime)) + Second(temptime)
                FstControl(station).Message = "Wait Fill Delay - " & Format(tempSec, "##0") & " sec"
                If Stn_DIO(station, isStorageFillSol).Value Then Stn_OutDigital station, isStorageFillSol, cOFF
                If Stn_DIO(station, isStorageFillRequest).Value Then Stn_OutDigital station, isStorageFillRequest, cOFF
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                If Now() > FstControl(station).Step_Time Then
                    FstControl(station).Step = 25
                End If
      
            Case 25
                FstControl(station).Message = "Check Level"
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                flag = (Stn_AIO(station, asStorageTankLevel).EUValue >= StationCfg_ADF(station, 1).FstFillShutOff) Or Stn_DIO(station, isStorageHiHiLevelLS).Value
                If flag Then
                    ' Tank is Filled
                    FstControl(station).Step_Time = Now() + TimeSerial(0, 0, 2)
                    FstControl(station).Step = 29
                Else
                    ' Resume Filling
                    FstControl(station).Step = 21
                End If
      
            Case 29
                FstControl(station).Message = "Fill Cycle Done"
                FstControl(station).LevelSP = Stn_AIO(station, asStorageTankLevel).EUValue
                ' is brief delay done??
                If Now() > FstControl(station).Step_Time Then
                    FstControl(station).Mode = 0
                    FstControl(station).Step = 0
                    FstControl(station).ButtonVisible_Drain = True
                    FstControl(station).ButtonVisible_Fill = True
                    FstControl(station).ButtonVisible_Stop = False
                End If
      
      
            Case 90
                FstControl(station).Message = "Abort-StnStopped"
                FstControl(station).Mode = 0
                FstControl(station).Step = 0
                FstControl(station).StepBeforePause = 0
                Stn_OutDigital station, isFuelPumpMotor, cOFF
                Stn_OutDigital station, isStorageDrainSol, cOFF
                Stn_OutDigital station, isStorageFillSol, cOFF
                Stn_OutDigital station, isStorageFillRequest, cOFF
      
            Case 91
                FstControl(station).Message = "Abort-Too Long Drain"
                Stn_OutDigital station, isFuelPumpMotor, cOFF
                Stn_OutDigital station, isStorageDrainSol, cOFF
                FstControl(station).ButtonVisible_Drain = True
                FstControl(station).ButtonVisible_Fill = False
                FstControl(station).ButtonVisible_Stop = True
      
            Case 92
                FstControl(station).Message = "Abort-Too Long Fill"
                Stn_OutDigital station, isStorageFillSol, cOFF
                Stn_OutDigital station, isStorageFillRequest, cOFF
                FstControl(station).ButtonVisible_Drain = True
                FstControl(station).ButtonVisible_Fill = True
                FstControl(station).ButtonVisible_Stop = True
            
            Case 95
                FstControl(station).Message = "Abort-StnPaused"
                Stn_OutDigital station, isFuelPumpMotor, cOFF
                Stn_OutDigital station, isStorageDrainSol, cOFF
                Stn_OutDigital station, isStorageFillSol, cOFF
                Stn_OutDigital station, isStorageFillRequest, cOFF
                FstControl(station).ButtonVisible_Drain = True
                FstControl(station).ButtonVisible_Fill = True
                FstControl(station).ButtonVisible_Stop = False
    
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

Sub UpdateLeakInputs(ByVal iStation As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 984

Dim T1 As Double

    If USINGSIMULATION Then
        ' simulation only
        Stn_AIO(iStation, asNitrogenFlow).EUValue = Cfg_LeakTest.InitialN2Flow
        Stn_AIO(iStation, asLtInletPress).EUValue = 13.25
        Stn_AIO(iStation, asLtN2Temp).EUValue = Com_AIO(acAmbTempSensor).EUValue
    End If
    
    ' LeakTest-Calculation Input Values
    ' PIN = Inlet pressure
    ' PATM = Atmospneric Pressure
    QN2 = Stn_AIO(iStation, asNitrogenFlow).EUValue * 0.000016666    ' l/m to m3/s
    Patm = Com_AIO(acAmbBaroSensor).EUValue * 0.1                    ' mbar to kPa
' ###  TEST
'    Pin = Patm + Stn_AIO(iStation, asLtInletPress).EUValue          ' kPa (gauge to absolute)
    
    Pin = Patm + (Stn_AIO(iStation, asLtInletPress).EUValue) * 0.1         ' kPa (gauge to absolute)

'  #####  TEST
'    T1 = IIf(USINGC, Stn_AIO(iStation, asLtN2Temp).EUValue, DegFtoC(Stn_AIO(iStation, asLtN2Temp).EUValue))  ' degC

    T1 = DegFtoC(Abs(Stn_AIO(iStation, asLtN2Temp).EUValue))  ' degC

    TN2 = T1 + 273.15                                            ' degC to degK
    
    With CurrLT2_Data
        .InPress = Pin
        .AtmPress = Patm
        .NitFlow = QN2
        .NitTemp = TN2
    End With

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
DeffCalcMsg = "calculation NOT ok"
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

Sub CalcEffLeakDia(ByVal iStation As Integer)
'
' Calculate Effective Leak Diameter
'   per 40 CFR 1066.985
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 985

Dim dVal0 As Double
Dim dVal1 As Double
Dim dVal2 As Double
Dim dVal3 As Double
Dim dVal4 As Double

    dVal0 = (Pin - Patm) * (Pin + Patm)
    If (dVal0 > 0#) Then
        dVal1 = dVal0 / (SGN2 * TN2)
        dVal2 = Sqr(dVal1)
        If (dVal2 <> 0#) Then
            dVal3 = (QN2 / dVal2)
            dVal4 = dVal3 ^ CDbl(0.5057)
            
'  ####  TEST
' Deff = (dVal4 * CDbl(7.844))
            Deff = (dVal4 * CDbl(7.844)) * 0.1
            DeffCalcFlag = True
            DeffCalcMsg = "calculation ok"
        Else
            DeffCalcFlag = False
            DeffCalcMsg = "divide by zero; calc aborted"
        End If
    Else
        DeffCalcFlag = False
        DeffCalcMsg = "Patm > Pin; calc aborted"
    End If
    
    With CurrLT2_Data
        .ClkTime = Now()
        .SecTimer = Timer
        .EffDia = Deff
    End With

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
DeffCalcMsg = "calculation NOT ok"
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

Sub Common_Valves()
'
'   Controls Common Valves to match needs of stations
'
'   14 March 2005
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 373

' Paused or Normal Operation ?
If Pause_Alarm = NOTPAUSED Then
    ' Normal Operation
    ' Butane valve is always on
    If systemhasBUTANE And Not Com_DIO(icButaneShutoffSol).Value Then Com_OutDigital icButaneShutoffSol, cON
Else
    ' System is Paused
    ' Shutoff Butane Valve
    If systemhasBUTANE And Com_DIO(icButaneShutoffSol).Value Then Com_OutDigital icButaneShutoffSol, cOFF
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

Sub LiveFuel_Controller()
'
'       Control the Live Fuel Tank(s) Valves, Heaters, etc.
'
'
Dim iStn As Integer
Dim delta, Target As Single

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 3333
Dim iShift As Integer
Dim sMsg As String


    '
    '   ADF SEQUENCE CONTROL
    '

    ChgErrModule 8, 3334
    
    For iStn = 1 To LAST_STN
    
'       Only applies to Live Fuel Stations
        If ((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)) Then
        
            iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
            
            ' any request to set LiveFuel State to "OK" ???
            If AdfControl(iStn).SetOkRequest Then
                AdfControl(iStn).SetOkRequest = False
                AdfControl(iStn).LiveFuelState = fuelOK
                AdfControl(iStn).LiveFuelDensityOkCnt = 0
                AdfControl(iStn).LiveFuelDensityWeakCnt = 0
                AdfControl(iStn).LiveFuelDensityDeadCnt = 0
                sMsg = "LiveFuel Refill Complete; fuel state is set to OK"
                Write_ELog "Station #" & Format(iStn, "0") & " " & sMsg
                If Len(StationControl(iStn, iShift).DBFile) > 0 Then OOT_Write iStn, iShift, sMsg
            End If
            
            If (AdfControl(iStn).LiveFuel) Then
                Select Case StationControl(iStn, iShift).Mode
                    Case VBPURGE, VBPURGEWAIT, VBPOSTPURGE, VBPURGECONT
                        AdfControl(iStn).Enable = True
                        AdfControl(iStn).ManScreen_Enable = True
                    Case VBCOURSEWAIT, VBCOURSEPAUSE
                        AdfControl(iStn).Enable = True
                        AdfControl(iStn).ManScreen_Enable = True
                    Case VBPAUSEOOT
                        AdfControl(iStn).Enable = True
                        AdfControl(iStn).ManScreen_Enable = False
                    Case VBGASPAUSE
                        AdfControl(iStn).Enable = True
                        AdfControl(iStn).ManScreen_Enable = True
                    Case VBFIDPAUSE
                        AdfControl(iStn).Enable = True
                        AdfControl(iStn).ManScreen_Enable = False
                    Case VBPRELOAD
                        AdfControl(iStn).Enable = False
                        AdfControl(iStn).ManScreen_Enable = False
                    Case VBLOAD
                        Select Case LoadControl(iStn, iShift).Phase
                            Case LoadPrep, LoadStarting
                                '
                                ' Preliminaries
                                ' Getting Started
                                '
                                AdfControl(iStn).Enable = False
                                AdfControl(iStn).ManScreen_Enable = False
                                '
                            Case LoadLoading, LoadComplete, LoadStopping, LoadPause
                                '
                                ' LOAD CYCLE    "are we done yet?"
                                ' turn off load cycle mfc's
                                ' after delay, turn off load cycle valves
                                ' after scale values settle, end this load cycle
                                '
                                AdfControl(iStn).Enable = IIf(((AdfControl(iStn).LiveFuelState = fuelDead) Or AdfControl(iStn).Enable), True, False)
'                                AdfControl(iStn).Enable = IIf((AdfControl(iStn).LiveFuelState = fuelDead), True, False)
                                AdfControl(iStn).ManScreen_Enable = AdfControl(iStn).Enable
                                '
                            Case Else
                                AdfControl(iStn).Enable = False
                                AdfControl(iStn).ManScreen_Enable = False
                                '
                        End Select
                    Case VBPOSTLOAD
                        AdfControl(iStn).Enable = IIf(((AdfControl(iStn).LiveFuelState = fuelDead) Or AdfControl(iStn).Enable), True, False)
'                        AdfControl(iStn).Enable = IIf((AdfControl(iStn).LiveFuelState = fuelDead), True, False)
                        AdfControl(iStn).ManScreen_Enable = AdfControl(iStn).Enable
                    Case Else
                        AdfControl(iStn).Enable = False
                        AdfControl(iStn).ManScreen_Enable = False
                End Select
                
                ' Need to Refill the LiveFuel Tank ?
                If (AnyMoreLiveFuelLoads(iStn, iShift) Or (AdfControl(iStn).LiveFuelState = fuelDead)) Then
                    If (Not AdfControl(iStn).InitialFill_Complete) _
                        Or (StationControl(iStn, iShift).LiveFuelCycleCount >= AdfControl(iStn).LiveFuelChgFreq) _
                        Or ((AdfControl(iStn).LiveFuelState = fuelWeak) And StationControl(iStn, iShift).Mode <> VBLOAD) _
                        Or (AdfControl(iStn).LiveFuelState = fuelDead) Then
                        ' not already requested ??
                        If Not AdfControl(iStn).RefillRequest And Not AdfControl(iStn).ReadyForRefill Then
                            ' request a refill
                            AdfControl(iStn).RefillRequest = True
                            AdfControl(iStn).ReadyForRefill = True
                            AdfControl(iStn).ReadyForLoad = False
                            sMsg = "none"
                            If (StationControl(iStn, iShift).LiveFuelCycleCount >= AdfControl(iStn).LiveFuelChgFreq) Then sMsg = "Tank refill requested due to Fuel Change Frequency"
                            If (AdfControl(iStn).LiveFuelState = fuelDead) Then sMsg = "Tank refill requested due to Dead fuel"
                            If ((AdfControl(iStn).LiveFuelState = fuelWeak) And StationControl(iStn, iShift).Mode <> VBLOAD And AnyMoreLiveFuelLoads(iStn, iShift)) Then sMsg = "Tank refill requested due to Weak fuel"
                            If (sMsg <> "none") Then
                                Write_ELog "ADF Station #" & Format(iStn, "#0") & " " & sMsg
                                If StationControl(iStn, iShift).TestTimerIsRunning Then Write_JLog iStn, iShift, sMsg
                            End If
                        End If
                    End If
                End If
                
                ' Need to Initiate a LiveFuel Refill Cycle ?
                If AdfControl(iStn).ReadyForRefill And AdfControl(iStn).RefillRequest And AdfControl(iStn).Enable Then
                        
                    ' Normal; Initiate a Refill
                    AdfControl(iStn).RefillRequest = False
                    If (StationControl(iStn, iShift).Mode = VBLOAD) Then frmStnDetail.StationPause "ADF"
                    If AdfControl(iStn).LiveFuelChgAuto Then
                        AdfControl(iStn).Mode = 2
                        AdfControl(iStn).Step = 0
                    Else
                        If (STN_INFO(iStn).ADF_DEF.hasADF_WaterBath And StationRecipe(iStn, iShift).ADF_Heater) Then
                            If LoadControl(iStn, iShift).WaterBathTempOK Then
                                Select Case StationConfig(iStn, iShift).WaterBathControl
                                    Case wbDirect
                                        AdfControl(iStn).Message = "Waiting for WaterBath Temp"
                                    Case wbFuelTemp
                                        AdfControl(iStn).Message = "Waiting for Fuel Temp"
                                    Case wbVaporTemp
                                        AdfControl(iStn).Message = "Waiting for Vapor Temp"
                                End Select
                                AdfControl(iStn).Mode = 0
                                AdfControl(iStn).Step = 0
                            Else
                                AdfControl(iStn).Message = "Waiting for Manual Fuel Change"
                                AdfControl(iStn).Mode = 0
                                AdfControl(iStn).Step = 0
                                AdfControl(iStn).ButtonVisible_Done = True
                            End If
                        Else
                            AdfControl(iStn).Message = "Waiting for Manual Fuel Change"
                            AdfControl(iStn).Mode = 0
                            AdfControl(iStn).Step = 0
                            AdfControl(iStn).ButtonVisible_Done = True
                        End If
                    End If
                    If (StationControl(iStn, iShift).Mode = VBLOAD) Then StationControl(iStn, iShift).Mode = VBGASPAUSE
                                       
                End If
                               
                If (STN_INFO(iStn).ADF_TANKTYPE <> 90) Then
                   ' AutoDrainFill Sequence Control
                    ADF_Sequence CInt(iStn)
                    ' Fuel Storage Drain & Fill Sequence Control
                    FST_Sequence CInt(iStn)
                            
                ElseIf (STN_INFO(iStn).ADF_TANKTYPE = 90) Then
                
                    ' WaterBath Only
                    WaterBath_Controller CInt(iStn)
                End If
            End If
     
        Else
            
            ' Not a Live Fuel Station
            AdfControl(iStn).Enable = False
            AdfControl(iStn).Heater_Enable = False
            AdfControl(iStn).ManScreen_Enable = False
            FstControl(iStn).Enable = False
                
        End If
        
    Next iStn
    

    '
    '   HEATER CONTROL
    '
    ChgErrModule 8, 3335
    ' time to check the Tank Temp(s) ??
    If Now() > ADF_HeaterCheckTime + TimeSerial(0, 0, 2) Then
        ' do a check
        ADF_HeaterCheckTime = Now()
        ' check all LiveFuel Stations that have a Heater
        For iStn = 1 To LAST_STN
            If (AdfControl(iStn).AdfDefinition.hasLIVEFUEL) Then
                If (AdfControl(iStn).AdfDefinition.hasADF_Heater) Then
                    ADF_Heater CInt(iStn)
                ElseIf (AdfControl(iStn).AdfDefinition.hasADF_WaterBath) Then
'                    ADF_WaterBath CInt(iStn)
                End If
            End If
        Next iStn
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

Sub PAS_LocalControl()
'
'   Local Control of Purge Air Temperature and Humidity
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 3130
Dim errFlag As Boolean
Dim rdyFlag As Boolean
Dim rdyInt As Integer
Dim reqFlag As Boolean

    ' *********
    ' PAS READY
    ' *********
    '   Check PAS Temperature
    PAS_Check pasTEMPERATURE
    '   Check PAS Moisture
    PAS_Check pasMOISTURE
    '   PAS Ready Output
    errFlag = IIf((PAS_INFO(pasTEMPERATURE).timeOut Or PAS_INFO(pasMOISTURE).timeOut), True, False)
    rdyFlag = IIf((PAS_INFO(pasTEMPERATURE).Ok And PAS_INFO(pasMOISTURE).Ok), True, False)
    rdyInt = IIf(rdyFlag, cYES, cNO)
'    reqFlag = Com_DIO(icPASPowerOnIn).Value Or MasterPagData.ReqIn Or LocalPagControl.ReqIn
    reqFlag = PAG_Request Or MasterPagData.ReqIn
    ' PAG MODE
    If (Not Com_DIO(icPASPowerOnIn).Value) Then
        MasterPagData.Status = "SOFF"
    ElseIf (errFlag) Then
        MasterPagData.Status = "SERR"
    ElseIf (Not reqFlag) Then
        MasterPagData.Status = "SIDL"
    ElseIf (Not rdyFlag) Then
        MasterPagData.Status = "STBY"
    Else
        MasterPagData.Status = "SRDY"
    End If
    ' PAG Current Values
    MasterPagData.Temperature = PATemp
    MasterPagData.Humidity = PAHum
    MasterPagData.Moisture = PAMoisture
    MasterPagData.TempSP = SysConfig.Temp_Target
    MasterPagData.MoistSP = SysConfig.Moisture_Target
    MasterPagData.TempTol = SysConfig.Tol_Temp
    MasterPagData.MoistTol = SysConfig.Tol_Moisture
    MasterPagData.RdyOut = IIf(rdyFlag, True, False)
    Com_OutDigital icPASReadyOut, rdyInt
    
    ' ***********************
    ' PAS TEMPERATURE CONTROL
    ' ***********************
    '   PAS Heater On/Off Controller
    Controller_OnOff pasTEMPERATURE
    ' ********************
    ' PAS MOISTURE CONTROL
    ' ********************
    '   PAS Moisturizer PID Controller
    Controller_PID pasMOISTURE
    
    ' ***********************
    '   Optionally, write PAS values to the zLog
    ' ***********************
    If Not NotDebugPAS Then
    
        If PAS_INFO(pasTEMPERATURE).LastUpdate > Debug_ZlogPAS_LastUpdate _
                Or _
           PAS_INFO(pasMOISTURE).LastUpdate > Debug_ZlogPAS_LastUpdate Then
           
                ' normal interval
                ' write to the PAS_Log in the zLog db file
                Write_Zlog_PAS "Normal Update"
                
        ElseIf PAS_INFO(pasTEMPERATURE).LastUpdate < 1 _
                Or _
           PAS_INFO(pasMOISTURE).LastUpdate < 1 Then
           
                ' it is just after midnight
                ' write to the PAS_Log in the zLog db file
                Write_Zlog_PAS "Rollover Update"
                
        ElseIf (Debug_ZlogPAS_LastUpdate - PAS_INFO(pasTEMPERATURE).LastUpdate) > 2 _
                Or _
           (Debug_ZlogPAS_LastUpdate - PAS_INFO(pasMOISTURE).LastUpdate) > 2 Then
           
                ' we are out of sync; time to catchup
                ' write to the PAS_Log in the zLog db file
                Dim txt As String
                txt = "Catchup Update to zLog_PAS at " & Now
                txt = txt & "   (Last Log Update = " & Format(Debug_ZlogPAS_LastUpdate, "####0.000")
                txt = txt & " ;Last Temp Update = " & Format(PAS_INFO(pasTEMPERATURE).LastUpdate, "####0.000")
                txt = txt & " ;Last Moisture Update = " & Format(PAS_INFO(pasMOISTURE).LastUpdate, "####0.000")
                txt = txt & ")"
                Write_ELog txt
                Write_Zlog_PAS "Catchup Update"
                
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

Sub PAS_Check(ByVal Idx As Integer)
'
'   Checks PAS Parameter for within limits (or not)
'
Dim NowTmr, DeltaTmr As Double
Dim lolimit, hilimit, currVal As Single
Dim sStr As String
Dim iStn, iShift As Integer
Dim cntrlOn As Boolean
Dim reqFlag As Boolean

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 3131

    ' how long since last check
    NowTmr = IIf(Timer > PAS_INFO(Idx).LastUpdate, Timer, Timer + 86400)
    DeltaTmr = NowTmr - PAS_INFO(Idx).LastUpdate
    ' at least 1 second between checks
    If DeltaTmr > 1# Then
        ' if within limits for desired duration then OK is true
        PAS_INFO(Idx).Ok = IIf((PAS_INFO(Idx).Duration < PAS_INFO(Idx).DurationTarget), False, True)
        ' enable timeout if PAS is running
        reqFlag = PAG_Request Or MasterPagData.ReqIn
        If (reqFlag And Com_DIO(icPASisRunningIn).Value) Then
'        If Com_DIO(icPASPowerOnIn).Value Then
'            If Com_DIO(icPASisRunningIn).Value Then
            Select Case Idx
                Case pasTEMPERATURE
                    ' Temperature
                    currVal = PATemp
                    lolimit = SysConfig.Temp_Target - SysConfig.Tol_Temp
                    hilimit = SysConfig.Temp_Target + SysConfig.Tol_Temp
                Case pasMOISTURE
                    ' Moisture
                    currVal = PAMoisture
                    lolimit = SysConfig.Moisture_Target - SysConfig.Tol_Moisture
                    hilimit = SysConfig.Moisture_Target + SysConfig.Tol_Moisture
            End Select
            If (currVal > lolimit And currVal < hilimit) Then
                ' currently within limits
                ' add delta time to duration (also don't let duration count overflow)
                PAS_INFO(Idx).Duration = PAS_INFO(Idx).Duration + DeltaTmr
                If PAS_INFO(Idx).Duration > (10 * PAS_INFO(Idx).DurationTarget) Then PAS_INFO(Idx).Duration = (10 * PAS_INFO(Idx).DurationTarget)
                ' check for timeout if controller is running
                If PID_INFO(Idx).Enable And Not PID_INFO(Idx).Inhibit Then
                    If ((PAS_INFO(Idx).timeOut) And (Not last_INFO(Idx).timeOut)) Then
                        Select Case Idx
                            Case pasTEMPERATURE
                                ' Temperature
                                If USINGC Then sStr = " deg C"
                                If USINGF Then sStr = " deg F"
                                sStr = "PAS Temperature of " & Format(SysConfig.Temp_Target, "##0.0") & sStr & " is now within tolerance limits"
                            Case pasMOISTURE
                                ' Moisture
                                If USINGMoist_RH Then sStr = " % rH"
                                If USINGMoist_Grains Then sStr = " grains/lb"
                                sStr = "PAS Moisture of " & Format(SysConfig.Moisture_Target, "##0.0") & sStr & " is now within tolerance limits"
                        End Select
                        Write_ELog sStr
                        For iStn = 1 To LAST_STN
                            For iShift = 1 To NR_SHIFT
                                If StationControl(iStn, iShift).Mode <> VBIDLE _
                                    And StationControl(iStn, iShift).Mode <> VBIDLEWAITING _
                                    And StationControl(iStn, iShift).Mode <> VBCOMPLETE Then
                                        ALM_Write CInt(iStn), CInt(iShift), sStr
                                End If
                            Next iShift
                        Next iStn
                    End If
                    PAS_INFO(Idx).timeOut = False
                    PAS_INFO(Idx).TimeOutDuration = 0#
                Else
                    ' controller not enabled or is inhibited
                    PAS_INFO(Idx).timeOut = False
                    PAS_INFO(Idx).TimeOutDuration = 0#
                End If
            Else
                ' Not currently within limits
                PAS_INFO(Idx).Duration = 0#
'                PAS_INFO(Idx).Ok = False
                ' check for timeout if controller is running
                If PID_INFO(Idx).Enable And Not PID_INFO(Idx).Inhibit Then
                    ' add delta time to timeout duration (also don't let timeout duration count overflow)
                    PAS_INFO(Idx).TimeOutDuration = PAS_INFO(Idx).TimeOutDuration + DeltaTmr
                    If PAS_INFO(Idx).TimeOutDuration > (10 * PAS_INFO(Idx).TimeOutTarget) Then PAS_INFO(Idx).TimeOutDuration = (10 * PAS_INFO(Idx).TimeOutTarget)
                    ' if outside limits for too long then Timeout is true
                    If (PAS_INFO(Idx).TimeOutDuration > PAS_INFO(Idx).TimeOutTarget) Then
                        If ((Not PAS_INFO(Idx).timeOut) And (last_INFO(Idx).timeOut)) Then
                            Select Case Idx
                                Case pasTEMPERATURE
                                    ' Temperature
                                    If USINGC Then sStr = " deg C"
                                    If USINGF Then sStr = " deg F"
                                    sStr = "PAS Temperature Timeout; Failed to reach " & Format(SysConfig.Temp_Target, "##0.0") & sStr
                                Case pasMOISTURE
                                    ' Moisture
                                    If USINGMoist_RH Then sStr = " % rH"
                                    If USINGMoist_Grains Then sStr = " grains/lb"
                                    sStr = "PAS Moisture Timeout; Failed to reach " & Format(SysConfig.Moisture_Target, "##0.0") & sStr
                            End Select
                            Write_ELog sStr
                            For iStn = 1 To LAST_STN
                                For iShift = 1 To NR_SHIFT
                                    If StationControl(iStn, iShift).Mode <> VBIDLE _
                                        And StationControl(iStn, iShift).Mode <> VBIDLEWAITING _
                                        And StationControl(iStn, iShift).Mode <> VBCOMPLETE Then
                                        ALM_Write CInt(iStn), CInt(iShift), sStr
                                    End If
                                Next iShift
                            Next iStn
                            PAS_INFO(Idx).timeOut = True
                        End If
                    Else
                        PAS_INFO(Idx).timeOut = False
                    End If
                Else
                    ' controller not enabled or is inhibited
                    PAS_INFO(Idx).timeOut = False
                    PAS_INFO(Idx).TimeOutDuration = 0#
                End If
            End If
        Else
            ' local control is off
            PAS_INFO(Idx).timeOut = False
            PAS_INFO(Idx).TimeOutDuration = 0#
'            PAS_INFO(Idx).Ok = False
            PAS_INFO(Idx).Duration = 0#
        End If
        ' remember
        last_INFO(Idx) = PAS_INFO(Idx)
        ' LastUpdate = Timer
        PAS_INFO(Idx).LastUpdate = IIf((NowTmr < 86400), NowTmr, NowTmr - 86400)
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

Sub Controller_OnOff(ByVal Idx As Integer)
'
'   On/Off Controlller
'
'       this sub only sets the Output bit in the PID Control Block
'       the actual physical output is set elsewhere
'
Dim CurOnOff As PIDcontrol
Dim CntrlTrue, CntrlFalse As Boolean
Dim NowTmr, DeltaTmr As Double
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 3132

    ' Copy Control Block Values
    CurOnOff = PID_INFO(Idx)
    ' Update Controller Input Values
    Select Case Idx
        Case pasTEMPERATURE
            ' PAS Temperature
            ' enable
            CurOnOff.Enable = Com_DIO(icPASPowerOnIn).Value
            ' inhibit
            '   inhibit if "PASPower Input is Off" OR "TempTimeout" OR "MoistureTimeout"
            CurOnOff.Inhibit = IIf(Not Com_DIO(icPASisRunningIn).Value, True, _
                                IIf(PAS_INFO(pasTEMPERATURE).timeOut, True, _
                                IIf(PAS_INFO(pasMOISTURE).timeOut, True, False)))
            CurOnOff.PV = PATemp
            CurOnOff.SP = SysConfig.Temp_Target
            ' error
            CurOnOff.Er = CurOnOff.PV - CurOnOff.SP
            ' On & Off Duty Cycle
            '   On Duty in seconds
            CurOnOff.OnDuty = 2
            '   Off Duty in seconds
            If USINGC Then
                ' temperature units are deg Celsius
                Select Case CurOnOff.Er
                    Case Is > -0.5
                        ' Very Near the SP
                        CurOnOff.OffDuty = 11
                    Case Is > -1#
                        ' Near the SP
                        CurOnOff.OffDuty = 9
                    Case Is > -2#
                        ' "Sort Of" Near the SP
                        CurOnOff.OffDuty = 7
                    Case Else
                        ' Not Near the SP
                        CurOnOff.OffDuty = 6
                End Select
            End If
            If USINGF Then
                ' temperature units are deg Fahrenheit
                Select Case CurOnOff.Er
                    Case Is > -0.9
                        ' Very Near the SP
                        CurOnOff.OffDuty = 11
                    Case Is > -1.8
                        ' Near the SP
                        CurOnOff.OffDuty = 9
                    Case Is > -3.6
                        ' "Sort Of" Near the SP
                        CurOnOff.OffDuty = 7
                    Case Else
                        ' Not Near the SP
                        CurOnOff.OffDuty = 6
                End Select
            End If
            ' adjust Duty Cycle for higher Moisture Loads
            If PID_INFO(pasMOISTURE).Enable = True _
                And PID_INFO(pasMOISTURE).Inhibit = False Then
                Select Case PID_INFO(pasMOISTURE).SP
                    Case Is > 35
                        CurOnOff.OffDuty = 0.65 * CurOnOff.OffDuty
                        CurOnOff.OnDuty = 1.15 * CurOnOff.OnDuty
                    Case Is > 25
                        CurOnOff.OffDuty = 0.85 * CurOnOff.OffDuty
                        CurOnOff.OnDuty = 1.05 * CurOnOff.OnDuty
                    Case Else
                        CurOnOff.OffDuty = CurOnOff.OffDuty
                        CurOnOff.OnDuty = CurOnOff.OnDuty
                End Select
            End If
            ' adjust Duty Cycle per Config settings
            CurOnOff.OffDuty = CurOnOff.OffDutyMult * CurOnOff.OffDuty
            CurOnOff.OnDuty = CurOnOff.OnDutyMult * CurOnOff.OnDuty
        Case Else
            ' not defined
            CurOnOff.Enable = False
            CurOnOff.Inhibit = True
            CurOnOff.PV = 1
            CurOnOff.SP = 1
            CurOnOff.Er = 0
    End Select
    ' Reverse Acting Output?
    CntrlFalse = IIf(CurOnOff.Rev, True, False)
    CntrlTrue = IIf(CurOnOff.Rev, False, True)
        
    If Pause_Alarm = SYSTEMPAUSED Then
        ' System is Paused, Turn Off All Outputs
        CurOnOff.Output = CntrlFalse
    Else
        ' Normal Operation (system is not paused)
        ' Enabled and Not Inhibited?
        If CurOnOff.Enable And Not CurOnOff.Inhibit Then
            ' How long since last update
            NowTmr = IIf(Timer > CurOnOff.LastUpdate, Timer, Timer + 86400)
            DeltaTmr = NowTmr - CurOnOff.LastUpdate
            CurOnOff.LastUpdate = Timer
            ' update Heater On & Off timers
            Select Case CurOnOff.Output
                Case CntrlTrue
                    CurOnOff.OnTimer = IIf(CurOnOff.OnTimer > 10 * CurOnOff.OnDuty, CurOnOff.OnTimer, CurOnOff.OnTimer + DeltaTmr)
                    CurOnOff.OffTimer = 0
                Case CntrlFalse
                    CurOnOff.OffTimer = IIf(CurOnOff.OffTimer > 10 * CurOnOff.OffDuty, CurOnOff.OffTimer, CurOnOff.OffTimer + DeltaTmr)
                    CurOnOff.OnTimer = 0
            End Select
            '
            '   ON/OFF CONTROL
            '
            '   If pv is above the OffLimit then the Output is Always Off
            '   If pv is between the Off & On Limits then the Output is On/Off per Duty Cycle
            '   If pv is below the OnLimit then the Output is Always On
            '
            '   notes: Error = (PV - SP) & everything is relative to the SP
            '
            Select Case CurOnOff.Er
                Case Is > CurOnOff.OffLimitDelta
                    ' Temp > OffLimit
                    ' Heater Off
                    CurOnOff.Output = CntrlFalse
                Case Is < CurOnOff.OnLimitDelta
                    ' Temp < OnLimit
                    ' Heater On at Max Duty Cycle
                    CurOnOff.Output = CntrlTrue
                Case Else
                    ' OffLimit > Temp > OnLimit
                    ' Heater On/Off per On/Off Duty Cycle
                    Select Case CurOnOff.Output
                        Case True
                            ' if on, stay on for onduty
                            CurOnOff.Output = IIf(CurOnOff.OnTimer < CurOnOff.OnDuty, CntrlTrue, CntrlFalse)
                        Case False
                            ' if off, stay off for offduty
                            CurOnOff.Output = IIf(CurOnOff.OffTimer < CurOnOff.OffDuty, CntrlFalse, CntrlTrue)
                    End Select
            End Select
        Else
            ' Turn Output Off if Just-Lost-Enable Or Inhibit is True
            If PID_INFO(Idx).Enable Then CurOnOff.Output = CntrlFalse
            If CurOnOff.Inhibit Then CurOnOff.Output = CntrlFalse
        End If
    End If
    ' Update Controller Output
    Select Case Idx
        Case pasTEMPERATURE
            ' PAS Temperature    (note: CurOnOff.Output is a Boolean, cON & cOFF are integers)
            If CurOnOff.Enable And Not CurOnOff.Inhibit Then
                ' Enabled and Not Inhibited
                If Com_DIO(icPASHeaterSSR).Value <> CurOnOff.Output Then _
                    Com_OutDigital icPASHeaterSSR, IIf(CurOnOff.Output, cON, cOFF)
            ElseIf CurOnOff.Inhibit Then
                ' Inhibited
                If Com_DIO(icPASHeaterSSR).Value <> CurOnOff.Output Then _
                    Com_OutDigital icPASHeaterSSR, IIf(CurOnOff.Output, cON, cOFF)
            Else
                ' Overheat Protection
                If CurOnOff.Er > CurOnOff.OffLimitDelta Then _
                    Com_OutDigital icPASHeaterSSR, cOFF
            End If
        Case Else
            ' not defined
    End Select
    ' Update Control Block
    PID_INFO(Idx) = CurOnOff

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

Sub Controller_PID(ByVal iController As Integer)
'
'   PID Controller
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 31330
Dim iStn As Integer
Dim iShift As Integer
Dim Idx As Integer
Dim outEU As Single
Dim outInit As Single
Dim outmax As Single
Dim outputMax As Single
Dim outputMin As Single
Dim cumMax As Single
Dim cumMin As Single
Dim outRaw As Single
Dim iMfc As Integer
Dim iFuncPV As Integer
Dim iFuncSP As Integer
Dim tmpSpan As Single
Dim NowTmr, DeltaTmr As Double
Dim DeltaI, NewOut, tempRH As Single
Dim CurPID As PIDcontrol
    
    CurPID = PID_INFO(iController)
    Select Case iController
        Case pasMOISTURE
ChgErrModule 8, 31331
            ' PAS Moisture Controller  (assume Range = 0 to 100 Grains/lb; i.e. EU = %)
            ' enable
            CurPID.Enable = Com_DIO(icPASPowerOnIn).Value
            ' inhibit
            '   inhibit if "PASPower Input is OFF" or "Temp.Timeout" or "Moisture.Timeout"
            CurPID.Inhibit = IIf(Not Com_DIO(icPASisRunningIn).Value, True, _
                                IIf(PAS_INFO(pasTEMPERATURE).timeOut, True, _
                                IIf(PAS_INFO(pasMOISTURE).timeOut, True, False)))
            ' process value
            CurPID.PV = PAMoisture
            ' set point
            CurPID.SP = SysConfig.Moisture_Target
            ' error  (note: Reverse Action switches polarity of Error)
            CurPID.Er = IIf(CurPID.Rev, (CurPID.SP - CurPID.PV), (CurPID.PV - CurPID.SP))
        Case wbSuperTemp
ChgErrModule 8, 31332
            ' WaterBath Supervisory Temperature Controller  (convert EU Range to %)
            For Idx = 1 To NR_STN
                If STN_INFO(Idx).ADF_DEF.hasADF_WaterBath Then iStn = Idx
            Next Idx
            iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
            Select Case StationConfig(iStn, iShift).WaterBathControl
                Case wbDirect
                    ' enable
                    CurPID.Enable = IIf(((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)), True, False)
                    ' inhibit
                    CurPID.Inhibit = True
                    ' process value
                    CurPID.PV = 1
                    ' set point
                    CurPID.SP = 1
                    ' error
                    CurPID.Er = 0
                    ' cumulative integral term
                    CurPID.CumI = 0
                    ' output
                    If USINGC Then
                        tmpSpan = WB_AIO.EuMax - WB_AIO.EuMin
                        outRaw = (StationRecipe(iStn, iShift).ADF_HeaterSP - WB_AIO.EuMin) / tmpSpan
                    ElseIf USINGF Then
                        tmpSpan = DegCtoF(WB_AIO.EuMax) - DegCtoF(WB_AIO.EuMin)
                        outRaw = (StationRecipe(iStn, iShift).ADF_HeaterSP - DegCtoF(WB_AIO.EuMin)) / tmpSpan
                    End If
                    CurPID.out = outRaw * CSng(100)
                Case wbFuelTemp
                    ' enable
                    CurPID.Enable = IIf(((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)), True, False)
                    ' inhibit
                    CurPID.Inhibit = False
                    ' process value
                    iFuncPV = asFuelTankTemp
                    CurPID.PV = Stn_AIO(iStn, iFuncPV).EUValue
                    ' set point
                    CurPID.SP = StationRecipe(iStn, iShift).ADF_HeaterSP
                    ' error  (note: Reverse Action switches polarity of Error)
                    CurPID.Er = IIf(CurPID.Rev, (CurPID.SP - CurPID.PV), (CurPID.PV - CurPID.SP))
                Case wbVaporTemp
                    ' enable
                    CurPID.Enable = IIf(((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)), True, False)
                    ' inhibit
                    CurPID.Inhibit = False
                    ' process value
                    iFuncPV = asFuelVaporTemp
                    CurPID.PV = Stn_AIO(iStn, iFuncPV).EUValue
                    ' set point
                    CurPID.SP = StationRecipe(iStn, iShift).ADF_HeaterSP
                    ' error  (note: Reverse Action switches polarity of Error)
                    CurPID.Er = IIf(CurPID.Rev, (CurPID.SP - CurPID.PV), (CurPID.PV - CurPID.SP))
            End Select
        Case stn1LoadRate To stn9LoadRate
ChgErrModule 8, 31333
            ' STN Loadrate Controller  (convert EU Range to %)
            iStn = iController - 10
            iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
            Select Case STN_INFO(iStn).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                     outmax = (SlpmToGramsPerHour(Stn_AIO(iStn, asButaneFlow).EuMax, (GramsPerLiter * STN_INFO(iStn).ButMfc2DensityMult))) * 0.95
                Case STN_ORVR2_TYPE
                    If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                        outmax = (SlpmToGramsPerHour(Stn_AIO(iStn, asButaneORVRFlow).EuMax, (GramsPerLiter * STN_INFO(iStn).ButMfc2DensityMult))) * 0.95
                    Else
                        outmax = (SlpmToGramsPerHour(Stn_AIO(iStn, asButaneFlow).EuMax, (GramsPerLiter * STN_INFO(iStn).ButMfc2DensityMult))) * 0.95
                    End If
                Case STN_LIVEFUEL_TYPE, STN_LIVEREG_TYPE
                    outmax = (SlpmToGramsPerHour(Stn_AIO(iStn, asLiveFuelVaporFlow).EuMax, LiveFuelVaporDensity)) * 0.95
                Case STN_LIVEORVR2_TYPE
                    If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                        outmax = (SlpmToGramsPerHour(Stn_AIO(iStn, asLiveFuelVaporORVRFlow).EuMax, LiveFuelVaporDensity)) * 0.95
                    Else
                        outmax = (SlpmToGramsPerHour(Stn_AIO(iStn, asLiveFuelVaporFlow).EuMax, LiveFuelVaporDensity)) * 0.95
                    End If
                Case Else
                    outmax = 0
            End Select
            ' enable
            CurPID.Enable = IIf(((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)), True, False)
            ' inhibit
            '   inhibit if "LoadRate Timeout"  ???
            CurPID.Inhibit = IIf((StationControl(iStn, iShift).Mode <> VBLOAD) Or (LoadControl(iStn, iShift).Phase < LoadLoading), True, False)
            ' process value
            CurPID.PV = LoadControl(iStn, iShift).TotalWtChgRate * (100# / outmax)      ' convert to 0-100%; max rate = 60 grams/hr per slpm
            CurPID.PV = IIf(CurPID.PV > 150, 150, CurPID.PV)                                ' clip hi to 100
            CurPID.PV = IIf(CurPID.PV < 0, 0, CurPID.PV)                                    ' clip lo to 0
            ' set point
            CurPID.SP = LoadControl(iStn, iShift).LoadRateTarget * (100# / outmax)         ' convert to 0-100%; max rate = 60 grams/hr per slpm
            CurPID.SP = IIf(CurPID.SP > 150, 150, CurPID.SP)                                ' clip hi to 100
            CurPID.SP = IIf(CurPID.SP < 0, 0, CurPID.SP)                                    ' clip lo to 0
            ' error  (note: Reverse Action switches polarity of Error)
            CurPID.Rev = True
            CurPID.Er = IIf(CurPID.Rev, (CurPID.SP - CurPID.PV), (CurPID.PV - CurPID.SP))
        Case stn1LeakTest To stn9LeakTest
ChgErrModule 8, 31334
            ' STN LeakTest Controller  (convert EU Range to %)
            iStn = iController - 20
            iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
            iFuncPV = asNitrogenFlow
            iFuncSP = asNitrogenFlowSP
            iMfc = MFCNITROGEN
            outmax = (Stn_AIO(iStn, iFuncPV).EuMin) + (0.95 * (Stn_AIO(iStn, iFuncPV).EuMax - Stn_AIO(iStn, iFuncPV).EuMin))
            ' enable
            CurPID.Enable = IIf((StationControl(iStn, iShift).Mode = VBLEAKTEST), True, False)
            ' inhibit
            CurPID.Inhibit = IIf(((SEQ_Step(iStn, iShift) < 3) Or (SEQ_Step(iStn, iShift) > 9)), True, False)
            ' process value
            CurPID.PV = 100# * ((Stn_AIO(iStn, asLtInletPress).EUValue - Stn_AIO(iStn, asLtInletPress).EuMin) / (Stn_AIO(iStn, asLtInletPress).EuMax - Stn_AIO(iStn, asLtInletPress).EuMin))
            ' set point
            CurPID.SP = 100# * ((Rcp_LeakTest.TargetPress - Stn_AIO(iStn, asLtInletPress).EuMin) / (Stn_AIO(iStn, asLtInletPress).EuMax - Stn_AIO(iStn, asLtInletPress).EuMin))
            ' error  (note: Reverse Action switches polarity of Error)
            CurPID.Er = IIf(CurPID.Rev, (CurPID.SP - CurPID.PV), (CurPID.PV - CurPID.SP))
        Case Else
ChgErrModule 8, 31335
            ' undefined controller
            CurPID.Enable = False
            CurPID.Inhibit = True
            CurPID.PV = 1
            CurPID.SP = 1
            CurPID.Er = 0
            CurPID.CumI = 0
    End Select
ChgErrModule 8, 31338
    cumMax = CurPID.CumImax
    cumMin = CurPID.CumImin
    outputMax = CurPID.outmax
    outputMin = CurPID.outmin
ChgErrModule 8, 31339
    If Pause_Alarm = SYSTEMPAUSED Then
'ChgErrModule 8, 31340
        ' System is Paused, Turn Off All Outputs
        CurPID.out = 0#
    Else
'ChgErrModule 8, 31341
        ' Normal Operation (system is not paused)
        ' Enabled and Not Inhibited?
        If CurPID.Enable And Not CurPID.Inhibit Then
            If Not PID_INFO(iController).Enable Or PID_INFO(iController).Inhibit Then
                ' Just-Got-Enable Or Just-Lost-Inhibit; i.e. first pass
                CurPID.LastUpdate = Timer
                Select Case iController
                    Case pasMOISTURE
'ChgErrModule 8, 31342
                        Select Case CurPID.SP
                            Case Is > 75
                                NewOut = 95
                            Case Is > 70
                                NewOut = 60
                            Case Is > 65
                                NewOut = 50
                            Case Is > 55
                                NewOut = 40
                            Case Is > 50
                                NewOut = 25
                            Case Is > 30
                                NewOut = 7
                            Case Else
                                NewOut = 3.5
                        End Select
                    Case wbSuperTemp
'ChgErrModule 8, 31343
                        NewOut = CSng(100) * ((CurPID.SP - WB_AIO.EuMin) / (WB_AIO.EuMax - WB_AIO.EuMin))
                    Case stn1LoadRate To stn9LoadRate
'ChgErrModule 8, 31344
                         Select Case CurPID.SP
                            Case Is > 75
                                NewOut = 45
                            Case Is > 70
                                NewOut = 40
                            Case Is > 65
                                NewOut = 35
                            Case Is > 55
                                NewOut = 25
                            Case Is > 50
                                NewOut = 20
                            Case Is > 30
                                NewOut = 10
                            Case Else
                                NewOut = 5
                        End Select
                    Case stn1LeakTest To stn9LeakTest
'ChgErrModule 8, 31345
                        ' STN LeakTest Controller
                        tmpSpan = Stn_AIO(iStn, iFuncSP).EuMax - Stn_AIO(iStn, iFuncSP).EuMin
                        outInit = 100# * ((Cfg_LeakTest.InitialN2Flow - Stn_AIO(iStn, iFuncSP).EuMin) / tmpSpan)
                         Select Case outInit
                            Case Is > 75
                                NewOut = 45
                            Case Is > 70
                                NewOut = 40
                            Case Is > 65
                                NewOut = 35
                            Case Is > 55
                                NewOut = 25
                            Case Is > 50
                                NewOut = 20
                            Case Is > 30
                                NewOut = 10
                            Case Else
                                NewOut = 5
                        End Select
                   Case Else
                        ' undefined controller
                End Select
                ' calculate required CumI for the desired initial output value  SEQ_Step(iStation, iShift) = 4
                CurPID.CumI = NewOut - (50# + (CurPID.Pgain * CurPID.Er))
                CurPID.CumI = IIf(CurPID.CumI > cumMax, cumMax, CurPID.CumI)                                    ' clip hi to cumMax
                CurPID.CumI = IIf(CurPID.CumI < cumMin, cumMin, CurPID.CumI)                                    ' clip lo to cumMin
                CurPID.out = NewOut
            Else
'ChgErrModule 8, 31346
                ' not first pass
                ' How long since last update
                NowTmr = IIf(Timer > CurPID.LastUpdate, Timer, Timer + 86400)
                DeltaTmr = NowTmr - CurPID.LastUpdate
                CurPID.LastUpdate = Timer
                DeltaI = CurPID.Igain * CurPID.Er * (DeltaTmr / 60)
                If DeltaI > 0 And (CurPID.CumI < 0 Or CurPID.out <= 100) Then CurPID.CumI = CurPID.CumI + DeltaI
                If DeltaI < 0 And (CurPID.CumI > 0 Or CurPID.out >= 0) Then CurPID.CumI = CurPID.CumI + DeltaI
                CurPID.CumI = IIf(CurPID.CumI > cumMax, cumMax, CurPID.CumI)                                    ' clip hi to cumMax
                CurPID.CumI = IIf(CurPID.CumI < cumMin, cumMin, CurPID.CumI)                                    ' clip lo to cumMin
                NewOut = (CurPID.Pgain * CurPID.Er) + CurPID.CumI + 50#
                CurPID.out = IIf(NewOut < 0, 0, IIf(NewOut > outputMax, outputMax, NewOut))
            End If
        Else
ChgErrModule 8, 31347
            If (iController <> wbSuperTemp) Then
                ' Set Output=0% if Just-Lost-Enable Or Inhibit is True
                If PID_INFO(iController).Enable Then CurPID.out = 0
                If CurPID.Inhibit Then CurPID.out = 0
            Else
ChgErrModule 8, 31348
                If (StationConfig(iStn, iShift).WaterBathControl <> wbDirect) Then
ChgErrModule 8, 31349
                    ' Set Output=0% if Just-Lost-Enable Or Inhibit is True
                    If PID_INFO(iController).Enable Then CurPID.out = 0
ChgErrModule 8, 31350
                    If CurPID.Inhibit Then CurPID.out = 0
                End If
            End If
        End If
    End If
    ' Update Controller Output Values
    Select Case iController
        Case pasMOISTURE
'ChgErrModule 8, 31351
            ' PAS Moisture
            If CurPID.Enable And Not CurPID.Inhibit Then
                ' Enabled and Not Inhibited
                If Com_AIO(acPASMoistCntrlOut).EUValue <> CurPID.out Then _
                    Com_OutAnalog acPASMoistCntrlOut, CurPID.out, outNORMAL
            ElseIf CurPID.Inhibit Then
                ' Inhibited
                If Com_AIO(acPASMoistCntrlOut).EUValue <> CurPID.out Then _
                    Com_OutAnalog acPASMoistCntrlOut, 0, outZERO
            Else
                ' overmoisturizing protection
                '       none
            End If
        Case wbSuperTemp
'ChgErrModule 8, 31352
            ' WaterBath Supervisory Temperature Control
            If CurPID.Enable And Not CurPID.Inhibit Then
                ' Enabled and Not Inhibited
                tmpSpan = WB_AIO.EuMax - WB_AIO.EuMin
                outEU = WB_AIO.EuMin + (CurPID.out * tmpSpan * 0.01)
                WaterBathSP = outEU
            ElseIf CurPID.Inhibit Then
                ' Inhibited
                Select Case StationConfig(iStn, iShift).WaterBathControl
                    Case wbDirect
                        tmpSpan = WB_AIO.EuMax - WB_AIO.EuMin
                        outEU = WB_AIO.EuMin + (CurPID.out * tmpSpan * 0.01)
                        WaterBathSP = outEU
                    Case Else
                        ' don't change the WaterBathSP
                End Select
            Else
                ' Not Enabled but Not Inhibited
                If USINGC Then WaterBathSP = Com_AIO(acAmbTempSensor).EUValue
                If USINGF Then WaterBathSP = DegFtoC(Com_AIO(acAmbTempSensor).EUValue)
            End If
        Case stn1LoadRate To stn9LoadRate
'ChgErrModule 8, 31353
            ' STN Loadrate Controller
            Select Case STN_INFO(iStn).Type
                Case STN_LIVEFUEL_TYPE, STN_LIVEREG_TYPE
                    If CurPID.Enable And Not CurPID.Inhibit Then
                        ' Enabled and Not Inhibited
                        tmpSpan = Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMin
                        outEU = Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMin + (CurPID.out * tmpSpan * 0.01)
                        ' update MFC SP
                        If (Stn_AIO(iStn, asLiveFuelVaporFlowSP).EUValue <> outEU) Then
                            StationRecipe(iStn, iShift).NitrogenFlow = outEU
                            LoadSetPoint_Update iStn, iShift
                        End If
                    ElseIf CurPID.Inhibit Then
                        ' Inhibited
                        outEU = 0
                        ' update MFC SP
                        If (Stn_AIO(iStn, asLiveFuelVaporFlowSP).EUValue <> outEU) Then
                            StationRecipe(iStn, iShift).NitrogenFlow = outEU
                            LoadSetPoint_Update iStn, iShift
                        End If
                    End If
                Case STN_LIVEORVR2_TYPE
                    If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                        If CurPID.Enable And Not CurPID.Inhibit Then
                            ' Enabled and Not Inhibited
                            tmpSpan = Stn_AIO(iStn, asLiveFuelVaporORVRFlowSP).EuMax - Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMin
                            outEU = Stn_AIO(iStn, asLiveFuelVaporORVRFlowSP).EuMin + (CurPID.out * tmpSpan * 0.01)
                            ' update MFC SP
                            If (Stn_AIO(iStn, asLiveFuelVaporORVRFlowSP).EUValue <> outEU) Then
                                StationRecipe(iStn, iShift).NitrogenFlow = outEU
                                LoadSetPoint_Update iStn, iShift
                            End If
                        ElseIf CurPID.Inhibit Then
                            ' Inhibited
                            outEU = 0
                            ' update MFC SP
                            If (Stn_AIO(iStn, asLiveFuelVaporORVRFlowSP).EUValue <> outEU) Then
                                StationRecipe(iStn, iShift).NitrogenFlow = outEU
                                LoadSetPoint_Update iStn, iShift
                            End If
                        End If
                    Else
                        If CurPID.Enable And Not CurPID.Inhibit Then
                            ' Enabled and Not Inhibited
                            tmpSpan = Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMax - Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMin
                            outEU = Stn_AIO(iStn, asLiveFuelVaporFlowSP).EuMin + (CurPID.out * tmpSpan * 0.01)
                            ' update MFC SP
                            If (Stn_AIO(iStn, asLiveFuelVaporFlowSP).EUValue <> outEU) Then
                                StationRecipe(iStn, iShift).NitrogenFlow = outEU
                                LoadSetPoint_Update iStn, iShift
                            End If
                        ElseIf CurPID.Inhibit Then
                            ' Inhibited
                            outEU = 0
                            ' update MFC SP
                            If (Stn_AIO(iStn, asLiveFuelVaporFlowSP).EUValue <> outEU) Then
                                StationRecipe(iStn, iShift).NitrogenFlow = outEU
                                LoadSetPoint_Update iStn, iShift
                            End If
                        End If
                    End If
                Case Else
                    outmax = 0
            End Select
        Case stn1LeakTest To stn9LeakTest
ChgErrModule 8, 31354
            ' STN LeakTest Controller
            If CurPID.Enable And Not CurPID.Inhibit Then
                ' Enabled and Not Inhibited
                tmpSpan = Stn_AIO(iStn, iFuncSP).EuMax - Stn_AIO(iStn, iFuncSP).EuMin
                outEU = (CurPID.out * 0.01 * tmpSpan) + Stn_AIO(iStn, iFuncSP).EuMin
                outRaw = Stn_AIO(iStn, iFuncSP).EuMin + (tmpSpan * Cal_MfcOutput(outEU, iStn, iMfc, Stn_MfcCal(iStn, iMfc)))
                ' update MFC SP
                If (Stn_AIO(iStn, iFuncSP).EUValue <> outEU) Then
                    ' set Nitrogen MFC setpoint
                    Stn_OutAnalog iStn, iFuncSP, outRaw, outNORMAL
                End If
            ElseIf CurPID.Inhibit Then
                ' Inhibited
                outEU = 0
                outRaw = 0
                ' update MFC SP
                If (Stn_AIO(iStn, iFuncSP).EUValue <> outEU) Then
                    ' set Nitrogen MFC setpoint
                    Stn_OutAnalog iStn, iFuncSP, outRaw, outZERO
                End If
            End If
        Case Else
ChgErrModule 8, 31356
            ' not defined
    End Select
ChgErrModule 8, 31359
    ' Update Control Block
    PID_INFO(iController) = CurPID

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

Sub PurgeAir_Controller()
'
'   Controls Purge "Source(s)" Valves to match needs of stations
'
'   16 February 2007
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 313
Dim Idx As Integer
Dim temptime As Date
Dim tempSec As Integer
Dim turnOffpulse As Boolean
Dim turnOnRequest As Boolean
Dim usingAkRequest As Boolean
Dim usingHdwRequest As Boolean
Dim pasAvailable As Boolean
Dim pagFlag1 As Boolean
Dim pagFlag2 As Boolean
Dim rdyFlag As Boolean
     
    turnOffpulse = False
    turnOnRequest = False
    usingAkRequest = False
    usingHdwRequest = False
    ' Is PAS (local or remote) Available (i.e. Not Shutdown or Timed Out)
    pagFlag1 = IIf((LocalPagControl.Type = pagClient), True, False)
    pagFlag2 = IIf((USINGPASLOCALCONTROL) And (PAS_INFO(pasTEMPERATURE).timeOut Or PAS_INFO(pasMOISTURE).timeOut), False, True)
    pasAvailable = IIf((pagFlag1 Or pagFlag2), True, False)
     
    ' Paused OR Normal Operation ?
    If Pause_Alarm = NOTPAUSED Then
        ' System is Not Paused; Normal Operation
        For Idx = 1 To NR_PRGAIR
        
            temptime = Now() - PRG_INFO(Idx).lastTime
            tempSec = (60 * Minute(temptime)) + Second(temptime)
            If tempSec > PRG_INFO(Idx).CheckSecs Then
                If PRG_INFO(Idx).StandbyRequest Or PRG_INFO(Idx).LastStandbyRequest Then
                    PRG_INFO(Idx).StandingBy = True
                Else
                    PRG_INFO(Idx).StandingBy = False
                End If
                If PRG_INFO(Idx).RequestRdy Or PRG_INFO(Idx).LastRequestRdy Then
                    PRG_INFO(Idx).Requested = True
                Else
                    PRG_INFO(Idx).Requested = False
                End If
                If PRG_INFO(Idx).RequestRun Or PRG_INFO(Idx).LastRequestRun Then
                    PRG_INFO(Idx).Running = True
                Else
                    PRG_INFO(Idx).Running = False
                End If
                PRG_INFO(Idx).LastRequestRdy = PRG_INFO(Idx).RequestRdy
                PRG_INFO(Idx).LastRequestRun = PRG_INFO(Idx).RequestRun
                PRG_INFO(Idx).LastStandbyRequest = PRG_INFO(Idx).StandbyRequest
                PRG_INFO(Idx).lastTime = Now()
                PRG_INFO(Idx).RequestRdy = False
                PRG_INFO(Idx).RequestRun = False
                PRG_INFO(Idx).StandbyRequest = False
            End If
            
            ' Note: No Hdw Request/Ready for Positive Pressure Purges
            If (((PRG_INFO(Idx).UsingPrgReqAK) Or (PRG_INFO(Idx).UsingPrgReqHdw)) And (Not SysConfig.PosPressPurge)) Then
            
                ' Using Request/Ready Hardware or AKinterface (i.e. a Remote Purge Air Conditioning System; usually shared by all PurgeAir Sources)
                If (PRG_INFO(Idx).UsingPrgReqAK) Then usingAkRequest = True
                If (PRG_INFO(Idx).UsingPrgReqHdw) Then usingHdwRequest = True
                If (PRG_INFO(Idx).Running Or PRG_INFO(Idx).Requested Or LocalPagControl.ReqIn) And pasAvailable Then
                    ' a PurgeAir Source needs Conditioned Purge Air
                    turnOnRequest = True
                End If
                 
                rdyFlag = (Com_DIO(icPurgeReadyIn).Value Or MasterPagData.RdyOut)
                If (rdyFlag And (PRG_INFO(Idx).Running Or PRG_INFO(Idx).Requested)) Then
                    PRG_INFO(Idx).Ready = True
                Else
                    If PRG_INFO(Idx).Ready Then turnOffpulse = True
                    PRG_INFO(Idx).Ready = False
                End If
                
            Else
            
                ' Regular
                If PRG_INFO(Idx).Requested Or PRG_INFO(Idx).Running Then
                    PRG_INFO(Idx).Ready = True
                Else
                    ' Not Running AND Not Requested
                    If PRG_INFO(Idx).Ready Then turnOffpulse = True
                    PRG_INFO(Idx).Ready = False
                End If
                    
            End If
                 
    
            If PRG_INFO(Idx).Running Then
                ' Turn the Valves On
                '   and keep them On
                If SysConfig.PosPressPurge Then
                    ' POSITIVE PRESSURE PURGE
                    If Prg_DIO(Idx, ipPiabSol).Value Then Prg_OutDigital Idx, ipPiabSol, cOFF                                               ' turn off Vacuum Purge Valve
                    If Not Prg_DIO(Idx, ipPosPrsPrgSol).Value Then Prg_OutDigital Idx, ipPosPrsPrgSol, cON                                  ' turn oN Positive Pressure Purge Valve
                Else
                    ' VACUUM PURGE (Normal)
                    If Not Prg_DIO(Idx, ipPiabSol).Value Then Prg_OutDigital Idx, ipPiabSol, cON                                            ' turn ON Vacuum Purge Valve
                    If Prg_DIO(Idx, ipPosPrsPrgSol).Value Then Prg_OutDigital Idx, ipPosPrsPrgSol, cOFF                                     ' turn off Positive Pressure Purge Valve
                End If
                If (Not Prg_DIO(Idx, ipAuxAirSol).Value And PRG_INFO(Idx).UsingAuxAirSol) Then Prg_OutDigital Idx, ipAuxAirSol, cON     ' turn ON Aux Air
            ElseIf turnOffpulse Or PRG_INFO(Idx).StandingBy Then
                ' Turn Off the valves
                '   but only keep them off if one of this PurgeAir's stations is active
                '  (don't want to interfere with manual control of I/O of an idle station)
                If Prg_DIO(Idx, ipPiabSol).Value Then Prg_OutDigital Idx, ipPiabSol, cOFF                                               ' turn off Vacuum Purge Valve
                If (Prg_DIO(Idx, ipPosPrsPrgSol).Value And PRG_INFO(Idx).UsingPosPrsPrg) Then Prg_OutDigital Idx, ipPosPrsPrgSol, cOFF  ' turn off Positive Pressure Purge Valve
                If (Prg_DIO(Idx, ipAuxAirSol).Value And PRG_INFO(Idx).UsingAuxAirSol) Then Prg_OutDigital Idx, ipAuxAirSol, cOFF        ' turn off Aux Air
            End If
                
        Next Idx
        
        ' Actually turn the Remote Request On or Off
        PAG_Request = turnOnRequest
        If usingAkRequest Then
            If turnOnRequest Then
                LocalPagControl.ReqOut = True          ' turn ON the request to Purge Cabinet
            Else
                LocalPagControl.ReqOut = False         ' turn off the request to Purge Cabinet
            End If
        End If
        If usingHdwRequest Then
            If (turnOnRequest Or MasterPagData.ReqIn) Then
                If Not Com_DIO(icPurgeRequestOut).Value Then Com_OutDigital icPurgeRequestOut, cON          ' turn ON the request to Purge Cabinet
            Else
                If Com_DIO(icPurgeRequestOut).Value Then Com_OutDigital icPurgeRequestOut, cOFF             ' turn off the request to Purge Cabinet
            End If
        End If
         
         
         
    Else
    
        ' System is Paused; Turn Off Everything
        For Idx = 1 To NR_PRGAIR
        
            If Prg_DIO(Idx, ipPiabSol).Value Then Prg_OutDigital Idx, ipPiabSol, cOFF                                               ' turn off Vacuum Purge Valve
            If Prg_DIO(Idx, ipPosPrsPrgSol).Value Then Prg_OutDigital Idx, ipPosPrsPrgSol, cOFF                                     ' turn off Positive Pressure Purge Valve
            If (Prg_DIO(Idx, ipAuxAirSol).Value And PRG_INFO(Idx).UsingAuxAirSol) Then Prg_OutDigital Idx, ipAuxAirSol, cOFF        ' turn off Aux Air
       
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

Sub PurgeOven_Controller()
'
'   Controls Purge Oven
'
'   6 November 2017
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 3137
Dim iStation As Integer
Dim iShift As Integer
Dim okDelta As Single
Dim deltaTemp As Single
Dim deltaSP As Single
'Dim ovenAvailable
                     
    
    ' Note: One oven per station
    If USINGPURGEOVEN Then
        ' Paused OR Normal Operation ?
        If Pause_Alarm = NOTPAUSED Then
        
            ' System is Not Paused; Normal Operation
            For iStation = 1 To NR_STN
                For iShift = 1 To NR_SHIFT
                    If (STN_INFO(iStation).USINGPURGEOVEN) Then
                        If ((Not StationControl(iStation, iShift).ModeIsIdle_Debounced) And StationRecipe(iStation, iShift).PurgeOven) Then
                            Stn_OutAnalog iStation, asPurgeOvenTempSP, StationRecipe(iStation, iShift).PurgeOvenSP, outNORMAL
                            PurgeControl(iStation, iShift).PurgeOvenPV = Stn_AIO(iStation, asPurgeOvenTemp).EUValue
                            PurgeControl(iStation, iShift).PurgeOvenSP = Stn_AIO(iStation, asPurgeOvenTempSP).EUValue
                            deltaTemp = Abs(PurgeControl(iStation, iShift).PurgeOvenSP - PurgeControl(iStation, iShift).PurgeOvenPV)
                            okDelta = StationConfig(iStation, iShift).PurgeOvenBand
                            If USINGC Then
                                deltaSP = Abs(Stn_AIO(iStation, asPurgeOvenTempSP).EUValue - StationRecipe(iStation, iShift).PurgeOvenSP)
                            Else
                                deltaSP = Abs(DegFtoC(Stn_AIO(iStation, asPurgeOvenTempSP).EUValue) - DegFtoC(StationRecipe(iStation, iShift).ADF_HeaterSP))
                            End If
                            PurgeControl(iStation, iShift).PurgeOvenTempOK = IIf(((deltaTemp <= okDelta) And (deltaSP <= 0.1)), True, False)
                            antiRepeat(iStation, iShift) = True
                        Else
                            If antiRepeat(iStation, iShift) Then Stn_OutAnalog iStation, asPurgeOvenTempSP, 0#, outZERO
                            PurgeControl(iStation, iShift).PurgeOvenPV = Stn_AIO(iStation, asPurgeOvenTemp).EUValue
                            PurgeControl(iStation, iShift).PurgeOvenSP = Stn_AIO(iStation, asPurgeOvenTempSP).EUValue
                            PurgeControl(iStation, iShift).PurgeOvenTempOK = False
                            antiRepeat(iStation, iShift) = False
                        End If
                    End If
                Next iShift
            Next iStation
            
        Else
        
            ' System is Paused; Turn Off Everything
            For iStation = 1 To NR_STN
                For iShift = 1 To NR_SHIFT
                    If (STN_INFO(iStation).USINGPURGEOVEN) Then
                        If antiRepeat(iStation, iShift) Then Stn_OutAnalog iStation, asPurgeOvenTempSP, 0#, outZERO
                        PurgeControl(iStation, iShift).PurgeOvenPV = Stn_AIO(iStation, asPurgeOvenTemp).EUValue
                        PurgeControl(iStation, iShift).PurgeOvenSP = Stn_AIO(iStation, asPurgeOvenTempSP).EUValue
                        PurgeControl(iStation, iShift).PurgeOvenTempOK = False
                        antiRepeat(iStation, iShift) = False
                    End If
                Next iShift
            Next iStation
            
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

Sub WaterBath_Controller(ByVal iStation As Integer)
'
'   Controls WaterBath
'
'   7 November 2017
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 7137
Dim iShift As Integer
Dim okDelta As Single
Dim deltaSP As Single
Dim deltaTemp As Single
'Dim ovenAvailable
                     
    ' Note: Only One WaterBath per system
    
    'Active Shift
    iShift = IIf((Stn_ActiveShift(iStation) > 0), Stn_ActiveShift(iStation), 1)
            

    If USINGWATERBATH Then
        ' Paused OR Normal Operation ?
        If Pause_Alarm = NOTPAUSED Then
        
            ' System is Not Paused; Normal Operation
            If ((Not StationControl(iStation, iShift).ModeIsIdle_Debounced) And StationRecipe(iStation, iShift).ADF_Heater) Then
                ' *****************************************
                ' WaterBath Supervisory Temperature Control
                ' *****************************************
                LF_Chiller.OutOut = pumpFAST
                Controller_PID wbSuperTemp
                LoadControl(iStation, iShift).WaterBathPV = LF_Chiller.PvIn
                LoadControl(iStation, iShift).WaterBathSP = LF_Chiller.SpIn
                Select Case StationConfig(iStation, iShift).WaterBathControl
                    Case wbDirect
                        deltaTemp = Abs(LoadControl(iStation, iShift).WaterBathSP - LoadControl(iStation, iShift).WaterBathPV)
                        okDelta = IIf(USINGC, StationConfig(iStation, iShift).Tol_WaterBathTemp, DegFtoC(CSng(StationConfig(iStation, iShift).Tol_WaterBathTemp)))
                        If USINGC Then
                            deltaSP = Abs(LF_Chiller.SpIn - LF_Chiller.SpOut)
                        Else
                            deltaSP = Abs(LF_Chiller.SpIn - DegFtoC(StationRecipe(iStation, iShift).ADF_HeaterSP))
                        End If
                        LoadControl(iStation, iShift).WaterBathTempOK = IIf(((deltaTemp <= okDelta) And (deltaSP <= 0.1)), True, False)
                    Case wbFuelTemp
                        deltaTemp = Abs(PID_INFO(wbSuperTemp).SP - PID_INFO(wbSuperTemp).PV)
                        okDelta = IIf(USINGC, StationConfig(iStation, iShift).Tol_WaterBathTemp, DegFtoC(CSng(StationConfig(iStation, iShift).Tol_WaterBathTemp)))
                        deltaSP = Abs(LF_Chiller.SpIn - LF_Chiller.SpOut)
                        LoadControl(iStation, iShift).WaterBathTempOK = IIf((deltaTemp <= okDelta), True, False)
                    Case wbVaporTemp
                        deltaTemp = Abs(PID_INFO(wbSuperTemp).SP - PID_INFO(wbSuperTemp).PV)
                        okDelta = IIf(USINGC, StationConfig(iStation, iShift).Tol_WaterBathTemp, DegFtoC(CSng(StationConfig(iStation, iShift).Tol_WaterBathTemp)))
                        deltaSP = Abs(LF_Chiller.SpIn - LF_Chiller.SpOut)
                        LoadControl(iStation, iShift).WaterBathTempOK = IIf((deltaTemp <= okDelta), True, False)
                End Select
            Else
                WaterBathSP = Com_AIO(acAmbTempSensor).EUValue
                LoadControl(iStation, iShift).WaterBathPV = LF_Chiller.PvIn
                LoadControl(iStation, iShift).WaterBathSP = LF_Chiller.SpIn
                LoadControl(iStation, iShift).WaterBathTempOK = False
                LF_Chiller.OutOut = pumpSLOW
            End If
            
        Else
        
            ' System is Paused; Turn Off Everything
            WaterBathSP = Com_AIO(acAmbTempSensor).EUValue
            LoadControl(iStation, iShift).WaterBathPV = LF_Chiller.PvIn
            LoadControl(iStation, iShift).WaterBathSP = LF_Chiller.SpIn
            LoadControl(iStation, iShift).WaterBathTempOK = False
            
        End If
        AdfControl(iStation).TempOK = LoadControl(iStation, iShift).WaterBathTempOK
    
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

Sub Seq_Controller()
' Routine Name:  Station Sequence Controller
' Author:        MMW
' Description:
' Controls various Station Sequences
'
'   Sequence Number
'   0       Idle
'   1       PostPurge N2 Feed (for CanVentAlarm)
'   2       unused
'   3       unused
'   4       LeakTest (LeakTest stations only)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 221
Dim iStation As Integer
Dim iShift As Integer
Dim iPid As Integer
Dim deltaPress As Single
Dim span As Single
Dim temptime As Date
Dim tempSec, tempMin As Integer
Dim TempTol As Single
Static tempDeff As Single
Dim Nitrogen_Rate As Single
Dim Nitrogen_Output As Single


    For iStation = 1 To LAST_STN
    
        For iShift = 1 To NR_SHIFT
    
            '   Which Sequence are we Executing now ?
            '
            Select Case SEQ_Nmbr(iStation, iShift)
                
                ' *************************************
                ' *************************************
                '                 IDLE
                ' *************************************
                ' *************************************
                Case seqIdle
                    SEQ_Task(iStation, iShift) = "Idle"
                    SEQ_Message(iStation, iShift) = "Idle"
                    SEQ_Step(iStation, iShift) = 0
                
                ' *************************************
                ' *************************************
                ' PostPurge N2 Feed (with CanVentAlarm)
                ' *************************************
                ' *************************************
                Case seqCanVentN2Feed
                    
                    SEQ_Task(iStation, iShift) = "N2 Feed @1SLPM in Progress"
                    
                    Select Case SEQ_Step(iStation, iShift)
            
                        Case 0
                            SEQ_Message(iStation, iShift) = "Idle"
                        
                        Case 1
                            SEQ_Message(iStation, iShift) = "Start N2 Feed"
                            SEQ_StartTime(iStation, iShift) = Now()                               ' sequence start time
                            SEQ_Step_Time(iStation, iShift) = Now() + TimeSerial(0, 10, 0)        ' setup 10 minute timeout
                            SEQ_OOT(iStation, iShift) = False
                            SEQ_Step(iStation, iShift) = 2                                        ' Continue
                  
                        
                        Case 2
                            SEQ_Message(iStation, iShift) = "Open Valves; Set MFC SetPoint"
                            If StationRecipe(iStation, iShift).UsePriScale And StationControl(iStation, iShift).PriScaleStn > 0 _
                               And StationControl(iStation, iShift).PriScaleStn < FIRST_REMOTESCALE Then
                                Stn_OutDigital StationControl(iStation, iShift).PriScaleStn, isPriAuxVentSol, cON  ' Diff station valve
                            End If
                            ' Shift valves
                            Select Case iShift
                                Case 1
                                    ' nothing to do
                                Case 2
                                    Stn_OutDigital iStation, isLoadShift2Sol, cON
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cON
                                Case 3
                                    Stn_OutDigital iStation, isLoadShift3Sol, cON
                                    Stn_OutDigital iStation, isLoadShift2Sol, cON
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cON
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
                            ' Nitrogen
                            Stn_OutDigital iStation, isNitrogenSol, cON
                            ' N2 rate = (0.9 * EU Max) OR 1.0, whichever is smaller)
                            Nitrogen_Rate = CSng(0.9 * Stn_AIO(iStation, asNitrogenFlowSP).EuMax)
                            If Nitrogen_Rate > CSng(1#) Then Nitrogen_Rate = CSng(1#)
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            SEQ_Step(iStation, iShift) = 3                                       ' Continue
                 
                        Case 3
                            temptime = SEQ_Step_Time(iStation, iShift) - Now()
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            SEQ_Message(iStation, iShift) = "Waiting for CanVent Flow Switch - " & Format(tempSec, "##0") & " sec"
                            If Stn_DIO(iStation, isCanVentAlarmSw).Value Then
                                ' FlowSw Made; Continue
                                SEQ_Step_StartTime(iStation, iShift) = Now()
                                SEQ_Step(iStation, iShift) = 4                                   ' Continue
                            ElseIf Now() > SEQ_Step_Time(iStation, iShift) Then
                                ' Timeout
                                SEQ_EndTime(iStation, iShift) = Now()
                                SEQ_OOT(iStation, iShift) = True                                 ' declare OOT
                                If Len(StationControl(iStation, iShift).DBFile) > 0 Then OOT_Write CInt(iStation), CInt(iShift), "Too Long to Make CanVent FS (PostPurge)"
                                SEQ_Step(iStation, iShift) = 91                                  ' abort
                            ElseIf Not Stn_DIO(iStation, isNitrogenSol).Value Then
                                SEQ_Step(iStation, iShift) = 2                                   ' Restart N2 Feed
                            End If
                 
                        Case 4
                            temptime = SEQ_Step_Time(iStation, iShift) - Now()
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            SEQ_Message(iStation, iShift) = "Monitoring CanVent Flow Switch - " & Format(tempSec, "##0") & " sec"
                            temptime = SEQ_Step_StartTime(iStation, iShift) - Now()
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            If Not Stn_DIO(iStation, isNitrogenSol).Value Then
                                SEQ_Step(iStation, iShift) = 2                                   ' Go Back and Restart N2 Feed
                            ElseIf Not Stn_DIO(iStation, isCanVentAlarmSw).Value Then
                                SEQ_Step(iStation, iShift) = 2                                   ' Go Back and Wait for CanVent FS
                            ElseIf tempSec > 3 Then
                                SEQ_Step(iStation, iShift) = 8                                   ' Success; time to shut things off
                            End If
                 
                        Case 8
                            SEQ_Message(iStation, iShift) = "Close Valves; Reset MFC SetPoint"
                            If StationRecipe(iStation, iShift).UsePriScale And StationControl(iStation, iShift).PriScaleStn > 0 _
                               And StationControl(iStation, iShift).PriScaleStn < FIRST_REMOTESCALE Then
                                Stn_OutDigital StationControl(iStation, iShift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
                            End If
                            ' Shift valves
                            Select Case iShift
                                Case 1
                                    ' nothing to do
                                Case 2
                                    Stn_OutDigital iStation, isLoadShift2Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cON
                                Case 3
                                    Stn_OutDigital iStation, isLoadShift2Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cON
                                    Stn_OutDigital iStation, isLoadShift3Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift3Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift3Sol, cON
                                Case 4
                                    Stn_OutDigital iStation, isLoadShift2Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cON
                                    Stn_OutDigital iStation, isLoadShift4Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift4Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift4Sol, cON
                            End Select
                            ' Nitrogen
                            Stn_OutDigital iStation, isNitrogenSol, cOFF
                            Nitrogen_Rate = CSng(0)                                             ' N2 rate = 0 SLPM
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            ' Done
                            SEQ_EndTime(iStation, iShift) = Now()
                            SEQ_Step(iStation, iShift) = 9                                        ' Done
                 
                        Case 9
                            temptime = SEQ_EndTime(iStation, iShift) - SEQ_StartTime(iStation, iShift)
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            SEQ_Message(iStation, iShift) = "Completed Successfully in " & Format(tempSec, "##0") & " sec"
                            ' Done; Wait for Reset
                                    
                       
                        Case 90
                            SEQ_Message(iStation, iShift) = "Aborted - Station Stopped"
                        
                        Case 91
                            temptime = SEQ_EndTime(iStation, iShift) - SEQ_StartTime(iStation, iShift)
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            SEQ_Message(iStation, iShift) = "Aborted after " & Format(tempSec, "##0") & "sec - Timeout"
                            If StationRecipe(iStation, iShift).UsePriScale And StationControl(iStation, iShift).PriScaleStn > 0 _
                               And StationControl(iStation, iShift).PriScaleStn < FIRST_REMOTESCALE Then
                                Stn_OutDigital StationControl(iStation, iShift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
                            End If
                            ' Shift valves
                            Select Case iShift
                                Case 1
                                    ' nothing to do
                                Case 2
                                    Stn_OutDigital iStation, isLoadShift2Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cOFF
                                Case 3
                                    Stn_OutDigital iStation, isLoadShift2Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cOFF
                                    Stn_OutDigital iStation, isLoadShift3Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift3Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift3Sol, cOFF
                                Case 4
                                    Stn_OutDigital iStation, isLoadShift2Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift2Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift2Sol, cOFF
                                    Stn_OutDigital iStation, isLoadShift4Sol, cOFF
                                    Stn_OutDigital iStation, isPurgeShift4Sol, cOFF
                                    Stn_OutDigital iStation, isVentShift4Sol, cOFF
                            End Select
                            ' Nitrogen
                            Stn_OutDigital iStation, isNitrogenSol, cOFF
                            Nitrogen_Rate = CSng(0)                                                 ' N2 rate = 0 SLPM
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            ' Aborted - Timeout; Wait for Reset or Restart
                        
                        Case 95
                            SEQ_Message(iStation, iShift) = "Aborted - Station Paused"
                        
                    End Select  ' Step
                

                ' *************************************
                ' *************************************
                ' LeakTest Station LeakTest Sequence
                ' *************************************
                ' *************************************
                Case seqLeakTest
                    
                    SEQ_Task(iStation, iShift) = "LeakTest in Progress"
                    
                    Select Case SEQ_Step(iStation, iShift)
            
                        Case 0
'                            SEQ_Message(iStation, iShift) = "Idle"
                        
                        Case 1
                            SEQ_Message(iStation, iShift) = "Start LeakTest"
                            SEQ_Message2(iStation, iShift) = "Start LeakTest"
                            SEQ_StartTime(iStation, iShift) = Now()                               ' sequence start time
                            SEQ_Time(iStation, iShift) = Now() + TimeSerial(0, 40, 0)             ' setup 40 minute timeout
 '                           SEQ_Step_Time(iStation, iShift) = Now() + TimeSerial(0, 40, 0)        ' setup 40 minute timeout
                            SEQ_OOT(iStation, iShift) = False
                            ' First DB write
                            LT_Write iStation, iShift, LT_BEGIN
                            CurrLT2_Data.ClkTime = Now
                            CurrLT2_Data.SecTimer = Timer
                            LT2_Write iStation, iShift, "Start LeakTest", CurrLT2_Data
                            SEQ_Step(iStation, iShift) = 2                                        ' Continue
                        
                        Case 2
                            SEQ_Message(iStation, iShift) = "Open N2 Valve"
                            ' Nitrogen
                            ' Nitrogen
                            If (Stn_DIO(iStation, isNitrogenSol).Value) Then
                                SEQ_Step(iStation, iShift) = 3                      ' Continue
                            Else
                                Stn_OutDigital iStation, isNitrogenSol, cON
                            End If
                 
                        Case 3
                            SEQ_Message(iStation, iShift) = "Set initial N2 Flow"
                            ' N2 rate = (EUmin + (0.1 * EUspan)) OR 0.25, whichever is smaller)
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Rate = Cfg_LeakTest.InitialN2Flow
                            If (Nitrogen_Rate > 0.25) Then Nitrogen_Rate = 0.25
                            ' is Actual SP within 2% of Target ???
                            If ((Stn_AIO(iStation, asNitrogenFlowSP).EUValue > (0.98 * Nitrogen_Rate)) And (Stn_AIO(iStation, asNitrogenFlowSP).EUValue < (1.02 * Nitrogen_Rate))) Then
                                ' remember Now & move on
                                SEQ_Step_Time(iStation, iShift) = Now()
                                SEQ_Step(iStation, iShift) = 4                                       ' Continue
                            Else
                                ' update MFC SetPoint
'  ####CHANGED   #####
                                 'Nitrogen_Output = Cal_Output(Nitrogen_Rate, iStation, iShift, MFCNITROGEN)
                                 Nitrogen_Output = span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN))
                                Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            End If
                 
                        Case 4
                            SEQ_Message(iStation, iShift) = "Waiting for stable Pressure"
                            SEQ_Message2(iStation, iShift) = "Waiting for stable Pressure"
                            deltaPress = Rcp_LeakTest.TargetPress - Stn_AIO(iStation, asLtInletPress).EUValue
                            tempSec = DateDiff("s", SEQ_Step_Time(iStation, iShift), Now)
                            SEQ_Message(iStation, iShift) = "Waiting for stable Pressure - " & Format(Stn_AIO(iStation, asLtInletPress).EUValue, "####0.0#") & " kPa"
                            If (deltaPress > Cfg_LeakTest.PressTol) Then
                                ' Press OOT; Reset Time counter
                                SEQ_Step_Time(iStation, iShift) = Now()                          ' Reset Timer
                            ElseIf (tempSec > CInt(Cfg_LeakTest.PressTolDuration)) Then
                                ' Press at Target & Stable; Continue
                                If (Len(StationControl(iStation, iShift).DBFile) > 3) Then Write_JLog iStation, iShift, "Pressure stable at " & Format(Stn_AIO(iStation, asLtInletPress).EUValue, "####0.0#") & " kPa"
                                LT2_Write iStation, iShift, "Press at Target & Stable", CurrLT2_Data
                                SEQ_Step(iStation, iShift) = 5                                   ' Continue
                            ElseIf (tempSec > CInt(Cfg_LeakTest.PressTimeout)) Then
                                ' Pressurization Timeout
                                SEQ_EndTime(iStation, iShift) = Now()
                                SEQ_OOT(iStation, iShift) = True                                 ' declare OOT
                                If Len(StationControl(iStation, iShift).DBFile) > 0 Then OOT_Write CInt(iStation), CInt(iShift), "Too Long to Achieve Target Pressure"
                                LT2_Write iStation, iShift, "Abort - Too Long to Achieve Target Pressure", CurrLT2_Data
                                SEQ_Step(iStation, iShift) = 91                                  ' Abort
                            ElseIf Not Stn_DIO(iStation, isNitrogenSol).Value Then
                                SEQ_Step(iStation, iShift) = 2                                   ' Restart N2 Feed
                            End If
                 
                        Case 5
                            SEQ_Message(iStation, iShift) = "Monitor Effective Diameter - " & Format(Deff, "###0.0###") & " in"
                            tempSec = DateDiff("s", SEQ_Step_Time(iStation, iShift), Now)
                            tempDeff = Deff
                            SEQ_Step_Time(iStation, iShift) = Now()
                            SEQ_Step(iStation, iShift) = 6                                      ' Continue
                 
                        Case 6
                            tempSec = DateDiff("s", SEQ_Step_Time(iStation, iShift), Now)
                            SEQ_Message(iStation, iShift) = "Monitoring Effective Diameter - " & Format(Deff, "###0.0###") & " in for " & Format(tempSec, "######0") & " sec"
                            SEQ_Message2(iStation, iShift) = "Monitoring Effective Diameter - " & Format(Deff, "###0.0###") & " in for " & Format(tempSec, "######0") & " sec"
                            TempTol = Abs(Deff - tempDeff)
                            If (tempSec > CInt(Rcp_LeakTest.HoldDuration)) Then
                                ' Deff Stable; Continue
                                SEQ_Step(iStation, iShift) = 8                                   ' Continue
                            ElseIf (TempTol > Cfg_LeakTest.DeffTol) Then
                                ' Deff OOT; Reset Time counter
                                SEQ_Step_Time(iStation, iShift) = Now()                          ' Reset Timer
                            ElseIf (tempSec > CInt(Cfg_LeakTest.timeOut)) Then
                                ' Deff Stablization Timeout
                                SEQ_EndTime(iStation, iShift) = Now()
                                SEQ_OOT(iStation, iShift) = True                                 ' declare OOT
                                If Len(StationControl(iStation, iShift).DBFile) > 0 Then OOT_Write CInt(iStation), CInt(iShift), "Too Long to Achieve Stable Deff"
                                LT2_Write iStation, iShift, "Abort - Too Long to Achieve Stable Deff", CurrLT2_Data
                                SEQ_Step(iStation, iShift) = 91                                  ' Abort
                            ElseIf Not Stn_DIO(iStation, isNitrogenSol).Value Then
                                SEQ_Step(iStation, iShift) = 2                                   ' Restart N2 Feed
                            End If
                 
                        Case 8
                            tempSec = DateDiff("s", SEQ_Step_Time(iStation, iShift), Now)
                            SEQ_Message(iStation, iShift) = "Effective Diameter Stable - " & Format(Deff, "###0.0###") & " in."
                            SEQ_Message2(iStation, iShift) = "Effective Diameter Stable - " & Format(Deff, "###0.0###") & " in. for " & Format(tempSec, "######0") & " sec  at " & Format(CurrLT2_Data.InPress, "###0.0###") & " kPa"
                            ' Last DB writes
                            LT2_Write iStation, iShift, "Effective Diameter Stable", CurrLT2_Data
                            LT_Write iStation, iShift, LT_DONE
                            ' Valves
                            Stn_OutDigital iStation, isNitrogenSol, cOFF
                            ' Nitrogen
                            Nitrogen_Rate = CSng(0)                                             ' N2 rate = 0 SLPM
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            ' Press PID
'                            PID_INFO(iStation).Enable = False
                            ' Done
                            SEQ_EndTime(iStation, iShift) = Now()
                            SEQ_Step(iStation, iShift) = 9                                       ' Done
                 
                        Case 9
                            temptime = SEQ_EndTime(iStation, iShift) - SEQ_StartTime(iStation, iShift)
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            SEQ_Message(iStation, iShift) = "Completed Successfully in " & Format(tempSec, "##0") & " sec"
                            ' Update Header data in data file
                            JobInfo(iStation, iShift).End_OK = True
                            StationControl(iStation, iShift).End_Time = Now
                            JobInfo(iStation, iShift).End_Baro = AmbBaro
                            Header_Update iStation, iShift
                            ' Done; Wait for Reset
                            SEQ_Step(iStation, iShift) = 10                                       ' Done
                       
                        Case 10
'                            SEQ_Message(iStation, iShift) = "Resetting - LeakTest Complete"
                            ' Valves
                            If (Stn_DIO(iStation, isNitrogenSol).Value) Then Stn_OutDigital iStation, isNitrogenSol, cOFF
                            ' Nitrogen
                            Nitrogen_Rate = CSng(0)                                             ' N2 rate = 0 SLPM
                            If (Stn_AIO(iStation, asNitrogenFlowSP).EUValue <> Nitrogen_Rate) Then
                                span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                                Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                                Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            End If
                            ' Note: SEQ is reset to idle by the LeakTest_Check(via Station_Finish) routine
                        
                        Case 90
                            SEQ_Message(iStation, iShift) = "Aborted - Station Stopped"
                            ' Valves
                            Stn_OutDigital iStation, isNitrogenSol, cOFF
                            ' Nitrogen
                            Nitrogen_Rate = CSng(0)                                             ' N2 rate = 0 SLPM
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                        
                        Case 91
                            temptime = SEQ_EndTime(iStation, iShift) - SEQ_StartTime(iStation, iShift)
                            tempSec = (60 * Minute(temptime)) + Second(temptime)
                            SEQ_Message(iStation, iShift) = "Aborted after " & Format(tempSec, "##0") & "sec - Timeout"
                            If StationRecipe(iStation, iShift).UsePriScale And StationControl(iStation, iShift).PriScaleStn > 0 _
                               And StationControl(iStation, iShift).PriScaleStn < FIRST_REMOTESCALE Then
                                Stn_OutDigital StationControl(iStation, iShift).PriScaleStn, isPriAuxVentSol, cOFF  ' Diff station valve
                            End If
                            ' Valves
                            Stn_OutDigital iStation, isNitrogenSol, cOFF
                            ' Nitrogen
                            Nitrogen_Rate = CSng(0)                                                 ' N2 rate = 0 SLPM
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                            ' Update Header data in data file
                            JobInfo(iStation, iShift).End_OK = False
                            StationControl(iStation, iShift).End_Time = Now
                            JobInfo(iStation, iShift).End_Baro = AmbBaro
                            Header_Update iStation, iShift
                            ' Aborted - Timeout; Wait for Reset or Restart
                        
                        Case 95
                            SEQ_Message(iStation, iShift) = "Aborted - Station Paused"
                            ' Valves
                            Stn_OutDigital iStation, isNitrogenSol, cOFF
                            ' Nitrogen
                            Nitrogen_Rate = CSng(0)                                                 ' N2 rate = 0 SLPM
                            span = Stn_AIO(iStation, asNitrogenFlowSP).EuMax - Stn_AIO(iStation, asNitrogenFlowSP).EuMin
                            
'                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_Output(Nitrogen_Rate, iStation, iShift, MFCNITROGEN))
                            Nitrogen_Output = Stn_AIO(iStation, asNitrogenFlowSP).EuMin + (span * Cal_MfcOutput(Nitrogen_Rate, iStation, MFCNITROGEN, Stn_MfcCal(iStation, MFCNITROGEN)))
                            
                            Stn_OutAnalog iStation, asNitrogenFlowSP, Nitrogen_Output, outNORMAL
                        
                    End Select  ' Step
                
                Case Else
                    ' Nothing to do; sequence not defined
                    SEQ_Task(iStation, iShift) = "undefined"
                    SEQ_Message(iStation, iShift) = "Idle"
                    SEQ_Step(iStation, iShift) = 0
                
            End Select
        Next iShift
        
    Next iStation
                
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

Sub SimulateIO()
Dim iStn, iShift, iPrg As Integer
Dim iTC, iFunc, iAddr, iChan As Integer
Dim iFuncIn, iFuncOut As Integer
Dim fVal1, fVal2, fVal3 As Double
Dim addr, chan As Integer
Dim tempPerc, tempVdc As Single
Dim flag As Boolean

SetErrModule 8, 1010
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' *************************
    ' Opto IO Communications OK
    IoComm_Flag = True
    ' *************************
    
     ' ****************************
    ' AK PurgeAir Communications OK
    PaComm_Flag = True
    ' *****************************
    
   ' *********************
    ' COMMON DIGITAL VALUES
    ' *********************
    ' do on first scan only
    If (Not firstPassSim) Then
        ' starting simulation
        frmAbout.UpdateMsg "Starting IO Simulation" & vbCrLf
        ' ESTOP, LEL, DOORS, BLOWER
        iFunc = icEStopSw
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
        iFunc = ic20LelGasSw
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
        iFunc = icDoorSw
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
        iFunc = icExhaustFlowFS
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
        ' LiveFuel PurgePS
        iFunc = icExhaustFlowFS
        If systemhasLIVEFUEL Then OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
        ' UPS Monitoring
        iFunc = icUpsFaultSw
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), True, False)
        iFunc = icUpsActiveSw
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), True, False)
        ' PAS Request In (PAS Local Control)
        iFunc = icPASPowerOnIn
        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
        ' AlarmSilence PB
        iFunc = icHornSilencePB
        If Com_DIO(iFunc).UseInverse Then
            OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = False
        Else
            OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = True
        End If
        ' Maintenance Mode DI
        iFunc = icMaintSw
        If Com_DIO(iFunc).UseInverse Then
            OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = False
'Debug.Print "Maintenance Mode RawValue set to False" & " @ " & Format(Timer, "###,##0.000")
        Else
            OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = True
'Debug.Print "Maintenance Mode RawValue set to True" & " @ " & Format(Timer, "###,##0.000")
        End If
        ' System Vacuum Switch DI
        OptoDIO(Com_DIO(icSystemVacSw).addr, Com_DIO(icSystemVacSw).chan).RawValue = True
    End If
    

        ' PAS Request In (PAS Local Control)
'        iFunc = icPASPowerOnIn
'        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = IIf((Com_DIO(iFunc).UseInverse), False, True)
'        OptoDIO(Com_DIO(iFunc).addr, Com_DIO(iFunc).chan).RawValue = True
    
    ' PAS RunLocalControl In (PAS Local Control)
    iFuncIn = icPASisRunningIn
    iFuncOut = icPurgeRequestOut
'    OptoDIO(Com_DIO(iFuncIn).addr, Com_DIO(iFuncIn).chan).RawValue = OptoDIO(Com_DIO(iFuncOut).addr, Com_DIO(iFuncOut).chan).RawValue
    OptoDIO(Com_DIO(iFuncIn).addr, Com_DIO(iFuncIn).chan).RawValue = True
    
    ' PAS Ready Out (PAS Local Control)
    iFuncIn = icPurgeReadyIn
    iFuncOut = icPASReadyOut
    OptoDIO(Com_DIO(iFuncIn).addr, Com_DIO(iFuncIn).chan).RawValue = OptoDIO(Com_DIO(iFuncOut).addr, Com_DIO(iFuncOut).chan).RawValue
    
    ' Map Common Digital I/O
    Map_ComDigitals
    ' Map PurgeAir Digital I/O
    For iPrg = 1 To NR_PRGAIR
        Map_PrgDigitals iPrg
    Next iPrg
    
    ' ********************
    ' COMMON ANALOG VALUES
    ' ********************
    ' PurgeAir Temp, Humidity, Baro
    If USINGSIMNOISE Then
        Com_AIO(acAmbTempSensor).EUValue = SysConfig.Temp_Target - 0.5 + Rnd + Sim_PasError(pasTEMPERATURE)
        Com_AIO(acAmbBaroSensor).EUValue = 1013# - 1.25 + (1.25 * Rnd)
        Com_AIO(acAmbHumiditySensor).EUValue = 40# - 0.25 + (0.5 * Rnd) + Sim_PasError(pasMOISTURE)
        MasterPagData.Temperature = SysConfig.Temp_Target - 0.5 + Rnd + Sim_PasError(pasTEMPERATURE)
        MasterPagData.Humidity = 40# - 0.25 + (0.5 * Rnd) + Sim_PasError(pasMOISTURE)
        MasterPagData.Moisture = RHtoGrains(Com_AIO(acAmbBaroSensor).EUValue, MasterPagData.Temperature, MasterPagData.Humidity) - 0.5 + Rnd
    Else
        Com_AIO(acAmbTempSensor).EUValue = SysConfig.Temp_Target + Sim_PasError(pasTEMPERATURE)
        Com_AIO(acAmbBaroSensor).EUValue = 1013#
        Com_AIO(acAmbHumiditySensor).EUValue = 40# + Sim_PasError(pasMOISTURE)
        MasterPagData.Temperature = SysConfig.Temp_Target + Sim_PasError(pasTEMPERATURE)
        MasterPagData.Humidity = 40# + Sim_PasError(pasMOISTURE)
        MasterPagData.Moisture = RHtoGrains(Com_AIO(acAmbBaroSensor).EUValue, MasterPagData.Temperature, MasterPagData.Humidity)
    End If
    If (USINGPASLOCALCONTROL Or (USINGDRYPURGEAIR And SysConfig.DryAirPurge)) Then
        Com_AIO(acPasTempSensor).EUValue = Com_AIO(acAmbTempSensor).EUValue + 0.05
        Com_AIO(acPasHumiditySensor).EUValue = Com_AIO(acAmbHumiditySensor).EUValue + 0.05
    End If
    Map_AIO(Com_AIO(acPasTempSensor).addr, Com_AIO(acPasTempSensor).chan).EUValue = Com_AIO(acPasTempSensor).EUValue
    Map_AIO(Com_AIO(acAmbTempSensor).addr, Com_AIO(acAmbTempSensor).chan).EUValue = Com_AIO(acAmbTempSensor).EUValue
    Map_AIO(Com_AIO(acAmbBaroSensor).addr, Com_AIO(acAmbBaroSensor).chan).EUValue = Com_AIO(acAmbBaroSensor).EUValue
    Map_AIO(Com_AIO(acAmbHumiditySensor).addr, Com_AIO(acAmbHumiditySensor).chan).EUValue = Com_AIO(acAmbHumiditySensor).EUValue
    Map_AIO(Com_AIO(acPasHumiditySensor).addr, Com_AIO(acPasHumiditySensor).chan).EUValue = Com_AIO(acPasHumiditySensor).EUValue
    If (USINGPASLOCALCONTROL Or (USINGDRYPURGEAIR And SysConfig.DryAirPurge)) Then
        PATemp = Com_AIO(acPasTempSensor).EUValue
        PAHum = Com_AIO(acPasHumiditySensor).EUValue
        PAMoisture = RHtoGrains(AmbBaro, PATemp, PAHum)
    ElseIf (LocalPagControl.Type = pagClient) Then
        PATemp = MasterPagData.Temperature
        PAHum = MasterPagData.Humidity
        PAMoisture = MasterPagData.Moisture
    Else
        PATemp = Com_AIO(acAmbTempSensor).EUValue
        PAHum = Com_AIO(acAmbHumiditySensor).EUValue
        PAMoisture = RHtoGrains(AmbBaro, PATemp, PAHum)
    End If
    AmbTemp = Com_AIO(acAmbTempSensor).EUValue
    AmbBaro = Com_AIO(acAmbBaroSensor).EUValue
    AmbHum = Com_AIO(acAmbHumiditySensor).EUValue
    AmbMoisture = RHtoGrains(AmbBaro, AmbTemp, AmbHum)
    
    Map_AIO(Com_AIO(acPASMoistCntrlOut).addr, Com_AIO(acPASMoistCntrlOut).chan).RawValue = OptoAIO(Com_AIO(acPASMoistCntrlOut).addr, Com_AIO(acPASMoistCntrlOut).chan).RawValue
    addr = Com_AIO(acPASMoistCntrlOut).addr
    chan = Com_AIO(acPASMoistCntrlOut).chan
    If (Map_AIO(addr, chan).VdcMax > Map_AIO(addr, chan).VdcMin) Then
        tempVdc = 10# * (Map_AIO(addr, chan).RawValue / FULLSCALE)          ' Vdc out of 0-10Vdc
        tempVdc = tempVdc - Map_AIO(addr, chan).VdcMin                      ' Vdc above VdcMin
        tempPerc = tempVdc / (Map_AIO(addr, chan).VdcMax - Map_AIO(addr, chan).VdcMin)
        tempPerc = CSng(100# * tempPerc)
    Else
        tempPerc = CSng(100# * (Map_AIO(addr, chan).RawValue / FULLSCALE))  ' % of 0-10Vdc
    End If
    Map_AIO(addr, chan).EUValue = tempPerc
    Com_AIO(acPASMoistCntrlOut).EUValue = Map_AIO(addr, chan).EUValue
    
    ' Common TC's
    If USINGCOMMONTC Then
        flag = False
        For iStn = 1 To LAST_STN
            For iShift = 1 To NR_SHIFT
                If StationControl(iStn, iShift).Mode = VBLOAD Then flag = True
            Next iShift
        Next iStn
        For iTC = 1 To 6
            iFunc = iTC - 1 + acCommonTC1
            If flag Then
                Com_AIO(iFunc).EUValue = Com_AIO(iFunc).EUValue + 0.0025
            Else
                Com_AIO(iFunc).EUValue = Com_AIO(iFunc).EUValue - 0.001
            End If
            If Com_AIO(iFunc).EUValue > 100 Then Com_AIO(iFunc).EUValue = 100
            If Com_AIO(iFunc).EUValue < 10 Then Com_AIO(iFunc).EUValue = 10
        Next iTC
    End If
    
    ' Pressure Sensor
    If LeakCheckControl.station = 0 Or LeakCheckControl.station <> Sim_LcPtUser Or LeakCheckControl.Phase = LeakPurging Then
        Com_AIO(acComnPressSensor).EUValue = 0#
        Sim_LcPtUser = LeakCheckControl.station
    ElseIf LeakCheckControl.Phase = LeakPressurizing Then
        Sim_LastPTcheckTimer = StationControl(LeakCheckControl.station, LeakCheckControl.Shift).TestTimer
        If Com_AIO(acComnPressSensor).EUValue < Com_AIO(acComnPressSensor).EuMax Then
            fVal2 = 0.05 * (StationControl(LeakCheckControl.station, LeakCheckControl.Shift).TestTimer - LeakCheckControl.PhaseStartTimer)
            Com_AIO(acComnPressSensor).EUValue = CSng(fVal2) * Com_AIO(acComnPressSensor).EuMax
        End If
    ElseIf LeakCheckControl.Phase = LeakTesting Then
        If StationControl(LeakCheckControl.station, LeakCheckControl.Shift).TestTimer > Sim_LastPTcheckTimer Then
            fVal1 = StationControl(LeakCheckControl.station, LeakCheckControl.Shift).TestTimer - Sim_LastPTcheckTimer     ' number of seconds
            fVal2 = Sim_LeakError(LeakCheckControl.station) / 60                                               ' percent leak per second
            fVal3 = fVal1 * fVal2 * (1 / 100) * StationConfig(LeakCheckControl.station, LeakCheckControl.Shift).LCSetPoint
            Com_AIO(acComnPressSensor).EUValue = PTinvalue - (CSng(fVal3))
        End If
        Sim_LastPTcheckTimer = StationControl(LeakCheckControl.station, LeakCheckControl.Shift).TestTimer
    End If
    PTinvalue = Com_AIO(acComnPressSensor).EUValue
    
    
    ' **********************
    ' STATION DIGITAL VALUES
    ' **********************
    For iStn = 1 To LAST_STN
        If ((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)) Then
            ' Live Fuel Tank AutoDrain&Fill
            SimulateADF iStn
        End If
        ' Map Station Digital I/O
        Map_StnDigitals iStn
        ' WaterBath
        If USINGWATERBATH Then
            If (Not ChillComOn) Then
                LF_Chiller.SpIn = WaterBathSP
                If USINGSIMNOISE Then
                    LF_Chiller.PvIn = WaterBathSP - 0.25 + (Rnd / 5)
                Else
                    LF_Chiller.PvIn = WaterBathSP
                End If
            End If
        End If
    Next iStn
    
    ' *********************
    ' STATION ANALOG VALUES
    ' *********************
    For iStn = 1 To LAST_STN
        iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
        ' Purge Oven
        If USINGPURGEOVEN Then
            If StationRecipe(iStn, iShift).PurgeOven Then
                Stn_AIO(iStn, asPurgeOvenTempSP).EUValue = StationRecipe(iStn, iShift).PurgeOvenSP
                If USINGSIMNOISE Then
                    Stn_AIO(iStn, asPurgeOvenTemp).EUValue = StationRecipe(iStn, iShift).PurgeOvenSP - 0.25 + (Rnd / 5)
                Else
                    Stn_AIO(iStn, asPurgeOvenTemp).EUValue = StationRecipe(iStn, iShift).PurgeOvenSP
                End If
            End If
        End If
        ' Station MFC's
        Select Case STN_INFO(iStn).Type
            Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                SimulateMFC CInt(iStn), asNitrogenFlowSP, asNitrogenFlow
                SimulateMFC CInt(iStn), asButaneFlowSP, asButaneFlow
                SimulateMFC CInt(iStn), asPurgeAirFlowSP, asPurgeAirFlow
                Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asNitrogenFlow).EUValue
                Stn_Btn_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asButaneFlow).EUValue
            Case STN_ORVR2_TYPE
                SimulateMFC CInt(iStn), asNitrogenFlowSP, asNitrogenFlow
                SimulateMFC CInt(iStn), asButaneFlowSP, asButaneFlow
                SimulateMFC CInt(iStn), asNitrogenORVRFlowSP, asNitrogenORVRFlow
                SimulateMFC CInt(iStn), asButaneORVRFlowSP, asButaneORVRFlow
                SimulateMFC CInt(iStn), asPurgeAirFlowSP, asPurgeAirFlow
                If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                    ' Use Higher Range MFC's
                    Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asNitrogenORVRFlow).EUValue
                    Stn_Btn_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asButaneORVRFlow).EUValue
                Else
                    ' Use Lower Range MFC's
                    Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asNitrogenFlow).EUValue
                    Stn_Btn_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asButaneFlow).EUValue
                End If
            Case STN_LIVEFUEL_TYPE
                SimulateMFC CInt(iStn), asLiveFuelVaporFlowSP, asLiveFuelVaporFlow
                SimulateMFC CInt(iStn), asPurgeAirFlowSP, asPurgeAirFlow
                Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asLiveFuelVaporFlow).EUValue
                Stn_Btn_Flow_PV(iStn, iShift) = 0
    '            If USINGSIMNOISE Then
    '                Stn_AIO(iStn, asFuelHeaterTemp).EUValue = StationRecipe(iStn, iShift).ADF_HeaterSP - 0.5 + Rnd
    '                Stn_AIO(iStn, asFuelTankTemp).EUValue = StationRecipe(iStn, iShift).ADF_HeaterSP - 0.25 + (Rnd / 5)
    '            Else
    '                Stn_AIO(iStn, asFuelHeaterTemp).EUValue = StationRecipe(iStn, iShift).ADF_HeaterSP
    '                Stn_AIO(iStn, asFuelTankTemp).EUValue = StationRecipe(iStn, iShift).ADF_HeaterSP
    '            End If
            Case STN_LIVEREG_TYPE
                SimulateMFC CInt(iStn), asNitrogenFlowSP, asNitrogenFlow
                SimulateMFC CInt(iStn), asButaneFlowSP, asButaneFlow
                SimulateMFC CInt(iStn), asPurgeAirFlowSP, asPurgeAirFlow
                SimulateMFC CInt(iStn), asLiveFuelVaporFlowSP, asLiveFuelVaporFlow
                If (StationRecipe(iStn, iShift).LiveFuel) Then
                    ' use Live Fuel
                    Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asLiveFuelVaporFlow).EUValue
                    Stn_Btn_Flow_PV(iStn, iShift) = 0
                Else
                    ' use Butane/Nitrogen
                    Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asNitrogenFlow).EUValue
                    Stn_Btn_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asButaneFlow).EUValue
                End If
            Case STN_LIVEORVR2_TYPE
                SimulateMFC CInt(iStn), asNitrogenFlowSP, asNitrogenFlow
                SimulateMFC CInt(iStn), asButaneFlowSP, asButaneFlow
                SimulateMFC CInt(iStn), asNitrogenORVRFlowSP, asNitrogenORVRFlow
                SimulateMFC CInt(iStn), asButaneORVRFlowSP, asButaneORVRFlow
                SimulateMFC CInt(iStn), asPurgeAirFlowSP, asPurgeAirFlow
                SimulateMFC CInt(iStn), asLiveFuelVaporFlowSP, asLiveFuelVaporFlow
                SimulateMFC CInt(iStn), asLiveFuelVaporORVRFlowSP, asLiveFuelVaporORVRFlow
                If (StationRecipe(iStn, iShift).LiveFuel) Then
                    ' use Live Fuel
                    If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                        ' Use Higher Range MFC's
                        Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asLiveFuelVaporORVRFlow).EUValue
                        Stn_Btn_Flow_PV(iStn, iShift) = 0
                    Else
                        ' Use Lower Range MFC's
                        Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asLiveFuelVaporFlow).EUValue
                        Stn_Btn_Flow_PV(iStn, iShift) = 0
                    End If
                Else
                    ' use Butane/Nitrogen
                    If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                        ' Use Higher Range MFC's
                        Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asNitrogenORVRFlow).EUValue
                        Stn_Btn_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asButaneORVRFlow).EUValue
                    Else
                        ' Use Lower Range MFC's
                        Stn_Nit_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asNitrogenFlow).EUValue
                        Stn_Btn_Flow_PV(iStn, iShift) = Stn_AIO(iStn, asButaneFlow).EUValue
                    End If
                End If
            Case STN_COMBO3_TYPE
                'future
            
            Case STN_LEAKTEST_TYPE
                UpdateLeakInputs iStn
                
        End Select
        
    Next iStn
    
    firstPassSim = True
    OptoReadAllOnce = True

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

Sub SimulateMFC(iStn As Integer, iSPfunc As Integer, iPVfunc As Integer)
Dim newraw As Long
Dim craw, clim, neweu, newsp As Single
Dim cmax, cmin, cspan, ctemp As Single
Dim emax, emin, espan, etemp As Single
Dim erreu, enoise As Single
Dim Control As Boolean

SetErrModule 8, 1011
If UseLocalErrorHandler Then On Error GoTo localhandler

clim = CSng(FULLSCALE)
Control = True

If Stn_AIO(iStn, iSPfunc).addr <> 0 Or Stn_AIO(iStn, iSPfunc).chan <> 0 Then

        newraw = OptoAIO(Stn_AIO(iStn, iSPfunc).addr, Stn_AIO(iStn, iSPfunc).chan).RawValue
        craw = CSng(newraw)
        
        ' SP's & PV's are Linear, MinMax of 0-10 Vdc Input
        cmax = clim * (Stn_AIO(iStn, iSPfunc).VdcMax / 10#)
        cmin = clim * (Stn_AIO(iStn, iSPfunc).VdcMin / 10#)
        cspan = cmax - cmin
        emax = Stn_AIO(iStn, iSPfunc).EuMax
        emin = Stn_AIO(iStn, iSPfunc).EuMin
        espan = emax - emin
        ' optional max 1% noise
        enoise = IIf(USINGSIMNOISE, espan * (-0.005 + (0.01 * Rnd)), 0)
        If cspan <> 0# Then
            Select Case iSPfunc            ' Note: Only AIs & AOs for MFCs are Calibrated
                Case asNitrogenFlowSP
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCNITROGEN, Stn_MfcCal(iStn, MFCNITROGEN))
                    erreu = espan * ((Sim_MfcError(iStn, MFCNITROGEN) / 100))
                    neweu = erreu + newsp + enoise
                 Case asButaneFlowSP
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCBUTANE, Stn_MfcCal(iStn, MFCBUTANE))
                    erreu = espan * ((Sim_MfcError(iStn, MFCBUTANE) / 100))
                    neweu = erreu + newsp + enoise
                Case asNitrogenORVRFlowSP   ' shares SimError with MFCNITROGEN
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCORVRNIT, Stn_MfcCal(iStn, MFCORVRNIT))
                    erreu = espan * ((Sim_MfcError(iStn, MFCNITROGEN) / 100))
                    neweu = erreu + newsp + enoise
                 Case asButaneORVRFlowSP   ' shares SimError with MFCBUTANE
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCORVRBUT, Stn_MfcCal(iStn, MFCORVRBUT))
                    erreu = espan * ((Sim_MfcError(iStn, MFCBUTANE) / 100))
                    neweu = erreu + newsp + enoise
                Case asPurgeAirFlowSP
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCPURGEAIR, Stn_MfcCal(iStn, MFCPURGEAIR))
                    erreu = espan * ((Sim_MfcError(iStn, MFCPURGEAIR) / 100))
                    neweu = erreu + newsp + enoise
                Case asLiveFuelVaporFlowSP
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCLIVEFUEL, Stn_MfcCal(iStn, MFCLIVEFUEL))
                    erreu = espan * ((Sim_MfcError(iStn, MFCLIVEFUEL) / 100))
                    neweu = erreu + newsp + enoise
                 Case asLiveFuelVaporORVRFlowSP   ' shares SimError with MFCLIVEFUEL
                    ctemp = (craw - cmin) / cspan
                    etemp = emin + (espan * ctemp)
                    newsp = Cal_MfcInput(ctemp, iStn, MFCORVRLIVE, Stn_MfcCal(iStn, MFCORVRLIVE))
                    erreu = espan * ((Sim_MfcError(iStn, MFCLIVEFUEL) / 100))
                    neweu = erreu + newsp + enoise
                Case Else
                    newsp = emin + (espan * ((craw - cmin) / cspan))
                    neweu = emin + (espan * ((craw - cmin) / cspan)) + enoise
            End Select
        Else
            neweu = 0#
        End If
            
        Stn_AIO(iStn, iSPfunc).EUValue = newsp
        Stn_AIO(iStn, iPVfunc).EUValue = neweu
                            
End If

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

Sub SimulateScales()
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 1012
Dim iStn, iShift, iScale As Integer
Dim purgePriDensityFactor As Single
Dim purgeAuxDensityFactor As Single
Dim purgeLitersEmpty As Single
Dim purgeLitersMax As Single
Dim purgeWtChgFraction As Single
Dim deltaLiters As Single
Dim deltaWeight As Single
Dim loadTotalGrams As Single
Dim maxWtChg As Single
Dim netWtChg As Single
Dim newweight As Single
Dim CanWC As Single
Dim loadBreakthruWt As Single
Dim loadBreakthruPriPortion As Single
Dim loadBreakthruAuxPortion As Single

    If Not ScalesReadAllOnce Then
        frmAbout.UpdateMsg "Starting Scale Simulation" & vbCrLf
    End If
    
    ' Scale OK
    For iScale = 1 To NR_SCALES
        Scale_OK(iScale) = True
    Next iScale
    
    ' Scale Values
    For iStn = 1 To LAST_STN
        iShift = IIf((Stn_ActiveShift(iStn) > 0), Stn_ActiveShift(iStn), 1)
        If (iShift > 4) Then iShift = 1
        
        ' set calculation factors
        CanWC = IIf((StationCanister(iStn, iShift).WorkingCapacity > 0), StationCanister(iStn, iShift).WorkingCapacity, (DefCanVol2CanWcMult * StationCanister(iStn, iShift).WorkingVolume))
        loadBreakthruWt = 0.975 * CanWC
        loadBreakthruAuxPortion = 0.825
        loadBreakthruPriPortion = CSng(1) - loadBreakthruAuxPortion
        purgeLitersMax = CSng(303) * StationCanister(iStn, iShift).WorkingVolume
        purgeLitersEmpty = CSng(300) * StationCanister(iStn, iShift).WorkingVolume
        purgePriDensityFactor = 1.935
        purgeAuxDensityFactor = 1.625
        
        ' Scales & ModeStart
        If Sim_Mode(iStn, iShift) <> StationControl(iStn, iShift).Mode Then Sim_ModeStartComplete(iStn, iShift) = False
        Select Case StationControl(iStn, iShift).Mode
            Case VBLOAD
                If LoadControl(iStn, iShift).Phase = LoadLoading Then
                    ' Load Cycle
                    If Sim_Cycle(iStn, iShift) <> StationControl(iStn, iShift).CurrCycle Then Sim_ModeStartComplete(iStn, iShift) = False
        
                    ' need to simulate live fuel loads
                    If ((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE And StationRecipe(iStn, iShift).LiveFuel) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE And StationRecipe(iStn, iShift).LiveFuel)) Then
                        ' Calculate the fuel vapor used (in grams) during this load
                        loadTotalGrams = LoadControl(iStn, iShift).LoadTotalLiters * Sim_LiveFuelDensity
                    Else
                        loadTotalGrams = LoadControl(iStn, iShift).loadTotalGrams
                    End If
                    
                    ' Primary Scale
                    If StationRecipe(iStn, iShift).UsePriScale Then
                        iScale = StationRecipe(iStn, iShift).PriScaleNo
                        Scale_OK(iScale) = True
                        ' determine weight change
                        netWtChg = Sim_PriWt_Current(iStn) - LoadControl(iStn, iShift).PriWt_Start
                        Select Case netWtChg
                            Case Is <= loadBreakthruWt
                                Sim_PriWt_Current(iStn) = loadTotalGrams + LoadControl(iStn, iShift).PriWt_Start
                            Case Else
                                Sim_PriWt_Current(iStn) = loadBreakthruWt + (loadBreakthruPriPortion * (loadTotalGrams - loadBreakthruWt)) + LoadControl(iStn, iShift).PriWt_Start
                        End Select
                        newweight = CSng(0.1) * CSng(CLng(CSng(10) * Sim_PriWt_Current(iStn)))
                        If (Sim_PriWt(iStn) <> newweight) Then
                            Sim_PriWt(iStn) = newweight
                            Scale_Weight(iScale) = Sim_PriWt(iStn)
                            Scale_Value(iScale) = CStr(Scale_Weight(iScale))
                            ' Optionally, log New Reading
                            If Not NotDebugSCALES Then
                                Write_Zlog_Scales iScale, 0, Scale_Value(iScale), "Simulated Scale #" & Format(iScale, "#0")
                            End If
                        End If
                    End If
                    ' Aux Scale
                    If StationRecipe(iStn, iShift).UseAuxScale Then
                        iScale = StationRecipe(iStn, iShift).AuxScaleNo
                        Scale_OK(iScale) = True
                        netWtChg = Sim_AuxWt_Current(iStn) - LoadControl(iStn, iShift).AuxWt_Start
                        Sim_LastLoadAuxWtChg(iStn) = netWtChg
                        If (loadTotalGrams > loadBreakthruWt) Then
                            Sim_AuxWt_Current(iStn) = (loadBreakthruAuxPortion * (loadTotalGrams - loadBreakthruWt)) + LoadControl(iStn, iShift).AuxWt_Start
                        End If
                        newweight = CSng(0.1) * CSng(CLng(CSng(10) * Sim_AuxWt_Current(iStn)))
                        If (Sim_AuxWt(iStn) <> newweight) Then
                            Sim_AuxWt(iStn) = newweight
                            Scale_Weight(iScale) = Sim_AuxWt(iStn)
                            Scale_Value(iScale) = CStr(Scale_Weight(iScale))
                            ' Optionally, log New Reading
                            If Not NotDebugSCALES Then
                                Write_Zlog_Scales iScale, 0, Scale_Value(iScale), "Simulated Scale #" & Format(iScale, "#0")
                            End If
                        End If
                    End If
                    Sim_ModeStartComplete(iStn, iShift) = True
                    Sim_Cycle(iStn, iShift) = StationControl(iStn, iShift).CurrCycle
                End If
            Case VBPURGE
                If PurgeControl(iStn, iShift).Phase = PurgePurging Then
                    ' Purge Cycle
                    purgeWtChgFraction = CSng(1) + (CSng(0.275) / StationControl(iStn, iShift).CurrCycle)
                    ' check if first pass
                    If Sim_Cycle(iStn, iShift) <> StationControl(iStn, iShift).CurrCycle Then Sim_ModeStartComplete(iStn, iShift) = False
                    ' Primary Scale
                    If StationRecipe(iStn, iShift).UsePriScale Then
                        iScale = StationRecipe(iStn, iShift).PriScaleNo
                        If (PurgeControl(iStn, iShift).Purge_Total > 0) Then
                            If ((StationControl(iStn, iShift).CompletedCycles > 0) Or (StationControl(iStn, iShift).Course > 1)) Then
                                maxWtChg = purgeWtChgFraction * CanWC
                            Else
                                maxWtChg = PurgeControl(iStn, iShift).PriWt_Start
                            End If
                            If (maxWtChg < 0) Then maxWtChg = 0
                            netWtChg = PurgeControl(iStn, iShift).PriWt_Start - Sim_PriWt_Current(iStn)
                            If (maxWtChg = 0) Then
                                deltaLiters = CSng(0)
                                deltaWeight = CSng(0)
                            Else
                                If (PurgeControl(iStn, iShift).Purge_Total <> Purge_Total_Last(iStn, iShift)) Then
                                    deltaLiters = PurgeControl(iStn, iShift).Purge_Total - Purge_Total_Last(iStn, iShift)
                                Else
                                    deltaLiters = CSng(0)
                                End If
                                Select Case netWtChg
                                    Case Is <= maxWtChg
                                        'deltaWeight = deltaLiters * ((purgeLitersMax - PurgeControl(iStn, iShift).Purge_Total) / purgeLitersMax)
                                        deltaWeight = purgePriDensityFactor * ((deltaLiters / purgeLitersEmpty) * maxWtChg) * ((purgeLitersMax - PurgeControl(iStn, iShift).Purge_Total) / purgeLitersMax)
                                    Case Else
                                        deltaWeight = CSng(0)
                                End Select
                            End If
                            If (deltaWeight < CSng(0)) Then deltaWeight = CSng(-1) * deltaWeight
                            Sim_PriWt_Current(iStn) = Sim_PriWt_Current(iStn) - deltaWeight
                        Else
    '                        fVal1 = 0
    '                        Scale_Weight(iScale) = Sim_PriWt_Start(iStn, iShift)
                        End If
                        newweight = CSng(0.1) * CSng(CLng(CSng(10) * Sim_PriWt_Current(iStn)))
                        If (Sim_PriWt(iStn) <> newweight) Then
                            Sim_PriWt(iStn) = newweight
                            Scale_Weight(iScale) = Sim_PriWt(iStn)
                            Scale_Value(iScale) = CStr(Scale_Weight(iScale))
                            ' Optionally, log New Reading
                            If Not NotDebugSCALES Then
                                Write_Zlog_Scales iScale, 0, Scale_Value(iScale), "Simulated Scale #" & Format(iScale, "#0")
                            End If
                        End If
                    End If
                    ' Aux Scale
                    If StationRecipe(iStn, iShift).UseAuxScale Then
                        iScale = StationRecipe(iStn, iShift).AuxScaleNo
                        If (PurgeControl(iStn, iShift).Purge_Total > 0) Then
                            If ((StationControl(iStn, iShift).CompletedCycles > 0) Or (StationControl(iStn, iShift).Course > 1)) Then
                                maxWtChg = purgeWtChgFraction * Sim_LastLoadAuxWtChg(iStn)
                            Else
                                maxWtChg = PurgeControl(iStn, iShift).AuxWt_Start
                            End If
                            If (maxWtChg < 0) Then maxWtChg = 0
                            netWtChg = PurgeControl(iStn, iShift).AuxWt_Start - Sim_AuxWt_Current(iStn)
                            If (maxWtChg = 0) Then
                                deltaLiters = CSng(0)
                                deltaWeight = CSng(0)
                            Else
                                deltaLiters = PurgeControl(iStn, iShift).Purge_Total - Purge_Total_Last(iStn, iShift)
                                Select Case netWtChg
                                    Case Is <= maxWtChg
                                        deltaWeight = purgeAuxDensityFactor * ((deltaLiters / purgeLitersEmpty) * maxWtChg) * ((purgeLitersMax - PurgeControl(iStn, iShift).Purge_Total) / purgeLitersMax)
                                    Case Else
                                        deltaWeight = CSng(0)
                                End Select
                            End If
                            If (deltaWeight < CSng(0)) Then deltaWeight = CSng(-1) * deltaWeight
                            Sim_AuxWt_Current(iStn) = Sim_AuxWt_Current(iStn) - deltaWeight
                        Else
    '                        fVal1 = 0
    '                        Scale_Weight(iScale) = Sim_PriWt_Start(iStn, iShift)
                        End If
                        newweight = CSng(0.1) * CSng(CLng(CSng(10) * Sim_AuxWt_Current(iStn)))
                        If (Sim_AuxWt(iStn) <> newweight) Then
                            Sim_AuxWt(iStn) = newweight
                            Scale_Weight(iScale) = Sim_AuxWt(iStn)
                            Scale_Value(iScale) = CStr(Scale_Weight(iScale))
                            ' Optionally, log New Reading
                            If Not NotDebugSCALES Then
                                Write_Zlog_Scales iScale, 0, Scale_Value(iScale), "Simulated Scale #" & Format(iScale, "#0")
                            End If
                        End If
                    End If
                    Sim_ModeStartComplete(iStn, iShift) = True
                    Sim_Cycle(iStn, iShift) = StationControl(iStn, iShift).CurrCycle
                End If
            Case VBLEAK
                ' Leak Check
                Sim_ModeStartComplete(iStn, iShift) = False
            Case VBIDLE
                ' Idle
                Sim_ModeStartComplete(iStn, iShift) = False
                ' Primary Scale
                If StationRecipe(iStn, iShift).UsePriScale Then
                    iScale = StationRecipe(iStn, iShift).PriScaleNo
                    Sim_PriWt(iStn) = Scale_Weight(iScale)
                    Sim_PriWt_Current(iStn) = Scale_Weight(iScale)
                End If
                ' Aux Scale
                If StationRecipe(iStn, iShift).UseAuxScale Then
                    iScale = StationRecipe(iStn, iShift).AuxScaleNo
                    Sim_AuxWt(iStn) = Scale_Weight(iScale)
                    Sim_AuxWt_Current(iStn) = Scale_Weight(iScale)
                End If
            Case Else
                ' Primary Scale
                If StationRecipe(iStn, iShift).UsePriScale Then
                    iScale = StationRecipe(iStn, iShift).PriScaleNo
                    Sim_PriWt(iStn) = CSng(0.1) * CSng(CLng(CSng(10) * Sim_PriWt_Current(iStn)))
                    Scale_Weight(iScale) = Sim_PriWt(iStn)
                    Scale_Value(iScale) = CStr(Scale_Weight(iScale))
                End If
                ' Aux Scale
                If StationRecipe(iStn, iShift).UseAuxScale Then
                    iScale = StationRecipe(iStn, iShift).AuxScaleNo
                    Sim_AuxWt(iStn) = CSng(0.1) * CSng(CLng(CSng(10) * Sim_AuxWt_Current(iStn)))
                    Scale_Weight(iScale) = Sim_AuxWt(iStn)
                    Scale_Value(iScale) = CStr(Scale_Weight(iScale))
                End If
                Sim_ModeStartComplete(iStn, iShift) = False
        End Select
        Sim_PriWt_Last(iStn) = Sim_PriWt_Current(iStn)
        Sim_AuxWt_Last(iStn) = Sim_AuxWt_Current(iStn)
        Purge_Total_Last(iStn, iShift) = PurgeControl(iStn, iShift).Purge_Total
        Sim_Mode(iStn, iShift) = StationControl(iStn, iShift).Mode
    
    Next iStn
    
    ScalesReadAllOnce = True
    
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

Private Sub SimulateADF(ByVal iStn As Integer)
Static ThisTimer As Double
Static LastTimer(0 To MAX_STN) As Double
Dim DeltaTimer As Double

    ' 111111111111111111111111111111111111111
    ' initialize values at startup, then exit
    ' 111111111111111111111111111111111111111
    If (Not OptoReadAllOnce) Then
        ' tank is not pressurized, is empty  and is at ambient temp
        ADFsim_StorageLevel(iStn) = 50.5
        ADFsim_TankLevel(iStn) = 0
        ADFsim_TankPressure(iStn) = 0
        ADFsim_TankTemperature(iStn) = PATemp
        ADFsim_TankSheathTemp(iStn) = ADFsim_TankTemperature(iStn)
        ADFsim_PVoverT(iStn) = (ADFsim_TankPressure(iStn) * (100 - ADFsim_TankLevel(iStn))) / ADFsim_TankTemperature(iStn)
        LastTimer(iStn) = Timer
        Exit Sub
    End If
    
       
    ' ----------
    ' delta time
    ' ----------
    ThisTimer = Timer
    DeltaTimer = ThisTimer - LastTimer(iStn)
    If DeltaTimer < 0 Then DeltaTimer = 0.1
    LastTimer(iStn) = Timer
    
    ' ******************
    ' STORAGE TANK LEVEL
    ' ******************
    ' update level
    If STN_INFO(iStn).ADF_DEF.hasADF_FST Then
        If Stn_DIO(iStn, isFuelFillSol).Value Then
            ' transfer in progress
            If ADFsim_StorageLevel(iStn) > 0 Then
                ' transfer rate is 0.5%/sec (i.e. transfer all in 200 sec.)
                ADFsim_StorageLevel(iStn) = ADFsim_StorageLevel(iStn) - (DeltaTimer * 0.5)
                If ADFsim_StorageLevel(iStn) < 0 Then ADFsim_StorageLevel(iStn) = 0
            End If
        ElseIf (Stn_DIO(DispStn, isStorageDrainSol).Value And Stn_DIO(DispStn, isFuelPumpMotor).Value) Then
            ' drain-to-waste in progress
            If ADFsim_StorageLevel(iStn) > 0 Then
                ' drain rate is 1%/sec (i.e. drain all in 100 sec.)
                ADFsim_StorageLevel(iStn) = ADFsim_StorageLevel(iStn) - (DeltaTimer * 1#)
                If ADFsim_StorageLevel(iStn) < 0 Then ADFsim_StorageLevel(iStn) = 0
            End If
        ElseIf Stn_DIO(iStn, isStorageFillSol).Value Then
            ' fill in progress
            If ADFsim_StorageLevel(iStn) < 100 Then
                ' fill rate is 1.66%/sec (i.e. fill in 60 sec.)
                ADFsim_StorageLevel(iStn) = ADFsim_StorageLevel(iStn) + (DeltaTimer * 1.66)
                If ADFsim_StorageLevel(iStn) > 100 Then ADFsim_StorageLevel(iStn) = 100
            End If
        End If
    Else
        ADFsim_StorageLevel(iStn) = 50
    End If
    Stn_AIO(iStn, asStorageTankLevel).EUValue = ADFsim_StorageLevel(iStn)
    ' update Level Switches
    OptoDIO(Stn_DIO(iStn, isStorageLowLevelLS).addr, Stn_DIO(iStn, isStorageLowLevelLS).chan).RawValue = IIf(ADFsim_StorageLevel(iStn) > 10, True, False)
    OptoDIO(Stn_DIO(iStn, isStorageHiHiLevelLS).addr, Stn_DIO(iStn, isStorageHiHiLevelLS).chan).RawValue = IIf(ADFsim_StorageLevel(iStn) > 85, IIf(Stn_DIO(iStn, isStorageHiHiLevelLS).UseInverse, False, True), IIf(Stn_DIO(iStn, isStorageHiHiLevelLS).UseInverse, cON, cOFF))
    
    ' ****************
    ' VAPOR TANK LEVEL
    ' ****************
    ' update level
'    If Stn_DIO(iStn, isFuelPumpMotor).Value Then
        If ((Stn_DIO(iStn, isFuelDrainSol).Value) And (Not Stn_DIO(iStn, isFuelRecircSol).Value)) Then
            ' drain in progress
            If ADFsim_TankLevel(iStn) > 0 Then
                ' drain rate is 5%/sec (i.e. drain in 20 sec.)
                ADFsim_TankLevel(iStn) = ADFsim_TankLevel(iStn) - (DeltaTimer * 5#)
                If ADFsim_TankLevel(iStn) < 0 Then ADFsim_TankLevel(iStn) = 0
            End If
        ElseIf (Stn_DIO(iStn, isFuelFillSol).Value) Then
            ' fill in progress
            If ((ADFsim_TankLevel(iStn) < 100) And ((ADFsim_StorageLevel(iStn) > 0) Or (Not STN_INFO(iStn).ADF_DEF.hasADF_FST))) Then
                ' fill rate is 3.33%/sec (i.e. fill in 30 sec.)
                ADFsim_TankLevel(iStn) = ADFsim_TankLevel(iStn) + (DeltaTimer * 3.33)
                If ADFsim_TankLevel(iStn) > 100 Then ADFsim_TankLevel(iStn) = 100
            End If
        End If
'    End If
    Stn_AIO(iStn, asFuelTankLevel).EUValue = ADFsim_TankLevel(iStn)
    ' update Level Switches
    
'If iStn = 1 Then Debug.Print "HiHi Raw = " & IIf(OptoDIO(Stn_DIO(iStn, isFuelHiHiLevelLS).addr, Stn_DIO(iStn, isFuelHiHiLevelLS).chan).RawValue, "True", "False") & " @ " & Format(Timer, "###,##0.000")
    
    OptoDIO(Stn_DIO(iStn, isFuelLowLevelLS).addr, Stn_DIO(iStn, isFuelLowLevelLS).chan).RawValue = IIf(ADFsim_TankLevel(iStn) > 10, True, False)
    OptoDIO(Stn_DIO(iStn, isFuelSafetyLevelLS).addr, Stn_DIO(iStn, isFuelSafetyLevelLS).chan).RawValue = IIf(ADFsim_TankLevel(iStn) > 40, True, False)
    OptoDIO(Stn_DIO(iStn, isFuelHighLevelLS).addr, Stn_DIO(iStn, isFuelHighLevelLS).chan).RawValue = IIf(ADFsim_TankLevel(iStn) > 65, IIf(Stn_DIO(iStn, isFuelHighLevelLS).UseInverse, False, True), IIf(Stn_DIO(iStn, isFuelHighLevelLS).UseInverse, True, False))
    OptoDIO(Stn_DIO(iStn, isFuelHiHiLevelLS).addr, Stn_DIO(iStn, isFuelHiHiLevelLS).chan).RawValue = IIf(ADFsim_TankLevel(iStn) > 95, IIf(Stn_DIO(iStn, isFuelHiHiLevelLS).UseInverse, False, True), IIf(Stn_DIO(iStn, isFuelHiHiLevelLS).UseInverse, True, False))

'If iStn = 1 Then Debug.Print "SimLevel = " & Format(ADFsim_TankLevel(iStn), "#,##0.00") & " @ " & Format(Timer, "###,##0.000")
'If iStn = 1 Then Debug.Print "Sim HiHi = " & IIf(OptoDIO(Stn_DIO(iStn, isFuelHiHiLevelLS).addr, Stn_DIO(iStn, isFuelHiHiLevelLS).chan).RawValue, "True", "False") & " @ " & Format(Timer, "###,##0.000")
    
    ' **********************
    ' VAPOR TANK TEMPERATURE
    ' **********************
    ' update temperatures
    If Stn_DIO(iStn, isFuelHeaterSSR).Value Then
        ' heater is on
        If Stn_DIO(iStn, isFuelSafetyLevelLS).Value Then
            If Stn_DIO(iStn, isFuelPumpMotor).Value Then
                ' there is circulation
                ' Tank heat rate is 0.1 deg/sec
                ADFsim_TankTemperature(iStn) = ADFsim_TankTemperature(iStn) + (DeltaTimer * 0.1)
                ' Sheath heat rate is 0.25 deg/sec
                ADFsim_TankSheathTemp(iStn) = ADFsim_TankSheathTemp(iStn) + (DeltaTimer * 0.25)
            Else
                ' no circulation
                ' Tank heat rate is 0.05 deg/sec
                ADFsim_TankTemperature(iStn) = ADFsim_TankTemperature(iStn) + (DeltaTimer * 0.05)
                ' Sheath heat rate is 5.0 deg/sec
                ADFsim_TankSheathTemp(iStn) = ADFsim_TankSheathTemp(iStn) + (DeltaTimer * 5#)
            End If
        Else
            ' tc & heater are exposed
            ADFsim_TankTemperature(iStn) = PATemp
            ' Sheath heat rate is 10.0 deg/sec
            ADFsim_TankSheathTemp(iStn) = ADFsim_TankSheathTemp(iStn) + (DeltaTimer * 10#)
        End If
    Else
        If Stn_DIO(iStn, isFuelSafetyLevelLS).Value Then
            ' level is above min and heater is off
            ' Tank cool rate is 0.005 deg/sec
            ADFsim_TankTemperature(iStn) = ADFsim_TankTemperature(iStn) - (DeltaTimer * 0.005)
            If ADFsim_TankTemperature(iStn) < PATemp Then ADFsim_TankTemperature(iStn) = PATemp
            ADFsim_TankSheathTemp(iStn) = ADFsim_TankTemperature(iStn)
        Else
            ' level is below min and heater is off
            ADFsim_TankTemperature(iStn) = PATemp
            ADFsim_TankSheathTemp(iStn) = ADFsim_TankTemperature(iStn)
        End If
    End If
    Stn_AIO(iStn, asFuelTankTemp).EUValue = ADFsim_TankTemperature(iStn)
    Stn_AIO(iStn, asFuelHeaterTemp).EUValue = ADFsim_TankSheathTemp(iStn)
    ' update OverTemp Switch
    OptoDIO(Stn_DIO(iStn, isFuelOverTempSw).addr, Stn_DIO(iStn, isFuelOverTempSw).chan).RawValue = IIf(ADFsim_TankTemperature(iStn) > 150, True, False)
    
    
    ' *******************
    ' VAPOR TANK PRESSURE
    ' *******************
    If Stn_DIO(iStn, isFuelPressSol).Value Then
        ' Pressure increasing at 6%/sec (i.e. done in 16 sec.)
        ADFsim_TankPressure(iStn) = ADFsim_TankPressure(iStn) + (DeltaTimer * 6#)
        If ADFsim_TankPressure(iStn) > 100 Then ADFsim_TankPressure(iStn) = 100
        ADFsim_PVoverT(iStn) = (ADFsim_TankPressure(iStn) * (100 - ADFsim_TankLevel(iStn))) / ADFsim_TankTemperature(iStn)
    Else
        ' PV = nrT
        If ADFsim_TankLevel(iStn) < 100 Then
            ADFsim_TankPressure(iStn) = (ADFsim_PVoverT(iStn) * ADFsim_TankTemperature(iStn)) / (100 - ADFsim_TankLevel(iStn))
            If ADFsim_TankPressure(iStn) < 0 Then ADFsim_TankPressure(iStn) = 0
            If ADFsim_TankPressure(iStn) > 100 Then ADFsim_TankPressure(iStn) = 100
            ADFsim_PVoverT(iStn) = (ADFsim_TankPressure(iStn) * (100 - ADFsim_TankLevel(iStn))) / ADFsim_TankTemperature(iStn)
        End If
    End If
    ' update Pressure Switch
    Select Case STN_INFO(iStn).ADF_TANKTYPE
        Case 12
            ' Mahle
            OptoDIO(Stn_DIO(iStn, isFuelPressPS).addr, Stn_DIO(iStn, isFuelPressPS).chan).RawValue = IIf(ADFsim_TankPressure(iStn) > 50, True, False)
        Case 20
            ' Stant
            If Stn_DIO(iStn, isFuelVentSol).Value Then
                OptoDIO(Stn_DIO(iStn, isFuelPressPS).addr, Stn_DIO(iStn, isFuelPressPS).chan).RawValue = False
            Else
                OptoDIO(Stn_DIO(iStn, isFuelPressPS).addr, Stn_DIO(iStn, isFuelPressPS).chan).RawValue = True
            End If
    End Select
End Sub

Sub AIR_Check(ByVal Idx As Integer)
'
'   Checks AIR Parameter for within limits (or not) when not using Local PAS Control
'
Dim lolimit, hilimit, currVal As Single
Dim sStr As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 13813
    Select Case Idx
        Case pasTEMPERATURE
            ' Temperature
            currVal = PATemp
            lolimit = SysConfig.Temp_Target - SysConfig.Tol_Temp
            hilimit = SysConfig.Temp_Target + SysConfig.Tol_Temp
        Case pasMOISTURE
            ' Moisture
            currVal = PAMoisture
            lolimit = SysConfig.Moisture_Target - SysConfig.Tol_Moisture
            hilimit = SysConfig.Moisture_Target + SysConfig.Tol_Moisture
    End Select
    If (currVal > lolimit And currVal < hilimit) Then
        ' currently within limits
        If Not PAS_INFO(Idx).Ok Then
            PAS_INFO(Idx).Ok = True
            Select Case Idx
                Case pasTEMPERATURE
                    ' Temperature
                    If USINGC Then sStr = " deg C"
                    If USINGF Then sStr = " deg F"
                    sStr = "PAS Temperature of " & Format(PATemp, "##0.0") & sStr & " is now within tolerance limits"
                Case pasMOISTURE
                    ' Moisture
                    If USINGMoist_RH Then sStr = " % rH"
                    If USINGMoist_Grains Then sStr = " grains/lb"
                    sStr = "PAS Moisture of " & Format(PAMoisture, "##0.0") & sStr & " is now within tolerance limits"
            End Select
            Write_ELog sStr
        End If
    Else
        ' Not currently within limits
        If PAS_INFO(Idx).Ok Then
            PAS_INFO(Idx).Ok = False
            Select Case Idx
                Case pasTEMPERATURE
                    ' Temperature
                    If USINGC Then sStr = " deg C"
                    If USINGF Then sStr = " deg F"
                    sStr = "PAS Temperature of " & Format(PATemp, "##0.0") & sStr & " is outside tolerance limits"
                Case pasMOISTURE
                    ' Moisture
                    If USINGMoist_RH Then sStr = " % rH"
                    If USINGMoist_Grains Then sStr = " grains/lb"
                    sStr = "PAS Moisture of " & Format(PAMoisture, "##0.0") & sStr & " is outside tolerance limits"
            End Select
            Write_ELog sStr
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

Sub PurgeProfile_Controller(ByVal iStation As Integer, ByVal iShift As Integer)
' Routine Name:  PurgeProfile_Controller
' Author:        MMW
' Description:
' Controls PurgeByProfile
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 3030
Dim span As Single
Dim temptime As Date
Dim tempSec, tempMin As Integer
Dim Nitrogen_Rate As Single
Dim Nitrogen_Output As Single


    For iStation = 1 To LAST_STN
    
        For iShift = 1 To NR_SHIFT
        
        
        Next iShift
        
    Next iStation
    
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

Public Sub UpdateErrorStatus()
'************************************************************
'
'   Update SCSII Error Status for AK Command Interface
'
'************************************************************
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 8, 6464
Dim Idx As Integer
Dim flag As Boolean
        
    ' update current error status
    ErrorStatus_ScratchPad = ErrorStatus_NoErrors
    If USINGPASLOCALCONTROL Then
        With ErrorStatus_ScratchPad
    '        If PAS_INFO(pasTEMPERATURE).OOT Then
    '            .TempOOT = True
    '            .AnyError = True
    '        End If
            If PAS_INFO(pasTEMPERATURE).timeOut Then
                .TempTO = True
                .AnyError = True
            End If
    '        If PAS_INFO(pasMOISTURE).OOT Then
    '            .MoistOOT = True
    '            .AnyError = True
    '        End If
            If PAS_INFO(pasMOISTURE).timeOut Then
                .MoistTO = True
                .AnyError = True
            End If
        End With
    End If
    ErrorStatus_Current = ErrorStatus_ScratchPad
            
    ' any changes in current error status ??
    flag = False
    With ErrorStatus_Current
        If .TempOOT <> ErrorStatus_Last.TempOOT Then flag = True
        If .TempTO <> ErrorStatus_Last.TempTO Then flag = True
        If .MoistOOT <> ErrorStatus_Last.MoistOOT Then flag = True
        If .MoistTO <> ErrorStatus_Last.MoistTO Then flag = True
        If .TestBit <> ErrorStatus_Last.TestBit Then flag = True
    End With
    If flag Then
        If ErrorStatus_Current.AnyError Then
            ' new combination of error(s)
            ErrorValue_Current = IIf((ErrorValue_Current < 9), (ErrorValue_Current + 1), 1)
        Else
            ' no errors at this time
            ErrorValue_Current = 0
        End If
    End If
    
    ' remember error status
    ErrorStatus_Last = ErrorStatus_Current
            
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




