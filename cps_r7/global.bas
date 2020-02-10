Attribute VB_Name = "Module15"
'error module 15 ''''''''''''''''''' program GLOBAL.bas '''''''''''''''''''
'*********************************************************************************
'     New CPS Comms module for standardization
'
'
'*********************************************************************************
'
'     The below code was provided by OPTO to communicate with their BRICK B3000
'
'*********************************************************************************
'*********************************************************************************
Option Explicit
' global variable used to indicate view file. (For Help text)
Global ViewFile As String
' Declarations for global variable passed to the
' SendMIO function.
Global Brick_Handle As Long
Global Brick_Error As Integer
Global Brick_Address As Long
Global Brick_Cmd As Long
Global Brick_Pos(0 To 1) As Long
Global Brick_Send(0 To 15) As Long
Global Brick_Rece(0 To 15) As Long
' Global parameters for the port configuration
Global Brick_ioPort As Long
Global Brick_Port As Long
Global Brick_Baud As Long
Global Brick_TimeOut As Single   ' in seconds
Global Brick_Retry As Long
Global Brick_ProtoType As Long
Global Brick_CheckType As Long
' general purpose global variables
Global MisticErr As String
Global Connected As Integer
Global PortOpen As Integer
' defined constants used in this program
Global Const COM1 = 1
Global Const COM2 = 2
Global Const COM3 = 3
Global Const COM4 = 4
Global Const COM5 = 5
Global Const COM6 = 6
Global Const COM7 = 7
Global Const COM8 = 8
Global Const TypeOutput = 1
Global Const TypeInput = 0
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4

Sub ComboValueAdd(Box As ComboBox, text$, Value&)
    Box.AddItem text$
    Box.ItemData(Box.NewIndex) = Value&
End Sub

Function ComboValueGet&(Box As ComboBox)
  '' Return the value of the selected item.
  If Box.ListIndex < 0 Then
    ComboValueGet& = 0
  Else
    ComboValueGet& = Box.ItemData(Box.ListIndex)
  End If
End Function

Sub ComboSelectFromValue(Box As ComboBox, Value&)
  '' Using the value, select the appropriate entry.
  '' reverse operation of ComboValueGet&(PortType)
  Dim Index_%

  For Index_% = 0 To Box.ListCount - 1
    If Value& = Box.ItemData(Index_%) Then
      Box.ListIndex = Index_%
      Exit Sub
    End If
  Next Index_%
End Sub

Function MisticError$(ErrorCode%)
      Dim Buffer$   ' will hold string passed back from optoerr
      Dim Actual0&  ' length of string passed back
      Dim bOk%
      Dim StringLength%
      
      Buffer$ = String(O22_ERROR_MAX_STRING_LENGTH0, 32)
      
      ' call optoerr dll function
      bOk% = O22ErrorAsString( _
        ErrorCode%, _
        Buffer$, _
        O22_ERROR_MAX_STRING_LENGTH0, _
        Actual0&)
      
      StringLength = Actual0& - 1
      If StringLength < 0 Then StringLength = 0
      Buffer$ = Left$(Buffer$, StringLength)
      MisticError$ = Buffer$
      If ErrorCode% <> 0 Then
        frmMainForm.txtCount = frmMainForm.txtCount + 1
      End If
End Function

Function Map_All()
'*********************************************************************************
'     Convert Opto RawValues to Functional Values
'*********************************************************************************
Dim prg As Integer
Dim shft As Integer
Dim stn As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler

' Common Functions
SetErrModule 15, 2300
Map_ComDigitals
Map_ComAnalogs

DoEvents

' Station Functions
ChgErrModule 15, 2301
For stn = 1 To NR_STN
    
    Map_StnDigitals stn
    Map_StnAnalogs stn

Next stn

DoEvents

' Map Scale Values to StationScales
ChgErrModule 15, 2303
For stn = 1 To NR_STN
    For shft = 1 To NR_SHIFT
    
        If Not StationRecipe(stn, shft).UseAuxScale And Not StationRecipe(stn, shft).UsePriScale Then
            StationControl(stn, shft).Scale_OK = False
        Else
            If StationRecipe(stn, shft).UseAuxScale Then
                StationControl(stn, shft).Scale_OK = Scale_OK(StationRecipe(stn, shft).AuxScaleNo)
                StationControl(stn, shft).AuxScaleWt = Scale_Weight(StationRecipe(stn, shft).AuxScaleNo)
            Else
                StationControl(stn, shft).AuxScaleWt = CSng(0)
            End If
            If StationRecipe(stn, shft).UsePriScale Then
                StationControl(stn, shft).Scale_OK = Scale_OK(StationRecipe(stn, shft).PriScaleNo)
                StationControl(stn, shft).PriScaleWt = Scale_Weight(StationRecipe(stn, shft).PriScaleNo)
            Else
                StationControl(stn, shft).PriScaleWt = CSng(0)
            End If
        End If

    Next shft
Next stn

DoEvents

' PurgeAir Functions
SetErrModule 15, 2304
For prg = 1 To NR_PRGAIR
    Map_PrgDigitals prg
    Map_PrgAnalogs prg
Next prg


ResetErrModule

Exit Function
 
localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Function Map_ComAnalogs()
'*********************************************************************************
'     Convert Common Analog Opto Counts to Engr Unit Value
'*********************************************************************************
Dim func As Integer
Dim newraw As Long
Dim craw, clim, neweu As Single
Dim cmax, cmin, cspan As Single
Dim emax, emin, espan As Single
Dim ctemp, etemp As Single

SetErrModule 15, 2330
If UseLocalErrorHandler Then On Error GoTo localhandler

    For func = 1 To MAX_ANA_COM
    
        neweu = 0#
    
        If Com_AIO(func).addr <> 0 Or Com_AIO(func).chan <> 0 Then
        
            newraw = OptoAIO(Com_AIO(func).addr, Com_AIO(func).chan).RawValue
            craw = CSng(newraw)
            
            Select Case OptoAIO(Com_AIO(func).addr, Com_AIO(func).chan).Type
            
                Case optotypeAI, optotypeAO
                    '
                    ' Linear AI or AO, MinMax of 0-10 Vdc Input
                    '
                    clim = CSng(FULLSCALE)
                    cmax = clim * (Com_AIO(func).VdcMax / 10#)
                    cmin = clim * (Com_AIO(func).VdcMin / 10#)
                    cspan = cmax - cmin
                    emax = Com_AIO(func).EuMax
                    emin = Com_AIO(func).EuMin
                    espan = emax - emin
                
                    If cspan <> 0# Then
                        ctemp = (craw - cmin) / cspan
                        neweu = Cal_AnalogInput(ctemp, calgrpComm, func, Com_AiCal(func))
                    End If
                
                    
                Case optotypeTcJ, optotypeTcK
                    '
                    ' Thermocouple
                    '
                    clim = CSng(FULLSCALE / 10)
                    emin = Com_AIO(func).EuMin
                    If USINGC Then    ' using deg C
                        neweu = (craw / clim) + emin
                    End If
                    If USINGF Then    ' using deg F
                        neweu = (((craw / clim) * 9#) / 5#) + 32# + emin
                    End If
                
                
                Case optotypeRTD
                    '
                    ' RTD 100 Ohms
                    '
                    clim = CSng(FULLSCALE / 10)
                    emin = Com_AIO(func).EuMin
                    If USINGC Then    ' using deg C
                        neweu = (craw / clim) + emin
                    End If
                    If USINGF Then    ' using deg F
                        neweu = (((craw / clim) * 9#) / 5#) + 32# + emin
                    End If
                
                Case Else           ' eu value = raw value
                    neweu = craw
                       
            End Select
                
            Map_AIO(Com_AIO(func).addr, Com_AIO(func).chan).RawValue = newraw
            Map_AIO(Com_AIO(func).addr, Com_AIO(func).chan).EUValue = neweu
        
        End If
    
        Com_AIO(func).EUValue = neweu
        
    Next func
         
       
    '*****************************************************************************
    ChgErrModule 15, 2331
    
    AmbBaro = Com_AIO(acAmbBaroSensor).EUValue
    AmbTemp = Com_AIO(acAmbTempSensor).EUValue
    AmbHum = Com_AIO(acAmbHumiditySensor).EUValue
    PTinvalue = Com_AIO(acComnPressSensor).EUValue
    If ((USINGPASLOCALCONTROL And Com_DIO(icPASPowerOnIn).Value) Or (USINGDRYPURGEAIR And SysConfig.DryAirPurge)) Then
        PATemp = Com_AIO(acPasTempSensor).EUValue
        PAHum = Com_AIO(acPasHumiditySensor).EUValue
        If USINGMoist_RH Then
            AmbMoisture = PAHum
            PAMoisture = PAHum
        ElseIf USINGMoist_Grains Then
            If ((AmbTemp > 0) And (AmbBaro > 0) And (AmbHum > 0)) Then
                If ((AmbTemp < 200) And (AmbBaro < 2000) And (AmbHum < 200)) Then
                    AmbMoisture = RHtoGrains(AmbBaro, AmbTemp, AmbHum)
                End If
            End If
            If ((PATemp > 0) And (AmbBaro > 0) And (PAHum > 0)) Then
                If ((PATemp < 200) And (AmbBaro < 2000) And (PAHum < 200)) Then
                    PAMoisture = RHtoGrains(AmbBaro, PATemp, PAHum)
                End If
            End If
        End If
    ElseIf (LocalPagControl.Type = pagClient) Then
        PATemp = MasterPagData.Temperature
        PAHum = MasterPagData.Humidity
        PAMoisture = MasterPagData.Moisture
        If USINGMoist_RH Then
            AmbMoisture = PAHum
        ElseIf USINGMoist_Grains Then
            If ((AmbTemp > 0) And (AmbBaro > 0) And (AmbHum > 0)) Then
                If ((AmbTemp < 200) And (AmbBaro < 2000) And (AmbHum < 200)) Then
                    AmbMoisture = RHtoGrains(AmbBaro, AmbTemp, AmbHum)
                End If
            End If
        End If
    Else
        PATemp = Com_AIO(acAmbTempSensor).EUValue
        PAHum = Com_AIO(acAmbHumiditySensor).EUValue
        If USINGMoist_RH Then
            AmbMoisture = PAHum
            PAMoisture = PAHum
        ElseIf USINGMoist_Grains Then
            If ((AmbTemp > 0) And (AmbBaro > 0) And (AmbHum > 0)) Then
                If ((AmbTemp < 200) And (AmbBaro < 2000) And (AmbHum < 200)) Then
                    AmbMoisture = RHtoGrains(AmbBaro, AmbTemp, AmbHum)
                End If
            End If
            If ((PATemp > 0) And (AmbBaro > 0) And (PAHum > 0)) Then
                If ((PATemp < 200) And (AmbBaro < 2000) And (PAHum < 200)) Then
                    PAMoisture = RHtoGrains(AmbBaro, PATemp, PAHum)
                End If
            End If
        End If
    End If
    '*****************************************************************************
     
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function Map_ComDigitals()
'*********************************************************************************
'     Convert Common Digital Opto RawValue to Functional Value
'*********************************************************************************
Dim func As Integer

SetErrModule 15, 2310
If UseLocalErrorHandler Then On Error GoTo localhandler

    If (USINGPASLOCALCONTROL And Com_DIO(icPASisRunningIn).Value) Then
        func = icPASPowerOnIn
        If Com_DIO(func).UseInverse Then
            If Com_DIO(func).Value And OptoDIO(Com_DIO(func).addr, Com_DIO(func).chan).RawValue Then
                Write_ELog "PAS Power Off (Not Enabled)"
            End If
            If Not Com_DIO(func).Value And Not OptoDIO(Com_DIO(func).addr, Com_DIO(func).chan).RawValue Then
                Write_ELog "PAS Power On (Enabled)"
            End If
        Else
            If Com_DIO(func).Value And Not OptoDIO(Com_DIO(func).addr, Com_DIO(func).chan).RawValue Then
                Write_ELog "PAS Power Off (Not Enabled)"
            End If
            If Not Com_DIO(func).Value And OptoDIO(Com_DIO(func).addr, Com_DIO(func).chan).RawValue Then
                Write_ELog "PAS Power On (Enabled)"
            End If
        End If
    End If
    
    ChgErrModule 15, 2311
        
    For func = 1 To MAX_DIG_COM
    
        If Com_DIO(func).UseInverse Then
            Com_DIO(func).Value = IIf(OptoDIO(Com_DIO(func).addr, Com_DIO(func).chan).RawValue, False, True)
'            If (func = icMaintSw) Then
'    Debug.Print "use inv - MaintMonInRaw - " & IIf(OptoDIO(Com_DIO(icMaintSw).addr, Com_DIO(icMaintSw).chan).RawValue, "ON", "OFF") & " @ " & Format(Timer, "###,##0.000")
'    Debug.Print "use inv - MaintMonIn - " & IIf(Com_DIO(icMaintSw).Value, "ON", "OFF") & " @ " & Format(Timer, "###,##0.000")
'            End If
        Else
            Com_DIO(func).Value = IIf(OptoDIO(Com_DIO(func).addr, Com_DIO(func).chan).RawValue, True, False)
'            If (func = icMaintSw) Then
'    Debug.Print "MaintMonInRaw - " & IIf(OptoDIO(Com_DIO(icMaintSw).addr, Com_DIO(icMaintSw).chan).RawValue, "ON", "OFF") & " @ " & Format(Timer, "###,##0.000")
'    Debug.Print "MaintMonIn - " & IIf(Com_DIO(icMaintSw).Value, "ON", "OFF") & " @ " & Format(Timer, "###,##0.000")
'            End If
        End If
    
    Next func
    
     
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Function Map_PrgAnalogs(ByVal PURGE As Integer)
'*********************************************************************************
'     Convert PurgeAir Analog Opto RawValue to Engr Unit Value
'*********************************************************************************
Dim func As Integer
Dim newraw As Long
Dim craw, clim, neweu As Single
Dim cmax, cmin, cspan As Single
Dim emax, emin, espan As Single
Dim ctemp, etemp As Single

SetErrModule 15, 4382
If UseLocalErrorHandler Then On Error GoTo localhandler

For func = 1 To MAX_ANA_PRG

    neweu = 0#
    
    If Prg_AIO(PURGE, func).addr <> 0 Or Prg_AIO(PURGE, func).chan <> 0 Then
    
        newraw = OptoAIO(Prg_AIO(PURGE, func).addr, Prg_AIO(PURGE, func).chan).RawValue
        craw = CSng(newraw)
        
        Select Case OptoAIO(Prg_AIO(PURGE, func).addr, Prg_AIO(PURGE, func).chan).Type
                   
            Case optotypeAI, optotypeAO
                '
                ' Linear AI or AO, MinMax of 0-10 Vdc Input
                '
                clim = CSng(FULLSCALE)
                cmax = clim * (Prg_AIO(PURGE, func).VdcMax / 10#)
                cmin = clim * (Prg_AIO(PURGE, func).VdcMin / 10#)
                cspan = cmax - cmin
                emax = Prg_AIO(PURGE, func).EuMax
                emin = Prg_AIO(PURGE, func).EuMin
                espan = emax - emin
                
                If cspan <> 0# Then
                    ctemp = (craw - cmin) / cspan
                    neweu = Cal_AnalogInput(ctemp, (PURGE + 10), func, Prg_AiCal(PURGE, func))
                End If
                
                
             Case optotypeTcJ, optotypeTcK
                '
                ' Thermocouple
                '
                clim = CSng(FULLSCALE / 10)
                emin = Prg_AIO(PURGE, func).EuMin
                If USINGC Then    ' using deg C
                    neweu = (craw / clim) + emin
                End If
                If USINGF Then    ' using deg F
                    neweu = (((craw / clim) * 9#) / 5#) + 32# + emin
                End If
            
             Case optotypeRTD
                '
                ' RTD 100 Ohms
                '
                clim = CSng(FULLSCALE / 10)
                emin = Prg_AIO(PURGE, func).EuMin
                If USINGC Then    ' using deg C
                    neweu = (craw / clim) + emin
                End If
                If USINGF Then    ' using deg F
                    neweu = (((craw / clim) * 9#) / 5#) + 32# + emin
                End If
            
            Case Else           ' eu value = raw value
                neweu = craw
                   
        End Select
            
        Map_AIO(Prg_AIO(PURGE, func).addr, Prg_AIO(PURGE, func).chan).RawValue = newraw
        Map_AIO(Prg_AIO(PURGE, func).addr, Prg_AIO(PURGE, func).chan).EUValue = neweu
    
    End If
        
    Prg_AIO(PURGE, func).EUValue = neweu

Next func
     

ResetErrModule

Exit Function
 
localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Function Map_PrgDigitals(ByVal PURGE As Integer)
'*********************************************************************************
'     Convert PurgeAir Digital Opto RawValue to Functional Value
'*********************************************************************************
Dim func As Integer

SetErrModule 15, 4390
If UseLocalErrorHandler Then On Error GoTo localhandler

For func = 1 To MAX_DIG_PRG

    If Prg_DIO(PURGE, func).UseInverse Then
        Prg_DIO(PURGE, func).Value = IIf(OptoDIO(Prg_DIO(PURGE, func).addr, Prg_DIO(PURGE, func).chan).RawValue, False, True)
    Else
        Prg_DIO(PURGE, func).Value = IIf(OptoDIO(Prg_DIO(PURGE, func).addr, Prg_DIO(PURGE, func).chan).RawValue, True, False)
    End If

Next func

ResetErrModule

Exit Function
 
localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Function Map_StnAnalogs(ByVal station As Integer)
'*********************************************************************************
'     Convert Station Analog Opto RawValue to Engr Unit Value
'*********************************************************************************
Dim func, Shift As Integer
Dim newraw As Long
Dim craw, clim, neweu As Single
Dim cmax, cmin, cspan As Single
Dim emax, emin, espan As Single
Dim ctemp, etemp As Single

SetErrModule 15, 2382
If UseLocalErrorHandler Then On Error GoTo localhandler

    For func = 1 To MAX_ANA_STN
    
        neweu = 0#
    
        If Stn_AIO(station, func).addr <> 0 Or Stn_AIO(station, func).chan <> 0 Then
        
            newraw = OptoAIO(Stn_AIO(station, func).addr, Stn_AIO(station, func).chan).RawValue
            craw = CSng(newraw)
            
            Select Case OptoAIO(Stn_AIO(station, func).addr, Stn_AIO(station, func).chan).Type
            
               
                Case optotypeAI, optotypeAO
                    '
                    ' Linear AI or AO, MinMax of 0-10 Vdc Input
                    '
                    clim = CSng(FULLSCALE)
                    cmax = clim * (Stn_AIO(station, func).VdcMax / 10#)
                    cmin = clim * (Stn_AIO(station, func).VdcMin / 10#)
                    cspan = cmax - cmin
                    emax = Stn_AIO(station, func).EuMax
                    emin = Stn_AIO(station, func).EuMin
                    espan = emax - emin
                    If cspan <> 0# Then
                        ctemp = (craw - cmin) / cspan
                        etemp = emin + (espan * ctemp)
                        Select Case func            ' Note: MFC-SP-inputs use the MFC-PV-input calibration
                            Case asNitrogenFlow, asNitrogenFlowSP
                                ' Nitrogen MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCNITROGEN, Stn_MfcCal(station, MFCNITROGEN))
                            Case asButaneFlow, asButaneFlowSP
                                ' Butane MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCBUTANE, Stn_MfcCal(station, MFCBUTANE))
                            Case asNitrogenORVRFlow, asNitrogenORVRFlowSP
                                ' Nitrogen Hi-Range MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCORVRNIT, Stn_MfcCal(station, MFCORVRNIT))
                            Case asButaneORVRFlow, asButaneORVRFlowSP
                                ' Butane Hi-Range MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCORVRBUT, Stn_MfcCal(station, MFCORVRBUT))
                            Case asPurgeAirFlow, asPurgeAirFlowSP
                                ' Purge MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCPURGEAIR, Stn_MfcCal(station, MFCPURGEAIR))
                            Case asLiveFuelVaporFlow, asLiveFuelVaporFlowSP
                                ' LiveFuelVapor MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCLIVEFUEL, Stn_MfcCal(station, MFCLIVEFUEL))
                            Case asLiveFuelVaporORVRFlow, asLiveFuelVaporORVRFlowSP
                                ' LiveFuelVapor Hi-Range MFC AI & AO
                                neweu = Cal_MfcInput(ctemp, station, MFCORVRLIVE, Stn_MfcCal(station, MFCORVRLIVE))
                            Case Else
                                ' Independent Analog Input
                                neweu = Cal_AnalogInput(ctemp, station, func, Stn_AiCal(station, func))
                        End Select
                    End If
                
                    
                 Case optotypeTcJ, optotypeTcK
                    '
                    ' Thermocouple
                    '
                    clim = CSng(FULLSCALE / 10)
                    emax = Stn_AIO(station, func).EuMax
                    emin = Stn_AIO(station, func).EuMin
                    espan = emax - emin
                    If USINGC Then    ' using deg C
                        neweu = (craw / clim) + emin
                    End If
                    If USINGF Then    ' using deg F
                        neweu = (((craw / clim) * 9#) / 5#) + 32# + emin
                    End If
                
                
                 Case optotypeRTD
                    '
                    ' RTD 100 Ohms
                    '
                    clim = CSng(FULLSCALE / 10)
                    emax = Stn_AIO(station, func).EuMax
                    emin = Stn_AIO(station, func).EuMin
                    espan = emax - emin
                    If USINGC Then    ' using deg C
                        neweu = (craw / clim) + emin
                    End If
                    If USINGF Then    ' using deg F
                        neweu = (((craw / clim) * 9#) / 5#) + 32# + emin
                    End If
                
                Case Else
                    ' eu value = raw value
                    neweu = craw
                       
            End Select
                
            Map_AIO(Stn_AIO(station, func).addr, Stn_AIO(station, func).chan).RawValue = newraw
            Map_AIO(Stn_AIO(station, func).addr, Stn_AIO(station, func).chan).EUValue = neweu
        
        End If
    
        Stn_AIO(station, func).EUValue = neweu
                
    Next func
         
    
    '*****************************************************************************
    ' Update Instantaneous Load Flow Values, etc.
    Shift = IIf((Stn_ActiveShift(station) > 0), Stn_ActiveShift(station), 1)
    Select Case STN_INFO(station).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asNitrogenFlow).EUValue
            Stn_Btn_Flow_PV(station, Shift) = Stn_AIO(station, asButaneFlow).EUValue
                
        Case STN_ORVR2_TYPE
            If StationRecipe(station, Shift).UseHiRangeMFC Then
                ' Use Higher Range MFC's
                Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asNitrogenORVRFlow).EUValue
                Stn_Btn_Flow_PV(station, Shift) = Stn_AIO(station, asButaneORVRFlow).EUValue
            Else
                ' Use Lower Range MFC's
                Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asNitrogenFlow).EUValue
                Stn_Btn_Flow_PV(station, Shift) = Stn_AIO(station, asButaneFlow).EUValue
            End If
                
        Case STN_LIVEFUEL_TYPE
            Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asLiveFuelVaporFlow).EUValue
            Stn_Btn_Flow_PV(station, Shift) = 0
        
        Case STN_LIVEREG_TYPE
            If (StationRecipe(station, Shift).LiveFuel) Then
                ' use Live Fuel
                Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asLiveFuelVaporFlow).EUValue
                Stn_Btn_Flow_PV(station, Shift) = 0
            Else
                ' use Butane/Nitrogen
                Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asNitrogenFlow).EUValue
                Stn_Btn_Flow_PV(station, Shift) = Stn_AIO(station, asButaneFlow).EUValue
            End If
        
        Case STN_LIVEORVR2_TYPE
            If (StationRecipe(station, Shift).LiveFuel) Then
                ' use Live Fuel
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' Use Higher Range MFC's
                    Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asLiveFuelVaporORVRFlow).EUValue
                    Stn_Btn_Flow_PV(station, Shift) = 0
                Else
                    ' Use Lower Range MFC's
                    Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asLiveFuelVaporFlow).EUValue
                    Stn_Btn_Flow_PV(station, Shift) = 0
                End If
            Else
                ' use Butane/Nitrogen
                If StationRecipe(station, Shift).UseHiRangeMFC Then
                    ' Use Higher Range MFC's
                    Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asNitrogenORVRFlow).EUValue
                    Stn_Btn_Flow_PV(station, Shift) = Stn_AIO(station, asButaneORVRFlow).EUValue
                Else
                    ' Use Lower Range MFC's
                    Stn_Nit_Flow_PV(station, Shift) = Stn_AIO(station, asNitrogenFlow).EUValue
                    Stn_Btn_Flow_PV(station, Shift) = Stn_AIO(station, asButaneFlow).EUValue
                End If
            End If
        
        Case STN_COMBO3_TYPE
            ' future
        
        Case STN_LEAKTEST_TYPE
            ' LeakTest Station
            UpdateLeakInputs station
        
        Case Else
            ' Do Nothing
    End Select
        
 
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function Map_StnDigitals(ByVal station As Integer)
'*********************************************************************************
'     Convert Station Digital Opto RawValue to Functional Value
'*********************************************************************************
Dim func As Integer

SetErrModule 15, 2390
If UseLocalErrorHandler Then On Error GoTo localhandler

    For func = 1 To MAX_DIG_STN
    
        If ((func = isFuelSafetyLevelLS) And (station = 2)) Then
            func = func
        End If
    
    
        If Stn_DIO(station, func).UseInverse Then
            ' Input ON = FALSE
            Stn_DIO(station, func).Value = IIf(OptoDIO(Stn_DIO(station, func).addr, Stn_DIO(station, func).chan).RawValue, False, True)
    '        If ((func = isFuelHiHiLevelLS) And (station = 1)) Then
    'Debug.Print "mapInv HiHi raw  = " & OptoDIO(Stn_DIO(station, isFuelHiHiLevelLS).addr, Stn_DIO(station, isFuelHiHiLevelLS).chan).RawValue & " @ " & Format(Timer, "###,##0.000")
    'Debug.Print "mapInv HiHi LSw  = " & Stn_DIO(station, isFuelHiHiLevelLS).Value & " @ " & Format(Timer, "###,##0.000")
    '        End If
         Else
            ' Input ON = TRUE
            Stn_DIO(station, func).Value = IIf(OptoDIO(Stn_DIO(station, func).addr, Stn_DIO(station, func).chan).RawValue, True, False)
    '        If ((func = isFuelHiHiLevelLS) And (station = 1)) Then
    'Debug.Print "mapNor HiHi raw  = " & OptoDIO(Stn_DIO(station, isFuelHiHiLevelLS).addr, Stn_DIO(station, isFuelHiHiLevelLS).chan).RawValue & " @ " & Format(Timer, "###,##0.000")
    'Debug.Print "mapNor HiHi LSw  = " & Stn_DIO(station, isFuelHiHiLevelLS).Value & " @ " & Format(Timer, "###,##0.000")
    '        End If
        End If
    
    Next func
    
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function Com_OutDigital(ByVal func As Integer, ByVal Control As Integer)
'*********************************************************************************
'   TURN ON/OFF A COMMON FUNCTION (DIGITAL) OUTPUT
'*********************************************************************************

Dim address As Integer
Dim channel As Integer
Dim badaddr As Boolean

SetErrModule 15, 2262
If UseLocalErrorHandler Then On Error GoTo localhandler

    address = CInt(Com_DIO(func).addr)
    channel = CInt(Com_DIO(func).chan)
        
    If address + channel <> 0 Then
        OPTO_WriteDigital address, channel, Control
    Else
        badaddr = True
    End If
    
    ' debug logging to Zlog
    If Not NotDebugPURGE Then
        If func = icPurgeRequestOut Or func = icPurgeRequestOut Then
            Dim txt As String
            txt = IIf(badaddr, "Write Common DO - BadAddress Error", "Write Common DO")
            Write_Zlog_Purge 0, func, address, channel, Control, txt
        End If
    End If
    
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function Com_OutAnalog(ByVal func As Integer, ByVal euval As Single, ByVal Control As Integer)
'*********************************************************************************
'   SET AN ANALOG OUTPUT
'*********************************************************************************
'*********************************************************************************
'       euval is the value to output in Engr Units
'       control must be = 1 to set an Analog Output > 0; (control = 0 sets rawval=0)
'       rawval is the value to send to OPTO
'***********************************************************************************************
'***********************************************************************************************
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 15, 2277

Dim address As Integer
Dim channel As Integer
Dim RawVal As Long
Dim cmax, cmin, cspan As Long
Dim emax, emin, espan As Single

    address = CInt(Com_AIO(func).addr)
    channel = CInt(Com_AIO(func).chan)
    
    If ((address <> 0) Or (channel <> 0)) Then
        
        If Control = outZERO Then
            
            RawVal = 0
            OPTO_WriteAnalog address, channel, RawVal
        
        Else
            
            ' Linear AO, MinMax of 0-10 Vdc Output
            cmax = FULLSCALE * (Com_AIO(func).VdcMax / 10#)
            cmin = FULLSCALE * (Com_AIO(func).VdcMin / 10#)
            cspan = cmax - cmin
            emax = Com_AIO(func).EuMax
            emin = Com_AIO(func).EuMin
            espan = emax - emin
            
            If (espan <> 0) Then
                
                RawVal = cmin + (cspan * ((euval - emin) / espan))
                OPTO_WriteAnalog address, channel, RawVal
                
            Else
                
                Write_ELog "Com AO #" & Com_AnaDef(func).desc & " (Espan=0)Error - (Func=" & Format(func, "####0") & ")"
    
            End If
                
        End If
                
    Else
        
        Write_ELog "Com AO #" & Com_AnaDef(func).desc & " (Addr+Chan=0)Error - (Func=" & Format(func, "####0") & ")"
    
    End If

ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function Stn_OutAnalog(ByVal station As Integer, ByVal func As Integer, ByVal euval As Single, ByVal Control As Integer)
'*********************************************************************************
'   SET AN ANALOG OUTPUT
'*********************************************************************************
'*********************************************************************************
'       euval is the value to output in Engr Units
'       rawval is the value to send to OPTO
'***********************************************************************************************
'***********************************************************************************************
'
Dim address As Integer
Dim channel As Integer
Dim RawVal As Long
Dim cmax, cmin, cspan As Long
Dim emax, emin, espan As Single

SetErrModule 15, 2273
If UseLocalErrorHandler Then On Error GoTo localhandler

    address = CInt(Stn_AIO(station, func).addr)
    channel = CInt(Stn_AIO(station, func).chan)
    If ((address = 0) And (channel = 0)) Then
        ResetErrModule
        Exit Function
    End If
    
    If Control = outZERO Then
        
        RawVal = 0
    
    Else
        
        ' Linear AO, MinMax of 0-10 Vdc Output
        cmax = FULLSCALE * (Stn_AIO(station, func).VdcMax / 10#)
        cmin = FULLSCALE * (Stn_AIO(station, func).VdcMin / 10#)
        cspan = cmax - cmin
        emax = Stn_AIO(station, func).EuMax
        emin = Stn_AIO(station, func).EuMin
        espan = emax - emin
            
        RawVal = cmin + (cspan * ((euval - emin) / espan))
        
    End If
            
    OPTO_WriteAnalog address, channel, RawVal

ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function Stn_OutDigital(ByVal station As Integer, ByVal func As Integer, ByVal Control As Integer)
'*********************************************************************************
'   TURN ON/OFF A STATION FUNCTION (DIGITAL) OUTPUT
'*********************************************************************************

Dim address As Integer
Dim channel As Integer
Dim outstate As Integer
Dim badaddr As Boolean

SetErrModule 15, 2272
If UseLocalErrorHandler Then On Error GoTo localhandler

    address = CInt(Stn_DIO(station, func).addr)
    channel = CInt(Stn_DIO(station, func).chan)
    
    If (Stn_DIO(station, func).UseInverse) Then
        ' invert output
        outstate = IIf(Control = cON, cOFF, cON)
    Else
        ' normal output
        outstate = IIf(Control = cON, cON, cOFF)
    End If
    
    
    If address + channel <> 0 Then
        OPTO_WriteDigital address, channel, outstate
    Else
        badaddr = True
    End If

ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function OPTO_WriteAnalog(ByVal addr As Integer, ByVal chan As Integer, ByVal control2 As Long)
'*********************************************************************************
'   SET AN ANALOG OUTPUT
'*********************************************************************************
'       Control2 is the value to set in increments of 1/65536
'***********************************************************************************************
'

Dim address As Integer
Dim command As Integer
Dim position0 As Long
Dim position1 As Long

SetErrModule 15, 2443
If UseLocalErrorHandler Then On Error GoTo localhandler

    position1 = 0
    address = CInt(addr)
    position0 = CInt(chan)
    Opto_Send_Data(0) = control2
    command = 443                       ' Output analog command
        
    frmMainForm.Send_Opto_Command address, command, position0, position1
    
    ' Set the Read Value for the AO to the value just written
     OptoAIO(address, position0).RawValue = control2


ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function OPTO_WriteDigital(ByVal addr As Integer, ByVal chan As Integer, ByVal Control As Integer)
'*********************************************************************************
'   TURN ON/OFF A DIGITAL OUTPUT
'*********************************************************************************

Dim address As Integer
Dim command As Integer
Dim position0 As Long
Dim position1 As Long

SetErrModule 15, 2202
If UseLocalErrorHandler Then On Error GoTo localhandler

    position1 = 0
    address = CInt(addr)
    position0 = CInt(chan)
    If CInt(Control) = cON Then
       command = 202        'on/energize
    Else
       command = 201        'off/deenergize
    End If
        
    If address + position0 <> 0 Then
        frmMainForm.Send_Opto_Command address, command, position0, position1
    Else
        ' invalid DO address/channel
        position1 = 0
    End If
    
    ' Set the Read Value for the DO to the value just written
    OptoDIO(address, position0).RawValue = IIf(CInt(Control) = cON, True, False)


ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function OPTO_ReadAnalog(ByVal addr As Integer)
'*********************************************************************************
'*********************************************************************************
'
' Normal de-centralized multiple channel AIO reads from opto
'
'*********************************************************************************
'*********************************************************************************
'
SetErrModule 15, 2230
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim channel As Integer
Dim goodstuff As Integer
Dim return_val As String
Dim ErrTxt$
Dim address As Integer
Dim command As Integer
Dim position0 As Long
Dim position1 As Long

    '*****************************************************************************
    Opto_Rec_Data(16) = " "
    
    address = CInt(addr)
    position0 = 0                       ' channel mask; 255 for 0 - 7 ...65279 is for 8 - 15 ...65535 is for 0 - 15
    For channel = 0 To 15
        If OptoAIO(address, channel).Type <> 0 Then
            position0 = position0 + (2 ^ channel)
        End If
    Next channel
    
    position1 = 0                      ' not used for this command
    command = 307                      ' command = multi-channel AI read
    
    return_val = frmMainForm.Send_Opto_Command(address, command, position0, position1) ' read analog values 8 - 15 = FF00
    
    goodstuff = 0
    For channel = 0 To 15
        If Val(Opto_Rec_Data(channel)) > 0 Then
            goodstuff = 1
        End If
    Next channel
    
    If goodstuff = 1 Then
        If Opto_Rec_Data(16) = " " Then
            
            For channel = 0 To 15
            
                ' Valid Data
                OptoAIO(address, channel).RawValue = Val(Opto_Rec_Data(channel))
    
            Next channel
                                      
        End If
    Else
        ' Set raw values to zero
        For channel = 0 To 15
            
            OptoAIO(address, channel).RawValue = 0
    
        Next channel
    End If
    
    ErrTxt$ = MisticError(Brick_Error)
    If Opto_Rec_Data(16) <> " " Then                                'This opto board has an error
        Delay_Box ErrTxt$, MSGDELAY, msgSHOW                        '   + " opto error"
        Write_ELog ErrTxt$ & "Addr=" & CStr(address)                '   + " opto error"
        Opto_COMM_ERROR(address) = Opto_Rec_Data(16)                ' Current error string for this function
    End If
    '*****************************************************************************
 
 
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function OPTO_ReadDigital(ByVal addr As Integer)
'*********************************************************************************
'*********************************************************************************
'
' Normal de-centralized multiple channel DIO reads from opto
'
'*********************************************************************************
'*********************************************************************************
'

SetErrModule 15, 2210
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim channel As Integer
Dim return_val As String
Dim ErrTxt$
Dim address As Integer
Dim command As Integer
Dim position0 As Long
Dim position1 As Long

SetErrModule 15, 2211

    Opto_Rec_Data(16) = " "
    address = CInt(addr)
    position0 = 0                      ' pos 0 is 0
    position1 = 0                      ' pos 1 is 0
    command = 102                      ' Read block digital
    
    return_val = frmMainForm.Send_Opto_Command(address, command, position0, position1)  ' read digital values 0 - 15
    ' return_val = frmMainForm.Send_Opto_Command(0, 102, 0, 0)  ' read digital values 0 - 15 ' Hard coded if needed
    
    If Opto_Rec_Data(16) = " " Then                                              ' valid data continue
    
        For channel = 0 To 15
        
            OptoDIO(address, channel).RawValue = IIf(Opto_Rec_Data(channel) > 0, True, False)
    
        Next channel
    
    End If
    
    
    ErrTxt$ = MisticError(Brick_Error)
    If Opto_Rec_Data(16) <> " " Then                                'This opto board has an error
        Delay_Box ErrTxt$, MSGDELAY, msgSHOW                        ' + " opto error"
        Write_ELog ErrTxt$                                          ' + " opto error"
        Opto_COMM_ERROR(address) = Opto_Rec_Data(16)                ' Current error string for this function
    End If


ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function


'*******************************************************************************************
Function Close_Main_Valves()
'
' Close all Main valves
' This  is important to start all I/O's at a given position
' And to end all functions in a given position
'
Dim address As Integer
Dim channel As Integer
Dim func As Integer

SetErrModule 15, 330

    ' Turn Off All Functional DO's that are InUse
    For func = 1 To MAX_DIG_COM
        If func <> icButaneShutoffSol Or Not systemhasBUTANE Or Pause_Alarm = SYSTEMPAUSED Then         ' don't shutoff Butane except for alarm
            If Com_DIO(func).addr <> 0 And Com_DIO(func).chan <> 0 Then
                    address = CInt(Com_DIO(func).addr)
                    channel = CInt(Com_DIO(func).chan)
                    If OptoDIO(address, channel).Type = optotypeDO Then
                        OPTO_WriteDigital address, channel, cOFF
                    End If
            End If
        End If
    Next func
    
    ' Release the Common (Leak) Pressure Transducer
    LeakCheckControl.station = 0
    LeakCheckControl.Shift = 0
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

'*******************************************************************************************
Function Close_Scl_Valves(ByVal scalenum As Integer)
'
' Close all valves for a given scale
' It is important to start and end all tests
' With all I/O in a known position
'
Dim address As Integer
Dim channel As Integer

SetErrModule 15, 4033

    If scalenum <= NR_STN Then
    
        ' Turn Off Functional DO's that are InUse
        '
        '   (scale valves are physically installed with the corresponding station's valves)
                
        ' AUX PURGE VALVE
        address = CInt(Stn_DIO(scalenum, isAuxPurgeSol).addr)
        channel = CInt(Stn_DIO(scalenum, isAuxPurgeSol).chan)
        If OptoDIO(address, channel).Type = optotypeDO Then
            OPTO_WriteDigital address, channel, cOFF
        End If
        
        ' PRI/AUX VENT VALVE
        address = CInt(Stn_DIO(scalenum, isPriAuxVentSol).addr)
        channel = CInt(Stn_DIO(scalenum, isPriAuxVentSol).chan)
        If OptoDIO(address, channel).Type = optotypeDO Then
            OPTO_WriteDigital address, channel, cOFF
        End If
        
    End If

ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

'*******************************************************************************************
Function Close_Stn_Valves(ByVal stn As Integer, ByVal Shift As Integer)
'
' Close all valves for a given station. Shift is not really relevant
'   Scale Valves are done elsewhere
' It is important to start and end all tests
' With all I/O in a known position
'
SetErrModule 15, 30

Dim address As Integer
Dim channel As Integer
Dim func As Integer

    ' Turn Off All Functional DO's that are InUse
    For func = 1 To MAX_DIG_STN
        If func <> isAuxPurgeSol And func <> isPriAuxVentSol Then       ' don't do scale funtions here
            address = CInt(Stn_DIO(stn, func).addr)
            channel = CInt(Stn_DIO(stn, func).chan)
            If OptoDIO(address, channel).Type = optotypeDO Then
                OPTO_WriteDigital address, channel, cOFF
            End If
        End If
    Next func
    
    
    ' Turn Off All Functional AO's that are InUse
    Select Case STN_INFO(stn).Type
    
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
            Stn_OutAnalog stn, asButaneFlowSP, 0, outZERO
            Stn_OutAnalog stn, asNitrogenFlowSP, 0, outZERO
          
        Case STN_ORVR2_TYPE
            Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
            If StationRecipe(stn, Shift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog stn, asButaneORVRFlowSP, 0, outZERO
                Stn_OutAnalog stn, asNitrogenORVRFlowSP, 0, outZERO
            Else
                ' use lower range MFC
                Stn_OutAnalog stn, asButaneFlowSP, 0, outZERO
                Stn_OutAnalog stn, asNitrogenFlowSP, 0, outZERO
            End If
        
        Case STN_LIVEFUEL_TYPE
            Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
            Stn_OutAnalog stn, asLiveFuelVaporFlowSP, 0, outZERO
        
        Case STN_LIVEREG_TYPE
            If (StationRecipe(stn, Shift).LiveFuel) Then
                Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
                Stn_OutAnalog stn, asLiveFuelVaporFlowSP, 0, outZERO
            Else
                Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
                Stn_OutAnalog stn, asButaneFlowSP, 0, outZERO
                Stn_OutAnalog stn, asNitrogenFlowSP, 0, outZERO
            End If
        
        Case STN_LIVEORVR2_TYPE
            If (StationRecipe(stn, Shift).LiveFuel) Then
                Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
                If StationRecipe(stn, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog stn, asLiveFuelVaporORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog stn, asLiveFuelVaporFlowSP, 0, outZERO
                End If
            Else
                Stn_OutAnalog stn, asPurgeAirFlowSP, 0, outZERO
                If StationRecipe(stn, Shift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog stn, asButaneORVRFlowSP, 0, outZERO
                    Stn_OutAnalog stn, asNitrogenORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog stn, asButaneFlowSP, 0, outZERO
                    Stn_OutAnalog stn, asNitrogenFlowSP, 0, outZERO
                End If
            End If
        
        Case STN_COMBO3_TYPE
            ' future
        
        Case STN_LEAKTEST_TYPE
            ' future
        
        Case STN_DUMMY_TYPE
            ' Nothing to do
            
    End Select
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function Reset_Valves()     ' This is for all valves
'
' This is a power up / close type clear to let us start/end clean
'
Dim stn As Integer
Dim scl As Integer

SetErrModule 15, 31

    ' reset each station's valves
    For stn = 1 To NR_STN
         Close_Stn_Valves stn, 1                        ' One per opto board stn level
    Next stn
    
    ' reset each scale's valves
    For scl = 1 To NR_SCALES
         If scl <= NR_STN Then Close_Scl_Valves scl     ' share opto board with stn valves
    Next scl
    
    ' now reset the common valves
    Close_Main_Valves                                   ' one opto board

ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Sub ShutdownStnMFCs(ByVal iStn As Integer, ByVal iShift As Integer)
'
'
SetErrModule 15, 3722
If UseLocalErrorHandler Then On Error GoTo localhandler

    '  Turn Off MFC Outputs
   
    ' purge mfc
    Stn_OutAnalog iStn, asPurgeAirFlowSP, 0, outZERO
    ' load mfc variations
    Select Case STN_INFO(iStn).Type
        Case STN_REGULAR_TYPE, STN_ORVR_TYPE
            Stn_OutAnalog iStn, asNitrogenFlowSP, 0, outZERO
            Stn_OutAnalog iStn, asButaneFlowSP, 0, outZERO
        Case STN_ORVR2_TYPE
            If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                ' use higher range MFC
                Stn_OutAnalog iStn, asNitrogenORVRFlowSP, 0, outZERO
                Stn_OutAnalog iStn, asButaneORVRFlowSP, 0, outZERO
            Else
                ' use lower range MFC
                Stn_OutAnalog iStn, asNitrogenFlowSP, 0, outZERO
                Stn_OutAnalog iStn, asButaneFlowSP, 0, outZERO
            End If
        Case STN_LIVEFUEL_TYPE
            Stn_OutAnalog iStn, asLiveFuelVaporFlowSP, 0, outZERO
        Case STN_LIVEREG_TYPE
            If StationRecipe(iStn, iShift).LiveFuel Then
                ' use Live Fuel
                Stn_OutAnalog iStn, asLiveFuelVaporFlowSP, 0, outZERO
                Stn_OutAnalog iStn, asNitrogenFlowSP, 0, outZERO
                Stn_OutAnalog iStn, asButaneFlowSP, 0, outZERO
            Else
                ' use Butane/Nitrogen
                Stn_OutAnalog iStn, asLiveFuelVaporFlowSP, 0, outZERO
                Stn_OutAnalog iStn, asNitrogenFlowSP, 0, outZERO
                Stn_OutAnalog iStn, asButaneFlowSP, 0, outZERO
            End If
        Case STN_LIVEORVR2_TYPE
            If StationRecipe(iStn, iShift).LiveFuel Then
                ' use Live Fuel
                If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog iStn, asLiveFuelVaporORVRFlowSP, 0, outZERO
'                            Stn_OutAnalog iStn, asNitrogenORVRFlowSP, 0, outZERO
'                            Stn_OutAnalog iStn, asButaneORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog iStn, asLiveFuelVaporFlowSP, 0, outZERO
'                            Stn_OutAnalog iStn, asNitrogenFlowSP, 0, outZERO
'                            Stn_OutAnalog iStn, asButaneFlowSP, 0, outZERO
                End If
            Else
                ' use Butane/Nitrogen
                If StationRecipe(iStn, iShift).UseHiRangeMFC Then
                    ' use higher range MFC
                    Stn_OutAnalog iStn, asNitrogenORVRFlowSP, 0, outZERO
                    Stn_OutAnalog iStn, asButaneORVRFlowSP, 0, outZERO
                Else
                    ' use lower range MFC
                    Stn_OutAnalog iStn, asNitrogenFlowSP, 0, outZERO
                    Stn_OutAnalog iStn, asButaneFlowSP, 0, outZERO
                End If
            End If
        Case STN_COMBO3_TYPE
            ' future
        Case STN_LEAKTEST_TYPE
            ' future
        Case Else
            ' nothing to do
    End Select

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

Function Prg_OutDigital(ByVal PURGE As Integer, ByVal func As Integer, ByVal Control As Integer)
'*********************************************************************************
'   TURN ON/OFF A PURGEAIR FUNCTION (DIGITAL) OUTPUT
'*********************************************************************************

Dim address As Integer
Dim channel As Integer
Dim badaddr As Boolean

SetErrModule 15, 3272
If UseLocalErrorHandler Then On Error GoTo localhandler

    address = CInt(Prg_DIO(PURGE, func).addr)
    channel = CInt(Prg_DIO(PURGE, func).chan)
    
    If address + channel <> 0 Then
        OPTO_WriteDigital address, channel, Control
    Else
        badaddr = True
    End If
    
    ' debug logging to Zlog
    If Not NotDebugPURGE Then
        Dim txt As String
        txt = IIf(badaddr, "Write DO - BadAddress Error", "Write DO")
        Write_Zlog_Purge PURGE, func, address, channel, Control, txt
    End If

ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function Reset_Bar_Graph(ByVal stn As Integer, ByVal Shift As Integer)
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 15, 31

    StationControl(stn, Shift).Target = 0
    StationControl(stn, Shift).Actual = 0
    Stn_Nit_FlowSP(stn, Shift) = 0
    Stn_Btn_FlowSP(stn, Shift) = 0
    
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

