Attribute VB_Name = "Module4"
' error module 4 '''''''''''''''program CONFIG.bas '''''''''''''''''''''''''
Option Explicit
'
Sub Save_ButaneSupply()
' Save "Butane Supply" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2330

    SaveSetting "cps_r7", "Butane Supply", "Actual Butane", Format(ButaneSupply.CurrentOnHand, "######0.00")
    SaveSetting "cps_r7", "Butane Supply", "Butane SetPoint", Format(ButaneSupply.CylinderWeight, "######0.00")
    SaveSetting "cps_r7", "Butane Supply", "Warning SetPoint", Format(ButaneSupply.WarningSetPoint, "##0.0##")
    SaveSetting "cps_r7", "Butane Supply", "Butane Change DTS", FormatDateTime(ButaneSupply.Date)
    SaveSetting "cps_r7", "Butane Supply", "LastUpdate DTS", FormatDateTime(Now)

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

Sub Save_StartupSettings()
' Save "StartupSettings" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 23301
Dim def As String

    def = IIf(STARTUPVERBOSE, "VERBOSE", "TERSE")
    SaveSetting "cps_r7", "Startup", "Messages", def

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

Sub Load_StartupSettings()
' Load "StartupSettings" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 24201

Dim def As String
Dim val As String

    ' set to defaults
    STARTUPVERBOSE = False

    ' Load Last Saved Values
        def = IIf(STARTUPVERBOSE, "VERBOSE", "TERSE")
    val = CStr(GetSetting("cps_r7", "Startup", "Messages", def))
    STARTUPVERBOSE = IIf((val = "VERBOSE"), True, False)
    
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

Sub Load_ButaneSupply()
' Load "Butane Supply" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2420

Dim def As String

    ' set to defaults
    ButaneSupply.CurrentOnHand = 1
    ButaneSupply.CylinderWeight = 10
    ButaneSupply.WarningSetPoint = 1
    ButaneSupply.Date = CStr(Now())

    ' Load Last Saved Values
        def = Format(ButaneSupply.CurrentOnHand, "######0.00")
    ButaneSupply.CurrentOnHand = CSng(GetSetting("cps_r7", "Butane Supply", "Actual Butane", def))
        def = Format(ButaneSupply.CylinderWeight, "######0.00")
    ButaneSupply.CylinderWeight = CSng(GetSetting("cps_r7", "Butane Supply", "Butane SetPoint", def))
        def = Format(ButaneSupply.WarningSetPoint, "##0.0##")
    ButaneSupply.WarningSetPoint = CSng(GetSetting("cps_r7", "Butane Supply", "Warning SetPoint", def))
        def = FormatDateTime(ButaneSupply.Date)
    ButaneSupply.Date = GetSetting("cps_r7", "Butane Supply", "Butane Change DTS", def)
    
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

Sub Save_AdfConfig()
' Save "AutoDrainFill" configuration information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2310

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim icnt1, icnt2 As Integer

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
    For icnt1 = 1 To NR_STN
        For icnt2 = 1 To NR_SHIFT
    
            ' Save Station AutoDrainFill Configuration Records
            Criteria = "SELECT * FROM [AdfConfig] WHERE [Station] = " & icnt1 & "  and [Shift] = " & icnt2 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Station") = icnt1
                rsRecord("Shift") = icnt2
            Else
                rsRecord.MoveFirst
                rsRecord.Edit
            End If
               
            rsRecord("DrainDelay") = StationCfg_ADF(icnt1, icnt2).DrainDelay
            rsRecord("DrainTimeout") = StationCfg_ADF(icnt1, icnt2).DrainTimeout
            rsRecord("DrainShutoff") = StationCfg_ADF(icnt1, icnt2).DrainShutOff
            rsRecord("FillDelay") = StationCfg_ADF(icnt1, icnt2).FillDelay
            rsRecord("FillTimeout") = StationCfg_ADF(icnt1, icnt2).FillTimeout
            rsRecord("FillShutoff") = StationCfg_ADF(icnt1, icnt2).FillShutOff
            rsRecord("PurgeDrainDelay") = StationCfg_ADF(icnt1, icnt2).PurgeDrainDelay
            rsRecord("PurgeFillDelay") = StationCfg_ADF(icnt1, icnt2).PurgeFillDelay
            rsRecord("PurgeTimeout") = StationCfg_ADF(icnt1, icnt2).PurgeTimeout
            rsRecord("HeaterTimeout") = StationCfg_ADF(icnt1, icnt2).HeaterTimeout
            rsRecord("VaporTankVolume") = StationCfg_ADF(icnt1, icnt2).VaporGenTankVol
            rsRecord("VaporTankLevelTol") = StationCfg_ADF(icnt1, icnt2).VaporGenLevelTol
            rsRecord("StorageTankVolume") = StationCfg_ADF(icnt1, icnt2).FuelStorageTankVol
            rsRecord("StorageTankLevelTol") = StationCfg_ADF(icnt1, icnt2).FuelStorageLevelTol
            rsRecord("FstDrainDelay") = StationCfg_ADF(icnt1, icnt2).FstDrainDelay
            rsRecord("FstDrainTimeout") = StationCfg_ADF(icnt1, icnt2).FstDrainTimeout
            rsRecord("FstDrainShutoff") = StationCfg_ADF(icnt1, icnt2).FstDrainShutOff
            rsRecord("FstFillDelay") = StationCfg_ADF(icnt1, icnt2).FstFillDelay
            rsRecord("FstFillTimeout") = StationCfg_ADF(icnt1, icnt2).FstFillTimeout
            rsRecord("FstFillShutoff") = StationCfg_ADF(icnt1, icnt2).FstFillShutOff
               
            rsRecord.Update
            rsRecord.Close
            
        Next icnt2
    Next icnt1
    
    
    ' Close Database
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

Sub Load_AdfConfig()
' Load "AutoDrainFill" configuration information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2320

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim icnt1, icnt2 As Integer

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
    For icnt1 = 1 To NR_STN
        For icnt2 = 1 To NR_SHIFT
    
            ' Init Station AutoDrainFill Configuration Records
            StationCfg_ADF(icnt1, icnt2).DrainDelay = 0
            StationCfg_ADF(icnt1, icnt2).DrainTimeout = 0
            StationCfg_ADF(icnt1, icnt2).DrainShutOff = 0
            StationCfg_ADF(icnt1, icnt2).FillDelay = 0
            StationCfg_ADF(icnt1, icnt2).FillTimeout = 0
            StationCfg_ADF(icnt1, icnt2).FillShutOff = 0
            StationCfg_ADF(icnt1, icnt2).PurgeDrainDelay = 0
            StationCfg_ADF(icnt1, icnt2).PurgeFillDelay = 0
            StationCfg_ADF(icnt1, icnt2).PurgeTimeout = 0
            StationCfg_ADF(icnt1, icnt2).HeaterTimeout = 0
            StationCfg_ADF(icnt1, icnt2).FuelStorageTankVol = CSng(0)
            StationCfg_ADF(icnt1, icnt2).FuelStorageLevelTol = CSng(0)
            StationCfg_ADF(icnt1, icnt2).VaporGenTankVol = CSng(0)
            StationCfg_ADF(icnt1, icnt2).VaporGenLevelTol = CSng(0)
            StationCfg_ADF(icnt1, icnt2).FstDrainDelay = 0
            StationCfg_ADF(icnt1, icnt2).FstDrainTimeout = 0
            StationCfg_ADF(icnt1, icnt2).FstDrainShutOff = 0
            StationCfg_ADF(icnt1, icnt2).FstFillDelay = 0
            StationCfg_ADF(icnt1, icnt2).FstFillTimeout = 0
            StationCfg_ADF(icnt1, icnt2).FstFillShutOff = 0
                
            ' Read Station AutoDrainFill Configuration Records
            Criteria = "SELECT * FROM [AdfConfig] WHERE [Station] = " & icnt1 & "  and [Shift] = " & icnt2 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If (Not rsRecord.BOF) Then
                rsRecord.MoveFirst
                If (Not IsNull(rsRecord("DrainDelay"))) Then StationCfg_ADF(icnt1, icnt2).DrainDelay = rsRecord("DrainDelay")
                If (Not IsNull(rsRecord("DrainTimeout"))) Then StationCfg_ADF(icnt1, icnt2).DrainTimeout = rsRecord("DrainTimeout")
                If (Not IsNull(rsRecord("DrainShutoff"))) Then StationCfg_ADF(icnt1, icnt2).DrainShutOff = rsRecord("DrainShutoff")
                If (Not IsNull(rsRecord("FillDelay"))) Then StationCfg_ADF(icnt1, icnt2).FillDelay = rsRecord("FillDelay")
                If (Not IsNull(rsRecord("FillTimeout"))) Then StationCfg_ADF(icnt1, icnt2).FillTimeout = rsRecord("FillTimeout")
                If (Not IsNull(rsRecord("FillShutoff"))) Then StationCfg_ADF(icnt1, icnt2).FillShutOff = rsRecord("FillShutoff")
                If (Not IsNull(rsRecord("PurgeDrainDelay"))) Then StationCfg_ADF(icnt1, icnt2).PurgeDrainDelay = rsRecord("PurgeDrainDelay")
                If (Not IsNull(rsRecord("PurgeFillDelay"))) Then StationCfg_ADF(icnt1, icnt2).PurgeFillDelay = rsRecord("PurgeFillDelay")
                If (Not IsNull(rsRecord("PurgeTimeout"))) Then StationCfg_ADF(icnt1, icnt2).PurgeTimeout = rsRecord("PurgeTimeout")
                If (Not IsNull(rsRecord("HeaterTimeout"))) Then StationCfg_ADF(icnt1, icnt2).HeaterTimeout = rsRecord("HeaterTimeout")
                If (Not IsNull(rsRecord("StorageTankVolume"))) Then StationCfg_ADF(icnt1, icnt2).FuelStorageTankVol = rsRecord("StorageTankVolume")
                If (Not IsNull(rsRecord("StorageTankLevelTol"))) Then StationCfg_ADF(icnt1, icnt2).FuelStorageLevelTol = rsRecord("StorageTankLevelTol")
                If (Not IsNull(rsRecord("VaporTankVolume"))) Then StationCfg_ADF(icnt1, icnt2).VaporGenTankVol = rsRecord("VaporTankVolume")
                If (Not IsNull(rsRecord("VaporTankLevelTol"))) Then StationCfg_ADF(icnt1, icnt2).VaporGenLevelTol = rsRecord("VaporTankLevelTol")
                If (Not IsNull(rsRecord("FstDrainDelay"))) Then StationCfg_ADF(icnt1, icnt2).FstDrainDelay = rsRecord("FstDrainDelay")
                If (Not IsNull(rsRecord("FstDrainTimeout"))) Then StationCfg_ADF(icnt1, icnt2).FstDrainTimeout = rsRecord("FstDrainTimeout")
                If (Not IsNull(rsRecord("FstDrainShutoff"))) Then StationCfg_ADF(icnt1, icnt2).FstDrainShutOff = rsRecord("FstDrainShutoff")
                If (Not IsNull(rsRecord("FstFillDelay"))) Then StationCfg_ADF(icnt1, icnt2).FstFillDelay = rsRecord("FstFillDelay")
                If (Not IsNull(rsRecord("FstFillTimeout"))) Then StationCfg_ADF(icnt1, icnt2).FstFillTimeout = rsRecord("FstFillTimeout")
                If (Not IsNull(rsRecord("FstFillShutoff"))) Then StationCfg_ADF(icnt1, icnt2).FstFillShutOff = rsRecord("FstFillShutoff")
            End If
               
            rsRecord.Close
            
        Next icnt2
    Next icnt1
    
    ' Close Database
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

Sub Save_StationConfig()
' Save "Station Configuration" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2610

Dim inct As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim icnt1, icnt2 As Integer

' Open Database
Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
For icnt1 = 1 To LAST_STN
    For icnt2 = 1 To NR_SHIFT

        ' Save Station Configuration Records
        Criteria = "SELECT * FROM [StationConfig] WHERE [Station] = " & icnt1 & "  and [Shift] = " & icnt2 & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        If rsRecord.BOF Then
            rsRecord.AddNew
            rsRecord("Station") = icnt1
            rsRecord("Shift") = icnt2
        Else
            rsRecord.MoveFirst
            rsRecord.Edit
        End If
           
'        rsRecord("UseCommonTC") = Stn_CommonTC(icnt1, icnt2)
           
        rsRecord.Update
        rsRecord.Close
        
    Next icnt2
Next icnt1

' Close Database
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

Sub Load_StationConfig()
' Load "Station Configuration" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 3620

Dim inct As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim icnt1, icnt2 As Integer

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
    For icnt1 = 1 To LAST_STN
        For icnt2 = 1 To NR_SHIFT
    
            ' Read Station Configuration Records
            Criteria = "SELECT * FROM [StationConfig] WHERE [Station] = " & icnt1 & "  and [Shift] = " & icnt2 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
    '            Stn_CommonTC(icnt1, icnt2) = False
            Else
                rsRecord.MoveFirst
    '            Stn_CommonTC(icnt1, icnt2) = rsRecord("UseCommonTC")
            End If
               
            rsRecord.Close
            
        Next icnt2
    Next icnt1
    
    ' Close Database
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

Sub Save_Simulation()
' Save "Common Simulation" information
' Save "Station Simulation" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2410

Dim iStation As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
    ' Read Common Simulation Information Records
    Criteria = "SELECT * FROM [CommonSimulation] "
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    If rsRecord.BOF Then
        rsRecord.AddNew
    Else
        rsRecord.MoveFirst
        rsRecord.Edit
    End If
       
    rsRecord("Temperature_Error") = Sim_PasError(pasTEMPERATURE)
    rsRecord("Rh_Error") = Sim_PasError(pasMOISTURE)
    rsRecord("LiveFuelDensity") = Sim_LiveFuelDensity
       
    rsRecord.Update
    rsRecord.Close
    
    
    
    For iStation = 1 To NR_STN
    
            ' Read Station Simulation Information Records
            Criteria = "SELECT * FROM [StationSimulation] WHERE [Station] = " & iStation & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Station") = iStation
            Else
                rsRecord.MoveFirst
                rsRecord.Edit
            End If
               
            rsRecord("Nitrogen_MfcError") = Sim_MfcError(iStation, MFCNITROGEN)
            rsRecord("Butane_MfcError") = Sim_MfcError(iStation, MFCBUTANE)
            rsRecord("PurgeAir_MfcError") = Sim_MfcError(iStation, MFCPURGEAIR)
            rsRecord("LiveFuel_MfcError") = Sim_MfcError(iStation, MFCLIVEFUEL)
            rsRecord("LeakError") = Sim_LeakError(iStation)
            rsRecord("AuxCan_StartPercentFull") = Sim_AuxCan_JobStartPercentLoaded(iStation)
            rsRecord("PriCan_StartPercentFull") = Sim_PriCan_JobStartPercentLoaded(iStation)
               
            rsRecord.Update
            rsRecord.Close
            
    Next iStation
    
    ' Close Database
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

Sub Load_Simulation()
' Load "Common Simulation" information
' Load "Station Simulation" information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 3420

Dim iStation As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
    ' Read Common Simulation Information Records
    Criteria = "SELECT * FROM [CommonSimulation] "
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    If rsRecord.BOF Then
        Sim_PasError(pasTEMPERATURE) = 0
        Sim_PasError(pasMOISTURE) = 0
        Sim_LiveFuelDensity = 2.25
    Else
        rsRecord.MoveFirst
        Sim_PasError(pasTEMPERATURE) = rsRecord("Temperature_Error")
        Sim_PasError(pasMOISTURE) = rsRecord("Rh_Error")
        Sim_LiveFuelDensity = rsRecord("LiveFuelDensity")
    End If
       
    rsRecord.Close
    
    For iStation = 1 To NR_STN
    
            ' Read Station Simulation Information Records
            Criteria = "SELECT * FROM [StationSimulation] WHERE [Station] = " & iStation & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                Sim_MfcError(iStation, MFCNITROGEN) = 0
                Sim_MfcError(iStation, MFCBUTANE) = 0
                Sim_MfcError(iStation, MFCPURGEAIR) = 0
                Sim_MfcError(iStation, MFCLIVEFUEL) = 0
                Sim_LeakError(iStation) = 0
                Sim_AuxCan_JobStartPercentLoaded(iStation) = 0
                Sim_PriCan_JobStartPercentLoaded(iStation) = 0
            Else
                rsRecord.MoveFirst
                Sim_MfcError(iStation, MFCNITROGEN) = rsRecord("Nitrogen_MfcError")
                Sim_MfcError(iStation, MFCBUTANE) = rsRecord("Butane_MfcError")
                Sim_MfcError(iStation, MFCPURGEAIR) = rsRecord("PurgeAir_MfcError")
                Sim_MfcError(iStation, MFCLIVEFUEL) = rsRecord("LiveFuel_MfcError")
                Sim_LeakError(iStation) = rsRecord("LeakError")
                Sim_AuxCan_JobStartPercentLoaded(iStation) = rsRecord("AuxCan_StartPercentFull")
                Sim_PriCan_JobStartPercentLoaded(iStation) = rsRecord("PriCan_StartPercentFull")
            End If
               
            rsRecord.Close
    
    Next iStation
    
    ' Close Database
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

Sub Load_StationCanisters()
' Load Station Canister Parameters
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1420

Dim iStation As Integer
Dim iShift As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For iStation = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
    
            ' Read Station Canister Information Records
            Criteria = "SELECT * FROM [StationCanister] WHERE [Station] = " & iStation & " AND [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                StationCanister(iStation, iShift).Number = 0
                StationCanister(iStation, iShift).Description = "undefined"
                StationCanister(iStation, iShift).WorkingCapacity = 0
                StationCanister(iStation, iShift).WorkingVolume = 0
                StationCanister(iStation, iShift).Validated = False
            Else
                StationCanister(iStation, iShift).Number = rsRecord("Number")
                StationCanister(iStation, iShift).Description = rsRecord("Description")
                StationCanister(iStation, iShift).WorkingCapacity = rsRecord("WorkingCapacity")
                StationCanister(iStation, iShift).WorkingVolume = rsRecord("WCVolume")
                StationCanister(iStation, iShift).Validated = True
            End If
               
            rsRecord.Close
    
            ' Check Canister Number
            ' note: Station Canister Tracking was added as part of CfgRevLvl=5
            If CfgRevLvl < 5 Then StationCanister(iStation, iShift).Number = 0
    
        Next iShift
    Next iStation
    
    ' Close Database
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

Sub Load_StationSequences()
' Load Sequence Sequences
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1920

Dim iStation As Integer
Dim iShift As Integer
Dim iCourse As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

' Open Database
Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)

For iStation = 1 To NR_STN
    For iShift = 1 To NR_SHIFT

        ' Clear CourseData
        For iCourse = 1 To MAX_COURSES
            StationSequence(iStation, iShift).CourseData(iCourse).CourseNumber = 0
            StationSequence(iStation, iShift).CourseData(iCourse).Type = courseUndefined
            StationSequence(iStation, iShift).CourseData(iCourse).OkToProceed = False
            StationSequence(iStation, iShift).CourseData(iCourse).PauseDuration = 0
            StationSequence(iStation, iShift).CourseData(iCourse).RecipeNumber = 0
            StationSequence(iStation, iShift).CourseData(iCourse).Cycles = 0
            StationSequence(iStation, iShift).CourseData(iCourse).LoadRate = 0
            StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate = 0
            StationSequence(iStation, iShift).CourseData(iCourse).MsgText = "-na-"
            StationSequence(iStation, iShift).CourseData(iCourse).EstCourseDuration = 0
            StationSequence(iStation, iShift).CourseData(iCourse).DtsStart = Now
            StationSequence(iStation, iShift).CourseData(iCourse).DtsEnd = Now
        Next iCourse
           
        ' Many Job Sequences OR Just One ??
        If (NR_JOBSEQ <= 1) Then
            ' JUST ONE JOB SEQUENCE (THE DEFAULT)
            ' Set Station Sequence to the Default Sequence
            StationSequence(iStation, iShift).Number = CInt(0)
            StationSequence(iStation, iShift).Description = "default station sequence"
            StationSequence(iStation, iShift).PriScaleNo = STN_INFO(iStation).DefPriScale
            StationSequence(iStation, iShift).AuxScaleNo = STN_INFO(iStation).DefAuxScale
            StationSequence(iStation, iShift).EstSeqDuration = EstimatedRcpDuration(StationRecipe(iStation, iShift), StationCanister(iStation, iShift), StationProfile(iStation, iShift))
            StationSequence(iStation, iShift).EstSeqDurDesc = DurationDescription(StationSequence(iStation, iShift).EstSeqDuration)
            StationSequence(iStation, iShift).NumCourses = 1
            StationSequence(iStation, iShift).IDLoad = 0
            StationSequence(iStation, iShift).IDPurge = 0
            StationSequence(iStation, iShift).IDVent = 0
            StationSequence(iStation, iShift).LoadL = 0
            StationSequence(iStation, iShift).LoadV = 0
            StationSequence(iStation, iShift).PurgeL = 0
            StationSequence(iStation, iShift).PurgeV = 0
            StationSequence(iStation, iShift).VentL = 0
            StationSequence(iStation, iShift).VentV = 0
            StationSequence(iStation, iShift).Validated = True
            ' Set first(& only) course data
            StationSequence(iStation, iShift).CourseData(1).CourseNumber = 1
            StationSequence(iStation, iShift).CourseData(1).Type = courseRecipe
            StationSequence(iStation, iShift).CourseData(1).EstCourseDuration = StationSequence(iStation, iShift).EstSeqDuration
        Else
            ' MANY JOB SEQUENCES
            ' Read Station Sequence Information Records
            Criteria = "SELECT * FROM [StationSequence] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            If rsRecord.BOF Then
                StationSequence(iStation, iShift).Number = 0
                StationSequence(iStation, iShift).Description = "undefined"
                StationSequence(iStation, iShift).NumCourses = 0
                StationSequence(iStation, iShift).PriScaleNo = 0
                StationSequence(iStation, iShift).AuxScaleNo = 0
                StationSequence(iStation, iShift).IDLoad = 0
                StationSequence(iStation, iShift).IDPurge = 0
                StationSequence(iStation, iShift).IDVent = 0
                StationSequence(iStation, iShift).LoadL = 0
                StationSequence(iStation, iShift).LoadV = 0
                StationSequence(iStation, iShift).PurgeL = 0
                StationSequence(iStation, iShift).PurgeV = 0
                StationSequence(iStation, iShift).VentL = 0
                StationSequence(iStation, iShift).VentV = 0
                StationSequence(iStation, iShift).Validated = False
                StationSequence(iStation, iShift).EstSeqDuration = 0
                StationSequence(iStation, iShift).EstSeqDurDesc = "undefined"
            Else
    '            StationSequence(iStation, iShift).Number = rsRecord("Number")
                StationSequence(iStation, iShift).Number = 0
                StationSequence(iStation, iShift).Description = rsRecord("Description")
                StationSequence(iStation, iShift).NumCourses = rsRecord("Courses")
                StationSequence(iStation, iShift).PriScaleNo = rsRecord("PriScale")
                StationSequence(iStation, iShift).AuxScaleNo = rsRecord("AuxScale")
                StationSequence(iStation, iShift).IDLoad = rsRecord("IDLoad")
                StationSequence(iStation, iShift).IDPurge = rsRecord("IDPurge")
                StationSequence(iStation, iShift).IDVent = rsRecord("IDVent")
                StationSequence(iStation, iShift).LoadL = rsRecord("LoadL")
                StationSequence(iStation, iShift).LoadV = rsRecord("LoadV")
                StationSequence(iStation, iShift).PurgeL = rsRecord("PurgeL")
                StationSequence(iStation, iShift).PurgeV = rsRecord("PurgeV")
                StationSequence(iStation, iShift).VentL = rsRecord("VentL")
                StationSequence(iStation, iShift).VentV = rsRecord("VentV")
                StationSequence(iStation, iShift).Validated = rsRecord("Validated")
                StationSequence(iStation, iShift).EstSeqDuration = rsRecord("EstSeqDuration")
                StationSequence(iStation, iShift).EstSeqDurDesc = rsRecord("EstSeqDurDesc")
            End If
               
            rsRecord.Close
    
            ' Read Station Sequence Course Information Records
            Criteria = "SELECT * FROM [StationSequenceCourses] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            If Not rsRecord.BOF Then
                rsRecord.MoveFirst
                While Not rsRecord.EOF
                    iCourse = rsRecord("CourseNumber")
                    StationSequence(iStation, iShift).CourseData(iCourse).CourseNumber = iCourse
                    StationSequence(iStation, iShift).CourseData(iCourse).Type = rsRecord("Type")
                    StationSequence(iStation, iShift).CourseData(iCourse).OkToProceed = False
                    StationSequence(iStation, iShift).CourseData(iCourse).PauseDuration = rsRecord("PauseDuration")
                    StationSequence(iStation, iShift).CourseData(iCourse).RecipeNumber = rsRecord("RecipeNumber")
                    StationSequence(iStation, iShift).CourseData(iCourse).Cycles = rsRecord("Cycles")
                    StationSequence(iStation, iShift).CourseData(iCourse).LoadRate = rsRecord("LoadRate")
                    StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate = rsRecord("PurgeRate")
                    If (IsNull(rsRecord("MsgText"))) Then
                        StationSequence(iStation, iShift).CourseData(iCourse).MsgText = "-na-"
                    Else
                        StationSequence(iStation, iShift).CourseData(iCourse).MsgText = rsRecord("MsgText")
                    End If
                    StationSequence(iStation, iShift).CourseData(iCourse).EstCourseDuration = rsRecord("EstCourseDuration")
                    rsRecord.MoveNext
                Wend
            End If
                       
            rsRecord.Close
        End If

    Next iShift
Next iStation

' Close Database
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

Sub Load_StationProfiles()
' Load Station PurgeProfile Parameters
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 7420

Dim iStation As Integer
Dim iShift As Integer
Dim iStep As Integer
Dim dbDbase As Database
Dim rsProfile  As Recordset
Dim rsSteps  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For iStation = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
    
            ' Read Station Profile Information Records
            Criteria = "SELECT * FROM [StationProfiles] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
            Set rsProfile = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsProfile.BOF Then
                StationProfile(iStation, iShift).Number = CInt(0)
                StationProfile(iStation, iShift).Description = "undefined"
                StationProfile(iStation, iShift).Duration = CSng(0)
                StationProfile(iStation, iShift).DurDesc = ProfileDurationDescription(StationProfile(iStation, iShift).Duration)
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
    
            Else
                StationProfile(iStation, iShift).Number = rsProfile("Number")
                StationProfile(iStation, iShift).Description = rsProfile("Description")
                
                StationProfile(iStation, iShift).Duration = rsProfile("TotalDuration")
                StationProfile(iStation, iShift).DurDesc = ProfileDurationDescription(StationProfile(iStation, iShift).Duration)
                StationProfile(iStation, iShift).EndStep = rsProfile("Steps")
                StationProfile(iStation, iShift).ProjectedLiters = rsProfile("ProjectedLiters")
                StationProfile(iStation, iShift).ProjectedVolumes = rsProfile("ProjectedVolumes")
                StationProfile(iStation, iShift).Validated = True
                ' Read Station Profile Steps Information Records
                Criteria = "SELECT * FROM [StationProfileSteps] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
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
    
        Next iShift
    Next iStation
    
    ' Close Database
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

Sub Load_NodeInfo()
' Load OPTO Node Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2794
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim Idx As Integer
Dim iNode As Integer

    sFileName = FILEPATH_cfg & "nodeinfo.@@@"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    '  OPTO Node Information
    For iStation = 0 To MAX_NODE
    
        Input #iFileNumber, Node_Info(iStation)
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
    Next iStation
    
    Close #iFileNumber

    Idx = 0
    For iNode = 0 To MAX_NODE
        If (Node_Info(iNode) > 0) Then
            Idx = iNode
        End If
    Next iNode
    OptoMaxNodeNum = (Idx * 4) + 3
    
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

Sub Load_OptoInfo()
' Load OPTO Module Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2224

Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim iShift As Integer
Dim filler As String

    sFileName = FILEPATH_cfg & "optoinfo.@@@"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    '  OPTO Module Information
    For iStation = 0 To MAX_ADDR
        For iShift = 0 To MAX_SLOT
    
        Input #iFileNumber, Opto_Info(iStation, iShift)
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
        Next iShift
    Next iStation
    
    Close #iFileNumber

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

Sub Load_PurgeInfo()
' Load PurgeAir Source Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 4475

Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim filler As String

    If CfgRevLvl > 0 Then
        
        ' purgeair.@@@ already exists
        sFileName = FILEPATH_cfg & "purgeair.@@@"
        iFileNumber = FreeFile
        Open sFileName For Input As #iFileNumber
        
        '  PurgeAir Definition Information
        For iStation = 1 To MAX_PRG
        
            Input #iFileNumber, PRG_INFO(iStation).desc
            Input #iFileNumber, PRG_INFO(iStation).CheckSecs
            Input #iFileNumber, PRG_INFO(iStation).UsingPrgReqHdw
            Input #iFileNumber, PRG_INFO(iStation).UsingVacSwHdw
            Input #iFileNumber, PRG_INFO(iStation).UsingAuxAirSol
            
            Input #iFileNumber, PRG_INFO(iStation).UsingPosPrsPrg
            Input #iFileNumber, PRG_INFO(iStation).UsingPrgReqAK
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
        Next iStation
        
        Close #iFileNumber
        
    Else
        
        ' first time running this version
        PRG_INFO(1).desc = "Common PIAB"
        PRG_INFO(1).CheckSecs = 3
        PRG_INFO(1).UsingPrgReqAK = False
        PRG_INFO(1).UsingPrgReqHdw = False
        PRG_INFO(1).UsingVacSwHdw = False
        PRG_INFO(1).UsingAuxAirSol = False
        PRG_INFO(1).UsingPosPrsPrg = False
    
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

Sub Load_StationRecipes()
' Load Station Recipes
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1422

Dim iStation, iShift As Integer
Dim iAux As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For iStation = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
    
            ' Read Station Recipe Information Records
            Criteria = "SELECT * FROM [StationRecipe] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                ' Station Recipe is not defined
                StationRecipe(iStation, iShift) = EmptyRecipe
            
            Else
            
                ' Load Station Recipe
                StationRecipe(iStation, iShift).Name = rsRecord("Name")
                
                StationRecipe(iStation, iShift).CycleType = rsRecord("CycleType")
                If (StationRecipe(iStation, iShift).CycleType = CycleUndefined) Then StationRecipe(iStation, iShift).CycleType = CyclePurgeLoad
                
                StationRecipe(iStation, iShift).Load_Method = rsRecord("Load_Method")
                StationRecipe(iStation, iShift).UseHiRangeMFC = rsRecord("UseHiRangeMFC")
                StationRecipe(iStation, iShift).UseLoadRatePID = rsRecord("UseLoadRatePID")
                StationRecipe(iStation, iShift).NitrogenFlow = rsRecord("NitrogenFlow")
                StationRecipe(iStation, iShift).NitrogenFlowSave = StationRecipe(iStation, iShift).NitrogenFlow
                StationRecipe(iStation, iShift).Load_Rate = rsRecord("Load_Rate")
                StationRecipe(iStation, iShift).Load_RateSave = StationRecipe(iStation, iShift).Load_Rate
                StationRecipe(iStation, iShift).Mix_Percent = rsRecord("Mix_Percent")
                StationRecipe(iStation, iShift).WC_Mult = rsRecord("WC_Mult")
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
                If IsNumeric(rsRecord("Purge_TargetMode")) Then
                    StationRecipe(iStation, iShift).Purge_TargetMode = rsRecord("Purge_TargetMode")
                Else
                    StationRecipe(iStation, iShift).Purge_TargetMode = NOTARGET
                End If
                If IsNumeric(rsRecord("Purge_AuxTime")) Then
                    StationRecipe(iStation, iShift).Purge_AuxTime = rsRecord("Purge_AuxTime")
                Else
                    StationRecipe(iStation, iShift).Purge_AuxTime = 0
                End If
                StationRecipe(iStation, iShift).Purge_Time = rsRecord("Purge_Time")
                StationRecipe(iStation, iShift).Purge_Flow = rsRecord("Purge_Flow")
                StationRecipe(iStation, iShift).Purge_Liters = rsRecord("Purge_Liters")
                StationRecipe(iStation, iShift).Purge_Can_Vol = rsRecord("Purge_Can_Vol")
                StationRecipe(iStation, iShift).Purge_ProfileNumber = rsRecord("Purge_ProfileNumber")
                StationRecipe(iStation, iShift).Purge_TargetWC = rsRecord("Purge_TargetWC")
                StationRecipe(iStation, iShift).Purge_TargetWeight = rsRecord("Purge_TargetWeight")
                StationRecipe(iStation, iShift).Purge_MaxVolumes = rsRecord("Purge_MaxVolumes")
                StationRecipe(iStation, iShift).Purge_TargetPurge = rsRecord("Purge_TargetPurge")
                StationRecipe(iStation, iShift).Purge_TargetPause = rsRecord("Purge_TargetPause")
                
                StationRecipe(iStation, iShift).UseAuxScale = rsRecord("UseAuxScale")
                StationRecipe(iStation, iShift).PurgeAuxCan = rsRecord("PurgeAuxCan")
                StationRecipe(iStation, iShift).PurgeOven = rsRecord("PurgeInOven")
                StationRecipe(iStation, iShift).PurgeOvenSP = rsRecord("PurgeOvenSP")
                StationRecipe(iStation, iShift).AuxScaleNo = rsRecord("AuxScaleNo")
                StationRecipe(iStation, iShift).PauseLeakTime = rsRecord("PauseLeakTime")
                StationRecipe(iStation, iShift).PauseLoadTime = rsRecord("PauseLoadTime")
                StationRecipe(iStation, iShift).PausePurgeTime = rsRecord("PausePurgeTime")
                StationRecipe(iStation, iShift).UsePriScale = rsRecord("UsePriScale")
                StationRecipe(iStation, iShift).PriScaleNo = rsRecord("PriScaleNo")
                StationRecipe(iStation, iShift).PauseAfterLeak = rsRecord("PauseAfterLeak")
                StationRecipe(iStation, iShift).PauseAfterLoad = rsRecord("PauseAfterLoad")
                StationRecipe(iStation, iShift).PauseAfterLoadForOper = rsRecord("PauseAfterLoadForOper")
                StationRecipe(iStation, iShift).PauseAfterPurge = rsRecord("PauseAfterPurge")
                StationRecipe(iStation, iShift).PauseAfterPurgeForOper = rsRecord("PauseAfterPurgeForOper")
'                StationRecipe(iStation, iShift).TargetConcentration = rsRecord("TargetConcentration")
'                StationRecipe(iStation, iShift).DwellTime = rsRecord("DwellTime")
                StationRecipe(iStation, iShift).LeakCheck = rsRecord("LeakCheck")
                StationRecipe(iStation, iShift).LeakPrimary = rsRecord("LeakPrimary")
                StationRecipe(iStation, iShift).LeakAux = rsRecord("LeakAux")
'                StationRecipe(iStation, iShift).UseAnalyzer = rsRecord("UseAnalyzer")
                StationRecipe(iStation, iShift).MaxLoadTime = rsRecord("MaxLoadTime")
                
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
                
                StationRecipe(iStation, iShift).StartMethod = rsRecord("StartMethod")
                StationRecipe(iStation, iShift).StartDelay = rsRecord("StartDelay")
                StationRecipe(iStation, iShift).StartDate = rsRecord("StartDate")
                
                ' end method
                StationRecipe(iStation, iShift).EndMethod = rsRecord("EndMethod")
                StationRecipe(iStation, iShift).Cycles = rsRecord("Cycles")
                StationRecipe(iStation, iShift).EndWeightTolerance = rsRecord("EndWeightTolerance")
                StationRecipe(iStation, iShift).EndConsecutiveCycles = rsRecord("EndConsecutiveCycles")
                StationRecipe(iStation, iShift).EndMaximumCycles = rsRecord("EndMaximumCycles")
                StationRecipe(iStation, iShift).EndMinimumCycles = rsRecord("EndMinimumCycles")
                StationRecipe(iStation, iShift).UpdateCanWc = rsRecord("UpdateCanWc")
        
                StationRecipe(iStation, iShift).AuxOutputs = rsRecord("AuxOutputs")
                StationRecipe(iStation, iShift).AuxOutputs_Load(1) = rsRecord("AuxOutput1_Load")
                StationRecipe(iStation, iShift).AuxOutputs_Load(2) = rsRecord("AuxOutput2_Load")
                StationRecipe(iStation, iShift).AuxOutputs_Load(3) = rsRecord("AuxOutput3_Load")
                StationRecipe(iStation, iShift).AuxOutputs_Load(4) = rsRecord("AuxOutput4_Load")
                StationRecipe(iStation, iShift).AuxOutputs_Purge(1) = rsRecord("AuxOutput1_Purge")
                StationRecipe(iStation, iShift).AuxOutputs_Purge(2) = rsRecord("AuxOutput2_Purge")
                StationRecipe(iStation, iShift).AuxOutputs_Purge(3) = rsRecord("AuxOutput3_Purge")
                StationRecipe(iStation, iShift).AuxOutputs_Purge(4) = rsRecord("AuxOutput4_Purge")
            End If
               
            rsRecord.Close
    
            
            StationRecipe(iStation, iShift).CyclesSave = StationRecipe(iStation, iShift).Cycles
            StationRecipe(iStation, iShift).Load_MethodSave = StationRecipe(iStation, iShift).Load_Method
            StationRecipe(iStation, iShift).Load_RateSave = StationRecipe(iStation, iShift).Load_Rate
            StationRecipe(iStation, iShift).NitrogenFlowSave = StationRecipe(iStation, iShift).NitrogenFlow
            StationRecipe(iStation, iShift).WC_MultSave = StationRecipe(iStation, iShift).WC_Mult
            
            ' Update Station Recipe description fields
            UpdateStnRcpDsc iStation, iShift
            
            ' update live fuel parameters for adf
            LiveFuel_Update iStation, iShift
            
            ' Check Recipe Number
            ' note: Station Canister Recipe Tracking was added as part of CfgRevLvl=5
            If CfgRevLvl < 5 Then StationRecipe(iStation, iShift).Number = 0
    
        Next iShift
    Next iStation
    
    ' Close Database
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

Sub Load_ScaleConfig()
' Load Scale Configuration
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1434
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iScale As Integer

    sFileName = FILEPATH_cfg & "scales.@@@"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    '  Scale configurations
    For iScale = 1 To MAX_SCALES
    
        Input #iFileNumber, Scale_Port(iScale)
        Input #iFileNumber, Scale_Type(iScale)
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
    Next iScale
    
    Close #iFileNumber

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

Sub Update_AnalogFuncDef(ByVal oldMaxAnaStn As Integer, ByVal newMaxAnaStn As Integer)
'
' Update Analog Functions Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 9494
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim iPurge As Integer
Dim Idx As Integer
Dim inct As Integer
Dim inct2 As Integer
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")

    ' Does the config file exist
    sFileName = FILEPATH_cfg & "funcdefa.@@@"
    If fs.FileExists(sFileName) = True Then
    
        ' Analog Function Definitions configuration file exists; Read it
        iFileNumber = FreeFile
        Open sFileName For Input As #iFileNumber
        
        '  COMMON
        For Idx = 0 To MAX_ANA_COM
        
            Input #iFileNumber, Com_AIO(Idx).EuMax
            Input #iFileNumber, Com_AIO(Idx).EuMin
            Input #iFileNumber, Com_AIO(Idx).VdcMax
            Input #iFileNumber, Com_AIO(Idx).VdcMin
            Input #iFileNumber, Com_AIO(Idx).addr
            Input #iFileNumber, Com_AIO(Idx).chan
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
        Next Idx
        
        '  FID
        For Idx = 0 To MAX_ANA_FID
        
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
        Next Idx
        
        '  STATIONS
        For iStation = 1 To MAX_STN
            For Idx = 0 To oldMaxAnaStn
        
                Input #iFileNumber, Stn_AIO(iStation, Idx).EuMax
                Input #iFileNumber, Stn_AIO(iStation, Idx).EuMin
                Input #iFileNumber, Stn_AIO(iStation, Idx).VdcMax
                Input #iFileNumber, Stn_AIO(iStation, Idx).VdcMin
                Input #iFileNumber, Stn_AIO(iStation, Idx).addr
                Input #iFileNumber, Stn_AIO(iStation, Idx).chan
                
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                
            Next Idx
            
        Next iStation
        
        If CfgRevLvl > 0 Then
            '  PurgeAir Sources
            For iPurge = 1 To MAX_PRG
                For Idx = 0 To MAX_ANA_PRG
            
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).EuMax
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).EuMin
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).VdcMax
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).VdcMin
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).addr
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).chan
                    
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    
                Next Idx
                
            Next iPurge
        End If
        
        Close #iFileNumber
        
    End If
    
    sFileName = FILEPATH_cfg & "funcdefa.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    '  COMMON
    '  Analog Functions
    For inct2 = 0 To MAX_ANA_COM
    
        Write #iFileNumber, Com_AIO(inct2).EuMax
        Write #iFileNumber, Com_AIO(inct2).EuMin
        Write #iFileNumber, Com_AIO(inct2).VdcMax
        Write #iFileNumber, Com_AIO(inct2).VdcMin
        Write #iFileNumber, Com_AIO(inct2).addr
        Write #iFileNumber, Com_AIO(inct2).chan
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next inct2
    
    '  FID
    '  Analog Functionss
    For inct2 = 0 To MAX_ANA_FID
    
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next inct2
    
    '  Stations
    For inct = 1 To MAX_STN
        '  Analog Functions
        For inct2 = 0 To newMaxAnaStn
    
            Write #iFileNumber, Stn_AIO(inct, inct2).EuMax
            Write #iFileNumber, Stn_AIO(inct, inct2).EuMin
            Write #iFileNumber, Stn_AIO(inct, inct2).VdcMax
            Write #iFileNumber, Stn_AIO(inct, inct2).VdcMin
            Write #iFileNumber, Stn_AIO(inct, inct2).addr
            Write #iFileNumber, Stn_AIO(inct, inct2).chan
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
        Next inct2
        
    Next inct
    
    '  PurgeAir Sources
    For inct = 1 To MAX_PRG
        '  Analog Functionss
        For inct2 = 0 To MAX_ANA_PRG
    
            Write #iFileNumber, Prg_AIO(inct, inct2).EuMax
            Write #iFileNumber, Prg_AIO(inct, inct2).EuMin
            Write #iFileNumber, Prg_AIO(inct, inct2).VdcMax
            Write #iFileNumber, Prg_AIO(inct, inct2).VdcMin
            Write #iFileNumber, Prg_AIO(inct, inct2).addr
            Write #iFileNumber, Prg_AIO(inct, inct2).chan
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
        Next inct2
        
    Next inct
    
    Close #iFileNumber
    
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

Sub Load_AnalogFuncDef()
' Load Station Analog Functions Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1494
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim iPurge As Integer
Dim inct3 As Integer
Dim Idx As Integer
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")

    ' Clear the Analog Definition Arrays
    '  COMMON
    For Idx = 0 To MAX_ANA_COM
        Com_AIO(Idx).EuMax = 0
        Com_AIO(Idx).EuMin = 0
        Com_AIO(Idx).VdcMax = 0
        Com_AIO(Idx).VdcMin = 0
        Com_AIO(Idx).addr = 0
        Com_AIO(Idx).chan = 0
    Next Idx
    
    '  STATIONS
    For iStation = 1 To MAX_STN
        For Idx = 0 To MAX_ANA_STN
            Stn_AIO(iStation, Idx).EuMax = 0
            Stn_AIO(iStation, Idx).EuMin = 0
            Stn_AIO(iStation, Idx).VdcMax = 0
            Stn_AIO(iStation, Idx).VdcMin = 0
            Stn_AIO(iStation, Idx).addr = 0
            Stn_AIO(iStation, Idx).chan = 0
        Next Idx
    Next iStation
    
    '  PurgeAir Sources
    For iPurge = 1 To MAX_PRG
        For Idx = 0 To MAX_ANA_PRG
            Prg_AIO(iPurge, Idx).EuMax = 0
            Prg_AIO(iPurge, Idx).EuMin = 0
            Prg_AIO(iPurge, Idx).VdcMax = 0
            Prg_AIO(iPurge, Idx).VdcMin = 0
            Prg_AIO(iPurge, Idx).addr = 0
            Prg_AIO(iPurge, Idx).chan = 0
        Next Idx
    Next iPurge

    ' Does the config file exist
    sFileName = FILEPATH_cfg & "funcdefa.@@@"
    If fs.FileExists(sFileName) = True Then
    
        ' Analog Function Definitions configuration file exists; Read it
        iFileNumber = FreeFile
        Open sFileName For Input As #iFileNumber
        
        '  COMMON
        For Idx = 0 To MAX_ANA_COM
        
            Input #iFileNumber, Com_AIO(Idx).EuMax
            Input #iFileNumber, Com_AIO(Idx).EuMin
            Input #iFileNumber, Com_AIO(Idx).VdcMax
            Input #iFileNumber, Com_AIO(Idx).VdcMin
            Input #iFileNumber, Com_AIO(Idx).addr
            Input #iFileNumber, Com_AIO(Idx).chan
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            If (Com_AIO(Idx).addr > OptoMaxNodeNum) Then
                inct3 = Com_AIO(Idx).addr
                Com_AIO(Idx).addr = 0
                Com_AIO(Idx).chan = 0
                Com_AIO(Idx).EuMax = 0#
                Com_AIO(Idx).EuMin = 0#
                Com_AIO(Idx).VdcMax = 0#
                Com_AIO(Idx).VdcMin = 0#
            End If
            
        Next Idx
        
        '  FID
        For Idx = 0 To MAX_ANA_FID
        
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
        Next Idx
        
        '  STATIONS
        For iStation = 1 To MAX_STN
            For Idx = 0 To MAX_ANA_STN
'            For idx = 0 To 29
                Input #iFileNumber, Stn_AIO(iStation, Idx).EuMax
                Input #iFileNumber, Stn_AIO(iStation, Idx).EuMin
                Input #iFileNumber, Stn_AIO(iStation, Idx).VdcMax
                Input #iFileNumber, Stn_AIO(iStation, Idx).VdcMin
                Input #iFileNumber, Stn_AIO(iStation, Idx).addr
                Input #iFileNumber, Stn_AIO(iStation, Idx).chan
                
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                
                If (Stn_AIO(iStation, Idx).addr > OptoMaxNodeNum) Then
                    inct3 = Stn_AIO(iStation, Idx).addr
                    Stn_AIO(iStation, Idx).addr = 0
                    Stn_AIO(iStation, Idx).chan = 0
                    Stn_AIO(iStation, Idx).EuMax = 0#
                    Stn_AIO(iStation, Idx).EuMin = 0#
                    Stn_AIO(iStation, Idx).VdcMax = 0#
                    Stn_AIO(iStation, Idx).VdcMin = 0#
                End If
            
            Next Idx
        Next iStation
        
        If CfgRevLvl > 0 Then
            '  PurgeAir Sources
            For iPurge = 1 To MAX_PRG
                For Idx = 0 To MAX_ANA_PRG
            
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).EuMax
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).EuMin
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).VdcMax
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).VdcMin
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).addr
                    Input #iFileNumber, Prg_AIO(iPurge, Idx).chan
                    
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                    Input #iFileNumber, filler
                        
                    If (Prg_AIO(iPurge, Idx).addr > OptoMaxNodeNum) Then
                        inct3 = Prg_AIO(iPurge, Idx).addr
                        Prg_AIO(iPurge, Idx).addr = 0
                        Prg_AIO(iPurge, Idx).chan = 0
                        Prg_AIO(iPurge, Idx).EuMax = 0#
                        Prg_AIO(iPurge, Idx).EuMin = 0#
                        Prg_AIO(iPurge, Idx).VdcMax = 0#
                        Prg_AIO(iPurge, Idx).VdcMin = 0#
                    End If
            
                Next Idx
                
            Next iPurge
        End If
        
        Close #iFileNumber
        
    Else
    
        ' No Analog Function Definition Configuration File
        Delay_Box "Analog Function Definition File Not Found; values set to zero", MSGDELAY, msgSHOW
            
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

Sub Load_StationInfo()
' Save Station Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1475
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim iScale As Integer

    sFileName = FILEPATH_cfg & "stations.@@@"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    '  Station Definition Information
    For iStation = 1 To NR_STN
    
        Input #iFileNumber, STN_INFO(iStation).ADF_StnNum
        Input #iFileNumber, STN_INFO(iStation).ADF_TANKTYPE
        Input #iFileNumber, STN_INFO(iStation).AspiratorNum
        Input #iFileNumber, STN_INFO(iStation).DefAuxScale
        Input #iFileNumber, STN_INFO(iStation).DefPriScale
        Input #iFileNumber, STN_INFO(iStation).desc
        Input #iFileNumber, STN_INFO(iStation).Type
        
        Input #iFileNumber, STN_INFO(iStation).ButMfcDensityMult
        Input #iFileNumber, STN_INFO(iStation).ButMfc2DensityMult
        Input #iFileNumber, STN_INFO(iStation).ADF_HEATERTYPE
        Input #iFileNumber, STN_INFO(iStation).USINGPURGEOVEN
        Input #iFileNumber, STN_INFO(iStation).Abrev
        
        Input #iFileNumber, STN_INFO(iStation).SysID
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
        ' default Butane MFC Density Multiplier is 1.0; acceptable range is 0.9 - 1.1
        If STN_INFO(iStation).ButMfcDensityMult = 0 Then STN_INFO(iStation).ButMfcDensityMult = 1
        If STN_INFO(iStation).ButMfc2DensityMult = 0 Then STN_INFO(iStation).ButMfc2DensityMult = 1
        
        If (USINGHARDPIPEDSCALES And (STN_INFO(iStation).Type <> STN_LEAKTEST_TYPE)) Then
            ' two scales per station, fixed assignments for pri & aux for each station; stn#1 pri = 1, stn#1 aux = 2, etc.
            STN_INFO(iStation).DefPriScale = 1 + (2 * (iStation - 1))
            STN_INFO(iStation).DefAuxScale = 1 + STN_INFO(iStation).DefPriScale
            STN_INFO(STN_INFO(iStation).DefPriScale).AspiratorNum = STN_INFO(iStation).AspiratorNum
            STN_INFO(STN_INFO(iStation).DefAuxScale).AspiratorNum = STN_INFO(iStation).AspiratorNum
        End If
        ' set ADF_DEF bits (from TankType)
        DecodeAdfDef iStation
        AdfDef(iStation) = STN_INFO(iStation).ADF_DEF
        AdfControl(iStation).AdfDefinition = STN_INFO(iStation).ADF_DEF
        
    Next iStation
    
    Close #iFileNumber
    
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

Sub Load_SysDef()
' Load System definition Values
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 141
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim temp As Single
Dim debugInt As Integer
Dim usingInt1 As Integer
Dim usingInt2 As Integer
Dim usingInt3 As Integer
Dim usingInt4 As Integer
Dim usingInt5 As Integer
Dim usingInt6 As Integer
Dim usingInt7 As Integer
Dim usingInt8 As Integer
Dim usingInt9 As Integer
Dim usingInt10 As Integer
Dim debugBits As SixteenBits
Dim usingBits1 As SixteenBits
Dim usingBits2 As SixteenBits
Dim usingBits3 As SixteenBits
Dim usingBits4 As SixteenBits
Dim usingBits5 As SixteenBits
Dim usingBits6 As SixteenBits
Dim usingBits7 As SixteenBits
Dim usingBits8 As SixteenBits
Dim usingBits9 As SixteenBits
Dim usingBits10 As SixteenBits

    sFileName = FILEPATH_cfg & "sysdef.@@@"
    iFileNumber = FreeFile
    
    Open sFileName For Input As #iFileNumber
' ####  1 - 10
    Input #iFileNumber, OPTOCOM_PORT               ' opto AC24 = com2 ; sealevel = com5 ; Moxa = com6
    Input #iFileNumber, filler
    Input #iFileNumber, LoadMfcDelayTime           ' number of seconds between switching load valves & load mfc's
    Input #iFileNumber, NR_SCALES                  ' Number of Scales to use S & T & A
    Input #iFileNumber, NR_SHIFT                   ' Number of shifts
    Input #iFileNumber, NR_STN                     ' Number of stations
    Input #iFileNumber, NR_DUMMYSTN                ' Number of dummy stations (i.e. io board only)
    Input #iFileNumber, NR_REMOTESCALES            ' Number of Remote Scales (i.e. not mounted in the Station Cabinet); 0 = None
    Input #iFileNumber, LocalPagControl.Type       ' 0=None, 1=Stand-Alone, 2=AkMaster, 3=AkClient)
    Input #iFileNumber, MSGDELAY                   ' DelayBox open time in ms
' ####  11 - 20
    Input #iFileNumber, MFC_Settle_Time            ' In seconds
    Input #iFileNumber, LoadEqlDelayTime           ' Time in seconds for Load Flow & Scales to reach "equalibrium"
    Input #iFileNumber, usingInt1
    Input #iFileNumber, usingInt2
    Input #iFileNumber, usingInt3
    Input #iFileNumber, usingInt4
    Input #iFileNumber, usingInt5
    Input #iFileNumber, usingInt6
    Input #iFileNumber, usingInt7
    Input #iFileNumber, usingInt8
' ####  21 - 30
    Input #iFileNumber, usingInt9
    Input #iFileNumber, usingInt10
    Input #iFileNumber, PAGSERVERIP                ' IP address of this client's pag server
    Input #iFileNumber, TOM_2Gm_Recipe
    Input #iFileNumber, TOM_Wcm_Recipe
    Input #iFileNumber, Chiller_PORT
    Input #iFileNumber, Chiller_Timeout
    Input #iFileNumber, USING_EXT_CONTACTS          ' External alarm contacts - Remote alarm Pause
    Input #iFileNumber, DESC_EXT_CONTACTS           ' Description of external alarm
    Input #iFileNumber, USINGUPS                    ' No ups installed=0/Large ups timed=1/small down now ups=2
' ####  31 - 40
    Input #iFileNumber, MinDataLogSeconds           ' Minimum allowed data log interval in seconds
    Input #iFileNumber, USING_AUX_OUTPUTS           ' Aux (12vdc or Dry Contact) Outputs
    Input #iFileNumber, NR_AUX_OUTPUTS              ' Number of Aux Outputs/Station
    Input #iFileNumber, DESC_AUX_OUTPUT1            ' Description of aux output #1
    Input #iFileNumber, DESC_AUX_OUTPUT2            ' Description of aux output #2
    Input #iFileNumber, DESC_AUX_OUTPUT3            ' Description of aux output #3
    Input #iFileNumber, DESC_AUX_OUTPUT4            ' Description of aux output #4
    Input #iFileNumber, MAXNOTSTABLECOUNT           ' Max Allowed Consecutive Unstable Scale reads - if exceeded then read is "declared" Stable
    Input #iFileNumber, WEIGHTQUEUESIZE
    Input #iFileNumber, MaxSheathTempForAdfDrain
' ####  41 - 50
    Input #iFileNumber, DefScaleMax                 ' Default Max Scale Reading (Default Min = 0)
    Input #iFileNumber, GramsPerLiter               ' grams per Slpm (default = 2.40633)
    Input #iFileNumber, WB_AIO.EuMin
    Input #iFileNumber, NR_CAN                      ' Number of Master Canister Definitions
    Input #iFileNumber, NR_RCP                      ' Number of Master Recipes
    Input #iFileNumber, WB_AIO.EuMax
    Input #iFileNumber, AutoLogon                   ' 0 = No Auto Logon; 1 = user; 2 = cps; 3 = admin; 4 = aps; 4 = ApsUser
    Input #iFileNumber, debugInt                    ' Combined Debug Bits
    Input #iFileNumber, NR_JOBSEQ                   ' Number of Master Job Sequences
    Input #iFileNumber, MfcSpMin
' ####  51 - 60
    Input #iFileNumber, DeadLiveFuelDensity
    Input #iFileNumber, WeakLiveFuelDensity
    Input #iFileNumber, SystemTimers(9).Interval
    Input #iFileNumber, SystemTimers(8).Interval
    Input #iFileNumber, SystemTimers(7).Interval
    Input #iFileNumber, NR_PRGAIR                   ' Number of PurgeAir Sources
    Input #iFileNumber, SystemTimers(1).Interval
    Input #iFileNumber, SystemTimers(2).Interval
    Input #iFileNumber, SystemTimers(3).Interval
    Input #iFileNumber, SystemTimers(4).Interval
' ####  61 - 64
    Input #iFileNumber, SystemTimers(5).Interval
    Input #iFileNumber, SystemTimers(6).Interval
            ' last three values are Read-ONLY
    Input #iFileNumber, filler                      ' saved program revision level (can't be loaded; it is a global const)
    Input #iFileNumber, filler                      ' Revision Level required of the Report Generator program
    Input #iFileNumber, filler                      ' Revision Level of the Cfg & SysDef Files
    
endofnew:

    Close #iFileNumber


    ' separate debug Bits
    '
    '   bit 0 = UseLocalErrorHandler
    '   bit 1 = NotDebugADF
    '   bit 2 = NotDebugMMW
    '   bit 3 = NotDebugPURGE
    '
    '   bit 4 = NotDebugSCALES
    '   bit 5 = NotDebugPAS
    '   bit 6 = unused
    '   bit 7 = unused
    '
    debugBits = Bits_UnPack(debugInt)
    UseLocalErrorHandler = debugBits.B00
    NotDebugADF = debugBits.B01
    NotDebugMMW = debugBits.B02
    NotDebugPURGE = debugBits.B03
    NotDebugSCALES = debugBits.B04
    NotDebugPAS = debugBits.B05
'    unused = debugBits.B06
'    unused = debugBits.B07
     
    ' separate using Bits #1
    '
    '   bit 0 = USINGCOMMONTC              ' Six common thermocouples (Expert so far) 12/18/02
    '   bit 1 = USINGDOOROPEN              ' Door open switch installed
    '   bit 2 = USINGF                     ' Fht scale for temps
    '   bit 3 = USINGHIGHTEMPPAS
    '
    '   bit 4 = USINGPASLOCALCONTROL       ' Local PAS Control
    '   bit 5 = USINGLEAKCHECKEXHAUSTSOL   ' Leak Check Exhaust Sol on this system
    '   bit 6 = unused
    '   bit 7 = USINGSIMNOISE              ' add noise (max 1 %) to simulated eu values
    '
    usingBits1 = Bits_UnPack(usingInt1)
    USINGCOMMONTC = usingBits1.B00
    USINGDOOROPEN = usingBits1.B01
    USINGF = usingBits1.B02
    USINGHIGHTEMPPAS = usingBits1.B03
    USINGPASLOCALCONTROL = usingBits1.B04
    USINGLEAKCHECKEXHAUSTSOL = usingBits1.B05
'    unused = usingBits1.B06
    USINGSIMNOISE = usingBits1.B07
    
    ' separate using Bits #2
    '
    '   bit 0 = USINGBUTANEMASSLIMIT           ' One is for (was Toyota only) Limit Exceeded...Aborting
    '   bit 1 = USINGLOADTIMELIMIT       ' For Mitsu abort by operator entered time on recipe "MAX LOAD TIME"
    '   bit 2 = USINGLOADPRESSURE          ' Abort load on excessive load pressure (was Toyota only)
    '   bit 3 = USINGCUSTOMERLOWGAS        ' No remote inputs (was Toyota only)
    '
    '   bit 4 = USINGCANVENTALARM          ' Carb #1 to use (Hardware required in opto stn)
    '   bit 5 = USINGCONTAFTERLCFAIL       ' Allow Continue after a Leak Check Failure
    '   bit 6 = USINGLINEVOLUME            ' Volume compensation for line lengths (CARB first user) 5/18/03 Smitty
    '   bit 7 = USINGOOTPAUSE              ' Allow Pause of station/shift on OOT condition (Config option can still turn it off)
    '
    usingBits2 = Bits_UnPack(usingInt2)
    USINGBUTANEMASSLIMIT = usingBits2.B00
    USINGLOADTIMELIMIT = usingBits2.B01
    USINGLOADPRESSURE = usingBits2.B02
    USINGCUSTOMERLOWGAS = usingBits2.B03
    USINGCANVENTALARM = usingBits2.B04
    USINGCONTAFTERLCFAIL = usingBits2.B05
    USINGLINEVOLUME = usingBits2.B06
    USINGOOTPAUSE = usingBits2.B07
    
    ' separate using Bits #3
    '
    '   bit 0 = USINGREMCANLOAD            ' Task Order Manager Tasks
    '   bit 1 = LogTempRh                  ' Log Air Temperature and Humidity
    '   bit 2 = USING_ESTOP_INPUT          ' ESTOP (12vdc or Dry Contact) Input
    '   bit 3 = USINGSTNTC                 ' Use ThermoCouple (HowMany) Zero None + one for each per station
    '
    '   bit 4 = USINGMoist_RH              ' Display Moisture in grains per pound
    '   bit 5 = USINGPRESSUREPURGE         ' Yes system has config option of positive pressure purge (0 = normal)
    '   bit 6 = USINGSIMULATION            ' Run Simulation when IOScan is Off
    '   bit 7 = USINGLVol_SI               ' Using SI Units (mm & meters) for Line Vol ID & Length
    '
    usingBits3 = Bits_UnPack(usingInt3)
    USINGREMCANLOAD = usingBits3.B00
    LogTempRh = usingBits3.B01
    USING_ESTOP_INPUT = usingBits3.B02
    USINGSTNTC = usingBits3.B03
    USINGMoist_RH = usingBits3.B04
    USINGPRESSUREPURGE = usingBits3.B05
    USINGSIMULATION = usingBits3.B06
    USINGLVol_SI = usingBits3.B07
    
    ' separate using Bits #4
    '
    '   bit 0 = USINGHARDPIPEDSCALES       ' Scales are hard piped at 2 scales per station; stn#1 pri = 1, stn#1 aux = 2, etc.
    '   bit 1 = USINGAUXLEAKCHECK          ' Allows Leakcheck of Aux plumbing; Leakcheck Aux Only and Leakcheck Both (pri & aux)
    '   bit 2 = USINGERRORMSGBYPASS        ' Allow Error Handler to "bypass" the Error Message MsgBox
    '   bit 3 = USINGSYSTEMVACSW           ' System has a Master Vacuum Switch to be monitored
    '
    '   bit 4 = USINGPURGEDP               ' System has a DP transmitter for use during Purge
    '   bit 5 = USINGPURGESERIES           ' System is capable of Purging (Pri & Aux) Canisters in Series
    '   bit 6 = IoComOn                    ' True = Run with I/O comm, False = testing without IO
    '   bit 7 = SclComOn                   ' True = Run with Scale Port comm, False = testing without reading Scales
    '
    usingBits4 = Bits_UnPack(usingInt4)
    USINGHARDPIPEDSCALES = usingBits4.B00
    USINGAUXLEAKCHECK = usingBits4.B01
    USINGERRORMSGBYPASS = usingBits4.B02
    USINGSYSTEMVACSW = usingBits4.B03
    USINGPURGEDP = usingBits4.B04
    USINGPURGESERIES = usingBits4.B05
    IoComOn = usingBits4.B06
    SclComOn = usingBits4.B07
    
    ' combine using Bits #5
    '
    '   bit 0 = ChillComOn                 ' True = Run with Chiller comm, False = testing without Chiller
    '   bit 1 = USINGPURGEOVEN             ' System has one or more Purge Ovens
    '   bit 2 = USINGWATERBATH             ' System has a WaterBath Heater/Chiller for LiveFuel
    '   bit 3 = USINGFUELLEVELOOT          ' enable check of livefuel tank level
    '
    '   bit 4 = USINGDRYPURGEAIR
    '   bit 5 = USINGREMAVLFILES
    '   bit 6 = USINGREMSTSMON
    '   bit 7 = REMCHGSENABLED
    '
    usingBits5 = Bits_UnPack(usingInt5)
    ChillComOn = usingBits5.B00
    USINGPURGEOVEN = usingBits5.B01
    USINGWATERBATH = usingBits5.B02
    USINGFUELLEVELOOT = usingBits5.B03
    USINGDRYPURGEAIR = usingBits5.B04
    USINGREMAVLFILES = usingBits5.B05
    USINGREMSTSMON = usingBits5.B06
    REMCHGSENABLED = usingBits5.B07
    
    ' combine using Bits #6
    '
    '   bit 0 =
    '   bit 1 =
    '   bit 2 =
    '   bit 3 =
    '
    '   bit 4 =
    '   bit 5 =
    '   bit 6 =
    '   bit 7 =
    '
    usingBits6 = Bits_UnPack(usingInt6)
    USINGTOMCANLOAD = usingBits6.B00
    Spare1 = usingBits6.B01
    Spare2 = usingBits6.B02
    Spare3 = usingBits6.B03
    Spare4 = usingBits6.B04
    Spare5 = usingBits6.B05
    Spare6 = usingBits6.B06
    Spare7 = usingBits6.B07
    '
    ' combine using Bits #7
    '
    '   bit 0 =
    '   bit 1 =
    '   bit 2 =
    '   bit 3 =
    '
    '   bit 4 =
    '   bit 5 =
    '   bit 6 =
    '   bit 7 =
    '
    usingBits7 = Bits_UnPack(usingInt7)

    ' combine using Bits #8
    '
    '   bit 0 =
    '   bit 1 =
    '   bit 2 =
    '   bit 3 =
    '
    '   bit 4 =
    '   bit 5 =
    '   bit 6 =
    '   bit 7 =
    '
    usingBits8 = Bits_UnPack(usingInt8)
 
    ' combine using Bits #9
    '
    '   bit 0 =
    '   bit 1 =
    '   bit 2 =
    '   bit 3 =
    '
    '   bit 4 =
    '   bit 5 =
    '   bit 6 =
    '   bit 7 =
    '
    usingBits9 = Bits_UnPack(usingInt9)

    ' combine using Bits #10
    '
    '   bit 0 =
    '   bit 1 =
    '   bit 2 =
    '   bit 3 =
    '
    '   bit 4 =
    '   bit 5 =
    '   bit 6 =
    '   bit 7 =
    '
    usingBits10 = Bits_UnPack(usingInt10)
        
' ###################################################################################################################
    NotDebugProf = True                             ' NotDebugProf is hardcoded; i.e. not included in sysdef
    NotDebugREM = True                              ' NotDebugREM  is hardcoded; i.e. not included in sysdef

    If MSGDELAY < 1 Then MSGDELAY = 1000            ' default DelayBox open time = 1000ms
    If WEIGHTQUEUESIZE < 1 Then WEIGHTQUEUESIZE = 1 ' size of the queue used for Scale Weight Running Averages
    If WEIGHTQUEUESIZE > MAXWEIGHTQUEUE Then WEIGHTQUEUESIZE = MAXWEIGHTQUEUE

    If Not IsNumeric(MfcSpMin) Then MfcSpMin = 5#   ' default MFC SetPoint Minimum = 5.0 % of Fullscale
    If MfcSpMin < 1# Then MfcSpMin = 5#             ' default MFC SetPoint Minimum = 5.0 % of Fullscale
    If MfcSpMin > 5# Then MfcSpMin = 5#             ' default MFC SetPoint Minimum = 5.0 % of Fullscale
    
    LAST_STN = NR_STN - NR_DUMMYSTN                         ' Station Number of the Last Regular (i.e. Not Dummy) Station
    FIRST_REMOTESCALE = NR_SCALES - NR_REMOTESCALES + 1     ' Scale Number of the First Remote Scale
    
    USINGC = IIf(USINGF, False, True)
    USINGLVol_Engl = IIf(USINGLVol_SI, False, True)
    USINGMoist_Grains = IIf(USINGMoist_RH, False, True)
    
    ' No Interface if nothing to interface
    If ((Not USINGREMCANLOAD) And (Not USINGREMSTSMON)) Then USINGREMAVLFILES = False
    
    ' Check Interval Timer Settings
    ' note: recommended settings were changed as part of CfgRevLvl=3
    If CfgRevLvl < 3 Then
        SystemTimers(1).Interval = 70       ' scan io
        SystemTimers(2).Interval = 10       ' scan (scale) comm ports
        SystemTimers(3).Interval = 250      ' alarm/oot logic
        SystemTimers(4).Interval = 20       ' datalogger
        SystemTimers(5).Interval = 100      ' controllers logic (prg, adf, n2push, pas, sim, etc)
        SystemTimers(6).Interval = 50       ' station logic
        SystemTimers(7).Interval = 10       ' timer control logic
        SystemTimers(8).Interval = 1000     ' unused
        SystemTimers(9).Interval = 1000     ' unused
    End If
    
    ' Check Using_Simulation
    ' note: Simulation Controls were added as part of CfgRevLvl=3
    If CfgRevLvl < 3 Then USINGSIMULATION = False  ' default is OFF
    
    ' Check Auto Logon Settings
    ' note: recommended settings were changed as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then AutoLogon = 0       ' No AutoLogon
    
    ' Check Using Continue After Leak Check Failure Settings
    ' note: recommended settings were added as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then USINGCONTAFTERLCFAIL = False       ' No Continue After Leak Check Failure
    
    ' Check Number of Canisters/Recipes Settings
    ' note: settings were changed as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then
        If NR_CAN < 1 Or NR_CAN > MAX_CANRCP Then NR_CAN = MAX_CANRCP
        If NR_RCP < 1 Or NR_RCP > MAX_RCP Then NR_RCP = MAX_RCP
    Else
        If NR_CAN < 1 Or NR_CAN > 999 Then NR_CAN = 999
        If NR_RCP < 1 Or NR_RCP > 999 Then NR_RCP = 999
    End If
    
    ' Check GramsPerLiter Setting
    ' note: GramsPerLiter conversion factor was added as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then GramsPerLiter = 2.40633                       ' default
    If IsEmpty(GramsPerLiter) Then GramsPerLiter = 2.40633              ' default
    If Not IsNumeric(GramsPerLiter) Then GramsPerLiter = 2.40633        ' default
    If GramsPerLiter < 2 Then GramsPerLiter = 2.40633                   ' default
    If GramsPerLiter > 3 Then GramsPerLiter = 2.40633                   ' default
    
    ' check DefaultScaleMax (min Max = 100; max Max = 100,000)
    If DefScaleMax < 100 Then DefScaleMax = 100                         ' default
    If DefScaleMax > 100000 Then DefScaleMax = 100000                   ' default
    
    DebugAIRLOG = True
'    DebugAIRLOG = False

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

Sub Save_NodeInfo()
' Save OPTO Node Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2792
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim inct As Integer

sFileName = FILEPATH_cfg & "nodeinfo.@@@"
iFileNumber = FreeFile
Open sFileName For Output As #iFileNumber

'  OPTO Node Information
For inct = 0 To MAX_NODE

    Write #iFileNumber, Node_Info(inct)
    
    Write #iFileNumber, filler
    Write #iFileNumber, filler
    Write #iFileNumber, filler
    Write #iFileNumber, filler
    Write #iFileNumber, filler
    
Next inct

Close #iFileNumber

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

Sub Save_OptoInfo()
' Save OPTO Module Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2222
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim inct As Integer
Dim inct2 As Integer

    sFileName = FILEPATH_cfg & "optoinfo.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    '  OPTO Module Information
    For inct = 0 To MAX_ADDR
        For inct2 = 0 To MAX_SLOT
    
        Write #iFileNumber, Opto_Info(inct, inct2)
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Next inct2
    Next inct
    
    Close #iFileNumber

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

Sub Save_Controllers()
' Save Controller information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 7510

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim iController As Integer


    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
        For iController = 1 To MAX_CONTROLLER
            ' Save Controller Configuration Records
            Criteria = "SELECT * FROM [Controllers] WHERE [Number] = " & iController & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Number") = iController
            Else
                rsRecord.MoveFirst
                rsRecord.Edit
            End If
               
            rsRecord("Pgain") = PID_INFO(iController).Pgain
            rsRecord("Igain") = PID_INFO(iController).Igain
            rsRecord("Dgain") = PID_INFO(iController).Dgain
            rsRecord("ReverseAction") = PID_INFO(iController).Rev
            rsRecord("OffDuty") = PID_INFO(iController).OffDuty
            rsRecord("OnDuty") = PID_INFO(iController).OnDuty
            rsRecord("OffLimitDelta") = PID_INFO(iController).OffLimitDelta
            rsRecord("OnLimitDelta") = PID_INFO(iController).OnLimitDelta
            rsRecord("OffDutyMult") = PID_INFO(iController).OffDutyMult
            rsRecord("OnDutyMult") = PID_INFO(iController).OnDutyMult
               
            rsRecord.Update
            rsRecord.Close
        Next iController
    
    ' Close Database
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

Sub Load_LocalPAS()
' Load Local PAS Controller configuration information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 6520

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim Idx As Integer

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
        For Idx = 1 To 2
            ' Read Local PAS Controller Configuration Records
            Criteria = "SELECT * FROM [Local_PAS] WHERE [Number] = " & Idx & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                PAS_INFO(Idx).DurationTarget = 120
                PAS_INFO(Idx).TimeOutTarget = 900
            Else
                rsRecord.MoveFirst
                PAS_INFO(Idx).DurationTarget = rsRecord("InTolTargetDuration")
                PAS_INFO(Idx).TimeOutTarget = rsRecord("TimeoutDuration")
            End If
               
            rsRecord.Close
    
            ' PAS Info tracking
            last_INFO(Idx) = PAS_INFO(Idx)
        
        Next Idx
            
    ' Close Database
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

Sub Save_LocalPAS()
' Save Local PAS Controller information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 6510

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim Idx As Integer


    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
        For Idx = 1 To 2
            ' Save Local PAS Controller Configuration Records
            Criteria = "SELECT * FROM [Local_PAS] WHERE [Number] = " & Idx & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Number") = Idx
            Else
                rsRecord.MoveFirst
                rsRecord.Edit
            End If
               
            rsRecord("InTolTargetDuration") = PAS_INFO(Idx).DurationTarget
            rsRecord("TimeoutDuration") = PAS_INFO(Idx).TimeOutTarget
               
            rsRecord.Update
            rsRecord.Close
        Next Idx
    
    ' Close Database
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

Sub Load_Controllers()
' Load Controller configuration information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 7520

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim iController As Integer

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
        For iController = 1 To MAX_CONTROLLER
            ' Read Controller Configuration Records
            Criteria = "SELECT * FROM [Controllers] WHERE [Number] = " & iController & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                PID_INFO(iController).Pgain = 1#
                PID_INFO(iController).Igain = 1#
                PID_INFO(iController).Dgain = 1#
                PID_INFO(iController).Rev = False
                PID_INFO(iController).OffDuty = 1#
                PID_INFO(iController).OnDuty = 1#
                PID_INFO(iController).OffLimitDelta = 1#
                PID_INFO(iController).OnLimitDelta = 1#
                PID_INFO(iController).OffDutyMult = 1#
                PID_INFO(iController).OnDutyMult = 1#
                PID_INFO(iController).CumImax = CSng(70)
                PID_INFO(iController).CumImin = CSng(-70)
                PID_INFO(iController).outmax = CSng(95)
                PID_INFO(iController).outmin = CSng(5)
            Else
                rsRecord.MoveFirst
                PID_INFO(iController).Pgain = rsRecord("Pgain")
                PID_INFO(iController).Igain = rsRecord("Igain")
                PID_INFO(iController).Dgain = rsRecord("Dgain")
                PID_INFO(iController).Rev = rsRecord("ReverseAction")
                PID_INFO(iController).OffDuty = rsRecord("OffDuty")
                PID_INFO(iController).OnDuty = rsRecord("OnDuty")
                PID_INFO(iController).OffLimitDelta = rsRecord("OffLimitDelta")
                PID_INFO(iController).OnLimitDelta = rsRecord("OnLimitDelta")
                PID_INFO(iController).OffDutyMult = rsRecord("OffDutyMult")
                PID_INFO(iController).OnDutyMult = rsRecord("OnDutyMult")
                PID_INFO(iController).CumImax = rsRecord("CumImax")
                PID_INFO(iController).CumImin = rsRecord("CumImin")
                PID_INFO(iController).outmax = rsRecord("Outmax")
                PID_INFO(iController).outmin = rsRecord("Outmin")
            End If
               
            rsRecord.Close
        Next iController
            
    ' Close Database
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

Sub Save_PurgeInfo()
' Save PurgeAir Source Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 4495

Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iPurge As Integer

    sFileName = FILEPATH_cfg & "purgeair.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    '  PurgeAir Definition Information
    For iPurge = 1 To MAX_PRG
    
        Write #iFileNumber, PRG_INFO(iPurge).desc
        Write #iFileNumber, PRG_INFO(iPurge).CheckSecs
        Write #iFileNumber, PRG_INFO(iPurge).UsingPrgReqHdw
        Write #iFileNumber, PRG_INFO(iPurge).UsingVacSwHdw
        Write #iFileNumber, PRG_INFO(iPurge).UsingAuxAirSol
        
        Write #iFileNumber, PRG_INFO(iPurge).UsingPosPrsPrg
        Write #iFileNumber, PRG_INFO(iPurge).UsingPrgReqAK
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next iPurge
    
    Close #iFileNumber
    
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

Sub Save_StationRecipes()
' Save Station Recipes
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1423

Dim iStation, iShift As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For iStation = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
    
            ' Update Station Recipe Records
            Criteria = "SELECT * FROM [StationRecipe] WHERE [Station] = " & iStation & "  and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Station") = iStation
                rsRecord("Shift") = iShift
            Else
              rsRecord.MoveFirst
              rsRecord.Edit
            End If
               
            ' Update Station Recipe Record
            If (Len(StationRecipe(iStation, iShift).Name) > 1) Then
                rsRecord("Name") = StationRecipe(iStation, iShift).Name
            Else
                rsRecord("Name") = "Station #" & Format(iStation, "#0") & " Shift#" & Format(iShift, "#0") & "Recipe"
            End If
            rsRecord("Number") = StationRecipe(iStation, iShift).Number
            
            rsRecord("CycleType") = StationRecipe(iStation, iShift).CycleType
            rsRecord("CycleTypeDesc") = CycleTypeDesc(StationRecipe(iStation, iShift).CycleType)
            
            rsRecord("Cycles") = StationRecipe(iStation, iShift).CyclesSave
            rsRecord("Load_Method") = StationRecipe(iStation, iShift).Load_MethodSave
            rsRecord("Load_MethodDesc") = LoadMethodDesc(StationRecipe(iStation, iShift).Load_MethodSave)
            rsRecord("UseHiRangeMFC") = StationRecipe(iStation, iShift).UseHiRangeMFC
            rsRecord("UseLoadRatePID") = StationRecipe(iStation, iShift).UseLoadRatePID
            rsRecord("NitrogenFlow") = StationRecipe(iStation, iShift).NitrogenFlowSave
            rsRecord("Load_Rate") = StationRecipe(iStation, iShift).Load_RateSave
            rsRecord("Mix_Percent") = StationRecipe(iStation, iShift).Mix_Percent
            rsRecord("WC_Mult") = StationRecipe(iStation, iShift).WC_MultSave
            rsRecord("EPAFill") = StationRecipe(iStation, iShift).EPAFill
            rsRecord("Load_Wt") = StationRecipe(iStation, iShift).Load_Wt
            rsRecord("LoadBreakthrough") = StationRecipe(iStation, iShift).LoadBreakthrough
            rsRecord("FIDmg") = StationRecipe(iStation, iShift).FIDmg
            rsRecord("Load_Time") = StationRecipe(iStation, iShift).Load_Time
            
            rsRecord("Purge_Method") = StationRecipe(iStation, iShift).Purge_Method
            rsRecord("Purge_MethodDesc") = PurgeMethodDesc(StationRecipe(iStation, iShift).Purge_Method)
            rsRecord("Purge_AuxTime") = StationRecipe(iStation, iShift).Purge_AuxTime
            rsRecord("Purge_Time") = StationRecipe(iStation, iShift).Purge_Time
            rsRecord("Purge_Flow") = StationRecipe(iStation, iShift).Purge_Flow
            rsRecord("Purge_Liters") = StationRecipe(iStation, iShift).Purge_Liters
            rsRecord("Purge_Can_Vol") = StationRecipe(iStation, iShift).Purge_Can_Vol
            rsRecord("Purge_ProfileNumber") = StationRecipe(iStation, iShift).Purge_ProfileNumber
            rsRecord("Purge_TargetMode") = StationRecipe(iStation, iShift).Purge_TargetMode
            rsRecord("Purge_TargetModeDesc") = PurgeTargetDesc(StationRecipe(iStation, iShift).Purge_TargetMode)
            rsRecord("Purge_TargetWC") = StationRecipe(iStation, iShift).Purge_TargetWC
            rsRecord("Purge_TargetWeight") = StationRecipe(iStation, iShift).Purge_TargetWeight
            rsRecord("Purge_MaxVolumes") = StationRecipe(iStation, iShift).Purge_MaxVolumes
            rsRecord("Purge_TargetPurge") = StationRecipe(iStation, iShift).Purge_TargetPurge
            rsRecord("Purge_TargetPause") = StationRecipe(iStation, iShift).Purge_TargetPause
            
            rsRecord("PurgeAuxCan") = StationRecipe(iStation, iShift).PurgeAuxCan
            rsRecord("PurgeCansInSeries") = StationRecipe(iStation, iShift).PurgeCansInSeries
            rsRecord("PurgeInOven") = StationRecipe(iStation, iShift).PurgeOven
            rsRecord("PurgeOvenSP") = StationRecipe(iStation, iShift).PurgeOvenSP
            rsRecord("UseAuxScale") = StationRecipe(iStation, iShift).UseAuxScale
            rsRecord("AuxScaleNo") = StationRecipe(iStation, iShift).AuxScaleNo
            rsRecord("PauseLeakTime") = StationRecipe(iStation, iShift).PauseLeakTime
            rsRecord("PauseLoadTime") = StationRecipe(iStation, iShift).PauseLoadTime
            rsRecord("PausePurgeTime") = StationRecipe(iStation, iShift).PausePurgeTime
            rsRecord("UsePriScale") = StationRecipe(iStation, iShift).UsePriScale
            rsRecord("PriScaleNo") = StationRecipe(iStation, iShift).PriScaleNo
            rsRecord("PauseAfterLeak") = StationRecipe(iStation, iShift).PauseAfterLeak
            rsRecord("PauseAfterLoad") = StationRecipe(iStation, iShift).PauseAfterLoad
            rsRecord("PauseAfterLoadForOper") = StationRecipe(iStation, iShift).PauseAfterLoadForOper
            rsRecord("PauseAfterPurge") = StationRecipe(iStation, iShift).PauseAfterPurge
            rsRecord("PauseAfterPurgeForOper") = StationRecipe(iStation, iShift).PauseAfterPurgeForOper
'            rsRecord("TargetConcentration") = StationRecipe(iStation, iShift).TargetConcentration
'            rsRecord("DwellTime") = StationRecipe(iStation, iShift).DwellTime
            rsRecord("LeakCheck") = StationRecipe(iStation, iShift).LeakCheck
            rsRecord("LeakPrimary") = StationRecipe(iStation, iShift).LeakPrimary
            rsRecord("LeakAux") = StationRecipe(iStation, iShift).LeakAux
'            rsRecord("UseAnalyzer") = StationRecipe(iStation, iShift).UseAnalyzer
            rsRecord("MaxLoadTime") = StationRecipe(iStation, iShift).MaxLoadTime
            
            rsRecord("IDLoad") = StationRecipe(iStation, iShift).IDLoad
            rsRecord("LoadL") = StationRecipe(iStation, iShift).LoadL
            rsRecord("LoadV") = StationRecipe(iStation, iShift).LoadV
            rsRecord("IDPurge") = StationRecipe(iStation, iShift).IDPurge
            rsRecord("PurgeL") = StationRecipe(iStation, iShift).PurgeL
            rsRecord("PurgeV") = StationRecipe(iStation, iShift).PurgeV
            rsRecord("IDVent") = StationRecipe(iStation, iShift).IDVent
            rsRecord("VentL") = StationRecipe(iStation, iShift).VentL
            rsRecord("VentV") = StationRecipe(iStation, iShift).VentV
            
            rsRecord("LiveFuel") = StationRecipe(iStation, iShift).LiveFuel
            rsRecord("LiveFuelChgAuto") = StationRecipe(iStation, iShift).LiveFuelChgAuto
            rsRecord("LiveFuelChgFreq") = StationRecipe(iStation, iShift).LiveFuelChgFreq
            rsRecord("ADF_Heater") = StationRecipe(iStation, iShift).ADF_Heater
            rsRecord("ADF_HeaterSP") = StationRecipe(iStation, iShift).ADF_HeaterSP
            
            ' start method
            rsRecord("StartMethod") = StationRecipe(iStation, iShift).StartMethod
            rsRecord("StartDelay") = StationRecipe(iStation, iShift).StartDelay
            rsRecord("StartDate") = StationRecipe(iStation, iShift).StartDate
            rsRecord("StartMethodDesc") = StartMethodDesc(StationRecipe(iStation, iShift).StartMethod)
                
            ' end method
            rsRecord("EndMethod") = StationRecipe(iStation, iShift).EndMethod
            rsRecord("EndMaximumCycles") = StationRecipe(iStation, iShift).EndMaximumCycles
            rsRecord("EndMinimumCycles") = StationRecipe(iStation, iShift).EndMinimumCycles
            rsRecord("EndConsecutiveCycles") = StationRecipe(iStation, iShift).EndConsecutiveCycles
            rsRecord("EndWeightTolerance") = StationRecipe(iStation, iShift).EndWeightTolerance
            rsRecord("UpdateCanWc") = StationRecipe(iStation, iShift).UpdateCanWc
            rsRecord("Cycles") = StationRecipe(iStation, iShift).CyclesSave
            rsRecord("EndMethodDesc") = EndMethodDesc(StationRecipe(iStation, iShift).EndMethod)
                
            ' aux outputs
            rsRecord("AuxOutputs") = StationRecipe(iStation, iShift).AuxOutputs
            rsRecord("AuxOutput1_Load") = StationRecipe(iStation, iShift).AuxOutputs_Load(1)
            rsRecord("AuxOutput2_Load") = StationRecipe(iStation, iShift).AuxOutputs_Load(2)
            rsRecord("AuxOutput3_Load") = StationRecipe(iStation, iShift).AuxOutputs_Load(3)
            rsRecord("AuxOutput4_Load") = StationRecipe(iStation, iShift).AuxOutputs_Load(4)
            rsRecord("AuxOutput1_Purge") = StationRecipe(iStation, iShift).AuxOutputs_Purge(1)
            rsRecord("AuxOutput2_Purge") = StationRecipe(iStation, iShift).AuxOutputs_Purge(2)
            rsRecord("AuxOutput3_Purge") = StationRecipe(iStation, iShift).AuxOutputs_Purge(3)
            rsRecord("AuxOutput4_Purge") = StationRecipe(iStation, iShift).AuxOutputs_Purge(4)
            
            rsRecord.Update
            rsRecord.Close
    
        Next iShift
    Next iStation
    
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

Sub Save_ScaleConfig()
' Save Scale Configuration
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1424
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iScale As Integer

    sFileName = FILEPATH_cfg & "scales.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    '  Scale configurations
    For iScale = 1 To MAX_SCALES
    
        Write #iFileNumber, Scale_Port(iScale)
        Write #iFileNumber, Scale_Type(iScale)
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next iScale
    
    Close #iFileNumber

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

Sub Save_StationInfo()
' Save Station Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1474
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer

    sFileName = FILEPATH_cfg & "stations.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    
    '  Station Definition Information
    For iStation = 1 To MAX_STN
    
        Write #iFileNumber, STN_INFO(iStation).ADF_StnNum
        Write #iFileNumber, STN_INFO(iStation).ADF_TANKTYPE
        Write #iFileNumber, STN_INFO(iStation).AspiratorNum
        Write #iFileNumber, STN_INFO(iStation).DefAuxScale
        Write #iFileNumber, STN_INFO(iStation).DefPriScale
        Write #iFileNumber, STN_INFO(iStation).desc
        Write #iFileNumber, STN_INFO(iStation).Type
        
        Write #iFileNumber, STN_INFO(iStation).ButMfcDensityMult
        Write #iFileNumber, STN_INFO(iStation).ButMfc2DensityMult
        Write #iFileNumber, STN_INFO(iStation).ADF_HEATERTYPE
        Write #iFileNumber, STN_INFO(iStation).USINGPURGEOVEN
        Write #iFileNumber, STN_INFO(iStation).Abrev
        
        Write #iFileNumber, STN_INFO(iStation).SysID
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next iStation
    
    Close #iFileNumber
    
    ' Setup Station Information
    SetupStations
    
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

Sub Save_SysDef()
' Save System definition Values
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 142
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim debugInt As Integer
Dim usingInt1 As Integer
Dim usingInt2 As Integer
Dim usingInt3 As Integer
Dim usingInt4 As Integer
Dim usingInt5 As Integer
Dim usingInt6 As Integer
Dim usingInt7 As Integer
Dim usingInt8 As Integer
Dim usingInt9 As Integer
Dim usingInt10 As Integer
Dim temp As Single

    ' combine debug Bits
    '
    '   bit 0 = UseLocalErrorHandler
    '   bit 1 = NotDebugADF
    '   bit 2 = NotDebugMMW
    '   bit 3 = NotDebugPURGE
    '
    '   bit 4 = NotDebugSCALES
    '   bit 5 = NotDebugPAS
    '   bit 6 = unused
    '   bit 7 = unused
    '
    debugInt = Bits_Pack(UseLocalErrorHandler, NotDebugADF, NotDebugMMW, NotDebugPURGE, NotDebugSCALES, NotDebugPAS, False, False)
    
    ' combine using Bits #1
    '
    '   bit 0 = USINGCOMMONTC              ' Six common thermocouples (Expert so far) 12/18/02
    '   bit 1 = USINGDOOROPEN              ' Door open switch installed
    '   bit 2 = USINGF                     ' Fht scale for temps
    '   bit 3 = USINGHIGHTEMPPAS
    '
    '   bit 4 = USINGPASLOCALCONTROL       ' Local PAS Control
    '   bit 5 = USINGLEAKCHECKEXHAUSTSOL   ' Leak Check Exhaust Sol on this system
    '   bit 6 = unused
    '   bit 7 = USINGSIMNOISE              ' add noise (max 1 %) to simulated eu values
    '
    usingInt1 = Bits_Pack(USINGCOMMONTC, USINGDOOROPEN, USINGF, USINGHIGHTEMPPAS, USINGPASLOCALCONTROL, USINGLEAKCHECKEXHAUSTSOL, False, USINGSIMNOISE)
    '
    ' combine using Bits #2
    '
    '   bit 0 = USINGBUTANEMASSLIMIT           ' One is for (was Toyota only) Limit Exceeded...Aborting
    '   bit 1 = USINGLOADTIMELIMIT       ' For Mitsu abort by operator entered time on recipe "MAX LOAD TIME"
    '   bit 2 = USINGLOADPRESSURE          ' Abort load on excessive load pressure (was Toyota only)
    '   bit 3 = USINGCUSTOMERLOWGAS        ' No remote inputs (was Toyota only)
    '
    '   bit 4 = USINGCANVENTALARM          ' Carb #1 to use (Hardware required in opto stn)
    '   bit 5 = USINGCONTAFTERLCFAIL       ' Allow Continue after a Leak Check Failure
    '   bit 6 = USINGLINEVOLUME            ' Volume compensation for line lengths (CARB first user) 5/18/03 Smitty
    '   bit 7 = USINGOOTPAUSE              ' Allow Pause of station/shift on OOT condition (Config option can still turn it off)
    '
    usingInt2 = Bits_Pack(USINGBUTANEMASSLIMIT, USINGLOADTIMELIMIT, USINGLOADPRESSURE, USINGCUSTOMERLOWGAS, USINGCANVENTALARM, USINGCONTAFTERLCFAIL, USINGLINEVOLUME, USINGOOTPAUSE)
    '
    ' combine using Bits #3
    '
    '   bit 0 = USINGREMCANLOAD            ' Task Order Manager Tasks
    '   bit 1 = LogTempRh                  ' Log Air Temperature and Humidity
    '   bit 2 = USING_ESTOP_INPUT          ' ESTOP (12vdc or Dry Contact) Input
    '   bit 3 = USINGSTNTC                 ' Use ThermoCouple (HowMany) Zero None + one for each per station
    '
    '   bit 4 = USINGMoist_RH              ' Display Moisture in grains per pound
    '   bit 5 = USINGPRESSUREPURGE         ' Yes system has config option of positive pressure purge (0 = normal)
    '   bit 6 = USINGSIMULATION            ' Run Simulation when IOScan is Off
    '   bit 7 = USINGLVol_SI               ' Using SI Units (mm & meters) for Line Vol ID & Length
    '
    usingInt3 = Bits_Pack(USINGREMCANLOAD, LogTempRh, USING_ESTOP_INPUT, USINGSTNTC, USINGMoist_RH, USINGPRESSUREPURGE, USINGSIMULATION, USINGLVol_SI)
    '
    ' combine using Bits #4
    '
    '   bit 0 = USINGHARDPIPEDSCALES       ' Scales are hard piped at 2 scales per station; stn#1 pri = 1, stn#1 aux = 2, etc.
    '   bit 1 = USINGAUXLEAKCHECK          ' Allows Leakcheck of Aux plumbing; Leakcheck Aux Only and Leakcheck Both (pri & aux)
    '   bit 2 = USINGERRORMSGBYPASS        ' Allow Error Handler to "bypass" the Error Message MsgBox
    '   bit 3 = USINGSYSTEMVACSW           ' System has a Master Vacuum Switch to be monitored
    '
    '   bit 4 = USINGPURGEDP               ' System has a DP transmitter for use during Purge
    '   bit 5 = USINGPURGESERIES           ' System is capable of Purging (Pri & Aux) Canisters in Series
    '   bit 6 = IoComOn                    ' True = Run with I/O comm, False = testing without IO
    '   bit 7 = SclComOn                   ' True = Run with Scale Port comm, False = testing without reading Scales
    '
    usingInt4 = Bits_Pack(USINGHARDPIPEDSCALES, USINGAUXLEAKCHECK, USINGERRORMSGBYPASS, USINGSYSTEMVACSW, USINGPURGEDP, USINGPURGESERIES, IoComOn, SclComOn)
    '
    ' combine using Bits #5
    '
    '   bit 0 = ChillComOn                 ' True = Run with Chiller comm, False = testing without Chiller
    '   bit 1 = USINGPURGEOVEN             ' System has one or more Purge Ovens
    '   bit 2 = USINGWATERBATH             ' System has a WaterBath Heater/Chiller for LiveFuel
    '   bit 3 = USINGFUELLEVELOOT          ' enable check of livefuel tank level
    '
    '   bit 4 = USINGDRYPURGEAIR
    '   bit 5 = USINGREMAVLFILES
    '   bit 6 = USINGREMSTSMON
    '   bit 7 = REMCHGSENABLED
    '
    usingInt5 = Bits_Pack(ChillComOn, USINGPURGEOVEN, USINGWATERBATH, USINGFUELLEVELOOT, USINGDRYPURGEAIR, USINGREMAVLFILES, USINGREMSTSMON, REMCHGSENABLED)
    '
    ' combine using Bits #6
    '
    '   bit 0 = USINGTOMCANLOAD
    '   bit 1 = Spare1
    '   bit 2 = Spare2
    '   bit 3 = Spare3
    '
    '   bit 4 = Spare4
    '   bit 5 = Spare5
    '   bit 6 = Spare6
    '   bit 7 = Spare7
    '
    usingInt6 = Bits_Pack(USINGTOMCANLOAD, Spare1, Spare2, Spare3, Spare4, Spare5, Spare6, Spare7)
    '
    ' combine using Bits #7
    '
    '   bit 0 = Spare0
    '   bit 1 = Spare1
    '   bit 2 = Spare2
    '   bit 3 = Spare3
    '
    '   bit 4 = Spare4
    '   bit 5 = Spare5
    '   bit 6 = Spare6
    '   bit 7 = Spare7
    '
    usingInt7 = Bits_Pack(Spare0, Spare1, Spare2, Spare3, Spare4, Spare5, Spare6, Spare7)
    '
    ' combine using Bits #8
    '
    '   bit 0 = Spare0
    '   bit 1 = Spare1
    '   bit 2 = Spare2
    '   bit 3 = Spare3
    '
    '   bit 4 = Spare4
    '   bit 5 = Spare5
    '   bit 6 = Spare6
    '   bit 7 = Spare7
    '
    usingInt8 = Bits_Pack(Spare0, Spare1, Spare2, Spare3, Spare4, Spare5, Spare6, Spare7)
    '
    ' combine using Bits #9
    '
    '   bit 0 = Spare0
    '   bit 1 = Spare1
    '   bit 2 = Spare2
    '   bit 3 = Spare3
    '
    '   bit 4 = Spare4
    '   bit 5 = Spare5
    '   bit 6 = Spare6
    '   bit 7 = Spare7
    '
    usingInt9 = Bits_Pack(Spare0, Spare1, Spare2, Spare3, Spare4, Spare5, Spare6, Spare7)
    '
    ' combine using Bits #10
    '
    '   bit 0 = Spare0
    '   bit 1 = Spare1
    '   bit 2 = Spare2
    '   bit 3 = Spare3
    '
    '   bit 4 = Spare4
    '   bit 5 = Spare5
    '   bit 6 = Spare6
    '   bit 7 = Spare7
    '
    usingInt10 = Bits_Pack(Spare0, Spare1, Spare2, Spare3, Spare4, Spare5, Spare6, Spare7)
    '
'#############################################################################################################
    sFileName = FILEPATH_cfg & "sysdef.@@@"
    iFileNumber = FreeFile
    
    Open sFileName For Output As #iFileNumber
'  ####   1 - 10
    Write #iFileNumber, OPTOCOM_PORT               ' opto AC24 = com2 ; sealevel = com5 ; Moxa = com6
    Write #iFileNumber, filler
    Write #iFileNumber, LoadMfcDelayTime           ' number of seconds between switching load valves & load mfc's
    Write #iFileNumber, NR_SCALES                  ' Number of Scales (includes # of remote scales)
    Write #iFileNumber, NR_SHIFT                   ' Number of shifts
    Write #iFileNumber, NR_STN                     ' Number of stations
    Write #iFileNumber, NR_DUMMYSTN                ' Number of dummy stations (i.e. io board only)
    Write #iFileNumber, NR_REMOTESCALES            ' Number of Remote Scales (i.e. not mounted in the Station Cabinet); 0 = None
    Write #iFileNumber, LocalPagControl.Type       ' 0=None, 1=Stand-Alone, 2=AkMaster, 3=AkClient)
    Write #iFileNumber, MSGDELAY                   ' DelayBox open time in ms
' ####  11 - 20
    Write #iFileNumber, MFC_Settle_Time            ' In seconds
    Write #iFileNumber, LoadEqlDelayTime           ' Time in seconds for Load Flow & Scales to reach "equalibrium"
    Write #iFileNumber, usingInt1                  ' Combined USING Bits #1
    Write #iFileNumber, usingInt2                  ' Combined USING Bits #2
    Write #iFileNumber, usingInt3                  ' Combined USING Bits #3
    Write #iFileNumber, usingInt4                  ' Combined USING Bits #4
    Write #iFileNumber, usingInt5                  ' Combined USING Bits #5
    Write #iFileNumber, usingInt6                  ' Combined USING Bits #6
    Write #iFileNumber, usingInt7                  ' Combined USING Bits #7
    Write #iFileNumber, usingInt8                  ' Combined USING Bits #8
'  ####   21 - 30
    Write #iFileNumber, usingInt9                  ' Combined USING Bits #9
    Write #iFileNumber, usingInt10                  ' Combined USING Bits #10
    Write #iFileNumber, PAGSERVERIP                ' IP address of this client's pag server
    Write #iFileNumber, TOM_2Gm_Recipe
    Write #iFileNumber, TOM_Wcm_Recipe
    Write #iFileNumber, Chiller_PORT
    Write #iFileNumber, Chiller_Timeout
    Write #iFileNumber, USING_EXT_CONTACTS          ' External alarm contacts - Remote alarm Pause
    Write #iFileNumber, DESC_EXT_CONTACTS           ' Description of external alarm
    Write #iFileNumber, USINGUPS                    ' No ups installed=0/Large ups timed=1/small down now ups=2
'  ####   31 - 40
    Write #iFileNumber, MinDataLogSeconds           ' Minimum allowed data log interval in seconds
    Write #iFileNumber, USING_AUX_OUTPUTS           ' Aux (12vdc or Dry Contact) Outputs
    Write #iFileNumber, NR_AUX_OUTPUTS              ' Number of Aux Outputs/Station
    Write #iFileNumber, DESC_AUX_OUTPUT1            ' Description of Aux Output #1
    Write #iFileNumber, DESC_AUX_OUTPUT2            ' Description of Aux Output #2
    Write #iFileNumber, DESC_AUX_OUTPUT3            ' Description of Aux Output #3
    Write #iFileNumber, DESC_AUX_OUTPUT4            ' Description of Aux Output #4
    Write #iFileNumber, MAXNOTSTABLECOUNT           ' Max Allowed Consecutive Unstable Scale reads - if exceeded then read is "declared" Stable
    Write #iFileNumber, WEIGHTQUEUESIZE
    Write #iFileNumber, MaxSheathTempForAdfDrain
'  ####   41 - 50
    Write #iFileNumber, DefScaleMax                 ' Default Max Scale Reading (Default Min = 0)
    Write #iFileNumber, GramsPerLiter               ' grams per Slpm of Butane (default = 2.40633)
    Write #iFileNumber, WB_AIO.EuMin
    Write #iFileNumber, NR_CAN                      ' Number of Canister Definitions
    Write #iFileNumber, NR_RCP                      ' Number of Recipes
    Write #iFileNumber, WB_AIO.EuMax
    Write #iFileNumber, AutoLogon                   ' 0 = No Auto Logon; 1 = user; 2 = cps; 3 = admin; 4 = aps; 4 = ApsUser
    Write #iFileNumber, debugInt                    ' Combined Debug Bits
    Write #iFileNumber, NR_JOBSEQ                   ' Number of Job Sequences
    Write #iFileNumber, MfcSpMin
'  ####   51 - 60
    Write #iFileNumber, DeadLiveFuelDensity
    Write #iFileNumber, WeakLiveFuelDensity
    Write #iFileNumber, SystemTimers(9).Interval
    Write #iFileNumber, SystemTimers(8).Interval
    Write #iFileNumber, SystemTimers(7).Interval
    Write #iFileNumber, NR_PRGAIR                   ' Number of PurgeAir Sources
    Write #iFileNumber, SystemTimers(1).Interval
    Write #iFileNumber, SystemTimers(2).Interval
    Write #iFileNumber, SystemTimers(3).Interval
    Write #iFileNumber, SystemTimers(4).Interval
'  ####   61 - 64
    Write #iFileNumber, SystemTimers(5).Interval
    Write #iFileNumber, SystemTimers(6).Interval
    Write #iFileNumber, USINGRELEASEDATE           ' program revision level
    Write #iFileNumber, CfgRevLvl                  ' Revision Level of the Cfg & SysDef Files
    Write #iFileNumber, ReportGenRevLvl            ' Revision Level required of Report Generator program
    
    Close #iFileNumber
    
    
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

Sub Load_Config()
'
' Procedure Name:   Load_config
' Created By:       Analytical Process Programmer     8/96
' Description:
' This procedure loads the user configuration data from a file located on
' the default FILEPATH in the file data\config.@@@.
' The file is loaded on a boot sequence and may also be called up during
' program operation.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1

Dim reportInt As Integer
Dim reportBits As SixteenBits
Dim reportInt2 As Integer
Dim reportBits2 As SixteenBits
Dim iFileNumber As Integer
Dim sFileName As String
Dim fillerstr As String
Dim fs, f As Object
Set fs = CreateObject("Scripting.FileSystemObject")

    sFileName = FILEPATH_cfg & "config.@@@"
    iFileNumber = FreeFile
    
    Open sFileName For Input As #iFileNumber
    Input #iFileNumber, fillerstr
    Input #iFileNumber, fillerstr
    Input #iFileNumber, reportInt
    Input #iFileNumber, SysConfig.Tol_Mix_Ratio
    Input #iFileNumber, SysConfig.LoadPressure
    Input #iFileNumber, SysConfig.Tol_Purge_Total
    Input #iFileNumber, SysConfig.Tol_Load_Total
    Input #iFileNumber, SysConfig.Tol_Nit_Flow
    Input #iFileNumber, SysConfig.Tol_Btn_Flow
    Input #iFileNumber, SysConfig.Tol_Pur_Flow
    Input #iFileNumber, SysConfig.Load_Interval
    Input #iFileNumber, SysConfig.Purge_Interval
    Input #iFileNumber, SysConfig.Tol_Temp
    Input #iFileNumber, SysConfig.Tol_Moisture
    Input #iFileNumber, SysConfig.Temp_Target
    Input #iFileNumber, SysConfig.Moisture_Target
    Input #iFileNumber, SysConfig.Next_File
    Input #iFileNumber, SysConfig.AutoLogon
    Input #iFileNumber, SysConfig.Heading
    Input #iFileNumber, SysConfig.Heading2
    Input #iFileNumber, SysConfig.ReportBackup_Active
    Input #iFileNumber, SysConfig.ReportBackup_Path
    Input #iFileNumber, SysConfig.JobRecs
    Input #iFileNumber, SysConfig.LCMinDelay
    Input #iFileNumber, SysConfig.LCSetPoint
    Input #iFileNumber, SysConfig.LCTime
    Input #iFileNumber, SysConfig.PressureDecay
    Input #iFileNumber, SysConfig.NitrogenPurgeTime
    Input #iFileNumber, SysConfig.DoorOpenDelay
    Input #iFileNumber, SysConfig.UPSOpenDelay
    Input #iFileNumber, SysConfig.OOTtimeDelay
    Input #iFileNumber, SysConfig.PosPressPurge
    Input #iFileNumber, SysConfig.Tol_FuelTemp
    Input #iFileNumber, SysConfig.Default_Interval
    
    Input #iFileNumber, SysConfig.CanVent_Delay_Max
    Input #iFileNumber, SysConfig.LoLim_Load_Flow             ' Low Limit for Tolerance Checking in %
    Input #iFileNumber, SysConfig.LoLim_Purge_Flow            ' Low Limit for Tolerance Checking in %
    Input #iFileNumber, SysConfig.LoadTotal_Interval
    Input #iFileNumber, SysConfig.PurgeTotal_Interval
    
    Input #iFileNumber, SysConfig.RemStatus_Interval
    Input #iFileNumber, fillerstr
    Input #iFileNumber, fillerstr
    Input #iFileNumber, fillerstr
    Input #iFileNumber, SysConfig.WaterBathControl
    
    Input #iFileNumber, SysConfig.Tol_Lfv_Flow
    Input #iFileNumber, SysConfig.PurgeDP_HiLimit
    Input #iFileNumber, SysConfig.EventRecs
    Input #iFileNumber, SysConfig.LeakCheckFailResponse
    Input #iFileNumber, SysConfig.LeakCheck_Interval
    
    Input #iFileNumber, SysConfig.DbFileBackup_Active
    Input #iFileNumber, SysConfig.DbFileBackup_Path
    Input #iFileNumber, SysConfig.ReportFileName1stPart
    Input #iFileNumber, SysConfig.ReportFileName2ndPart
    Input #iFileNumber, SysConfig.ReportFileName3rdPart
    
    Input #iFileNumber, SysConfig.Tol_ORVRNit_Flow
    Input #iFileNumber, SysConfig.Tol_ORVRBtn_Flow
    Input #iFileNumber, SysConfig.LoadSettleTime
    Input #iFileNumber, SysConfig.PurgeSettleTime
    Input #iFileNumber, SysConfig.AutoLogonUser
    
    Input #iFileNumber, SysConfig.ButaneMassLimit
    Input #iFileNumber, SysConfig.LoadTimeLimit
    Input #iFileNumber, SysConfig.TempRhLogInterval
    Input #iFileNumber, SysConfig.TempRhLogVerbose
    Input #iFileNumber, reportInt2
    
    Input #iFileNumber, SysConfig.BtnFlowResp
    Input #iFileNumber, SysConfig.NitFlowResp
    Input #iFileNumber, SysConfig.PurFlowResp
    Input #iFileNumber, SysConfig.AirMoistResp
    Input #iFileNumber, SysConfig.AirTempResp
    
    Input #iFileNumber, SysConfig.FuelTempResp
    Input #iFileNumber, SysConfig.CanVentResp
    Input #iFileNumber, SysConfig.LoadRateResp
    Input #iFileNumber, SysConfig.PurgeDpResp
    Input #iFileNumber, SysConfig.FuelLevelResp
    
    Input #iFileNumber, SysConfig.PurgeOvenBand
    Input #iFileNumber, SysConfig.DryAirPurge
    Input #iFileNumber, SysConfig.PurgeOvenResp
    Input #iFileNumber, SysConfig.WaterBathResp
    Input #iFileNumber, SysConfig.Tol_PurgeOvenTemp
    
    Input #iFileNumber, SysConfig.Tol_WaterBathTemp
    Input #iFileNumber, fillerstr
    Input #iFileNumber, fillerstr
    Input #iFileNumber, fillerstr
    Input #iFileNumber, fillerstr
    
    Close #iFileNumber
    
    ' separate report Bits
    '
    '   bit 0 = TextEotReporting
    '   bit 1 = TextEotSummary
    '   bit 2 = TextEotSummary_AutoPrint
    '   bit 3 = TextEotDetail
    '
    '   bit 4 = TextGenReporting
    '   bit 5 = TextGenSummary
    '   bit 6 = TextGenDetail
    '
    '   bit 7 = XlsEotReporting
    '   bit 8 = XlsEotSummary
    '   bit 9 = XlsEotDetail
    '
    '   bit 10 = XlsEotReporting
    '   bit 11 = XlsEotSummary
    '   bit 12 = XlsEotDetail
    '
    '   bit 13 = unused
    '   bit 14 = unused
    '
    reportBits = Bits_UnPack(reportInt)
    With SysConfig.RptConfig
        .TextEotReporting = reportBits.B00
        .TextEotSummary = reportBits.B01
        .TextEotSummary_AutoPrint = reportBits.B02
        .TextEotDetail = reportBits.B03
        .TextGenReporting = reportBits.B04
        .TextGenSummary = reportBits.B05
        .TextGenDetail = reportBits.B06
        .XlsEotReporting = reportBits.B07
        .XlsEotSummary = reportBits.B08
        .XlsEotDetail = reportBits.B09
        .XlsGenReporting = reportBits.B10
        .XlsGenSummary = reportBits.B11
        .XlsGenDetail = reportBits.B12
    End With
    
    ' separate report Bits2
    '
    '   bit 0 = CsvEotReporting
    '   bit 1 = CsvEotSummary
    '   bit 2 = CsvEotDetail
    '   bit 3 = unused
    '
    '   bit 4 = CsvGenReporting
    '   bit 5 = CsvGenSummary
    '   bit 6 = CsvGenDetail
    '
    '   bit 7 = unused
    '   bit 8 = unused
    '   bit 9 = unused
    '
    '   bit 10 = unused
    '   bit 11 = unused
    '   bit 12 = unused
    '
    '   bit 13 = unused
    '   bit 14 = unused
    '
    reportBits2 = Bits_UnPack(reportInt2)
    With SysConfig.RptConfig
        .CsvEotReporting = reportBits2.B00
        .CsvEotSummary = reportBits2.B01
        .CsvEotDetail = reportBits2.B02
        .CsvGenReporting = reportBits2.B04
        .CsvGenSummary = reportBits2.B05
        .CsvGenDetail = reportBits2.B06
    End With
    
    '
    ' Set Leak Check Totalize Interval (always = 1 sec)
    '
    SysConfig.LeakTotal_Interval = 1
    
    
    ' Check Sysdef controlled options
    SysConfig.BtnFlowResp = IIf(USINGOOTPAUSE, SysConfig.BtnFlowResp, ootrspContinue)
    SysConfig.NitFlowResp = IIf(USINGOOTPAUSE, SysConfig.NitFlowResp, ootrspContinue)
    SysConfig.FuelTempResp = IIf(USINGOOTPAUSE, SysConfig.FuelTempResp, ootrspContinue)
    SysConfig.PurFlowResp = IIf(USINGOOTPAUSE, SysConfig.PurFlowResp, ootrspContinue)
    SysConfig.AirMoistResp = IIf(USINGOOTPAUSE, SysConfig.AirMoistResp, ootrspContinue)
    SysConfig.AirTempResp = IIf(USINGOOTPAUSE, SysConfig.AirTempResp, ootrspContinue)
    SysConfig.CanVentResp = IIf(USINGOOTPAUSE, SysConfig.CanVentResp, ootrspContinue)
    SysConfig.LoadRateResp = IIf(USINGOOTPAUSE, SysConfig.LoadRateResp, ootrspContinue)
    
    ' Check Totalize Intervals for valid values
    ' note: Totalize Intervals were added as part of CfgRevLvl=2
    If ((SysConfig.LoadTotal_Interval < 0.1) Or (SysConfig.LoadTotal_Interval > SysConfig.Load_Interval)) Then SysConfig.LoadTotal_Interval = 1
    If ((SysConfig.PurgeTotal_Interval < 0.1) Or (SysConfig.PurgeTotal_Interval > SysConfig.Purge_Interval)) Then SysConfig.PurgeTotal_Interval = 1
    
    ' Check Vapor Carrier Flow Tolerance
    ' note: Separate Vapor Carrier Flow Tolerance was added as part of CfgRevLvl=3
    If CfgRevLvl < 3 Then SysConfig.Tol_Lfv_Flow = SysConfig.Tol_Nit_Flow   ' default Vapor Carrier Flow Tolerance
    
    ' Check SysConfig.EventRecs for valid values
    ' note: Max Event Log Entries were added as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then SysConfig.EventRecs = 100                         ' default Max Event Log Entries
    
    ' Check SysConfig.DbFile Backup for valid values
    ' note: DbFile Backup Entries were added as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then SysConfig.DbFileBackup_Active = False             ' default DbFile Backup values
    If CfgRevLvl < 5 Then SysConfig.DbFileBackup_Path = ""                  ' default DbFile Backup values
    
    ' Check SysConfig.LeakCheckFailResponse & SysConfig.LeakCheck_Interval for valid values
    ' note: Leak Check Failure Response was added as part of CfgRevLvl=5
    ' note: Leak Check Interval was added as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then SysConfig.LeakCheckFailResponse = 0               ' default = only STOP on Leak Check Failure
    If CfgRevLvl < 5 Then SysConfig.LeakCheck_Interval = 1                  ' default = 1 second
    
    ' Check Report File Naming
    ' note: User changeable Report File Names were added as part of CfgRevLvl=5
    If CfgRevLvl < 5 Then SysConfig.ReportFileName1stPart = 1   ' default = <Job #>_
    If CfgRevLvl < 5 Then SysConfig.ReportFileName2ndPart = 0   ' default = nothing
    If CfgRevLvl < 5 Then SysConfig.ReportFileName3rdPart = 0   ' default = nothing
    If SysConfig.ReportFileName1stPart = 0 Then SysConfig.ReportFileName1stPart = 1     ' default = <Job #>_
    
    If SysConfig.TempRhLogInterval = 0 Then SysConfig.TempRhLogInterval = 1             ' default = 1 minute
    
    If Not IntroDone Then
        frmAbout.UpdateMsg "Configuration File Loaded" & vbCrLf
        Delay_Box "", INTRODELAY, msgNOSHOW
    Else
        Delay_Box "Configuration File Loaded", MSGDELAY, msgSHOW
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

Sub Load_DigitalFuncDef()
' Load Digital Functions Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1493

Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim inct As Integer
Dim inct2 As Integer
Dim inct3 As Integer
Dim iAddr As Integer
Dim iChan As Integer
Dim iUseInverse As Boolean

    sFileName = FILEPATH_cfg & "funcdefd.@@@"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    '  COMMON
    For inct2 = 0 To MAX_DIG_COM
    
        Input #iFileNumber, iUseInverse
        Input #iFileNumber, iAddr
        Input #iFileNumber, iChan
        
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
        If (Com_DIO(inct2).addr > OptoMaxNodeNum) Then
            inct3 = Com_DIO(inct2).addr
            Com_DIO(inct2).addr = 0
            Com_DIO(inct2).chan = 0
            Com_DIO(inct2).UseInverse = False
        Else
            Com_DIO(inct2).UseInverse = iUseInverse
            Com_DIO(inct2).addr = iAddr
            Com_DIO(inct2).chan = iChan
        End If
            
    Next inct2
    
    '  FID
    For inct2 = 0 To MAX_DIG_FID
    
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        
    Next inct2
    
    '  Stations
    For inct = 1 To MAX_STN
        For inct2 = 0 To MAX_DIG_STN
    
            Input #iFileNumber, iUseInverse
            Input #iFileNumber, iAddr
            Input #iFileNumber, iChan
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            Input #iFileNumber, filler
            
            If (Stn_DIO(inct, inct2).addr > OptoMaxNodeNum) Then
                inct3 = Stn_DIO(inct, inct2).addr
                Stn_DIO(inct, inct2).addr = 0
                Stn_DIO(inct, inct2).chan = 0
                Stn_DIO(inct, inct2).UseInverse = False
            Else
                Stn_DIO(inct, inct2).UseInverse = iUseInverse
                Stn_DIO(inct, inct2).addr = iAddr
                Stn_DIO(inct, inct2).chan = iChan
            End If
            
            If ((Stn_DIO(inct, inct2).addr = 4) And (Stn_DIO(inct, inct2).chan = 15)) Then
                filler = filler
            End If
            
        Next inct2
        
    Next inct
    
    If CfgRevLvl > 0 Then
        '  PurgeAir Sources
        For inct = 1 To MAX_PRG
            For inct2 = 0 To MAX_DIG_PRG
        
                Input #iFileNumber, iUseInverse
                Input #iFileNumber, iAddr
                Input #iFileNumber, iChan
                
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                Input #iFileNumber, filler
                
                If (Prg_DIO(inct, inct2).addr > OptoMaxNodeNum) Then
                    inct3 = Prg_DIO(inct, inct2).addr
                    Prg_DIO(inct, inct2).addr = 0
                    Prg_DIO(inct, inct2).chan = 0
                    Prg_DIO(inct, inct2).UseInverse = False
                Else
                    Prg_DIO(inct, inct2).UseInverse = iUseInverse
                    Prg_DIO(inct, inct2).addr = iAddr
                    Prg_DIO(inct, inct2).chan = iChan
                End If
            Next inct2
            
        Next inct
    End If
    Close #iFileNumber
    
    
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

Sub Save_StationCanisters()
' Save Station Canister Recipes
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1421

Dim icnt1, icnt2 As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For icnt1 = 1 To LAST_STN
        For icnt2 = 1 To NR_SHIFT
    
            ' Update Station Canister Information Records
            Criteria = "SELECT * FROM [StationCanister] WHERE [Station] = " & icnt1 & "  and [Shift] = " & icnt2 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Station") = icnt1
                rsRecord("Shift") = icnt2
            Else
              rsRecord.MoveFirst
              rsRecord.Edit
            End If
               
            rsRecord("Number") = StationCanister(icnt1, icnt2).Number
            If (Len(StationCanister(icnt1, icnt2).Description) > 1) Then
                rsRecord("Description") = StationCanister(icnt1, icnt2).Description
            Else
                rsRecord("Description") = "Station #" & Format(icnt1, "#0") & " Shift#" & Format(icnt2, "#0") & "Canister"
            End If
            rsRecord("WorkingCapacity") = StationCanister(icnt1, icnt2).WorkingCapacity
            rsRecord("WCVolume") = StationCanister(icnt1, icnt2).WorkingVolume
            
            rsRecord.Update
            rsRecord.Close
    
        Next icnt2
    Next icnt1
    
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

Sub Save_StationProfiles()
' Save Station Purge Profiles
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 14213

Dim iStn As Integer
Dim iShift As Integer
Dim iStep As Integer
Dim dbDbase As Database
Dim rsProfile  As Recordset
Dim rsSteps  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For iStn = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
    
            ' Update Station Profile Information Records
            Criteria = "SELECT * FROM [StationProfiles] WHERE [Station] = " & iStn & "  and [Shift] = " & iShift & " "
            Set rsProfile = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsProfile.BOF Then
                rsProfile.AddNew
                rsProfile("Station") = iStn
                rsProfile("Shift") = iShift
            Else
              rsProfile.MoveFirst
              rsProfile.Edit
            End If
               
            ' Update Station PurgeProfile Record
            If (Len(StationProfile(iStn, iShift).Description) > 1) Then
                rsProfile("Description") = StationProfile(iStn, iShift).Description
            Else
                rsProfile("Description") = "Station #" & Format(iStn, "#0") & " Shift#" & Format(iShift, "#0") & "Profile"
            End If
            rsProfile("TotalDuration") = StationProfile(iStn, iShift).Duration
            rsProfile("Steps") = StationProfile(iStn, iShift).EndStep
            rsProfile("ProjectedLiters") = StationProfile(iStn, iShift).ProjectedLiters
            rsProfile("ProjectedVolumes") = StationProfile(iStn, iShift).ProjectedVolumes
            rsProfile.Update
            rsProfile.Close
        
            ' Save Station PurgeProfile Steps
            Criteria = "SELECT * FROM [StationProfileSteps] WHERE [Station] = " & iStn & "  and [Shift] = " & iShift & " ORDER BY [StepNumber] ASC"
            Set rsSteps = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            ' first remove existing steps
            If rsSteps.BOF Then
                ' nothing to do; no steps for this profile exist in db
            Else
                rsSteps.MoveLast
                If (Not rsSteps.BOF) Then
                    While Not rsSteps.BOF
                        rsSteps.Delete
                        rsSteps.MovePrevious
                    Wend
                End If
            End If
            rsSteps.Close
            
            ' Update Station PurgeProfile Steps
            Set rsSteps = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            For iStep = 1 To MAX_PROFILESTEPS
                If StationProfile(iStn, iShift).StepType(iStep) <> NOSTEP Then
                    rsSteps.AddNew
                    rsSteps("Station") = iStn
                    rsSteps("Shift") = iShift
                    rsSteps("StepNumber") = iStep
                    rsSteps("Duration") = StationProfile(iStn, iShift).StepDuration(iStep)
                    rsSteps("InitialSP") = StationProfile(iStn, iShift).StepStartSetpoint(iStep)
                    rsSteps("StepType") = StationProfile(iStn, iShift).StepType(iStep)
                    rsSteps("StepTypeDesc") = PurgeProfileStepDesc(StationProfile(iStn, iShift).StepType(iStep))
                    rsSteps.Update
                End If
            Next iStep
            rsSteps.Close
    
        Next iShift
    Next iStn
    
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

Sub Save_StationSequences()
' Save Station Sequence Recipes
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1921

Dim iStation As Integer
Dim iShift As Integer
Dim iCourse As Integer
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    For iStation = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
    
            ' Update Station Sequence Information Records
            Criteria = "SELECT * FROM [StationSequence] WHERE [Station] = " & iStation & "  and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("Station") = iStation
                rsRecord("Shift") = iShift
            Else
              rsRecord.MoveFirst
              rsRecord.Edit
            End If
               
    '        rsRecord("Number") = StationSequence(iStation, iShift).Number
            rsRecord("Number") = 0
            If (Len(StationSequence(iStation, iShift).Description) > 1) Then
                rsRecord("Description") = StationSequence(iStation, iShift).Description
            Else
                rsRecord("Description") = "Station #" & Format(iStation, "#0") & " Shift#" & Format(iShift, "#0") & "Sequence"
            End If
            rsRecord("Courses") = StationSequence(iStation, iShift).NumCourses
            rsRecord("PriScale") = StationSequence(iStation, iShift).PriScaleNo
            rsRecord("AuxScale") = StationSequence(iStation, iShift).AuxScaleNo
            rsRecord("IDLoad") = StationSequence(iStation, iShift).IDLoad
            rsRecord("IDPurge") = StationSequence(iStation, iShift).IDPurge
            rsRecord("IDVent") = StationSequence(iStation, iShift).IDVent
            rsRecord("LoadL") = StationSequence(iStation, iShift).LoadL
            rsRecord("LoadV") = StationSequence(iStation, iShift).LoadV
            rsRecord("PurgeL") = StationSequence(iStation, iShift).PurgeL
            rsRecord("PurgeV") = StationSequence(iStation, iShift).PurgeV
            rsRecord("VentL") = StationSequence(iStation, iShift).VentL
            rsRecord("VentV") = StationSequence(iStation, iShift).VentV
            rsRecord("Validated") = StationSequence(iStation, iShift).Validated
            rsRecord("EstSeqDuration") = StationSequence(iStation, iShift).EstSeqDuration
            rsRecord("EstSeqDurDesc") = StationSequence(iStation, iShift).EstSeqDurDesc
            
            rsRecord.Update
            rsRecord.Close
    
            ' Delete Existing Station Sequence Course Information Records
            Criteria = "SELECT * FROM [StationSequenceCourses] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            If Not rsRecord.BOF Then
                rsRecord.MoveLast
                Do While Not rsRecord.BOF
                    rsRecord.Delete
                    rsRecord.MoveLast
                Loop
            End If
                       
            ' Update Station Sequence Course Information Records
            Criteria = "SELECT * FROM [StationSequenceCourses] WHERE [Station] = " & iStation & " and [Shift] = " & iShift & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            If rsRecord.BOF Then
                iCourse = 1
                While (iCourse <= StationSequence(iStation, iShift).NumCourses)
                    rsRecord.AddNew
                    rsRecord("Station") = iStation
                    rsRecord("Shift") = iShift
                    rsRecord("CourseNumber") = StationSequence(iStation, iShift).CourseData(iCourse).CourseNumber
                    rsRecord("Type") = StationSequence(iStation, iShift).CourseData(iCourse).Type
                    rsRecord("PauseDuration") = StationSequence(iStation, iShift).CourseData(iCourse).PauseDuration
                    rsRecord("RecipeNumber") = StationSequence(iStation, iShift).CourseData(iCourse).RecipeNumber
                    rsRecord("Cycles") = StationSequence(iStation, iShift).CourseData(iCourse).Cycles
                    rsRecord("LoadRate") = StationSequence(iStation, iShift).CourseData(iCourse).LoadRate
                    rsRecord("PurgeRate") = StationSequence(iStation, iShift).CourseData(iCourse).PurgeRate
                    rsRecord("MsgText") = StationSequence(iStation, iShift).CourseData(iCourse).MsgText
                    rsRecord("EstCourseDuration") = StationSequence(iStation, iShift).CourseData(iCourse).EstCourseDuration
                    rsRecord.Update
                    iCourse = iCourse + 1
                Wend
            End If
                       
            rsRecord.Close

        Next iShift
    Next iStation
    
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

Sub Save_AnalogFuncDef()
' Save Station Analog Functions Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1492
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim inct As Integer
Dim inct2 As Integer

    sFileName = FILEPATH_cfg & "funcdefa.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    '  COMMON
    '  Analog Functions
    For inct2 = 0 To MAX_ANA_COM
    
        Write #iFileNumber, Com_AIO(inct2).EuMax
        Write #iFileNumber, Com_AIO(inct2).EuMin
        Write #iFileNumber, Com_AIO(inct2).VdcMax
        Write #iFileNumber, Com_AIO(inct2).VdcMin
        Write #iFileNumber, Com_AIO(inct2).addr
        Write #iFileNumber, Com_AIO(inct2).chan
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next inct2
    
    '  FID
    '  Analog Functionss
    For inct2 = 0 To MAX_ANA_FID
    
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next inct2
    
    '  Stations
    For inct = 1 To MAX_STN
        '  Analog Functions
        For inct2 = 0 To MAX_ANA_STN
    
            Write #iFileNumber, Stn_AIO(inct, inct2).EuMax
            Write #iFileNumber, Stn_AIO(inct, inct2).EuMin
            Write #iFileNumber, Stn_AIO(inct, inct2).VdcMax
            Write #iFileNumber, Stn_AIO(inct, inct2).VdcMin
            Write #iFileNumber, Stn_AIO(inct, inct2).addr
            Write #iFileNumber, Stn_AIO(inct, inct2).chan
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
        Next inct2
        
    Next inct
    
    '  PurgeAir Sources
    For inct = 1 To MAX_PRG
        '  Analog Functionss
        For inct2 = 0 To MAX_ANA_PRG
    
            Write #iFileNumber, Prg_AIO(inct, inct2).EuMax
            Write #iFileNumber, Prg_AIO(inct, inct2).EuMin
            Write #iFileNumber, Prg_AIO(inct, inct2).VdcMax
            Write #iFileNumber, Prg_AIO(inct, inct2).VdcMin
            Write #iFileNumber, Prg_AIO(inct, inct2).addr
            Write #iFileNumber, Prg_AIO(inct, inct2).chan
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
        Next inct2
        
    Next inct
    
    Close #iFileNumber
    
    
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

Sub Save_Config()
'
' Procedure Name:   Save_config
' Created By:       Analytical Process Programmer     8/96
' Description:
' This procedure saves the user configuration data to a file located on
' the default FILEPATH to the file data\config.@@@.
' This file is used by the Save button on the config screen and also
' by the check_stations routine to update the next file number
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 3

Dim iFileNumber As Integer
Dim sFileName As String
Dim fillerstr As String
Dim reportInt As Integer
Dim reportInt2 As Integer

    fillerstr = "123456789012345678901234567890"
    
    ' combine report Bits
    '
    '   bit 0 = TextEotReporting
    '   bit 1 = TextEotSummary
    '   bit 2 = TextEotSummary_AutoPrint
    '   bit 3 = TextEotDetail
    '
    '   bit 4 = TextGenReporting
    '   bit 5 = TextGenSummary
    '   bit 6 = TextGenDetail
    '
    '   bit 7 = XlsEotReporting
    '   bit 8 = XlsEotSummary
    '   bit 9 = XlsEotDetail
    '
    '   bit 10 = XlsGenReporting
    '   bit 11 = XlsGenSummary
    '   bit 12 = XlsGenDetail
    '
    '   bit 13 = unused
    '   bit 14 = unused
    '

    With SysConfig.RptConfig
        reportInt = Bits_Pack(.TextEotReporting, .TextEotSummary, .TextEotSummary_AutoPrint, .TextEotDetail, .TextGenReporting, .TextGenSummary, .TextGenDetail, .XlsEotReporting, .XlsEotSummary, .XlsEotDetail, .XlsGenReporting, .XlsGenSummary, .XlsGenDetail, False, False)
    End With
    
    ' combine report Bits2
    '
    '   bit 0 = CsvEotReporting
    '   bit 1 = CsvEotSummary
    '   bit 2 = CsvEotDetail
    '   bit 3 = unused
    '
    '   bit 4 = TextGenReporting
    '   bit 5 = TextGenSummary
    '   bit 6 = TextGenDetail
    '
    '   bit 7 = unused
    '   bit 8 = unused
    '   bit 9 = unused
    '
    '   bit 10 = unused
    '   bit 11 = unused
    '   bit 12 = unused
    '
    '   bit 13 = unused
    '   bit 14 = unused
    '   bit 15 = unused
    '

    With SysConfig.RptConfig
        reportInt2 = Bits_Pack(.CsvEotReporting, .CsvEotSummary, .CsvEotDetail, False, .CsvGenReporting, .CsvGenSummary, .CsvGenDetail, False, False, False, False, False, False, False, False)
    End With
    
    ' This file does not update the configuration variables, only saves them
    sFileName = FILEPATH_cfg & "config.@@@"
    iFileNumber = FreeFile
    
    Open sFileName For Output As #iFileNumber
    Write #iFileNumber, fillerstr
    Write #iFileNumber, fillerstr
    Write #iFileNumber, reportInt
    Write #iFileNumber, SysConfig.Tol_Mix_Ratio
    Write #iFileNumber, SysConfig.LoadPressure
    Write #iFileNumber, SysConfig.Tol_Purge_Total
    Write #iFileNumber, SysConfig.Tol_Load_Total
    Write #iFileNumber, SysConfig.Tol_Nit_Flow
    Write #iFileNumber, SysConfig.Tol_Btn_Flow
    Write #iFileNumber, SysConfig.Tol_Pur_Flow
    Write #iFileNumber, SysConfig.Load_Interval
    Write #iFileNumber, SysConfig.Purge_Interval
    Write #iFileNumber, SysConfig.Tol_Temp
    Write #iFileNumber, SysConfig.Tol_Moisture
    Write #iFileNumber, SysConfig.Temp_Target
    Write #iFileNumber, SysConfig.Moisture_Target
    Write #iFileNumber, SysConfig.Next_File
    Write #iFileNumber, SysConfig.AutoLogon
    Write #iFileNumber, SysConfig.Heading
    Write #iFileNumber, SysConfig.Heading2
    Write #iFileNumber, SysConfig.ReportBackup_Active
    Write #iFileNumber, SysConfig.ReportBackup_Path
    Write #iFileNumber, SysConfig.JobRecs
    Write #iFileNumber, SysConfig.LCMinDelay
    Write #iFileNumber, SysConfig.LCSetPoint
    Write #iFileNumber, SysConfig.LCTime
    Write #iFileNumber, SysConfig.PressureDecay
    Write #iFileNumber, SysConfig.NitrogenPurgeTime
    Write #iFileNumber, SysConfig.DoorOpenDelay
    Write #iFileNumber, SysConfig.UPSOpenDelay
    Write #iFileNumber, SysConfig.OOTtimeDelay
    Write #iFileNumber, SysConfig.PosPressPurge
    Write #iFileNumber, SysConfig.Tol_FuelTemp
    Write #iFileNumber, SysConfig.Default_Interval
    
    Write #iFileNumber, SysConfig.CanVent_Delay_Max
    Write #iFileNumber, SysConfig.LoLim_Load_Flow             ' Low Limit for Tolerance Checking in %
    Write #iFileNumber, SysConfig.LoLim_Purge_Flow            ' Low Limit for Tolerance Checking in %
    Write #iFileNumber, SysConfig.LoadTotal_Interval
    Write #iFileNumber, SysConfig.PurgeTotal_Interval
    
    Write #iFileNumber, SysConfig.RemStatus_Interval
    Write #iFileNumber, fillerstr
    Write #iFileNumber, fillerstr
    Write #iFileNumber, fillerstr
    Write #iFileNumber, SysConfig.WaterBathControl
    
    Write #iFileNumber, SysConfig.Tol_Lfv_Flow
    Write #iFileNumber, SysConfig.PurgeDP_HiLimit
    Write #iFileNumber, SysConfig.EventRecs
    Write #iFileNumber, SysConfig.LeakCheckFailResponse
    Write #iFileNumber, SysConfig.LeakCheck_Interval
    
    Write #iFileNumber, SysConfig.ReportBackup_Active
    Write #iFileNumber, SysConfig.ReportBackup_Path
    Write #iFileNumber, SysConfig.ReportFileName1stPart
    Write #iFileNumber, SysConfig.ReportFileName2ndPart
    Write #iFileNumber, SysConfig.ReportFileName3rdPart
    
    Write #iFileNumber, SysConfig.Tol_ORVRNit_Flow
    Write #iFileNumber, SysConfig.Tol_ORVRBtn_Flow
    Write #iFileNumber, SysConfig.LoadSettleTime
    Write #iFileNumber, SysConfig.PurgeSettleTime
    Write #iFileNumber, SysConfig.AutoLogonUser
    
    Write #iFileNumber, SysConfig.ButaneMassLimit
    Write #iFileNumber, SysConfig.LoadTimeLimit
    Write #iFileNumber, SysConfig.TempRhLogInterval
    Write #iFileNumber, SysConfig.TempRhLogVerbose
    Write #iFileNumber, reportInt2
    
    Write #iFileNumber, SysConfig.BtnFlowResp
    Write #iFileNumber, SysConfig.NitFlowResp
    Write #iFileNumber, SysConfig.PurFlowResp
    Write #iFileNumber, SysConfig.AirMoistResp
    Write #iFileNumber, SysConfig.AirTempResp
    
    Write #iFileNumber, SysConfig.FuelTempResp
    Write #iFileNumber, SysConfig.CanVentResp
    Write #iFileNumber, SysConfig.LoadRateResp
    Write #iFileNumber, SysConfig.PurgeDpResp
    Write #iFileNumber, SysConfig.FuelLevelResp
    
    Write #iFileNumber, SysConfig.PurgeOvenBand
    Write #iFileNumber, SysConfig.DryAirPurge
    Write #iFileNumber, SysConfig.PurgeOvenResp
    Write #iFileNumber, SysConfig.WaterBathResp
    Write #iFileNumber, SysConfig.Tol_PurgeOvenTemp
    
    Write #iFileNumber, SysConfig.Tol_WaterBathTemp
    Write #iFileNumber, fillerstr
    Write #iFileNumber, fillerstr
    Write #iFileNumber, fillerstr
    Write #iFileNumber, fillerstr
    
    Close #iFileNumber
    
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

Sub Save_DigitalFuncDef()
' Save Digital Functions Definition Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 1491

Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim inct As Integer
Dim inct2 As Integer

    sFileName = FILEPATH_cfg & "funcdefd.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    '  COMMON
    '  Digital Functions
    For inct2 = 0 To MAX_DIG_COM
    
        Write #iFileNumber, Com_DIO(inct2).UseInverse
        Write #iFileNumber, Com_DIO(inct2).addr
        Write #iFileNumber, Com_DIO(inct2).chan
        
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next inct2
    
    '  FID
    '  Digital Functions
    For inct2 = 0 To MAX_DIG_FID
    
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        
    Next inct2
    
    '  Stations
    For inct = 1 To MAX_STN
        '  Digital Functions
        For inct2 = 0 To MAX_DIG_STN
    
            Write #iFileNumber, Stn_DIO(inct, inct2).UseInverse
            Write #iFileNumber, Stn_DIO(inct, inct2).addr
            Write #iFileNumber, Stn_DIO(inct, inct2).chan
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
        Next inct2
        
    Next inct
    
    '  PurgeAir Sources
    For inct = 1 To MAX_PRG
        '  Digital Functions
        For inct2 = 0 To MAX_DIG_PRG
    
            Write #iFileNumber, Prg_DIO(inct, inct2).UseInverse
            Write #iFileNumber, Prg_DIO(inct, inct2).addr
            Write #iFileNumber, Prg_DIO(inct, inct2).chan
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            Write #iFileNumber, filler
            
        Next inct2
        
    Next inct
    
    Close #iFileNumber
    
    
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

Public Function ProfileDurationDescription(ByVal ppd As Single) As String
' Routine Name: ProfileDurationDescription
' Created by:   Brunrose
' Function:
' This routine converts a Purge Profile Duration (in minutes) into a string description.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 9093
Dim Hrs As Integer
Dim remainDur As Single
Dim sDesc As String

    Hrs = 0
    remainDur = ppd
    Do While remainDur >= 60
        remainDur = remainDur - 60
        Hrs = Hrs + 1
    Loop
    
    sDesc = ""
    If Hrs > 0 Then
        ' duration >= 1 hour
        sDesc = sDesc & Format(Hrs, "#,##0") & ":"
        sDesc = sDesc & Format(Int(remainDur), "00") & ":"
        sDesc = sDesc & Format((60 * (remainDur - Int(remainDur))), "00")
    ElseIf ppd > 1.5 Then
        ' duration > 90 seconds
        sDesc = sDesc & Format(Int(remainDur), "#0") & ":"
        sDesc = sDesc & Format((60 * (remainDur - Int(remainDur))), "00")
    Else
        ' duration <= 90 seconds
        sDesc = sDesc & Format((60 * remainDur), "##0") & " sec"
    End If

    ' Result
    ProfileDurationDescription = sDesc

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

Public Sub UpdateStnRcpDsc(ByVal iStn As Integer, ByVal iShift As Integer)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 3210
Dim sStart, sLeak, sCycles, sPurge, sLoad, sTxt As String

        ' Start - add StartMethod phase description
        Select Case StationRecipe(iStn, iShift).StartMethod
            Case STARTNOW
                sStart = ""
            Case STARTDELAYED
                sStart = "Delay Start by " & Format(StationRecipe(iStn, iShift).StartDelay, "##0")
                sStart = sStart & StartTypeDesc2(STARTDELAYED)
            Case STARTATDATE
                sStart = StartTypeDesc2(STARTATDATE)
                sStart = sStart & Format(StationRecipe(iStn, iShift).StartDate, "D MMM, YYYY   h:mm")
            Case Else
                sStart = "Do Not Start"
        End Select
        ' Cycles - add Number of Cycles phase description
        Select Case StationRecipe(iStn, iShift).EndMethod
            Case ENDCYCLES
                If (StationRecipe(iStn, iShift).Cycles = 1) Then
                    sCycles = Format(StationRecipe(iStn, iShift).Cycles, "###0") & " cycle of"
                Else
                    sCycles = Format(StationRecipe(iStn, iShift).Cycles, "###0") & " cycles of"
                End If
            Case ENDWEIGHTCHG
                If (StationRecipe(iStn, iShift).EndMinimumCycles = 1) Then
                    sCycles = "at least " & Format(StationRecipe(iStn, iShift).EndMinimumCycles, "###0") & " cycle of"
                Else
                    sCycles = "at least " & Format(StationRecipe(iStn, iShift).EndMinimumCycles, "###0") & " cycles of"
                End If
            Case Else
                If (StationRecipe(iStn, iShift).Cycles = 1) Then
                    sCycles = Format(StationRecipe(iStn, iShift).Cycles, "###0") & " cycle of"
                Else
                    sCycles = Format(StationRecipe(iStn, iShift).Cycles, "###0") & " cycles of"
                End If
        End Select
        ' Leak Check - add leak check phase description
        sLeak = IIf(StationRecipe(iStn, iShift).LeakCheck, "LeakCheck", "No LeakCheck")
        ' Purge - add Purge method description
        Select Case StationRecipe(iStn, iShift).Purge_Method
            Case NOPURGE
                sPurge = "No Purge"
            Case PURGEBYTIME
                sPurge = "Purge for " & StationRecipe(iStn, iShift).Purge_Time & " Minute"
                If StationRecipe(iStn, iShift).Purge_Time > 1 Then sPurge = sPurge & "s"
            Case PURGEBYLITERS
                sPurge = "Purge " & StationRecipe(iStn, iShift).Purge_Liters & " liter"
                If StationRecipe(iStn, iShift).Purge_Liters <> 1 Then sPurge = sPurge & "s"
            Case PURGEBYVOLUME
                sPurge = "Purge " & StationRecipe(iStn, iShift).Purge_Can_Vol & " Canister Volume"
                If StationRecipe(iStn, iShift).Purge_Can_Vol <> 1 Then sPurge = sPurge & "s"
            Case PURGEAUXONLY
                sPurge = "Purge Aux Can for " & StationRecipe(iStn, iShift).Purge_AuxTime & " Minute"
                If StationRecipe(iStn, iShift).Purge_AuxTime > 1 Then sPurge = sPurge & "s"
            Case PURGEBYPROFILE
                ' PURGE BY PROFILE
                sPurge = "Purge by Profile"
            Case PURGEBYWC
                ' PURGE BY WORKING CAPACITY
                sPurge = "Purge " & StationRecipe(iStn, iShift).Purge_TargetWC & " % of Work Cap"
            Case PURGETOTARGET
                ' PURGE TO TARGET WEIGHT
                sPurge = "Purge to " & StationRecipe(iStn, iShift).Purge_TargetWeight & " Grams"
            Case PURGETOUNDOLOAD
                ' PURGE TO UNDO LOAD
                sPurge = "Purge to Undo Load"
            Case Else
                sPurge = "Do Not Purge"
        End Select
        ' Load - add load method description
        Select Case StationRecipe(iStn, iShift).Load_MethodSave
            Case NOLOAD
                sLoad = LoadTypeDesc(NOLOAD)
            Case LOADBYTIME
                sLoad = "Load for "
                sLoad = sLoad & Format(StationRecipe(iStn, iShift).Load_Time, "##0")
                sLoad = sLoad & " Minute"
                If StationRecipe(iStn, iShift).Load_Time > 1 Then sLoad = sLoad & "s"
            Case LOADBYWC
                sLoad = LoadTypeDesc(LOADBYWC)
                sLoad = sLoad & Format(StationRecipe(iStn, iShift).WC_MultSave, "##0.#")
                sLoad = sLoad & LoadTypeDesc2(LOADBYWC)
                sLoad = sLoad & Format(StationRecipe(iStn, iShift).EPAFill, "##0")
                sLoad = sLoad & LoadTypeDesc3(LOADBYWC)
            Case LOADBYWEIGHT
                sLoad = "Load "
                If Int(StationRecipe(iStn, iShift).Load_Wt) = StationRecipe(iStn, iShift).Load_Wt Then
                    ' no digits to the right of the decimal point
                    sLoad = sLoad & Format(StationRecipe(iStn, iShift).Load_Wt, "##0")
                Else
                    ' digit(s) to the right of the decimal point
                    sLoad = sLoad & Format(StationRecipe(iStn, iShift).Load_Wt, "##0.##")
                End If
                sLoad = sLoad & LoadTypeDesc2(LOADBYWEIGHT)
            Case LOADBYBREAKTHRU
                sLoad = LoadTypeDesc(LOADBYBREAKTHRU)
                If Int(StationRecipe(iStn, iShift).LoadBreakthrough) = StationRecipe(iStn, iShift).LoadBreakthrough Then
                    ' no digits to the right of the decimal point
                    sLoad = sLoad & Format(StationRecipe(iStn, iShift).LoadBreakthrough, "##0")
                Else
                    ' digit(s) to the right of the decimal point
                    sLoad = sLoad & Format(StationRecipe(iStn, iShift).LoadBreakthrough, "##0.##")
                End If
                sLoad = sLoad & LoadTypeDesc2(LOADBYBREAKTHRU)
                sLoad = sLoad
                sLoad = sLoad & LoadTypeDesc3(LOADBYBREAKTHRU)
            Case LOADBYFID
                sLoad = LoadTypeDesc(LOADBYFID)
                sLoad = sLoad & Format(StationRecipe(iStn, iShift).FIDmg, "#####0")
                sLoad = sLoad & LoadTypeDesc2(LOADBYFID)
                sLoad = sLoad
                sLoad = sLoad & LoadTypeDesc3(LOADBYFID)
            Case Else
                sLoad = "Do Not Load"
        End Select
    
        ' Update Station Recipe Description Array
        If Len(sStart) > 1 Then
            ' delayed start of some kind; use Start Desc; do not mention LeakCheck
            StationRecipe(iStn, iShift).desc(0) = sStart & " then " & sCycles
        Else
            ' no start delay;  do not mention Start; use LeakCheck desc
            StationRecipe(iStn, iShift).desc(0) = sLeak & " then " & sCycles
        End If
        ' Purge/Load OR Load/Purge  ??
        Select Case StationRecipe(iStn, iShift).CycleType
            Case CyclePurgeLoad
                StationRecipe(iStn, iShift).desc(1) = sPurge
                StationRecipe(iStn, iShift).desc(2) = sLoad
            Case CycleLoadPurge
                StationRecipe(iStn, iShift).desc(1) = sLoad
                StationRecipe(iStn, iShift).desc(2) = sPurge
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

Sub Load_ScreenColors()
' Load ScreenColors configuration information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 4420

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim Idx As Integer

    Idx = 1

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
            ' Read FID Configuration Records
            Criteria = "SELECT * FROM [ScreenColors] WHERE [ColorSet] = " & Idx & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                SetDefault_ScreenColors
            Else
                rsRecord.MoveFirst
                Common_BackColor = rsRecord("Common_BackColor")
                Entry_BackColor = rsRecord("Entry_BackColor")
                EntryInvalid_BackColor = rsRecord("EntryInvalid_BackColor")
                EntryUnsaved_BackColor = rsRecord("EntryUnsaved_BackColor")
                EntryNotChangeable_BackColor = rsRecord("EntryNotChangeable_BackColor")
                MasterMode_BackColor = rsRecord("MasterMode_BackColor")
                StationMode_BackColor = rsRecord("StationMode_BackColor")
                Alarm_ForeColor = rsRecord("Alarm_ForeColor")
                BarActual_ForeColor = rsRecord("BarActual_ForeColor")
                Data_ForeColor = rsRecord("Data_ForeColor")
                DataBold_ForeColor = rsRecord("DataBold_ForeColor")
                DataHiLite_ForeColor = rsRecord("DataHiLite_ForeColor")
                Entry_ForeColor = rsRecord("Entry_ForeColor")
                Good_ForeColor = rsRecord("Good_ForeColor")
                Message_ForeColor = rsRecord("Message_ForeColor")
                Titles_ForeColor = rsRecord("Titles_ForeColor")
                TitlesData_Forecolor = rsRecord("TitlesData_ForeColor")
                TitlesLabel_ForeColor = rsRecord("TitlesLabel_ForeColor")
                Warning_ForeColor = rsRecord("Warning_ForeColor")
            End If
               
            rsRecord.Close
            
    ' Close Database
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

Sub Save_ScreenColors()
' Save Screen Colors configuration information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 4410

Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim Idx As Integer

    Idx = 1

    ' Open Database
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATASYSDEF)
    
            ' Save FID Configuration Records
            Criteria = "SELECT * FROM [ScreenColors] WHERE [ColorSet] = " & Idx & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
            If rsRecord.BOF Then
                rsRecord.AddNew
                rsRecord("ColorSet") = Idx
            Else
                rsRecord.MoveFirst
                rsRecord.Edit
            End If
               
            rsRecord("Common_BackColor") = Common_BackColor
            rsRecord("Entry_BackColor") = Entry_BackColor
            rsRecord("EntryInvalid_BackColor") = EntryInvalid_BackColor
            rsRecord("EntryUnsaved_BackColor") = EntryUnsaved_BackColor
            rsRecord("EntryNotChangeable_BackColor") = EntryNotChangeable_BackColor
            rsRecord("MasterMode_BackColor") = MasterMode_BackColor
            rsRecord("StationMode_BackColor") = StationMode_BackColor
            rsRecord("Alarm_ForeColor") = Alarm_ForeColor
            rsRecord("BarActual_ForeColor") = BarActual_ForeColor
            rsRecord("Data_ForeColor") = Data_ForeColor
            rsRecord("DataBold_ForeColor") = DataBold_ForeColor
            rsRecord("DataHiLite_ForeColor") = DataHiLite_ForeColor
            rsRecord("Entry_ForeColor") = Entry_ForeColor
            rsRecord("Good_ForeColor") = Good_ForeColor
            rsRecord("Message_ForeColor") = Message_ForeColor
            rsRecord("Titles_ForeColor") = Titles_ForeColor
            rsRecord("TitlesData_ForeColor") = TitlesData_Forecolor
            rsRecord("TitlesLabel_ForeColor") = TitlesLabel_ForeColor
            rsRecord("Warning_ForeColor") = Warning_ForeColor
               
            rsRecord.Update
            rsRecord.Close
    
    ' Close Database
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

Sub SetDefault_ScreenColors()
' Set Default Screen Colors
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 4400

    '   backcolors
    If VarInitDone Then
        Common_BackColor = frmDisplayProperties.Point(90, 90)
    Else
        Common_BackColor = vbButtonFace
    End If
    Entry_BackColor = GhostWhite
    EntryInvalid_BackColor = Yellow
    EntryUnsaved_BackColor = Wheat
    EntryNotChangeable_BackColor = CornSilk
'    MasterMode_BackColor = RGBFromRedGreenBlue(CLng(250), CLng(210), CLng(160))
    MasterMode_BackColor = Tan
    StationMode_BackColor = Common_BackColor
    
    '   forecolors
    Alarm_ForeColor = Red
    BarActual_ForeColor = RoyalBlue
    Data_ForeColor = SteelBlue
    DataBold_ForeColor = MediumBlue
    DataHiLite_ForeColor = DodgerBlue
    Entry_ForeColor = Black
    Good_ForeColor = LimeGreen
    Message_ForeColor = Purple
    Titles_ForeColor = Teal
    TitlesData_Forecolor = RoyalBlue
    TitlesLabel_ForeColor = SaddleBrown
    Warning_ForeColor = DarkOrange

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

Private Sub DecodeAdfDef(ByVal iStn As Integer)
' set ADF_DEF elements from LiveFuel TankType
    Select Case STN_INFO(iStn).ADF_TANKTYPE
        Case 1        '(Mark IV)
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = True
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = True
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = False
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = False
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = False
            STN_INFO(iStn).ADF_DEF.hasADF_FST = False
            STN_INFO(iStn).ADF_DEF.hasADF_LT = False
        Case 7        '(Honda R&D)
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = True
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = True
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = False
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = False
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = True
            STN_INFO(iStn).ADF_DEF.hasADF_FST = False
            STN_INFO(iStn).ADF_DEF.hasADF_LT = False
        Case 12       '(Mahle)
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = True
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = True
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = True
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = True
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = False
            STN_INFO(iStn).ADF_DEF.hasADF_FST = False
            STN_INFO(iStn).ADF_DEF.hasADF_LT = False
        Case 20       '(Stant)
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = True
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = True
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = True
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = True
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = False
            STN_INFO(iStn).ADF_DEF.hasADF_FST = False
            STN_INFO(iStn).ADF_DEF.hasADF_LT = True
        Case 22       '(Chrysler)
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = True
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = True
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = True
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = False
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = False
            STN_INFO(iStn).ADF_DEF.hasADF_FST = True
            STN_INFO(iStn).ADF_DEF.hasADF_LT = True
        Case 90       '(Honda R&D)
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = True
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = False
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = False
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = False
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = True
            STN_INFO(iStn).ADF_DEF.hasADF_FST = False
            STN_INFO(iStn).ADF_DEF.hasADF_LT = False
        Case 0
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = False
            STN_INFO(iStn).ADF_DEF.hasAUTODRAINFILL = False
            STN_INFO(iStn).ADF_DEF.hasADF_VaporValve = False
            STN_INFO(iStn).ADF_DEF.hasADF_Heater = False
            STN_INFO(iStn).ADF_DEF.hasADF_WaterBath = False
            STN_INFO(iStn).ADF_DEF.hasADF_FST = False
            STN_INFO(iStn).ADF_DEF.hasADF_LT = False
            Write_ELog "TankType(=" & Format(STN_INFO(iStn).ADF_TANKTYPE, "#0") & ") for Station #" & Format(iStn, "#0")
        Case Else
            STN_INFO(iStn).ADF_DEF.hasLIVEFUEL = False
            Write_ELog "Invalid TankType(=" & Format(STN_INFO(iStn).ADF_TANKTYPE, "#0") & ") for Station #" & Format(iStn, "#0")
    End Select
        
End Sub

Sub Save_LeakTest()
' Save LeakTest configuration and recipe information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 6571

Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStn As Integer

    filler = "1234567890"
    
    sFileName = FILEPATH_cfg & "leaktest.@@@"
    iFileNumber = FreeFile
    Open sFileName For Output As #iFileNumber
    
    For iStn = 1 To MAX_STN
        '  LeakTest Configuration
        Write #iFileNumber, Cfg_LeakTest.DeffTol
        Write #iFileNumber, Cfg_LeakTest.InitialN2Flow
        Write #iFileNumber, Cfg_LeakTest.PressTimeout
        Write #iFileNumber, Cfg_LeakTest.PressTol
        Write #iFileNumber, Cfg_LeakTest.PressTolDuration
    
        Write #iFileNumber, Cfg_LeakTest.timeOut
        Write #iFileNumber, Cfg_LeakTest.ReportInterval
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
    
        '  LeakTest Recipe
        Write #iFileNumber, Rcp_LeakTest.HoldDuration
        Write #iFileNumber, Rcp_LeakTest.TargetPress
        Write #iFileNumber, filler
        Write #iFileNumber, filler
        Write #iFileNumber, filler
    Next iStn
    
    Close #iFileNumber
    
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

Sub Load_LeakTest()
' Load LeakTest configuration and recipe information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 6572

Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStn As Integer

    sFileName = FILEPATH_cfg & "leaktest.@@@"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    For iStn = 1 To MAX_STN
        '  LeakTest Configuration
        Input #iFileNumber, Cfg_LeakTest.DeffTol
        Input #iFileNumber, Cfg_LeakTest.InitialN2Flow
        Input #iFileNumber, Cfg_LeakTest.PressTimeout
        Input #iFileNumber, Cfg_LeakTest.PressTol
        Input #iFileNumber, Cfg_LeakTest.PressTolDuration
    
        Input #iFileNumber, Cfg_LeakTest.timeOut
        Input #iFileNumber, Cfg_LeakTest.ReportInterval
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
    
        '  LeakTest Recipe
        Input #iFileNumber, Rcp_LeakTest.HoldDuration
        Input #iFileNumber, Rcp_LeakTest.TargetPress
        Input #iFileNumber, filler
        Input #iFileNumber, filler
        Input #iFileNumber, filler
    Next iStn

    Close #iFileNumber

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

Sub Load_Test()
' Load Test
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 4, 2756

Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim Idx As Integer
Dim TestSP(0 To 10) As Single
Dim TestDur(0 To 10) As Single

    sFileName = FILEPATH_rcp & "importfile.prg"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    For Idx = 1 To 10
        '  Test Configuration
        Input #iFileNumber, TestSP(Idx), TestDur(Idx)
    Next Idx

    Close #iFileNumber

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


